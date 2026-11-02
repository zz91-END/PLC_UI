// =====================================================================================
//  TiaService.cs —— TIA Portal V19 Openness 核心逻辑封装（V0.2）
//  -------------------------------------------------------------------------------------
//  职责（全部在后台线程执行，界面不会卡死）：
//    1. 把 SCL 文本写入本地文件（自动建目录，ANSI/GBK 无 BOM，避免中文乱码）
//    2. 附加到正在运行的 TIA Portal V19
//    3. 遍历项目设备，用 device.GetService<SoftwareContainer>() 取得 PlcSoftware
//    4. ExternalSourceGroup.ExternalSources.CreateFromFile 导入 SCL 外部源
//    5. GenerateBlocksFromSource(GenerateBlockOption.None) 生成程序块
//    6. GetService<ICompilable>().Compile() 编译，并收集编译消息
//    7. V0.2 新增：读取 PLC 变量表（Tag Table），供 AI 生成代码时复用真实变量
//
//  与之前控制台版本的区别：用 IProgress<string> 汇报进度，所有异常统一转换成
//  TiaOperationResult.FailureMessage，界面层不需要引用任何 Siemens 类型。
// =====================================================================================

using Siemens.Engineering;
using Siemens.Engineering.Compiler;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.ExternalSources;
using Siemens.Engineering.SW.Tags;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace PLC_UI
{
    /// <summary>SCL 写入 → 导入 → 生成块 → 编译 的整体结果。</summary>
    public sealed class TiaOperationResult
    {
        /// <summary>是否整体成功（块已生成且编译无错误）。</summary>
        public bool Succeeded { get; set; }

        /// <summary>最终生成的块名（由博图 BlockGroup 中实际查到的块名决定）。</summary>
        public string BlockName { get; set; }

        /// <summary>本地 SCL 文件路径。</summary>
        public string SclFilePath { get; set; }

        /// <summary>编译状态（CompilerResult.State 的字符串形式）。</summary>
        public string CompileState { get; set; }

        public int ErrorCount { get; set; }

        public int WarningCount { get; set; }

        /// <summary>失败时的详细信息（异常链 + 排查提示），成功时为 null。</summary>
        public string FailureMessage { get; set; }

        /// <summary>逐条格式化后的编译消息。</summary>
        public List<string> CompileMessages { get; private set; }

        public TiaOperationResult()
        {
            CompileMessages = new List<string>();
        }
    }

    /// <summary>把 SCL 源码导入运行中的 TIA Portal V19、生成块并编译。</summary>
    public sealed class TiaService
    {
        /// <summary>SCL 临时文件目录（不存在会自动创建），默认 C:\temp。</summary>
        public string SclDirectory { get; set; }

        public TiaService()
        {
            SclDirectory = @"C:\temp";
        }

        /// <summary>
        /// 在后台线程完成“写文件 → 附加博图 → 导入外部源 → 生成块 → 编译”。
        /// 不会抛出业务异常：失败信息放在 <see cref="TiaOperationResult.FailureMessage"/>。
        /// </summary>
        /// <param name="sclSource">完整的 SCL 文本（FUNCTION_BLOCK ... END_FUNCTION_BLOCK）。</param>
        /// <param name="blockName">期望的块名，可为 null（由生成结果自动识别）。</param>
        /// <param name="progress">进度/日志回调（由界面用 Progress&lt;string&gt; 传入，自动回到 UI 线程）。</param>
        public Task<TiaOperationResult> ImportAndCompileAsync(string sclSource, string blockName,
                                                              IProgress<string> progress)
        {
            if (string.IsNullOrWhiteSpace(sclSource))
            {
                throw new ArgumentException("SCL 源码不能为空。", "sclSource");
            }

            return Task.Run(() => RunPipeline(sclSource, blockName, progress));
        }

        // =================================================================================
        //  读取 PLC 变量表（V0.2 新增）
        // =================================================================================
        /// <summary>
        /// 读取当前博图项目里 PLC 变量表的全部变量，格式 "变量名 : 数据类型"（如 "StartButton : Bool"）。
        /// 该方法是同步阻塞的（内部会 Attach 博图、遍历 TagTable），在界面里请用
        /// <see cref="GetPlcVariablesAsync"/> 以免卡死 UI。
        /// 找不到 TIA / 项目 / PlcSoftware / 变量表时一律返回空列表，不抛异常。
        /// </summary>
        public List<string> GetPlcVariables()
        {
            List<string> variables = new List<string>();
            TiaPortal tiaPortal = null;

            try
            {
                tiaPortal = AttachToRunningTiaPortal(null);

                Project project = GetFirstOpenProject(tiaPortal);

                string ownerName;
                PlcSoftware plcSoftware = FindPlcSoftware(project, out ownerName);
                if (plcSoftware == null)
                {
                    return variables;
                }

                // PlcSoftware.TagTableGroup 是“PLC 变量表系统组”，下面是用户组（Groups）和变量表（TagTables）
                PlcTagTableGroup rootGroup = plcSoftware.TagTableGroup;
                if (rootGroup == null)
                {
                    return variables;
                }

                CollectVariables(rootGroup, variables);
                return variables;
            }
            catch (Exception)
            {
                // 读取变量表属于辅助功能，任何失败都只返回空列表，不影响主流程
                return variables;
            }
            finally
            {
                if (tiaPortal != null)
                {
                    try
                    {
                        tiaPortal.Dispose();
                    }
                    catch (Exception)
                    {
                        // 断开失败忽略
                    }
                }
            }
        }

        /// <summary>后台线程版本（界面用 await 调用，不卡 UI）。</summary>
        public Task<List<string>> GetPlcVariablesAsync()
        {
            return Task.Run(() => GetPlcVariables());
        }

        /// <summary>递归遍历变量表分组：先收集本组变量表，再递归子分组。</summary>
        private static void CollectVariables(PlcTagTableGroup group, List<string> output)
        {
            if (group == null)
            {
                return;
            }

            foreach (PlcTagTable table in group.TagTables)
            {
                foreach (PlcTag tag in table.Tags)
                {
                    string dataType = tag.DataTypeName;
                    if (string.IsNullOrEmpty(dataType))
                    {
                        dataType = "Unknown";
                    }

                    output.Add(tag.Name + " : " + dataType);
                }
            }

            foreach (PlcTagTableGroup childGroup in group.Groups)
            {
                CollectVariables(childGroup, output);
            }
        }

        // =================================================================================
        //  主流程
        // =================================================================================
        private TiaOperationResult RunPipeline(string sclSource, string blockName, IProgress<string> progress)
        {
            TiaOperationResult result = new TiaOperationResult();
            TiaPortal tiaPortal = null;

            try
            {
                // ---------- 1. 写本地 SCL 文件 ----------
                Report(progress, "[1/5] 写入本地 SCL 文件");
                result.SclFilePath = WriteSclFile(sclSource, blockName);
                Report(progress, "      已保存：" + result.SclFilePath +
                                 "（编码 " + Encoding.Default.WebName + "，无 BOM）");

                // ---------- 2. 连接 TIA Portal ----------
                Report(progress, "[2/5] 连接正在运行的 TIA Portal V19");
                tiaPortal = AttachToRunningTiaPortal(progress);
                Report(progress, "      已连接 TIA Portal");

                // ---------- 3. 找项目与 PlcSoftware ----------
                Report(progress, "[3/5] 读取项目并定位 PlcSoftware");
                Project project = GetFirstOpenProject(tiaPortal);
                Report(progress, "      当前项目：" + project.Name);

                string ownerName;
                PlcSoftware plcSoftware = FindPlcSoftware(project, out ownerName);
                if (plcSoftware == null)
                {
                    throw new InvalidOperationException(
                        "项目中没有找到带软件的 PLC 设备（PlcSoftware）。" +
                        "请确认项目中已组态 S7-1200 / S7-1500 CPU。");
                }

                Report(progress, "      PlcSoftware 来源：" + ownerName + " → " + plcSoftware.Name);

                // ---------- 4. 导入外部源并生成块 ----------
                Report(progress, "[4/5] 导入 SCL 外部源并生成程序块");
                PlcExternalSource externalSource = ImportExternalSource(plcSoftware, result.SclFilePath,
                                                                        blockName, progress);
                PlcBlock block = GenerateBlock(plcSoftware, externalSource, blockName, progress);
                result.BlockName = block.Name;
                Report(progress, "      已生成块：" + block.Name);

                // ---------- 5. 编译 ----------
                Report(progress, "[5/5] 编译");
                CompilerResult compileResult = CompilePlc(plcSoftware);
                result.CompileState = compileResult.State.ToString();
                result.ErrorCount = compileResult.ErrorCount;
                result.WarningCount = compileResult.WarningCount;

                CollectCompilerMessages(compileResult.Messages, 0, result.CompileMessages);
                Report(progress, "      编译状态：" + result.CompileState +
                                 "，错误 " + result.ErrorCount + "，警告 " + result.WarningCount);
                foreach (string message in result.CompileMessages)
                {
                    Report(progress, "      " + message);
                }

                result.Succeeded = result.ErrorCount == 0 && compileResult.State != CompilerResultState.Error;
                return result;
            }
            catch (EngineeringSecurityException ex)
            {
                result.Succeeded = false;
                result.FailureMessage = FormatException(ex) + Environment.NewLine +
                    "原因提示：博图弹出的“允许 Openness 连接”确认框被拒绝了。" + Environment.NewLine +
                    "请重新运行并在弹窗中点击“是”，并确认当前 Windows 用户已加入“Openness”用户组。";
                return result;
            }
            catch (LicenseNotFoundException ex)
            {
                result.Succeeded = false;
                result.FailureMessage = FormatException(ex) + Environment.NewLine +
                    "原因提示：该操作需要 SIMATIC Openness 授权，请检查 Automation License Manager。";
                return result;
            }
            catch (EngineeringException ex)
            {
                result.Succeeded = false;
                result.FailureMessage = FormatException(ex) + Environment.NewLine +
                    "原因提示：TIA Portal 拒绝了本次 Openness 操作。" +
                    "常见原因：CPU 处于在线状态（编译前请切到离线）、块名冲突、SCL 语法错误。";
                return result;
            }
            catch (Exception ex)
            {
                result.Succeeded = false;
                result.FailureMessage = FormatException(ex);
                return result;
            }
            finally
            {
                // 只断开 Openness 连接，不会关闭博图、不会关闭项目
                if (tiaPortal != null)
                {
                    try
                    {
                        tiaPortal.Dispose();
                    }
                    catch (Exception)
                    {
                        // 断开失败不影响本次结果展示
                    }
                }
            }
        }

        // =================================================================================
        //  1. 写 SCL 文件
        // =================================================================================
        private string WriteSclFile(string sclSource, string blockName)
        {
            if (!Directory.Exists(SclDirectory))
            {
                Directory.CreateDirectory(SclDirectory);
            }

            string fileName = string.IsNullOrEmpty(blockName) ? "AI_SCL" : SanitizeFileName(blockName);
            string filePath = Path.Combine(SclDirectory, fileName + ".scl");

            // Encoding.Default：中文系统即 ANSI/GBK 且不写 BOM。
            // UTF-8 BOM 会被 TIA 当成首行内容导致“首行语法错误”；UTF-8 无 BOM 则会被按 ANSI 解析导致中文乱码。
            File.WriteAllText(filePath, sclSource, Encoding.Default);

            return filePath;
        }

        private static string SanitizeFileName(string name)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            StringBuilder builder = new StringBuilder(name.Length);
            for (int i = 0; i < name.Length; i++)
            {
                builder.Append(Array.IndexOf(invalidChars, name[i]) >= 0 ? '_' : name[i]);
            }

            string sanitized = builder.ToString().Trim();
            return sanitized.Length == 0 ? "AI_SCL" : sanitized;
        }

        // =================================================================================
        //  2. 附加运行中的 TIA Portal
        // =================================================================================
        private static TiaPortal AttachToRunningTiaPortal(IProgress<string> progress)
        {
            List<TiaPortalProcess> processes = new List<TiaPortalProcess>();
            foreach (TiaPortalProcess process in TiaPortal.GetProcesses())
            {
                processes.Add(process);
            }

            if (processes.Count == 0)
            {
                throw new InvalidOperationException(
                    "没有检测到正在运行的 TIA Portal 进程。请先打开博图 V19 并加载项目后再试。");
            }

            Report(progress, "      检测到 " + processes.Count + " 个 TIA Portal 进程：");
            foreach (TiaPortalProcess process in processes)
            {
                Report(progress, "        - Id=" + process.Id +
                                 "，安装目录=" + process.Path +
                                 "，模式=" + process.Mode);
            }

            // 多个实例时默认连接第一个；如需指定，可改用 TiaPortal.GetProcess(进程Id, 超时)
            TiaPortalProcess target = processes[0];
            Report(progress, "      若博图弹出“允许 Openness 连接”的确认框，请点击“是”。");

            TiaPortal tiaPortal = target.Attach();
            if (tiaPortal == null)
            {
                throw new InvalidOperationException("Attach() 返回了 null，连接 TIA Portal 失败。");
            }

            return tiaPortal;
        }

        // =================================================================================
        //  3. 项目 / PlcSoftware
        // =================================================================================
        private static Project GetFirstOpenProject(TiaPortal tiaPortal)
        {
            List<Project> projects = new List<Project>();
            foreach (Project project in tiaPortal.Projects)
            {
                projects.Add(project);
            }

            if (projects.Count == 0)
            {
                throw new InvalidOperationException("TIA Portal 当前没有打开任何项目，请先在博图中打开项目。");
            }

            return projects[0];
        }

        private static PlcSoftware FindPlcSoftware(Project project, out string ownerName)
        {
            ownerName = null;

            List<Device> devices = new List<Device>();
            foreach (Device device in project.Devices)
            {
                devices.Add(device);
            }

            if (devices.Count == 0)
            {
                throw new InvalidOperationException(
                    "项目中没有找到任何设备，请确认博图项目中已组态 S7-1200 / S7-1500 站点。");
            }

            foreach (Device device in devices)
            {
                PlcSoftware plcSoftware = TryGetPlcSoftware(device);
                if (plcSoftware != null)
                {
                    ownerName = device.Name;
                    return plcSoftware;
                }

                int itemIndex = 0;
                foreach (DeviceItem deviceItem in device.DeviceItems)
                {
                    itemIndex++;
                    plcSoftware = TryGetPlcSoftware(deviceItem);
                    if (plcSoftware != null)
                    {
                        ownerName = device.Name + " / 设备项 " + itemIndex;
                        return plcSoftware;
                    }
                }
            }

            return null;
        }

        private static PlcSoftware TryGetPlcSoftware(Device device)
        {
            try
            {
                SoftwareContainer container = device.GetService<SoftwareContainer>();
                return container == null ? null : container.Software as PlcSoftware;
            }
            catch (EngineeringException)
            {
                // 该设备没有软件容器，忽略并继续遍历
                return null;
            }
        }

        private static PlcSoftware TryGetPlcSoftware(DeviceItem deviceItem)
        {
            try
            {
                SoftwareContainer container = deviceItem.GetService<SoftwareContainer>();
                return container == null ? null : container.Software as PlcSoftware;
            }
            catch (EngineeringException)
            {
                return null;
            }
        }

        // =================================================================================
        //  4. 导入外部源 + 生成块
        // =================================================================================
        private static PlcExternalSource ImportExternalSource(PlcSoftware plcSoftware, string sclFilePath,
                                                              string blockName, IProgress<string> progress)
        {
            PlcExternalSourceGroup sourceGroup = plcSoftware.ExternalSourceGroup;
            if (sourceGroup == null)
            {
                throw new InvalidOperationException("未能获取 ExternalSourceGroup（外部源组），无法导入 SCL。");
            }

            string sourceName = BuildExternalSourceName(blockName);
            PlcExternalSource existingSource = sourceGroup.ExternalSources.Find(sourceName);
            if (existingSource != null)
            {
                existingSource.Delete();
                Report(progress, "      已删除同名旧外部源：" + sourceName);
            }

            // 参数顺序：CreateFromFile(外部源名称, 文件路径)
            PlcExternalSource externalSource =
                sourceGroup.ExternalSources.CreateFromFile(sourceName, sclFilePath);

            if (externalSource == null)
            {
                throw new InvalidOperationException("CreateFromFile() 返回 null，SCL 外部源导入失败。");
            }

            Report(progress, "      已导入外部源：" + externalSource.Name);
            return externalSource;
        }

        private static string BuildExternalSourceName(string blockName)
        {
            if (string.IsNullOrEmpty(blockName))
            {
                return "AI_SCL_Source";
            }

            string name = "AI_" + SanitizeFileName(blockName);
            return name.Length > 100 ? name.Substring(0, 100) : name;
        }

        private static PlcBlock GenerateBlock(PlcSoftware plcSoftware, PlcExternalSource externalSource,
                                              string blockName, IProgress<string> progress)
        {
            // 同名旧块先删掉，保证每次都是“重新生成”的完整验证
            if (!string.IsNullOrEmpty(blockName))
            {
                PlcBlock oldBlock = plcSoftware.BlockGroup.Blocks.Find(blockName);
                if (oldBlock != null)
                {
                    try
                    {
                        oldBlock.Delete();
                        Report(progress, "      已删除同名旧块：" + blockName);
                    }
                    catch (Exception ex)
                    {
                        Report(progress, "      警告：删除旧块失败（" + ex.Message + "），继续生成。");
                    }
                }
            }

            IList<IEngineeringObject> generatedObjects =
                externalSource.GenerateBlocksFromSource(GenerateBlockOption.None);

            int generatedCount = generatedObjects == null ? 0 : generatedObjects.Count;
            Report(progress, "      GenerateBlocksFromSource 返回对象数：" + generatedCount);

            PlcBlock block = null;
            if (!string.IsNullOrEmpty(blockName))
            {
                block = plcSoftware.BlockGroup.Blocks.Find(blockName);
            }

            if (block == null && generatedObjects != null)
            {
                foreach (IEngineeringObject generatedObject in generatedObjects)
                {
                    if (generatedObject is PlcBlock)
                    {
                        block = (PlcBlock)generatedObject;
                        break;
                    }
                }
            }

            if (block == null)
            {
                throw new InvalidOperationException(
                    "外部源已导入，但没有生成任何程序块。" +
                    "请检查 SCL 语法：块声明、END_FUNCTION_BLOCK 是否完整、变量声明是否正确。");
            }

            return block;
        }

        // =================================================================================
        //  5. 编译
        // =================================================================================
        private static CompilerResult CompilePlc(PlcSoftware plcSoftware)
        {
            // V19 没有 PlcSoftware.Compile()，也没有公开的 CompileProvider 类；
            // 正确做法是通过服务接口 ICompilable 编译。
            ICompilable compilable = plcSoftware.GetService<ICompilable>();
            if (compilable == null)
            {
                throw new InvalidOperationException(
                    "当前 PlcSoftware 没有提供 ICompilable 编译服务。" +
                    "请确认设备处于离线状态，且 Openness 版本与博图 V19 一致（x86）。");
            }

            return compilable.Compile();
        }

        private static void CollectCompilerMessages(CompilerResultMessageComposition messages, int depth,
                                                    List<string> output)
        {
            if (messages == null)
            {
                return;
            }

            string indent = new string(' ', depth * 2);
            foreach (CompilerResultMessage message in messages)
            {
                string description = string.IsNullOrEmpty(message.Description)
                    ? string.Empty
                    : message.Description.Replace("\r", " ").Replace("\n", " ");

                output.Add(indent + "[" + message.State + "] " + message.Path + "  " + description);

                CollectCompilerMessages(message.Messages, depth + 1, output);
            }
        }

        // =================================================================================
        //  辅助
        // =================================================================================
        private static void Report(IProgress<string> progress, string message)
        {
            if (progress != null)
            {
                progress.Report(message);
            }
        }

        private static string FormatException(Exception exception)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("******************** 捕获到异常 ********************");

            int index = 1;
            for (Exception ex = exception; ex != null; ex = ex.InnerException)
            {
                builder.AppendLine("[" + index + "] 类型：" + ex.GetType().FullName);
                builder.AppendLine("    信息：" + ex.Message);
                index++;
            }

            builder.AppendLine("****************************************************");
            return builder.ToString();
        }
    }
}