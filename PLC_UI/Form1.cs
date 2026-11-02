// =====================================================================================
//  Form1.cs —— TIA SCL 生成导入助手（V0.3）
//  -------------------------------------------------------------------------------------
//  V0.3 升级点（动态块名）：
//    1. BuildSystemPrompt 两种模式都要求 AI 按需求动态生成“唯一且有语义的块名”
//       （流水灯 → WaterLight，电机正反转 → MotorControl），禁止再用固定的 Control_Logic / DB_Variables
//    2. 新增 ExtractBlockName / ExtractAllBlockNames：兼容提取 FUNCTION_BLOCK / FUNCTION / DATA_BLOCK 的块名
//    3. btnGenerate_Click 把提取到的块名传给 TiaService（不再传 null），
//       使 .scl 文件名与外部源名（AI_<块名>）随程序变化，生成第二个程序时不会覆盖第一个
//
//  保持不变：流式输出（stream=true）、System.Collections.IList 防呆 JSON 解析、
//  FC 里 VAR→VAR_TEMP、FB 智能判断、DB+FC+FB 多块清洗。
// =====================================================================================

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace PLC_UI
{
    public partial class Form1 : Form
    {
        // -----------------------------------------------------------------------------
        //  配置区（改这里）
        // -----------------------------------------------------------------------------
        private const string DeepSeekApiKey = "";                                        // ← 填你的 DeepSeek API Key
        private const string DeepSeekApiUrl = "https://api.deepseek.com/chat/completions";
        private const string DeepSeekModel = "deepseek-chat";
        private const string ButtonIdleText = "生成并导入（DeepSeek → TIA Portal）";

        // -----------------------------------------------------------------------------
        //  各种块起始关键字：AI 把代码挤成一行时，用来强制换行
        // -----------------------------------------------------------------------------
        private static readonly string[] SclLineStartKeywords =
        {
            "END_DATA_BLOCK", "END_FUNCTION_BLOCK", "END_FUNCTION",
            "DATA_BLOCK", "FUNCTION_BLOCK", "FUNCTION", "VERSION",
            "VAR_INPUT", "VAR_OUTPUT", "VAR_IN_OUT", "END_VAR", "VAR", "BEGIN"
        };

        private static readonly string SclLineStartPattern =
            @"(?<=\S)[ \t]*(?=(?<!\w)(?:" + string.Join("|", SclLineStartKeywords) + @")(?!\w))";

        private static readonly Regex SclBlockStartRegex = new Regex(
            @"(?<!\w)(?:DATA_BLOCK|FUNCTION_BLOCK|FUNCTION)(?!\w)", RegexOptions.IgnoreCase);

        private static readonly Regex SclBlockEndRegex = new Regex(
            @"(?<!\w)(?:END_DATA_BLOCK|END_FUNCTION|END_FUNCTION_BLOCK)(?!\w)",
            RegexOptions.IgnoreCase | RegexOptions.RightToLeft);

        private static readonly Regex SclBeginRegex = new Regex(
            @"(?<!\w)BEGIN(?!\w)", RegexOptions.IgnoreCase);

        private static readonly Regex SclVarEndRegex = new Regex(
            @"(?<!\w)END_VAR(?!\w)", RegexOptions.IgnoreCase);

        // { S7_Optimized_Access := 'TRUE' } 必须独占一行，不能粘在块声明后面
        private static readonly Regex SclAttributeRegex = new Regex(
            @"(?<=\S)[ \t]*(?=\{\s*S7_)", RegexOptions.IgnoreCase);

        // VERSION : 0.1 后面如果粘着变量声明，要拆成新行（版本号本身不能被拆开）
        private static readonly Regex SclVersionTailRegex = new Regex(
            @"(?<=VERSION\s*:\s*[0-9]+\.[0-9]+)[ \t]*(?=[A-Za-z_""#])", RegexOptions.IgnoreCase);

        // 多个变量声明挤在一行时（"A : Bool; B : Bool;"），按分号拆行
        private static readonly Regex SclDeclarationSplitRegex = new Regex(
            @";[ \t]*(?=[A-Za-z_$][\w$]*[ \t]*:(?!=))");

        // 关键字后面直接粘着下一句代码时（"END_VAR IF ..."、"BEGIN x := 1;"），在关键字后补换行
        private static readonly Regex SclKeywordTailRegex = new Regex(
            @"(?<!\w)(?:END_VAR|END_FUNCTION_BLOCK|END_FUNCTION|END_DATA_BLOCK|BEGIN|VAR)(?!\w)[ \t]*(?=[A-Za-z_""'])",
            RegexOptions.IgnoreCase);

        // AI 把 END_VAR 和后面的语句完全粘死时（"END_VARIF ..."），也要拆开
        private static readonly Regex SclGluedKeywordRegex = new Regex(
            @"END_VAR(?=(?:IF|CASE|FOR|WHILE|REPEAT|RETURN|EXIT|CONTINUE)\b)",
            RegexOptions.IgnoreCase);

        // -----------------------------------------------------------------------------
        //  运行期字段
        // -----------------------------------------------------------------------------
        private static readonly HttpClient Http = CreateHttpClient();
        private static readonly JavaScriptSerializer JsonSerializer = new JavaScriptSerializer();
        private readonly TiaService _tiaService = new TiaService();

        public Form1()
        {
            InitializeComponent();
        }

        // =================================================================================
        //  修改点 1：双模 System Prompt
        // =================================================================================
        private static string BuildSystemPrompt(IList<string> variables)
        {
            // 模式 A：博图里没有任何变量表 → 让 AI 自己建 DB + FC
            if (variables == null || variables.Count == 0)
            {
                return BuildModeAPrompt();
            }

            // 模式 B：博图里已有变量表 → AI 评估后决定生成 FC 或 FB，块名同样动态生成
            return BuildModeBPrompt(variables);
        }

        /// <summary>模式 A：AI 必须生成 DATA_BLOCK "DB_程序名" + FUNCTION "程序名"（块名由 AI 按需求动态命名）。</summary>
        private static string BuildModeAPrompt()
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("你是西门子 TIA Portal 的 SCL 编程专家。当前博图项目里【没有任何现成的 PLC 变量表】，");
            builder.AppendLine("所以你必须自己创建一个全局数据块来存放变量，并写一个函数来使用它们。");
            builder.AppendLine();
            builder.AppendLine("【块名必须动态生成（非常重要）】");
            builder.AppendLine("请根据用户的需求，自动生成一个唯一且具有语义的块名（例如：流水灯程序叫 WaterLight，电机正反转叫 MotorControl）。");
            builder.AppendLine("不要每次都用固定的名字（如 Control_Logic / DB_Variables），也不要把示例名当成默认值照抄。");
            builder.AppendLine("块名只能用英文字母、数字、下划线，不能以数字开头，不能用中文。");
            builder.AppendLine("两个块必须使用同一个程序名；数据块名 = DB_ + 程序名（例如程序名 WaterLight → 数据块 DB_WaterLight）。");
            builder.AppendLine();
            builder.AppendLine("请一次性输出两个块（先 DB，后 FC），顺序和格式必须完全如下（把 <程序名> 替换成你起的名字），中间不要插入任何解释文字：");
            builder.AppendLine();
            builder.AppendLine("DATA_BLOCK \"DB_<程序名>\"");
            builder.AppendLine("{ S7_Optimized_Access := 'TRUE' }");
            builder.AppendLine("VERSION : 0.1");
            builder.AppendLine("   <变量名> : <数据类型>;");
            builder.AppendLine("   ...（按用户需求列出全部变量，每行一个，声明直接写在 DB 里）");
            builder.AppendLine("BEGIN");
            builder.AppendLine("END_DATA_BLOCK");
            builder.AppendLine();
            builder.AppendLine("FUNCTION \"<程序名>\" : Void");
            builder.AppendLine("{ S7_Optimized_Access := 'TRUE' }");
            builder.AppendLine("VERSION : 0.1");
            builder.AppendLine("BEGIN");
            builder.AppendLine("   ...（这里写控制逻辑）");
            builder.AppendLine("END_FUNCTION");
            builder.AppendLine();
            builder.AppendLine("强制规则（违反任何一条，代码都会导入失败）：");
            builder.AppendLine("1. 只输出 SCL 代码，不要 Markdown 围栏，不要任何解释、标题或多余文字。");
            builder.AppendLine("2. 绝对禁止生成 FUNCTION_BLOCK，绝对禁止生成 FB。");
            builder.AppendLine("3. 绝对禁止使用 VAR_INPUT / VAR_OUTPUT / VAR_IN_OUT 管脚声明。");
            builder.AppendLine("4. 所有变量都放在 DATA_BLOCK 里，声明格式为 \"变量名 : 类型;\"，");
            builder.AppendLine("   不要写 VAR / END_VAR，变量名不要带 # 号，不要加引号。");
            builder.AppendLine("5. 函数体里访问变量必须写成 \"DB_<程序名>\".变量名（与上面数据块的块名完全一致），例如：");
            builder.AppendLine("   \"DB_WaterLight\".Motor := TRUE;");
            builder.AppendLine("   IF \"DB_WaterLight\".StartButton THEN ... END_IF;");
            builder.AppendLine("6. 关键字必须各占一行：DATA_BLOCK、FUNCTION、VERSION、BEGIN、END_DATA_BLOCK、END_FUNCTION");
            builder.AppendLine("   不能和别的代码挤在同一行。");
            builder.AppendLine("7. 数据类型只能用标准 SCL 类型（Bool、Int、DInt、Real、Time、Timer 等）。");
            builder.AppendLine("8. 变量名只能用英文字母、数字、下划线，不能用中文；中文只允许出现在 // 注释里。");
            builder.AppendLine("9. 【致命警告】：在 FUNCTION 块中，绝对禁止使用 VAR 声明区！");
            builder.AppendLine("   所有的中间变量、自锁标志位，必须使用 VAR_TEMP ... END_VAR 声明！");
            builder.AppendLine("   （DATA_BLOCK 里的变量声明不受此限制，但 FUNCTION 里只能出现 VAR_TEMP。）");
            builder.AppendLine("10. 块名必须按需求动态生成，两次不同需求的程序不能出现相同的块名；");
            builder.AppendLine("    如果你不确定起什么名字，就用需求关键词的英文（例如 交通灯 → TrafficLight）。");
            builder.AppendLine();
            builder.AppendLine("请根据下面用户的需求，生成这两个完整的块。");
            return builder.ToString();
        }

        /// <summary>模式 B：AI 评估变量表后选择生成 FC 或 FB，块名由 AI 按需求动态命名。</summary>
        private static string BuildModeBPrompt(IList<string> variables)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("你是西门子 TIA Portal 的 SCL 编程专家。下面列出了本项目 PLC 变量表里【已有的全局变量】。");
            builder.AppendLine();
            builder.AppendLine("【块名必须动态生成（非常重要）】");
            builder.AppendLine("请根据用户的需求，自动生成一个唯一且具有语义的块名（例如：流水灯程序叫 WaterLight，电机正反转叫 MotorControl）。");
            builder.AppendLine("不要每次都用固定的名字（如 Control_Logic），也不要把示例名当成默认值照抄。");
            builder.AppendLine("FC 和 FB 都必须这样做；块名只能用英文字母、数字、下划线，不能以数字开头，不能用中文。");
            builder.AppendLine();
            builder.AppendLine("【第一步：先评估变量表，再决定用哪种方式】");
            builder.AppendLine("请先把用户需求与现有变量表做匹配判断，不要急着写代码：");
            builder.AppendLine("A. 如果现有变量表【能直接满足】当前需求（例如需求是电机正反转，表里正好有 电机正转/电机反转/启动/停止）：");
            builder.AppendLine("   生成一个 FC（FUNCTION 块），块名按需求动态命名，函数体里直接调用这些全局变量。");
            builder.AppendLine("   禁止生成 FUNCTION_BLOCK，禁止使用 VAR_INPUT / VAR_OUTPUT 管脚，禁止生成 DATA_BLOCK；");
            builder.AppendLine("   如果确实需要中间变量，只能用 VAR_TEMP ... END_VAR 声明。");
            builder.AppendLine("B. 如果现有变量表【无法满足】当前需求（例如用户要写流水灯，但变量表里全是电机控制相关的变量）：");
            builder.AppendLine("   请直接生成一个 FUNCTION_BLOCK（FB 块），块名同样按需求动态命名（见上面的块名规则），");
            builder.AppendLine("   在 FB 内部用 VAR 声明中间变量/自锁标志位，用 VAR_INPUT 声明启动、停止等输入，用 VAR_OUTPUT 声明输出；");
            builder.AppendLine("   绝对不要强行套用不相关的现有变量（流水灯代码里不能出现“电机正转/电机反转”）。");
            builder.AppendLine();
            builder.AppendLine("【第二步：必须在代码第一行用注释说明你的选择，二选一】");
            builder.AppendLine("  选择 A 时第一行写：// 采用现有变量");
            builder.AppendLine("  选择 B 时第一行写：// 采用 FB 块新建变量");
            builder.AppendLine("这句注释之后必须紧跟块声明，不要输出任何其它解释文字。");
            builder.AppendLine();
            builder.AppendLine("【方式 A（FC）的输出格式】");
            builder.AppendLine("FUNCTION \"<程序名>\" : Void");
            builder.AppendLine("{ S7_Optimized_Access := 'TRUE' }");
            builder.AppendLine("VERSION : 0.1");
            builder.AppendLine("VAR_TEMP");
            builder.AppendLine("   （可选的内部临时变量声明：变量名 : 类型;）");
            builder.AppendLine("END_VAR");
            builder.AppendLine("BEGIN");
            builder.AppendLine("   ...（这里写控制逻辑，直接引用 \"电机正转\" 这样的全局变量）");
            builder.AppendLine("END_FUNCTION");
            builder.AppendLine();
            builder.AppendLine("【方式 B（FB）的输出格式】");
            builder.AppendLine("FUNCTION_BLOCK \"<程序名>\"");
            builder.AppendLine("{ S7_Optimized_Access := 'TRUE' }");
            builder.AppendLine("VERSION : 0.1");
            builder.AppendLine("VAR_INPUT");
            builder.AppendLine("   StartButton : Bool;");
            builder.AppendLine("END_VAR");
            builder.AppendLine("VAR_OUTPUT");
            builder.AppendLine("   Motor : Bool;");
            builder.AppendLine("END_VAR");
            builder.AppendLine("VAR");
            builder.AppendLine("   temp : Bool;");
            builder.AppendLine("END_VAR");
            builder.AppendLine("BEGIN");
            builder.AppendLine("   ...（这里写控制逻辑）");
            builder.AppendLine("END_FUNCTION_BLOCK");
            builder.AppendLine();
            builder.AppendLine("强制规则（违反任何一条，代码都会导入失败）：");
            builder.AppendLine("1. 只输出 SCL 代码，不要 Markdown 围栏，不要任何标题或多余文字；唯一允许的注释是第一行那句选择说明。");
            builder.AppendLine("2. 严禁“为了用变量而用变量”：变量不匹配时走 B 方案，不要生搬硬套。");
            builder.AppendLine("3. 方式 A（FC）里绝对禁止出现 FUNCTION_BLOCK、VAR、VAR_INPUT、VAR_OUTPUT。");
            builder.AppendLine("4. 方式 A（FC）里绝对禁止使用 VAR 声明区！中间变量只能用 VAR_TEMP ... END_VAR（FC 无记忆，写 VAR 会报“区域 VAR 不可用于 FC 类型的块”）。");
            builder.AppendLine("5. 方式 B（FB）里必须用 VAR / VAR_INPUT / VAR_OUTPUT 声明自己的变量，不要生成 DATA_BLOCK，不要引用不相关变量。");
            builder.AppendLine("6. 关键字必须各占一行：FUNCTION、FUNCTION_BLOCK、VERSION、VAR、VAR_INPUT、VAR_OUTPUT、VAR_TEMP、END_VAR、BEGIN、END_FUNCTION、END_FUNCTION_BLOCK，不能和别的代码挤在同一行。");
            builder.AppendLine("7. 变量名只能用英文字母、数字、下划线，不能用中文；中文只允许出现在 // 注释里。");
            builder.AppendLine("8. 块名必须按需求动态生成：不同需求必须用不同块名（例如流水灯 → WaterLight、电机正反转 → MotorControl），");
            builder.AppendLine("   不要把示例名或 Control_Logic 当成固定答案。");
            builder.AppendLine();
            builder.AppendLine("【本项目 PLC 变量表里已有的变量（来自博图联机读取）】");

            int limit = Math.Min(variables.Count, 80);
            for (int i = 0; i < limit; i++)
            {
                builder.AppendLine("- " + variables[i]);
            }

            if (variables.Count > limit)
            {
                builder.AppendLine("- ...（共 " + variables.Count + " 个变量，此处仅列出前 " + limit + " 个）");
            }

            builder.AppendLine();
            builder.AppendLine("请先评估变量表，再决定生成 FC 还是 FB，并在第一行注释里写明你的选择。");
            return builder.ToString();
        }

        // =================================================================================
        //  按钮主流程（V0.3：提取 AI 生成的动态块名，传给 TiaService 避免覆盖旧程序）
        // =================================================================================
        private async void btnGenerate_Click(object sender, EventArgs e)
        {
            string userPrompt = txtPrompt.Text.Trim();
            if (userPrompt.Length == 0)
            {
                MessageBox.Show(this, "请先在左侧输入你的控制需求。", "提示",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtPrompt.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(DeepSeekApiKey))
            {
                MessageBox.Show(this, "请先在 Form1.cs 里填写 DeepSeekApiKey。", "提示",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnGenerate.Enabled = false;
            btnGenerate.Text = "生成中，请稍候...";
            Cursor = Cursors.WaitCursor;
            txtLog.Clear();

            try
            {
                // ---------- 步骤 0：读取博图真实变量表（决定双模） ----------
                AppendLog("[0/3] 读取博图中的 PLC 变量表 ...");
                List<string> variables = null;
                try
                {
                    // 后台线程执行 Openness 遍历，不会卡住界面
                    variables = await _tiaService.GetPlcVariablesAsync();
                }
                catch (Exception ex)
                {
                    AppendLog("      读取变量表失败，按“模式 A（无变量）”处理（" + ex.Message + "）");
                }

                if (variables == null)
                {
                    variables = new List<string>();
                }

                bool hasVariables = variables.Count > 0;
                AppendLog("      读取到 " + variables.Count + " 个变量 → 当前模式：" +
                          (hasVariables ? "B（复用已有变量，只生成 FUNCTION）"
                                        : "A（无变量，生成 DATA_BLOCK + FUNCTION）"));

                if (hasVariables)
                {
                    int previewCount = Math.Min(variables.Count, 15);
                    for (int i = 0; i < previewCount; i++)
                    {
                        AppendLog("        - " + variables[i]);
                    }

                    if (variables.Count > previewCount)
                    {
                        AppendLog("        ...（其余 " + (variables.Count - previewCount) + " 个已注入 Prompt）");
                    }
                }

                // ---------- 步骤 1：拼装双模 Prompt + 流式调用 DeepSeek ----------
                string systemPrompt = BuildSystemPrompt(variables);
                AppendLog("[1/3] 调用 DeepSeek 生成 SCL（流式输出，模式 " + (hasVariables ? "B" : "A") + "）...");
                AppendLog("---------------------------------------- SCL 开始 ----------------------------------------");

                string streamedText = await RequestSclFromDeepSeekAsync(systemPrompt, userPrompt, AppendStreamChunk);

                AppendLog("");
                AppendLog("---------------------------------------- SCL 结束 ----------------------------------------");

                if (string.IsNullOrWhiteSpace(streamedText))
                {
                    throw new InvalidOperationException("DeepSeek 返回内容为空。");
                }

                // ---------- 步骤 2：清洗 SCL（兼容 DB + FC + FB 多块） ----------
                AppendLog("[2/3] 清理 SCL 文本（DB + FC + FB 多块模式）...");
                string aiChoice = DetectAiBlockChoice(streamedText);
                if (aiChoice != null)
                {
                    AppendLog("      AI 选择：" + aiChoice);
                }

                string scl = CleanSclText(streamedText);
                if (string.IsNullOrWhiteSpace(scl))
                {
                    throw new InvalidOperationException("未能从返回内容中提取到 SCL 代码。");
                }

                AppendLog("      识别到块：" + DescribeSclBlocks(scl));

                // 提取 AI 生成的块名：不同程序用不同块名，从而得到不同的 .scl 文件名和外部源名，避免覆盖已有程序。
                List<string> allBlockNames = ExtractAllBlockNames(scl);
                if (allBlockNames.Count > 0)
                {
                    AppendLog("      提取到块名：" + string.Join("、", allBlockNames.ToArray()));
                }

                string blockName = ExtractBlockName(scl);
                if (string.IsNullOrEmpty(blockName))
                {
                    AppendLog("      ⚠ 未提取到块名，将由 Openness 按 SCL 内容自动识别（可能与已有块重名）。");
                }
                else
                {
                    AppendLog("      导入用主块名：" + blockName + "（TiaService 用它命名 .scl 文件和外部源）");
                }

                AppendLog("      SCL 共 " + CountLines(scl) + " 行。");
                AppendLog(TrimForLog(scl, 40, 4000));

                // ---------- 步骤 3：导入博图 + 生成块 + 编译（传入动态块名） ----------
                AppendLog("[3/3] 导入博图并编译（主块名：" + (blockName ?? "自动识别") + "）...");
                IProgress<string> progress = new Progress<string>(AppendLog);

                // blockName 非 null：TiaService 用它作为 .scl 文件名和外部源名（AI_<块名>），
                // 因此每次生成不同的程序时不会覆盖上一次的程序块。
                TiaOperationResult result = await _tiaService.ImportAndCompileAsync(scl, blockName, progress);

                if (result.Succeeded)
                {
                    AppendLog("✅ 完成：编译状态 " + result.CompileState +
                              "，错误 0，警告 " + result.WarningCount + "。");
                    MessageBox.Show(this,
                        "导入并编译成功！\n\n编译状态：" + result.CompileState +
                        "\n警告数：" + result.WarningCount +
                        "\nSCL 文件：" + result.SclFilePath +
                        "\n\n提示：本次生成的块名为 " + (blockName ?? "(自动识别)") +
                        "，请到博图的“程序块”里查看。",
                        "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    AppendLog("❌ 失败：" + (result.FailureMessage ?? "未知错误"));
                    MessageBox.Show(this, result.FailureMessage ?? "未知错误", "导入/编译失败",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                AppendLog("❌ 异常：" + ex.Message);
                MessageBox.Show(this, ex.Message, "异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnGenerate.Enabled = true;
                btnGenerate.Text = ButtonIdleText;
                Cursor = Cursors.Default;
            }
        }

        // =================================================================================
        //  修改点 2：CleanSclText（兼容 DB + FC 多块）
        // =================================================================================
        private static string CleanSclText(string rawText)
        {
            if (string.IsNullOrEmpty(rawText))
            {
                return string.Empty;
            }

            // 统一换行 + 去掉 BOM / 首尾空白
            string text = rawText.Replace("\r\n", "\n").Replace('\r', '\n');
            text = text.Trim('\uFEFF', ' ', '\t', '\n', '\u00A0');

            // 1) 去掉 Markdown 围栏
            StringBuilder withoutFences = new StringBuilder();
            string[] lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].Trim().StartsWith("```", StringComparison.Ordinal))
                {
                    continue;
                }

                withoutFences.Append(lines[i]);
                if (i < lines.Length - 1)
                {
                    withoutFences.Append('\n');
                }
            }

            text = withoutFences.ToString();

            // 2) AI 挤成一行时，强制在关键字前换行
            text = Regex.Replace(text, SclLineStartPattern, "\n");

            // 2.1) 属性行 { S7_Optimized_Access := 'TRUE' } 必须独占一行
            text = SclAttributeRegex.Replace(text, "\n");

            // 2.2) VERSION 行后面的变量声明要换到下一行
            text = SclVersionTailRegex.Replace(text, "\n");

            // 2.3) 关键字后面粘着下一句代码时，关键字后补换行
            text = SclKeywordTailRegex.Replace(text, "$&\n");
            text = SclGluedKeywordRegex.Replace(text, "END_VAR\n");

            // 2.4) 上面拆完后再做一次「关键字前换行」，保证 "Bool;END_VAR" 也能拆开
            text = Regex.Replace(text, SclLineStartPattern, "\n");

            // 3) 多块提取：从最早的 DATA_BLOCK / FUNCTION 开始，
            //    到最远的 END_DATA_BLOCK / END_FUNCTION / END_FUNCTION_BLOCK 结束
            Match blockStart = SclBlockStartRegex.Match(text);
            if (blockStart.Success)
            {
                Match blockEnd = SclBlockEndRegex.Match(text);
                if (blockEnd.Success && blockEnd.Index + blockEnd.Length > blockStart.Index)
                {
                    text = text.Substring(blockStart.Index,
                                          blockEnd.Index + blockEnd.Length - blockStart.Index);
                }
                else
                {
                    // 有块头但没有块尾：先截掉前面的废话，后面由补全逻辑处理
                    text = text.Substring(blockStart.Index);
                }
            }

            // 4) 清理声明区多余的 # 号（DB 声明行 / VAR..END_VAR 声明行）
            text = RemoveHashInDeclarations(text);

            // 4.1) 多个声明挤在一行时按分号拆行（放在去 # 之后，否则 "#Motor : ..." 匹配不到）
            text = SclDeclarationSplitRegex.Replace(text, ";\n");

            // 4.2) 去掉尾巴上的自然语言废话（比如 "Done."、"以上代码可以直接复制"）
            text = DropTrailingProse(text);

            // 5) 兜底补全 END_DATA_BLOCK / END_FUNCTION
            text = CompleteMissingBlockEnds(text);

            // 6) 兜底补全 BEGIN
            text = CompleteMissingBegin(text);

            // 6.5) 双重保险（兜底）：FC 块里绝对不允许 VAR 声明区，强制改成 VAR_TEMP
            text = ConvertFcVarToVarTemp(text);

            // 7) 整理空行（压缩连续空行）+ 统一 CRLF 结尾
            string[] finalLines = text.Replace("\r\n", "\n").Split('\n');
            StringBuilder finalBuilder = new StringBuilder();
            int blankRun = 0;
            for (int i = 0; i < finalLines.Length; i++)
            {
                string line = finalLines[i].TrimEnd();
                if (line.Length == 0)
                {
                    blankRun++;
                    if (blankRun > 1)
                    {
                        continue;
                    }
                }
                else
                {
                    blankRun = 0;
                }

                finalBuilder.Append(line);
                finalBuilder.Append('\n');
            }

            string normalized = finalBuilder.ToString().TrimEnd('\n');
            return normalized.Replace("\n", Environment.NewLine) + Environment.NewLine;
        }

        /// <summary>
        /// 去掉声明区多余的 # 号：
        ///  - 形如 "#StartButton : Bool;" 的整行声明（: 后面不是 =，所以不会误伤 body 里的 "#Motor := ..."）
        ///  - VAR ... END_VAR 区间内的 "#变量名" 用法
        /// </summary>
        private static string RemoveHashInDeclarations(string text)
        {
            // 情况 1：声明写法 "#X : Type;"（DB 声明区、VAR 区都可能出现，行首行中都算）
            text = Regex.Replace(text, @"#(?=[A-Za-z_$][\w$]*[ \t]*:(?!=))", string.Empty);

            // 情况 2：VAR ... END_VAR 之间所有 "#标识符"
            StringBuilder builder = new StringBuilder();
            string[] lines = text.Split('\n');
            bool insideVarSection = false;

            for (int i = 0; i < lines.Length; i++)
            {
                string trimmed = lines[i].Trim();

                if (!insideVarSection && Regex.IsMatch(trimmed, @"^VAR(_INPUT|_OUTPUT|_IN_OUT)?\b",
                                                       RegexOptions.IgnoreCase))
                {
                    insideVarSection = true;
                }

                string line = insideVarSection
                    ? Regex.Replace(lines[i], @"#(?=[A-Za-z_$][\w$]*)", string.Empty)
                    : lines[i];

                builder.Append(line);
                if (i < lines.Length - 1)
                {
                    builder.Append('\n');
                }

                if (insideVarSection && Regex.IsMatch(trimmed, @"^END_VAR\b", RegexOptions.IgnoreCase))
                {
                    insideVarSection = false;
                }
            }

            return builder.ToString();
        }

        /// <summary>
        /// 去掉代码尾部的自然语言废话（AI 有时会在块尾追加 "Done." / "以上代码可直接复制"）。
        /// 只从末尾往前删，遇到任何“像 SCL 的行”就停，不会吃掉正常代码。
        /// </summary>
        private static string DropTrailingProse(string text)
        {
            string[] lines = text.Replace("\r\n", "\n").Split('\n');

            int end = lines.Length - 1;
            while (end >= 0 && lines[end].Trim().Length == 0)
            {
                end--;
            }

            while (end >= 0 && !LooksLikeSclLine(lines[end]))
            {
                end--;
                while (end >= 0 && lines[end].Trim().Length == 0)
                {
                    end--;
                }
            }

            if (end < 0)
            {
                return text;
            }

            StringBuilder builder = new StringBuilder();
            for (int i = 0; i <= end; i++)
            {
                if (i > 0)
                {
                    builder.Append('\n');
                }

                builder.Append(lines[i]);
            }

            return builder.ToString();
        }

        private static bool LooksLikeSclLine(string line)
        {
            string trimmed = line.Trim();
            if (trimmed.Length == 0)
            {
                return true;
            }

            if (trimmed.IndexOf(';') >= 0 || trimmed.IndexOf(":=", StringComparison.Ordinal) >= 0)
            {
                return true;
            }

            return Regex.IsMatch(trimmed,
                @"^(//|\{)|(?<!\w)(BEGIN|END_\w+|IF|THEN|ELSE|ELSIF|CASE|FOR|TO|BY|WHILE|REPEAT|UNTIL|DO|VAR\w*|VERSION|DATA_BLOCK|FUNCTION\w*|TYPE|STRUCT)(?!\w)",
                RegexOptions.IgnoreCase);
        }

        /// <summary>缺 END_DATA_BLOCK / END_FUNCTION 时自动补在合适位置。</summary>
        private static string CompleteMissingBlockEnds(string text)
        {
            bool hasDataBlock = Regex.IsMatch(text, @"(?<!\w)DATA_BLOCK(?!\w)", RegexOptions.IgnoreCase);
            bool hasDataBlockEnd = Regex.IsMatch(text, @"(?<!\w)END_DATA_BLOCK(?!\w)", RegexOptions.IgnoreCase);

            if (hasDataBlock && !hasDataBlockEnd)
            {
                // 优先补在 DB 之后那个 FUNCTION 之前，避免把 FC 的代码吞进 DB
                Match functionStart = Regex.Match(text, @"(?<!\w)FUNCTION(?!\w)", RegexOptions.IgnoreCase);
                Match dataBlockStart = Regex.Match(text, @"(?<!\w)DATA_BLOCK(?!\w)", RegexOptions.IgnoreCase);

                if (functionStart.Success && dataBlockStart.Success &&
                    functionStart.Index > dataBlockStart.Index)
                {
                    text = InsertLineAt(text, functionStart.Index, "END_DATA_BLOCK");
                }
                else
                {
                    text = AppendLine(text, "END_DATA_BLOCK");
                }
            }

            bool hasFunction = Regex.IsMatch(text, @"(?<!\w)FUNCTION(?!\w)", RegexOptions.IgnoreCase);
            bool hasFunctionEnd = Regex.IsMatch(text, @"(?<!\w)END_FUNCTION(?!\w)", RegexOptions.IgnoreCase);
            bool hasFunctionBlockEnd = Regex.IsMatch(text, @"(?<!\w)END_FUNCTION_BLOCK(?!\w)",
                                                    RegexOptions.IgnoreCase);

            if (hasFunction && !hasFunctionEnd && !hasFunctionBlockEnd)
            {
                text = AppendLine(text, "END_FUNCTION");
            }

            return text;
        }

        /// <summary>某一块缺 BEGIN 时自动补在插入点上（DB 补在 END_DATA_BLOCK 前，FC 补在最后一个 END_VAR / VERSION 行之后）。</summary>
        private static string CompleteMissingBegin(string text)
        {
            // ---------- DB 段 ----------
            Match dataBlockStart = Regex.Match(text, @"(?<!\w)DATA_BLOCK(?!\w)", RegexOptions.IgnoreCase);
            if (dataBlockStart.Success)
            {
                Match dataBlockEnd = Regex.Match(text.Substring(dataBlockStart.Index),
                                                 @"(?<!\w)END_DATA_BLOCK(?!\w)", RegexOptions.IgnoreCase);
                if (dataBlockEnd.Success)
                {
                    int segmentStart = dataBlockStart.Index;
                    int segmentEnd = segmentStart + dataBlockEnd.Index;
                    string segment = text.Substring(segmentStart, segmentEnd - segmentStart);

                    if (!SclBeginRegex.IsMatch(segment))
                    {
                        // DB 的 BEGIN 紧挨在 END_DATA_BLOCK 前面最稳妥（声明都在它上面）
                        text = InsertLineAt(text, segmentEnd, "BEGIN");
                    }
                }
            }

            // ---------- FC 段 ----------
            Match functionStart = Regex.Match(text, @"(?<!\w)FUNCTION(?!\w)", RegexOptions.IgnoreCase);
            if (functionStart.Success)
            {
                Match functionEnd = Regex.Match(text.Substring(functionStart.Index),
                                                @"(?<!\w)END_FUNCTION(?!\w)", RegexOptions.IgnoreCase);
                if (functionEnd.Success)
                {
                    int segmentStart = functionStart.Index;
                    int segmentEnd = segmentStart + functionEnd.Index;
                    string segment = text.Substring(segmentStart, segmentEnd - segmentStart);

                    if (!SclBeginRegex.IsMatch(segment))
                    {
                        MatchCollection varEnds = SclVarEndRegex.Matches(segment);
                        int insertAt;

                        if (varEnds.Count > 0)
                        {
                            // 有 VAR 段：补在最后一个 END_VAR 之后
                            Match lastVarEnd = varEnds[varEnds.Count - 1];
                            insertAt = segmentStart + lastVarEnd.Index + lastVarEnd.Length;
                        }
                        else
                        {
                            // 没有 VAR 段：补在 VERSION 行之后
                            Match versionLine = Regex.Match(segment, @"VERSION\s*:\s*[0-9.]+",
                                                           RegexOptions.IgnoreCase);
                            insertAt = versionLine.Success
                                ? segmentStart + versionLine.Index + versionLine.Length
                                : segmentStart;
                        }

                        text = InsertLineAt(text, insertAt, "BEGIN");
                    }
                }
            }

            return text;
        }

        /// <summary>
        /// 兜底：FC（FUNCTION）块里绝对不允许 VAR 声明区（否则博图报“区域 VAR 不可用于 FC 类型的块”），
        /// 强制把 VAR 改成 VAR_TEMP。
        /// 只处理 FUNCTION ... END_FUNCTION 片段：FUNCTION_BLOCK / END_FUNCTION_BLOCK 里的 FUNCTION
        /// 因后面紧跟下划线（\w）不会被匹配，所以 FB 里的 VAR 区绝不会被改动。
        /// </summary>
        private static string ConvertFcVarToVarTemp(string text)
        {
            // 整份代码是 FB、且没有任何 FC 片段时，直接返回，避免任何误改
            bool hasFunctionBlock = Regex.IsMatch(text, @"(?<!\w)FUNCTION_BLOCK(?!\w)",
                                                  RegexOptions.IgnoreCase);
            MatchCollection functionStarts = Regex.Matches(
                text, @"(?<!\w)FUNCTION(?!\w)", RegexOptions.IgnoreCase);

            if (hasFunctionBlock && functionStarts.Count == 0)
            {
                return text;
            }

            // 从后往前处理，避免前面的替换影响后面片段的字符下标
            for (int i = functionStarts.Count - 1; i >= 0; i--)
            {
                int keywordIndex = functionStarts[i].Index;

                // 双保险：确认这里不是 FUNCTION_BLOCK 的前缀（正常正则已排除，这里再挡一次）
                if (string.Compare(text, keywordIndex, "FUNCTION_BLOCK", 0, "FUNCTION_BLOCK".Length,
                                   StringComparison.OrdinalIgnoreCase) == 0)
                {
                    continue;
                }

                int segmentStart = functionStarts[i].Index + functionStarts[i].Length;
                int segmentEnd = text.Length;

                Match functionEnd = Regex.Match(text.Substring(segmentStart),
                                                @"(?<!\w)END_FUNCTION(?!\w)", RegexOptions.IgnoreCase);
                if (functionEnd.Success)
                {
                    segmentEnd = segmentStart + functionEnd.Index + functionEnd.Length;
                }

                string segment = text.Substring(segmentStart, segmentEnd - segmentStart);
                string fixedSegment = Regex.Replace(segment, @"(?<!\w)VAR(?!\w)", "VAR_TEMP",
                                                   RegexOptions.IgnoreCase);

                if (!string.Equals(segment, fixedSegment, StringComparison.Ordinal))
                {
                    text = text.Substring(0, segmentStart) + fixedSegment + text.Substring(segmentEnd);
                }
            }

            return text;
        }

        /// <summary>
        /// 读取 AI 在代码第一行写下的选择注释（// 采用现有变量 或 // 采用 FB 块新建变量）。
        /// 返回 null 表示 AI 没写（或清洗时该注释已被裁掉）。
        /// </summary>
        private static string DetectAiBlockChoice(string rawText)
        {
            if (string.IsNullOrEmpty(rawText))
            {
                return null;
            }

            if (rawText.IndexOf("采用 FB 块新建变量", StringComparison.Ordinal) >= 0)
            {
                return "FB 块新建变量（现有变量表不匹配当前需求）";
            }

            if (rawText.IndexOf("采用现有变量", StringComparison.Ordinal) >= 0)
            {
                return "复用现有变量（FC + 全局变量）";
            }

            return null;
        }

        // =================================================================================
        //  DeepSeek 流式请求（保持原有流式逻辑不变）
        // =================================================================================
        private static async Task<string> RequestSclFromDeepSeekAsync(string systemPrompt, string userPrompt,
                                                                     Action<string> onDelta)
        {
            string requestJson = JsonSerializer.Serialize(BuildRequestPayload(systemPrompt, userPrompt, true));

            using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, DeepSeekApiUrl))
            {
                request.Headers.TryAddWithoutValidation("Authorization", "Bearer " + DeepSeekApiKey);
                request.Content = new StringContent(requestJson, Encoding.UTF8, "application/json");

                // ResponseHeadersRead：不把整个响应缓冲到内存，收到响应头就开始处理流
                using (HttpResponseMessage response = await Http.SendAsync(
                           request, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(false))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        string errorBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        throw new HttpRequestException(
                            "DeepSeek 返回 HTTP " + (int)response.StatusCode + " " + response.ReasonPhrase +
                            "\n" + TrimForLog(errorBody, 10, 2000));
                    }

                    using (Stream stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                    using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        StringBuilder builder = new StringBuilder();

                        while (true)
                        {
                            string line = await reader.ReadLineAsync().ConfigureAwait(false);
                            if (line == null)
                            {
                                break;
                            }

                            line = line.Trim();
                            if (line.Length == 0 || !line.StartsWith("data:", StringComparison.Ordinal))
                            {
                                continue;
                            }

                            string payloadText = line.Substring(5).Trim();
                            if (payloadText.Length == 0)
                            {
                                continue;
                            }

                            if (string.Equals(payloadText, "[DONE]", StringComparison.Ordinal))
                            {
                                break;
                            }

                            // 流式增量：choices[0].delta.content
                            string delta = ExtractStreamDelta(payloadText);
                            if (string.IsNullOrEmpty(delta))
                            {
                                continue;
                            }

                            builder.Append(delta);
                            if (onDelta != null)
                            {
                                onDelta(delta);
                            }
                        }

                        return builder.ToString();
                    }
                }
            }
        }

        private static Dictionary<string, object> BuildRequestPayload(string systemPrompt, string userPrompt,
                                                                      bool stream)
        {
            List<object> messages = new List<object>();
            messages.Add(new Dictionary<string, object>
            {
                { "role", "system" },
                { "content", systemPrompt }
            });
            messages.Add(new Dictionary<string, object>
            {
                { "role", "user" },
                { "content", userPrompt }
            });

            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["model"] = DeepSeekModel;
            payload["stream"] = stream;
            payload["temperature"] = 0.2;
            payload["max_tokens"] = 4096;
            payload["messages"] = messages;
            return payload;
        }

        // =================================================================================
        //  JSON 解析（IList 防呆，绝不改回 object[]）
        // =================================================================================
        private static string ExtractStreamDelta(string jsonText)
        {
            try
            {
                object root = JsonSerializer.DeserializeObject(jsonText);
                Dictionary<string, object> rootDict = root as Dictionary<string, object>;
                if (rootDict == null)
                {
                    return null;
                }

                object choicesObject;
                if (!rootDict.TryGetValue("choices", out choicesObject))
                {
                    return null;
                }

                // ★ 防呆：JavaScriptSerializer 把 JSON 数组反序列化成 ArrayList，
                //   直接 (object[])choicesObject 会得到 null（历史上导致“choices 为空”）。
                IList choices = choicesObject as IList;
                if (choices == null || choices.Count == 0)
                {
                    return null;
                }

                Dictionary<string, object> firstChoice = choices[0] as Dictionary<string, object>;
                if (firstChoice == null)
                {
                    return null;
                }

                object deltaObject;
                if (!firstChoice.TryGetValue("delta", out deltaObject))
                {
                    return null;
                }

                Dictionary<string, object> delta = deltaObject as Dictionary<string, object>;
                if (delta == null)
                {
                    return null;
                }

                object contentObject;
                if (!delta.TryGetValue("content", out contentObject))
                {
                    return null;
                }

                return contentObject as string;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string ExtractAssistantContent(string jsonText)
        {
            try
            {
                object root = JsonSerializer.DeserializeObject(jsonText);
                Dictionary<string, object> rootDict = root as Dictionary<string, object>;
                if (rootDict == null)
                {
                    return null;
                }

                object choicesObject;
                if (!rootDict.TryGetValue("choices", out choicesObject))
                {
                    return null;
                }

                // ★ 同样保持 IList 防呆
                IList choices = choicesObject as IList;
                if (choices == null || choices.Count == 0)
                {
                    return null;
                }

                Dictionary<string, object> firstChoice = choices[0] as Dictionary<string, object>;
                if (firstChoice == null)
                {
                    return null;
                }

                object messageObject;
                if (!firstChoice.TryGetValue("message", out messageObject))
                {
                    return null;
                }

                Dictionary<string, object> message = messageObject as Dictionary<string, object>;
                if (message == null)
                {
                    return null;
                }

                object contentObject;
                if (!message.TryGetValue("content", out contentObject))
                {
                    return null;
                }

                return contentObject as string;
            }
            catch (Exception)
            {
                return null;
            }
        }

        // =================================================================================
        //  界面日志 / 文本辅助
        // =================================================================================
        private void AppendLog(string message)
        {
            if (txtLog.IsDisposed)
            {
                return;
            }

            if (txtLog.InvokeRequired)
            {
                txtLog.BeginInvoke(new Action<string>(AppendLog), message);
                return;
            }

            string line = "[" + DateTime.Now.ToString("HH:mm:ss") + "] " + message + Environment.NewLine;
            txtLog.AppendText(line);
            txtLog.SelectionStart = txtLog.TextLength;
            txtLog.ScrollToCaret();
        }

        /// <summary>流式片段：不加时间戳、不加换行，直接拼到日志末尾。</summary>
        private void AppendStreamChunk(string delta)
        {
            if (string.IsNullOrEmpty(delta) || txtLog.IsDisposed)
            {
                return;
            }

            if (txtLog.InvokeRequired)
            {
                txtLog.BeginInvoke(new Action<string>(AppendStreamChunk), delta);
                return;
            }

            txtLog.AppendText(delta);
            txtLog.SelectionStart = txtLog.TextLength;
            txtLog.ScrollToCaret();
        }

        private static HttpClient CreateHttpClient()
        {
            HttpClient client = new HttpClient();
            client.Timeout = TimeSpan.FromMinutes(3);
            client.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "text/event-stream");
            return client;
        }

        /// <summary>把 SCL 里声明的块列出来，用于日志（多块模式下块名不传，仅用于显示）。</summary>
        private static string DescribeSclBlocks(string scl)
        {
            if (string.IsNullOrEmpty(scl))
            {
                return "(空)";
            }

            MatchCollection matches = Regex.Matches(scl,
                "(?<!\\w)(DATA_BLOCK|FUNCTION_BLOCK|FUNCTION)\\s+\"([^\"]+)\"",
                RegexOptions.IgnoreCase);

            if (matches.Count == 0)
            {
                return "(未识别到块声明)";
            }

            StringBuilder builder = new StringBuilder();
            foreach (Match match in matches)
            {
                if (builder.Length > 0)
                {
                    builder.Append("，");
                }

                builder.Append(match.Groups[1].Value.ToUpperInvariant());
                builder.Append(" \"");
                builder.Append(match.Groups[2].Value);
                builder.Append("\"");
            }

            return builder.ToString();
        }

        /// <summary>
        /// 从清洗后的 SCL 中提取“主块名”，用于生成唯一的 .scl 文件名 / 外部源名，避免覆盖博图里已有的程序块。
        /// 优先级：FUNCTION_BLOCK &gt; FUNCTION &gt; DATA_BLOCK（有 FC/FB 时优先返回 FC/FB 的名字）。
        /// </summary>
        private static string ExtractBlockName(string scl)
        {
            if (string.IsNullOrEmpty(scl))
            {
                return null;
            }

            Match functionBlock = Regex.Match(scl,
                "(?<!\\w)FUNCTION_BLOCK\\s+\"([^\"]+)\"", RegexOptions.IgnoreCase);
            if (functionBlock.Success)
            {
                return functionBlock.Groups[1].Value;
            }

            Match function = Regex.Match(scl,
                "(?<!\\w)FUNCTION(?!\\w)\\s+\"([^\"]+)\"", RegexOptions.IgnoreCase);
            if (function.Success)
            {
                return function.Groups[1].Value;
            }

            Match dataBlock = Regex.Match(scl,
                "(?<!\\w)DATA_BLOCK\\s+\"([^\"]+)\"", RegexOptions.IgnoreCase);
            if (dataBlock.Success)
            {
                return dataBlock.Groups[1].Value;
            }

            return null;
        }

        /// <summary>提取 SCL 中所有块名（含类型），用于在日志里确认 AI 是否真的起了新名字。</summary>
        private static List<string> ExtractAllBlockNames(string scl)
        {
            List<string> names = new List<string>();
            if (string.IsNullOrEmpty(scl))
            {
                return names;
            }

            MatchCollection matches = Regex.Matches(scl,
                "(?<!\\w)(DATA_BLOCK|FUNCTION_BLOCK|FUNCTION)\\s+\"([^\"]+)\"",
                RegexOptions.IgnoreCase);

            foreach (Match match in matches)
            {
                names.Add(match.Groups[1].Value.ToUpperInvariant() + " \"" + match.Groups[2].Value + "\"");
            }

            return names;
        }

        private static string InsertLineAt(string text, int index, string line)
        {
            if (index < 0)
            {
                index = 0;
            }

            if (index > text.Length)
            {
                index = text.Length;
            }

            StringBuilder builder = new StringBuilder();
            builder.Append(text, 0, index);

            if (index > 0 && text[index - 1] != '\n')
            {
                builder.Append('\n');
            }

            builder.Append(line);

            if (index < text.Length && text[index] != '\n')
            {
                builder.Append('\n');
            }

            builder.Append(text, index, text.Length - index);
            return builder.ToString();
        }

        private static string AppendLine(string text, string line)
        {
            string trimmed = text.TrimEnd();
            return trimmed + "\n" + line + "\n";
        }

        private static string TrimForLog(string text, int maxLines, int maxChars)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            string trimmed = text.Trim();
            string[] lines = trimmed.Replace("\r\n", "\n").Split('\n');
            if (lines.Length > maxLines)
            {
                StringBuilder builder = new StringBuilder();
                for (int i = 0; i < maxLines; i++)
                {
                    builder.AppendLine(lines[i]);
                }

                builder.AppendLine("...（共 " + lines.Length + " 行，已截断）");
                trimmed = builder.ToString();
            }

            if (trimmed.Length > maxChars)
            {
                trimmed = trimmed.Substring(0, maxChars) + "...（已截断）";
            }

            return trimmed;
        }

        private static int CountLines(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            return text.Replace("\r\n", "\n").Split('\n').Length;
        }
    }
}