// =====================================================================================
//  Form1.cs —— TIA SCL 生成导入助手（V0.4）
//  -------------------------------------------------------------------------------------
//  V0.4 功能清单：
//    1.【Excel 整表一次性注入】NPOI 读整张表 → Markdown 表格（序号/变量名/地址/类型/描述），只调 1 次 DeepSeek
//    2.【PLC 默认变量表自动建变量】按 Excel 顺序建变量：合法名照用，纯数字/非法名 → Tag_1、Tag_2 …；
//       同时生成 物理地址 → 变量名 映射（%I0.0 → "Tag_1"）交给 AI
//    3.【双块架构】固定生成 DATA_BLOCK "DB_Variables"（变量名照抄 Excel）+ FUNCTION "FC_MainControl" : Void
//    4.【变量名直接映射】FC 里禁止绝对地址、禁止 VAR_TEMP，直接写 "DB_Variables"."111" := "Tag_1";
//    5.【清洗兜底】FB → FC、去 Tag_/Var 前缀、非法变量名补双引号、绝对地址按映射表换成变量名
//    6.【预览 + 自动纠错】导入前显示清洗后的 SCL；编译失败把报错交给 DeepSeek 修复，最多 3 轮
//    7.【API Key 独立保存】界面输入 → config.txt（UTF-8 无 BOM），不写死在代码里
//
//  历史版本：V0.1 窗体 + Openness 闭环；V0.2 读取变量表 + 流式生成；V0.3 动态块名 + FC 禁 VAR。
// =====================================================================================

using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace PLC_UI
{
    public partial class Form1 : Form
    {
        // -----------------------------------------------------------------------------
        //  配置区（改这里）
        // -----------------------------------------------------------------------------
        // API Key 不再写死在代码里：界面上输入后保存到程序目录的 config.txt（见 LoadApiKey / SaveApiKey）
        private const string ApiKeyConfigFileName = "config.txt";
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

        // V0.4：自动纠错最多修复 3 轮（首次生成不算）
        private const int MaxAutoFixRounds = 3;

        /// <summary>批量导入时 Excel 里的原始变量名（用于把 AI 私自加的 Var 前缀映射回去）。</summary>
        private List<string> _batchOriginalVariableNames;

        /// <summary>批量导入时 物理地址 → 本次在 PLC 默认变量表里创建好的变量名（如 %I0.0 → Tag_1）。</summary>
        private Dictionary<string, string> _batchAddressToTagName;

        // V0.4：运行时动态创建的代码预览控件
        private TextBox txtPreview;
        private CheckBox chkAutoFix;

        // V0.4：API Key 输入与保存控件（运行时动态创建）
        private TextBox txtApiKey;
        private Button btnSaveKey;

        // V0.5：批量导入 Excel 按钮（运行时动态创建）
        private Button btnImportExcel;

        public Form1()
        {
            InitializeComponent();
            CreatePreviewControls();
            CreateApiKeyControls();
            CreateImportExcelControls();
            Load += Form1_Load;
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
        //  V0.4 功能二：运行时创建代码预览区（txtPreview）
        // =================================================================================
        /// <summary>在窗体底部动态创建“代码预览 + 自动修复开关”区域（不依赖 Designer，避免冲突）。</summary>
        private void CreatePreviewControls()
        {
            Panel previewPanel = new Panel();
            previewPanel.Name = "previewPanel";
            previewPanel.Dock = DockStyle.Bottom;
            previewPanel.Height = 240;
            previewPanel.Padding = new Padding(4, 2, 4, 4);

            // 预览框（先添加 → 最后参与停靠，填满剩余空间）
            txtPreview = new TextBox();
            txtPreview.Name = "txtPreview";
            txtPreview.Multiline = true;
            txtPreview.ReadOnly = true;
            txtPreview.ScrollBars = ScrollBars.Both;
            txtPreview.WordWrap = false;
            txtPreview.Dock = DockStyle.Fill;
            txtPreview.Font = new Font("Consolas", 9F);
            txtPreview.BackColor = Color.FromArgb(30, 30, 30);
            txtPreview.ForeColor = Color.Gainsboro;

            // 自动修复开关（默认勾选：导入失败时自动让 AI 修复）
            chkAutoFix = new CheckBox();
            chkAutoFix.Name = "chkAutoFix";
            chkAutoFix.Text = "编译失败时自动让 AI 修复（最多 " + MaxAutoFixRounds + " 轮；取消勾选则每轮询问）";
            chkAutoFix.Checked = true;
            chkAutoFix.Dock = DockStyle.Top;
            chkAutoFix.Height = 24;

            // 标题
            Label lblPreview = new Label();
            lblPreview.Name = "lblPreview";
            lblPreview.Text = "③ SCL 预览（导入博图前显示，可人工核对）";
            lblPreview.Dock = DockStyle.Top;
            lblPreview.Height = 20;
            lblPreview.TextAlign = ContentAlignment.MiddleLeft;

            previewPanel.Controls.Add(txtPreview);
            previewPanel.Controls.Add(chkAutoFix);
            previewPanel.Controls.Add(lblPreview);

            Controls.Add(previewPanel);
        }

        /// <summary>把当前要导入的 SCL 显示到预览框（每次导入前/修复后都会刷新）。</summary>
        private void DisplaySclPreview(string scl)
        {
            if (txtPreview == null || txtPreview.IsDisposed)
            {
                return;
            }

            txtPreview.Text = scl ?? string.Empty;
            txtPreview.SelectionStart = 0;
            txtPreview.SelectionLength = 0;
        }

        /// <summary>是否继续自动修复：勾选了 chkAutoFix 就直接继续，否则询问用户。</summary>
        private bool ShouldContinueAutoFix(int nextRound)
        {
            if (chkAutoFix != null && chkAutoFix.Checked)
            {
                return true;
            }

            DialogResult answer = MessageBox.Show(this,
                "第 " + nextRound + " 轮自动修复：编译再次失败。\n\n是否让 DeepSeek 修复这段 SCL 并重新导入？\n" +
                "（选“否”将中断流程，保留预览区里的代码供你手动修改）",
                "自动纠错", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            return answer == DialogResult.Yes;
        }

        // =================================================================================
        //  V0.5 功能：批量导入 Excel（变量名 / 数据类型 / 功能描述）→ 逐行生成 SCL → 汇总导入
        // =================================================================================
        /// <summary>Excel 里的一行点位定义。</summary>
        private sealed class VariableRow
        {
            public string Name;                 // 变量名
            public string Address;              // IO 地址（如 i0.0 / Q0.0 / %I0.0）
            public string DataType;             // 数据类型（缺失时已被防呆为 Bool）
            public string Description;          // 功能描述
            public int RowNumber;               // Excel 行号（用于日志）
            public bool DataTypeWasMissing;     // Excel 里这一行是否没填数据类型（防呆标记）
            public bool DataTypeColumnPresent;  // Excel 表头里是否存在“数据类型”列
        }

        /// <summary>导入 + 自动修复的结果容器（批量导入使用）。</summary>
        private sealed class ImportOutcome
        {
            public TiaOperationResult Result;
            public string Scl;
            public string BlockName;
            public int RepairRounds;
            public int ImportAttempts;
        }

        /// <summary>Excel 单元格取文本用的格式化器（能正确处理数字、日期、公式结果）。</summary>
        private static readonly DataFormatter ExcelCellFormatter = new DataFormatter();

        /// <summary>在左侧面板最下方动态创建“批量导入 Excel”按钮（不改 Designer，避免冲突）。</summary>
        private void CreateImportExcelControls()
        {
            btnImportExcel = new Button();
            btnImportExcel.Name = "btnImportExcel";
            btnImportExcel.Text = "批量导入 Excel 生成代码";
            btnImportExcel.Dock = DockStyle.Fill;
            btnImportExcel.Height = 34;
            btnImportExcel.MinimumSize = new Size(0, 34);
            btnImportExcel.Margin = new Padding(0, 4, 0, 0);
            btnImportExcel.Click += btnImportExcel_Click;

            tableLayoutPanel1.RowCount = tableLayoutPanel1.RowCount + 1;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tableLayoutPanel1.Controls.Add(btnImportExcel, 0, tableLayoutPanel1.RowCount - 1);
        }

        /// <summary>点击“批量导入 Excel”：选文件 → 读整张变量表 → 只调 1 次 DeepSeek 生成 DB+逻辑块 → 导入博图。</summary>
        private async void btnImportExcel_Click(object sender, EventArgs e)
        {
            string apiKey = txtApiKey != null && txtApiKey.Text != null ? txtApiKey.Text.Trim() : string.Empty;
            if (apiKey.Length == 0)
            {
                MessageBox.Show(this,
                    "请先在左下角的“API Key”输入框填写 DeepSeek API Key，并点“保存Key”。",
                    "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtApiKey.Focus();
                return;
            }

            string excelPath = string.Empty;
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "选择变量清单 Excel（列：序号 / 变量名 / 地址 / 描述，可选数据类型）";
                dialog.Filter = "Excel 文件 (*.xlsx;*.xls)|*.xlsx;*.xls|Excel 2007 及以上 (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls";
                dialog.CheckFileExists = true;
                dialog.CheckPathExists = true;
                dialog.Multiselect = false;

                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                excelPath = dialog.FileName;
            }

            try
            {
                await RunBatchImportAsync(excelPath);
            }
            catch (Exception ex)
            {
                AppendLog("❌ 批量导入异常：" + ex.Message);
                MessageBox.Show(this, ex.Message, "批量导入失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 批量导入主流程（全表一次性注入）：Excel 整表 → Markdown 表格 → 只调 1 次 DeepSeek
        /// → CleanSclText 清洗 → 预览 → 导入编译（失败自动修复，最多 3 轮）。
        /// </summary>
        private async Task RunBatchImportAsync(string excelPath)
        {
            btnImportExcel.Enabled = false;
            btnGenerate.Enabled = false;
            Cursor = Cursors.WaitCursor;
            txtLog.Clear();
            DisplaySclPreview(string.Empty);

            try
            {
                AppendLog("==================== 批量导入：Excel → DeepSeek → TIA Portal ====================");
                AppendLog("[1/6] 读取 Excel：" + excelPath);

                List<VariableRow> rows = ReadVariableRows(excelPath);
                _batchOriginalVariableNames = GetOriginalVariableNames(rows);
                AppendLog("      读取到 " + rows.Count + " 个变量（整表一次性交给 AI，不再逐行调用）。");
                if (rows.Count == 0)
                {
                    MessageBox.Show(this,
                        "Excel 里没有读到有效数据。\n\n请确认：\n1) 第一行是表头：序号 / 变量名 / 地址 / 描述（可选：数据类型）；\n2) 至少有 1 行数据。",
                        "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int previewCount = Math.Min(rows.Count, 10);
                for (int i = 0; i < previewCount; i++)
                {
                    AppendLog("        - 第 " + rows[i].RowNumber + " 行：" + rows[i].Name +
                              " [地址 " + FormatAddress(rows[i].Address) + "] " + rows[i].Description);
                }

                if (rows.Count > previewCount)
                {
                    AppendLog("        ...（其余 " + (rows.Count - previewCount) + " 行）");
                }

                // 统计“数据类型”缺失的行数，给用户一个总结性反馈（防呆数量）
                int missingTypeCount = 0;
                for (int i = 0; i < rows.Count; i++)
                {
                    if (rows[i].DataTypeWasMissing)
                    {
                        missingTypeCount++;
                    }
                }

                if (missingTypeCount > 0)
                {
                    AppendLog("⚠ [防呆] 本次导入共有 " + missingTypeCount + " 个变量缺失类型，已全部进行防呆处理（临时按 Bool 兜底，并交给 AI 结合功能描述重新推断）。");
                }
                else
                {
                    AppendLog("      √ 所有变量都填写了数据类型，无需防呆。");
                }

                IProgress<string> progress = new Progress<string>(AppendLog);

                // ---- [2/6] 先把 Excel 点位创建到 PLC 的默认变量表（顺序命名 + 地址映射）----
                AppendLog("[2/6] 在 PLC 默认变量表里创建变量（顺序命名 Tag_1…Tag_N，并收集 地址 → 变量名 映射）...");
                List<TiaTagDefinition> tagDefinitions;
                Dictionary<string, string> ioMapping = BuildTagDefinitions(rows, out tagDefinitions);
                _batchAddressToTagName = ioMapping;
                AppendLog("      准备创建 " + tagDefinitions.Count + " 个 PLC 变量。");

                TiaTagCreationResult tagResult = await _tiaService.CreateTagsAsync(tagDefinitions, progress);
                if (tagResult == null)
                {
                    throw new InvalidOperationException("创建 PLC 变量失败：TiaService 没有返回结果。");
                }

                if (!tagResult.Succeeded)
                {
                    AppendLog("⚠ [变量表] 创建变量没有完全成功：" + (tagResult.FailureMessage ?? "未知原因"));
                }

                AppendLog("      地址 → 变量名 映射（共 " + ioMapping.Count + " 条，AI 将用右边的变量名写 FC）：");
                foreach (KeyValuePair<string, string> pair in ioMapping)
                {
                    AppendLog("        " + pair.Key + "  →  \"" + pair.Value + "\"");
                }

                AppendLog("[3/6] 拼装全表 Markdown 变量表 + 变量名映射 + 单次 User Prompt（全表一次性注入）...");
                string systemPrompt = BuildBatchSystemPrompt();
                string userPrompt = BuildBatchUserPrompt(rows, ioMapping);
                AppendLog("      变量表：" + rows.Count + " 个变量；User Prompt 字符数：" + userPrompt.Length);

                AppendLog("---------------------- User Prompt 预览（前 40 行）----------------------");
                string[] userPromptLines = userPrompt.Replace("\r\n", "\n").Split('\n');
                int showLines = Math.Min(userPromptLines.Length, 40);
                for (int i = 0; i < showLines; i++)
                {
                    AppendLog("  | " + userPromptLines[i]);
                }

                if (userPromptLines.Length > showLines)
                {
                    AppendLog("  | ...（共 " + userPromptLines.Length + " 行）");
                }

                AppendLog("[4/6] 调用 DeepSeek（全表一次性生成，流式输出，只调用 1 次）...");
                string rawText = await RequestSclFromDeepSeekAsync(systemPrompt, userPrompt, AppendStreamChunk);
                AppendLog("");
                if (string.IsNullOrWhiteSpace(rawText))
                {
                    throw new InvalidOperationException("DeepSeek 没有返回任何内容，请检查 API Key、网络或余额。");
                }

                AppendLog("[5/6] 清洗 SCL → 预览 → 识别块名 ...");
                string combinedScl = CleanSclText(rawText, true, _batchOriginalVariableNames, _batchAddressToTagName);
                if (string.IsNullOrWhiteSpace(combinedScl))
                {
                    throw new InvalidOperationException("AI 返回的内容里没有可用的 SCL，请重试。");
                }

                DisplaySclPreview(combinedScl);
                List<string> blockNames = ExtractAllBlockNames(combinedScl);
                AppendLog("      识别到块：" + DescribeSclBlocks(combinedScl));
                AppendLog("      块清单：" + (blockNames.Count == 0 ? "(无)" : string.Join("、", blockNames.ToArray())));
                string blockName = ExtractBlockName(combinedScl);
                AppendLog("      导入用主块名：" + (blockName ?? "(自动识别)"));
                AppendLog("      SCL 共 " + CountLines(combinedScl) + " 行。");

                AppendLog("[6/6] 导入博图 → 生成块 → 编译（失败自动修复，最多 " + MaxAutoFixRounds + " 轮）...");
                ImportOutcome outcome = await ImportSclWithAutoFixAsync(combinedScl, blockName, progress);

                if (outcome.Result != null && outcome.Result.Succeeded)
                {
                    AppendLog("✅ 批量导入完成：编译状态 " + outcome.Result.CompileState +
                              "，警告 " + outcome.Result.WarningCount +
                              (outcome.RepairRounds > 0 ? "（自动修复 " + outcome.RepairRounds + " 轮后成功）" : "（首次尝试即成功）") + "。");

                    MessageBox.Show(this,
                        "批量导入成功！\n\n" +
                        "Excel 变量数：" + rows.Count + "\n" +
                        "生成的块数：" + blockNames.Count + "（" + DescribeSclBlocks(combinedScl) + "）\n" +
                        "主块名：" + (outcome.BlockName ?? "(自动识别)") + "\n" +
                        "编译状态：" + outcome.Result.CompileState + "\n" +
                        "警告数：" + outcome.Result.WarningCount + "\n" +
                        "SCL 文件：" + outcome.Result.SclFilePath,
                        "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    string lastError = outcome.Result == null
                        ? "（没有拿到导入结果）"
                        : (outcome.Result.FailureMessage ?? "未知错误");

                    AppendLog("❌ 批量导入失败：已尝试导入 " + outcome.ImportAttempts +
                              " 次、自动修复 " + outcome.RepairRounds + " 轮。");

                    MessageBox.Show(this,
                        "批量导入失败，已尝试自动修复 " + outcome.RepairRounds + " 轮。\n\n" +
                        "最后一次错误：" + lastError + "\n\n" +
                        "预览区保留了最后一次代码，可手动修改后重试。",
                        "导入/编译失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            finally
            {
                btnImportExcel.Enabled = true;
                btnGenerate.Enabled = true;
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// 全表 System Prompt（V0.4：DB 照抄 Excel 变量名 + FC 内用“已创建的变量表映射”直接赋值）。
        /// </summary>
        private static string BuildBatchSystemPrompt()
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("你是一个冷酷的 PLC 代码生成器。请根据我提供的 IO 表，生成 DB 块和 FC 块。");
            builder.AppendLine("【铁律 1】：DB 块中的变量名 100% 照抄 Excel 中的【变量名】（如 `\"111\"`）。");
            builder.AppendLine("【铁律 2】：在 FC 块中，绝对禁止使用任何绝对物理地址（如 %I0.0），也绝对禁止生成 VAR_TEMP 临时变量。");
            builder.AppendLine("【铁律 3】：请直接使用我提供的【已创建的变量表映射】（如 `%I0.0 映射为 \"Tag_1\"`），进行直接赋值！");
            builder.AppendLine("【代码格式要求】：");
            builder.AppendLine("- 输入信号：`\"DB_Variables\".\"111\" := \"Tag_1\";`");
            builder.AppendLine("- 输出信号：`\"Tag_2\" := \"DB_Variables\".\"222\";`");
            builder.AppendLine("【块类型】：必须是 `FUNCTION \"FC_MainControl\" : Void`，结尾 `END_FUNCTION`。绝对禁止用 FB。只输出代码！");
            builder.AppendLine();
            builder.AppendLine("【技术格式（否则博图导入失败，必须遵守）】");
            builder.AppendLine("- 每个块都要单独一行写：{ S7_Optimized_Access := 'TRUE' } 和 VERSION : 0.1；");
            builder.AppendLine("- DB 块结构：DATA_BLOCK \"DB_Variables\" / 属性行 / VERSION : 0.1 / VAR / 变量声明 / END_VAR / BEGIN / END_DATA_BLOCK；");
            builder.AppendLine("- 关键字各占一行：DATA_BLOCK、FUNCTION、VERSION、VAR、END_VAR、BEGIN、END_DATA_BLOCK、END_FUNCTION；");
            builder.AppendLine("- FC 里只允许出现两种东西：\"DB_Variables\" 的变量（带双引号的变量名）和【已创建的变量表映射】里的变量名（带双引号）；");
            builder.AppendLine("- 不要输出 Markdown 围栏（```），中文只允许出现在 // 注释里。");
            return builder.ToString();
        }

        /// <summary>
        /// 把整张 Excel 变量表拼成【一次性】的 User Prompt：
        /// Markdown 表格（序号 / 变量名 / 地址 / 数据类型 / 描述）+ 已创建的 PLC 变量名映射表。
        /// 全表只调用 1 次 DeepSeek，避免逐行生成导致多个 FB 与脑补变量。
        /// </summary>
        private static string BuildBatchUserPrompt(List<VariableRow> rows,
                                                   IDictionary<string, string> ioMapping)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("请根据以下 IO 变量表生成 PLC 控制程序：");
            builder.AppendLine();
            builder.AppendLine("| 序号 | 变量名 | 地址 | 数据类型 | 描述 |");
            builder.AppendLine("|---|---|---|---|---|");

            for (int i = 0; i < rows.Count; i++)
            {
                VariableRow row = rows[i];
                string name = string.IsNullOrWhiteSpace(row.Name) ? "Unnamed" : row.Name.Trim();
                string address = FormatSclAddress(row.Address);
                string dataType = string.IsNullOrWhiteSpace(row.DataType) ? "Bool" : row.DataType.Trim();
                string description = string.IsNullOrWhiteSpace(row.Description)
                    ? (name + " 的控制逻辑")
                    : row.Description.Trim();
                description = description.Replace("|", "/").Replace("\r", " ").Replace("\n", " ");

                builder.AppendLine("| " + (i + 1) + " | " + name + " | " + address + " | " + dataType + " | " + description + " |");
            }

            builder.AppendLine();
            builder.AppendLine("【已创建的变量表映射】（FC 里必须直接用右边的变量名，不能再写 %I0.0 这类绝对地址）：");
            builder.AppendLine("| 物理地址 | 已创建的 PLC 变量名 |");
            builder.AppendLine("|---|---|");

            if (ioMapping != null)
            {
                foreach (KeyValuePair<string, string> pair in ioMapping)
                {
                    builder.AppendLine("| " + pair.Key + " | \"" + pair.Value + "\" |");
                }
            }

            builder.AppendLine();
            builder.AppendLine("记住：DB 块里的变量名照抄上面表格的【变量名】列（如 \"111\"）；FC 块里做映射时必须使用【已创建的 PLC 变量名】列（如 \"Tag_1\"）。");
            builder.AppendLine("例如输入映射写：\"DB_Variables\".\"111\" := \"Tag_1\";；输出映射写：\"Tag_2\" := \"DB_Variables\".\"222\";");
            builder.AppendLine("请严格按照 System Prompt 的约束，生成 DATA_BLOCK \"DB_Variables\" 和 FUNCTION \"FC_MainControl\" : Void 两个块。");
            return builder.ToString();
        }

        /// <summary>
        /// 把 Excel 里的地址统一成 SCL 绝对地址写法：i0.0 → %I0.0，Q0.0 → %Q0.0，%m10.1 → %M10.1。
        /// 无法识别的地址（如 DB1.DBX0.0）原样返回，交给 AI 自行判断。
        /// </summary>
        private static string FormatSclAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
            {
                return "-";
            }

            string text = address.Trim();

            // 已经是 %I0.0 / %Q0.0 这种写法：只把前缀字母统一成大写
            Match percentMatch = Regex.Match(text, @"^%\s*([A-Za-z])\s*(.+)$");
            if (percentMatch.Success)
            {
                return "%" + percentMatch.Groups[1].Value.ToUpperInvariant()
                    + percentMatch.Groups[2].Value.Trim();
            }

            // i0.0 / q0.0 / m0.0 / I 0.0 等：补 % 前缀并大写字母
            Match plainMatch = Regex.Match(text, @"^([IQM])\s*([0-9].*)$", RegexOptions.IgnoreCase);
            if (plainMatch.Success)
            {
                return "%" + plainMatch.Groups[1].Value.ToUpperInvariant()
                    + plainMatch.Groups[2].Value.Trim();
            }

            return text;
        }

        /// <summary>地址为空时统一显示“-”，保证 Markdown 表格列数对齐。</summary>
        private static string FormatAddress(string address)
        {
            return string.IsNullOrWhiteSpace(address) ? "-" : address.Trim();
        }

        /// <summary>取出 Excel 原始变量名列表（清洗时用于把 AI 私自加的 Var 前缀映射回原变量名）。</summary>
        private static List<string> GetOriginalVariableNames(List<VariableRow> rows)
        {
            List<string> names = new List<string>();
            if (rows == null)
            {
                return names;
            }

            for (int i = 0; i < rows.Count; i++)
            {
                if (rows[i] != null && !string.IsNullOrWhiteSpace(rows[i].Name))
                {
                    names.Add(rows[i].Name.Trim());
                }
            }

            return names;
        }

        // =================================================================================
        //  变量表创建（V0.4）：Excel 点位 → 默认变量表里的变量（顺序命名 Tag_1..Tag_N）
        // =================================================================================
        /// <summary>
        /// 根据 Excel 行生成“要创建到 PLC 默认变量表”的变量定义，并返回 物理地址 → 最终变量名 的映射。
        /// 命名规则：Excel 变量名合法（字母/下划线/中文开头，只含字母数字下划线中文）就直接用；
        /// 纯数字（111）或含非法字符时按顺序命名为 Tag_1、Tag_2 ……（每处理一个变量，计数器 +1）。
        /// </summary>
        private static Dictionary<string, string> BuildTagDefinitions(List<VariableRow> rows,
                                                                      out List<TiaTagDefinition> definitions)
        {
            definitions = new List<TiaTagDefinition>();
            Dictionary<string, string> ioMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (rows == null || rows.Count == 0)
            {
                return ioMapping;
            }

            HashSet<string> usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int tagCounter = 1;

            for (int i = 0; i < rows.Count; i++)
            {
                VariableRow row = rows[i];
                if (row == null || string.IsNullOrWhiteSpace(row.Address))
                {
                    // 没有物理地址的行在 ReadVariableRows 里已经跳过，这里再兜一层
                    continue;
                }

                string excelName = (row.Name ?? string.Empty).Trim();
                string finalName = BuildTagName(excelName, ref tagCounter, usedNames);
                string address = FormatSclAddress(row.Address);

                TiaTagDefinition definition = new TiaTagDefinition();
                definition.Name = finalName;
                definition.DataType = string.IsNullOrWhiteSpace(row.DataType) ? "Bool" : row.DataType.Trim();
                definition.Address = address;
                definition.Comment = string.IsNullOrWhiteSpace(row.Description) ? excelName : row.Description.Trim();

                definitions.Add(definition);
                ioMapping[address] = finalName;
            }

            return ioMapping;
        }

        /// <summary>
        /// 生成最终 PLC 变量名：合法名字直接用；纯数字/非法字符用 Tag_&lt;n&gt;；
        /// 与已用名字重复时追加 _2、_3 … 保底。每生成一个就 tagCounter++。
        /// </summary>
        private static string BuildTagName(string excelName, ref int tagCounter, HashSet<string> usedNames)
        {
            string candidate = IsValidTagName(excelName) ? excelName : "Tag_" + tagCounter;

            string unique = candidate;
            int suffix = 2;
            while (usedNames.Contains(unique))
            {
                unique = candidate + "_" + suffix;
                suffix++;
            }

            usedNames.Add(unique);
            tagCounter++;
            return unique;
        }

        /// <summary>
        /// 是否是能直接建到 PLC 变量表里的合法名字：
        /// 字母/下划线/中文开头，只含字母、数字、下划线、中文；纯数字（111）与含空格/符号的名字视为非法。
        /// </summary>
        private static bool IsValidTagName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            return Regex.IsMatch(name, @"^[A-Za-z_\u4e00-\u9fa5][A-Za-z0-9_\u4e00-\u9fa5]*$");
        }

        /// <summary>读取 Excel（支持 .xlsx / .xls）：识别 变量名 / 地址 / 数据类型 / 描述 列，返回所有点位。</summary>
        private List<VariableRow> ReadVariableRows(string excelPath)
        {
            if (string.IsNullOrWhiteSpace(excelPath))
            {
                throw new InvalidOperationException("没有选择 Excel 文件。");
            }

            if (!File.Exists(excelPath))
            {
                throw new FileNotFoundException("Excel 文件不存在。", excelPath);
            }

            IWorkbook workbook;
            using (FileStream stream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                bool isXlsx = excelPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                              || excelPath.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase);

                workbook = isXlsx ? (IWorkbook)new XSSFWorkbook(stream) : (IWorkbook)new HSSFWorkbook(stream);
            }

            ISheet sheet = workbook.GetSheetAt(0);
            if (sheet == null)
            {
                throw new InvalidOperationException("Excel 中没有工作表，请确认文件不是空的。");
            }

            IRow headerRow = sheet.GetRow(0);
            int columnName = FindColumnIndex(headerRow, "变量名", "变量", "名称", "点位", "name", "tag");
            int columnAddress = FindColumnIndex(headerRow, "地址", "io地址", "io 地址", "address", "adress", "absoluteaddress");
            int columnType = FindColumnIndex(headerRow, "数据类型", "类型", "type", "datatype");
            int columnDescription = FindColumnIndex(headerRow, "功能描述", "描述", "说明", "功能", "description", "comment");

            bool dataTypeColumnPresent = columnType >= 0;

            int firstDataRow;
            if (columnName >= 0)
            {
                firstDataRow = 1;
                if (columnDescription < 0)
                {
                    // 没找到“描述”列就取“变量名/地址/类型”之后第一个空列，避免把地址列当描述
                    int nextColumn = columnName + 1;
                    if (columnAddress >= 0 && columnAddress >= nextColumn)
                    {
                        nextColumn = columnAddress + 1;
                    }

                    if (columnType >= 0 && columnType >= nextColumn)
                    {
                        nextColumn = columnType + 1;
                    }

                    columnDescription = nextColumn;
                }
            }
            else
            {
                // 没有找到表头 → 退化为“第 1 列变量名 / 第 2 列地址 / 第 3 列描述”，第 1 行就是数据
                columnName = 0;
                columnAddress = 1;
                columnType = -1;
                columnDescription = 2;
                firstDataRow = 0;
            }

            List<VariableRow> rows = new List<VariableRow>();
            for (int rowIndex = firstDataRow; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null)
                {
                    continue;
                }

                string name = GetExcelCellText(row.GetCell(columnName));
                if (name.Length == 0 || name == "#")
                {
                    continue;
                }

                VariableRow item = new VariableRow();
                item.Name = name;

                // 防呆：物理地址为空 → 该行没有可映射的 IO，直接跳过并警告
                string rawAddress = columnAddress >= 0 ? GetExcelCellText(row.GetCell(columnAddress)) : string.Empty;
                if (rawAddress.Length == 0)
                {
                    AppendLog("[警告] 第 " + (rowIndex + 1) + " 行变量 \"" + name + "\" 缺失物理地址，已忽略！");
                    continue;
                }

                // 防呆：地址统一成 SCL 绝对地址（I0.0 → %I0.0；已是 %I0.0 则只规范大小写）
                item.Address = FormatSclAddress(rawAddress);
                item.DataType = columnType >= 0 ? GetExcelCellText(row.GetCell(columnType)) : string.Empty;
                item.Description = GetExcelCellText(row.GetCell(columnDescription));
                item.RowNumber = rowIndex + 1;
                item.DataTypeColumnPresent = dataTypeColumnPresent;

                // 防呆：数据类型为空（null / 空串 / 全空格）→ 强制 Bool，并立刻在日志里醒目警告
                if (item.DataType.Length == 0)
                {
                    item.DataType = "Bool";
                    item.DataTypeWasMissing = dataTypeColumnPresent;

                    if (dataTypeColumnPresent)
                    {
                        AppendLog("[警告] 第 " + item.RowNumber + " 行变量 \"" + item.Name
                                  + "\" 缺失数据类型，已默认按 Bool 类型处理，请核对！");
                    }
                }

                if (item.Description.Length == 0)
                {
                    item.Description = item.Name + " 的控制逻辑";
                }

                rows.Add(item);
            }

            return rows;
        }

        /// <summary>在表头行里按关键字找列号（忽略大小写，支持中英文表头）；找不到返回 -1。</summary>
        private static int FindColumnIndex(IRow headerRow, params string[] keywords)
        {
            if (headerRow == null || keywords == null || keywords.Length == 0)
            {
                return -1;
            }

            int lastCell = headerRow.LastCellNum;
            for (int columnIndex = 0; columnIndex < lastCell; columnIndex++)
            {
                string text = GetExcelCellText(headerRow.GetCell(columnIndex));
                if (text.Length == 0)
                {
                    continue;
                }

                for (int i = 0; i < keywords.Length; i++)
                {
                    if (text.IndexOf(keywords[i], StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return columnIndex;
                    }
                }
            }

            return -1;
        }

        /// <summary>取单元格文本（数字/布尔/公式都能转成字符串）。</summary>
        private static string GetExcelCellText(ICell cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }

            string text = string.Empty;
            try
            {
                text = ExcelCellFormatter.FormatCellValue(cell);
            }
            catch (Exception)
            {
                text = string.Empty;
            }

            return text == null ? string.Empty : text.Trim();
        }

        /// <summary>批量汇总时防止多个块同名（重名会让 Openness 只保留一个块）：冲突的块加 _2、_3 后缀。</summary>
        private List<string> EnsureUniqueBlockNames(List<string> parts)
        {
            Dictionary<string, int> usedNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            List<string> results = new List<string>(parts.Count);

            for (int i = 0; i < parts.Count; i++)
            {
                string part = parts[i];
                List<string> names = ExtractAllBlockNames(part);

                for (int n = 0; n < names.Count; n++)
                {
                    Match match = Regex.Match(names[n], "\"([^\"]+)\"");
                    string blockName = match.Success ? match.Groups[1].Value : names[n];

                    int usedCount;
                    if (!usedNames.TryGetValue(blockName, out usedCount))
                    {
                        usedNames[blockName] = 1;
                        continue;
                    }

                    usedCount++;
                    usedNames[blockName] = usedCount;

                    string newName = blockName + "_" + usedCount;
                    part = Regex.Replace(part, "\"" + Regex.Escape(blockName) + "\"", "\"" + newName + "\"");
                    AppendLog("      ⚠ 块名重复：" + blockName + " → 已改名为 " + newName);
                }

                results.Add(part);
            }

            return results;
        }

        // =================================================================================
        //  V0.5：批量模式合并 DB（每个点位都会生成 DATA_BLOCK "DB_Variables"，这里合并成一个）
        // =================================================================================
        private static readonly Regex DataBlockBlockRegex = new Regex(
            "(?is)DATA_BLOCK\\s+\"([^\"]+)\"(.*?)END_DATA_BLOCK", RegexOptions.Compiled);

        private static readonly Regex DataBlockVarSectionRegex = new Regex(
            "(?is)\\bVAR\\b(.*?)\\bEND_VAR\\b", RegexOptions.Compiled);

        /// <summary>
        /// 把每行生成结果里的 DATA_BLOCK "DB_Variables" 抽出来合并成一个全局 DB（变量按名称去重），
        /// 其余块（FB）保留在 <paramref name="parts"/> 里；FB 若引用了其它 DB 名，统一改成 "DB_Variables"。
        /// 数据类型优先级：表格里明确填写的类型 &gt; AI 推断出的类型（仅对缺失类型的行）&gt; Bool 兜底。
        /// </summary>
        private static string MergeDatabaseBlocks(List<string> parts, List<VariableRow> excelRows)
        {
            Dictionary<string, string> declarations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            List<string> otherDatabaseNames = new List<string>();
            List<string> remainingParts = new List<string>();

            // 1) Excel 里“明确填写了类型”的物理点位优先声明（保证类型与表格一致）
            if (excelRows != null)
            {
                for (int i = 0; i < excelRows.Count; i++)
                {
                    string excelName = excelRows[i].Name == null ? string.Empty : excelRows[i].Name.Trim();
                    if (excelName.Length == 0)
                    {
                        continue;
                    }

                    if (excelRows[i].DataTypeWasMissing)
                    {
                        // 缺失类型的行交给 AI 推断，见步骤 2.5
                        continue;
                    }

                    string excelType = string.IsNullOrWhiteSpace(excelRows[i].DataType) ? "Bool" : excelRows[i].DataType.Trim();
                    if (!declarations.ContainsKey(excelName))
                    {
                        declarations.Add(excelName, excelName + " : " + excelType);
                    }
                }
            }

            // 2) 从每个生成结果里抽出 DATA_BLOCK，合并 VAR 区声明
            for (int i = 0; i < parts.Count; i++)
            {
                string part = parts[i];
                MatchCollection matches = DataBlockBlockRegex.Matches(part);
                if (matches.Count == 0)
                {
                    remainingParts.Add(part);
                    continue;
                }

                StringBuilder stripped = new StringBuilder(part);
                for (int m = matches.Count - 1; m >= 0; m--)
                {
                    Match match = matches[m];
                    string databaseName = match.Groups[1].Value;

                    if (!string.Equals(databaseName, "DB_Variables", StringComparison.OrdinalIgnoreCase) &&
                        !otherDatabaseNames.Contains(databaseName))
                    {
                        otherDatabaseNames.Add(databaseName);
                    }

                    Match varSection = DataBlockVarSectionRegex.Match(match.Value);
                    if (varSection.Success)
                    {
                        AddVariableDeclarations(varSection.Groups[1].Value, declarations);
                    }

                    stripped.Remove(match.Index, match.Length);
                }

                string rest = stripped.ToString().Trim();
                if (rest.Length > 0)
                {
                    remainingParts.Add(rest);
                }
            }

            // 2.5) Excel 兜底：保证每个点位都出现在 DB 里
            //      - 表格里明确填了类型的行：以表格为准（覆盖 AI 的判断）
            //      - 表格里缺失类型的行：优先保留 AI 推断出的类型；AI 没声明时才用 Bool 兜底
            if (excelRows != null)
            {
                for (int i = 0; i < excelRows.Count; i++)
                {
                    string excelName = excelRows[i].Name == null ? string.Empty : excelRows[i].Name.Trim();
                    if (excelName.Length == 0)
                    {
                        continue;
                    }

                    string excelType = string.IsNullOrWhiteSpace(excelRows[i].DataType) ? "Bool" : excelRows[i].DataType.Trim();
                    if (excelRows[i].DataTypeWasMissing)
                    {
                        if (!declarations.ContainsKey(excelName))
                        {
                            declarations.Add(excelName, excelName + " : " + excelType);
                        }
                    }
                    else
                    {
                        declarations[excelName] = excelName + " : " + excelType;
                    }
                }
            }

            // 3) FB 里引用其它 DB 名 → 统一改成合并后的 DB_Variables
            for (int i = 0; i < remainingParts.Count; i++)
            {
                for (int n = 0; n < otherDatabaseNames.Count; n++)
                {
                    remainingParts[i] = Regex.Replace(remainingParts[i],
                        "\"" + Regex.Escape(otherDatabaseNames[n]) + "\"", "\"DB_Variables\"");
                }
            }

            parts.Clear();
            parts.AddRange(remainingParts);

            // 4) 重新拼出唯一的全局 DB
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("DATA_BLOCK \"DB_Variables\"");
            builder.AppendLine("{ S7_Optimized_Access := 'TRUE' }");
            builder.AppendLine("VERSION : 0.1");
            builder.AppendLine("VAR");
            foreach (KeyValuePair<string, string> pair in declarations)
            {
                builder.AppendLine("    " + pair.Value + ";");
            }

            builder.AppendLine("END_VAR");
            builder.AppendLine("BEGIN");
            builder.AppendLine("END_DATA_BLOCK");
            return builder.ToString();
        }

        /// <summary>解析 VAR 区文本里的声明（如 "Cylinder_1 : Bool;"），按变量名去重后并入 declarations。</summary>
        private static void AddVariableDeclarations(string varSectionText, Dictionary<string, string> declarations)
        {
            if (string.IsNullOrWhiteSpace(varSectionText))
            {
                return;
            }

            string[] lines = varSectionText.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (line.Length == 0 || line.StartsWith("//", StringComparison.Ordinal))
                {
                    continue;
                }

                string[] chunks = line.Split(';');
                for (int c = 0; c < chunks.Length; c++)
                {
                    string declaration = chunks[c].Trim();

                    int commentIndex = declaration.IndexOf("//", StringComparison.Ordinal);
                    if (commentIndex >= 0)
                    {
                        declaration = declaration.Substring(0, commentIndex).Trim();
                    }

                    int colon = declaration.IndexOf(':');
                    if (colon <= 0)
                    {
                        continue;
                    }

                    string variableName = declaration.Substring(0, colon).Trim();
                    if (variableName.Length == 0)
                    {
                        continue;
                    }

                    if (!declarations.ContainsKey(variableName))
                    {
                        declarations.Add(variableName, declaration);
                    }
                }
            }
        }

        /// <summary>取出 SCL 里所有块名（用于批量导入前把旧块删掉，避免同名冲突）。</summary>
        private static List<string> GetBlockNamesForDeletion(string scl)
        {
            List<string> result = new List<string>();
            List<string> formattedNames = ExtractAllBlockNames(scl);
            for (int i = 0; i < formattedNames.Count; i++)
            {
                Match match = Regex.Match(formattedNames[i], "\"([^\"]+)\"");
                if (match.Success)
                {
                    string name = match.Groups[1].Value;
                    if (!result.Contains(name))
                    {
                        result.Add(name);
                    }
                }
            }

            return result;
        }

        /// <summary>导入 SCL 并编译；失败时自动让 DeepSeek 修复并重试（最多 MaxAutoFixRounds 轮）。</summary>
        private async Task<ImportOutcome> ImportSclWithAutoFixAsync(string scl, string blockName, IProgress<string> progress)
        {
            ImportOutcome outcome = new ImportOutcome();
            outcome.Scl = scl;
            outcome.BlockName = blockName;
            outcome.RepairRounds = 0;
            outcome.ImportAttempts = 0;

            try
            {
                while (true)
                {
                    outcome.ImportAttempts++;
                    AppendLog("[导入] 第 " + outcome.ImportAttempts + " 次导入博图并编译（主块名：" +
                              (outcome.BlockName ?? "自动识别") + "）...");

                    List<string> blockNamesToDelete = GetBlockNamesForDeletion(outcome.Scl);
                    outcome.Result = await _tiaService.ImportAndCompileAsync(outcome.Scl, outcome.BlockName, progress,
                                                                             blockNamesToDelete);

                    if (outcome.Result.Succeeded)
                    {
                        return outcome;
                    }

                    AppendLog("❌ 第 " + outcome.ImportAttempts + " 次导入失败：" + (outcome.Result.FailureMessage ?? "未知错误"));

                    if (outcome.RepairRounds >= MaxAutoFixRounds)
                    {
                        AppendLog("❌ 自动修复已达上限（" + MaxAutoFixRounds + " 轮），停止尝试。");
                        return outcome;
                    }

                    if (!ShouldContinueAutoFix(outcome.RepairRounds + 1))
                    {
                        AppendLog("⚠ 用户取消了自动修复，流程中断。");
                        return outcome;
                    }

                    outcome.RepairRounds++;
                    AppendLog("============================================================");
                    AppendLog("[自动修复轮次 " + outcome.RepairRounds + "/" + MaxAutoFixRounds + "] 把“失败代码 + 博图报错”交给 DeepSeek 修复 ...");

                    string fixedRaw;
                    try
                    {
                        fixedRaw = await FixSclWithDeepSeekAsync(outcome.Scl, outcome.Result.FailureMessage ?? outcome.Result.CompileState);
                    }
                    catch (Exception fixEx)
                    {
                        AppendLog("❌ 自动修复请求本身出错：" + fixEx.Message);
                        return outcome;
                    }

                    if (string.IsNullOrWhiteSpace(fixedRaw))
                    {
                        AppendLog("❌ DeepSeek 没有返回修复代码，停止自动修复。");
                        return outcome;
                    }

                    string fixedScl = CleanSclText(fixedRaw, true, _batchOriginalVariableNames, _batchAddressToTagName);
                    if (string.IsNullOrWhiteSpace(fixedScl))
                    {
                        AppendLog("❌ 修复结果清洗后为空，停止自动修复。");
                        return outcome;
                    }

                    outcome.Scl = fixedScl;

                    string fixedBlockName = ExtractBlockName(outcome.Scl);
                    if (!string.IsNullOrEmpty(fixedBlockName))
                    {
                        outcome.BlockName = fixedBlockName;
                    }

                    AppendLog("      修复后块：" + DescribeSclBlocks(outcome.Scl));
                    DisplaySclPreview(outcome.Scl);
                }
            }
            catch (Exception ex)
            {
                AppendLog("❌ 导入过程异常：" + ex.Message);
                return outcome;
            }
        }

        // =================================================================================
        //  V0.4：API Key 界面输入 + 本地保存（config.txt）
        // =================================================================================
        /// <summary>在左侧面板最下方动态创建“API Key 输入框 + 保存Key 按钮”（不改 Designer，避免冲突）。</summary>
        private void CreateApiKeyControls()
        {
            TableLayoutPanel apiPanel = new TableLayoutPanel();
            apiPanel.Name = "apiKeyPanel";
            apiPanel.Dock = DockStyle.Fill;
            apiPanel.AutoSize = false;
            apiPanel.Height = 30;
            apiPanel.Margin = new Padding(0, 4, 0, 0);
            apiPanel.ColumnCount = 3;
            apiPanel.RowCount = 1;
            apiPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            apiPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            apiPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            // “API Key：”标签
            Label lblApiKey = new Label();
            lblApiKey.Name = "lblApiKey";
            lblApiKey.Text = "API Key：";
            lblApiKey.AutoSize = true;
            lblApiKey.Anchor = AnchorStyles.Left;
            lblApiKey.TextAlign = ContentAlignment.MiddleLeft;

            // 输入框（密码掩码，保护隐私）
            txtApiKey = new TextBox();
            txtApiKey.Name = "txtApiKey";
            txtApiKey.Dock = DockStyle.Fill;
            txtApiKey.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            txtApiKey.UseSystemPasswordChar = true;
            txtApiKey.Margin = new Padding(0, 3, 4, 3);

            // 保存按钮
            btnSaveKey = new Button();
            btnSaveKey.Name = "btnSaveKey";
            btnSaveKey.Text = "保存Key";
            btnSaveKey.AutoSize = true;
            btnSaveKey.Anchor = AnchorStyles.Right;
            btnSaveKey.Click += btnSaveKey_Click;

            apiPanel.Controls.Add(lblApiKey, 0, 0);
            apiPanel.Controls.Add(txtApiKey, 1, 0);
            apiPanel.Controls.Add(btnSaveKey, 2, 0);

            // 加到左侧 tableLayoutPanel1 的最后一行（最下方），原有三行的行号不变
            tableLayoutPanel1.RowCount = tableLayoutPanel1.RowCount + 1;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tableLayoutPanel1.Controls.Add(apiPanel, 0, tableLayoutPanel1.RowCount - 1);
        }

        /// <summary>窗体加载时把本地保存的 API Key 读回到输入框。</summary>
        private void Form1_Load(object sender, EventArgs e)
        {
            string savedKey = LoadApiKey();
            txtApiKey.Text = savedKey ?? string.Empty;
            if (!string.IsNullOrEmpty(txtApiKey.Text))
            {
                txtApiKey.SelectionStart = txtApiKey.Text.Length;
            }

            txtApiKey.Focus();
        }

        /// <summary>config.txt 的完整路径：程序运行目录（exe 所在目录）。</summary>
        private static string GetConfigFilePath()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ApiKeyConfigFileName);
        }

        /// <summary>读取本地 API Key；config.txt 不存在或读取失败时返回空字符串。</summary>
        private static string LoadApiKey()
        {
            try
            {
                string path = GetConfigFilePath();
                if (!File.Exists(path))
                {
                    return string.Empty;
                }

                string text = File.ReadAllText(path, Encoding.UTF8);
                // 兼容手动用记事本编辑后可能带的 BOM / 行尾空白
                return text.Trim().Trim('\uFEFF').Trim();
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>把 API Key 写入程序目录的 config.txt（UTF-8 无 BOM）。</summary>
        private void SaveApiKey(string key)
        {
            string path = GetConfigFilePath();
            File.WriteAllText(path, (key ?? string.Empty).Trim(), new UTF8Encoding(false));
        }

        /// <summary>点击“保存Key”：校验非空后写入 config.txt 并提示。</summary>
        private void btnSaveKey_Click(object sender, EventArgs e)
        {
            string key = txtApiKey.Text != null ? txtApiKey.Text.Trim() : string.Empty;
            if (key.Length == 0)
            {
                MessageBox.Show(this, "API Key 不能为空。", "提示",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtApiKey.Focus();
                return;
            }

            try
            {
                SaveApiKey(key);
                MessageBox.Show(this, "API Key 已保存。\n\n文件：" + GetConfigFilePath(),
                                "已保存", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "保存失败：" + ex.Message, "保存失败",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =================================================================================
        //  V0.4 功能一：自动纠错 —— 把“失败代码 + 博图报错”交给 DeepSeek 修复
        // =================================================================================
        /// <summary>
        /// 请求 DeepSeek 修复编译失败的 SCL：输入失败代码与博图报错信息，返回修复后的完整 SCL。
        /// 内部复用流式请求（边生成边输出到日志），返回累计文本后由调用方 CleanSclText 清洗。
        /// </summary>
        private async Task<string> FixSclWithDeepSeekAsync(string brokenScl, string errorMessage)
        {
            if (string.IsNullOrEmpty(errorMessage))
            {
                errorMessage = "（博图没有返回具体错误信息，请检查 SCL 语法、块名冲突、FC 里是否误用了 VAR。）";
            }

            StringBuilder system = new StringBuilder();
            system.AppendLine("你是西门子 TIA Portal V19 的 SCL 调试专家，负责修复导入/编译失败的 SCL 代码。");
            system.AppendLine("用户会给你：1) 之前失败的完整 SCL 代码；2) 博图的报错信息。");
            system.AppendLine("你必须输出修复后的、能成功导入并编译的完整 SCL。");
            system.AppendLine();
            system.AppendLine("【最高授权】");
            system.AppendLine("现在你是“修复模式”，不是“服从需求模式”。你有最高权限推翻之前任何不合理的约束：");
            system.AppendLine("- 如果最初的需求（例如“用 TON 但不给背景 DB/实例”“不要声明变量”“直接用定时器名而不给实例”）在 SCL 里根本不可能成立，你必须以“能通过博图编译”为唯一目标进行修改，不要为了遵守旧约束而保留错误代码。");
            system.AppendLine("- 不允许再以“按用户原始要求”为理由拒绝补全变量、定时器实例或声明区；先把代码修到能编译，再谈需求。");
            system.AppendLine("- 如果原块是 FUNCTION（FC）但必须使用定时器实例，允许把块类型升级为同名的 FUNCTION_BLOCK（带 VAR 实例），块名保持不变；这是被授权的修改。");
            system.AppendLine();
            system.AppendLine("【错误分析】");
            system.AppendLine("仔细阅读博图报错信息，逐条分析：");
            system.AppendLine("- 如果报错提示“Tag 未定义 / 未定义 / 找不到变量 / 参数无效 / 不是有效的实例 / 无法解析”等含义，必须（强制）在对应块的声明区中补充定义这些变量：FB/DB 用 VAR ... END_VAR，FC 用 VAR_TEMP ... END_VAR；数据类型必须与用法一致（Bool、Int、Real、Time、TON、TOF 等）。绝对不允许靠删除代码来“绕过”报错。");
            system.AppendLine("- 如果报错指出某行语法或类型不匹配，按该行及上下文修正，不要删除整段逻辑。");
            system.AppendLine("- 如果博图报“行 000004”之类的行号，对应的是这段代码的行号，据此定位真正出错的位置。");
            system.AppendLine();
            system.AppendLine("【定时器铁律】");
            system.AppendLine("如果代码中使用了 TON、TOF、TP、TONR 等定时器：");
            system.AppendLine("- 必须在声明区（FB 用 VAR，FC 用 VAR_TEMP）声明对应的 IEC_TIMER 类型实例变量，例如：MoveTimer : TON; 或 DelayTimer : TOF;");
            system.AppendLine("- 调用时必须把这个实例作为第一个参数传入，例如：\"MoveTimer\"(IN := #Start, PT := T#5S); 或 #MoveTimer(IN := #Start, PT := T#5S);");
            system.AppendLine("- 绝对不能再省略实例变量！TON 不是函数，没有实例必然编译失败；也不允许让多个定时器共享同一个实例反复复用。");
            system.AppendLine();
            system.AppendLine("修复规则（在满足上面三条强硬要求的前提下，必须全部遵守）：");
            system.AppendLine("1. 只输出 SCL 代码，不要 Markdown 围栏，不要任何解释、标题、道歉或多余文字。");
            system.AppendLine("2. 保持原块类型（FUNCTION_BLOCK / FUNCTION / DATA_BLOCK）和块名不变（仅“定时器需要实例”这一种情况允许 FC→同名 FB），不要改名，不要增删无关的块。");
            system.AppendLine("3. FC（FUNCTION）里绝对禁止 VAR 声明区，中间变量只能写 VAR_TEMP ... END_VAR（否则会报“区域 VAR 不可用于 FC 类型的块”）。");
            system.AppendLine("4. 关键字必须各占一行：DATA_BLOCK、FUNCTION_BLOCK、FUNCTION、VERSION、VAR、VAR_INPUT、VAR_OUTPUT、VAR_TEMP、END_VAR、BEGIN、END_DATA_BLOCK、END_FUNCTION、END_FUNCTION_BLOCK。");
            system.AppendLine("5. 原代码里有多个块（例如 DATA_BLOCK + FUNCTION）时，必须一起完整输出，不能只输出其中一个。");
            system.AppendLine("6. 块名/变量名只能用英文字母、数字、下划线，不能以数字开头；中文只允许出现在 // 注释里。");
            system.AppendLine("7. 只改真正出错的地方，不要大改逻辑；能保留的业务逻辑（起保停、联锁、延时等）必须保留。");
            system.AppendLine("8. 修复后的代码必须完整（从块声明开始，到 END_* 结束）。");

            StringBuilder user = new StringBuilder();
            user.AppendLine("代码编译报错了，报错信息如下：");
            user.AppendLine(errorMessage);
            user.AppendLine();
            user.AppendLine("出错的完整 SCL 代码如下（原文）：");
            user.AppendLine(brokenScl ?? string.Empty);
            user.AppendLine();
            user.AppendLine("请特别注意：如果报错涉及“Tag 未定义 / 参数无效 / 定时器没有实例”，必须按【错误分析】和【定时器铁律】强制补全变量声明和 TON/TOF 实例变量，不允许省略。");
            user.AppendLine("请修复代码并输出完整的 SCL。绝对不要输出多余解释！");

            AppendLog("      正在请求 DeepSeek 修复（流式输出）...");
            AppendLog("---------------------------------------- 修复代码开始 ----------------------------------------");

            string fixedText = await RequestSclFromDeepSeekAsync(system.ToString(), user.ToString(), AppendStreamChunk);

            AppendLog("");
            AppendLog("---------------------------------------- 修复代码结束 ----------------------------------------");
            return fixedText;
        }

        // =================================================================================
        //  按钮主流程（V0.4：生成 → 预览 → 导入 → 失败自动修复，最多 3 轮）
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

            // V0.4：API Key 从界面读取（不再写死在代码里）
            string apiKey = txtApiKey != null && txtApiKey.Text != null ? txtApiKey.Text.Trim() : string.Empty;
            if (apiKey.Length == 0)
            {
                MessageBox.Show(this,
                    "请先在左下角的“API Key”输入框填写 DeepSeek API Key，并点“保存Key”。",
                    "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtApiKey.Focus();
                return;
            }

            btnGenerate.Enabled = false;
            btnGenerate.Text = "生成中，请稍候...";
            Cursor = Cursors.WaitCursor;
            txtLog.Clear();
            DisplaySclPreview(string.Empty);

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

                // ---------- 步骤 3：预览 → 导入博图 → 生成块 → 编译 → 失败自动修复 ----------
                IProgress<string> progress = new Progress<string>(AppendLog);
                TiaOperationResult result = null;
                int repairRound = 0;      // 已完成的自动修复轮次（首次生成不算修复）
                int importAttempt = 0;    // 导入尝试次数（1 = 首次，之后每次修复后 +1）

                // 导入前先把代码展示出来，方便人工核对（V0.4 功能二）
                DisplaySclPreview(scl);

                while (true)
                {
                    importAttempt++;

                    AppendLog("[3/3] 第 " + importAttempt + " 次导入博图并编译（主块名：" +
                              (blockName ?? "自动识别") + "）...");

                    // blockName 非 null：TiaService 用它作为 .scl 文件名和外部源名（AI_<块名>），
                    // 因此每次生成不同的程序时不会覆盖上一次的程序块。
                    result = await _tiaService.ImportAndCompileAsync(scl, blockName, progress);

                    if (result.Succeeded)
                    {
                        break;
                    }

                    AppendLog("❌ 第 " + importAttempt + " 次导入失败：" + (result.FailureMessage ?? "未知错误"));

                    // 已经用掉全部自动修复轮次 → 不再修复
                    if (repairRound >= MaxAutoFixRounds)
                    {
                        AppendLog("❌ 自动修复已达上限（" + MaxAutoFixRounds + " 轮），停止尝试。");
                        break;
                    }

                    // 用户可在预览区取消勾选 chkAutoFix → 每轮弹窗询问，可中断
                    if (!ShouldContinueAutoFix(repairRound + 1))
                    {
                        AppendLog("⚠ 用户取消了自动修复，流程中断（预览区保留最后一次代码，可手动修改）。");
                        break;
                    }

                    repairRound++;
                    AppendLog("============================================================");
                    AppendLog("[自动修复轮次 " + repairRound + "/" + MaxAutoFixRounds + "] 把“失败代码 + 博图报错”交给 DeepSeek 修复 ...");

                    string fixedRaw = null;
                    try
                    {
                        fixedRaw = await FixSclWithDeepSeekAsync(scl, result.FailureMessage ?? result.CompileState);
                    }
                    catch (Exception fixEx)
                    {
                        AppendLog("❌ 自动修复请求本身出错：" + fixEx.Message);
                        break;
                    }

                    if (string.IsNullOrWhiteSpace(fixedRaw))
                    {
                        AppendLog("❌ DeepSeek 没有返回修复代码，停止自动修复。");
                        break;
                    }

                    string fixedScl = CleanSclText(fixedRaw);
                    if (string.IsNullOrWhiteSpace(fixedScl))
                    {
                        AppendLog("❌ 修复结果清洗后为空，停止自动修复。");
                        break;
                    }

                    if (string.Equals(fixedScl, scl, StringComparison.Ordinal))
                    {
                        AppendLog("⚠ AI 返回的代码与失败代码完全相同，下一轮很可能仍然失败。");
                    }

                    scl = fixedScl;

                    // 若 AI 在修复过程中改了块名，跟随新块名（否则 TiaService 仍用旧名字命名文件/外部源）
                    string fixedBlockName = ExtractBlockName(scl);
                    if (!string.IsNullOrEmpty(fixedBlockName))
                    {
                        blockName = fixedBlockName;
                    }

                    AppendLog("      修复后块：" + DescribeSclBlocks(scl));
                    AppendLog("      修复后 SCL 共 " + CountLines(scl) + " 行。");

                    // 每次修复后刷新预览，用户能看到最新代码
                    DisplaySclPreview(scl);
                }

                if (result != null && result.Succeeded)
                {
                    AppendLog("✅ 完成：编译状态 " + result.CompileState +
                              "，错误 0，警告 " + result.WarningCount +
                              (repairRound > 0 ? "（自动修复 " + repairRound + " 轮后成功）" : "（首次尝试即成功）") + "。");
                    MessageBox.Show(this,
                        "导入并编译成功！" + (repairRound > 0 ? "（自动修复 " + repairRound + " 轮后成功）" : string.Empty) +
                        "\n\n编译状态：" + result.CompileState +
                        "\n警告数：" + result.WarningCount +
                        "\nSCL 文件：" + result.SclFilePath +
                        "\n\n提示：本次生成的块名为 " + (blockName ?? "(自动识别)") +
                        "，请到博图的“程序块”里查看。",
                        "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    string lastError = result == null
                        ? "（没有拿到导入结果）"
                        : (result.FailureMessage ?? "未知错误");

                    AppendLog("❌ 彻底失败：已尝试导入 " + importAttempt + " 次、自动修复 " + repairRound + " 轮。");

                    MessageBox.Show(this,
                        "自动修复 " + repairRound + " 轮后仍然失败，已停止。" +
                        "\n\n最后一次错误：" + lastError +
                        "\n\n预览区保留了最后一次代码，可手动修改后重新生成。",
                        "导入/编译失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        private static string CleanSclText(string rawText, bool forceExcelMapping = false,
                                           IList<string> originalVariableNames = null,
                                           IDictionary<string, string> addressToTagName = null)
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

            // 6.6) 最后一道防线：纯映射的 FUNCTION_BLOCK（内部没有任何 VAR/VAR_INPUT/VAR_OUTPUT 静态声明）
            //      一律强制改成 FUNCTION (FC 块)，并把 END_FUNCTION_BLOCK 改成 END_FUNCTION
            text = ConvertEmptyFunctionBlockToFunction(text);

            // 6.7) Excel 物理 IO 映射模式专用兜底（只有批量导入路径会传 forceExcelMapping = true）：
            //      ① 去掉 Tag_ 前缀；② 强制 FUNCTION_BLOCK → FUNCTION "FC_MainControl" : Void；
            //      ③ 变量名不是合法标识符（纯数字/中文/特殊符号）时补双引号；④ 绝对地址补双引号；
            //      ⑤ 绝对地址仍未映射到变量名时，用本次创建好的 地址 → 变量名 映射表替换。
            if (forceExcelMapping)
            {
                text = ApplyExcelMappingFallbacks(text, originalVariableNames, addressToTagName);
            }

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
        /// 最后一道防线（Excel 纯映射场景）：AI 仍可能顽固地输出 FUNCTION_BLOCK。
        /// 对每个 FUNCTION_BLOCK 片段，如果内部没有任何 VAR / VAR_INPUT / VAR_OUTPUT 静态声明
        /// （最多只有 VAR_TEMP 或压根不声明），说明它其实就是纯变量映射，强制转成 FUNCTION (FC 块)：
        ///     FUNCTION_BLOCK "xxx"   →   FUNCTION "xxx" : Void
        ///     END_FUNCTION_BLOCK     →   END_FUNCTION
        /// 只要出现过静态 VAR 区（真正需要记忆的 FB），就保持原样不动。
        /// </summary>
        private static string ConvertEmptyFunctionBlockToFunction(string text)
        {
            MatchCollection blockStarts = Regex.Matches(
                text, @"(?<!\w)FUNCTION_BLOCK(?!\w)", RegexOptions.IgnoreCase);

            // 从后往前处理，避免前面的替换影响后面片段的字符下标
            for (int i = blockStarts.Count - 1; i >= 0; i--)
            {
                int blockStart = blockStarts[i].Index;
                int blockEnd = text.Length;

                Match endMatch = Regex.Match(text.Substring(blockStart),
                                             @"(?<!\w)END_FUNCTION_BLOCK(?!\w)", RegexOptions.IgnoreCase);
                if (endMatch.Success)
                {
                    blockEnd = blockStart + endMatch.Index + endMatch.Length;
                }

                string segment = text.Substring(blockStart, blockEnd - blockStart);

                // 出现 VAR / VAR_INPUT / VAR_OUTPUT 就认为它有静态记忆区，保持 FB 不动
                // 注意：VAR_TEMP 后面紧跟下划线（\w），不会被 (?<!\w)VAR(?!\w) 匹配到
                if (Regex.IsMatch(segment, @"(?<!\w)(?:VAR|VAR_INPUT|VAR_OUTPUT)(?!\w)",
                                  RegexOptions.IgnoreCase))
                {
                    continue;
                }

                // 头部的 FUNCTION_BLOCK "块名" 必须能匹配到，否则跳过（避免误伤注释）
                if (!Regex.IsMatch(segment, @"(?<!\w)FUNCTION_BLOCK\s+""([^""]+)""",
                                   RegexOptions.IgnoreCase))
                {
                    continue;
                }

                string fixedSegment = Regex.Replace(
                    segment,
                    @"(?<!\w)FUNCTION_BLOCK\s+""([^""]+)""",
                    "FUNCTION \"$1\" : Void",
                    RegexOptions.IgnoreCase);

                if (endMatch.Success)
                {
                    fixedSegment = Regex.Replace(fixedSegment, @"(?<!\w)END_FUNCTION_BLOCK(?!\w)",
                                                 "END_FUNCTION", RegexOptions.IgnoreCase);
                }
                else
                {
                    // 连结束关键字都没生成时，顺手补一个 END_FUNCTION，避免清洗后不完整
                    fixedSegment = fixedSegment.TrimEnd() + Environment.NewLine + "END_FUNCTION";
                }

                text = text.Substring(0, blockStart) + fixedSegment + text.Substring(blockEnd);
            }

            return text;
        }

        /// <summary>
        /// Excel 物理 IO 映射模式的二次兜底（最后一道防线）。只在批量导入路径启用：
        ///   ① 去掉变量名里多余的 Tag_ 前缀（Tag_111 → 111）；<br/>
        ///   ② 强制把任何 FUNCTION_BLOCK 改成 FUNCTION "FC_MainControl" : Void（丢弃原块名），END_FUNCTION_BLOCK → END_FUNCTION；<br/>
        ///   ③ 变量名不是合法 SCL 标识符（纯数字 / 中文 / 特殊符号）时补双引号（111 → "111"）；<br/>
        ///   ④ 物理绝对地址（%I0.0 等）按本次创建好的 地址 → 变量名 映射表换成 "Tag_1" 这种变量引用；<br/>
        ///   ⑤ 仍残留的绝对地址：按 QuoteAbsoluteAddresses 决定补/去双引号（默认去，博图规定地址不带引号）。
        /// </summary>
        private static string ApplyExcelMappingFallbacks(string text, IList<string> originalVariableNames,
                                                        IDictionary<string, string> addressToTagName)
        {
            // ⓪ AI 私自加的前缀：Var111 → 111（先把 Excel 原变量名精确映射一遍，再做通用兜底）
            if (originalVariableNames != null)
            {
                for (int i = 0; i < originalVariableNames.Count; i++)
                {
                    string originalName = originalVariableNames[i];
                    if (string.IsNullOrEmpty(originalName))
                    {
                        continue;
                    }

                    text = Regex.Replace(text, @"\bVar" + Regex.Escape(originalName) + @"\b", originalName);
                }
            }

            // 通用兜底：Var 后面紧跟数字或 IO 临时变量名（Var111、VarI0_0）→ 直接去掉 Var
            text = Regex.Replace(text, @"(?<![\w""])Var(?=[0-9])", string.Empty);
            text = Regex.Replace(text, @"(?<![\w""])Var(?=[IQM][0-9])", string.Empty);

            // ① 声明区 + DB 引用里的 Tag_ 前缀（只处理代码，不碰 // 注释）
            text = Regex.Replace(text, @"(?m)^([ \t]*)Tag_([A-Za-z0-9_\u4e00-\u9fa5]+)([ \t]*:(?!=))",
                                 "$1$2$3");
            text = Regex.Replace(text, @"(""DB_Variables""\s*\.\s*)Tag_([A-Za-z0-9_\u4e00-\u9fa5]+)",
                                 "$1$2");

            // ② 声明区变量名不是合法标识符时补双引号（"111" : Bool; / "启动按钮" : Bool;）
            text = Regex.Replace(text, @"(?m)^([ \t]*)([^\s:;]+)([ \t]*:(?!=))", delegate (Match m)
            {
                string line = m.Value.TrimStart();
                if (line.StartsWith("//", StringComparison.Ordinal))
                {
                    return m.Value;
                }

                string name = m.Groups[2].Value;
                if (IsValidSclIdentifier(name) || name.StartsWith("\"", StringComparison.Ordinal))
                {
                    return m.Value;
                }

                return m.Groups[1].Value + "\"" + name + "\"" + m.Groups[3].Value;
            });

            // ③ "DB_Variables".xxx 引用里的非法变量名补双引号（"DB_Variables".111 → "DB_Variables"."111"）
            text = Regex.Replace(text, @"(""DB_Variables""\s*\.\s*)([A-Za-z0-9_\u4e00-\u9fa5]+)",
                                 delegate (Match m)
                                 {
                                     string name = m.Groups[2].Value;
                                     return IsValidSclIdentifier(name) ? m.Value : m.Groups[1].Value + "\"" + name + "\"";
                                 });

            // ④ 万能块类型锁定：任何 FUNCTION_BLOCK → FUNCTION "FC_MainControl" : Void
            int fcIndex = 0;
            text = Regex.Replace(text, @"(?<!\w)FUNCTION_BLOCK(?:\s+""[^""]*""|\s+[A-Za-z_][A-Za-z0-9_]*)?",
                                 delegate (Match m)
                                 {
                                     fcIndex++;
                                     string fcName = fcIndex == 1 ? "FC_MainControl" : "FC_MainControl_" + fcIndex;
                                     return "FUNCTION \"" + fcName + "\" : Void";
                                 }, RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"(?<!\w)END_FUNCTION_BLOCK(?!\w)", "END_FUNCTION",
                                 RegexOptions.IgnoreCase);

            // ④.1 块名锁定：任何 FC（包括已被 6.6 转成 FC 的空 FB）统一改名为 FC_MainControl
            int functionIndex = 0;
            text = Regex.Replace(text,
                                 @"(?m)^[ \t]*(?<!\w)FUNCTION(?!\w)\s+(?:""[^""]*""|[A-Za-z_][A-Za-z0-9_]*)(?:[ \t]*:[ \t]*[A-Za-z_][A-Za-z0-9_]*)?",
                                 delegate (Match m)
                                 {
                                     functionIndex++;
                                     string name = functionIndex == 1
                                         ? "FC_MainControl"
                                         : "FC_MainControl_" + functionIndex;
                                     return "FUNCTION \"" + name + "\" : Void";
                                 },
                                 RegexOptions.IgnoreCase);

            // ④.5 绝对地址 → 已创建好的 PLC 变量名（如 %I0.0 → "Tag_1"）：
            //      AI 若还是写了 %I0.0，这里直接换成变量名引用；带引号/不带引号的写法都处理。
            if (addressToTagName != null && addressToTagName.Count > 0)
            {
                foreach (KeyValuePair<string, string> pair in addressToTagName)
                {
                    if (string.IsNullOrWhiteSpace(pair.Key) || string.IsNullOrWhiteSpace(pair.Value))
                    {
                        continue;
                    }

                    string escapedAddress = Regex.Escape(pair.Key.Trim());
                    string tagReference = "\"" + pair.Value.Trim() + "\"";

                    // 带双引号的地址："%I0.0" → "Tag_1"
                    text = Regex.Replace(text, "\"\\s*" + escapedAddress + "\\s*\"", tagReference,
                                         RegexOptions.IgnoreCase);

                    // 不带双引号的地址：%I0.0 → "Tag_1"（前后不能紧贴标识符字符，避免误伤 %I0.01）
                    text = Regex.Replace(text, "(?<![\\w.])" + escapedAddress + "(?![\\w.])", tagReference,
                                         RegexOptions.IgnoreCase);
                }
            }

            // ⑤ 物理绝对地址【绝对不准加双引号】："%I0.0" → %I0.0（只处理 %I/%Q/%M，不碰 DB 变量名）
            //    这是 AI 最容易犯的致命错误，所以默认强制去掉；若要反过来加引号，把 QuoteAbsoluteAddresses 改 true。
            if (!QuoteAbsoluteAddresses)
            {
                text = Regex.Replace(text, "\"\\s*(%[IQM][A-Za-z]?[0-9]+(?:\\.[0-9]+)?)\\s*\"", "$1");
            }
            else
            {
                text = Regex.Replace(text, @"(?<!"")%([IQM][A-Za-z]?[0-9]+(?:\.[0-9]+)?)(?!"")",
                                     "\"%$1\"");
            }

            return text;
        }

        /// <summary>是否合法 SCL 标识符（字母/下划线开头，只含字母数字下划线）。</summary>
        private static bool IsValidSclIdentifier(string name)
        {
            return !string.IsNullOrEmpty(name) && Regex.IsMatch(name, @"^[A-Za-z_][A-Za-z0-9_]*$");
        }

        /// <summary>
        /// 是否给绝对地址加双引号（"%I0.0"）。**博图规定绝对地址不准加双引号**，所以默认 false（强制去引号）；
        /// 只有特殊语法场景才改成 true。用 static readonly 而不是 const，避免另一分支被编译器判定为死代码。
        /// </summary>
        private static readonly bool QuoteAbsoluteAddresses = false;

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
        //  DeepSeek 流式请求（保持原有流式逻辑不变；API Key 改为从界面读取）
        // =================================================================================
        private async Task<string> RequestSclFromDeepSeekAsync(string systemPrompt, string userPrompt,
                                                               Action<string> onDelta)
        {
            // V0.4：使用界面输入的 Key；第一行 await 之前读取，始终在 UI 线程上执行
            string apiKey = txtApiKey != null && txtApiKey.Text != null ? txtApiKey.Text.Trim() : string.Empty;
            if (apiKey.Length == 0)
            {
                throw new InvalidOperationException("DeepSeek API Key 为空，请先在界面输入并点“保存Key”。");
            }

            string requestJson = JsonSerializer.Serialize(BuildRequestPayload(systemPrompt, userPrompt, true));

            using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, DeepSeekApiUrl))
            {
                request.Headers.TryAddWithoutValidation("Authorization", "Bearer " + apiKey);
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