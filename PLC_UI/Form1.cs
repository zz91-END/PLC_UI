// =====================================================================================
//  Form1.cs —— WinForms 界面（V0.1）
//  -------------------------------------------------------------------------------------
//  职责：收集自然语言需求 → 调 DeepSeek 生成 SCL → 交给 TiaService 导入并编译 → 日志回显
//  约束：所有 Openness 操作都在 TiaService 的后台线程里执行，界面不会卡死。
// =====================================================================================

using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Web.Script.Serialization;

namespace PLC_UI
{
    public partial class Form1 : Form
    {
        // =============================== 配置区（改这里） ===============================

        /// <summary>DeepSeek API Key，形如 sk-xxxxxxxxxxxxxxxx。（先留空，自己填）</summary>
        private const string DeepSeekApiKey = "";

        /// <summary>DeepSeek 对话接口（OpenAI 兼容协议）。</summary>
        private const string DeepSeekApiUrl = "https://api.deepseek.com/chat/completions";

        /// <summary>模型名：deepseek-chat（速度快，适合生成代码）。</summary>
        private const string DeepSeekModel = "deepseek-chat";

        /// <summary>System Prompt：强制模型只输出可直接导入 TIA 的 SCL 代码。</summary>
        private const string DeepSeekSystemPrompt =
            "你是一名西门子 TIA Portal SCL 编程专家。用户会提出一个 PLC 功能需求，" +
            "你只输出一段可以直接导入 TIA Portal 的 SCL 源代码，用于创建一个 FUNCTION_BLOCK。\n" +
            "硬性要求：\n" +
            "1. 只输出代码本身，不要任何解释、客套话，也不要 Markdown 代码围栏（```）。\n" +
            "2. 第一行必须是：FUNCTION_BLOCK \"块名\"（块名用英文/数字/下划线，用双引号括起来）。\n" +
            "3. 第二行必须单独占一行：{ S7_Optimized_Access := 'TRUE' }\n" +
            "4. 第三行必须是：VERSION : 0.1\n" +
            "5. 变量声明用 VAR_INPUT / VAR_OUTPUT / VAR_IN_OUT / VAR / VAR_TEMP，并以 END_VAR 结束。\n" +
            "6. 语法必须全部使用英文 ASCII；中文只能出现在 // 行注释里。\n" +
            "7. 只使用已声明的变量，临时变量放 VAR_TEMP，不要出现无法编译的写法。\n" +
            "8. 必须以 END_FUNCTION_BLOCK 结束，代码完整、可直接编译。";

        private const string ButtonIdleText = "生成并导入（DeepSeek → TIA Portal）";

        // =============================== 运行期字段 ===============================

        private static readonly HttpClient Http = CreateHttpClient();
        private readonly TiaService _tiaService = new TiaService();

        public Form1()
        {
            InitializeComponent();
        }

        private static HttpClient CreateHttpClient()
        {
            HttpClient client = new HttpClient();
            client.Timeout = TimeSpan.FromSeconds(180); // 生成较长的 SCL 时留足时间
            return client;
        }

        // =============================== 按钮事件 ===============================

        private async void btnGenerate_Click(object sender, EventArgs e)
        {
            string userPrompt = txtPrompt.Text.Trim();
            if (userPrompt.Length == 0)
            {
                MessageBox.Show(this, "请先在左侧输入你的需求，例如：写一个电机起保停的 FB 块。",
                    "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (string.IsNullOrWhiteSpace(DeepSeekApiKey))
            {
                MessageBox.Show(this, "请先在 Form1.cs 顶部把 DeepSeekApiKey 改成你自己的 Key。",
                    "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            btnGenerate.Enabled = false;
            btnGenerate.Text = "执行中，请稍候...";
            Cursor = Cursors.WaitCursor;
            txtLog.Clear();

            try
            {
                AppendLog("========== 开始 ==========");
                AppendLog("需求：" + userPrompt);

                // ---------- 第一步：调 DeepSeek 生成 SCL ----------
                AppendLog(string.Empty);
                AppendLog("[DeepSeek] 正在请求 SCL 代码 ...");
                string rawContent = await RequestSclFromDeepSeekAsync(userPrompt);
                string sclSource = CleanSclText(rawContent);
                if (sclSource.Length == 0)
                {
                    throw new InvalidOperationException(
                        "DeepSeek 没有返回可用的 SCL 代码。原始返回：" + TrimForLog(rawContent, 300));
                }

                string blockName = ExtractBlockName(sclSource);
                AppendLog("[DeepSeek] 已返回 " + CountLines(sclSource) + " 行 SCL，块名：" +
                          (blockName == null ? "（未识别，将由博图实际生成的块决定）" : blockName));
                AppendLog(string.Empty);
                AppendLog("---- AI 生成的 SCL 预览 ----");
                foreach (string line in FirstLines(sclSource, 40))
                {
                    AppendLog("  " + line);
                }
                AppendLog("---------------------------");

                // ---------- 第二步：交给 TiaService（后台线程执行 Openness） ----------
                AppendLog(string.Empty);
                IProgress<string> progress = new Progress<string>(AppendLog);
                TiaOperationResult result = await _tiaService.ImportAndCompileAsync(sclSource, blockName, progress);

                // ---------- 第三步：展示结果 ----------
                AppendLog(string.Empty);
                if (result.Succeeded)
                {
                    AppendLog("✔ 完成：块 " + result.BlockName + " 已生成并编译通过（错误 0，警告 " +
                              result.WarningCount + "）。");
                    MessageBox.Show(this, "块 " + result.BlockName + " 已导入博图并编译通过。",
                        "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    AppendLog("✘ 失败" + (string.IsNullOrEmpty(result.CompileState)
                        ? string.Empty
                        : "（编译状态：" + result.CompileState + "）"));
                    if (!string.IsNullOrEmpty(result.FailureMessage))
                    {
                        AppendLog(result.FailureMessage);
                    }
                }
            }
            catch (TaskCanceledException)
            {
                AppendLog("✘ 请求超时或被取消。");
            }
            catch (HttpRequestException ex)
            {
                AppendLog("✘ 网络请求失败：" + ex.Message);
            }
            catch (Exception ex)
            {
                AppendLog("✘ 未处理异常：" + ex.Message);
                AppendLog(ex.ToString());
            }
            finally
            {
                Cursor = Cursors.Default;
                btnGenerate.Text = ButtonIdleText;
                btnGenerate.Enabled = true;
            }
        }

        // =============================== 日志 ===============================

        private void AppendLog(string message)
        {
            if (txtLog.InvokeRequired)
            {
                txtLog.BeginInvoke(new Action<string>(AppendLog), message);
                return;
            }

            string line = DateTime.Now.ToString("HH:mm:ss") + "  " + message;
            if (txtLog.TextLength > 0)
            {
                txtLog.AppendText(Environment.NewLine);
            }

            txtLog.AppendText(line);
            txtLog.SelectionStart = txtLog.TextLength;
            txtLog.ScrollToCaret();
        }

        // =============================== DeepSeek 调用 ===============================

        private static async Task<string> RequestSclFromDeepSeekAsync(string userPrompt)
        {
            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["model"] = DeepSeekModel;
            payload["stream"] = false;
            payload["temperature"] = 0.2;
            payload["max_tokens"] = 4096;
            payload["messages"] = new object[]
            {
                new Dictionary<string, object>
                {
                    { "role", "system" },
                    { "content", DeepSeekSystemPrompt }
                },
                new Dictionary<string, object>
                {
                    { "role", "user" },
                    { "content", "请用 SCL 实现下面的需求（块名自定，英文）：\n" + userPrompt }
                }
            };

            JavaScriptSerializer serializer = new JavaScriptSerializer();
            string requestJson = serializer.Serialize(payload);

            using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, DeepSeekApiUrl))
            {
                request.Headers.TryAddWithoutValidation("Authorization", "Bearer " + DeepSeekApiKey);
                request.Content = new StringContent(requestJson, Encoding.UTF8, "application/json");

                using (HttpResponseMessage response = await Http.SendAsync(request).ConfigureAwait(false))
                {
                    string responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException(
                            "DeepSeek API 调用失败：HTTP " + (int)response.StatusCode + " " +
                            response.ReasonPhrase + "，返回内容：" + TrimForLog(responseBody, 500));
                    }

                    return ExtractAssistantContent(responseBody);
                }
            }
        }

        private static string ExtractAssistantContent(string responseBody)
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            Dictionary<string, object> root =
                serializer.Deserialize<Dictionary<string, object>>(responseBody);

            object choicesValue;
            if (root == null || !root.TryGetValue("choices", out choicesValue))
            {
                throw new InvalidOperationException("无法解析 DeepSeek 返回内容：" + TrimForLog(responseBody, 500));
            }

            System.Collections.IList choices = choicesValue as System.Collections.IList;
            if (choices == null || choices.Count == 0)
            {
                throw new InvalidOperationException("DeepSeek 返回的 choices 为空。原始返回：" + responseBody);
            }

            Dictionary<string, object> firstChoice = choices[0] as Dictionary<string, object>;
            if (firstChoice == null)
            {
                throw new InvalidOperationException("DeepSeek 返回的 choices 结构无法识别。");
            }

            object messageValue;
            if (!firstChoice.TryGetValue("message", out messageValue))
            {
                throw new InvalidOperationException("DeepSeek 返回内容缺少 message 字段。");
            }

            Dictionary<string, object> message = messageValue as Dictionary<string, object>;
            if (message == null)
            {
                throw new InvalidOperationException("DeepSeek 返回的 message 结构无法识别。");
            }

            object contentValue;
            if (!message.TryGetValue("content", out contentValue) || contentValue == null)
            {
                throw new InvalidOperationException("DeepSeek 返回的 message.content 为空。");
            }

            return contentValue as string;
        }

        // =============================== SCL 文本清洗/解析 ===============================

        /// <summary>去掉 Markdown 围栏、BOM、模型前言后语，只保留 FUNCTION_BLOCK ... END_FUNCTION_BLOCK。</summary>
        private static string CleanSclText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            string normalized = text.Replace("\r\n", "\n")
                                    .Replace('\uFEFF', ' ')
                                    .Replace('`', ' ');

            string[] lines = normalized.Split('\n');

            int start = 0;
            bool foundStart = false;
            for (int i = 0; i < lines.Length; i++)
            {
                string trimmed = lines[i].Trim();
                if (StartsWithIgnoreCase(trimmed, "FUNCTION_BLOCK") ||
                    StartsWithIgnoreCase(trimmed, "FUNCTION ") ||
                    StartsWithIgnoreCase(trimmed, "ORGANIZATION_BLOCK") ||
                    StartsWithIgnoreCase(trimmed, "DATA_BLOCK"))
                {
                    start = i;
                    foundStart = true;
                    break;
                }
            }

            if (!foundStart)
            {
                return string.Empty;
            }

            int end = lines.Length - 1;
            for (int i = start; i < lines.Length; i++)
            {
                string trimmed = lines[i].Trim();
                if (StartsWithIgnoreCase(trimmed, "END_FUNCTION_BLOCK") ||
                    StartsWithIgnoreCase(trimmed, "END_FUNCTION") ||
                    StartsWithIgnoreCase(trimmed, "END_ORGANIZATION_BLOCK") ||
                    StartsWithIgnoreCase(trimmed, "END_DATA_BLOCK"))
                {
                    end = i;
                    break;
                }
            }

            List<string> keptLines = new List<string>();
            for (int i = start; i <= end; i++)
            {
                keptLines.Add(lines[i].TrimEnd());
            }

            return string.Join(Environment.NewLine, keptLines) + Environment.NewLine;
        }

        /// <summary>从 SCL 中提取 FUNCTION_BLOCK 的块名，识别不到返回 null。</summary>
        private static string ExtractBlockName(string sclSource)
        {
            Match match = Regex.Match(sclSource, "FUNCTION_BLOCK\\s+\"?([A-Za-z_][A-Za-z0-9_]*)\"?",
                                      RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value : null;
        }

        private static bool StartsWithIgnoreCase(string text, string prefix)
        {
            return text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
        }

        private static int CountLines(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            return text.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).Length;
        }

        private static List<string> FirstLines(string text, int maxLines)
        {
            List<string> result = new List<string>();
            if (string.IsNullOrEmpty(text))
            {
                return result;
            }

            string[] lines = text.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            for (int i = 0; i < lines.Length && i < maxLines; i++)
            {
                result.Add(lines[i]);
            }

            if (lines.Length > maxLines)
            {
                result.Add("...（共 " + lines.Length + " 行，仅预览前 " + maxLines + " 行）");
            }

            return result;
        }

        private static string TrimForLog(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            string singleLine = text.Replace("\r", " ").Replace("\n", " ");
            return singleLine.Length <= maxLength
                ? singleLine
                : singleLine.Substring(0, maxLength) + " ...";
        }
    }
}