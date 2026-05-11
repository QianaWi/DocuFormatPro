using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocuFormatPro.Models;

namespace DocuFormatPro.Services
{
    /// <summary>
    /// 排版规则持久化服务
    /// 将规则以 JSON 文件保存到应用目录下的 Templates 文件夹
    /// </summary>
    public class RuleStorageService
    {
        private readonly string _templatesDir;

        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true,
            Converters = { new JsonStringEnumConverter() },
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        public RuleStorageService()
        {
            // 模板存储在可执行文件同级的 Templates 目录下
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            _templatesDir = Path.Combine(appDir, "Templates");
            if (!Directory.Exists(_templatesDir))
            {
                Directory.CreateDirectory(_templatesDir);
            }
        }

        /// <summary>
        /// 保存排版规则为命名模板
        /// </summary>
        public void SaveRule(FormattingRule rule, string name)
        {
            rule.RuleName = name;
            string safeName = SanitizeFileName(name);
            string filePath = Path.Combine(_templatesDir, $"{safeName}.json");
            string json = JsonSerializer.Serialize(rule, JsonOptions);
            File.WriteAllText(filePath, json);
        }

        /// <summary>
        /// 加载指定名称的模板
        /// </summary>
        public FormattingRule? LoadRule(string name)
        {
            string safeName = SanitizeFileName(name);
            string filePath = Path.Combine(_templatesDir, $"{safeName}.json");
            if (!File.Exists(filePath)) return null;

            string json = File.ReadAllText(filePath);
            return JsonSerializer.Deserialize<FormattingRule>(json, JsonOptions);
        }

        /// <summary>
        /// 删除指定名称的模板
        /// </summary>
        public bool DeleteRule(string name)
        {
            string safeName = SanitizeFileName(name);
            string filePath = Path.Combine(_templatesDir, $"{safeName}.json");
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
                return true;
            }
            return false;
        }

        /// <summary>
        /// 列出所有已保存的模板名称
        /// </summary>
        public List<string> ListSavedRules()
        {
            var result = new List<string>();
            if (!Directory.Exists(_templatesDir)) return result;

            foreach (var file in Directory.GetFiles(_templatesDir, "*.json"))
            {
                try
                {
                    string json = File.ReadAllText(file);
                    var rule = JsonSerializer.Deserialize<FormattingRule>(json, JsonOptions);
                    if (rule != null)
                    {
                        result.Add(rule.RuleName);
                    }
                }
                catch
                {
                    // 跳过无法解析的文件
                    result.Add(Path.GetFileNameWithoutExtension(file));
                }
            }
            return result;
        }

        /// <summary>
        /// 获取默认规则（如果已保存过默认模板则加载，否则创建新的）
        /// </summary>
        public FormattingRule GetDefaultRule()
        {
            var saved = LoadRule("默认模板");
            return saved ?? FormattingRule.CreateDefault();
        }

        /// <summary>
        /// 清理文件名中的非法字符
        /// </summary>
        private static string SanitizeFileName(string name)
        {
            char[] invalid = Path.GetInvalidFileNameChars();
            string safe = name;
            foreach (char c in invalid)
            {
                safe = safe.Replace(c, '_');
            }
            return safe;
        }
    }
}
