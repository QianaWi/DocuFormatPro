using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace DocuFormatPro.Services
{
    /// <summary>
    /// 正文文本规范化服务
    /// 处理 AI 风格文本问题：中英文空格、英文标点转中文全角
    /// 只处理正文段落，跳过表格、题注、代码样式段落
    /// </summary>
    public class TextNormalizationService
    {
        // 跳过处理的样式名（题注、代码、预格式化等）
        private static readonly HashSet<string> SkipStyleNames = new(StringComparer.OrdinalIgnoreCase)
        {
            "题注", "Caption",
            "代码", "Code", "HTML Code", "HTML Preformatted",
            "Preformatted Text", "Plain Text",
            "宏文字", "Macro Text"
        };

        /// <summary>规范化文档中所有正文段落的文本</summary>
        public void NormalizeBodyText(Document doc, IProgress<string>? progress = null)
        {
            progress?.Report("正在规范化正文文本格式...");

            foreach (Paragraph para in doc.Paragraphs)
            {
                try
                {
                    // 跳过表格内段落
                    bool isInTable = false;
                    try { isInTable = (bool)para.Range.Information[WdInformation.wdWithInTable]; } catch { }
                    if (isInTable) continue;

                    // 获取样式名
                    string styleName = "";
                    try { styleName = ((Style)para.get_Style()).NameLocal; } catch { continue; }

                    // 跳过题注、代码等样式
                    if (SkipStyleNames.Contains(styleName)) continue;

                    // 跳过标题样式（只处理正文）
                    if (styleName.Contains("标题") || styleName.Contains("Heading")) continue;

                    // 获取段落文本（排除段落标记）
                    var r = para.Range;
                    object unit = WdUnits.wdCharacter;
                    object cnt = -1;
                    r.MoveEnd(ref unit, ref cnt);

                    string original = r.Text;
                    if (string.IsNullOrEmpty(original)) continue;

                    string normalized = Normalize(original);
                    if (normalized != original)
                        r.Text = normalized;
                }
                catch { continue; }
            }

            progress?.Report("文本规范化完成");
        }

        /// <summary>对单个字符串执行所有规范化处理</summary>
        public static string Normalize(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            // 1. 英文引号转中文全角（需在标点转换前处理，保持配对）
            text = ConvertQuotes(text);

            // 2. 英文标点转中文全角（排除数字上下文中的小数点和连字符）
            text = ConvertPunctuation(text);

            // 3. 去掉中文与英文/数字之间的空格
            text = RemoveMixedSpaces(text);

            return text;
        }

        /// <summary>英文引号转中文全角引号（成对匹配）</summary>
        private static string ConvertQuotes(string text)
        {
            // 双引号：" → " 和 "
            var sb = new StringBuilder(text);
            bool inDouble = false;
            bool inSingle = false;

            for (int i = 0; i < sb.Length; i++)
            {
                char c = sb[i];
                if (c == '"')
                {
                    sb[i] = inDouble ? '\u201D' : '\u201C'; // " or "
                    inDouble = !inDouble;
                }
                else if (c == '\'')
                {
                    // 单引号：区分撇号和引号
                    bool prevIsChinese = i > 0 && IsChinese(sb[i - 1]);
                    bool nextIsChinese = i < sb.Length - 1 && IsChinese(sb[i + 1]);
                    bool isApostrophe = i > 0 && char.IsLetterOrDigit(sb[i - 1]) &&
                                       i < sb.Length - 1 && char.IsLetterOrDigit(sb[i + 1]);

                    if (!isApostrophe) // 不是撇号才转换
                    {
                        sb[i] = inSingle ? '\u2019' : '\u2018'; // ' or '
                        inSingle = !inSingle;
                    }
                }
            }

            return sb.ToString();
        }

        /// <summary>英文标点转中文全角（保护小数点、网址、版本号等）</summary>
        private static string ConvertPunctuation(string text)
        {
            // 逗号：不在数字之间时转换
            text = Regex.Replace(text, @"(?<![0-9]),(?![0-9])", "，");

            // 句号：不在数字之间（小数点）、不在字母之间（网址/缩写）时转换
            text = Regex.Replace(text, @"(?<![0-9a-zA-Z])\.(?![0-9a-zA-Z])", "。");

            // 冒号：不在数字之间（时间 12:30）时转换
            text = Regex.Replace(text, @"(?<![0-9]):(?![0-9/])", "：");

            // 感叹号
            text = text.Replace("!", "！");

            // 问号
            text = text.Replace("?", "？");

            // 括号（只转换前后有中文或空格的情况，避免转换代码中的括号）
            text = Regex.Replace(text, @"(?<=[\u4e00-\u9fff\s])\(|(?<=^)\((?=[\u4e00-\u9fff])", "（");
            text = Regex.Replace(text, @"\)(?=[\u4e00-\u9fff\s，。！？])", "）");

            return text;
        }

        /// <summary>去除中文字符与英文/数字之间多余的空格</summary>
        private static string RemoveMixedSpaces(string text)
        {
            // 中文后跟空格再跟英文/数字：去掉空格
            text = Regex.Replace(text, @"([\u4e00-\u9fff]) +([\u0021-\u007e])", "$1$2");

            // 英文/数字后跟空格再跟中文：去掉空格
            text = Regex.Replace(text, @"([\u0021-\u007e]) +([\u4e00-\u9fff])", "$1$2");

            return text;
        }

        private static bool IsChinese(char c) => c >= '\u4e00' && c <= '\u9fff';
    }
}
