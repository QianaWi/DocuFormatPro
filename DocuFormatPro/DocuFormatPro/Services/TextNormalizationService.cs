using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace DocuFormatPro.Services
{
    /// <summary>
    /// 正文文本规范化服务。
    /// 处理中英文空格、英文标点转中文全角，以及中文引号字体对齐。
    /// </summary>
    public class TextNormalizationService
    {
        private static readonly HashSet<string> SkipStyleNames = new(StringComparer.OrdinalIgnoreCase)
        {
            "题注", "Caption",
            "代码", "Code", "HTML Code", "HTML Preformatted",
            "Preformatted Text", "Plain Text",
            "宏文本", "Macro Text"
        };

        public void NormalizeBodyText(
            Document doc,
            string chineseFontName,
            string? englishFontName = null,
            IProgress<string>? progress = null)
        {
            progress?.Report("正在规范化正文文本格式...");

            foreach (Paragraph para in doc.Paragraphs)
            {
                try
                {
                    bool isInTable = false;
                    try { isInTable = (bool)para.Range.Information[WdInformation.wdWithInTable]; } catch { }
                    if (isInTable) continue;

                    string styleName = "";
                    try { styleName = ((Style)para.get_Style()).NameLocal; } catch { continue; }
                    if (SkipStyleNames.Contains(styleName)) continue;
                    if (styleName.Contains("标题") || styleName.Contains("Heading")) continue;

                    var r = para.Range;
                    object unit = WdUnits.wdCharacter;
                    object cnt = -1;
                    r.MoveEnd(ref unit, ref cnt);

                    string original = r.Text;
                    if (string.IsNullOrEmpty(original)) continue;

                    string normalized = Normalize(original);
                    if (normalized != original)
                        r.Text = normalized;

                    if (ContainsChineseQuotationOrPunctuation(normalized))
                    {
                        ApplyChineseFont(r, chineseFontName, englishFontName);
                        ApplyChineseFontToTargetCharacters(r, chineseFontName);
                    }
                }
                catch { continue; }
            }

            progress?.Report("文本规范化完成");
        }

        public void NormalizeTableText(
            Document doc,
            string chineseFontName,
            string? englishFontName = null,
            IProgress<string>? progress = null)
        {
            progress?.Report("姝ｅ湪瑙勮寖鍖栬〃鏍煎唴鏂囨湰...");

            foreach (Table table in doc.Tables)
            {
                foreach (Paragraph para in table.Range.Paragraphs)
                {
                    try
                    {
                        var r = para.Range;
                        object unit = WdUnits.wdCharacter;
                        object cnt = -1;
                        r.MoveEnd(ref unit, ref cnt);

                        string original = r.Text;
                        if (string.IsNullOrEmpty(original)) continue;

                        string normalized = Normalize(original);
                        if (normalized != original)
                            r.Text = normalized;

                        if (ContainsChineseQuotationOrPunctuation(normalized))
                        {
                            ApplyChineseFont(r, chineseFontName, englishFontName);
                            ApplyChineseFontToTargetCharacters(r, chineseFontName);
                        }
                    }
                    catch { continue; }
                }
            }

            progress?.Report("琛ㄦ牸鏂囨湰瑙勮寖鍖栧畬鎴?");
        }

        public static string Normalize(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            text = ConvertQuotes(text);
            text = ConvertPunctuation(text);
            text = RemoveMixedSpaces(text);

            return text;
        }

        private static string ConvertQuotes(string text)
        {
            var sb = new StringBuilder(text);
            bool inDouble = false;
            bool inSingle = false;

            for (int i = 0; i < sb.Length; i++)
            {
                char c = sb[i];
                if (c == '"')
                {
                    sb[i] = inDouble ? '\u201D' : '\u201C';
                    inDouble = !inDouble;
                }
                else if (c == '\'')
                {
                    bool isApostrophe = i > 0 && char.IsLetterOrDigit(sb[i - 1]) &&
                                        i < sb.Length - 1 && char.IsLetterOrDigit(sb[i + 1]);
                    if (!isApostrophe)
                    {
                        sb[i] = inSingle ? '\u2019' : '\u2018';
                        inSingle = !inSingle;
                    }
                }
            }

            return sb.ToString();
        }

        private static string ConvertPunctuation(string text)
        {
            text = Regex.Replace(text, @"(?<![0-9]),(?![0-9])", "，");
            text = Regex.Replace(text, @"(?<![0-9a-zA-Z])\.(?![0-9a-zA-Z])", "。");
            text = Regex.Replace(text, @"(?<![0-9]):(?![0-9/])", "：");
            text = text.Replace("!", "！");
            text = text.Replace("?", "？");
            text = Regex.Replace(text, @"(?<=[\u4e00-\u9fff\s])\(|(?<=^)\((?=[\u4e00-\u9fff])", "（");
            text = Regex.Replace(text, @"\)(?=[\u4e00-\u9fff\s，。！？：；、])", "）");

            return text;
        }

        private static string RemoveMixedSpaces(string text)
        {
            const string mixedSpace = @"[ \t\u00A0\u1680\u2000-\u200A\u202F\u205F\u3000]+";
            const string cjkOrFullWidth = @"[\p{IsCJKUnifiedIdeographs}\u3000-\u303F\uFF00-\uFFEF]";
            const string chinesePunctuation = @"[\u3001\u3002\uFF01\uFF1F\uFF1B\uFF1A\uFF08\uFF09\u300A\u300B\u300C\u300D\u300E\u300F\u3008\u3009\u3010\u3011\u201C\u201D\u2018\u2019]";
            const string latinOrDigit = @"[A-Za-z0-9]";
            const string glueToken = @"[A-Za-z0-9%±≥≤=/+\-]";
            const string digitOnly = @"[0-9]";

            text = Regex.Replace(text, $"(?<={cjkOrFullWidth}){mixedSpace}(?={latinOrDigit})", "");
            text = Regex.Replace(text, $"(?<={latinOrDigit}){mixedSpace}(?={cjkOrFullWidth})", "");
            text = Regex.Replace(text, $"(?<={glueToken}){mixedSpace}(?={digitOnly})", "");
            text = Regex.Replace(text, $"(?<={digitOnly}){mixedSpace}(?={glueToken})", "");
            text = Regex.Replace(text, $"(?<={cjkOrFullWidth}){mixedSpace}(?={glueToken})", "");
            text = Regex.Replace(text, $"(?<={glueToken}){mixedSpace}(?={cjkOrFullWidth})", "");

            text = Regex.Replace(text, $"(?<={chinesePunctuation}){mixedSpace}(?={latinOrDigit})", "");
            text = Regex.Replace(text, $"(?<={latinOrDigit}){mixedSpace}(?={chinesePunctuation})", "");
            text = Regex.Replace(text, $"(?<={chinesePunctuation}){mixedSpace}(?={glueToken})", "");
            text = Regex.Replace(text, $"(?<={glueToken}){mixedSpace}(?={chinesePunctuation})", "");

            text = Regex.Replace(text, $"(?<=[\\p{{IsCJKUnifiedIdeographs}}]){mixedSpace}(?={chinesePunctuation})", "");
            text = Regex.Replace(text, $"(?<={chinesePunctuation}){mixedSpace}(?=[\\p{{IsCJKUnifiedIdeographs}}])", "");

            return text;
        }

        private static bool ContainsChineseQuotationOrPunctuation(string text)
        {
            foreach (char c in text)
            {
                if (IsChineseQuotationOrPunctuationChar(c))
                    return true;
            }

            return false;
        }

        private static bool IsChineseQuotationOrPunctuationChar(char c)
            => c is '\u201C' or '\u201D' or '\u300C' or '\u300D' or '\u300E' or '\u300F' or '\u300A' or '\u300B'
                or '\uFF02' or '\uFF07'
                or '\u3001' or '\u3002' or '\uFF0C' or '\uFF1A' or '\uFF1B' or '\uFF1F' or '\uFF01';

        private static void ApplyChineseFont(Microsoft.Office.Interop.Word.Range range, string chineseFontName, string? englishFontName)
        {
            try
            {
                range.Font.Name = chineseFontName;
                range.Font.NameFarEast = chineseFontName;
                range.Font.NameOther = chineseFontName;

                if (!string.IsNullOrWhiteSpace(englishFontName))
                    range.Font.NameAscii = englishFontName;
            }
            catch
            {
                // Ignore
            }
        }

        private static void ApplyChineseFontToTargetCharacters(Microsoft.Office.Interop.Word.Range range, string chineseFontName)
        {
            try
            {
                foreach (Microsoft.Office.Interop.Word.Range characterRange in range.Characters)
                {
                    string text = characterRange.Text;
                    if (string.IsNullOrEmpty(text) || !IsChineseQuotationOrPunctuationChar(text[0]))
                        continue;

                    characterRange.Font.Name = chineseFontName;
                    characterRange.Font.NameFarEast = chineseFontName;
                    characterRange.Font.NameOther = chineseFontName;
                    characterRange.Font.NameAscii = chineseFontName;
                }
            }
            catch
            {
                // Ignore
            }
        }
    }
}
