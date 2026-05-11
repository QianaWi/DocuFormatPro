using System.Runtime.InteropServices;
using DocuFormatPro.Models;
using Microsoft.Office.Interop.Word;

namespace DocuFormatPro.Services
{
    /// <summary>
    /// 模板提取服务
    /// 从已排好版的 Word 文档中读取格式规则，生成 FormattingRule 对象
    /// </summary>
    public class TemplateExtractorService
    {
        /// <summary>
        /// 从指定文档中提取排版规则
        /// </summary>
        public async System.Threading.Tasks.Task<FormattingRule> ExtractFromDocumentAsync(
            string filePath,
            IProgress<string>? progress = null,
            CancellationToken cancellationToken = default)
        {
            FormattingRule rule = FormattingRule.CreateDefault();

            await System.Threading.Tasks.Task.Run(() =>
            {
                var thread = new Thread(() =>
                {
                    Application? wordApp = null;
                    Document? doc = null;
                    try
                    {
                        progress?.Report("正在启动 Word 引擎...");

                        wordApp = new Application
                        {
                            Visible = false,
                            ScreenUpdating = false,
                            DisplayAlerts = WdAlertLevel.wdAlertsNone
                        };

                        progress?.Report($"正在打开文档: {System.IO.Path.GetFileName(filePath)}");

                        doc = wordApp.Documents.Open(
                            FileName: filePath,
                            ReadOnly: true,
                            Visible: false);

                        cancellationToken.ThrowIfCancellationRequested();

                        // ===== 提取页面设置 =====
                        progress?.Report("正在提取页面设置...");
                        ExtractPageMargins(doc, wordApp, rule);

                        // ===== 提取正文格式 =====
                        progress?.Report("正在提取正文格式...");
                        ExtractBodyTextFormat(doc, rule);

                        // ===== 提取标题格式 =====
                        progress?.Report("正在提取标题格式...");
                        ExtractHeadingFormats(doc, rule);

                        progress?.Report("模板提取完成！");
                    }
                    catch (OperationCanceledException)
                    {
                        progress?.Report("提取已取消");
                        throw;
                    }
                    catch (Exception ex)
                    {
                        progress?.Report($"提取失败: {ex.Message}");
                        throw;
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            try { doc.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
                            Marshal.ReleaseComObject(doc);
                        }
                        if (wordApp != null)
                        {
                            try { wordApp.Quit(WdSaveOptions.wdDoNotSaveChanges); } catch { }
                            Marshal.ReleaseComObject(wordApp);
                        }
                    }
                });

                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();

            }, cancellationToken);

            return rule;
        }

        /// <summary>提取页面边距</summary>
        private void ExtractPageMargins(Document doc, Application wordApp, FormattingRule rule)
        {
            try
            {
                var pageSetup = doc.Sections[1].PageSetup;
                rule.PageMargins.TopMargin = (float)Math.Round(wordApp.PointsToCentimeters(pageSetup.TopMargin), 2);
                rule.PageMargins.BottomMargin = (float)Math.Round(wordApp.PointsToCentimeters(pageSetup.BottomMargin), 2);
                rule.PageMargins.LeftMargin = (float)Math.Round(wordApp.PointsToCentimeters(pageSetup.LeftMargin), 2);
                rule.PageMargins.RightMargin = (float)Math.Round(wordApp.PointsToCentimeters(pageSetup.RightMargin), 2);
            }
            catch { /* 保留默认值 */ }
        }

        /// <summary>提取正文格式（从 "正文" 样式或第一个正文段落）</summary>
        private void ExtractBodyTextFormat(Document doc, FormattingRule rule)
        {
            try
            {
                // 尝试从 "正文" 样式提取
                Style? normalStyle = null;
                try { normalStyle = doc.Styles[WdBuiltinStyle.wdStyleNormal]; } catch { }

                if (normalStyle != null)
                {
                    var font = normalStyle.Font;
                    if (!string.IsNullOrEmpty(font.Name))
                        rule.BodyText.ChineseFontName = font.Name;
                    if (!string.IsNullOrEmpty(font.NameAscii))
                        rule.BodyText.EnglishFontName = font.NameAscii;
                    if (font.Size > 0)
                    {
                        rule.BodyText.FontSizePoint = font.Size;
                        rule.BodyText.FontSizeName = FontSizeMapping.GetName(font.Size);
                    }
                    rule.BodyText.IsBold = (font.Bold == -1); // WdConstants: True = -1

                    // 段落格式
                    var pf = normalStyle.ParagraphFormat;
                    if (pf.FirstLineIndent > 0)
                    {
                        // 将磅值转换为字符数（近似：1字符 ≈ 字号磅值）
                        rule.Paragraph.FirstLineIndentChars = (float)Math.Round(pf.FirstLineIndent / rule.BodyText.FontSizePoint, 1);
                    }

                    ExtractLineSpacing(pf, rule.Paragraph);

                    if (pf.SpaceBefore >= 0)
                        rule.Paragraph.SpaceBeforeLines = (float)Math.Round(pf.SpaceBefore / 12f, 1);
                    if (pf.SpaceAfter >= 0)
                        rule.Paragraph.SpaceAfterLines = (float)Math.Round(pf.SpaceAfter / 12f, 1);
                }

                // 也尝试从第一个正文段落提取作为补充
                foreach (Paragraph para in doc.Paragraphs)
                {
                    try
                    {
                        var styleName = ((Style)para.get_Style()).NameLocal;
                        if (styleName == "正文" || styleName == "Normal")
                        {
                            var font = para.Range.Font;
                            if (!string.IsNullOrEmpty(font.Name) && font.Name != normalStyle?.Font?.Name)
                                rule.BodyText.ChineseFontName = font.Name;
                            if (!string.IsNullOrEmpty(font.NameAscii))
                                rule.BodyText.EnglishFontName = font.NameAscii;
                            break;
                        }
                    }
                    catch { continue; }
                }
            }
            catch { /* 保留默认值 */ }
        }

        /// <summary>提取标题样式</summary>
        private void ExtractHeadingFormats(Document doc, FormattingRule rule)
        {
            var headingStyles = new[]
            {
                (WdBuiltinStyle.wdStyleHeading1, 1),
                (WdBuiltinStyle.wdStyleHeading2, 2),
                (WdBuiltinStyle.wdStyleHeading3, 3)
            };

            rule.Headings.Clear();

            foreach (var (builtinStyle, level) in headingStyles)
            {
                try
                {
                    var style = doc.Styles[builtinStyle];
                    var heading = new Models.HeadingStyle
                    {
                        Level = level,
                        ChineseFontName = !string.IsNullOrEmpty(style.Font.Name) ? style.Font.Name : "黑体",
                        EnglishFontName = !string.IsNullOrEmpty(style.Font.NameAscii) ? style.Font.NameAscii : "Times New Roman",
                        FontSizePoint = style.Font.Size > 0 ? style.Font.Size : 16f - (level - 1) * 2f,
                        IsBold = (style.Font.Bold == -1),
                    };
                    heading.FontSizeName = FontSizeMapping.GetName(heading.FontSizePoint);

                    // 对齐方式
                    heading.Alignment = style.ParagraphFormat.Alignment switch
                    {
                        WdParagraphAlignment.wdAlignParagraphCenter => TextAlignment.Center,
                        WdParagraphAlignment.wdAlignParagraphRight => TextAlignment.Right,
                        WdParagraphAlignment.wdAlignParagraphJustify => TextAlignment.Justify,
                        _ => TextAlignment.Left
                    };

                    // 行距
                    var hPara = new ParagraphSettings();
                    ExtractLineSpacing(style.ParagraphFormat, hPara);
                    heading.LineSpacingType = hPara.LineSpacingType;
                    heading.LineSpacingValue = hPara.LineSpacingValue;

                    // 段前段后
                    if (style.ParagraphFormat.SpaceBefore >= 0)
                        heading.SpaceBeforeLines = (float)Math.Round(style.ParagraphFormat.SpaceBefore / 12f, 1);
                    if (style.ParagraphFormat.SpaceAfter >= 0)
                        heading.SpaceAfterLines = (float)Math.Round(style.ParagraphFormat.SpaceAfter / 12f, 1);

                    rule.Headings.Add(heading);
                }
                catch
                {
                    // 如果该级别标题样式不存在，使用默认值
                    rule.Headings.Add(new Models.HeadingStyle { Level = level });
                }
            }
        }

        /// <summary>从 ParagraphFormat 提取行距设置</summary>
        private void ExtractLineSpacing(ParagraphFormat pf, ParagraphSettings para)
        {
            try
            {
                var lineRule = pf.LineSpacingRule;
                float lineSpacing = pf.LineSpacing;

                switch (lineRule)
                {
                    case WdLineSpacing.wdLineSpaceSingle:
                        para.LineSpacingType = LineSpacingType.Single;
                        para.LineSpacingValue = 1f;
                        break;
                    case WdLineSpacing.wdLineSpace1pt5:
                        para.LineSpacingType = LineSpacingType.OneAndHalf;
                        para.LineSpacingValue = 1.5f;
                        break;
                    case WdLineSpacing.wdLineSpaceDouble:
                        para.LineSpacingType = LineSpacingType.Double;
                        para.LineSpacingValue = 2f;
                        break;
                    case WdLineSpacing.wdLineSpaceMultiple:
                        para.LineSpacingType = LineSpacingType.Multiple;
                        para.LineSpacingValue = (float)Math.Round(lineSpacing / 12f, 2);
                        break;
                    case WdLineSpacing.wdLineSpaceExactly:
                        para.LineSpacingType = LineSpacingType.Fixed;
                        para.LineSpacingValue = lineSpacing;
                        break;
                    case WdLineSpacing.wdLineSpaceAtLeast:
                        para.LineSpacingType = LineSpacingType.AtLeast;
                        para.LineSpacingValue = lineSpacing;
                        break;
                }
            }
            catch { /* 保留默认值 */ }
        }
    }
}
