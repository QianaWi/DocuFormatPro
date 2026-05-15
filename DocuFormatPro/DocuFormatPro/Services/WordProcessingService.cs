using System.Runtime.InteropServices;
using DocuFormatPro.Models;
using Microsoft.Office.Interop.Word;

namespace DocuFormatPro.Services
{
    /// <summary>
    /// Word COM 自动化服务类
    /// 根据 FormattingRule 对文档进行排版处理
    /// </summary>
    public class WordProcessingService : IDisposable
    {
        private bool _disposed;

        /// <summary>
        /// 异步处理单个 Word 文档
        /// </summary>
        public async System.Threading.Tasks.Task ProcessDocumentAsync(
            string filePath,
            FormattingRule rule,
            IProgress<string>? progress = null,
            CancellationToken cancellationToken = default)
        {
            await System.Threading.Tasks.Task.Run(() =>
            {
                Exception? threadException = null;

                var thread = new Thread(() =>
                {
                    Application? wordApp = null;
                    Document? doc = null;
                    try
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        // 每次处理在当前 STA 线程内创建独立的 Word Application 实例
                        // 避免跨线程复用 RCW 导致 "COM object that has been separated from its underlying RCW" 错误
                        wordApp = new Application
                        {
                            Visible = false,
                            ScreenUpdating = false,
                            DisplayAlerts = WdAlertLevel.wdAlertsNone
                        };

                        progress?.Report($"正在打开文档: {System.IO.Path.GetFileName(filePath)}");

                        doc = wordApp.Documents.Open(
                            FileName: filePath,
                            ReadOnly: false,
                            Visible: false);

                        cancellationToken.ThrowIfCancellationRequested();

                        // ===== 1. 设置页边距 =====
                        progress?.Report("正在标准化页边距...");
                        SetPageMargins(wordApp, doc, rule);
                        cancellationToken.ThrowIfCancellationRequested();

                        // ===== 2. 设置正文样式 =====
                        progress?.Report("正在设置正文格式...");
                        ApplyBodyTextStyle(doc, rule);
                        cancellationToken.ThrowIfCancellationRequested();

                        // ===== 3. 设置标题样式 =====
                        progress?.Report("正在设置标题格式...");
                        ApplyHeadingStyles(doc, rule);
                        cancellationToken.ThrowIfCancellationRequested();

                        // ===== 4. 逐段落应用格式 =====
                        progress?.Report("正在处理段落格式...");
                        ApplyParagraphFormatting(doc, rule);
                        cancellationToken.ThrowIfCancellationRequested();

                        // ===== 5. 格式化表格 =====
                        if (rule.Table.ApplyTableFormatting)
                        {
                            progress?.Report("正在重塑表格...");
                            FormatAllTables(doc, rule);
                            cancellationToken.ThrowIfCancellationRequested();
                        }

                        // ===== 6. 清除所有文字背景色 =====
                        progress?.Report("正在清除文字背景色...");
                        ClearAllTextBackground(doc);
                        cancellationToken.ThrowIfCancellationRequested();

                        // ===== 7. 处理表格和图片题注 =====
                        if (rule.Table.ApplyTableCaptions)
                        {
                            progress?.Report("正在处理题注...");
                            var captionService = new CaptionService();
                            captionService.ProcessCaptions(doc, progress);
                            cancellationToken.ThrowIfCancellationRequested();
                        }

                        // ===== 8. 插入前置页 =====
                        if (rule.FrontMatter.InsertFrontMatter && !string.IsNullOrWhiteSpace(rule.FrontMatter.TemplateFilePath))
                        {
                            progress?.Report("正在插入封面与前置页...");
                            InsertFrontMatterPages(wordApp, doc, rule.FrontMatter.TemplateFilePath);
                            cancellationToken.ThrowIfCancellationRequested();
                        }

                        // ===== 保存文档 =====
                        string directory = System.IO.Path.GetDirectoryName(filePath) ?? "";
                        string nameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(filePath);
                        string extension = System.IO.Path.GetExtension(filePath);
                        string outputPath = System.IO.Path.Combine(directory, $"{nameWithoutExt}_formatted{extension}");

                        progress?.Report("正在保存文档...");
                        doc.SaveAs2(FileName: outputPath);

                        progress?.Report($"文档处理完成: {System.IO.Path.GetFileName(outputPath)}");
                    }
                    catch (OperationCanceledException)
                    {
                        progress?.Report("操作已取消");
                        threadException = new OperationCanceledException();
                    }
                    catch (Exception ex)
                    {
                        progress?.Report($"处理失败: {ex.Message}\n{ex.StackTrace}");
                        threadException = ex;
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            try { doc.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
                            finally
                            {
                                Marshal.ReleaseComObject(doc);
                                doc = null;
                            }
                        }

                        if (wordApp != null)
                        {
                            try { wordApp.Quit(WdSaveOptions.wdDoNotSaveChanges); } catch { }
                            finally
                            {
                                Marshal.ReleaseComObject(wordApp);
                                wordApp = null;
                            }
                        }
                    }
                });

                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();

                // 将 STA 线程内的异常传递到调用方，确保失败状态能被正确标记
                if (threadException is OperationCanceledException)
                    throw new OperationCanceledException();
                if (threadException != null)
                    throw new Exception(threadException.Message, threadException);

            }, cancellationToken);
        }

        #region 排版实现

        /// <summary>设置页边距</summary>
        private void SetPageMargins(Application wordApp, Document doc, FormattingRule rule)
        {
            foreach (Section section in doc.Sections)
            {
                var ps = section.PageSetup;
                ps.TopMargin = wordApp.CentimetersToPoints(rule.PageMargins.TopMargin);
                ps.BottomMargin = wordApp.CentimetersToPoints(rule.PageMargins.BottomMargin);
                ps.LeftMargin = wordApp.CentimetersToPoints(rule.PageMargins.LeftMargin);
                ps.RightMargin = wordApp.CentimetersToPoints(rule.PageMargins.RightMargin);
            }
        }

        /// <summary>设置 "正文" 内置样式的字体和段落属性</summary>
        private void ApplyBodyTextStyle(Document doc, FormattingRule rule)
        {
            try
            {
                var normalStyle = doc.Styles[WdBuiltinStyle.wdStyleNormal];

                // 字体设置
                normalStyle.Font.Name = rule.BodyText.ChineseFontName;
                normalStyle.Font.NameAscii = rule.BodyText.EnglishFontName;
                normalStyle.Font.NameOther = rule.BodyText.ChineseFontName;
                normalStyle.Font.Size = rule.BodyText.FontSizePoint;
                normalStyle.Font.Bold = rule.BodyText.IsBold ? -1 : 0;
                normalStyle.Font.Color = rule.BodyText.UseCustomFontColor
                    ? ParseHexToWdColor(rule.BodyText.FontColorHex)
                    : WdColor.wdColorBlack;

                // 段落格式
                var pf = normalStyle.ParagraphFormat;
                pf.CharacterUnitFirstLineIndent = rule.Paragraph.FirstLineIndentChars;
                pf.FirstLineIndent = 0; // 让 CharacterUnit 控制
                ApplyLineSpacing(pf, rule.Paragraph.LineSpacingType, rule.Paragraph.LineSpacingValue);
                pf.SpaceBeforeAuto = 0;
                pf.SpaceAfterAuto = 0;
                pf.SpaceBefore = rule.Paragraph.SpaceBeforeLines * 12f;
                pf.SpaceAfter = rule.Paragraph.SpaceAfterLines * 12f;
            }
            catch { /* 忽略样式设置错误 */ }
        }

        /// <summary>设置标题内置样式，并清除段落级直接格式确保样式生效</summary>
        private void ApplyHeadingStyles(Document doc, FormattingRule rule)
        {
            var builtinMap = new Dictionary<int, WdBuiltinStyle>
            {
                { 1, WdBuiltinStyle.wdStyleHeading1 },
                { 2, WdBuiltinStyle.wdStyleHeading2 },
                { 3, WdBuiltinStyle.wdStyleHeading3 }
            };

            if (rule.Headings == null) return;

            foreach (var heading in rule.Headings)
            {
                if (heading == null) continue;

                if (!builtinMap.TryGetValue(heading.Level, out var builtinStyle))
                    continue;

                try
                {
                    var style = doc.Styles[builtinStyle];

                    // 字体
                    style.Font.Name = heading.ChineseFontName;
                    style.Font.NameAscii = heading.EnglishFontName;
                    style.Font.NameOther = heading.ChineseFontName;
                    style.Font.Size = heading.FontSizePoint;
                    style.Font.Bold = heading.IsBold ? -1 : 0;
                    style.Font.Color = heading.UseCustomFontColor
                        ? ParseHexToWdColor(heading.FontColorHex)
                        : WdColor.wdColorBlack;

                    // 对齐
                    style.ParagraphFormat.Alignment = heading.Alignment switch
                    {
                        TextAlignment.Center => WdParagraphAlignment.wdAlignParagraphCenter,
                        TextAlignment.Right => WdParagraphAlignment.wdAlignParagraphRight,
                        TextAlignment.Justify => WdParagraphAlignment.wdAlignParagraphJustify,
                        _ => WdParagraphAlignment.wdAlignParagraphLeft
                    };

                    // 行距
                    ApplyLineSpacing(style.ParagraphFormat, heading.LineSpacingType, heading.LineSpacingValue);

                    // 段前段后（固定 6pt）
                    style.ParagraphFormat.SpaceBeforeAuto = 0;
                    style.ParagraphFormat.SpaceAfterAuto = 0;
                    style.ParagraphFormat.SpaceBefore = 6f;
                    style.ParagraphFormat.SpaceAfter = 6f;
                }
                catch { continue; }
            }

            // 遍历所有标题段落，清除直接格式（direct formatting），强制让样式定义生效
            // 不清除直接格式时，段落级覆盖会屏蔽样式修改，需手动点击样式才能刷新
            var headingStyleNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "标题 1", "标题 2", "标题 3",
                "Heading 1", "Heading 2", "Heading 3"
            };

            foreach (Paragraph para in doc.Paragraphs)
            {
                try
                {
                    var styleName = ((Style)para.get_Style()).NameLocal;
                    if (!headingStyleNames.Contains(styleName)) continue;

                    // 清除段落级直接字体格式
                    para.Range.Font.Reset();

                    // 清除段落级直接段落格式（重新应用样式即可覆盖）
                    object styleObj = para.get_Style();
                    para.set_Style(ref styleObj);

                    // 强制左缩进为 0，清除可能从正文继承的首行缩进
                    para.Format.LeftIndent = 0;
                    para.Format.FirstLineIndent = 0;
                    para.Format.CharacterUnitFirstLineIndent = 0;
                    para.Format.CharacterUnitLeftIndent = 0;
                }
                catch { continue; }
            }
        }

        /// <summary>逐段落应用正文格式（处理直接格式覆盖样式的情况）</summary>
        private void ApplyParagraphFormatting(Document doc, FormattingRule rule)
        {
            foreach (Paragraph para in doc.Paragraphs)
            {
                try
                {
                    var styleName = ((Style)para.get_Style()).NameLocal;
                    bool isNormal = styleName == "正文" || styleName == "Normal"
                                    || styleName.Contains("Body");

                    if (isNormal)
                    {
                        // 正文段落：应用正文格式
                        var range = para.Range;
                        range.Font.Name = rule.BodyText.ChineseFontName;
                        range.Font.NameAscii = rule.BodyText.EnglishFontName;
                        range.Font.Size = rule.BodyText.FontSizePoint;
                        range.Font.Bold = rule.BodyText.IsBold ? -1 : 0;
                        range.Font.Color = rule.BodyText.UseCustomFontColor
                            ? ParseHexToWdColor(rule.BodyText.FontColorHex)
                            : WdColor.wdColorBlack;

                        bool isInTable = false;
                        try { isInTable = (bool)range.Information[WdInformation.wdWithInTable]; } catch { }

                        if (isInTable)
                        {
                            para.Format.CharacterUnitFirstLineIndent = 0;
                            para.Format.FirstLineIndent = 0;
                        }
                        else
                        {
                            para.Format.CharacterUnitFirstLineIndent = rule.Paragraph.FirstLineIndentChars;
                            para.Format.FirstLineIndent = 0;
                        }
                        ApplyLineSpacing(para.Format, rule.Paragraph.LineSpacingType, rule.Paragraph.LineSpacingValue);
                        para.Format.SpaceBefore = rule.Paragraph.SpaceBeforeLines * 12f;
                        para.Format.SpaceAfter = rule.Paragraph.SpaceAfterLines * 12f;
                    }
                }
                catch { continue; }
            }
        }

        /// <summary>清除所有文字的背景色（高亮和底纹），并逐个清除表格单元格背景</summary>
        private void ClearAllTextBackground(Document doc)
        {
            try
            {
                // 清除全文高亮
                doc.Content.HighlightColorIndex = WdColorIndex.wdNoHighlight;

                // 清除全文底纹（对正文段落有效，但对表格单元格可能产生混合状态）
                doc.Content.Shading.Texture = WdTextureIndex.wdTextureNone;
                doc.Content.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                doc.Content.Shading.ForegroundPatternColor = WdColor.wdColorWhite;
            }
            catch { }

            // 逐个单元格清除背景色，避免 doc.Content.Shading 对表格产生脏状态
            foreach (Table table in doc.Tables)
            {
                try
                {
                    foreach (Cell cell in table.Range.Cells)
                    {
                        try
                        {
                            cell.Shading.Texture = WdTextureIndex.wdTextureNone;
                            cell.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                            cell.Shading.ForegroundPatternColor = WdColor.wdColorWhite;
                        }
                        catch { continue; }
                    }
                }
                catch { continue; }
            }
        }

        /// <summary>格式化所有表格</summary>
        private void FormatAllTables(Document doc, FormattingRule rule)
        {
            foreach (Table table in doc.Tables)
            {
                try
                {
                    // ── 表格自动适应窗口宽度 ──
                    table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

                    // ── 设置边框 ──
                    ApplyTableBorders(table, rule.Table);

                    // ── 设置单元格对齐 ──
                    foreach (Cell cell in table.Range.Cells)
                    {
                        try
                        {
                            // 垂直对齐
                            cell.VerticalAlignment = rule.Table.CellVerticalAlignment switch
                            {
                                CellVerticalAlign.Top => WdCellVerticalAlignment.wdCellAlignVerticalTop,
                                CellVerticalAlign.Bottom => WdCellVerticalAlignment.wdCellAlignVerticalBottom,
                                _ => WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            };

                            // 水平对齐
                            cell.Range.ParagraphFormat.Alignment = rule.Table.CellHorizontalAlignment switch
                            {
                                CellHorizontalAlign.Left => WdParagraphAlignment.wdAlignParagraphLeft,
                                CellHorizontalAlign.Right => WdParagraphAlignment.wdAlignParagraphRight,
                                _ => WdParagraphAlignment.wdAlignParagraphCenter
                            };

                            // 字体（与正文一致）
                            if (rule.Table.UseSameAsBody)
                            {
                                cell.Range.Font.Name = rule.BodyText.ChineseFontName;
                                cell.Range.Font.NameAscii = rule.BodyText.EnglishFontName;
                                cell.Range.Font.Size = rule.BodyText.FontSizePoint;
                                cell.Range.Font.Bold = 0; // 默认不加粗
                                cell.Range.Font.Color = rule.BodyText.UseCustomFontColor
                                    ? ParseHexToWdColor(rule.BodyText.FontColorHex)
                                    : WdColor.wdColorBlack;
                            }

                            // 段落格式
                            cell.Range.ParagraphFormat.SpaceBefore = rule.Table.SpaceBeforeLines * 12f;
                            cell.Range.ParagraphFormat.SpaceAfter = rule.Table.SpaceAfterLines * 12f;
                            ApplyLineSpacing(cell.Range.ParagraphFormat, rule.Table.LineSpacingType, rule.Table.LineSpacingValue);

                            // 清除首行缩进
                            cell.Range.ParagraphFormat.FirstLineIndent = 0;
                            cell.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                        }
                        catch { continue; }
                    }

                    // ── 表头加粗 ──
                    if (rule.Table.HeaderBold && table.Rows.Count > 0)
                    {
                        try
                        {
                            var headerRow = table.Rows[1];
                            headerRow.Range.Font.Bold = -1;  // True

                            // 跨页重复表头
                            if (rule.Table.RepeatHeaderRow)
                            {
                                headerRow.HeadingFormat = -1; // True
                            }
                        }
                        catch { /* 忽略 */ }
                    }
                }
                catch { continue; }
            }
        }

        /// <summary>应用表格边框</summary>
        private void ApplyTableBorders(Table table, TableSettings ts)
        {
            if (ts.BorderStyle == TableBorderStyle.None)
            {
                table.Borders.Enable = 0;
                return;
            }

            WdLineStyle lineStyle = ts.BorderStyle switch
            {
                TableBorderStyle.SingleThick => WdLineStyle.wdLineStyleSingle,
                TableBorderStyle.Double => WdLineStyle.wdLineStyleDouble,
                _ => WdLineStyle.wdLineStyleSingle
            };

            WdLineWidth lineWidth = ts.BorderStyle switch
            {
                TableBorderStyle.SingleThick => WdLineWidth.wdLineWidth150pt,
                _ => WdLineWidth.wdLineWidth050pt
            };

            WdColor borderColor = WdColor.wdColorBlack;

            // 设置所有六个边框
            var borderTypes = new[]
            {
                WdBorderType.wdBorderTop,
                WdBorderType.wdBorderBottom,
                WdBorderType.wdBorderLeft,
                WdBorderType.wdBorderRight,
                WdBorderType.wdBorderHorizontal,
                WdBorderType.wdBorderVertical
            };

            foreach (var borderType in borderTypes)
            {
                try
                {
                    var border = table.Borders[borderType];
                    border.LineStyle = lineStyle;
                    border.LineWidth = lineWidth;
                    border.Color = borderColor;
                }
                catch { continue; }
            }
        }

        /// <summary>应用行距设置</summary>
        private void ApplyLineSpacing(ParagraphFormat pf, LineSpacingType type, float value)
        {
            switch (type)
            {
                case LineSpacingType.Single:
                    pf.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                    break;
                case LineSpacingType.OneAndHalf:
                    pf.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;
                    break;
                case LineSpacingType.Double:
                    pf.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble;
                    break;
                case LineSpacingType.Multiple:
                    pf.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                    pf.LineSpacing = value * 12f; // Word 内部用 12pt 为 1 倍
                    break;
                case LineSpacingType.Fixed:
                    pf.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                    pf.LineSpacing = value; // 直接用磅值
                    break;
                case LineSpacingType.AtLeast:
                    pf.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast;
                    pf.LineSpacing = value;
                    break;
            }
        }

        /// <summary>将十六进制颜色字符串转换为 WdColor（Word COM 使用 BGR 格式）</summary>
        private WdColor ParseHexToWdColor(string? hex)
        {
            if (string.IsNullOrWhiteSpace(hex)) return WdColor.wdColorBlack;

            try
            {
                string cleanHex = hex.Trim().TrimStart('#');
                if (cleanHex.Length == 6)
                {
                    int r = Convert.ToInt32(cleanHex.Substring(0, 2), 16);
                    int g = Convert.ToInt32(cleanHex.Substring(2, 2), 16);
                    int b = Convert.ToInt32(cleanHex.Substring(4, 2), 16);
                    // Word COM 使用 BGR 格式: (B << 16) | (G << 8) | R
                    return (WdColor)((b << 16) | (g << 8) | r);
                }
            }
            catch { }
            return WdColor.wdColorBlack;
        }

        #endregion

        #region 前置页处理

        /// <summary>从模板提取前4页并插入到目标文档开头</summary>
        private void InsertFrontMatterPages(Application wordApp, Document targetDoc, string templatePath)
        {
            if (!System.IO.File.Exists(templatePath)) return;

            Document? templateDoc = null;
            try
            {
                templateDoc = wordApp.Documents.Open(
                    FileName: templatePath,
                    ReadOnly: true,
                    Visible: false);

                // 获取前4页的 Range
                object what = WdGoToItem.wdGoToPage;
                object which = WdGoToDirection.wdGoToAbsolute;
                object count = 5; // 跳转到第5页开头
                object missing = Type.Missing;

                Microsoft.Office.Interop.Word.Range startRange = templateDoc.Range(0, 0);
                Microsoft.Office.Interop.Word.Range endRange = templateDoc.GoTo(ref what, ref which, ref count, ref missing);

                // 如果文档不到5页，endRange可能停留在开头或末尾
                if (endRange.Start <= startRange.Start) 
                {
                    endRange = templateDoc.Content;
                    endRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                }

                Microsoft.Office.Interop.Word.Range copyRange = templateDoc.Range(startRange.Start, endRange.Start);
                copyRange.Copy();

                // 粘贴到目标文档开头
                Microsoft.Office.Interop.Word.Range targetRange = targetDoc.Range(0, 0);
                targetRange.InsertBreak(Type: WdBreakType.wdSectionBreakNextPage);
                
                Microsoft.Office.Interop.Word.Range pasteRange = targetDoc.Range(0, 0);
                pasteRange.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting);
            }
            catch { }
            finally
            {
                if (templateDoc != null)
                {
                    object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                    object missing = Type.Missing;
                    templateDoc.Close(ref saveChanges, ref missing, ref missing);
                }
            }
        }

        #endregion

        #region 资源释放

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                // Word Application 实例由每次 ProcessDocumentAsync 在 STA 线程内独立创建并释放，
                // 这里不再持有跨调用的 COM 引用
                _disposed = true;
            }
        }

        ~WordProcessingService()
        {
            Dispose(false);
        }

        #endregion
    }
}
