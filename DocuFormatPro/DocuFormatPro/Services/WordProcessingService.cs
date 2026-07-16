using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
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
                        var stepTimer = Stopwatch.StartNew();
                        void ReportStepDone(string stepName)
                        {
                            progress?.Report($"{stepName}完成，用时 {stepTimer.Elapsed.TotalSeconds:F1} 秒");
                            stepTimer.Restart();
                        }

                        cancellationToken.ThrowIfCancellationRequested();

                        // 每次处理在当前 STA 线程内创建独立的 Word Application 实例
                        // 避免跨线程复用 RCW 导致 "COM object that has been separated from its underlying RCW" 错误
                        wordApp = new Application
                        {
                            Visible = false,
                            ScreenUpdating = false,
                            DisplayAlerts = WdAlertLevel.wdAlertsNone
                        };
                        try
                        {
                            wordApp.Options.CheckSpellingAsYouType = false;
                            wordApp.Options.CheckGrammarAsYouType = false;
                            wordApp.Options.Pagination = false;
                        }
                        catch { }

                        progress?.Report($"正在打开文档: {System.IO.Path.GetFileName(filePath)}");

                        doc = wordApp.Documents.Open(
                            FileName: filePath,
                            ReadOnly: false,
                            Visible: false);
                        try
                        {
                            doc.ShowSpellingErrors = false;
                            doc.ShowGrammaticalErrors = false;
                        }
                        catch { }
                        string inputExtension = System.IO.Path.GetExtension(filePath);
                        bool isDocxDocument = string.Equals(inputExtension, ".docx", StringComparison.OrdinalIgnoreCase);

                        cancellationToken.ThrowIfCancellationRequested();

                        // ===== 1. 设置页边距 =====
                        progress?.Report("正在标准化页边距...");
                        SetPageMargins(wordApp, doc, rule);
                        ReportStepDone("页边距");
                        cancellationToken.ThrowIfCancellationRequested();

                        // ===== 2. 设置正文样式 =====
                        progress?.Report("正在设置正文格式...");
                        ApplyBodyTextStyle(doc, rule);
                        ReportStepDone("正文样式");
                        cancellationToken.ThrowIfCancellationRequested();

                        // ===== 3. 设置标题样式 =====
                        progress?.Report("正在设置标题格式...");
                        ApplyHeadingStyles(doc, rule, resetDirectFormatting: !isDocxDocument);
                        ReportStepDone("标题样式");
                        cancellationToken.ThrowIfCancellationRequested();

                        // ===== 3b. 标题自动编号 =====
                        if (rule.HeadingNumbering.EnableNumbering)
                        {
                            progress?.Report("正在处理标题编号...");
                            ApplyHeadingNumbering(doc, rule.HeadingNumbering);
                            ReportStepDone("标题编号");
                            cancellationToken.ThrowIfCancellationRequested();
                        }

                        // ===== 4. 逐段落应用格式 =====
                        progress?.Report("正在处理段落格式...");
                        // ===== 5. 格式化表格 =====
                        if (!isDocxDocument)
                        {
                            ApplyParagraphFormatting(doc, rule);
                            ReportStepDone("段落格式");
                            cancellationToken.ThrowIfCancellationRequested();
                        }

                        if (rule.Table.ApplyTableFormatting)
                        {
                            progress?.Report("正在重塑表格...");
                            FormatAllTablesFast(doc, rule, isDocxDocument);
                            ReportStepDone("表格格式");
                            cancellationToken.ThrowIfCancellationRequested();
                        }

                        // ===== 6. 清除所有文字背景色 =====
                        progress?.Report("正在清除文字背景色...");
                        // 如果启用了首行底色，跳过首行单元格，避免清除后再设置的竞争问题
                        bool skipFirstRow = rule.Table.ApplyTableFormatting && rule.Table.UseHeaderShading;
                        if (!isDocxDocument)
                            ClearAllTextBackground(doc, skipFirstRow);
                        ReportStepDone("背景清理");
                        cancellationToken.ThrowIfCancellationRequested();

                        // ===== 6b. 重新应用首行底色（在清除背景后执行，避免被覆盖）=====
                        if (rule.Table.ApplyTableFormatting && rule.Table.UseHeaderShading && !isDocxDocument)
                        {
                            progress?.Report("正在设置表格首行底色...");
                            ApplyTableHeaderShading(doc, rule.Table);
                            ReportStepDone("表头底色");
                        }

                        // ===== 6b. 规范化正文文本 =====
                        if (rule.NormalizeBodyText && !isDocxDocument)
                        {
                            var normalizer = new TextNormalizationService();
                            normalizer.NormalizeBodyText(
                                doc,
                                rule.BodyText.ChineseFontName,
                                rule.BodyText.EnglishFontName,
                                progress);
                            normalizer.NormalizeTableText(
                                doc,
                                rule.Table.ChineseFontName,
                                rule.Table.EnglishFontName,
                                progress);
                            ReportStepDone("文本规范化");
                            cancellationToken.ThrowIfCancellationRequested();
                        }

                        // ===== 7. 处理表格和图片题注 =====
                        if (rule.Table.ApplyTableCaptions)
                        {
                            progress?.Report("正在处理题注...");
                            var captionService = new CaptionService();
                            captionService.ProcessCaptions(doc, progress);
                            ReportStepDone("题注");
                            cancellationToken.ThrowIfCancellationRequested();
                        }

                        // ===== 8. 插入前置页 =====
                        if (rule.FrontMatter.InsertFrontMatter && !string.IsNullOrWhiteSpace(rule.FrontMatter.TemplateFilePath))
                        {
                            progress?.Report("正在插入封面与前置页...");
                            InsertFrontMatterPages(wordApp, doc, rule.FrontMatter.TemplateFilePath);
                            ReportStepDone("前置页");
                            cancellationToken.ThrowIfCancellationRequested();
                        }

                        // ===== 保存文档 =====
                        string directory = System.IO.Path.GetDirectoryName(filePath) ?? "";
                        string nameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(filePath);
                        string extension = System.IO.Path.GetExtension(filePath);
                        string outputPath = System.IO.Path.Combine(directory, $"{nameWithoutExt}_formatted{extension}");

                        progress?.Report("正在保存文档...");
                        doc.SaveAs2(FileName: outputPath);
                        ReportStepDone("保存");
                        if (string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase))
                        {
                            doc.Close(WdSaveOptions.wdDoNotSaveChanges);
                            Marshal.ReleaseComObject(doc);
                            doc = null;

                            EnsureDocxPostProcessing(
                                outputPath,
                                rule,
                                normalizeBodyText: rule.NormalizeBodyText);
                            ReportStepDone("docx 快速后处理");
                        }

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
        private void ApplyHeadingStyles(Document doc, FormattingRule rule, bool resetDirectFormatting = true)
        {
            var builtinMap = new Dictionary<int, WdBuiltinStyle>
            {
                { 1, WdBuiltinStyle.wdStyleHeading1 },
                { 2, WdBuiltinStyle.wdStyleHeading2 },
                { 3, WdBuiltinStyle.wdStyleHeading3 },
                { 4, WdBuiltinStyle.wdStyleHeading4 },
                { 5, WdBuiltinStyle.wdStyleHeading5 },
                { 6, WdBuiltinStyle.wdStyleHeading6 },
                { 7, WdBuiltinStyle.wdStyleHeading7 },
                { 8, WdBuiltinStyle.wdStyleHeading8 },
                { 9, WdBuiltinStyle.wdStyleHeading9 }
            };

            if (rule.Headings == null) return;

            for (int headingLevel = 1; headingLevel <= 9; headingLevel++)
            {
                var heading = ResolveHeadingFormat(rule, headingLevel);
                if (heading == null) continue;

                if (!builtinMap.TryGetValue(headingLevel, out var builtinStyle))
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

                    // 段前段后（按标题配置的磅值）
                    style.ParagraphFormat.SpaceBeforeAuto = 0;
                    style.ParagraphFormat.SpaceAfterAuto = 0;
                    style.ParagraphFormat.SpaceBefore = heading.SpaceBeforePoints;
                    style.ParagraphFormat.SpaceAfter = heading.SpaceAfterPoints;
                    style.ParagraphFormat.FirstLineIndent = 0;
                    style.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                    style.ParagraphFormat.LeftIndent = 0;
                    style.ParagraphFormat.CharacterUnitLeftIndent = 0;
                }
                catch { continue; }
            }

            // 遍历所有标题段落，清除直接格式（direct formatting），强制让样式定义生效
            // 不清除直接格式时，段落级覆盖会屏蔽样式修改，需手动点击样式才能刷新
            if (!resetDirectFormatting) return;

            var headingStyleNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "标题 1", "标题 2", "标题 3", "标题 4", "标题 5", "标题 6", "标题 7", "标题 8", "标题 9",
                "Heading 1", "Heading 2", "Heading 3", "Heading 4", "Heading 5", "Heading 6", "Heading 7", "Heading 8", "Heading 9"
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

        /// <summary>为标题样式绑定多级列表，实现自动编号（新增/移动标题后编号自动更新）</summary>
        private DocuFormatPro.Models.HeadingStyle? ResolveHeadingFormat(FormattingRule rule, int level)
        {
            if (rule.Headings == null || rule.Headings.Count == 0) return null;

            DocuFormatPro.Models.HeadingStyle? exact = rule.Headings.FirstOrDefault(h => h.Level == level);
            if (exact != null) return exact;

            int fallbackLevel = level >= 4 ? 3 : level;
            return rule.Headings.FirstOrDefault(h => h.Level == fallbackLevel)
                ?? rule.Headings.ElementAtOrDefault(Math.Max(0, fallbackLevel - 1))
                ?? rule.Headings.LastOrDefault();
        }

        private static string BuildNumberFormat(int level)
            => string.Join(".", Enumerable.Range(1, Math.Clamp(level, 1, 9)).Select(i => $"%{i}"));

        private static int GetHeadingLevelFromStyleName(string? styleName)
        {
            if (string.IsNullOrWhiteSpace(styleName)) return 0;

            var match = Regex.Match(styleName, @"(?:Heading|标题)\s*([1-9])", RegexOptions.IgnoreCase);
            return match.Success ? int.Parse(match.Groups[1].Value) : 0;
        }

        private void ApplyHeadingNumbering(Document doc, HeadingNumberingSettings settings)
        {
            var headingStyleNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "标题 1", "标题 2", "标题 3", "标题 4", "标题 5", "标题 6", "标题 7", "标题 8", "标题 9",
                "Heading 1", "Heading 2", "Heading 3", "Heading 4", "Heading 5", "Heading 6", "Heading 7", "Heading 8", "Heading 9"
            };

            // 如果选择去除现有编号前缀，先剥离文字中的旧编号
            if (settings.StripExistingNumbers)
            {
                var stripPattern = new Regex(
                    @"^(\d+(\.\d+)*\.?\s*|第[一二三四五六七八九十百\d]+[章节条]\s*|[一二三四五六七八九十]+[、.]\s*|（[一二三四五六七八九十]+）\s*)",
                    RegexOptions.None);

                foreach (Paragraph para in doc.Paragraphs)
                {
                    try
                    {
                        var styleName = ((Style)para.get_Style()).NameLocal;
                        if (!headingStyleNames.Contains(styleName)) continue;

                        var r = para.Range;
                        object unit = WdUnits.wdCharacter;
                        object cnt = -1;
                        r.MoveEnd(ref unit, ref cnt);
                        string currentText = r.Text ?? "";
                        string stripped = stripPattern.Replace(currentText, "").TrimStart();
                        if (stripped != currentText)
                            r.Text = stripped;
                    }
                    catch { continue; }
                }
            }

            try
            {
                // 创建多级列表模板
                var lt = doc.ListTemplates.Add(OutlineNumbered: true);

                for (int level = 1; level <= 9; level++)
                {
                    var ll = lt.ListLevels[level];
                    ll.StartAt = 1;

                    switch (settings.Scheme)
                    {
                        case HeadingNumberingScheme.Numeric:
                            ll.NumberFormat = level switch
                            {
                                1 => "%1",
                                2 => "%1.%2",
                                _ => BuildNumberFormat(level)
                            };
                            ll.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic;
                            ll.TrailingCharacter = WdTrailingCharacter.wdTrailingSpace;
                            break;

                        case HeadingNumberingScheme.ChapterNumeric:
                            if (level == 1)
                            {
                                ll.NumberFormat = "第%1章";
                                ll.NumberStyle = WdListNumberStyle.wdListNumberStyleSimpChinNum2;
                            }
                            else
                            {
                                ll.NumberFormat = BuildNumberFormat(level);
                                ll.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic;
                            }
                            ll.TrailingCharacter = WdTrailingCharacter.wdTrailingSpace;
                            break;

                        case HeadingNumberingScheme.Traditional:
                            if (level == 1)
                            {
                                ll.NumberFormat = "%1、";
                                ll.NumberStyle = WdListNumberStyle.wdListNumberStyleSimpChinNum2;
                            }
                            else if (level == 2)
                            {
                                ll.NumberFormat = "（%2）";
                                ll.NumberStyle = WdListNumberStyle.wdListNumberStyleSimpChinNum2;
                            }
                            else if (level == 3)
                            {
                                ll.NumberFormat = "%3.";
                                ll.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic;
                            }
                            else
                            {
                                ll.NumberFormat = BuildNumberFormat(level);
                                ll.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic;
                            }
                            ll.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
                            break;
                    }

                    ll.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
                    ll.NumberPosition = 0;
                    ll.TextPosition = 0;
                    ll.TabPosition = 0;
                }

                // 对每个标题段落分别应用对应级别的列表
                foreach (Paragraph para in doc.Paragraphs)
                {
                    try
                    {
                        var styleName = ((Style)para.get_Style()).NameLocal;
                        if (!headingStyleNames.Contains(styleName)) continue;

                        // 根据样式名确定级别
                        int headingLevel = GetHeadingLevelFromStyleName(styleName);
                        if (headingLevel == 0) continue;

                        para.Range.ListFormat.ApplyListTemplateWithLevel(
                            lt,
                            ContinuePreviousList: true,
                            ApplyTo: WdListApplyTo.wdListApplyToWholeList,
                            DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior,
                            ApplyLevel: headingLevel);
                    }
                    catch { continue; }
                }

                // 更新域以刷新编号显示
                try { doc.Fields.Update(); } catch { }
            }
            catch { /* 多级列表创建失败，静默忽略 */ }
        }

        /// <summary>获取文档中对应级别的标题样式名称（中文或英文）</summary>
        private string GetHeadingStyleName(Document doc, int level)
        {
            try
            {
                var builtinStyle = level switch
                {
                    1 => WdBuiltinStyle.wdStyleHeading1,
                    2 => WdBuiltinStyle.wdStyleHeading2,
                    3 => WdBuiltinStyle.wdStyleHeading3,
                    4 => WdBuiltinStyle.wdStyleHeading4,
                    5 => WdBuiltinStyle.wdStyleHeading5,
                    6 => WdBuiltinStyle.wdStyleHeading6,
                    7 => WdBuiltinStyle.wdStyleHeading7,
                    8 => WdBuiltinStyle.wdStyleHeading8,
                    9 => WdBuiltinStyle.wdStyleHeading9,
                    _ => WdBuiltinStyle.wdStyleHeading3
                };
                return doc.Styles[builtinStyle].NameLocal;
            }
            catch
            {
                return $"Heading {level}";
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
                    if (GetHeadingLevelFromStyleName(styleName) > 0) continue;

                    var range = para.Range;
                    bool isInTable = false;
                    try { isInTable = (bool)range.Information[WdInformation.wdWithInTable]; } catch { }
                    if (isInTable) continue;

                    bool isNormal = styleName == "正文" || styleName == "Normal"
                                    || styleName.Contains("Body");

                    if (isNormal)
                    {
                        // 正文段落：应用正文格式
                        range.Font.Name = rule.BodyText.ChineseFontName;
                        range.Font.NameAscii = rule.BodyText.EnglishFontName;
                        range.Font.Size = rule.BodyText.FontSizePoint;
                        range.Font.Bold = rule.BodyText.IsBold ? -1 : 0;
                        range.Font.Color = rule.BodyText.UseCustomFontColor
                            ? ParseHexToWdColor(rule.BodyText.FontColorHex)
                            : WdColor.wdColorBlack;

                        para.Format.CharacterUnitFirstLineIndent = rule.Paragraph.FirstLineIndentChars;
                        para.Format.FirstLineIndent = 0;
                        ApplyLineSpacing(para.Format, rule.Paragraph.LineSpacingType, rule.Paragraph.LineSpacingValue);
                        para.Format.SpaceBefore = rule.Paragraph.SpaceBeforeLines * 12f;
                        para.Format.SpaceAfter = rule.Paragraph.SpaceAfterLines * 12f;
                    }
                }
                catch { continue; }
            }
        }

        /// <summary>清除所有文字的背景色（高亮和底纹），并逐个清除表格单元格背景</summary>
        private void ClearAllTextBackground(Document doc, bool skipFirstRowOfTables = false)
        {
            try
            {
                // 清除全文高亮
                doc.Content.HighlightColorIndex = WdColorIndex.wdNoHighlight;

                // 清除全文底纹
                ClearShading(doc.Content.Shading);
            }
            catch { }

            // 逐个单元格清除背景色，跳过需要保留底色的首行
            foreach (Table table in doc.Tables)
            {
                try
                {
                    foreach (Cell cell in table.Range.Cells)
                    {
                        try
                        {
                            if (skipFirstRowOfTables)
                            {
                                bool isFirstRow = false;
                                try { isFirstRow = cell.RowIndex == 1; } catch { }
                                if (isFirstRow) continue;
                            }

                            ClearShading(cell.Shading);
                            ClearShading(cell.Range.Shading);
                        }
                        catch { continue; }
                    }
                }
                catch { continue; }
            }
        }

        /// <summary>格式化所有表格</summary>
        private void FormatAllTablesFast(Document doc, FormattingRule rule, bool deferParagraphIndentReset)
        {
            foreach (Table table in doc.Tables)
            {
                try
                {
                    if (deferParagraphIndentReset)
                    {
                        try
                        {
                            table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
                            table.PreferredWidth = 100f;
                        }
                        catch { }
                    }
                    else
                    {
                        table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
                    }

                    try
                    {
                        table.ApplyStyleHeadingRows = false;
                        table.ApplyStyleFirstColumn = false;
                        table.ApplyStyleLastColumn = false;
                        table.ApplyStyleRowBands = false;
                        table.ApplyStyleColumnBands = false;
                    }
                    catch { }

                    try
                    {
                        table.Rows.HeightRule = WdRowHeightRule.wdRowHeightAtLeast;
                        table.Rows.Height = (float)(1.0 * 28.35);
                    }
                    catch { }

                    ApplyTableBorders(table, rule.Table);
                    ApplyTableRangeFormatting(table, rule.Table);
                    if (!deferParagraphIndentReset)
                        ResetTableParagraphIndents(table);

                    if (table.Rows.Count > 0)
                    {
                        try
                        {
                            Row headerRow = table.Rows[1];
                            if (rule.Table.HeaderBold)
                                headerRow.Range.Font.Bold = -1;

                            if (rule.Table.RepeatHeaderRow)
                                headerRow.HeadingFormat = -1;
                        }
                        catch { }
                    }
                }
                catch { continue; }
            }
        }

        private void FormatAllTables(Document doc, FormattingRule rule)
        {
            foreach (Table table in doc.Tables)
            {
                try
                {
                    // ── 表格自动适应窗口宽度 ──
                    table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

                    // ── 清除表格样式的行列格式标志，避免样式覆盖我们设置的底纹 ──
                    try
                    {
                        table.ApplyStyleHeadingRows = false;
                        table.ApplyStyleFirstColumn = false;
                        table.ApplyStyleLastColumn = false;
                        table.ApplyStyleRowBands = false;
                        table.ApplyStyleColumnBands = false;
                    }
                    catch { }

                    // ── 所有行行高：最小值 1 厘米 ──
                    try
                    {
                        table.Rows.HeightRule = WdRowHeightRule.wdRowHeightAtLeast;
                        table.Rows.Height = (float)(1.0 * 28.35);
                    }
                    catch { }

                    Row row = table.Rows[1];
                    for (int rowIndex = 0; rowIndex < 0; rowIndex++)
                    {
                        try
                        {
                            row.HeightRule = WdRowHeightRule.wdRowHeightAtLeast;
                            row.Height = (float)(1.0 * 28.35); // 1cm ≈ 28.35pt
                        }
                        catch { }
                    }

                    // ── 设置边框 ──
                    ApplyTableBorders(table, rule.Table);
                    ApplyTableRangeFormatting(table, rule.Table);

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

                            // 字体（使用表格独立配置）
                            cell.Range.Font.Name = rule.Table.ChineseFontName;
                            cell.Range.Font.NameAscii = rule.Table.EnglishFontName;
                            cell.Range.Font.Size = rule.Table.FontSizePoint;
                            cell.Range.Font.Bold = 0;
                            cell.Range.Font.Color = WdColor.wdColorBlack;

                            // 段落：不缩进，水平居中，单元格垂直居中（默认）
                            cell.Range.ParagraphFormat.SpaceBefore = rule.Table.SpaceBeforeLines * 12f;
                            cell.Range.ParagraphFormat.SpaceAfter = rule.Table.SpaceAfterLines * 12f;
                            ApplyLineSpacing(cell.Range.ParagraphFormat, LineSpacingType.Single, 1f);

                            // 清除首行缩进（默认行为）
                            cell.Range.ParagraphFormat.FirstLineIndent = 0;
                            cell.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                            cell.Range.ParagraphFormat.LeftIndent = 0;
                            cell.Range.ParagraphFormat.CharacterUnitLeftIndent = 0;
                        }
                        catch { continue; }
                    }

                    // ── 表头加粗 + 底色 ──
                    if (table.Rows.Count > 0)
                    {
                        try
                        {
                            var headerRow = table.Rows[1];

                            if (rule.Table.HeaderBold)
                                headerRow.Range.Font.Bold = -1;

                            if (rule.Table.RepeatHeaderRow)
                                headerRow.HeadingFormat = -1;

                            // 首行底色（在 ClearAllTextBackground 之后单独调用，此处移除）
                        }
                        catch { /* 忽略 */ }
                    }
                }
                catch { continue; }
            }
        }

        /// <summary>在清除背景色之后，单独为所有表格首行重新应用底色</summary>
        private void ApplyTableRangeFormatting(Table table, TableSettings ts)
        {
            var tableRange = table.Range;

            tableRange.Font.Name = ts.ChineseFontName;
            tableRange.Font.NameAscii = ts.EnglishFontName;
            tableRange.Font.Size = ts.FontSizePoint;
            tableRange.Font.Bold = 0;
            tableRange.Font.Color = WdColor.wdColorBlack;

            var pf = tableRange.ParagraphFormat;
            pf.Alignment = ts.CellHorizontalAlignment switch
            {
                CellHorizontalAlign.Left => WdParagraphAlignment.wdAlignParagraphLeft,
                CellHorizontalAlign.Right => WdParagraphAlignment.wdAlignParagraphRight,
                _ => WdParagraphAlignment.wdAlignParagraphCenter
            };
            pf.SpaceBefore = ts.SpaceBeforeLines * 12f;
            pf.SpaceAfter = ts.SpaceAfterLines * 12f;
            pf.FirstLineIndent = 0;
            pf.CharacterUnitFirstLineIndent = 0;
            pf.LeftIndent = 0;
            pf.CharacterUnitLeftIndent = 0;
            ApplyLineSpacing(pf, LineSpacingType.Single, 1f);

            WdCellVerticalAlignment verticalAlignment = ts.CellVerticalAlignment switch
            {
                CellVerticalAlign.Top => WdCellVerticalAlignment.wdCellAlignVerticalTop,
                CellVerticalAlign.Bottom => WdCellVerticalAlignment.wdCellAlignVerticalBottom,
                _ => WdCellVerticalAlignment.wdCellAlignVerticalCenter
            };

            try
            {
                table.Range.Cells.VerticalAlignment = verticalAlignment;
            }
            catch
            {
                foreach (Cell cell in table.Range.Cells)
                {
                    try { cell.VerticalAlignment = verticalAlignment; }
                    catch { }
                }
            }
        }

        private void ResetTableParagraphIndents(Table table)
        {
            foreach (Paragraph paragraph in table.Range.Paragraphs)
            {
                try
                {
                    paragraph.Format.LeftIndent = 0;
                    paragraph.Format.CharacterUnitLeftIndent = 0;
                    paragraph.Format.FirstLineIndent = 0;
                    paragraph.Format.CharacterUnitFirstLineIndent = 0;
                }
                catch { }
            }
        }

        private void ApplyTableHeaderShading(Document doc, TableSettings ts)
        {
            WdColor fillColor = ParseHexToWdColor(ts.HeaderShadingColorHex);

            foreach (Table table in doc.Tables)
            {
                try
                {
                    if (table.Rows.Count == 0) continue;

                    // 清除表格样式标志，避免 Word 内置表格样式覆盖我们设置的底色
                    try
                    {
                        table.ApplyStyleHeadingRows = false;
                        table.ApplyStyleFirstColumn = false;
                        table.ApplyStyleLastColumn = false;
                        table.ApplyStyleRowBands = false;
                        table.ApplyStyleColumnBands = false;
                    }
                    catch { }

                    Row headerRow = table.Rows[1];
                    ApplyHeaderRowShading(headerRow, fillColor);

                    foreach (Cell cell in headerRow.Cells)
                    {
                        try
                        {
                            // 先清除高亮
                            cell.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight;

                            // 设置单元格底色
                            ApplySolidShading(cell.Shading, fillColor);
                            cell.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight;
                            ApplySolidShading(cell.Shading, fillColor);
                            ApplySolidShading(cell.Range.Shading, fillColor);
                            ApplySolidShading(cell.Range.Font.Shading, fillColor);

                            foreach (Paragraph paragraph in cell.Range.Paragraphs)
                            {
                                try
                                {
                                    ApplySolidShading(paragraph.Range.Shading, fillColor);
                                    ApplySolidShading(paragraph.Range.Font.Shading, fillColor);
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }
                }
                catch { continue; }
            }
        }

        /// <summary>应用表格边框</summary>
        private void ApplyHeaderRowShading(Row headerRow, WdColor fillColor)
        {
            foreach (Cell cell in headerRow.Cells)
            {
                try
                {
                    cell.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight;

                    ApplySolidShading(cell.Shading, fillColor);
                    ApplySolidShading(cell.Range.Shading, fillColor);
                    ApplySolidShading(cell.Range.Font.Shading, fillColor);

                    foreach (Paragraph paragraph in cell.Range.Paragraphs)
                    {
                        try
                        {
                            paragraph.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight;
                            ApplySolidShading(paragraph.Range.Shading, fillColor);
                            ApplySolidShading(paragraph.Range.Font.Shading, fillColor);
                        }
                        catch { }
                    }
                }
                catch { }
            }
        }

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
        private static void ClearShading(Shading shading)
        {
            shading.Texture = WdTextureIndex.wdTextureNone;
            shading.BackgroundPatternColor = WdColor.wdColorWhite;
            shading.ForegroundPatternColor = WdColor.wdColorWhite;
        }

        private static void ApplySolidShading(Shading shading, WdColor fillColor)
        {
            shading.Texture = WdTextureIndex.wdTextureNone;
            shading.BackgroundPatternColor = fillColor;
            shading.ForegroundPatternColor = fillColor;
        }

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
        private void EnsureDocxPostProcessing(
            string docxPath,
            FormattingRule rule,
            bool normalizeBodyText)
        {
            if (!System.IO.File.Exists(docxPath)) return;

            string fillHex = NormalizeHexColor(rule.Table.HeaderShadingColorHex);

            using ZipArchive archive = ZipFile.Open(docxPath, ZipArchiveMode.Update);
            ZipArchiveEntry? documentEntry = archive.GetEntry("word/document.xml");
            if (documentEntry == null) return;

            XDocument documentXml;
            using (Stream stream = documentEntry.Open())
            {
                documentXml = XDocument.Load(stream);
            }

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XName tblName = w + "tbl";
            XName trName = w + "tr";
            XName tcName = w + "tc";
            XName tcPrName = w + "tcPr";
            XName pName = w + "p";
            XName pPrName = w + "pPr";
            XName indName = w + "ind";
            XName shdName = w + "shd";
            XName highlightName = w + "highlight";

            XName valAttr = w + "val";
            XName colorAttr = w + "color";
            XName fillAttr = w + "fill";
            XName leftAttr = w + "left";
            XName leftCharsAttr = w + "leftChars";
            XName firstLineAttr = w + "firstLine";
            XName firstLineCharsAttr = w + "firstLineChars";
            XName hangingAttr = w + "hanging";
            XName hangingCharsAttr = w + "hangingChars";

            foreach (XElement highlight in documentXml.Descendants(highlightName).ToList())
                highlight.Remove();

            foreach (XElement shading in documentXml.Descendants(shdName).ToList())
                shading.Remove();

            var styleNames = BuildDocxStyleNameMap(archive, w);
            ApplyDocxParagraphFormatting(documentXml, w, styleNames, rule, normalizeBodyText);

            if (!rule.Table.ApplyTableFormatting)
            {
                RewriteDocumentXmlEntry(archive, documentEntry, documentXml);
                return;
            }

            foreach (XElement table in documentXml.Descendants(tblName))
            {
                foreach (XElement paragraph in table.Descendants(pName))
                {
                    XElement? pPr = paragraph.Element(pPrName);
                    if (pPr == null)
                    {
                        pPr = new XElement(pPrName);
                        paragraph.AddFirst(pPr);
                    }

                    XElement? indent = pPr.Element(indName);
                    if (indent == null)
                    {
                        indent = new XElement(indName);
                        pPr.Add(indent);
                    }

                    indent.SetAttributeValue(leftAttr, "0");
                    indent.SetAttributeValue(leftCharsAttr, "0");
                    indent.SetAttributeValue(firstLineAttr, "0");
                    indent.SetAttributeValue(firstLineCharsAttr, "0");
                    indent.SetAttributeValue(hangingAttr, "0");
                    indent.SetAttributeValue(hangingCharsAttr, "0");

                    XElement spacing = EnsureFirstChild(paragraph, w + "spacing");
                    spacing.SetAttributeValue(w + "before", "0");
                    spacing.SetAttributeValue(w + "after", "0");
                    spacing.SetAttributeValue(w + "line", "240");
                    spacing.SetAttributeValue(w + "lineRule", "auto");

                    if (normalizeBodyText)
                    {
                        NormalizeDocxParagraphText(paragraph, w);
                    }

                    ApplyQuotationFontOverrides(paragraph, w, rule.Table.ChineseFontName);
                }

                if (!rule.Table.UseHeaderShading) continue;

                XElement? headerRow = table.Elements(trName).FirstOrDefault();
                if (headerRow == null) continue;

                foreach (XElement cell in headerRow.Elements(tcName))
                {
                    XElement? tcPr = cell.Element(tcPrName);
                    if (tcPr == null)
                    {
                        tcPr = new XElement(tcPrName);
                        cell.AddFirst(tcPr);
                    }

                    foreach (XElement innerShading in cell.Descendants(shdName)
                                 .Where(shading => shading.Parent?.Name != tcPrName)
                                 .ToList())
                    {
                        innerShading.Remove();
                    }

                    XElement? cellShading = tcPr.Element(shdName);
                    if (cellShading == null)
                    {
                        cellShading = new XElement(shdName);
                        tcPr.Add(cellShading);
                    }

                    cellShading.SetAttributeValue(valAttr, "clear");
                    cellShading.SetAttributeValue(colorAttr, "auto");
                    cellShading.SetAttributeValue(fillAttr, fillHex);
                }
            }

            RewriteDocumentXmlEntry(archive, documentEntry, documentXml);
        }

        private static Dictionary<string, string> BuildDocxStyleNameMap(ZipArchive archive, XNamespace w)
        {
            var styleNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            ZipArchiveEntry? stylesEntry = archive.GetEntry("word/styles.xml");
            if (stylesEntry == null) return styleNames;

            XName styleName = w + "style";
            XName nameName = w + "name";
            XName styleIdAttr = w + "styleId";
            XName valAttr = w + "val";

            using Stream stream = stylesEntry.Open();
            XDocument stylesXml = XDocument.Load(stream);

            foreach (XElement style in stylesXml.Descendants(styleName))
            {
                string? styleId = style.Attribute(styleIdAttr)?.Value;
                string? name = style.Element(nameName)?.Attribute(valAttr)?.Value;
                if (!string.IsNullOrWhiteSpace(styleId) && !string.IsNullOrWhiteSpace(name))
                    styleNames[styleId] = name;
            }

            return styleNames;
        }

        private void ApplyDocxParagraphFormatting(
            XDocument documentXml,
            XNamespace w,
            Dictionary<string, string> styleNames,
            FormattingRule rule,
            bool normalizeBodyText)
        {
            XName tblName = w + "tbl";
            XName pName = w + "p";
            XName pPrName = w + "pPr";
            XName pStyleName = w + "pStyle";
            XName rName = w + "r";
            XName tName = w + "t";
            XName valAttr = w + "val";

            foreach (XElement paragraph in documentXml.Descendants(pName))
            {
                if (paragraph.Ancestors(tblName).Any()) continue;

                XElement pPr = EnsureFirstChild(paragraph, pPrName);
                string? styleId = pPr.Element(pStyleName)?.Attribute(valAttr)?.Value;
                string styleName = GetDocxStyleName(styleId, styleNames);

                int headingLevel = GetHeadingLevelFromStyleName(styleName);
                if (headingLevel > 0)
                {
                    var heading = ResolveHeadingFormat(rule, headingLevel);
                    if (heading != null)
                    {
                        ApplyParagraphProperties(pPr, w, heading.LineSpacingType, heading.LineSpacingValue,
                            heading.SpaceBeforePoints, heading.SpaceAfterPoints, 0, spacingInPoints: true);
                        ApplyRunsFormatting(paragraph.Elements(rName), w, heading.ChineseFontName, heading.EnglishFontName,
                            heading.FontSizePoint, heading.IsBold,
                            heading.UseCustomFontColor ? NormalizeHexColor(heading.FontColorHex) : "000000");
                        ApplyQuotationFontOverrides(paragraph, w, heading.ChineseFontName);
                    }

                    continue;
                }

                if (!IsBodyStyle(styleId, styleName)) continue;

                ApplyParagraphProperties(pPr, w, rule.Paragraph.LineSpacingType, rule.Paragraph.LineSpacingValue,
                    rule.Paragraph.SpaceBeforeLines, rule.Paragraph.SpaceAfterLines, rule.Paragraph.FirstLineIndentChars);
                ApplyRunsFormatting(paragraph.Elements(rName), w, rule.BodyText.ChineseFontName, rule.BodyText.EnglishFontName,
                    rule.BodyText.FontSizePoint, rule.BodyText.IsBold,
                    rule.BodyText.UseCustomFontColor ? NormalizeHexColor(rule.BodyText.FontColorHex) : "000000");

                if (normalizeBodyText)
                {
                    NormalizeDocxParagraphText(paragraph, w);
                }

                ApplyQuotationFontOverrides(paragraph, w, rule.BodyText.ChineseFontName);
            }
        }

        private static void NormalizeDocxParagraphText(XElement paragraph, XNamespace w)
        {
            XName rName = w + "r";
            XName tName = w + "t";

            var runs = paragraph.Elements(rName).ToList();
            if (runs.Count == 0) return;

            string original = string.Concat(runs.SelectMany(run => run.Descendants(tName).Select(text => text.Value)));
            if (string.IsNullOrEmpty(original)) return;

            string normalized = TextNormalizationService.Normalize(original);
            if (normalized == original) return;

            XElement firstRun = runs[0];
            firstRun.Elements(tName).Remove();
            XElement firstText = new XElement(tName, normalized);
            if (normalized.StartsWith(" ") || normalized.EndsWith(" "))
                firstText.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            firstRun.Add(firstText);

            foreach (XElement run in runs.Skip(1).ToList())
                run.Remove();
        }

        private static string GetDocxStyleName(string? styleId, Dictionary<string, string> styleNames)
        {
            if (string.IsNullOrWhiteSpace(styleId)) return "Normal";
            return styleNames.TryGetValue(styleId, out string? styleName) ? styleName : styleId;
        }

        private static bool IsBodyStyle(string? styleId, string styleName)
            => string.IsNullOrWhiteSpace(styleId)
               || styleName.Equals("正文", StringComparison.OrdinalIgnoreCase)
               || styleName.Equals("Normal", StringComparison.OrdinalIgnoreCase)
               || styleName.Contains("Body", StringComparison.OrdinalIgnoreCase);

        private static XElement EnsureFirstChild(XElement parent, XName childName)
        {
            XElement? child = parent.Element(childName);
            if (child != null) return child;

            child = new XElement(childName);
            parent.AddFirst(child);
            return child;
        }

        private static XElement EnsureChild(XElement parent, XName childName)
        {
            XElement? child = parent.Element(childName);
            if (child != null) return child;

            child = new XElement(childName);
            parent.Add(child);
            return child;
        }

        private static void ApplyParagraphProperties(
            XElement pPr,
            XNamespace w,
            LineSpacingType lineSpacingType,
            float lineSpacingValue,
            float spaceBeforeLines,
            float spaceAfterLines,
            float firstLineIndentChars,
            bool spacingInPoints = false)
        {
            XName spacingName = w + "spacing";
            XName indName = w + "ind";
            XName beforeAttr = w + "before";
            XName afterAttr = w + "after";
            XName lineAttr = w + "line";
            XName lineRuleAttr = w + "lineRule";
            XName firstLineCharsAttr = w + "firstLineChars";
            XName firstLineAttr = w + "firstLine";
            XName hangingAttr = w + "hanging";
            XName hangingCharsAttr = w + "hangingChars";

            XElement spacing = EnsureChild(pPr, spacingName);
            float spacingUnit = spacingInPoints ? 20f : 240f;
            spacing.SetAttributeValue(beforeAttr, Math.Max(0, (int)Math.Round(spaceBeforeLines * spacingUnit)));
            spacing.SetAttributeValue(afterAttr, Math.Max(0, (int)Math.Round(spaceAfterLines * spacingUnit)));

            switch (lineSpacingType)
            {
                case LineSpacingType.Single:
                    spacing.SetAttributeValue(lineAttr, "240");
                    spacing.SetAttributeValue(lineRuleAttr, "auto");
                    break;
                case LineSpacingType.OneAndHalf:
                    spacing.SetAttributeValue(lineAttr, "360");
                    spacing.SetAttributeValue(lineRuleAttr, "auto");
                    break;
                case LineSpacingType.Double:
                    spacing.SetAttributeValue(lineAttr, "480");
                    spacing.SetAttributeValue(lineRuleAttr, "auto");
                    break;
                case LineSpacingType.Fixed:
                    spacing.SetAttributeValue(lineAttr, Math.Max(0, (int)Math.Round(lineSpacingValue * 20)));
                    spacing.SetAttributeValue(lineRuleAttr, "exact");
                    break;
                case LineSpacingType.AtLeast:
                    spacing.SetAttributeValue(lineAttr, Math.Max(0, (int)Math.Round(lineSpacingValue * 20)));
                    spacing.SetAttributeValue(lineRuleAttr, "atLeast");
                    break;
                default:
                    spacing.SetAttributeValue(lineAttr, Math.Max(0, (int)Math.Round(lineSpacingValue * 240)));
                    spacing.SetAttributeValue(lineRuleAttr, "auto");
                    break;
            }

            XElement indent = EnsureChild(pPr, indName);
            if (firstLineIndentChars <= 0)
            {
                indent.SetAttributeValue(firstLineAttr, "0");
                indent.SetAttributeValue(firstLineCharsAttr, "0");
                indent.SetAttributeValue(hangingAttr, "0");
                indent.SetAttributeValue(hangingCharsAttr, "0");
            }
            else
            {
                indent.SetAttributeValue(firstLineAttr, null);
                indent.SetAttributeValue(hangingAttr, null);
                indent.SetAttributeValue(hangingCharsAttr, null);
                indent.SetAttributeValue(firstLineCharsAttr, Math.Max(0, (int)Math.Round(firstLineIndentChars * 100)));
            }
        }

        private static void ApplyRunsFormatting(
            IEnumerable<XElement> runs,
            XNamespace w,
            string chineseFontName,
            string englishFontName,
            float fontSizePoint,
            bool bold,
            string colorHex)
        {
            XName rPrName = w + "rPr";
            XName rFontsName = w + "rFonts";
            XName szName = w + "sz";
            XName szCsName = w + "szCs";
            XName bName = w + "b";
            XName bCsName = w + "bCs";
            XName colorName = w + "color";
            XName valAttr = w + "val";

            XName asciiAttr = w + "ascii";
            XName hAnsiAttr = w + "hAnsi";
            XName eastAsiaAttr = w + "eastAsia";
            XName csAttr = w + "cs";

            string halfPoints = Math.Max(1, (int)Math.Round(fontSizePoint * 2)).ToString();

            foreach (XElement run in runs)
            {
                XElement rPr = EnsureFirstChild(run, rPrName);
                bool hasChineseQuotationOrPunctuation = ContainsChineseQuotationOrPunctuation(run.Value);

                XElement rFonts = EnsureChild(rPr, rFontsName);
                rFonts.SetAttributeValue(asciiAttr, englishFontName);
                rFonts.SetAttributeValue(hAnsiAttr, englishFontName);
                rFonts.SetAttributeValue(eastAsiaAttr, chineseFontName);
                rFonts.SetAttributeValue(csAttr, hasChineseQuotationOrPunctuation ? chineseFontName : englishFontName);

                EnsureChild(rPr, szName).SetAttributeValue(valAttr, halfPoints);
                EnsureChild(rPr, szCsName).SetAttributeValue(valAttr, halfPoints);
                EnsureChild(rPr, colorName).SetAttributeValue(valAttr, colorHex);

                if (bold)
                {
                    EnsureChild(rPr, bName).SetAttributeValue(valAttr, "1");
                    EnsureChild(rPr, bCsName).SetAttributeValue(valAttr, "1");
                }
                else
                {
                    rPr.Element(bName)?.Remove();
                    rPr.Element(bCsName)?.Remove();
                }
            }
        }

        private static void ApplyQuotationFontOverrides(XElement paragraph, XNamespace w, string chineseFontName)
        {
            XName rName = w + "r";
            XName tName = w + "t";
            XName rPrName = w + "rPr";
            XName rFontsName = w + "rFonts";
            XName asciiAttr = w + "ascii";
            XName hAnsiAttr = w + "hAnsi";
            XName eastAsiaAttr = w + "eastAsia";
            XName csAttr = w + "cs";

            foreach (XElement run in paragraph.Elements(rName).ToList())
            {
                if (run.Elements().Any(e => e.Name != rPrName && e.Name != tName))
                    continue;

                XElement? textElement = run.Element(tName);
                if (textElement == null)
                    continue;

                string text = textElement.Value;
                if (string.IsNullOrEmpty(text) || !text.Any(IsChineseQuotationOrPunctuationChar))
                    continue;

                XElement? originalRPr = run.Element(rPrName);
                var replacementRuns = new List<XElement>(text.Length);

                foreach (char c in text)
                {
                    XElement replacementRun = new XElement(rName);
                    if (originalRPr != null)
                        replacementRun.Add(new XElement(originalRPr));

                    XElement replacementText = new XElement(tName, c.ToString());
                    if (char.IsWhiteSpace(c))
                        replacementText.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                    replacementRun.Add(replacementText);

                    if (IsChineseQuotationOrPunctuationChar(c))
                    {
                        XElement rFonts = EnsureChild(EnsureFirstChild(replacementRun, rPrName), rFontsName);
                        rFonts.SetAttributeValue(asciiAttr, chineseFontName);
                        rFonts.SetAttributeValue(hAnsiAttr, chineseFontName);
                        rFonts.SetAttributeValue(eastAsiaAttr, chineseFontName);
                        rFonts.SetAttributeValue(csAttr, chineseFontName);
                    }

                    replacementRuns.Add(replacementRun);
                }

                run.AddBeforeSelf(replacementRuns);
                run.Remove();
            }
        }

        private static void RewriteDocumentXmlEntry(ZipArchive archive, ZipArchiveEntry documentEntry, XDocument documentXml)
        {
            documentEntry.Delete();
            ZipArchiveEntry newDocumentEntry = archive.CreateEntry("word/document.xml");
            using Stream outputStream = newDocumentEntry.Open();
            documentXml.Save(outputStream, SaveOptions.DisableFormatting);
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

        private static string NormalizeHexColor(string? hex)
        {
            if (string.IsNullOrWhiteSpace(hex)) return "D9D9D9";

            string cleanHex = hex.Trim().TrimStart('#');
            if (cleanHex.Length != 6 || !Regex.IsMatch(cleanHex, "^[0-9a-fA-F]{6}$"))
                return "D9D9D9";

            return cleanHex.ToUpperInvariant();
        }

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
