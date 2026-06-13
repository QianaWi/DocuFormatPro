using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace DocuFormatPro.Services
{
    /// <summary>
    /// 表格和图片题注自动处理服务
    /// 使用 Word 原生 SEQ 域代码实现自动编号，支持新增/删除后 F9 刷新
    /// </summary>
    public class CaptionService
    {
        /// <summary>处理文档中所有表格和图片的题注</summary>
        public void ProcessCaptions(Document doc, IProgress<string>? progress = null)
        {
            progress?.Report("正在分析文档结构...");

            // 收集所有表格和图片元素
            var elements = CollectElements(doc);

            // 按文档位置排序
            elements.Sort((a, b) => a.Position.CompareTo(b.Position));

            // 应用题注（从后往前处理，避免位置偏移）
            elements.Reverse();
            foreach (var elem in elements)
            {
                try
                {
                    if (elem.Type == ElementType.Table)
                    {
                        ApplyTableCaption(doc, elem);
                    }
                    else if (elem.Type == ElementType.Image)
                    {
                        ApplyImageCaption(doc, elem);
                    }
                }
                catch { continue; }
            }

            // 更新所有域代码，确保编号正确显示
            try
            {
                doc.Fields.Update();
            }
            catch { }

            progress?.Report("题注处理完成");
        }

        #region 数据结构

        private enum ElementType { Table, Image }

        private class DocElement
        {
            public ElementType Type;
            public int Position;
            public object? WordObject;
            public Paragraph? ExistingCaptionParagraph;
        }

        #endregion

        #region 收集文档元素

        private List<DocElement> CollectElements(Document doc)
        {
            var elements = new List<DocElement>();

            // 收集表格
            foreach (Table table in doc.Tables)
            {
                try
                {
                    var elem = new DocElement
                    {
                        Type = ElementType.Table,
                        Position = table.Range.Start,
                        WordObject = table
                    };
                    CheckTableExistingCaption(doc, table, elem);
                    elements.Add(elem);
                }
                catch { continue; }
            }

            // 收集图片
            foreach (InlineShape shape in doc.InlineShapes)
            {
                try
                {
                    if (shape.Type == WdInlineShapeType.wdInlineShapePicture ||
                        shape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture ||
                        shape.Type == WdInlineShapeType.wdInlineShapeEmbeddedOLEObject)
                    {
                        var elem = new DocElement
                        {
                            Type = ElementType.Image,
                            Position = shape.Range.Start,
                            WordObject = shape
                        };
                        CheckImageExistingCaption(doc, shape, elem);
                        elements.Add(elem);
                    }
                }
                catch { continue; }
            }

            return elements;
        }

        private void CheckTableExistingCaption(Document doc, Table table, DocElement elem)
        {
            try
            {
                Microsoft.Office.Interop.Word.Range beforeRange = doc.Range(Math.Max(0, table.Range.Start - 1), table.Range.Start);
                Paragraph? prevPara = null;
                try { prevPara = beforeRange.Paragraphs[1]; } catch { }

                if (prevPara != null)
                {
                    string text = prevPara.Range.Text.Trim();
                    string styleNameLocal = "";
                    try { styleNameLocal = ((Style)prevPara.get_Style()).NameLocal; } catch { }

                    bool isCaption = styleNameLocal == "题注" || styleNameLocal == "Caption";
                    bool hasTablePrefix = Regex.IsMatch(text, @"^表[\s\d]") || HasCorrectCaptionFormat(prevPara, "表");

                    if (isCaption || hasTablePrefix)
                    {
                        elem.ExistingCaptionParagraph = prevPara;
                    }
                }
            }
            catch { }
        }

        private void CheckImageExistingCaption(Document doc, InlineShape shape, DocElement elem)
        {
            try
            {
                var r = shape.Range;
                r.MoveEnd(WdUnits.wdParagraph, 1);
                if (r.Paragraphs.Count >= 1)
                {
                    var lastPara = r.Paragraphs[r.Paragraphs.Count];
                    string text = lastPara.Range.Text.Trim();
                    string styleNameLocal = "";
                    try { styleNameLocal = ((Style)lastPara.get_Style()).NameLocal; } catch { }

                    bool isCaption = styleNameLocal == "题注" || styleNameLocal == "Caption";
                    bool hasFigPrefix = Regex.IsMatch(text, @"^图[\s\d]") || HasCorrectCaptionFormat(lastPara, "图");

                    if (isCaption || hasFigPrefix)
                    {
                        elem.ExistingCaptionParagraph = lastPara;
                    }
                }
            }
            catch { }
        }

        #endregion

        #region 应用题注

        /// <summary>应用表格题注（表格上方，使用 SEQ 域）</summary>
        private void ApplyTableCaption(Document doc, DocElement elem)
        {
            Paragraph? targetPara = elem.ExistingCaptionParagraph;

            if (targetPara != null && HasCorrectCaptionFormat(targetPara, "表"))
            {
                // 域代码正确，只验证字体格式和描述文字
                EnsureCaptionDescriptionAndFormat(targetPara, "表");
                return;
            }

            // 从旧题注中提取描述文字
            string existingDescription = "";
            if (targetPara != null)
                existingDescription = ExtractCaptionDescription(targetPara, "表");

            if (targetPara == null)
            {
                var table = (Table)elem.WordObject!;
                var app = doc.Application;

                try
                {
                    table.Range.Cells[1].Select();
                    object direction = WdCollapseDirection.wdCollapseStart;
                    app.Selection.Collapse(ref direction);
                    app.Selection.SplitTable();

                    object count = 1;
                    targetPara = table.Range.Paragraphs[1].Previous(ref count);
                }
                catch { return; }
            }

            if (targetPara == null) return;

            try
            {
                object styleNormal = WdBuiltinStyle.wdStyleNormal;
                targetPara.set_Style(ref styleNormal);
                targetPara.Format.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;
            }
            catch { }

            ClearParagraphContent(targetPara);
            InsertCaptionFields(targetPara, "表", existingDescription);
            FormatCaptionParagraph(targetPara);
        }

        /// <summary>应用图片题注（图片下方，使用 SEQ 域）</summary>
        private void ApplyImageCaption(Document doc, DocElement elem)
        {
            Paragraph? targetPara = elem.ExistingCaptionParagraph;

            if (targetPara != null && HasCorrectCaptionFormat(targetPara, "图"))
            {
                // 域代码正确，只验证字体格式和描述文字
                EnsureCaptionDescriptionAndFormat(targetPara, "图");
                return;
            }

            string existingDescription = "";
            if (targetPara != null)
                existingDescription = ExtractCaptionDescription(targetPara, "图");

            if (targetPara == null)
            {
                var shape = (InlineShape)elem.WordObject!;
                shape.Range.Paragraphs[1].Range.InsertParagraphAfter();

                try
                {
                    object count = 1;
                    targetPara = shape.Range.Paragraphs[1].Next(ref count);
                }
                catch { return; }
            }

            if (targetPara == null) return;

            ClearParagraphContent(targetPara);
            InsertCaptionFields(targetPara, "图", existingDescription);
            FormatCaptionParagraph(targetPara);
        }

        /// <summary>对域代码正确的题注：强制应用字体格式 + 补充缺失描述</summary>
        private void EnsureCaptionDescriptionAndFormat(Paragraph para, string prefix)
        {
            try
            {
                // 获取显示文本，去掉域标记字符和段落标记
                string text = para.Range.Text ?? "";
                text = Regex.Replace(text, @"[\x13\x14\x15\r\n\a]", "").Trim();

                // 去掉标签+编号部分（如 "表0-4"），剩余的是描述文字
                string desc = Regex.Replace(text, $@"^{prefix}[\d.\-]+\s*", "").Trim();

                if (string.IsNullOrWhiteSpace(desc))
                {
                    // 没有描述文字，用 Range 在段落末尾插入
                    // 获取段落 range，排除最后一个字符（段落标记）
                    var paraRange = para.Range;
                    int insertPos = paraRange.End - 1; // 段落标记之前
                    var insertRange = para.Range.Document.Range(insertPos, insertPos);
                    insertRange.Text = " XXXXXX";
                }

                // 强制应用字体格式（黑体小四居中）
                FormatCaptionParagraph(para);
            }
            catch { }
        }

        /// <summary>从旧题注段落中提取描述文字（去掉标签和编号前缀后剩余部分）</summary>
        private string ExtractCaptionDescription(Paragraph para, string prefix)
        {
            try
            {
                string text = para.Range.Text.Trim();
                // 去掉段落标记
                text = text.TrimEnd('\r', '\n', '\a');
                // 去掉前缀标签+编号部分（如"表1-2"、"表1"、"图2.3-1"等）
                var match = Regex.Match(text, $@"^{prefix}[\d.\-]+\s*(.*)$");
                if (match.Success)
                    return match.Groups[1].Value.Trim();
                // 如果不匹配标准格式，但以前缀开头，取前缀后面的所有内容
                if (text.StartsWith(prefix) && text.Length > prefix.Length)
                {
                    string after = Regex.Replace(text.Substring(prefix.Length), @"^[\d.\-\s]+", "").Trim();
                    return after;
                }
            }
            catch { }
            return "";
        }

        /// <summary>检查已有题注是否严格匹配我们的格式（STYLEREF + SEQ 域）</summary>
        private bool HasCorrectCaptionFormat(Paragraph para, string seqName)
        {
            try
            {
                bool hasSeq = false;
                bool hasStyleRef = false;

                foreach (Field field in para.Range.Fields)
                {
                    string code = field.Code.Text.Trim();
                    if (field.Type == WdFieldType.wdFieldSequence)
                    {
                        if (code.Contains($"SEQ {seqName}") && code.Contains("\\* ARABIC") && code.Contains("\\s 1"))
                            hasSeq = true;
                    }
                    else if (field.Type == WdFieldType.wdFieldStyleRef)
                    {
                        if (code.Contains("1") && code.Contains("\\s"))
                            hasStyleRef = true;
                    }
                }

                return hasSeq && hasStyleRef;
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 在段落中插入题注域代码（始终使用章节号格式）
        /// 格式：前缀 + STYLEREF域 + - + SEQ域 + 空格 + 描述文字/XXXXXX
        /// </summary>
        private void InsertCaptionFields(Paragraph para, string prefix, string description)
        {
            object missing = Type.Missing;

            Microsoft.Office.Interop.Word.Range r = para.Range;
            object unit = WdUnits.wdCharacter;
            object cnt = -1;
            r.MoveEnd(ref unit, ref cnt); // 排除段落标记

            // 如果没有描述文字，用 XXXXXX 占位
            string descText = string.IsNullOrWhiteSpace(description) ? "XXXXXX" : description;

            // 始终使用章节号格式：前缀 + §STYLEREF§ + - + §SEQ§ + 空格 + 描述
            r.Text = $"{prefix}\x01-\x02 {descText}";

            var doc = para.Range.Document;

            // 替换 \x01 为 STYLEREF 域
            var findRange1 = para.Range;
            findRange1.Find.ClearFormatting();
            object findText1 = "\x01";
            object replaceWith1 = "";
            object findWrap = WdFindWrap.wdFindStop;
            if (findRange1.Find.Execute(ref findText1, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref findWrap,
                ref missing, ref replaceWith1, ref missing, ref missing,
                ref missing, ref missing, ref missing))
            {
                findRange1.Text = "";
                doc_AddField(doc, findRange1, WdFieldType.wdFieldStyleRef, "1 \\s", false);
            }

            // 替换 \x02 为 SEQ 域
            var findRange2 = para.Range;
            findRange2.Find.ClearFormatting();
            object findText2 = "\x02";
            object replaceWith2 = "";
            object findWrap2 = WdFindWrap.wdFindStop;
            if (findRange2.Find.Execute(ref findText2, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref findWrap2,
                ref missing, ref replaceWith2, ref missing, ref missing,
                ref missing, ref missing, ref missing))
            {
                findRange2.Text = "";
                doc_AddField(doc, findRange2, WdFieldType.wdFieldSequence, $"{prefix} \\* ARABIC \\s 1", false);
            }
        }

        /// <summary>检测文档中是否有带数字编号前缀的 Heading 1 段落（保留供 CheckExisting 用）</summary>
        private bool HasNumberedHeading1(Document doc)
        {
            foreach (Paragraph para in doc.Paragraphs)
            {
                try
                {
                    var style = (Style)para.get_Style();
                    string styleName = style.NameLocal;
                    if (styleName.Contains("1") && (styleName.Contains("标题") || styleName.Contains("Heading")))
                    {
                        string text = para.Range.Text.Trim();
                        if (Regex.IsMatch(text, @"^\d"))
                            return true;
                    }
                }
                catch { continue; }
            }
            return false;
        }

        /// <summary>通过 COM 在指定 Range 位置插入域</summary>
        private void doc_AddField(Document doc, Microsoft.Office.Interop.Word.Range range, WdFieldType fieldType, string fieldCode, bool preserveFormatting)
        {
            object missing = Type.Missing;
            object preserveFmt = preserveFormatting;
            range.Fields.Add(range, fieldType, fieldCode, ref preserveFmt);
        }

        /// <summary>清空段落内容（文字和域），保留段落标记</summary>
        private void ClearParagraphContent(Paragraph para)
        {
            try
            {
                // 删除段落中所有域
                while (para.Range.Fields.Count > 0)
                {
                    try { para.Range.Fields[1].Delete(); } catch { break; }
                }

                // 清空文字内容但保留段落标记（不用 Delete，避免段落消失被吸入表格）
                var r = para.Range;
                object unit = WdUnits.wdCharacter;
                object count = -1;
                r.MoveEnd(ref unit, ref count); // 排除段落标记
                if (r.Start < r.End)
                    r.Text = "";
            }
            catch { }
        }

        /// <summary>格式化题注段落：题注样式 + 宋体小四居中</summary>
        private void FormatCaptionParagraph(Paragraph para)
        {
            try
            {
                try
                {
                    object captionStyle = "题注";
                    para.set_Style(ref captionStyle);
                }
                catch
                {
                    try
                    {
                        object captionStyle = "Caption";
                        para.set_Style(ref captionStyle);
                    }
                    catch
                    {
                        object styleNormal = WdBuiltinStyle.wdStyleNormal;
                        para.set_Style(ref styleNormal);
                    }
                }

                para.Format.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;

                // 覆盖具体字体格式：黑体小四居中
                para.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.Font.Name = "黑体";
                para.Range.Font.NameAscii = "Times New Roman";
                para.Range.Font.Size = 12f;
                para.Range.Font.Bold = 0;
                para.Range.Font.Color = WdColor.wdColorBlack;
                para.Format.FirstLineIndent = 0;
                para.Format.CharacterUnitFirstLineIndent = 0;
                para.Format.LeftIndent = 0;
                para.Format.CharacterUnitLeftIndent = 0;
            }
            catch { }
        }

        #endregion
    }
}
