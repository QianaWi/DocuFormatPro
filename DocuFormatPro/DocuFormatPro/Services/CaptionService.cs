using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace DocuFormatPro.Services
{
    /// <summary>
    /// 表格和图片题注自动处理服务
    /// </summary>
    public class CaptionService
    {
        /// <summary>处理文档中所有表格和图片的题注</summary>
        public void ProcessCaptions(Document doc, IProgress<string>? progress = null)
        {
            progress?.Report("正在分析文档结构...");

            // 1. 构建标题映射：位置 -> 标题编号
            var headings = BuildHeadingMap(doc);

            // 2. 收集所有表格和图片元素
            var elements = CollectElements(doc);

            // 3. 按文档位置排序
            elements.Sort((a, b) => a.Position.CompareTo(b.Position));

            // 4. 分配题注编号
            AssignCaptionNumbers(elements, headings);

            // 5. 应用题注（从后往前处理，避免位置偏移）
            elements.Reverse();
            foreach (var elem in elements)
            {
                try
                {
                    if (elem.Type == ElementType.Table || elem.Type == ElementType.TableAsImage)
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

            progress?.Report("题注处理完成");
        }

        #region 数据结构

        private enum ElementType { Table, Image, TableAsImage }

        private class DocElement
        {
            public ElementType Type;
            public int Position;          // Range.Start
            public object? WordObject;    // Table or InlineShape
            public string? ExistingCaption;
            public Paragraph? CaptionParagraph;
            public string SectionNumber = "";
            public int SequenceInSection;
            public string FinalCaption = "";
        }

        private class HeadingInfo
        {
            public int Position;
            public int Level;
            public string Number = "";
        }

        #endregion

        #region 构建标题映射

        private List<HeadingInfo> BuildHeadingMap(Document doc)
        {
            var headings = new List<HeadingInfo>();
            foreach (Paragraph para in doc.Paragraphs)
            {
                try
                {
                    var style = (Style)para.get_Style();
                    string styleName = style.NameLocal;
                    int level = 0;

                    if (styleName.Contains("标题") || styleName.Contains("Heading"))
                    {
                        // 尝试从样式名提取级别
                        var m = Regex.Match(styleName, @"(\d+)");
                        if (m.Success) level = int.Parse(m.Value);
                    }

                    if (level > 0 && level <= 9)
                    {
                        string number = ExtractHeadingNumber(para);
                        if (!string.IsNullOrEmpty(number))
                        {
                            headings.Add(new HeadingInfo
                            {
                                Position = para.Range.Start,
                                Level = level,
                                Number = number
                            });
                        }
                    }
                }
                catch { continue; }
            }
            return headings;
        }

        /// <summary>从标题段落中提取编号，如 "2.3.2"</summary>
        private string ExtractHeadingNumber(Paragraph para)
        {
            string text = para.Range.Text.Trim();

            // 优先从文本提取：匹配开头的数字编号 (1, 1.2, 1.2.3 等)
            var match = Regex.Match(text, @"^(\d+(\.\d+)*)");
            if (match.Success) return match.Groups[1].Value;

            // 尝试从 ListFormat 提取
            try
            {
                string listStr = para.Range.ListFormat.ListString;
                if (!string.IsNullOrEmpty(listStr))
                {
                    var m2 = Regex.Match(listStr.Trim(), @"(\d+(\.\d+)*)");
                    if (m2.Success) return m2.Groups[1].Value;
                }
            }
            catch { }

            return "";
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

                    // 检查表格上方是否有题注
                    CheckTableExistingCaption(doc, table, elem);
                    elements.Add(elem);
                }
                catch { continue; }
            }

            // 收集图片 (InlineShapes)
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

        /// <summary>检查表格上方是否已有题注</summary>
        private void CheckTableExistingCaption(Document doc, Table table, DocElement elem)
        {
            try
            {
                // 获取表格上方的段落
                Microsoft.Office.Interop.Word.Range beforeRange = doc.Range(Math.Max(0, table.Range.Start - 1), table.Range.Start);
                Paragraph? prevPara = null;
                try { prevPara = beforeRange.Paragraphs[1]; } catch { }

                if (prevPara != null)
                {
                    string text = prevPara.Range.Text.Trim();
                    if (Regex.IsMatch(text, @"^表\s*\d"))
                    {
                        elem.ExistingCaption = text;
                        elem.CaptionParagraph = prevPara;
                    }
                    else if (Regex.IsMatch(text, @"^图\s*\d"))
                    {
                        // 特殊情况：表格被当作图片使用
                        elem.Type = ElementType.TableAsImage;
                        elem.ExistingCaption = text;
                        elem.CaptionParagraph = prevPara;
                    }
                }
            }
            catch { }
        }

        /// <summary>检查图片下方是否已有题注</summary>
        private void CheckImageExistingCaption(Document doc, InlineShape shape, DocElement elem)
        {
            try
            {
                Microsoft.Office.Interop.Word.Range afterRange = doc.Range(shape.Range.End, Math.Min(shape.Range.End + 2, doc.Content.End));
                try
                {
                    // 跳到下一段
                    var r = shape.Range;
                    r.MoveEnd(WdUnits.wdParagraph, 1);
                    if (r.Paragraphs.Count >= 1)
                    {
                        var lastPara = r.Paragraphs[r.Paragraphs.Count];
                        string text = lastPara.Range.Text.Trim();
                        if (Regex.IsMatch(text, @"^图\s*\d"))
                        {
                            elem.ExistingCaption = text;
                            elem.CaptionParagraph = lastPara;
                        }
                    }
                }
                catch { }
            }
            catch { }
        }

        #endregion

        #region 分配编号

        private void AssignCaptionNumbers(List<DocElement> elements, List<HeadingInfo> headings)
        {
            // 合并标题和元素，按位置排序
            string currentSection = "";
            var tableCounts = new Dictionary<string, int>();
            var imageCounts = new Dictionary<string, int>();

            int headingIdx = 0;

            foreach (var elem in elements)
            {
                // 更新当前章节
                while (headingIdx < headings.Count && headings[headingIdx].Position <= elem.Position)
                {
                    currentSection = headings[headingIdx].Number;
                    headingIdx++;
                }

                string section = string.IsNullOrEmpty(currentSection) ? "0" : currentSection;
                elem.SectionNumber = section;

                if (elem.Type == ElementType.Table)
                {
                    if (!tableCounts.ContainsKey(section)) tableCounts[section] = 0;
                    tableCounts[section]++;
                    elem.SequenceInSection = tableCounts[section];
                    elem.FinalCaption = $"表{section}-{elem.SequenceInSection}";
                }
                else // Image 或 TableAsImage 都按图片计数
                {
                    if (!imageCounts.ContainsKey(section)) imageCounts[section] = 0;
                    imageCounts[section]++;
                    elem.SequenceInSection = imageCounts[section];
                    elem.FinalCaption = $"图{section}-{elem.SequenceInSection}";
                }
            }
        }

        #endregion

        #region 应用题注

        /// <summary>应用表格题注（表格上方居中）</summary>
        private void ApplyTableCaption(Document doc, DocElement elem)
        {
            string prefix = elem.Type == ElementType.TableAsImage ? "图" : "表";
            // TableAsImage 的 FinalCaption 已经是 "图X-X"

            if (elem.CaptionParagraph != null)
            {
                // 修改已有题注
                SetParagraphTextSafely(elem.CaptionParagraph, elem.FinalCaption);
                FormatCaptionParagraph(elem.CaptionParagraph);
            }
            else
            {
                // 在表格上方添加新题注
                var table = (Table)elem.WordObject!;
                var app = doc.Application;
                
                try
                {
                    // 终极必杀技：调用 Word 原生的 SplitTable 功能模拟人工操作。
                    // 1. 选中表格的第一个单元格（使用 Range.Cells[1] 无视任何合并单元格异常）
                    table.Range.Cells[1].Select();
                    object direction = WdCollapseDirection.wdCollapseStart;
                    app.Selection.Collapse(ref direction);

                    // 2. 当光标在表格第一行时，SplitTable（拆分表格）会强行在表格外部正上方挤出一个纯文本空段落！
                    // 这个由 Word UI 层引擎触发的动作，彻底杜绝了 COM API 中 InsertParagraphBefore 导致的边界吸附和删标题 Bug。
                    app.Selection.SplitTable();
                    
                    // 3. 获取刚刚被挤出来的这个独立的空段落
                    object count = 1;
                    Paragraph pTarget = table.Range.Paragraphs[1].Previous(ref count);
                    
                    if (pTarget != null)
                    {
                        // 强制重置样式，彻底洗掉可能从上方标题继承来的大纲级别，防止进入目录
                        object styleNormal = WdBuiltinStyle.wdStyleNormal;
                        pTarget.set_Style(ref styleNormal);
                        pTarget.Format.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;

                        SetParagraphTextSafely(pTarget, elem.FinalCaption);
                        FormatCaptionParagraph(pTarget);
                    }
                }
                catch { }
            }
        }

        /// <summary>应用图片题注（图片下方居中）</summary>
        private void ApplyImageCaption(Document doc, DocElement elem)
        {
            if (elem.CaptionParagraph != null)
            {
                // 修改已有题注
                SetParagraphTextSafely(elem.CaptionParagraph, elem.FinalCaption);
                FormatCaptionParagraph(elem.CaptionParagraph);
            }
            else
            {
                // 在图片下方添加新题注
                var shape = (InlineShape)elem.WordObject!;
                shape.Range.Paragraphs[1].Range.InsertParagraphAfter();
                
                try
                {
                    object count = 1;
                    Paragraph pTarget = shape.Range.Paragraphs[1].Next(ref count);
                    if (pTarget != null)
                    {
                        SetParagraphTextSafely(pTarget, elem.FinalCaption);
                        FormatCaptionParagraph(pTarget);
                    }
                }
                catch { }
            }
        }

        /// <summary>格式化题注段落：居中、宋体小四</summary>
        private void FormatCaptionParagraph(Paragraph para)
        {
            try
            {
                // 强制重置样式和级别，清除可能从小标题继承下来的大纲级别，防止进入目录或覆盖标题
                object styleNormal = WdBuiltinStyle.wdStyleNormal;
                para.set_Style(ref styleNormal);
                para.Format.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;

                para.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.Font.Name = "宋体";
                para.Range.Font.NameAscii = "Times New Roman";
                para.Range.Font.Size = 12f; // 小四
                para.Range.Font.Bold = 0;
                para.Range.Font.Color = WdColor.wdColorBlack;
                para.Format.FirstLineIndent = 0;
                para.Format.CharacterUnitFirstLineIndent = 0;
            }
            catch { }
        }


        /// <summary>安全地设置段落文本，保留段落标记，防止文本吸入紧邻的表格单元格内</summary>
        private void SetParagraphTextSafely(Paragraph para, string text)
        {
            try
            {
                var r = para.Range;
                object unit = WdUnits.wdCharacter;
                object count = -1;
                r.MoveEnd(ref unit, ref count);
                r.Text = text;
            }
            catch { }
        }

        #endregion
    }
}
