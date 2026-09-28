using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using DocuFormatPro.Models;

namespace DocuFormatPro.Views
{
    public partial class ConfirmFormattingDialog : Window
    {
        public ConfirmFormattingDialog(FormattingRule rule, int fileCount)
        {
            InitializeComponent();
            BuildRulesSummary(rule, fileCount);
        }

        private void BuildRulesSummary(FormattingRule rule, int fileCount)
        {
            bool hasFrontMatter = rule.FrontMatter.InsertFrontMatter && !string.IsNullOrWhiteSpace(rule.FrontMatter.TemplateFilePath);

            // ---- 汇总：启用 / 跳过 ----
            var enabledNames = new List<string>();
            var skippedNames = new List<string>();

            if (rule.ApplyPageMargins) enabledNames.Add("页面设置"); else skippedNames.Add("页面设置");
            if (rule.ApplyBodyFormatting) enabledNames.Add("正文格式"); else skippedNames.Add("正文格式");
            if (rule.ApplyHeadingFormatting) enabledNames.Add("标题样式"); else skippedNames.Add("标题样式");
            if (rule.HeadingNumbering.EnableNumbering) enabledNames.Add("标题编号"); else skippedNames.Add("标题编号");
            if (rule.Table.ApplyTableFormatting) enabledNames.Add("表格样式"); else skippedNames.Add("表格样式");
            if (rule.Table.ApplyTableCaptions) enabledNames.Add("题注"); else skippedNames.Add("题注");
            if (hasFrontMatter) enabledNames.Add("前置页"); else skippedNames.Add("前置页");
            if (rule.NormalizeBodyText) enabledNames.Add("规范化文本"); else skippedNames.Add("规范化文本");
            if (rule.ClearTextBackground) enabledNames.Add("清除背景色"); else skippedNames.Add("清除背景色");

            // ---- 顶部：文件数 + 摘要横幅 ----
            AddItem(RulesPanel, $"📁  待处理文件：{fileCount} 个", "#1A1A2E", bold: true);
            AddSummaryBanner(enabledNames);

            // ---- 中部：仅展示启用功能的参数明细 ----
            if (rule.ApplyPageMargins)
            {
                AddSectionTitle(RulesPanel, "📐 页面设置");
                AddItem(RulesPanel, $"页边距：上 {rule.PageMargins.TopMargin}cm  下 {rule.PageMargins.BottomMargin}cm  左 {rule.PageMargins.LeftMargin}cm  右 {rule.PageMargins.RightMargin}cm");
            }

            if (rule.ApplyBodyFormatting)
            {
                AddSectionTitle(RulesPanel, "🔤 正文格式");
                AddItem(RulesPanel, $"字体：{rule.BodyText.ChineseFontName} / {rule.BodyText.EnglishFontName}，{rule.BodyText.FontSizeName}（{rule.BodyText.FontSizePoint}pt）{(rule.BodyText.IsBold ? "，加粗" : "")}");
                if (rule.BodyText.UseCustomFontColor)
                    AddItem(RulesPanel, $"字体颜色：{rule.BodyText.FontColorHex}");
                AddItem(RulesPanel, $"首行缩进：{rule.Paragraph.FirstLineIndentChars} 字符");
                AddItem(RulesPanel, $"对齐方式：{DescribeAlignment(rule.Paragraph.Alignment)}");
                AddItem(RulesPanel, $"行距：{DescribeLineSpacing(rule.Paragraph.LineSpacingType, rule.Paragraph.LineSpacingValue, rule.Paragraph.LineSpacingUnit)}");
                if (rule.Paragraph.SpaceBeforeLines > 0 || rule.Paragraph.SpaceAfterLines > 0)
                    AddItem(RulesPanel, $"段间距：段前 {rule.Paragraph.SpaceBeforeLines} 行，段后 {rule.Paragraph.SpaceAfterLines} 行");
            }

            if (rule.ApplyHeadingFormatting || rule.HeadingNumbering.EnableNumbering)
            {
                AddSectionTitle(RulesPanel, "📌 标题");
                if (rule.ApplyHeadingFormatting)
                {
                    if (rule.UseOriginalHeadingStyle)
                    {
                        AddItem(RulesPanel, "仅保持原文档标题样式，不应用下方自定义参数", "#9CA3AF");
                    }
                    else foreach (var h in rule.Headings ?? Enumerable.Empty<HeadingStyle>())
                    {
                        string colorNote = h.UseCustomFontColor ? $"，颜色 {h.FontColorHex}" : "";
                        AddItem(RulesPanel, $"标题 {h.Level}：{h.ChineseFontName}，{h.FontSizeName}，{DescribeAlignment(h.Alignment)}{(h.IsBold ? "，加粗" : "")}{colorNote}，段前 {h.SpaceBeforePoints}pt，段后 {h.SpaceAfterPoints}pt，行距 {DescribeLineSpacing(h.LineSpacingType, h.LineSpacingValue, h.LineSpacingUnit)}");
                    }
                }
                else
                {
                    AddItem(RulesPanel, "⬜ 标题样式保持原样（仅处理编号）", "#9CA3AF");
                }

                if (rule.HeadingNumbering.EnableNumbering)
                {
                    string strip = rule.HeadingNumbering.StripExistingNumbers ? "，去除旧编号" : "";
                    string promote = rule.HeadingNumbering.PromoteManualNumberedHeadings ? "，提升手动编号段落为标题" : "";
                    AddItem(RulesPanel, $"✅ 标题自动编号（{DescribeNumberingScheme(rule.HeadingNumbering.Scheme)}{strip}{promote}）");
                }
            }

            if (rule.Table.ApplyTableFormatting || rule.Table.ApplyTableCaptions)
            {
                AddSectionTitle(RulesPanel, "📊 表格与题注");
                if (rule.Table.ApplyTableFormatting)
                {
                    AddItem(RulesPanel, $"✅ 表格样式：{rule.Table.ChineseFontName} {rule.Table.FontSizeName}，{(rule.Table.HeaderBold ? "表头加粗" : "表头不加粗")}，{(rule.Table.RepeatHeaderRow ? "跨页重复表头" : "")}");
                    AddItem(RulesPanel, $"   行距：{DescribeLineSpacing(rule.Table.LineSpacingType, rule.Table.LineSpacingValue, rule.Table.LineSpacingUnit)}，边框：黑色单细线{(rule.Table.UseHeaderShading ? $"，首行底色 {rule.Table.HeaderShadingColorHex}" : "")}");
                }
                else
                {
                    AddItem(RulesPanel, "⬜ 表格样式保持原样", "#9CA3AF");
                }

                if (rule.Table.ApplyTableCaptions)
                    AddItem(RulesPanel, "✅ 自动处理题注编号（表格/图片）");
                else
                    AddItem(RulesPanel, "⬜ 题注保持原样", "#9CA3AF");
            }

            if (hasFrontMatter)
            {
                AddSectionTitle(RulesPanel, "📄 附加功能");
                AddItem(RulesPanel, $"✅ 插入前置页：{System.IO.Path.GetFileName(rule.FrontMatter.TemplateFilePath)}");
            }

            if (rule.NormalizeBodyText || rule.ClearTextBackground)
            {
                AddSectionTitle(RulesPanel, "🧹 文本处理");
                if (rule.NormalizeBodyText)
                    AddItem(RulesPanel, "✅ 规范化文本（去除中英文空格、英文标点转中文全角）");
                if (rule.ClearTextBackground)
                    AddItem(RulesPanel, "✅ 清除文字高亮和底纹背景色（黄色警示标记除外）");
            }

            // ---- 底部：跳过项汇总 ----
            AddSkippedFooter(skippedNames);
        }

        /// <summary>顶部摘要横幅：一眼看出本次将处理哪些内容；未启用任何项时黄色警示</summary>
        private void AddSummaryBanner(IReadOnlyList<string> enabledNames)
        {
            bool any = enabledNames.Count > 0;
            string text = any
                ? $"✅ 本次将处理 {enabledNames.Count} 项：{string.Join("、", enabledNames)}"
                : "⚠ 未启用任何排版项，文档将仅复制保存";

            var bg = any ? Color.FromRgb(0xEC, 0xFD, 0xF5) : Color.FromRgb(0xFF, 0xFB, 0xEB);
            var border = any ? Color.FromRgb(0xA7, 0xF3, 0xD0) : Color.FromRgb(0xFD, 0xE6, 0x8A);
            var fg = any ? Color.FromRgb(0x06, 0x5F, 0x46) : Color.FromRgb(0x92, 0x40, 0x0E);

            RulesPanel.Children.Add(new Border
            {
                Background = new SolidColorBrush(bg),
                BorderBrush = new SolidColorBrush(border),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(10, 8, 10, 8),
                Margin = new Thickness(0, 8, 0, 4),
                Child = new TextBlock
                {
                    Text = text,
                    FontSize = 12.5,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush(fg),
                    TextWrapping = TextWrapping.Wrap
                }
            });
        }

        /// <summary>底部灰色汇总：本次跳过（保留原样）的功能</summary>
        private void AddSkippedFooter(IReadOnlyList<string> skippedNames)
        {
            if (skippedNames.Count == 0) return;

            RulesPanel.Children.Add(new Separator
            {
                Background = new SolidColorBrush(Color.FromRgb(0xE5, 0xE7, 0xEB)),
                Margin = new Thickness(0, 12, 0, 6)
            });
            AddItem(RulesPanel, $"⏭ 本次跳过（保留原样）：{string.Join("、", skippedNames)}", "#9CA3AF");
        }

        private static string DescribeNumberingScheme(HeadingNumberingScheme scheme) => scheme switch
        {
            HeadingNumberingScheme.ChapterNumeric => "第X章 / 1.1 / 1.1.1",
            HeadingNumberingScheme.Traditional => "一、/（一）/ 1.（传统公文，8级）",
            HeadingNumberingScheme.NumericWithPeriod => "1. / 1.1. / 1.1.1.",
            _ => "1 / 1.1 / 1.1.1"
        };

        private void AddSectionTitle(Panel parent, string text)
        {
            parent.Children.Add(new TextBlock
            {
                Text = text,
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(0x6C, 0x63, 0xFF)),
                Margin = new Thickness(0, 10, 0, 4)
            });
        }

        private void AddItem(Panel parent, string text, string colorHex = "#374151", bool bold = false)
        {
            var color = colorHex == "#374151"
                ? Color.FromRgb(0x37, 0x41, 0x51)
                : colorHex == "#9CA3AF"
                    ? Color.FromRgb(0x9C, 0xA3, 0xAF)
                    : Color.FromRgb(0x1A, 0x1A, 0x2E);

            parent.Children.Add(new TextBlock
            {
                Text = text,
                FontSize = 12,
                FontWeight = bold ? FontWeights.SemiBold : FontWeights.Normal,
                Foreground = new SolidColorBrush(color),
                Margin = new Thickness(0, 2, 0, 0),
                TextWrapping = TextWrapping.Wrap
            });
        }

        private string DescribeLineSpacing(LineSpacingType type, float value, LineSpacingUnit unit) => type switch
        {
            LineSpacingType.Single => "单倍",
            LineSpacingType.OneAndHalf => "1.5 倍",
            LineSpacingType.Double => "双倍",
            LineSpacingType.Multiple => $"多倍 {value:0.##}{(unit == LineSpacingUnit.Points ? "pt" : "行")}",
            LineSpacingType.Fixed => $"固定值 {value:0.##}{(unit == LineSpacingUnit.Points ? "pt" : "行")}",
            LineSpacingType.AtLeast => $"最小值 {value:0.##}{(unit == LineSpacingUnit.Points ? "pt" : "行")}",
            _ => value.ToString()
        };

        private string DescribeAlignment(Models.TextAlignment alignment) => alignment switch
        {
            Models.TextAlignment.Center => "居中",
            Models.TextAlignment.Right => "右对齐",
            Models.TextAlignment.Justify => "两端对齐",
            _ => "左对齐"
        };

        private void Confirm_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
