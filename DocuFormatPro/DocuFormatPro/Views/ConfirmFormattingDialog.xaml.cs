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
            // 文件数量提示
            AddItem(RulesPanel, $"📁  待处理文件：{fileCount} 个", "#1A1A2E", bold: true);

            // 页面设置
            AddSectionTitle(RulesPanel, "📐 页面设置");
            AddItem(RulesPanel, $"页边距：上 {rule.PageMargins.TopMargin}cm  下 {rule.PageMargins.BottomMargin}cm  左 {rule.PageMargins.LeftMargin}cm  右 {rule.PageMargins.RightMargin}cm");

            // 正文格式
            AddSectionTitle(RulesPanel, "🔤 正文格式");
            AddItem(RulesPanel, $"字体：{rule.BodyText.ChineseFontName} / {rule.BodyText.EnglishFontName}，{rule.BodyText.FontSizeName}（{rule.BodyText.FontSizePoint}pt）{(rule.BodyText.IsBold ? "，加粗" : "")}");
            if (rule.BodyText.UseCustomFontColor)
                AddItem(RulesPanel, $"字体颜色：{rule.BodyText.FontColorHex}");
            AddItem(RulesPanel, $"首行缩进：{rule.Paragraph.FirstLineIndentChars} 字符");
            AddItem(RulesPanel, $"行距：{DescribeLineSpacing(rule.Paragraph.LineSpacingType, rule.Paragraph.LineSpacingValue)}");
            if (rule.Paragraph.SpaceBeforeLines > 0 || rule.Paragraph.SpaceAfterLines > 0)
                AddItem(RulesPanel, $"段间距：段前 {rule.Paragraph.SpaceBeforeLines} 行，段后 {rule.Paragraph.SpaceAfterLines} 行");

            // 标题格式
            if (rule.Headings?.Count > 0)
            {
                AddSectionTitle(RulesPanel, "📌 标题格式");
                foreach (var h in rule.Headings)
                {
                    string colorNote = h.UseCustomFontColor ? $"，颜色 {h.FontColorHex}" : "";
                    AddItem(RulesPanel, $"标题 {h.Level}：{h.ChineseFontName}，{h.FontSizeName}，{DescribeAlignment(h.Alignment)}{(h.IsBold ? "，加粗" : "")}{colorNote}");
                }
            }

            // 表格格式
            AddSectionTitle(RulesPanel, "📊 表格与题注");
            if (rule.Table.ApplyTableFormatting)
            {
                AddItem(RulesPanel, $"✅ 应用表格样式：{rule.Table.ChineseFontName} {rule.Table.FontSizeName}，{(rule.Table.HeaderBold ? "表头加粗" : "表头不加粗")}，{(rule.Table.RepeatHeaderRow ? "跨页重复表头" : "")}");
                AddItem(RulesPanel, $"   行距：{DescribeLineSpacing(rule.Table.LineSpacingType, rule.Table.LineSpacingValue)}，边框：黑色单细线{(rule.Table.UseHeaderShading ? $"，首行底色 {rule.Table.HeaderShadingColorHex}" : "")}");
            }
            else
            {
                AddItem(RulesPanel, "⬜ 跳过表格样式（保留原有格式）", "#9CA3AF");
            }

            if (rule.Table.ApplyTableCaptions)
                AddItem(RulesPanel, "✅ 自动处理题注编号（表格/图片）");
            else
                AddItem(RulesPanel, "⬜ 跳过题注处理（保留原有题注）", "#9CA3AF");

            // 附加功能
            bool hasFrontMatter = rule.FrontMatter.InsertFrontMatter && !string.IsNullOrWhiteSpace(rule.FrontMatter.TemplateFilePath);
            if (hasFrontMatter)
            {
                AddSectionTitle(RulesPanel, "📄 附加功能");
                AddItem(RulesPanel, $"✅ 插入前置页：{System.IO.Path.GetFileName(rule.FrontMatter.TemplateFilePath)}");
            }

            // 始终执行的操作
            AddSectionTitle(RulesPanel, "🧹 始终执行");
            AddItem(RulesPanel, "清除文字高亮和底纹背景色");
        }

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

        private string DescribeLineSpacing(LineSpacingType type, float value) => type switch
        {
            LineSpacingType.Single => "单倍",
            LineSpacingType.OneAndHalf => "1.5 倍",
            LineSpacingType.Double => "双倍",
            LineSpacingType.Multiple => $"多倍 {value}",
            LineSpacingType.Fixed => $"固定值 {value}pt",
            LineSpacingType.AtLeast => $"最小值 {value}pt",
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
