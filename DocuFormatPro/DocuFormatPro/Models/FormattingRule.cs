using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json.Serialization;

namespace DocuFormatPro.Models
{
    /// <summary>
    /// 完整排版规则数据模型，支持 UI 绑定和 JSON 序列化
    /// </summary>
    public class FormattingRule : INotifyPropertyChanged
    {
        private string _ruleName = "默认模板";
        private PageMarginSettings _pageMargins = new();
        private BodyTextSettings _bodyText = new();
        private ParagraphSettings _paragraph = new();
        private TableSettings _table = new();
        private List<HeadingStyle> _headings = new();
        private FrontMatterSettings _frontMatter = new();
        private HeadingNumberingSettings _headingNumbering = new();

        /// <summary>规则/模板名称</summary>
        public string RuleName
        {
            get => _ruleName;
            set { _ruleName = value; OnPropertyChanged(); }
        }

        /// <summary>页面边距设置</summary>
        public PageMarginSettings PageMargins
        {
            get => _pageMargins;
            set { _pageMargins = value; OnPropertyChanged(); }
        }

        /// <summary>正文字体设置</summary>
        public BodyTextSettings BodyText
        {
            get => _bodyText;
            set { _bodyText = value; OnPropertyChanged(); }
        }

        /// <summary>段落格式设置</summary>
        public ParagraphSettings Paragraph
        {
            get => _paragraph;
            set { _paragraph = value; OnPropertyChanged(); }
        }

        /// <summary>表格格式设置</summary>
        public TableSettings Table
        {
            get => _table;
            set { _table = value; OnPropertyChanged(); }
        }

        /// <summary>标题样式列表 (标题1~N)</summary>
        public List<HeadingStyle> Headings
        {
            get => _headings;
            set { _headings = value; OnPropertyChanged(); }
        }

        /// <summary>前置页设置</summary>
        public FrontMatterSettings FrontMatter
        {
            get => _frontMatter;
            set { _frontMatter = value; OnPropertyChanged(); }
        }

        /// <summary>标题自动编号设置</summary>
        public HeadingNumberingSettings HeadingNumbering
        {
            get => _headingNumbering;
            set { _headingNumbering = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// 返回包含用户指定默认值的排版规则
        /// </summary>
        public static FormattingRule CreateDefault()
        {
            return new FormattingRule
            {
                RuleName = "默认模板",
                PageMargins = new PageMarginSettings
                {
                    TopMargin = 2.54f,
                    BottomMargin = 2.54f,
                    LeftMargin = 2.8f,
                    RightMargin = 2.6f
                },
                BodyText = new BodyTextSettings
                {
                    ChineseFontName = "宋体",
                    EnglishFontName = "Times New Roman",
                    FontSizePoint = 12f,    // 小四 = 12pt
                    FontSizeName = "小四",
                    IsBold = false
                },
                Paragraph = new ParagraphSettings
                {
                    FirstLineIndentChars = 2f,
                    LineSpacingType = LineSpacingType.Multiple,
                    LineSpacingValue = 1.4f,
                    SpaceBeforeLines = 0f,
                    SpaceAfterLines = 0f
                },
                Table = new TableSettings
                {
                    UseSameAsBody = true,
                    HeaderBold = true,
                    SpaceBeforeLines = 0f,
                    SpaceAfterLines = 0f,
                    LineSpacingType = LineSpacingType.Fixed,
                    LineSpacingValue = 18f,  // 固定值 18 磅
                    BorderStyle = TableBorderStyle.SingleThin,
                    BorderColorHex = "#000000",
                    CellVerticalAlignment = CellVerticalAlign.Center,
                    CellHorizontalAlignment = CellHorizontalAlign.Center,
                    RepeatHeaderRow = true
                },
                Headings = new List<HeadingStyle>
                {
                    new HeadingStyle
                    {
                        Level = 1,
                        ChineseFontName = "宋体",
                        EnglishFontName = "Times New Roman",
                        FontSizePoint = 15f,    // 小三 = 15pt
                        FontSizeName = "小三",
                        IsBold = true,
                        Alignment = TextAlignment.Left,
                        SpaceBeforeLines = 0f,
                        SpaceAfterLines = 0f,
                        LineSpacingType = LineSpacingType.Multiple,
                        LineSpacingValue = 1.4f
                    },
                    new HeadingStyle
                    {
                        Level = 2,
                        ChineseFontName = "宋体",
                        EnglishFontName = "Times New Roman",
                        FontSizePoint = 14f,    // 四号 = 14pt
                        FontSizeName = "四号",
                        IsBold = true,
                        Alignment = TextAlignment.Left,
                        SpaceBeforeLines = 0f,
                        SpaceAfterLines = 0f,
                        LineSpacingType = LineSpacingType.Multiple,
                        LineSpacingValue = 1.4f
                    },
                    new HeadingStyle
                    {
                        Level = 3,
                        ChineseFontName = "宋体",
                        EnglishFontName = "Times New Roman",
                        FontSizePoint = 12f,    // 小四 = 12pt
                        FontSizeName = "小四",
                        IsBold = true,
                        Alignment = TextAlignment.Left,
                        SpaceBeforeLines = 0f,
                        SpaceAfterLines = 0f,
                        LineSpacingType = LineSpacingType.Multiple,
                        LineSpacingValue = 1.4f
                    }
                }
            };
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    #region 子设置类

    /// <summary>页面边距设置</summary>
    public class PageMarginSettings : INotifyPropertyChanged
    {
        private float _topMargin = 2.54f;
        private float _bottomMargin = 2.54f;
        private float _leftMargin = 2.8f;
        private float _rightMargin = 2.6f;

        public float TopMargin { get => _topMargin; set { _topMargin = value; OnPropertyChanged(); } }
        public float BottomMargin { get => _bottomMargin; set { _bottomMargin = value; OnPropertyChanged(); } }
        public float LeftMargin { get => _leftMargin; set { _leftMargin = value; OnPropertyChanged(); } }
        public float RightMargin { get => _rightMargin; set { _rightMargin = value; OnPropertyChanged(); } }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>正文字体设置</summary>
    public class BodyTextSettings : INotifyPropertyChanged
    {
        private string _chineseFontName = "宋体";
        private string _englishFontName = "Times New Roman";
        private float _fontSizePoint = 12f;
        private string _fontSizeName = "小四";
        private bool _isBold;
        private bool _useCustomFontColor = true;
        private string _fontColorHex = "#000000";

        public string ChineseFontName { get => _chineseFontName; set { _chineseFontName = value; OnPropertyChanged(); } }
        public string EnglishFontName { get => _englishFontName; set { _englishFontName = value; OnPropertyChanged(); } }
        public float FontSizePoint { get => _fontSizePoint; set { _fontSizePoint = value; OnPropertyChanged(); } }
        public string FontSizeName { get => _fontSizeName; set { _fontSizeName = value; OnPropertyChanged(); } }
        public bool IsBold { get => _isBold; set { _isBold = value; OnPropertyChanged(); } }
        /// <summary>是否启用自定义字体颜色</summary>
        public bool UseCustomFontColor { get => _useCustomFontColor; set { _useCustomFontColor = value; OnPropertyChanged(); } }
        /// <summary>字体颜色（十六进制，如 #000000）</summary>
        public string FontColorHex { get => _fontColorHex; set { _fontColorHex = value; OnPropertyChanged(); } }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>段落格式设置</summary>
    public class ParagraphSettings : INotifyPropertyChanged
    {
        private float _firstLineIndentChars = 2f;
        private LineSpacingType _lineSpacingType = LineSpacingType.Multiple;
        private float _lineSpacingValue = 1.4f;
        private float _spaceBeforeLines;
        private float _spaceAfterLines;

        public float FirstLineIndentChars { get => _firstLineIndentChars; set { _firstLineIndentChars = value; OnPropertyChanged(); } }
        public LineSpacingType LineSpacingType { get => _lineSpacingType; set { _lineSpacingType = value; OnPropertyChanged(); } }
        public float LineSpacingValue { get => _lineSpacingValue; set { _lineSpacingValue = value; OnPropertyChanged(); } }
        public float SpaceBeforeLines { get => _spaceBeforeLines; set { _spaceBeforeLines = value; OnPropertyChanged(); } }
        public float SpaceAfterLines { get => _spaceAfterLines; set { _spaceAfterLines = value; OnPropertyChanged(); } }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>表格格式设置</summary>
    public class TableSettings : INotifyPropertyChanged
    {
        private bool _applyTableFormatting = false;
        private bool _applyTableCaptions = false;
        private bool _useSameAsBody = true;
        private bool _headerBold = true;
        private float _spaceBeforeLines;
        private float _spaceAfterLines;
        private LineSpacingType _lineSpacingType = LineSpacingType.Fixed;
        private float _lineSpacingValue = 18f;
        private TableBorderStyle _borderStyle = TableBorderStyle.SingleThin;
        private string _borderColorHex = "#000000";
        private CellVerticalAlign _cellVerticalAlignment = CellVerticalAlign.Center;
        private CellHorizontalAlign _cellHorizontalAlignment = CellHorizontalAlign.Center;
        private bool _repeatHeaderRow = true;

        /// <summary>是否应用表格样式格式化</summary>
        public bool ApplyTableFormatting { get => _applyTableFormatting; set { _applyTableFormatting = value; OnPropertyChanged(); } }
        /// <summary>是否自动处理表格和图片题注</summary>
        public bool ApplyTableCaptions { get => _applyTableCaptions; set { _applyTableCaptions = value; OnPropertyChanged(); } }
        /// <summary>字体是否与正文一致</summary>
        public bool UseSameAsBody { get => _useSameAsBody; set { _useSameAsBody = value; OnPropertyChanged(); } }
        /// <summary>表头行加粗</summary>
        public bool HeaderBold { get => _headerBold; set { _headerBold = value; OnPropertyChanged(); } }
        public float SpaceBeforeLines { get => _spaceBeforeLines; set { _spaceBeforeLines = value; OnPropertyChanged(); } }
        public float SpaceAfterLines { get => _spaceAfterLines; set { _spaceAfterLines = value; OnPropertyChanged(); } }
        public LineSpacingType LineSpacingType { get => _lineSpacingType; set { _lineSpacingType = value; OnPropertyChanged(); } }
        public float LineSpacingValue { get => _lineSpacingValue; set { _lineSpacingValue = value; OnPropertyChanged(); } }
        public TableBorderStyle BorderStyle { get => _borderStyle; set { _borderStyle = value; OnPropertyChanged(); } }
        public string BorderColorHex { get => _borderColorHex; set { _borderColorHex = value; OnPropertyChanged(); } }
        public CellVerticalAlign CellVerticalAlignment { get => _cellVerticalAlignment; set { _cellVerticalAlignment = value; OnPropertyChanged(); } }
        public CellHorizontalAlign CellHorizontalAlignment { get => _cellHorizontalAlignment; set { _cellHorizontalAlignment = value; OnPropertyChanged(); } }
        /// <summary>跨页时重复显示表头</summary>
        public bool RepeatHeaderRow { get => _repeatHeaderRow; set { _repeatHeaderRow = value; OnPropertyChanged(); } }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>标题样式定义</summary>
    public class HeadingStyle : INotifyPropertyChanged
    {
        private int _level = 1;
        private string _chineseFontName = "黑体";
        private string _englishFontName = "Times New Roman";
        private float _fontSizePoint = 16f;
        private string _fontSizeName = "三号";
        private bool _isBold = true;
        private TextAlignment _alignment = TextAlignment.Left;
        private float _spaceBeforeLines = 0.5f;
        private float _spaceAfterLines = 0.5f;
        private LineSpacingType _lineSpacingType = LineSpacingType.Multiple;
        private float _lineSpacingValue = 1.4f;
        private bool _useCustomFontColor = true;
        private string _fontColorHex = "#000000";

        public int Level { get => _level; set { _level = value; OnPropertyChanged(); } }
        public string ChineseFontName { get => _chineseFontName; set { _chineseFontName = value; OnPropertyChanged(); } }
        public string EnglishFontName { get => _englishFontName; set { _englishFontName = value; OnPropertyChanged(); } }
        public float FontSizePoint { get => _fontSizePoint; set { _fontSizePoint = value; OnPropertyChanged(); } }
        public string FontSizeName { get => _fontSizeName; set { _fontSizeName = value; OnPropertyChanged(); } }
        public bool IsBold { get => _isBold; set { _isBold = value; OnPropertyChanged(); } }
        public TextAlignment Alignment { get => _alignment; set { _alignment = value; OnPropertyChanged(); } }
        public float SpaceBeforeLines { get => _spaceBeforeLines; set { _spaceBeforeLines = value; OnPropertyChanged(); } }
        public float SpaceAfterLines { get => _spaceAfterLines; set { _spaceAfterLines = value; OnPropertyChanged(); } }
        public LineSpacingType LineSpacingType { get => _lineSpacingType; set { _lineSpacingType = value; OnPropertyChanged(); } }
        public float LineSpacingValue { get => _lineSpacingValue; set { _lineSpacingValue = value; OnPropertyChanged(); } }
        /// <summary>是否启用自定义字体颜色</summary>
        public bool UseCustomFontColor { get => _useCustomFontColor; set { _useCustomFontColor = value; OnPropertyChanged(); } }
        /// <summary>字体颜色（十六进制，如 #000000）</summary>
        public string FontColorHex { get => _fontColorHex; set { _fontColorHex = value; OnPropertyChanged(); } }

        /// <summary>显示名称（用于 UI）</summary>
        [JsonIgnore]
        public string DisplayName => $"标题 {Level}";

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    #endregion

    #region 子设置类 - 附加功能

    /// <summary>前置页设置（例如：自动从模板粘贴封面和前言）</summary>
    public class FrontMatterSettings : INotifyPropertyChanged
    {
        private bool _insertFrontMatter = false;
        private string _templateFilePath = "";

        public bool InsertFrontMatter { get => _insertFrontMatter; set { _insertFrontMatter = value; OnPropertyChanged(); } }
        public string TemplateFilePath { get => _templateFilePath; set { _templateFilePath = value; OnPropertyChanged(); } }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>标题自动编号设置</summary>
    public class HeadingNumberingSettings : INotifyPropertyChanged
    {
        private bool _enableNumbering = false;
        private bool _stripExistingNumbers = false;
        private HeadingNumberingScheme _scheme = HeadingNumberingScheme.Numeric;

        /// <summary>是否启用标题自动编号</summary>
        public bool EnableNumbering
        {
            get => _enableNumbering;
            set { _enableNumbering = value; OnPropertyChanged(); }
        }

        /// <summary>是否先去除标题文字中已有的编号前缀</summary>
        public bool StripExistingNumbers
        {
            get => _stripExistingNumbers;
            set { _stripExistingNumbers = value; OnPropertyChanged(); }
        }

        /// <summary>编号方案</summary>
        public HeadingNumberingScheme Scheme
        {
            get => _scheme;
            set { _scheme = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    #endregion

    #region 枚举定义

    public enum LineSpacingType
    {
        Single,     // 单倍行距
        OneAndHalf, // 1.5 倍行距
        Double,     // 双倍行距
        Multiple,   // 多倍行距
        Fixed,      // 固定值 (磅)
        AtLeast     // 最小值 (磅)
    }

    public enum TextAlignment
    {
        Left,
        Center,
        Right,
        Justify // 两端对齐
    }

    public enum TableBorderStyle
    {
        None,
        SingleThin,     // 单细线
        SingleThick,    // 单粗线
        Double          // 双线
    }

    public enum CellVerticalAlign
    {
        Top,
        Center,
        Bottom
    }

    public enum CellHorizontalAlign
    {
        Left,
        Center,
        Right
    }

    public enum HeadingNumberingScheme
    {
        Numeric,        // 1 / 1.1 / 1.1.1
        ChapterNumeric, // 第一章 / 1.1 / 1.1.1
        Traditional     // 一、/ （一）/ 1.
    }

    #endregion

    #region 字号映射工具

    /// <summary>
    /// 中文字号 ↔ 磅值 映射工具
    /// </summary>
    public static class FontSizeMapping
    {
        private static readonly Dictionary<string, float> NameToPoint = new()
        {
            { "八号", 5f },
            { "七号", 5.5f },
            { "小六", 6.5f },
            { "六号", 7.5f },
            { "小五", 9f },
            { "五号", 10.5f },
            { "小四", 12f },
            { "四号", 14f },
            { "小三", 15f },
            { "三号", 16f },
            { "小二", 18f },
            { "二号", 22f },
            { "小一", 24f },
            { "一号", 26f },
            { "小初", 36f },
            { "初号", 42f }
        };

        public static readonly List<string> AllSizeNames = new(NameToPoint.Keys);

        public static float GetPoint(string sizeName)
            => NameToPoint.TryGetValue(sizeName, out var pt) ? pt : 12f;

        public static string GetName(float point)
        {
            foreach (var kvp in NameToPoint)
            {
                if (Math.Abs(kvp.Value - point) < 0.1f)
                    return kvp.Key;
            }
            return $"{point}pt";
        }
    }

    #endregion
}
