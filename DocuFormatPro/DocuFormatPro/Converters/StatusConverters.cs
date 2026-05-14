using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using DocuFormatPro.Models;

namespace DocuFormatPro.Converters
{
    /// <summary>
    /// 文件状态 → 颜色转换器
    /// </summary>
    public class StatusToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is FileStatus status)
            {
                return status switch
                {
                    FileStatus.Queued => new SolidColorBrush(Color.FromRgb(158, 158, 158)),      // 灰色
                    FileStatus.Processing => new SolidColorBrush(Color.FromRgb(108, 99, 255)),    // 紫蓝
                    FileStatus.Completed => new SolidColorBrush(Color.FromRgb(76, 175, 80)),      // 绿色
                    FileStatus.Failed => new SolidColorBrush(Color.FromRgb(244, 67, 54)),         // 红色
                    _ => new SolidColorBrush(Colors.Gray)
                };
            }
            return new SolidColorBrush(Colors.Gray);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }

    /// <summary>
    /// 文件状态 → 图标字符转换器
    /// </summary>
    public class StatusToIconConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is FileStatus status)
            {
                return status switch
                {
                    FileStatus.Queued => "⏳",
                    FileStatus.Processing => "⚙️",
                    FileStatus.Completed => "✅",
                    FileStatus.Failed => "❌",
                    _ => "•"
                };
            }
            return "•";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }

    /// <summary>
    /// 文件状态 → 中文文本转换器
    /// </summary>
    public class StatusToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is FileStatus status)
            {
                return status switch
                {
                    FileStatus.Queued => "排队中",
                    FileStatus.Processing => "处理中",
                    FileStatus.Completed => "已完成",
                    FileStatus.Failed => "失败",
                    _ => "未知"
                };
            }
            return "未知";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }

    /// <summary>
    /// 文件大小（字节）→ 可读字符串转换器 (KB/MB/GB)
    /// </summary>
    public class FileSizeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is long bytes)
            {
                string[] suffixes = { "B", "KB", "MB", "GB" };
                int order = 0;
                double size = bytes;
                while (size >= 1024 && order < suffixes.Length - 1)
                {
                    order++;
                    size /= 1024;
                }
                return $"{size:0.##} {suffixes[order]}";
            }
            return "0 B";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }

    /// <summary>
    /// 布尔值取反转换器（用于按钮启用/禁用状态）
    /// </summary>
    public class InverseBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool b) return !b;
            return true;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool b) return !b;
            return true;
        }
    }

    /// <summary>
    /// 布尔值 → Visibility 转换器
    /// </summary>
    public class BoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool b)
            {
                bool invert = parameter is string s && s == "Invert";
                return (b ^ invert) ? Visibility.Visible : Visibility.Collapsed;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }

    /// <summary>
    /// 集合数量 → Visibility 转换器（列表有项则显示）
    /// </summary>
    public class CountToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int count)
            {
                bool invert = parameter is string s && s == "Invert";
                bool hasItems = count > 0;
                return (hasItems ^ invert) ? Visibility.Visible : Visibility.Collapsed;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }

    /// <summary>
    /// 十六进制颜色字符串（#RRGGBB）→ SolidColorBrush 转换器，用于颜色预览色块
    /// </summary>
    public class HexColorToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string hex)
            {
                try
                {
                    hex = hex.TrimStart('#');
                    if (hex.Length == 6)
                    {
                        byte r = System.Convert.ToByte(hex.Substring(0, 2), 16);
                        byte g = System.Convert.ToByte(hex.Substring(2, 2), 16);
                        byte b = System.Convert.ToByte(hex.Substring(4, 2), 16);
                        return new SolidColorBrush(Color.FromRgb(r, g, b));
                    }
                }
                catch { }
            }
            return new SolidColorBrush(Colors.Black);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }
}
