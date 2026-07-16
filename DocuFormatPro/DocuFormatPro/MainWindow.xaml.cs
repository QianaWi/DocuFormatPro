using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using DocuFormatPro.Models;
using DocuFormatPro.ViewModels;
using DocuFormatPro.Views;

namespace DocuFormatPro
{
    public partial class MainWindow : Window
    {
        private readonly MainViewModel _viewModel;

        public MainWindow()
        {
            InitializeComponent();
            _viewModel = new MainViewModel();
            DataContext = _viewModel;

            // 注入排版前确认对话框
            _viewModel.ConfirmBeforeProcessing = (rule, fileCount) =>
            {
                var dialog = new ConfirmFormattingDialog(rule, fileCount)
                {
                    Owner = this
                };
                return dialog.ShowDialog() == true;
            };

            // 注入文件锁定警告（模态置顶）
            _viewModel.ShowFileLockedWarning = (filePath) =>
            {
                string fileName = System.IO.Path.GetFileName(filePath);
                var msg = new Window
                {
                    Title = "文件被占用",
                    Width = 380,
                    Height = 160,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    ResizeMode = ResizeMode.NoResize,
                    Owner = this,
                    Topmost = true,
                    WindowStyle = WindowStyle.ToolWindow,
                    Background = System.Windows.Media.Brushes.White
                };
                var panel = new StackPanel { Margin = new Thickness(24, 20, 24, 16) };
                panel.Children.Add(new TextBlock
                {
                    Text = $"⚠️  文档已被其他程序占用",
                    FontSize = 14,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0x1A, 0x1A, 0x2E)),
                    Margin = new Thickness(0, 0, 0, 8)
                });
                panel.Children.Add(new TextBlock
                {
                    Text = $"请先关闭 Word 中打开的「{fileName}」，再重新拖入。",
                    FontSize = 12,
                    Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0x6B, 0x72, 0x80)),
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 0, 0, 16)
                });
                var btn = new Button
                {
                    Content = "确定",
                    Width = 80,
                    Height = 30,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0x6C, 0x63, 0xFF)),
                    Foreground = System.Windows.Media.Brushes.White,
                    BorderThickness = new Thickness(0)
                };
                btn.Click += (_, _) => msg.Close();
                panel.Children.Add(btn);
                msg.Content = panel;
                msg.ShowDialog();
                return true;
            };
        }

        private void Window_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                bool hasWord = files.Any(f =>
                {
                    string ext = System.IO.Path.GetExtension(f).ToLowerInvariant();
                    return ext == ".doc" || ext == ".docx";
                });
                e.Effects = hasWord ? DragDropEffects.Copy : DragDropEffects.None;
                if (hasWord)
                {
                    DropOverlay.Visibility = Visibility.Visible;
                    DropZoneBorderBrush.Color = (Color)FindResource("PrimaryColor");
                }
            }
            else e.Effects = DragDropEffects.None;
            e.Handled = true;
        }

        private void Window_DragLeave(object sender, DragEventArgs e)
        {
            DropOverlay.Visibility = Visibility.Collapsed;
            DropZoneBorderBrush.Color = Color.FromRgb(0xD1, 0xD5, 0xDB);
        }

        private void Window_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            DropOverlay.Visibility = Visibility.Collapsed;
            DropZoneBorderBrush.Color = Color.FromRgb(0xD1, 0xD5, 0xDB);
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                foreach (var f in (string[])e.Data.GetData(DataFormats.FileDrop))
                    _viewModel.AddFileToQueue(f);
            }
            e.Handled = true;
        }

        private void DropZone_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (_viewModel.AddFilesCommand.CanExecute(null))
                _viewModel.AddFilesCommand.Execute(null);
        }

        private void BodyFontSize_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox cb && cb.SelectedItem is string sizeName)
                _viewModel.SyncBodyFontSize(sizeName);
        }

        private void HeadingFontSize_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox cb && cb.SelectedItem is string sizeName && cb.Tag is string tag)
            {
                if (int.TryParse(tag, out int level))
                    _viewModel.SyncHeadingFontSizeByLevel(level, sizeName);
            }
        }

        private void HeadingNumberingScheme_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox cb && cb.SelectedIndex >= 0)
                _viewModel.SyncHeadingNumberingScheme(cb.SelectedIndex);
        }

        private void HeadingLineSpacingType_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox cb && cb.SelectedIndex >= 0 && cb.Tag is string tag)
            {
                if (int.TryParse(tag, out int level))
                    _viewModel.SyncHeadingLineSpacingTypeByLevel(level, cb.SelectedIndex);
            }
        }

        private void TableFontSize_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox cb && cb.SelectedItem is string sizeName)
                _viewModel.SyncTableFontSize(sizeName);
        }

        private void BodyLineSpacingType_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox cb && cb.SelectedIndex >= 0)
            {
                _viewModel.CurrentRule.Paragraph.LineSpacingType = cb.SelectedIndex switch
                {
                    0 => LineSpacingType.Single,
                    1 => LineSpacingType.OneAndHalf,
                    2 => LineSpacingType.Double,
                    3 => LineSpacingType.Multiple,
                    4 => LineSpacingType.Fixed,
                    5 => LineSpacingType.AtLeast,
                    _ => LineSpacingType.Multiple
                };
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _viewModel.Cleanup();
        }
    }
}
