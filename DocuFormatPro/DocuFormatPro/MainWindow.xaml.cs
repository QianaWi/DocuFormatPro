using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using DocuFormatPro.Models;
using DocuFormatPro.ViewModels;

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