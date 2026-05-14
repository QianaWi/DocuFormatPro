using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using DocuFormatPro.Models;
using DocuFormatPro.Services;
using Microsoft.Win32;

namespace DocuFormatPro.ViewModels
{
    /// <summary>
    /// 主窗口 ViewModel，管理文件队列、排版规则和处理流程
    /// </summary>
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly WordProcessingService _wordService;
        private readonly TemplateExtractorService _extractorService;
        private readonly RuleStorageService _ruleStorage;
        private CancellationTokenSource? _cts;

        /// <summary>排版前确认回调，由 View 层注入；返回 false 则取消处理</summary>
        public Func<FormattingRule, int, bool>? ConfirmBeforeProcessing { get; set; }

        private bool _isProcessing;
        private double _progressValue;
        private string _statusText = "就绪 — 拖拽 Word 文档到窗口中开始";
        private bool _hasFiles;
        private FormattingRule _currentRule;
        private string? _selectedSavedRuleName;

        public MainViewModel()
        {
            _wordService = new WordProcessingService();
            _extractorService = new TemplateExtractorService();
            _ruleStorage = new RuleStorageService();

            // 加载默认规则
            _currentRule = _ruleStorage.GetDefaultRule();

            FileItems = new ObservableCollection<FileItem>();
            FileItems.CollectionChanged += (_, _) => HasFiles = FileItems.Count > 0;

            SavedRuleNames = new ObservableCollection<string>();
            RefreshSavedRuleNames();

            // 字体列表和字号列表
            ChineseFontNames = new ObservableCollection<string>
            {
                "宋体", "黑体", "楷体", "仿宋", "微软雅黑",
                "华文中宋", "华文楷体", "华文仿宋", "方正小标宋简体"
            };
            EnglishFontNames = new ObservableCollection<string>
            {
                "Times New Roman", "Arial", "Calibri", "Cambria",
                "Courier New", "Verdana", "Georgia"
            };
            FontSizeNames = new ObservableCollection<string>(FontSizeMapping.AllSizeNames);
            LineSpacingTypes = new ObservableCollection<string>
            {
                "单倍行距", "1.5倍行距", "双倍行距", "多倍行距", "固定值", "最小值"
            };
            AlignmentTypes = new ObservableCollection<string>
            {
                "左对齐", "居中", "右对齐", "两端对齐"
            };

            // 初始化命令
            AddFilesCommand = new RelayCommand(_ => AddFiles(), _ => !IsProcessing);
            RemoveFileCommand = new RelayCommand(param => RemoveFile(param), _ => !IsProcessing);
            StartProcessingCommand = new RelayCommand(_ => _ = StartProcessingAsync(), _ => HasFiles && !IsProcessing);
            ClearListCommand = new RelayCommand(_ => ClearList(), _ => HasFiles && !IsProcessing);
            CancelCommand = new RelayCommand(_ => CancelProcessing(), _ => IsProcessing);
            ImportTemplateCommand = new RelayCommand(_ => _ = ImportTemplateAsync(), _ => !IsProcessing);
            SaveRuleCommand = new RelayCommand(_ => SaveCurrentRule(), _ => !IsProcessing);
            LoadRuleCommand = new RelayCommand(_ => LoadSelectedRule(), _ => !IsProcessing && SelectedSavedRuleName != null);
            DeleteRuleCommand = new RelayCommand(_ => DeleteSelectedRule(), _ => !IsProcessing && SelectedSavedRuleName != null);
            ResetToDefaultCommand = new RelayCommand(_ => ResetToDefault(), _ => !IsProcessing);
            SelectFrontMatterTemplateCommand = new RelayCommand(_ => SelectFrontMatterTemplate(), _ => !IsProcessing);
        }

        #region Properties

        public ObservableCollection<FileItem> FileItems { get; }
        public ObservableCollection<string> SavedRuleNames { get; }
        public ObservableCollection<string> ChineseFontNames { get; }
        public ObservableCollection<string> EnglishFontNames { get; }
        public ObservableCollection<string> FontSizeNames { get; }
        public ObservableCollection<string> LineSpacingTypes { get; }
        public ObservableCollection<string> AlignmentTypes { get; }

        /// <summary>当前排版规则</summary>
        public FormattingRule CurrentRule
        {
            get => _currentRule;
            set { _currentRule = value; OnPropertyChanged(); }
        }

        /// <summary>已保存规则列表中选中的名称</summary>
        public string? SelectedSavedRuleName
        {
            get => _selectedSavedRuleName;
            set { _selectedSavedRuleName = value; OnPropertyChanged(); }
        }

        public bool IsProcessing
        {
            get => _isProcessing;
            set { _isProcessing = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsNotProcessing)); }
        }

        public bool IsNotProcessing => !_isProcessing;

        public double ProgressValue
        {
            get => _progressValue;
            set { _progressValue = value; OnPropertyChanged(); }
        }

        public string StatusText
        {
            get => _statusText;
            set { _statusText = value; OnPropertyChanged(); }
        }

        public bool HasFiles
        {
            get => _hasFiles;
            set { _hasFiles = value; OnPropertyChanged(); }
        }

        #endregion

        #region Commands

        public RelayCommand AddFilesCommand { get; }
        public RelayCommand RemoveFileCommand { get; }
        public RelayCommand StartProcessingCommand { get; }
        public RelayCommand ClearListCommand { get; }
        public RelayCommand CancelCommand { get; }
        public RelayCommand ImportTemplateCommand { get; }
        public RelayCommand SaveRuleCommand { get; }
        public RelayCommand LoadRuleCommand { get; }
        public RelayCommand DeleteRuleCommand { get; }
        public RelayCommand ResetToDefaultCommand { get; }
        public RelayCommand SelectFrontMatterTemplateCommand { get; }

        #endregion

        #region 文件管理

        private void AddFiles()
        {
            var dialog = new OpenFileDialog
            {
                Title = "选择 Word 文档",
                Filter = "Word 文档|*.doc;*.docx|所有文件|*.*",
                Multiselect = true
            };
            if (dialog.ShowDialog() == true)
            {
                foreach (var f in dialog.FileNames) AddFileToQueue(f);
            }
        }

        public void AddFileToQueue(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".doc" && ext != ".docx") return;

            var existing = FileItems.FirstOrDefault(f => f.FilePath == filePath);
            if (existing != null)
            {
                // 终态（Completed/Failed）条目重置为 Queued，允许重新处理
                // 活跃态（Queued/Processing）保持不变，避免干扰进行中的任务
                if (existing.Status == FileStatus.Completed || existing.Status == FileStatus.Failed)
                {
                    existing.Status = FileStatus.Queued;
                    existing.StatusMessage = "排队中";
                }
                return;
            }

            var fi = new FileInfo(filePath);
            FileItems.Add(new FileItem
            {
                FileName = fi.Name,
                FilePath = filePath,
                FileSize = fi.Length,
                Status = FileStatus.Queued,
                StatusMessage = "排队中"
            });
        }

        private void RemoveFile(object? param)
        {
            if (param is FileItem item) FileItems.Remove(item);
        }

        private void ClearList()
        {
            FileItems.Clear();
            ProgressValue = 0;
            StatusText = "就绪 — 拖拽 Word 文档到窗口中开始";
        }

        private void CancelProcessing()
        {
            _cts?.Cancel();
            StatusText = "正在取消任务...";
        }

        #endregion

        #region 排版处理

        private async Task StartProcessingAsync()
        {
            if (FileItems.Count == 0) return;

            // 弹出二次确认对话框
            if (ConfirmBeforeProcessing != null)
            {
                bool confirmed = ConfirmBeforeProcessing(CurrentRule, FileItems.Count);
                if (!confirmed) return;
            }

            IsProcessing = true;
            ProgressValue = 0;
            _cts = new CancellationTokenSource();

            int total = FileItems.Count;
            int processed = 0;

            try
            {
                foreach (var item in FileItems.ToList())
                {
                    if (_cts.Token.IsCancellationRequested) break;

                    item.Status = FileStatus.Processing;
                    item.StatusMessage = "处理中...";
                    StatusText = $"正在处理 第 {processed + 1}/{total} 个文件: {item.FileName}...";

                    var reporter = new Progress<string>(msg => item.StatusMessage = msg);

                    try
                    {
                        await _wordService.ProcessDocumentAsync(item.FilePath, CurrentRule, reporter, _cts.Token);
                        item.Status = FileStatus.Completed;
                        item.StatusMessage = "已完成";
                    }
                    catch (OperationCanceledException)
                    {
                        item.Status = FileStatus.Failed;
                        item.StatusMessage = "已取消";
                        break;
                    }
                    catch (Exception ex)
                    {
                        item.Status = FileStatus.Failed;
                        item.StatusMessage = $"失败: {ex.Message}";
                    }

                    processed++;
                    ProgressValue = (double)processed / total * 100;
                }

                if (_cts.Token.IsCancellationRequested)
                {
                    StatusText = "任务已取消";
                    foreach (var i in FileItems.Where(f => f.Status == FileStatus.Queued))
                    {
                        i.Status = FileStatus.Failed;
                        i.StatusMessage = "已取消";
                    }
                }
                else
                {
                    int ok = FileItems.Count(f => f.Status == FileStatus.Completed);
                    int fail = FileItems.Count(f => f.Status == FileStatus.Failed);
                    StatusText = $"处理完成 — 成功: {ok}, 失败: {fail}";
                    ProgressValue = 100;
                }
            }
            finally
            {
                IsProcessing = false;
                _cts?.Dispose();
                _cts = null;
            }
        }

        #endregion

        #region 模板管理

        /// <summary>从文档提取模板</summary>
        private async Task ImportTemplateAsync()
        {
            var dialog = new OpenFileDialog
            {
                Title = "选择模板文档（将从中提取排版格式）",
                Filter = "Word 文档|*.doc;*.docx"
            };
            if (dialog.ShowDialog() != true) return;

            IsProcessing = true;
            StatusText = "正在从文档提取模板...";

            try
            {
                var reporter = new Progress<string>(msg => StatusText = msg);
                var rule = await _extractorService.ExtractFromDocumentAsync(dialog.FileName, reporter);
                rule.RuleName = Path.GetFileNameWithoutExtension(dialog.FileName);
                CurrentRule = rule;
                StatusText = $"模板提取完成: {rule.RuleName}";
            }
            catch (Exception ex)
            {
                StatusText = $"模板提取失败: {ex.Message}";
            }
            finally
            {
                IsProcessing = false;
            }
        }

        /// <summary>保存当前规则</summary>
        private void SaveCurrentRule()
        {
            string name = CurrentRule.RuleName;
            if (string.IsNullOrWhiteSpace(name)) name = "自定义模板";

            _ruleStorage.SaveRule(CurrentRule, name);
            RefreshSavedRuleNames();
            StatusText = $"模板已保存: {name}";
        }

        /// <summary>加载选中的规则</summary>
        private void LoadSelectedRule()
        {
            if (SelectedSavedRuleName == null) return;
            var rule = _ruleStorage.LoadRule(SelectedSavedRuleName);
            if (rule != null)
            {
                CurrentRule = rule;
                StatusText = $"已加载模板: {rule.RuleName}";
            }
        }

        /// <summary>删除选中的规则</summary>
        private void DeleteSelectedRule()
        {
            if (SelectedSavedRuleName == null) return;
            _ruleStorage.DeleteRule(SelectedSavedRuleName);
            RefreshSavedRuleNames();
            StatusText = "模板已删除";
        }

        /// <summary>恢复默认</summary>
        private void ResetToDefault()
        {
            CurrentRule = FormattingRule.CreateDefault();
            StatusText = "已恢复默认排版规则";
        }

        private void RefreshSavedRuleNames()
        {
            SavedRuleNames.Clear();
            foreach (var name in _ruleStorage.ListSavedRules())
                SavedRuleNames.Add(name);
        }

        /// <summary>选择前置页模板文档</summary>
        private void SelectFrontMatterTemplate()
        {
            var dialog = new OpenFileDialog
            {
                Title = "选择前置页模板文档",
                Filter = "Word 文档|*.doc;*.docx"
            };
            if (dialog.ShowDialog() == true)
            {
                CurrentRule.FrontMatter.TemplateFilePath = dialog.FileName;
            }
        }

        #endregion

        #region 字号同步辅助

        /// <summary>当字号名称改变时同步磅值（用于 UI 下拉选择）</summary>
        public void SyncBodyFontSize(string sizeName)
        {
            CurrentRule.BodyText.FontSizeName = sizeName;
            CurrentRule.BodyText.FontSizePoint = FontSizeMapping.GetPoint(sizeName);
        }

        public void SyncHeadingFontSize(HeadingStyle heading, string sizeName)
        {
            heading.FontSizeName = sizeName;
            heading.FontSizePoint = FontSizeMapping.GetPoint(sizeName);
        }

        #endregion

        public void Cleanup()
        {
            _cts?.Cancel();
            _wordService.Dispose();
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
