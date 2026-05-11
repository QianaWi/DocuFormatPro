using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace DocuFormatPro.Models
{
    /// <summary>
    /// 文件处理状态枚举
    /// </summary>
    public enum FileStatus
    {
        Queued,      // 排队中
        Processing,  // 处理中
        Completed,   // 完成
        Failed       // 失败
    }

    /// <summary>
    /// 文件数据模型，用于在文件列表中展示和追踪处理状态
    /// </summary>
    public class FileItem : INotifyPropertyChanged
    {
        private string _fileName = string.Empty;
        private string _filePath = string.Empty;
        private long _fileSize;
        private FileStatus _status = FileStatus.Queued;
        private string _statusMessage = "排队中";

        /// <summary>
        /// 文件名（不含路径）
        /// </summary>
        public string FileName
        {
            get => _fileName;
            set { _fileName = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// 文件完整路径
        /// </summary>
        public string FilePath
        {
            get => _filePath;
            set { _filePath = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// 文件大小（字节）
        /// </summary>
        public long FileSize
        {
            get => _fileSize;
            set { _fileSize = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// 当前处理状态
        /// </summary>
        public FileStatus Status
        {
            get => _status;
            set { _status = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// 详细状态信息
        /// </summary>
        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
