using System;
using System.ComponentModel;
using System.IO;

namespace AllAutoOfficePDF2.Models
{
    /// <summary>
    /// ファイルアイテムモデル
    /// </summary>
    public class FileItem : INotifyPropertyChanged
    {
        private bool _isSelected;
        private string _targetPages = "";
        private int _number;
        private string _relativePath = "";

        /// <summary>
        /// 選択状態
        /// </summary>
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged(nameof(IsSelected));
            }
        }

        /// <summary>
        /// 対象ページ
        /// </summary>
        public string TargetPages
        {
            get => _targetPages;
            set
            {
                _targetPages = value;
                OnPropertyChanged(nameof(TargetPages));
            }
        }

        /// <summary>
        /// 番号
        /// </summary>
        public int Number
        {
            get => _number;
            set
            {
                _number = value;
                OnPropertyChanged(nameof(Number));
            }
        }

        /// <summary>
        /// ファイル名
        /// </summary>
        public string FileName { get; set; } = "";

        /// <summary>
        /// ファイルパス
        /// </summary>
        public string FilePath { get; set; } = "";

        /// <summary>
        /// 拡張子
        /// </summary>
        public string Extension { get; set; } = "";

        /// <summary>
        /// 最終更新日時
        /// </summary>
        public DateTime LastModified { get; set; }

        /// <summary>
        /// PDFステータス
        /// </summary>
        public string PdfStatus { get; set; } = "";

        /// <summary>
        /// 表示順序
        /// </summary>
        public int DisplayOrder { get; set; } = 0;

        /// <summary>
        /// 相対パス（サブフォルダ読み込み用）
        /// </summary>
        public string RelativePath
        {
            get => _relativePath;
            set
            {
                _relativePath = value;
                OnPropertyChanged(nameof(RelativePath));
                OnPropertyChanged(nameof(FolderName));
            }
        }

        /// <summary>
        /// フォルダ名のみ（相対パスからフォルダ名部分のみを抽出）
        /// </summary>
        public string FolderName
        {
            get
            {
                if (string.IsNullOrEmpty(RelativePath))
                {
                    return "";
                }

                // RelativePathがファイルパスの場合、ディレクトリ部分のみを取得
                var directoryPath = Path.GetDirectoryName(RelativePath);
                
                if (string.IsNullOrEmpty(directoryPath))
                {
                    return "";
                }

                // 最後のフォルダ名のみを返す
                return Path.GetFileName(directoryPath);
            }
        }

        /// <summary>
        /// プロパティ変更イベント
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged;

        /// <summary>
        /// プロパティ変更通知
        /// </summary>
        /// <param name="propertyName">プロパティ名</param>
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}