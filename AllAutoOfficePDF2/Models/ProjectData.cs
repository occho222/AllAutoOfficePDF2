using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text.Json.Serialization;

namespace AllAutoOfficePDF2.Models
{
    /// <summary>
    /// プロジェクトデータモデル
    /// </summary>
    public class ProjectData : INotifyPropertyChanged
    {
        private string _name = "";
        private bool _isActive = false;

        /// <summary>
        /// プロジェクトID
        /// </summary>
        public string Id { get; set; } = Guid.NewGuid().ToString();

        /// <summary>
        /// プロジェクト名
        /// </summary>
        public string Name
        {
            get => _name;
            set
            {
                _name = value;
                OnPropertyChanged(nameof(Name));
            }
        }

        /// <summary>
        /// アクティブ状態
        /// </summary>
        public bool IsActive
        {
            get => _isActive;
            set
            {
                _isActive = value;
                OnPropertyChanged(nameof(IsActive));
            }
        }

        /// <summary>
        /// プロジェクトフォルダのパス
        /// </summary>
        public string FolderPath { get; set; } = "";

        /// <summary>
        /// PDF出力フォルダのパス
        /// </summary>
        public string PdfOutputFolder { get; set; } = "";

        /// <summary>
        /// 結合PDFファイル名
        /// </summary>
        public string MergeFileName { get; set; } = "結合PDF";

        /// <summary>
        /// ページ番号追加フラグ
        /// </summary>
        public bool AddPageNumber { get; set; } = false;

        /// <summary>
        /// 最新の結合PDFファイルパス
        /// </summary>
        public string LatestMergedPdfPath { get; set; } = "";

        /// <summary>
        /// 作成日時
        /// </summary>
        public DateTime CreatedDate { get; set; } = DateTime.Now;

        /// <summary>
        /// 最終アクセス日時
        /// </summary>
        public DateTime LastAccessDate { get; set; } = DateTime.Now;

        /// <summary>
        /// ファイルアイテムリスト
        /// </summary>
        public List<FileItemData> FileItems { get; set; } = new List<FileItemData>();

        /// <summary>
        /// 表示名（JSON非対象）
        /// </summary>
        [JsonIgnore]
        public string DisplayName => $"{Name} ({Path.GetFileName(FolderPath)})";

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