using System;

namespace AllAutoOfficePDF2.Models
{
    /// <summary>
    /// ファイルアイテムデータ（保存用）
    /// </summary>
    public class FileItemData
    {
        /// <summary>
        /// 選択状態
        /// </summary>
        public bool IsSelected { get; set; }

        /// <summary>
        /// 対象ページ
        /// </summary>
        public string TargetPages { get; set; } = "";

        /// <summary>
        /// ファイルパス
        /// </summary>
        public string FilePath { get; set; } = "";

        /// <summary>
        /// 最終更新日時
        /// </summary>
        public DateTime LastModified { get; set; }

        /// <summary>
        /// 表示順序
        /// </summary>
        public int DisplayOrder { get; set; } = 0;
    }
}