using System;

namespace AllAutoOfficePDF2.Models
{
    /// <summary>
    /// �t�@�C���A�C�e���f�[�^�i�ۑ��p�j
    /// </summary>
    public class FileItemData
    {
        /// <summary>
        /// �I�����
        /// </summary>
        public bool IsSelected { get; set; }

        /// <summary>
        /// �Ώۃy�[�W
        /// </summary>
        public string TargetPages { get; set; } = "";

        /// <summary>
        /// �t�@�C���p�X
        /// </summary>
        public string FilePath { get; set; } = "";

        /// <summary>
        /// �ŏI�X�V����
        /// </summary>
        public DateTime LastModified { get; set; }

        /// <summary>
        /// �\������
        /// </summary>
        public int DisplayOrder { get; set; } = 0;
    }
}