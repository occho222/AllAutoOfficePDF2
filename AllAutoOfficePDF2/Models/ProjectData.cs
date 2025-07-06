using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text.Json.Serialization;

namespace AllAutoOfficePDF2.Models
{
    /// <summary>
    /// �v���W�F�N�g�f�[�^���f��
    /// </summary>
    public class ProjectData : INotifyPropertyChanged
    {
        private string _name = "";
        private bool _isActive = false;

        /// <summary>
        /// �v���W�F�N�gID
        /// </summary>
        public string Id { get; set; } = Guid.NewGuid().ToString();

        /// <summary>
        /// �v���W�F�N�g��
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
        /// �A�N�e�B�u���
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
        /// �v���W�F�N�g�t�H���_�̃p�X
        /// </summary>
        public string FolderPath { get; set; } = "";

        /// <summary>
        /// PDF�o�̓t�H���_�̃p�X
        /// </summary>
        public string PdfOutputFolder { get; set; } = "";

        /// <summary>
        /// ����PDF�t�@�C����
        /// </summary>
        public string MergeFileName { get; set; } = "����PDF";

        /// <summary>
        /// �y�[�W�ԍ��ǉ��t���O
        /// </summary>
        public bool AddPageNumber { get; set; } = false;

        /// <summary>
        /// �ŐV�̌���PDF�t�@�C���p�X
        /// </summary>
        public string LatestMergedPdfPath { get; set; } = "";

        /// <summary>
        /// �쐬����
        /// </summary>
        public DateTime CreatedDate { get; set; } = DateTime.Now;

        /// <summary>
        /// �ŏI�A�N�Z�X����
        /// </summary>
        public DateTime LastAccessDate { get; set; } = DateTime.Now;

        /// <summary>
        /// �t�@�C���A�C�e�����X�g
        /// </summary>
        public List<FileItemData> FileItems { get; set; } = new List<FileItemData>();

        /// <summary>
        /// �\�����iJSON��Ώہj
        /// </summary>
        [JsonIgnore]
        public string DisplayName => $"{Name} ({Path.GetFileName(FolderPath)})";

        /// <summary>
        /// �v���p�e�B�ύX�C�x���g
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged;

        /// <summary>
        /// �v���p�e�B�ύX�ʒm
        /// </summary>
        /// <param name="propertyName">�v���p�e�B��</param>
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}