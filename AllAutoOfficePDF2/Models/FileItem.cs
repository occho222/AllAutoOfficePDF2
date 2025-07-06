using System;
using System.ComponentModel;

namespace AllAutoOfficePDF2.Models
{
    /// <summary>
    /// �t�@�C���A�C�e�����f��
    /// </summary>
    public class FileItem : INotifyPropertyChanged
    {
        private bool _isSelected;
        private string _targetPages = "";
        private int _number;

        /// <summary>
        /// �I�����
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
        /// �Ώۃy�[�W
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
        /// �ԍ�
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
        /// �t�@�C����
        /// </summary>
        public string FileName { get; set; } = "";

        /// <summary>
        /// �t�@�C���p�X
        /// </summary>
        public string FilePath { get; set; } = "";

        /// <summary>
        /// �g���q
        /// </summary>
        public string Extension { get; set; } = "";

        /// <summary>
        /// �ŏI�X�V����
        /// </summary>
        public DateTime LastModified { get; set; }

        /// <summary>
        /// PDF�X�e�[�^�X
        /// </summary>
        public string PdfStatus { get; set; } = "";

        /// <summary>
        /// �\������
        /// </summary>
        public int DisplayOrder { get; set; } = 0;

        /// <summary>
        /// ���΃p�X�i�T�u�t�H���_�ǂݍ��ݗp�j
        /// </summary>
        public string RelativePath { get; set; } = "";

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