using System;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using MessageBox = System.Windows.MessageBox;

namespace AllAutoOfficePDF2.Views
{
    /// <summary>
    /// �v���W�F�N�g�ҏW�_�C�A���O
    /// </summary>
    public partial class ProjectEditDialog : Window
    {
        #region �v���p�e�B
        /// <summary>
        /// �v���W�F�N�g��
        /// </summary>
        public string ProjectName { get; set; } = "";

        /// <summary>
        /// �t�H���_�p�X
        /// </summary>
        public string FolderPath { get; set; } = "";
        #endregion

        #region �R���X�g���N�^
        public ProjectEditDialog()
        {
            InitializeComponent();
        }
        #endregion

        #region �C�x���g�n���h��
        /// <summary>
        /// �E�B���h�E�ǂݍ��ݎ�
        /// </summary>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            txtProjectName.Text = ProjectName;
            txtFolderPath.Text = FolderPath;
        }

        /// <summary>
        /// �t�H���_�I���{�^���N���b�N��
        /// </summary>
        private void BtnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "�v���W�F�N�g�t�H���_��I�����Ă�������";
                if (!string.IsNullOrEmpty(txtFolderPath.Text))
                {
                    dialog.SelectedPath = txtFolderPath.Text;
                }

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtFolderPath.Text = dialog.SelectedPath;
                    
                    // �v���W�F�N�g������̏ꍇ�̓t�H���_����ݒ�
                    if (string.IsNullOrEmpty(txtProjectName.Text))
                    {
                        txtProjectName.Text = Path.GetFileName(dialog.SelectedPath);
                    }
                }
            }
        }

        /// <summary>
        /// OK�{�^���N���b�N��
        /// </summary>
        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtProjectName.Text))
            {
                MessageBox.Show("�v���W�F�N�g������͂��Ă��������B", "�G���[", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtFolderPath.Text))
            {
                MessageBox.Show("�t�H���_��I�����Ă��������B", "�G���[", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!Directory.Exists(txtFolderPath.Text))
            {
                MessageBox.Show("�I�����ꂽ�t�H���_�����݂��܂���B", "�G���[", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            ProjectName = txtProjectName.Text.Trim();
            FolderPath = txtFolderPath.Text.Trim();
            DialogResult = true;
            Close();
        }

        /// <summary>
        /// �L�����Z���{�^���N���b�N��
        /// </summary>
        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
        #endregion
    }
}