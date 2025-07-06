using System;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using MessageBox = System.Windows.MessageBox;

namespace AllAutoOfficePDF2.Views
{
    /// <summary>
    /// プロジェクト編集ダイアログ
    /// </summary>
    public partial class ProjectEditDialog : Window
    {
        #region プロパティ
        /// <summary>
        /// プロジェクト名
        /// </summary>
        public string ProjectName { get; set; } = "";

        /// <summary>
        /// フォルダパス
        /// </summary>
        public string FolderPath { get; set; } = "";
        #endregion

        #region コンストラクタ
        public ProjectEditDialog()
        {
            InitializeComponent();
        }
        #endregion

        #region イベントハンドラ
        /// <summary>
        /// ウィンドウ読み込み時
        /// </summary>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            txtProjectName.Text = ProjectName;
            txtFolderPath.Text = FolderPath;
        }

        /// <summary>
        /// フォルダ選択ボタンクリック時
        /// </summary>
        private void BtnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "プロジェクトフォルダを選択してください";
                if (!string.IsNullOrEmpty(txtFolderPath.Text))
                {
                    dialog.SelectedPath = txtFolderPath.Text;
                }

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtFolderPath.Text = dialog.SelectedPath;
                    
                    // プロジェクト名が空の場合はフォルダ名を設定
                    if (string.IsNullOrEmpty(txtProjectName.Text))
                    {
                        txtProjectName.Text = Path.GetFileName(dialog.SelectedPath);
                    }
                }
            }
        }

        /// <summary>
        /// OKボタンクリック時
        /// </summary>
        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtProjectName.Text))
            {
                MessageBox.Show("プロジェクト名を入力してください。", "エラー", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtFolderPath.Text))
            {
                MessageBox.Show("フォルダを選択してください。", "エラー", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!Directory.Exists(txtFolderPath.Text))
            {
                MessageBox.Show("選択されたフォルダが存在しません。", "エラー", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            ProjectName = txtProjectName.Text.Trim();
            FolderPath = txtFolderPath.Text.Trim();
            DialogResult = true;
            Close();
        }

        /// <summary>
        /// キャンセルボタンクリック時
        /// </summary>
        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
        #endregion
    }
}