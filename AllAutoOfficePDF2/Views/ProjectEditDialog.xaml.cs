using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using AllAutoOfficePDF2.Models;
using AllAutoOfficePDF2.Services;
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
        /// プロジェクトカテゴリ
        /// </summary>
        public string Category { get; set; } = "";

        /// <summary>
        /// フォルダパス
        /// </summary>
        public string FolderPath { get; set; } = "";

        /// <summary>
        /// サブフォルダを含むかどうか
        /// </summary>
        public bool IncludeSubfolders { get; set; } = false;

        /// <summary>
        /// カスタムPDF保存パスを使用するかどうか
        /// </summary>
        public bool UseCustomPdfPath { get; set; } = false;

        /// <summary>
        /// カスタムPDF保存パス
        /// </summary>
        public string CustomPdfPath { get; set; } = "";

        /// <summary>
        /// 利用可能なカテゴリリスト
        /// </summary>
        private List<string> availableCategories = new List<string>();
        #endregion

        #region コンストラクタ
        public ProjectEditDialog()
        {
            InitializeComponent();
            LoadAvailableCategories();
        }
        #endregion

        #region 初期化
        /// <summary>
        /// 利用可能なカテゴリを読み込み
        /// </summary>
        private void LoadAvailableCategories()
        {
            var allProjects = ProjectManager.LoadProjects();
            availableCategories = ProjectManager.GetAvailableCategories(allProjects);
            
            // よく使われるカテゴリを追加
            var defaultCategories = new List<string> { "業務", "個人", "開発", "資料", "アーカイブ" };
            foreach (var category in defaultCategories)
            {
                if (!availableCategories.Contains(category))
                {
                    availableCategories.Add(category);
                }
            }
            
            cmbCategory.ItemsSource = availableCategories;
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
            cmbCategory.Text = Category;
            chkIncludeSubfolders.IsChecked = IncludeSubfolders;
            chkUseCustomPdfPath.IsChecked = UseCustomPdfPath;
            txtCustomPdfPath.Text = CustomPdfPath;

            // カスタムPDFパスの有効/無効を設定
            UpdateCustomPdfPathEnabled();
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
        /// カスタムPDF保存パス選択ボタンクリック時
        /// </summary>
        private void BtnSelectCustomPdfPath_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "PDF保存フォルダを選択してください（フォルダパスのみが設定されます）";
                if (!string.IsNullOrEmpty(txtCustomPdfPath.Text))
                {
                    // 既存のパスがファイル名を含む場合は、ディレクトリパスのみを取得
                    var existingPath = txtCustomPdfPath.Text;
                    if (File.Exists(existingPath))
                    {
                        dialog.SelectedPath = Path.GetDirectoryName(existingPath) ?? "";
                    }
                    else if (Directory.Exists(existingPath))
                    {
                        dialog.SelectedPath = existingPath;
                    }
                    else
                    {
                        // パスの親ディレクトリが存在するかチェック
                        var parentDir = Path.GetDirectoryName(existingPath);
                        if (!string.IsNullOrEmpty(parentDir) && Directory.Exists(parentDir))
                        {
                            dialog.SelectedPath = parentDir;
                        }
                    }
                }

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    // フォルダパスのみを設定（ファイル名は含めない）
                    txtCustomPdfPath.Text = dialog.SelectedPath;
                }
            }
        }

        /// <summary>
        /// サブフォルダ読み込みチェック時
        /// </summary>
        private void ChkIncludeSubfolders_Checked(object sender, RoutedEventArgs e)
        {
            // サブフォルダを含む場合、カスタムPDFパスを必須にする
            chkUseCustomPdfPath.IsChecked = true;
            UpdateCustomPdfPathEnabled();
        }

        /// <summary>
        /// サブフォルダ読み込みチェック解除時
        /// </summary>
        private void ChkIncludeSubfolders_Unchecked(object sender, RoutedEventArgs e)
        {
            // サブフォルダを含まない場合は任意
            UpdateCustomPdfPathEnabled();
        }

        /// <summary>
        /// カスタムPDF保存パス使用チェック時
        /// </summary>
        private void ChkUseCustomPdfPath_Checked(object sender, RoutedEventArgs e)
        {
            UpdateCustomPdfPathEnabled();
        }

        /// <summary>
        /// カスタムPDF保存パス使用チェック解除時
        /// </summary>
        private void ChkUseCustomPdfPath_Unchecked(object sender, RoutedEventArgs e)
        {
            UpdateCustomPdfPathEnabled();
        }

        /// <summary>
        /// カスタムPDF保存パス入力欄の有効/無効を更新
        /// </summary>
        private void UpdateCustomPdfPathEnabled()
        {
            var includeSubfolders = chkIncludeSubfolders.IsChecked == true;
            var useCustomPdfPath = chkUseCustomPdfPath.IsChecked == true;
            
            // サブフォルダを含む場合は、カスタムPDFパスを強制的に有効にする
            if (includeSubfolders)
            {
                chkUseCustomPdfPath.IsChecked = true;
                chkUseCustomPdfPath.IsEnabled = false; // チェックボックスを無効化（必須）
                gridCustomPdfPath.IsEnabled = true;
            }
            else
            {
                chkUseCustomPdfPath.IsEnabled = true; // チェックボックスを有効化（任意）
                gridCustomPdfPath.IsEnabled = useCustomPdfPath;
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

            if (chkIncludeSubfolders.IsChecked == true && chkUseCustomPdfPath.IsChecked != true)
            {
                MessageBox.Show("サブフォルダを含む設定の場合、カスタムPDF保存パスの設定が必須です。", "エラー", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (chkUseCustomPdfPath.IsChecked == true)
            {
                if (string.IsNullOrWhiteSpace(txtCustomPdfPath.Text))
                {
                    MessageBox.Show("カスタムPDF保存パスを選択してください。", "エラー", 
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!Directory.Exists(txtCustomPdfPath.Text))
                {
                    var result = MessageBox.Show("指定されたPDF保存フォルダが存在しません。作成しますか？", "確認", 
                        MessageBoxButton.YesNo, MessageBoxImage.Question);
                    
                    if (result == MessageBoxResult.Yes)
                    {
                        try
                        {
                            Directory.CreateDirectory(txtCustomPdfPath.Text);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"フォルダの作成に失敗しました: {ex.Message}", "エラー", 
                                MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }
                    }
                    else
                    {
                        return;
                    }
                }
            }

            ProjectName = txtProjectName.Text.Trim();
            FolderPath = txtFolderPath.Text.Trim();
            Category = cmbCategory.Text?.Trim() ?? "";
            IncludeSubfolders = chkIncludeSubfolders.IsChecked == true;
            UseCustomPdfPath = chkUseCustomPdfPath.IsChecked == true;
            CustomPdfPath = txtCustomPdfPath.Text.Trim();
            
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