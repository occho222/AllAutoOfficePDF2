using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using AllAutoOfficePDF2.Models;
using AllAutoOfficePDF2.Services;
using AllAutoOfficePDF2.Views;
using MessageBox = System.Windows.MessageBox;
using DragEventArgs = System.Windows.DragEventArgs;
using DataFormats = System.Windows.DataFormats;
using DragDropEffects = System.Windows.DragDropEffects;

namespace AllAutoOfficePDF2
{
    /// <summary>
    /// メインウィンドウ
    /// </summary>
    public partial class MainWindow : Window
    {
        #region プライベートフィールド
        private ObservableCollection<FileItem> fileItems = new ObservableCollection<FileItem>();
        private ObservableCollection<ProjectCategoryGroup> categoryGroups = new ObservableCollection<ProjectCategoryGroup>();
        private ProjectData? currentProject = null;
        private string selectedFolderPath = "";
        private string pdfOutputFolder = "";
        #endregion

        #region コンストラクタ
        public MainWindow()
        {
            InitializeComponent();
            InitializeDataBindings();
            LoadProjects();
            RestoreActiveProject();
            UpdateProjectDisplay();
        }
        #endregion

        #region 初期化
        private void InitializeDataBindings()
        {
            dgFiles.ItemsSource = fileItems;
            treeProjects.ItemsSource = categoryGroups;
        }

        private void LoadProjects()
        {
            categoryGroups.Clear();
            var projectList = ProjectManager.LoadProjects();
            
            // 既存プロジェクトのアイコンを修正
            FixExistingProjectIcons(projectList);
            
            // 各プロジェクトのアイコンを確認・設定
            foreach (var project in projectList)
            {
                if (string.IsNullOrEmpty(project.CategoryIcon))
                {
                    project.CategoryIcon = GetDefaultCategoryIcon(project.Category);
                }
            }
            
            // カテゴリ別にグループ化
            var groupedProjects = projectList.GroupBy(p => string.IsNullOrEmpty(p.Category) ? "未分類" : p.Category)
                                            .OrderBy(g => g.Key == "未分類" ? "z" : g.Key)
                                            .ToList();

            foreach (var group in groupedProjects)
            {
                var categoryGroup = new ProjectCategoryGroup
                {
                    CategoryName = group.Key,
                    CategoryIcon = GetCategoryIcon(group.Key, group.First().CategoryIcon),
                    CategoryColor = GetCategoryColor(group.Key, group.First().CategoryColor)
                };

                // カテゴリ内でプロジェクト名順に並び替え
                var sortedProjects = group.OrderBy(p => p.Name).ToList();
                foreach (var project in sortedProjects)
                {
                    categoryGroup.Projects.Add(project);
                }

                categoryGroups.Add(categoryGroup);
            }
        }

        /// <summary>
        /// 既存プロジェクトのアイコンを修正
        /// </summary>
        private void FixExistingProjectIcons(List<ProjectData> projects)
        {
            bool needsSave = false;
            
            foreach (var project in projects)
            {
                // 空のアイコンやデフォルト値の修正
                if (string.IsNullOrEmpty(project.CategoryIcon) || project.CategoryIcon == "??")
                {
                    project.CategoryIcon = GetDefaultCategoryIcon(project.Category);
                    needsSave = true;
                }
                
                // 空の色やデフォルト値の修正
                if (string.IsNullOrEmpty(project.CategoryColor))
                {
                    project.CategoryColor = GetCategoryColor(project.Category, "");
                    needsSave = true;
                }
            }
            
            // 修正があった場合は保存
            if (needsSave)
            {
                ProjectManager.SaveProjects(projects);
            }
        }

        private string GetDefaultCategoryIcon(string category)
        {
            return category switch
            {
                "業務" => "💼",
                "プロジェクト" => "📊",
                "資料" => "📋",
                "マニュアル" => "📖",
                "提案書" => "📝",
                "報告書" => "📄",
                "会議" => "🗣️",
                "設計" => "⚙️",
                "テスト" => "🧪",
                "開発" => "💻",
                "運用" => "🔧",
                "保守" => "🛠️",
                "バックアップ" => "💾",
                "アーカイブ" => "📦",
                "一時的" => "⏱️",
                "進行中" => "🔄",
                "完了" => "✅",
                "保留" => "⏸️",
                "重要" => "⭐",
                "緊急" => "🚨",
                _ => "📁"
            };
        }

        private string GetCategoryIcon(string categoryName, string existingIcon)
        {
            if (!string.IsNullOrEmpty(existingIcon) && existingIcon != "📁")
            {
                return existingIcon;
            }
            return GetDefaultCategoryIcon(categoryName);
        }

        private string GetCategoryColor(string categoryName, string existingColor)
        {
            if (!string.IsNullOrEmpty(existingColor) && existingColor != "#E9ECEF")
            {
                return existingColor;
            }
            
            return categoryName switch
            {
                "業務" => "#007ACC",
                "プロジェクト" => "#28A745",
                "資料" => "#6C757D",
                "マニュアル" => "#17A2B8",
                "提案書" => "#FFC107",
                "報告書" => "#DC3545",
                "会議" => "#6F42C1",
                "設計" => "#FD7E14",
                "テスト" => "#20C997",
                "開発" => "#E83E8C",
                "運用" => "#6C757D",
                "保守" => "#495057",
                "バックアップ" => "#ADB5BD",
                "アーカイブ" => "#868E96",
                "一時的" => "#F8F9FA",
                "進行中" => "#007BFF",
                "完了" => "#28A745",
                "保留" => "#FFC107",
                "重要" => "#FF6B6B",
                "緊急" => "#DC3545",
                _ => "#E9ECEF"
            };
        }

        private void RestoreActiveProject()
        {
            var activeProject = GetAllProjects().FirstOrDefault(p => p.IsActive);
            if (activeProject != null)
            {
                SwitchToProject(activeProject);
            }
            else
            {
                UpdateLatestMergedPdfDisplay();
            }
        }

        /// <summary>
        /// 全プロジェクトを取得
        /// </summary>
        /// <returns>全プロジェクトのリスト</returns>
        private List<ProjectData> GetAllProjects()
        {
            var allProjects = new List<ProjectData>();
            foreach (var categoryGroup in categoryGroups)
            {
                allProjects.AddRange(categoryGroup.Projects);
            }
            return allProjects;
        }
        #endregion

        #region プロジェクト管理
        private void BtnNewProject_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new ProjectEditDialog();
            if (dialog.ShowDialog() == true)
            {
                var newProject = new ProjectData
                {
                    Name = dialog.ProjectName,
                    FolderPath = dialog.FolderPath,
                    Category = dialog.Category,
                    IncludeSubfolders = dialog.IncludeSubfolders,
                    UseCustomPdfPath = dialog.UseCustomPdfPath,
                    CustomPdfPath = dialog.CustomPdfPath,
                    CategoryIcon = GetDefaultCategoryIcon(dialog.Category),
                    CategoryColor = GetCategoryColor(dialog.Category, "")
                };

                // カテゴリグループに追加
                AddProjectToCategoryGroup(newProject);
                SwitchToProject(newProject);
                
                // プロジェクトリストを再構築
                RefreshProjectList();
            }
        }

        private void BtnEditProject_Click(object sender, RoutedEventArgs e)
        {
            if (treeProjects.SelectedItem is ProjectData selectedProject)
            {
                var dialog = new ProjectEditDialog();
                dialog.ProjectName = selectedProject.Name;
                dialog.FolderPath = selectedProject.FolderPath;
                dialog.Category = selectedProject.Category;
                dialog.IncludeSubfolders = selectedProject.IncludeSubfolders;
                dialog.UseCustomPdfPath = selectedProject.UseCustomPdfPath;
                dialog.CustomPdfPath = selectedProject.CustomPdfPath;

                if (dialog.ShowDialog() == true)
                {
                    selectedProject.Name = dialog.ProjectName;
                    selectedProject.FolderPath = dialog.FolderPath;
                    selectedProject.Category = dialog.Category;
                    selectedProject.IncludeSubfolders = dialog.IncludeSubfolders;
                    selectedProject.UseCustomPdfPath = dialog.UseCustomPdfPath;
                    selectedProject.CustomPdfPath = dialog.CustomPdfPath;
                    
                    // カテゴリが変更された場合、アイコンと色を更新
                    selectedProject.CategoryIcon = GetDefaultCategoryIcon(dialog.Category);
                    selectedProject.CategoryColor = GetCategoryColor(dialog.Category, "");

                    if (selectedProject == currentProject)
                    {
                        selectedFolderPath = selectedProject.FolderPath;
                        pdfOutputFolder = selectedProject.PdfOutputFolder;
                        txtFolderPath.Text = selectedFolderPath;
                        UpdateProjectDisplay();
                    }

                    SaveProjects();
                    RefreshProjectList();
                }
            }
            else
            {
                MessageBox.Show("編集するプロジェクトを選択してください。", "情報", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnDeleteProject_Click(object sender, RoutedEventArgs e)
        {
            if (treeProjects.SelectedItem is ProjectData selectedProject)
            {
                var result = MessageBox.Show($"プロジェクト '{selectedProject.Name}' を削除しますか？",
                    "確認", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    // カテゴリグループから削除
                    RemoveProjectFromCategoryGroup(selectedProject);

                    if (selectedProject == currentProject)
                    {
                        currentProject = null;
                        fileItems.Clear();
                        selectedFolderPath = "";
                        pdfOutputFolder = "";
                        txtFolderPath.Text = "";
                        UpdateProjectDisplay();
                    }

                    SaveProjects();
                }
            }
            else
            {
                MessageBox.Show("削除するプロジェクトを選択してください。", "情報", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnSwitchProject_Click(object sender, RoutedEventArgs e)
        {
            if (treeProjects.SelectedItem is ProjectData selectedProject)
            {
                SwitchToProject(selectedProject);
            }
            else
            {
                MessageBox.Show("プロジェクトを選択してください。", "情報", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void TreeProjects_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (treeProjects.SelectedItem is ProjectData selectedProject)
            {
                SwitchToProject(selectedProject);
            }
        }

        private void BtnConvertToProject_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFolderPath))
            {
                MessageBox.Show("先にフォルダを選択してください。", "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var folderName = Path.GetFileName(selectedFolderPath);
            var dialog = new ProjectEditDialog();
            dialog.ProjectName = folderName;
            dialog.FolderPath = selectedFolderPath;

            if (dialog.ShowDialog() == true)
            {
                var newProject = new ProjectData
                {
                    Name = dialog.ProjectName,
                    FolderPath = dialog.FolderPath,
                    Category = dialog.Category,
                    IncludeSubfolders = dialog.IncludeSubfolders,
                    UseCustomPdfPath = dialog.UseCustomPdfPath,
                    CustomPdfPath = dialog.CustomPdfPath,
                    MergeFileName = txtMergeFileName.Text,
                    AddPageNumber = chkAddPageNumber.IsChecked ?? false,
                    CategoryIcon = GetDefaultCategoryIcon(dialog.Category),
                    CategoryColor = GetCategoryColor(dialog.Category, "")
                };

                // 現在のファイル状態を保存
                foreach (var item in fileItems)
                {
                    newProject.FileItems.Add(new FileItemData
                    {
                        IsSelected = item.IsSelected,
                        TargetPages = item.TargetPages,
                        FilePath = item.FilePath,
                        LastModified = item.LastModified,
                        DisplayOrder = item.DisplayOrder,
                        RelativePath = item.RelativePath
                    });
                }

                AddProjectToCategoryGroup(newProject);
                SwitchToProject(newProject);

                MessageBox.Show($"プロジェクト '{newProject.Name}' を作成しました。", "完了",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void SwitchToProject(ProjectData project)
        {
            SaveCurrentProjectState();

            // 全プロジェクトのアクティブ状態をリセット
            foreach (var p in GetAllProjects())
            {
                p.IsActive = false;
            }

            // 新しいプロジェクトをアクティブに設定
            project.IsActive = true;
            currentProject = project;

            // UIを更新
            selectedFolderPath = project.FolderPath;
            pdfOutputFolder = project.PdfOutputFolder;
            txtFolderPath.Text = selectedFolderPath;
            txtMergeFileName.Text = project.MergeFileName;
            chkAddPageNumber.IsChecked = project.AddPageNumber;

            UpdateLatestMergedPdfDisplay();
            RestoreFileItems(project);
            UpdateProjectDisplay();
            SaveProjects();
        }

        private void SaveProjects()
        {
            var allProjects = GetAllProjects();
            ProjectManager.SaveProjects(allProjects);
        }

        private void SaveCurrentProjectState()
        {
            if (currentProject != null)
            {
                currentProject.FolderPath = selectedFolderPath;
                currentProject.PdfOutputFolder = pdfOutputFolder;
                currentProject.MergeFileName = txtMergeFileName.Text;
                currentProject.AddPageNumber = chkAddPageNumber.IsChecked ?? false;
                currentProject.LastAccessDate = DateTime.Now;

                // ファイルアイテムの状態を保存
                currentProject.FileItems.Clear();
                foreach (var item in fileItems)
                {
                    currentProject.FileItems.Add(new FileItemData
                    {
                        IsSelected = item.IsSelected,
                        TargetPages = item.TargetPages,
                        FilePath = item.FilePath,
                        LastModified = item.LastModified,
                        DisplayOrder = item.DisplayOrder,
                        RelativePath = item.RelativePath
                    });
                }

                SaveProjects();
            }
        }

        private void UpdateProjectDisplay()
        {
            if (currentProject != null)
            {
                lblCurrentProject.Content = $"現在のプロジェクト: {currentProject.Name}";
                Title = $"AllAutoOfficePDF2 - {currentProject.Name}";
            }
            else
            {
                lblCurrentProject.Content = "現在のプロジェクト: なし";
                Title = "AllAutoOfficePDF2";
            }
            
            // ヒントテキストの表示制御
            if (txtDropHint != null)
            {
                txtDropHint.Visibility = string.IsNullOrEmpty(selectedFolderPath) ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        /// <summary>
        /// プロジェクトリストをカテゴリ順序で再構築
        /// </summary>
        private void RefreshProjectList()
        {
            var currentSelectedProject = treeProjects.SelectedItem as ProjectData;
            
            // カテゴリグループをクリアして再構築
            LoadProjects();
            
            // 選択状態を復元
            if (currentSelectedProject != null)
            {
                SelectProjectInTree(currentSelectedProject.Id);
            }
        }

        /// <summary>
        /// TreeViewで指定IDのプロジェクトを選択
        /// </summary>
        private void SelectProjectInTree(string projectId)
        {
            foreach (var categoryGroup in categoryGroups)
            {
                var project = categoryGroup.Projects.FirstOrDefault(p => p.Id == projectId);
                if (project != null)
                {
                    // TreeViewItemを見つけて選択
                    var treeViewItem = FindTreeViewItem(treeProjects, project);
                    if (treeViewItem != null)
                    {
                        treeViewItem.IsSelected = true;
                    }
                    break;
                }
            }
        }

        /// <summary>
        /// TreeViewItemを検索
        /// </summary>
        private TreeViewItem FindTreeViewItem(System.Windows.Controls.TreeView treeView, object item)
        {
            return FindTreeViewItem(treeView, item, treeView.ItemContainerGenerator);
        }

        /// <summary>
        /// TreeViewItemを再帰的に検索
        /// </summary>
        private TreeViewItem FindTreeViewItem(ItemsControl parent, object item, ItemContainerGenerator generator)
        {
            if (parent == null || item == null) return null;

            for (int i = 0; i < parent.Items.Count; i++)
            {
                var container = generator.ContainerFromIndex(i) as TreeViewItem;
                if (container != null)
                {
                    if (container.DataContext == item)
                        return container;

                    var child = FindTreeViewItem(container, item, container.ItemContainerGenerator);
                    if (child != null)
                        return child;
                }
            }
            return null;
        }

        private void BtnCategoryManage_Click(object sender, RoutedEventArgs e)
        {
            var allProjects = GetAllProjects();
            var categories = ProjectManager.GetAvailableCategories(allProjects);
            var categoryList = string.Join("\n", categories.Select((c, i) => $"{i + 1}. {c}"));
            
            var message = "現在のカテゴリ一覧:\n\n";
            if (categories.Any())
            {
                message += categoryList;
            }
            else
            {
                message += "カテゴリはまだ設定されていません。";
            }
            
            message += "\n\nプロジェクト編集画面でカテゴリを設定・変更できます。";
            
            MessageBox.Show(message, "カテゴリ管理", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>
        /// プロジェクトをカテゴリグループに追加
        /// </summary>
        private void AddProjectToCategoryGroup(ProjectData project)
        {
            var categoryName = string.IsNullOrEmpty(project.Category) ? "未分類" : project.Category;
            var existingGroup = categoryGroups.FirstOrDefault(g => g.CategoryName == categoryName);
            
            if (existingGroup == null)
            {
                existingGroup = new ProjectCategoryGroup
                {
                    CategoryName = categoryName,
                    CategoryIcon = GetCategoryIcon(categoryName, project.CategoryIcon),
                    CategoryColor = GetCategoryColor(categoryName, project.CategoryColor)
                };
                categoryGroups.Add(existingGroup);
            }
            
            existingGroup.Projects.Add(project);
        }

        /// <summary>
        /// プロジェクトをカテゴリグループから削除
        /// </summary>
        private void RemoveProjectFromCategoryGroup(ProjectData project)
        {
            foreach (var categoryGroup in categoryGroups.ToList())
            {
                if (categoryGroup.Projects.Contains(project))
                {
                    categoryGroup.Projects.Remove(project);
                    
                    // プロジェクトが空になったカテゴリグループは削除
                    if (categoryGroup.Projects.Count == 0)
                    {
                        categoryGroups.Remove(categoryGroup);
                    }
                    break;
                }
            }
        }
        #endregion

        #region ファイル管理
        private void BtnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "対象フォルダを選択してください（フォルダパスのみが設定されます）";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    // フォルダパスのみを設定（ファイル名は含めない）
                    selectedFolderPath = dialog.SelectedPath;
                    txtFolderPath.Text = selectedFolderPath;
                    
                    // PDFアウトプットフォルダもフォルダパスのみに設定
                    if (currentProject != null && currentProject.UseCustomPdfPath && !string.IsNullOrEmpty(currentProject.CustomPdfPath))
                    {
                        pdfOutputFolder = currentProject.CustomPdfPath;
                    }
                    else
                    {
                        pdfOutputFolder = Path.Combine(selectedFolderPath, "PDF");
                    }
                    
                    if (currentProject != null)
                    {
                        currentProject.FolderPath = selectedFolderPath;
                        // PdfOutputFolderはプロパティで自動計算されるので直接設定しない
                        SaveProjects();
                    }
                    
                    txtStatus.Text = "フォルダが選択されました";
                }
            }
        }

        private void BtnReadFolder_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFolderPath))
            {
                MessageBox.Show("フォルダを選択してください", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var includeSubfolders = currentProject?.IncludeSubfolders ?? false;
            var loadedFileItems = FileManagementService.LoadFilesFromFolder(selectedFolderPath, pdfOutputFolder, includeSubfolders);
            
            fileItems.Clear();
            for (int i = 0; i < loadedFileItems.Count; i++)
            {
                loadedFileItems[i].Number = i + 1;
                loadedFileItems[i].DisplayOrder = i;
                fileItems.Add(loadedFileItems[i]);
            }

            var statusMessage = $"{fileItems.Count}個のファイルを読み込みました";
            if (includeSubfolders)
            {
                statusMessage += " (サブフォルダを含む)";
            }
            txtStatus.Text = statusMessage;
            SaveCurrentProjectState();
        }

        private void BtnUpdateFiles_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFolderPath))
            {
                MessageBox.Show("フォルダを選択してください", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var includeSubfolders = currentProject?.IncludeSubfolders ?? false;
            var (updatedItems, changedFiles, addedFiles, deletedFiles) = 
                FileManagementService.UpdateFiles(selectedFolderPath, pdfOutputFolder, fileItems.ToList(), includeSubfolders);

            fileItems.Clear();
            foreach (var item in updatedItems)
            {
                fileItems.Add(item);
            }

            // 結果メッセージを作成
            var statusMessages = new List<string>();
            statusMessages.Add($"{fileItems.Count}個のファイルを更新しました");
            
            if (includeSubfolders)
            {
                statusMessages.Add("(サブフォルダを含む)");
            }

            if (changedFiles.Any())
                statusMessages.Add($"変更されたファイル: {changedFiles.Count}個");

            if (addedFiles.Any())
                statusMessages.Add($"追加されたファイル: {addedFiles.Count}個");

            if (deletedFiles.Any())
            {
                statusMessages.Add($"削除されたファイル: {deletedFiles.Count}個");
                
                var deletedMessage = $"以下のファイルが削除されました：\n{string.Join("\n", deletedFiles)}";
                if (deletedFiles.Any(f => !f.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)))
                {
                    deletedMessage += "\n\n対応するPDFファイルも削除されました。";
                }
                MessageBox.Show(deletedMessage, "削除されたファイル", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            txtStatus.Text = string.Join(" / ", statusMessages);
            SaveCurrentProjectState();
        }

        private void RestoreFileItems(ProjectData project)
        {
            fileItems.Clear();

            if (string.IsNullOrEmpty(project.FolderPath) || !Directory.Exists(project.FolderPath))
            {
                txtStatus.Text = "プロジェクトフォルダが見つかりません";
                return;
            }

            var loadedFileItems = FileManagementService.LoadFilesFromFolder(project.FolderPath, project.PdfOutputFolder, project.IncludeSubfolders);
            
            // 保存された状態を復元
            foreach (var item in loadedFileItems)
            {
                var savedItem = project.FileItems.FirstOrDefault(f => f.FilePath == item.FilePath);
                if (savedItem != null)
                {
                    item.IsSelected = savedItem.IsSelected;
                    item.TargetPages = savedItem.TargetPages;
                    item.DisplayOrder = savedItem.DisplayOrder >= 0 ? savedItem.DisplayOrder : loadedFileItems.Count;
                }
            }

            // 表示順序で並び替え
            var orderedItems = loadedFileItems
                .OrderBy(f => f.DisplayOrder)
                .ThenBy(f => f.RelativePath)
                .ThenBy(f => f.FileName)
                .ToList();

            for (int i = 0; i < orderedItems.Count; i++)
            {
                orderedItems[i].Number = i + 1;
                orderedItems[i].DisplayOrder = i;
                fileItems.Add(orderedItems[i]);
            }

            var statusMessage = $"プロジェクト '{project.Name}' を読み込みました ({fileItems.Count}個のファイル)";
            if (project.IncludeSubfolders)
            {
                statusMessage += " (サブフォルダを含む)";
            }
            txtStatus.Text = statusMessage;
        }
        #endregion

        #region ファイル操作
        private void BtnMoveUp_Click(object sender, RoutedEventArgs e)
        {
            if (dgFiles.SelectedIndex > 0)
            {
                var selectedIndex = dgFiles.SelectedIndex;
                var item = fileItems[selectedIndex];
                fileItems.RemoveAt(selectedIndex);
                fileItems.Insert(selectedIndex - 1, item);
                
                UpdateFileNumbers();
                dgFiles.SelectedIndex = selectedIndex - 1;
                SaveCurrentProjectState();
            }
        }

        private void BtnMoveDown_Click(object sender, RoutedEventArgs e)
        {
            if (dgFiles.SelectedIndex >= 0 && dgFiles.SelectedIndex < fileItems.Count - 1)
            {
                var selectedIndex = dgFiles.SelectedIndex;
                var item = fileItems[selectedIndex];
                fileItems.RemoveAt(selectedIndex);
                fileItems.Insert(selectedIndex + 1, item);
                
                UpdateFileNumbers();
                dgFiles.SelectedIndex = selectedIndex + 1;
                SaveCurrentProjectState();
            }
        }

        private void BtnSortByName_Click(object sender, RoutedEventArgs e)
        {
            var sortedItems = fileItems.OrderBy(f => f.FileName).ToList();
            fileItems.Clear();
            
            for (int i = 0; i < sortedItems.Count; i++)
            {
                sortedItems[i].Number = i + 1;
                sortedItems[i].DisplayOrder = i;
                fileItems.Add(sortedItems[i]);
            }
            
            SaveCurrentProjectState();
            txtStatus.Text = "ファイル名順に並び替えました";
        }

        private void UpdateFileNumbers()
        {
            for (int i = 0; i < fileItems.Count; i++)
            {
                fileItems[i].Number = i + 1;
                fileItems[i].DisplayOrder = i;
            }
        }

        private void ChkSelectAll_Click(object sender, RoutedEventArgs e)
        {
            var isChecked = chkSelectAll.IsChecked ?? false;
            foreach (var item in fileItems)
            {
                item.IsSelected = isChecked;
            }
            SaveCurrentProjectState();
        }

        private void FileName_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBlock textBlock && textBlock.DataContext is FileItem fileItem)
            {
                OpenFile(fileItem.FilePath);
            }
        }

        private void OpenFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = filePath,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"ファイルを開けませんでした: {ex.Message}", "エラー",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("ファイルが見つかりません。", "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        #endregion

        #region PDF処理
        private async void BtnConvertPDF_Click(object sender, RoutedEventArgs e)
        {
            var selectedFiles = fileItems.Where(f => f.IsSelected).ToList();
            if (!selectedFiles.Any())
            {
                MessageBox.Show("変換するファイルを選択してください", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!Directory.Exists(pdfOutputFolder))
                Directory.CreateDirectory(pdfOutputFolder);

            var includeSubfolders = currentProject?.IncludeSubfolders ?? false;
            var baseFolderPath = selectedFolderPath;

            progressBar.Visibility = Visibility.Visible;
            progressBar.Maximum = selectedFiles.Count;
            progressBar.Value = 0;

            await Task.Run(() =>
            {
                foreach (var file in selectedFiles)
                {
                    try
                    {
                        // サブフォルダ構造を考慮した変換
                        if (includeSubfolders)
                        {
                            FileManagementService.EnsurePdfOutputDirectory(file.FilePath, pdfOutputFolder, baseFolderPath, includeSubfolders);
                        }

                        PdfConversionService.ConvertToPdf(file.FilePath, pdfOutputFolder, file.TargetPages, baseFolderPath, includeSubfolders);

                        Dispatcher.Invoke(() =>
                        {
                            file.PdfStatus = "変換済";
                            file.IsSelected = false;
                            progressBar.Value++;
                            txtStatus.Text = $"変換中: {file.FileName}";
                        });
                    }
                    catch (Exception ex)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            MessageBox.Show($"変換エラー: {file.FileName}\n{ex.Message}", "エラー",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                        });
                    }
                }
            });

            progressBar.Visibility = Visibility.Collapsed;
            txtStatus.Text = "PDF変換が完了しました";
        }

        private async void BtnMergePDF_Click(object sender, RoutedEventArgs e)
        {
            var allFiles = fileItems.OrderBy(f => f.DisplayOrder).ToList();
            if (!allFiles.Any())
            {
                MessageBox.Show("結合対象のファイルがありません", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // PDFファイルパスを取得
            var pdfFilePaths = new List<string>();
            var missingPdfFiles = new List<string>();
            var includeSubfolders = currentProject?.IncludeSubfolders ?? false;
            var baseFolderPath = selectedFolderPath;

            foreach (var file in allFiles)
            {
                string pdfPath;
                if (file.Extension.ToLower() == "pdf")
                {
                    pdfPath = file.FilePath;
                }
                else
                {
                    if (includeSubfolders)
                    {
                        // サブフォルダ構造を考慮したパス
                        var fileInfo = new FileInfo(file.FilePath);
                        var relativePath = GetRelativePath(baseFolderPath, fileInfo.DirectoryName!);
                        var outputDir = Path.Combine(pdfOutputFolder, relativePath);
                        pdfPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(file.FileName) + ".pdf");
                    }
                    else
                    {
                        pdfPath = Path.Combine(pdfOutputFolder, Path.GetFileNameWithoutExtension(file.FileName) + ".pdf");
                    }
                }

                if (File.Exists(pdfPath))
                {
                    pdfFilePaths.Add(pdfPath);
                }
                else
                {
                    missingPdfFiles.Add(file.FileName);
                }
            }

            if (missingPdfFiles.Any())
            {
                var message = "以下のファイルに対応するPDFファイルが見つかりません:\n\n";
                message += string.Join("\n", missingPdfFiles);
                message += "\n\n先にPDF変換を実行してください。";
                
                MessageBox.Show(message, "PDFファイル不足", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // mergePDFフォルダの場所を決定（カスタムPDF保存パスを考慮）
            var mergeFolder = GetMergePdfFolderPath();
            
            if (!Directory.Exists(mergeFolder))
                Directory.CreateDirectory(mergeFolder);

            var mergeFileName = txtMergeFileName.Text;
            var addPageNumber = chkAddPageNumber.IsChecked == true;
            var timestamp = DateTime.Now.ToString("yyMMddHHmmss");
            var outputFileName = $"{mergeFileName}_{timestamp}.pdf";
            var outputPath = Path.Combine(mergeFolder, outputFileName);

            progressBar.Visibility = Visibility.Visible;
            progressBar.IsIndeterminate = true;
            txtStatus.Text = "PDF結合中...";

            bool mergeSuccess = false;
            
            await Task.Run(() =>
            {
                try
                {
                    PdfMergeService.MergePdfFiles(pdfFilePaths, outputPath, addPageNumber);
                    mergeSuccess = true;
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        MessageBox.Show($"PDF結合エラー: {ex.Message}", "エラー",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    });
                }
            });

            progressBar.Visibility = Visibility.Collapsed;
            
            if (mergeSuccess)
            {
                if (currentProject != null)
                {
                    currentProject.LatestMergedPdfPath = outputPath;
                    SaveProjects();
                }

                UpdateLatestMergedPdfDisplay();
                txtStatus.Text = "PDF結合が完了しました";
                
                var result = MessageBox.Show("PDFを開きますか？", "PDF結合完了", 
                    MessageBoxButton.YesNo, MessageBoxImage.Question);
                
                if (result == MessageBoxResult.Yes)
                {
                    OpenFile(outputPath);
                }
                else
                {
                    try
                    {
                        Process.Start("explorer.exe", mergeFolder);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"フォルダを開けませんでした: {ex.Message}", "エラー",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private void UpdateLatestMergedPdfDisplay()
        {
            if (currentProject != null && !string.IsNullOrEmpty(currentProject.LatestMergedPdfPath))
            {
                if (File.Exists(currentProject.LatestMergedPdfPath))
                {
                    txtLatestMergedPDF.Text = Path.GetFileName(currentProject.LatestMergedPdfPath);
                    btnOpenLatestMergedPDF.IsEnabled = true;
                }
                else
                {
                    txtLatestMergedPDF.Text = "ファイルが見つかりません";
                    btnOpenLatestMergedPDF.IsEnabled = false;
                }
            }
            else
            {
                txtLatestMergedPDF.Text = "まだ結合されていません";
                btnOpenLatestMergedPDF.IsEnabled = false;
            }
        }

        private void BtnOpenLatestMergedPDF_Click(object sender, RoutedEventArgs e)
        {
            if (currentProject != null && !string.IsNullOrEmpty(currentProject.LatestMergedPdfPath))
            {
                if (File.Exists(currentProject.LatestMergedPdfPath))
                {
                    OpenFile(currentProject.LatestMergedPdfPath);
                }
                else
                {
                    MessageBox.Show("結合PDFファイルが見つかりません。", "エラー",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    UpdateLatestMergedPdfDisplay();
                }
            }
            else
            {
                MessageBox.Show("結合PDFファイルがありません。", "情報",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        #endregion

        #region フォルダ操作
        private void BtnOpenProjectFolder_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button button && button.Tag is ProjectData project)
            {
                if (Directory.Exists(project.FolderPath))
                {
                    try
                    {
                        Process.Start("explorer.exe", project.FolderPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"フォルダを開けませんでした: {ex.Message}", "エラー",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    MessageBox.Show("フォルダが見つかりません。", "エラー",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        private void BtnOpenCurrentProjectFolder_Click(object sender, RoutedEventArgs e)
        {
            if (currentProject != null && Directory.Exists(currentProject.FolderPath))
            {
                try
                {
                    Process.Start("explorer.exe", currentProject.FolderPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"フォルダを開けませんでした: {ex.Message}", "エラー",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else if (currentProject == null)
            {
                MessageBox.Show("現在のプロジェクトが選択されていません。", "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                MessageBox.Show("フォルダが見つかりません。", "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        #endregion

        #region イベントハンドラ
        private void TreeProjects_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is ProjectData selectedProject)
            {
                // プロジェクトが選択された場合の処理はここに追加
                // 現在は何もしない（ダブルクリックで切り替え）
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            SaveCurrentProjectState();
            base.OnClosed(e);
        }
        #endregion

        #region ドラッグ&ドロップ処理
        private void Window_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void Window_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string droppedPath = files[0];
                    
                    // フォルダかファイルかを判定
                    if (Directory.Exists(droppedPath))
                    {
                        SetFolderPath(droppedPath);
                    }
                    else if (File.Exists(droppedPath))
                    {
                        // ファイルの場合は親フォルダを使用
                        string parentFolder = Path.GetDirectoryName(droppedPath);
                        if (!string.IsNullOrEmpty(parentFolder))
                        {
                            SetFolderPath(parentFolder);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 新しいドラッグ&ドロップエリアのDragEnter
        /// </summary>
        private void DropArea_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                // ドラッグオーバー時の視覚的フィードバック
                if (sender is Border border)
                {
                    border.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.LightBlue);
                    border.BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.DodgerBlue);
                }
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        /// <summary>
        /// 新しいドラッグ&ドロップエリアのDragOver
        /// </summary>
        private void DropArea_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        /// <summary>
        /// 新しいドラッグ&ドロップエリアのDragLeave
        /// </summary>
        private void DropArea_DragLeave(object sender, DragEventArgs e)
        {
            // ドラッグリーブ時の視覚的フィードバックを元に戻す
            if (sender is Border border)
            {
                border.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(248, 249, 250));
                border.BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 122, 204));
            }
        }

        /// <summary>
        /// 新しいドラッグ&ドロップエリアのDrop
        /// </summary>
        private void DropArea_Drop(object sender, DragEventArgs e)
        {
            // ドラッグオーバー時の視覚的フィードバックを元に戻す
            if (sender is Border border)
            {
                border.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(248, 249, 250));
                border.BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 122, 204));
            }
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string droppedPath = files[0];
                    
                    // フォルダかファイルかを判定
                    if (Directory.Exists(droppedPath))
                    {
                        SetFolderPath(droppedPath);
                    }
                    else if (File.Exists(droppedPath))
                    {
                        // ファイルの場合は親フォルダを使用
                        string parentFolder = Path.GetDirectoryName(droppedPath);
                        if (!string.IsNullOrEmpty(parentFolder))
                        {
                            SetFolderPath(parentFolder);
                        }
                    }
                }
            }
        }

        private void TxtFolderPath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                // ドラッグオーバー時の視覚的フィードバック
                txtFolderPath.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.LightBlue);
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void TxtFolderPath_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void TxtFolderPath_DragLeave(object sender, DragEventArgs e)
        {
            // ドラッグリーブ時の視覚的フィードバックを元に戻す
            txtFolderPath.Background = System.Windows.Media.Brushes.White;
        }

        private void TxtFolderPath_Drop(object sender, DragEventArgs e)
        {
            // ドラッグオーバー時の視覚的フィードバックを元に戻す
            txtFolderPath.Background = System.Windows.Media.Brushes.White;
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string droppedPath = files[0];
                    
                    // フォルダかファイルかを判定
                    if (Directory.Exists(droppedPath))
                    {
                        SetFolderPath(droppedPath);
                    }
                    else if (File.Exists(droppedPath))
                    {
                        // ファイルの場合は親フォルダを使用
                        string parentFolder = Path.GetDirectoryName(droppedPath);
                        if (!string.IsNullOrEmpty(parentFolder))
                        {
                            SetFolderPath(parentFolder);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// フォルダパスを設定する共通メソッド
        /// </summary>
        /// <param name="folderPath">設定するフォルダパス</param>
        private void SetFolderPath(string folderPath)
        {
            // フォルダパスのみを設定（ファイル名は含めない）
            selectedFolderPath = folderPath;
            txtFolderPath.Text = selectedFolderPath;
            
            // ヒントテキストを非表示にする
            if (txtDropHint != null)
            {
                txtDropHint.Visibility = string.IsNullOrEmpty(selectedFolderPath) ? Visibility.Visible : Visibility.Collapsed;
            }
            
            // PDFアウトプットフォルダもフォルダパスのみに設定
            if (currentProject != null && currentProject.UseCustomPdfPath && !string.IsNullOrEmpty(currentProject.CustomPdfPath))
            {
                pdfOutputFolder = currentProject.CustomPdfPath;
            }
            else
            {
                pdfOutputFolder = Path.Combine(selectedFolderPath, "PDF");
            }
            
            if (currentProject != null)
            {
                currentProject.FolderPath = selectedFolderPath;
                // PdfOutputFolderはプロパティで自動計算されるので直接設定しない
                SaveProjects();
            }
            
            txtStatus.Text = "フォルダがドラッグ&ドロップで選択されました";
        }
        #endregion

        #region ヘルパーメソッド
        /// <summary>
        /// 相対パスを取得
        /// </summary>
        /// <param name="basePath">基準パス</param>
        /// <param name="fullPath">完全パス</param>
        /// <returns>相対パス</returns>
        private string GetRelativePath(string basePath, string fullPath)
        {
            var baseUri = new Uri(basePath.EndsWith(Path.DirectorySeparatorChar.ToString()) ? basePath : basePath + Path.DirectorySeparatorChar);
            var fullUri = new Uri(fullPath);
            
            if (baseUri.Scheme != fullUri.Scheme)
            {
                return fullPath;
            }

            var relativeUri = baseUri.MakeRelativeUri(fullUri);
            var relativePath = Uri.UnescapeDataString(relativeUri.ToString());
            
            return relativePath.Replace('/', Path.DirectorySeparatorChar);
        }

        /// <summary>
        /// mergePDFフォルダのパスを取得
        /// </summary>
        /// <returns>mergePDFフォルダのパス</returns>
        private string GetMergePdfFolderPath()
        {
            if (currentProject != null)
            {
                return currentProject.MergePdfFolder;
            }
            else
            {
                // プロジェクトがない場合は従来通り
                return Path.Combine(selectedFolderPath, "mergePDF");
            }
        }
        #endregion
    }
}