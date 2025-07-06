using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Application = System.Windows.Application;
using MessageBox = System.Windows.MessageBox;
using WordInterop = Microsoft.Office.Interop.Word;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace AllAutoOfficePDF2
{
    // データモデル
    public class FileItem : INotifyPropertyChanged
    {
        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged(nameof(IsSelected));
            }
        }

        private string _targetPages = "";
        public string TargetPages
        {
            get => _targetPages;
            set
            {
                _targetPages = value;
                OnPropertyChanged(nameof(TargetPages));
            }
        }

        public int Number { get; set; }
        public string FileName { get; set; } = "";
        public string FilePath { get; set; } = "";
        public string Extension { get; set; } = "";
        public DateTime LastModified { get; set; }
        public string PdfStatus { get; set; } = "";

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    // メインウィンドウ
    public partial class MainWindow : System.Windows.Window
    {
        private ObservableCollection<FileItem> fileItems = new ObservableCollection<FileItem>();
        private string selectedFolderPath = "";
        private string pdfOutputFolder = "";

        public MainWindow()
        {
            InitializeComponent();
            dgFiles.ItemsSource = fileItems;
        }

        // フォルダ選択
        private void BtnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "対象フォルダを選択してください";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    selectedFolderPath = dialog.SelectedPath;
                    txtFolderPath.Text = selectedFolderPath;
                    pdfOutputFolder = Path.Combine(selectedFolderPath, "PDF");
                    txtStatus.Text = "フォルダが選択されました";
                }
            }
        }

        // ファイル読込
        private void BtnReadFolder_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFolderPath))
            {
                System.Windows.MessageBox.Show("フォルダを選択してください", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            fileItems.Clear();

            var extensions = new[] { "*.xls", "*.xlsx", "*.xlsm", "*.doc", "*.docx", "*.ppt", "*.pptx", "*.pdf" };
            int number = 1;

            foreach (var ext in extensions)
            {
                var files = Directory.GetFiles(selectedFolderPath, ext);
                foreach (var file in files.OrderBy(f => f))
                {
                    var fileInfo = new FileInfo(file);
                    string extensionUpper = fileInfo.Extension.TrimStart('.').ToUpper();
                    var item = new FileItem
                    {
                        Number = number++,
                        FileName = fileInfo.Name,
                        FilePath = fileInfo.FullName,
                        Extension = extensionUpper,
                        LastModified = fileInfo.LastWriteTime,
                        IsSelected = true,
                        PdfStatus = CheckPdfExists(fileInfo) ? "変換済" : "未変換",
                        TargetPages = (extensionUpper == "XLS" || extensionUpper == "XLSX" || extensionUpper == "XLSM") ? "1-1" : ""
                    };
                    fileItems.Add(item);
                }
            }

            txtStatus.Text = $"{fileItems.Count}個のファイルを読み込みました";
        }

        // PDFの存在確認
        private bool CheckPdfExists(FileInfo fileInfo)
        {
            if (fileInfo.Extension.ToLower() == ".pdf") return true;

            var pdfPath = Path.Combine(pdfOutputFolder, Path.GetFileNameWithoutExtension(fileInfo.Name) + ".pdf");
            return File.Exists(pdfPath);
        }

        // ファイル更新
        private void BtnUpdateFiles_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFolderPath))
            {
                System.Windows.MessageBox.Show("フォルダを選択してください", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 現在のファイル一覧を保存（変更・削除検出用）
            var previousFiles = fileItems.ToDictionary(f => f.FilePath, f => f);

            // 新しいファイル一覧を取得
            var newFileItems = new List<FileItem>();
            var extensions = new[] { "*.xls", "*.xlsx", "*.xlsm", "*.doc", "*.docx", "*.ppt", "*.pptx", "*.pdf" };
            int number = 1;
            var changedFiles = new List<string>();
            var addedFiles = new List<string>();

            foreach (var ext in extensions)
            {
                var files = Directory.GetFiles(selectedFolderPath, ext);
                foreach (var file in files.OrderBy(f => f))
                {
                    var fileInfo = new FileInfo(file);
                    string extensionUpper = fileInfo.Extension.TrimStart('.').ToUpper();

                    bool isSelected = true; // デフォルトで選択
                    string targetPages = (extensionUpper == "XLS" || extensionUpper == "XLSX" || extensionUpper == "XLSM") ? "1-1" : "";

                    // 既存ファイルの場合は更新日時をチェック
                    if (previousFiles.TryGetValue(file, out var existingFile))
                    {
                        if (existingFile.LastModified != fileInfo.LastWriteTime)
                        {
                            // 更新日時が変更された場合
                            changedFiles.Add(fileInfo.Name);
                            isSelected = true;
                        }
                        else
                        {
                            // 変更されていない場合は前の選択状態を保持
                            isSelected = existingFile.IsSelected;
                            targetPages = existingFile.TargetPages;
                        }
                    }
                    else
                    {
                        // 新規ファイルの場合
                        addedFiles.Add(fileInfo.Name);
                        isSelected = true;
                    }

                    var item = new FileItem
                    {
                        Number = number++,
                        FileName = fileInfo.Name,
                        FilePath = fileInfo.FullName,
                        Extension = extensionUpper,
                        LastModified = fileInfo.LastWriteTime,
                        IsSelected = isSelected,
                        PdfStatus = CheckPdfExists(fileInfo) ? "変換済" : "未変換",
                        TargetPages = targetPages
                    };
                    newFileItems.Add(item);
                }
            }

            // 削除されたファイルを検出してPDFファイルを削除
            var deletedFiles = new List<string>();
            var currentFilePaths = newFileItems.Select(f => f.FilePath).ToHashSet();

            foreach (var previousFile in previousFiles.Values)
            {
                if (!currentFilePaths.Contains(previousFile.FilePath))
                {
                    deletedFiles.Add(previousFile.FileName);

                    // 対応するPDFファイルを削除
                    if (previousFile.Extension.ToLower() != "pdf")
                    {
                        var pdfPath = Path.Combine(pdfOutputFolder, Path.GetFileNameWithoutExtension(previousFile.FileName) + ".pdf");
                        if (File.Exists(pdfPath))
                        {
                            try
                            {
                                File.Delete(pdfPath);
                            }
                            catch (Exception ex)
                            {
                                System.Windows.MessageBox.Show($"PDFファイルの削除に失敗しました: {Path.GetFileName(pdfPath)}\n{ex.Message}",
                                    "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                            }
                        }
                    }
                }
            }

            // ファイル一覧を更新
            fileItems.Clear();
            foreach (var item in newFileItems)
            {
                fileItems.Add(item);
            }

            // 結果メッセージを作成
            var statusMessages = new List<string>();
            statusMessages.Add($"{fileItems.Count}個のファイルを更新しました");

            if (changedFiles.Any())
            {
                statusMessages.Add($"変更されたファイル: {changedFiles.Count}個");
            }

            if (addedFiles.Any())
            {
                statusMessages.Add($"追加されたファイル: {addedFiles.Count}個");
            }

            if (deletedFiles.Any())
            {
                statusMessages.Add($"削除されたファイル: {deletedFiles.Count}個");

                // 削除されたファイルの詳細をメッセージボックスで表示
                var deletedMessage = $"以下のファイルが削除されました：\n{string.Join("\n", deletedFiles)}";
                if (deletedFiles.Any(f => !f.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)))
                {
                    deletedMessage += "\n\n対応するPDFファイルも削除されました。";
                }

                System.Windows.MessageBox.Show(deletedMessage, "削除されたファイル", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            txtStatus.Text = string.Join(" / ", statusMessages);
        }

        // PDF変換
        private async void BtnConvertPDF_Click(object sender, RoutedEventArgs e)
        {
            // 選択されているファイルはすべて変換対象にする
            var selectedFiles = fileItems
                .Where(f => f.IsSelected)
                .ToList();
            if (!selectedFiles.Any())
            {
                System.Windows.MessageBox.Show("変換するファイルを選択してください", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // PDFフォルダ作成
            if (!Directory.Exists(pdfOutputFolder))
                Directory.CreateDirectory(pdfOutputFolder);

            progressBar.Visibility = Visibility.Visible;
            progressBar.Maximum = selectedFiles.Count;
            progressBar.Value = 0;

            await System.Threading.Tasks.Task.Run(() =>
            {
                foreach (var file in selectedFiles)
                {
                    try
                    {
                        ConvertToPdf(file.FilePath);

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
                            System.Windows.MessageBox.Show($"変換エラー: {file.FileName}\n{ex.Message}", "エラー",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                        });
                    }
                }
            });

            progressBar.Visibility = Visibility.Collapsed;
            txtStatus.Text = "PDF変換が完了しました";
        }

        // Office文書をPDFに変換
        private void ConvertToPdf(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLower();
            var outputPath = Path.Combine(pdfOutputFolder, Path.GetFileNameWithoutExtension(filePath) + ".pdf");

            // 対象ページを取得
            var fileItem = fileItems.FirstOrDefault(f => f.FilePath == filePath);
            var targetPages = fileItem?.TargetPages ?? "";

            switch (extension)
            {
                case ".xls":
                case ".xlsx":
                case ".xlsm":
                    ConvertExcelToPdf(filePath, outputPath, targetPages);
                    break;
                case ".doc":
                case ".docx":
                    ConvertWordToPdf(filePath, outputPath, targetPages);
                    break;
                case ".ppt":
                case ".pptx":
                    ConvertPowerPointToPdf(filePath, outputPath, targetPages);
                    break;
                case ".pdf":
                    // PDFファイルは対象外
                    if (!File.Exists(outputPath))
                        File.Copy(filePath, outputPath, overwrite: false);
                    break;
            }
        }

        // ページ範囲解析
        private List<int> ParsePageRange(string pageRange)
        {
            var pages = new List<int>();
            if (string.IsNullOrWhiteSpace(pageRange))
                return pages;

            try
            {
                var parts = pageRange.Split(',');
                foreach (var part in parts)
                {
                    var trimmed = part.Trim();
                    if (trimmed.Contains('-'))
                    {
                        var range = trimmed.Split('-');
                        if (range.Length == 2 && int.TryParse(range[0], out int start) && int.TryParse(range[1], out int end))
                        {
                            for (int i = start; i <= end; i++)
                            {
                                if (i > 0 && !pages.Contains(i))
                                    pages.Add(i);
                            }
                        }
                    }
                    else if (int.TryParse(trimmed, out int page))
                    {
                        if (page > 0 && !pages.Contains(page))
                            pages.Add(page);
                    }
                }
            }
            catch
            {
                throw new ArgumentException($"無効なページ範囲指定: {pageRange}");
            }

            return pages.OrderBy(p => p).ToList();
        }

        // Excel→PDF変換
        private void ConvertExcelToPdf(string inputPath, string outputPath, string targetPages = "")
        {
            ExcelInterop.Application? excelApp = null;
            ExcelInterop.Workbook? workbook = null;

            try
            {
                excelApp = new ExcelInterop.Application();
                excelApp.Visible = false;
                workbook = excelApp.Workbooks.Open(inputPath);

                var targetSheets = ParsePageRange(targetPages);

                if (targetSheets.Any())
                {
                    // 指定シートのみ変換
                    var totalSheets = workbook.Worksheets.Count;

                    // 存在しないシートのチェック
                    var invalidSheets = targetSheets.Where(s => s > totalSheets).ToList();
                    if (invalidSheets.Any())
                    {
                        throw new ArgumentException($"存在しないシート番号が指定されています: {string.Join(", ", invalidSheets)} (総シート数: {totalSheets})");
                    }

                    // 指定シートを選択
                    var sheetsToSelect = new ExcelInterop.Worksheet[targetSheets.Count];
                    for (int i = 0; i < targetSheets.Count; i++)
                    {
                        sheetsToSelect[i] = (ExcelInterop.Worksheet)workbook.Worksheets[targetSheets[i]];
                    }

                    // 複数シートを選択
                    if (sheetsToSelect.Length > 1)
                    {
                        sheetsToSelect[0].Select();
                        for (int i = 1; i < sheetsToSelect.Length; i++)
                        {
                            sheetsToSelect[i].Select(false); // 追加選択
                        }
                    }
                    else
                    {
                        sheetsToSelect[0].Select();
                    }

                    // 選択されたシートをPDF変換
                    ((ExcelInterop.Worksheet)excelApp.ActiveSheet).ExportAsFixedFormat(ExcelInterop.XlFixedFormatType.xlTypePDF, outputPath);
                }
                else
                {
                    // 全シート変換
                    workbook.ExportAsFixedFormat(ExcelInterop.XlFixedFormatType.xlTypePDF, outputPath);
                }
            }
            finally
            {
                workbook?.Close(false);
                excelApp?.Quit();
                if (workbook != null) ReleaseComObject(workbook);
                if (excelApp != null) ReleaseComObject(excelApp);
            }
        }

        // Word→PDF変換
        private void ConvertWordToPdf(string inputPath, string outputPath, string targetPages = "")
        {
            WordInterop.Application? wordApp = null;
            WordInterop.Document? document = null;

            try
            {
                wordApp = new WordInterop.Application();
                wordApp.Visible = false;
                document = wordApp.Documents.Open(inputPath);

                var targetPageList = ParsePageRange(targetPages);

                if (targetPageList.Any())
                {
                    // 指定ページのみ変換
                    var totalPages = document.ComputeStatistics(WordInterop.WdStatistic.wdStatisticPages);

                    // 存在しないページのチェック
                    var invalidPages = targetPageList.Where(p => p > totalPages).ToList();
                    if (invalidPages.Any())
                    {
                        throw new ArgumentException($"存在しないページ番号が指定されています: {string.Join(", ", invalidPages)} (総ページ数: {totalPages})");
                    }

                    // ページ範囲指定でPDF出力
                    document.ExportAsFixedFormat(outputPath, WordInterop.WdExportFormat.wdExportFormatPDF,
                        Range: WordInterop.WdExportRange.wdExportFromTo,
                        From: targetPageList.Min(),
                        To: targetPageList.Max());
                }
                else
                {
                    // 全ページ変換
                    document.ExportAsFixedFormat(outputPath, WordInterop.WdExportFormat.wdExportFormatPDF);
                }
            }
            finally
            {
                document?.Close(false);
                wordApp?.Quit();
                if (document != null) ReleaseComObject(document);
                if (wordApp != null) ReleaseComObject(wordApp);
            }
        }

        // PowerPoint→PDF変換
        private void ConvertPowerPointToPdf(string inputPath, string outputPath, string targetPages = "")
        {
            dynamic? pptApp = null;
            dynamic? presentation = null;
            dynamic? tempPresentation = null;

            try
            {
                pptApp = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));
                presentation = pptApp.Presentations.Open(inputPath);

                var targetSlides = ParsePageRange(targetPages);

                if (targetSlides.Any())
                {
                    // 指定スライドのみ変換
                    var totalSlides = presentation.Slides.Count;

                    // 存在しないスライドのチェック
                    var invalidSlides = targetSlides.Where(s => s > totalSlides).ToList();
                    if (invalidSlides.Any())
                    {
                        throw new ArgumentException($"存在しないスライド番号が指定されています: {string.Join(", ", invalidSlides)} (総スライド数: {totalSlides})");
                    }

                    // 新しいプレゼンテーションを作成
                    tempPresentation = pptApp.Presentations.Add();

                    // 指定スライドをコピー
                    foreach (var slideIndex in targetSlides.OrderBy(x => x))
                    {
                        var sourceSlide = presentation.Slides[slideIndex];
                        sourceSlide.Copy();
                        tempPresentation.Slides.Paste();
                    }

                    // 最初の空白スライドを削除（存在する場合）
                    if (tempPresentation.Slides.Count > targetSlides.Count)
                    {
                        tempPresentation.Slides[1].Delete();
                    }

                    // 一時プレゼンテーションをPDF保存
                    tempPresentation.SaveAs(outputPath, 32); // 32 = ppSaveAsPDF
                }
                else
                {
                    // 全スライド変換
                    presentation.SaveAs(outputPath, 32); // 32 = ppSaveAsPDF
                }

                presentation.Close();
                tempPresentation?.Close();
            }
            finally
            {
                try { pptApp?.Quit(); } catch { }
                if (tempPresentation != null) ReleaseComObject(tempPresentation);
                if (presentation != null) ReleaseComObject(presentation);
                if (pptApp != null) ReleaseComObject(pptApp);

                // 念のためGCを2回
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();

                // PowerPointプロセスが残っていれば強制終了（最終手段）
                foreach (var proc in System.Diagnostics.Process.GetProcessesByName("POWERPNT"))
                {
                    try { proc.Kill(); } catch { }
                }
            }
        }

        // PDF結合
        private async void BtnMergePDF_Click(object sender, RoutedEventArgs e)
        {
            var pdfFiles = Directory.GetFiles(pdfOutputFolder, "*.pdf").OrderBy(f => f).ToList();
            if (!pdfFiles.Any())
            {
                System.Windows.MessageBox.Show("結合するPDFファイルがありません", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var mergeFolder = Path.Combine(selectedFolderPath, "mergePDF");
            if (!Directory.Exists(mergeFolder))
                Directory.CreateDirectory(mergeFolder);

            // UIスレッドで値を取得してローカル変数に保存
            var mergeFileName = txtMergeFileName.Text;
            var addPageNumber = chkAddPageNumber.IsChecked == true;

            var timestamp = DateTime.Now.ToString("yyMMddHHmmss");
            var outputFileName = $"{mergeFileName}_{timestamp}.pdf";
            var outputPath = Path.Combine(mergeFolder, outputFileName);

            progressBar.Visibility = Visibility.Visible;
            progressBar.IsIndeterminate = true;
            txtStatus.Text = "PDF結合中...";

            await System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    using (var document = new Document())
                    using (var copy = new PdfCopy(document, new FileStream(outputPath, FileMode.Create)))
                    {
                        document.Open();

                        foreach (var pdfPath in pdfFiles)
                        {
                            using (var reader = new PdfReader(pdfPath))
                            {
                                for (int i = 1; i <= reader.NumberOfPages; i++)
                                {
                                    var page = copy.GetImportedPage(reader, i);
                                    copy.AddPage(page);
                                }
                            }
                        }
                    }

                    // ページ番号追加（オプション）
                    if (addPageNumber)
                    {
                        AddPageNumbers(outputPath);
                    }
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        System.Windows.MessageBox.Show($"PDF結合エラー: {ex.Message}", "エラー",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    });
                }
            });

            progressBar.Visibility = Visibility.Collapsed;
            txtStatus.Text = "PDF結合が完了しました";

            // 結合フォルダを開く
            System.Diagnostics.Process.Start("explorer.exe", mergeFolder);
        }

        // ページ番号追加
        private void AddPageNumbers(string pdfPath)
        {
            var tempPath = pdfPath + ".tmp";

            using (var reader = new PdfReader(pdfPath))
            using (var stamper = new PdfStamper(reader, new FileStream(tempPath, FileMode.Create)))
            {
                var totalPages = reader.NumberOfPages;

                for (int i = 1; i <= totalPages; i++)
                {
                    var cb = stamper.GetOverContent(i);
                    var pageSize = reader.GetPageSize(i);

                    cb.BeginText();
                    cb.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false), 10);
                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER,
                        $"{i} / {totalPages}",
                        pageSize.Width / 2, 20, 0);
                    cb.EndText();
                }
            }

            File.Delete(pdfPath);
            File.Move(tempPath, pdfPath);
        }

        // 全選択チェックボックス
        private void ChkSelectAll_Click(object sender, RoutedEventArgs e)
        {
            var isChecked = chkSelectAll.IsChecked ?? false;
            foreach (var item in fileItems)
            {
                item.IsSelected = isChecked;
            }
        }

        // COMオブジェクト解放
        private void ReleaseComObject(object? obj)
        {
            if (obj != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}