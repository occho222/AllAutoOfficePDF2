using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Threading.Tasks;
using System.Diagnostics;
using AllAutoOfficePDF2.Models;
using System.Runtime.InteropServices;

namespace AllAutoOfficePDF2.Services
{
    /// <summary>
    /// PDF変換サービス
    /// </summary>
    public class PdfConversionService
    {
        /// <summary>
        /// Office文書をPDFに変換
        /// </summary>
        /// <param name="filePath">変換元ファイルパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="targetPages">対象ページ</param>
        /// <param name="baseFolderPath">基準フォルダパス（サブフォルダ構造維持用）</param>
        /// <param name="maintainSubfolderStructure">サブフォルダ構造を維持するかどうか</param>
        public static void ConvertToPdf(string filePath, string pdfOutputFolder, string targetPages = "", 
            string baseFolderPath = "", bool maintainSubfolderStructure = false)
        {
            var extension = Path.GetExtension(filePath).ToLower();
            
            string outputPath;
            if (maintainSubfolderStructure && !string.IsNullOrEmpty(baseFolderPath))
            {
                // サブフォルダ構造を維持
                var fileInfo = new FileInfo(filePath);
                var relativePath = GetRelativePath(baseFolderPath, fileInfo.DirectoryName!);
                var outputDir = Path.Combine(pdfOutputFolder, relativePath);
                
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                
                outputPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(filePath) + ".pdf");
            }
            else
            {
                // 従来通り、すべて同じフォルダに出力
                outputPath = Path.Combine(pdfOutputFolder, Path.GetFileNameWithoutExtension(filePath) + ".pdf");
            }

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
                    // PDFファイルの場合はコピー
                    if (!File.Exists(outputPath))
                        File.Copy(filePath, outputPath, overwrite: false);
                    break;
                default:
                    throw new NotSupportedException($"対応していないファイル形式: {extension}");
            }
        }

        /// <summary>
        /// Office文書をPDFに変換（従来の互換性メソッド）
        /// </summary>
        /// <param name="filePath">変換元ファイルパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="targetPages">対象ページ</param>
        public static void ConvertToPdf(string filePath, string pdfOutputFolder, string targetPages = "")
        {
            ConvertToPdf(filePath, pdfOutputFolder, targetPages, "", false);
        }

        /// <summary>
        /// Excel→PDF変換
        /// </summary>
        /// <param name="inputPath">入力ファイルパス</param>
        /// <param name="outputPath">出力ファイルパス</param>
        /// <param name="targetPages">対象ページ</param>
        private static void ConvertExcelToPdf(string inputPath, string outputPath, string targetPages = "")
        {
            dynamic? excelApp = null;
            dynamic? workbook = null;

            try
            {
                // 既存のExcelプロセスを強制終了
                KillExistingExcelProcesses();

                // Excelアプリケーションを動的に作成
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    throw new InvalidOperationException("Excel Applicationが見つかりません。");
                }

                excelApp = Activator.CreateInstance(excelType);
                if (excelApp == null)
                {
                    throw new InvalidOperationException("Excel Applicationの起動ができませんでした。");
                }

                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;
                excelApp.ScreenUpdating = false;
                excelApp.EnableEvents = false;
                
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
                    for (int i = 0; i < targetSheets.Count; i++)
                    {
                        var sheet = workbook.Worksheets[targetSheets[i]];
                        
                        if (i == 0)
                        {
                            sheet.Select();
                        }
                        else
                        {
                            sheet.Select(false); // 追加選択
                        }
                    }

                    // 選択されたシートをPDF変換 (xlTypePDF = 0)
                    excelApp.ActiveSheet.ExportAsFixedFormat(0, outputPath);
                }
                else
                {
                    // 全シート変換 (xlTypePDF = 0)
                    workbook.ExportAsFixedFormat(0, outputPath);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Excel変換中にエラーが発生しました: {ex.Message}", ex);
            }
            finally
            {
                try
                {
                    workbook?.Close(false);
                    excelApp?.Quit();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Excel終了処理でエラー: {ex.Message}");
                }

                if (workbook != null) ReleaseComObject(workbook);
                if (excelApp != null) ReleaseComObject(excelApp);

                // 強制ガベージコレクション
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // 残存プロセスを強制終了
                KillExistingExcelProcesses();
            }
        }

        /// <summary>
        /// Word→PDF変換
        /// </summary>
        /// <param name="inputPath">入力ファイルパス</param>
        /// <param name="outputPath">出力ファイルパス</param>
        /// <param name="targetPages">対象ページ</param>
        private static void ConvertWordToPdf(string inputPath, string outputPath, string targetPages = "")
        {
            dynamic? wordApp = null;
            dynamic? document = null;

            try
            {
                // 既存のWordプロセスをクリーンアップ
                KillExistingWordProcesses();

                // Wordアプリケーションを動的に作成
                var wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null)
                {
                    throw new InvalidOperationException("Word Applicationが見つかりません。");
                }

                wordApp = Activator.CreateInstance(wordType);
                if (wordApp == null)
                {
                    throw new InvalidOperationException("Word Applicationの起動ができませんでした。");
                }

                wordApp.Visible = false;
                wordApp.DisplayAlerts = 0; // wdAlertsNone = 0
                document = wordApp.Documents.Open(inputPath);

                var targetPageList = ParsePageRange(targetPages);

                if (targetPageList.Any())
                {
                    // 指定ページのみ変換
                    var totalPages = document.ComputeStatistics(4); // wdStatisticPages = 4

                    // 存在しないページのチェック
                    var invalidPages = targetPageList.Where(p => p > totalPages).ToList();
                    if (invalidPages.Any())
                    {
                        throw new ArgumentException($"存在しないページ番号が指定されています: {string.Join(", ", invalidPages)} (総ページ数: {totalPages})");
                    }

                    // ページ範囲指定でPDF出力
                    // wdExportFormatPDF = 17, wdExportFromTo = 3
                    document.ExportAsFixedFormat(outputPath, 17, Range: 3, From: targetPageList.Min(), To: targetPageList.Max());
                }
                else
                {
                    // 全ページ変換
                    // wdExportFormatPDF = 17
                    document.ExportAsFixedFormat(outputPath, 17);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Word変換中にエラーが発生しました: {ex.Message}", ex);
            }
            finally
            {
                try
                {
                    document?.Close(false);
                    wordApp?.Quit();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Word終了処理でエラー: {ex.Message}");
                }

                if (document != null) ReleaseComObject(document);
                if (wordApp != null) ReleaseComObject(wordApp);

                // 強制ガベージコレクション
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // 残存プロセスを強制終了
                KillExistingWordProcesses();
            }
        }

        /// <summary>
        /// PowerPoint→PDF変換
        /// </summary>
        /// <param name="inputPath">入力ファイルパス</param>
        /// <param name="outputPath">出力ファイルパス</param>
        /// <param name="targetPages">対象ページ</param>
        private static void ConvertPowerPointToPdf(string inputPath, string outputPath, string targetPages = "")
        {
            dynamic? pptApp = null;
            dynamic? presentation = null;
            string? tempPdfPath = null;

            try
            {
                var pptType = Type.GetTypeFromProgID("PowerPoint.Application");
                if (pptType == null)
                {
                    throw new InvalidOperationException("PowerPointアプリケーションが見つかりません。");
                }
                
                pptApp = Activator.CreateInstance(pptType);
                if (pptApp == null)
                {
                    throw new InvalidOperationException("PowerPointアプリケーションの起動ができませんでした。");
                }
                
                presentation = pptApp.Presentations.Open(inputPath);

                var targetSlides = ParsePageRange(targetPages);
                var totalSlides = presentation.Slides.Count;

                if (targetSlides.Any())
                {
                    // 存在しないスライドのチェック
                    var invalidSlides = targetSlides.Where(s => s > totalSlides).ToList();
                    if (invalidSlides.Any())
                    {
                        throw new ArgumentException($"存在しないスライド番号が指定されています: {string.Join(", ", invalidSlides)} (総スライド数: {totalSlides})");
                    }

                    // 一時的に全スライドをPDFに変換
                    tempPdfPath = Path.Combine(Path.GetTempPath(), $"temp_ppt_{Guid.NewGuid()}.pdf");
                    presentation.SaveAs(tempPdfPath, 32); // 32 = ppSaveAsPDF

                    // PowerPointを閉じる
                    presentation.Close();
                    pptApp.Quit();
                    ReleaseComObject(presentation);
                    ReleaseComObject(pptApp);
                    presentation = null;
                    pptApp = null;

                    // 一時的なGC実行
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    // 指定されたページのみを抽出してPDFを作成
                    ExtractPdfPages(tempPdfPath, outputPath, targetSlides);
                }
                else
                {
                    // 全スライド変換
                    presentation.SaveAs(outputPath, 32); // 32 = ppSaveAsPDF
                }

                presentation?.Close();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"PowerPoint変換中にエラーが発生しました: {ex.Message}", ex);
            }
            finally
            {
                try 
                { 
                    presentation?.Close();
                    pptApp?.Quit(); 
                } 
                catch { }
                
                if (presentation != null) ReleaseComObject(presentation);
                if (pptApp != null) ReleaseComObject(pptApp);

                // 一時ファイルを削除
                if (!string.IsNullOrEmpty(tempPdfPath) && File.Exists(tempPdfPath))
                {
                    try
                    {
                        File.Delete(tempPdfPath);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"一時ファイル削除に失敗: {ex.Message}");
                    }
                }

                // 念のためGCを2回
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();

                // PowerPointプロセスが残っていれば強制終了（最終手段）
                try
                {
                    foreach (var proc in Process.GetProcessesByName("POWERPNT"))
                    {
                        if (!proc.HasExited)
                        {
                            proc.Kill();
                            proc.WaitForExit(1000);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"PowerPointプロセス終了処理でエラー: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 既存のExcelプロセスを強制終了
        /// </summary>
        private static void KillExistingExcelProcesses()
        {
            try
            {
                var processes = Process.GetProcessesByName("EXCEL");
                foreach (var proc in processes)
                {
                    try
                    {
                        if (!proc.HasExited)
                        {
                            proc.Kill();
                            proc.WaitForExit(3000);
                        }
                    }
                    catch { }
                    finally
                    {
                        proc.Dispose();
                    }
                }
                
                // 少し待機
                System.Threading.Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Excelプロセス終了処理でエラー: {ex.Message}");
            }
        }

        /// <summary>
        /// 既存のWordプロセスを強制終了
        /// </summary>
        private static void KillExistingWordProcesses()
        {
            try
            {
                var processes = Process.GetProcessesByName("WINWORD");
                foreach (var proc in processes)
                {
                    try
                    {
                        if (!proc.HasExited)
                        {
                            proc.Kill();
                            proc.WaitForExit(3000);
                        }
                    }
                    catch { }
                    finally
                    {
                        proc.Dispose();
                    }
                }
                
                // 少し待機
                System.Threading.Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Wordプロセス終了処理でエラー: {ex.Message}");
            }
        }

        /// <summary>
        /// ページ範囲を解析
        /// </summary>
        /// <param name="pageRange">ページ範囲文字列</param>
        /// <returns>ページ番号リスト</returns>
        private static List<int> ParsePageRange(string pageRange)
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

        /// <summary>
        /// PDFから指定されたページのみを抽出
        /// </summary>
        /// <param name="inputPdfPath">入力PDFファイルパス</param>
        /// <param name="outputPdfPath">出力PDFファイルパス</param>
        /// <param name="pageNumbers">ページ番号リスト</param>
        private static void ExtractPdfPages(string inputPdfPath, string outputPdfPath, List<int> pageNumbers)
        {
            using (var inputReader = new PdfReader(inputPdfPath))
            using (var outputDocument = new Document())
            using (var outputWriter = new PdfCopy(outputDocument, new FileStream(outputPdfPath, FileMode.Create)))
            {
                outputDocument.Open();

                foreach (var pageNumber in pageNumbers.OrderBy(p => p))
                {
                    if (pageNumber <= inputReader.NumberOfPages)
                    {
                        var page = outputWriter.GetImportedPage(inputReader, pageNumber);
                        outputWriter.AddPage(page);
                    }
                }
            }
        }

        /// <summary>
        /// COMオブジェクトを解放
        /// </summary>
        /// <param name="obj">解放対象オブジェクト</param>
        private static void ReleaseComObject(object? obj)
        {
            if (obj != null)
            {
                try
                {
                    Marshal.ReleaseComObject(obj);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"COMオブジェクト解放エラー: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 相対パスを取得
        /// </summary>
        /// <param name="basePath">基準パス</param>
        /// <param name="fullPath">完全パス</param>
        /// <returns>相対パス</returns>
        private static string GetRelativePath(string basePath, string fullPath)
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
    }
}