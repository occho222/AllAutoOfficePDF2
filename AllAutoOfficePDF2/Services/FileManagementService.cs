using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AllAutoOfficePDF2.Models;

namespace AllAutoOfficePDF2.Services
{
    /// <summary>
    /// ファイル管理サービス
    /// </summary>
    public class FileManagementService
    {
        /// <summary>
        /// 指定フォルダからファイルを読み込み
        /// </summary>
        /// <param name="folderPath">フォルダパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="includeSubfolders">サブフォルダを含むかどうか</param>
        /// <returns>ファイルアイテムリスト</returns>
        public static List<FileItem> LoadFilesFromFolder(string folderPath, string pdfOutputFolder, bool includeSubfolders = false)
        {
            var fileItems = new List<FileItem>();
            var extensions = new[] { "*.xls", "*.xlsx", "*.xlsm", "*.doc", "*.docx", "*.ppt", "*.pptx", "*.pdf" };

            var searchOption = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

            foreach (var ext in extensions)
            {
                var files = Directory.GetFiles(folderPath, ext, searchOption);
                foreach (var file in files)
                {
                    var fileInfo = new FileInfo(file);
                    string extensionUpper = fileInfo.Extension.TrimStart('.').ToUpper();
                    
                    var item = new FileItem
                    {
                        FileName = fileInfo.Name,
                        FilePath = fileInfo.FullName,
                        Extension = extensionUpper,
                        LastModified = fileInfo.LastWriteTime,
                        IsSelected = true,
                        PdfStatus = CheckPdfExists(fileInfo, pdfOutputFolder, folderPath, includeSubfolders) ? "変換済" : "未変換",
                        TargetPages = GetDefaultTargetPages(extensionUpper),
                        RelativePath = GetRelativePath(folderPath, fileInfo.FullName)
                    };
                    fileItems.Add(item);
                }
            }

            return fileItems.OrderBy(f => f.RelativePath).ThenBy(f => f.FileName).ToList();
        }

        /// <summary>
        /// 指定フォルダからファイルを読み込み（従来の互換性メソッド）
        /// </summary>
        /// <param name="folderPath">フォルダパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <returns>ファイルアイテムリスト</returns>
        public static List<FileItem> LoadFilesFromFolder(string folderPath, string pdfOutputFolder)
        {
            return LoadFilesFromFolder(folderPath, pdfOutputFolder, false);
        }

        /// <summary>
        /// ファイルの更新をチェック
        /// </summary>
        /// <param name="folderPath">フォルダパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="currentFileItems">現在のファイルアイテムリスト</param>
        /// <param name="includeSubfolders">サブフォルダを含むかどうか</param>
        /// <returns>更新されたファイルアイテムリスト</returns>
        public static (List<FileItem> UpdatedItems, List<string> ChangedFiles, List<string> AddedFiles, List<string> DeletedFiles) 
            UpdateFiles(string folderPath, string pdfOutputFolder, List<FileItem> currentFileItems, bool includeSubfolders = false)
        {
            var previousFiles = currentFileItems.ToDictionary(f => f.FilePath, f => f);
            var newFileItems = new List<FileItem>();
            var changedFiles = new List<string>();
            var addedFiles = new List<string>();
            var extensions = new[] { "*.xls", "*.xlsx", "*.xlsm", "*.doc", "*.docx", "*.ppt", "*.pptx", "*.pdf" };

            var searchOption = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

            foreach (var ext in extensions)
            {
                var files = Directory.GetFiles(folderPath, ext, searchOption);
                foreach (var file in files)
                {
                    var fileInfo = new FileInfo(file);
                    string extensionUpper = fileInfo.Extension.TrimStart('.').ToUpper();

                    bool isSelected = true;
                    string targetPages = GetDefaultTargetPages(extensionUpper);
                    int displayOrder = 0;

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
                            displayOrder = existingFile.DisplayOrder;
                        }
                    }
                    else
                    {
                        // 新規ファイルの場合
                        addedFiles.Add(fileInfo.Name);
                        isSelected = true;
                        displayOrder = previousFiles.Count + addedFiles.Count - 1;
                    }

                    var item = new FileItem
                    {
                        FileName = fileInfo.Name,
                        FilePath = fileInfo.FullName,
                        Extension = extensionUpper,
                        LastModified = fileInfo.LastWriteTime,
                        IsSelected = isSelected,
                        PdfStatus = CheckPdfExists(fileInfo, pdfOutputFolder, folderPath, includeSubfolders) ? "変換済" : "未変換",
                        TargetPages = targetPages,
                        DisplayOrder = displayOrder,
                        RelativePath = GetRelativePath(folderPath, fileInfo.FullName)
                    };
                    newFileItems.Add(item);
                }
            }

            // 削除されたファイルを検出
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
                        var pdfPath = GetPdfPath(previousFile.FilePath, pdfOutputFolder, folderPath, includeSubfolders);
                        if (File.Exists(pdfPath))
                        {
                            try
                            {
                                File.Delete(pdfPath);
                            }
                            catch (Exception ex)
                            {
                                // ログに記録するか、エラーを蓄積
                                System.Diagnostics.Debug.WriteLine($"PDFファイル削除エラー: {ex.Message}");
                            }
                        }
                    }
                }
            }

            // 表示順序で並び替え
            var orderedItems = newFileItems.OrderBy(f => f.DisplayOrder).ThenBy(f => f.RelativePath).ThenBy(f => f.FileName).ToList();
            
            // 番号を再設定
            for (int i = 0; i < orderedItems.Count; i++)
            {
                orderedItems[i].Number = i + 1;
                orderedItems[i].DisplayOrder = i;
            }

            return (orderedItems, changedFiles, addedFiles, deletedFiles);
        }

        /// <summary>
        /// ファイルの更新をチェック（従来の互換性メソッド）
        /// </summary>
        /// <param name="folderPath">フォルダパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="currentFileItems">現在のファイルアイテムリスト</param>
        /// <returns>更新されたファイルアイテムリスト</returns>
        public static (List<FileItem> UpdatedItems, List<string> ChangedFiles, List<string> AddedFiles, List<string> DeletedFiles) 
            UpdateFiles(string folderPath, string pdfOutputFolder, List<FileItem> currentFileItems)
        {
            return UpdateFiles(folderPath, pdfOutputFolder, currentFileItems, false);
        }

        /// <summary>
        /// PDFファイルの存在を確認
        /// </summary>
        /// <param name="fileInfo">ファイル情報</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="baseFolderPath">基準フォルダパス</param>
        /// <param name="includeSubfolders">サブフォルダを含むかどうか</param>
        /// <returns>PDFファイルが存在するかどうか</returns>
        private static bool CheckPdfExists(FileInfo fileInfo, string pdfOutputFolder, string baseFolderPath, bool includeSubfolders)
        {
            if (fileInfo.Extension.ToLower() == ".pdf") return true;

            var pdfPath = GetPdfPath(fileInfo.FullName, pdfOutputFolder, baseFolderPath, includeSubfolders);
            return File.Exists(pdfPath);
        }

        /// <summary>
        /// PDFファイルのパスを取得
        /// </summary>
        /// <param name="originalFilePath">元のファイルパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="baseFolderPath">基準フォルダパス</param>
        /// <param name="includeSubfolders">サブフォルダを含むかどうか</param>
        /// <returns>PDFファイルのパス</returns>
        private static string GetPdfPath(string originalFilePath, string pdfOutputFolder, string baseFolderPath, bool includeSubfolders)
        {
            var fileInfo = new FileInfo(originalFilePath);
            var fileName = Path.GetFileNameWithoutExtension(fileInfo.Name) + ".pdf";

            if (includeSubfolders)
            {
                // サブフォルダ構造を維持
                var relativePath = GetRelativePath(baseFolderPath, fileInfo.DirectoryName!);
                var outputDir = Path.Combine(pdfOutputFolder, relativePath);
                return Path.Combine(outputDir, fileName);
            }
            else
            {
                // すべてのファイルを同じフォルダに出力
                return Path.Combine(pdfOutputFolder, fileName);
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

        /// <summary>
        /// 拡張子に基づいてデフォルトの対象ページを取得
        /// </summary>
        /// <param name="extension">拡張子</param>
        /// <returns>デフォルトの対象ページ</returns>
        private static string GetDefaultTargetPages(string extension)
        {
            return extension switch
            {
                "XLS" or "XLSX" or "XLSM" => "1-1",
                _ => ""
            };
        }

        /// <summary>
        /// サブフォルダ用のPDF出力ディレクトリを作成
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <param name="pdfOutputFolder">PDF出力フォルダ</param>
        /// <param name="baseFolderPath">基準フォルダパス</param>
        /// <param name="includeSubfolders">サブフォルダを含むかどうか</param>
        public static void EnsurePdfOutputDirectory(string filePath, string pdfOutputFolder, string baseFolderPath, bool includeSubfolders)
        {
            if (includeSubfolders)
            {
                var fileInfo = new FileInfo(filePath);
                var relativePath = GetRelativePath(baseFolderPath, fileInfo.DirectoryName!);
                var outputDir = Path.Combine(pdfOutputFolder, relativePath);
                
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
            }
            else
            {
                if (!Directory.Exists(pdfOutputFolder))
                {
                    Directory.CreateDirectory(pdfOutputFolder);
                }
            }
        }
    }
}