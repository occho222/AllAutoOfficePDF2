using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AllAutoOfficePDF2.Models;

namespace AllAutoOfficePDF2.Services
{
    /// <summary>
    /// �t�@�C���Ǘ��T�[�r�X
    /// </summary>
    public class FileManagementService
    {
        /// <summary>
        /// �w��t�H���_����t�@�C����ǂݍ���
        /// </summary>
        /// <param name="folderPath">�t�H���_�p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="includeSubfolders">�T�u�t�H���_���܂ނ��ǂ���</param>
        /// <returns>�t�@�C���A�C�e�����X�g</returns>
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
                        PdfStatus = CheckPdfExists(fileInfo, pdfOutputFolder, folderPath, includeSubfolders) ? "�ϊ���" : "���ϊ�",
                        TargetPages = GetDefaultTargetPages(extensionUpper),
                        RelativePath = GetRelativePath(folderPath, fileInfo.FullName)
                    };
                    fileItems.Add(item);
                }
            }

            return fileItems.OrderBy(f => f.RelativePath).ThenBy(f => f.FileName).ToList();
        }

        /// <summary>
        /// �w��t�H���_����t�@�C����ǂݍ��݁i�]���̌݊������\�b�h�j
        /// </summary>
        /// <param name="folderPath">�t�H���_�p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <returns>�t�@�C���A�C�e�����X�g</returns>
        public static List<FileItem> LoadFilesFromFolder(string folderPath, string pdfOutputFolder)
        {
            return LoadFilesFromFolder(folderPath, pdfOutputFolder, false);
        }

        /// <summary>
        /// �t�@�C���̍X�V���`�F�b�N
        /// </summary>
        /// <param name="folderPath">�t�H���_�p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="currentFileItems">���݂̃t�@�C���A�C�e�����X�g</param>
        /// <param name="includeSubfolders">�T�u�t�H���_���܂ނ��ǂ���</param>
        /// <returns>�X�V���ꂽ�t�@�C���A�C�e�����X�g</returns>
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

                    // �����t�@�C���̏ꍇ�͍X�V�������`�F�b�N
                    if (previousFiles.TryGetValue(file, out var existingFile))
                    {
                        if (existingFile.LastModified != fileInfo.LastWriteTime)
                        {
                            // �X�V�������ύX���ꂽ�ꍇ
                            changedFiles.Add(fileInfo.Name);
                            isSelected = true;
                        }
                        else
                        {
                            // �ύX����Ă��Ȃ��ꍇ�͑O�̑I����Ԃ�ێ�
                            isSelected = existingFile.IsSelected;
                            targetPages = existingFile.TargetPages;
                            displayOrder = existingFile.DisplayOrder;
                        }
                    }
                    else
                    {
                        // �V�K�t�@�C���̏ꍇ
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
                        PdfStatus = CheckPdfExists(fileInfo, pdfOutputFolder, folderPath, includeSubfolders) ? "�ϊ���" : "���ϊ�",
                        TargetPages = targetPages,
                        DisplayOrder = displayOrder,
                        RelativePath = GetRelativePath(folderPath, fileInfo.FullName)
                    };
                    newFileItems.Add(item);
                }
            }

            // �폜���ꂽ�t�@�C�������o
            var deletedFiles = new List<string>();
            var currentFilePaths = newFileItems.Select(f => f.FilePath).ToHashSet();

            foreach (var previousFile in previousFiles.Values)
            {
                if (!currentFilePaths.Contains(previousFile.FilePath))
                {
                    deletedFiles.Add(previousFile.FileName);

                    // �Ή�����PDF�t�@�C�����폜
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
                                // ���O�ɋL�^���邩�A�G���[��~��
                                System.Diagnostics.Debug.WriteLine($"PDF�t�@�C���폜�G���[: {ex.Message}");
                            }
                        }
                    }
                }
            }

            // �\�������ŕ��ёւ�
            var orderedItems = newFileItems.OrderBy(f => f.DisplayOrder).ThenBy(f => f.RelativePath).ThenBy(f => f.FileName).ToList();
            
            // �ԍ����Đݒ�
            for (int i = 0; i < orderedItems.Count; i++)
            {
                orderedItems[i].Number = i + 1;
                orderedItems[i].DisplayOrder = i;
            }

            return (orderedItems, changedFiles, addedFiles, deletedFiles);
        }

        /// <summary>
        /// �t�@�C���̍X�V���`�F�b�N�i�]���̌݊������\�b�h�j
        /// </summary>
        /// <param name="folderPath">�t�H���_�p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="currentFileItems">���݂̃t�@�C���A�C�e�����X�g</param>
        /// <returns>�X�V���ꂽ�t�@�C���A�C�e�����X�g</returns>
        public static (List<FileItem> UpdatedItems, List<string> ChangedFiles, List<string> AddedFiles, List<string> DeletedFiles) 
            UpdateFiles(string folderPath, string pdfOutputFolder, List<FileItem> currentFileItems)
        {
            return UpdateFiles(folderPath, pdfOutputFolder, currentFileItems, false);
        }

        /// <summary>
        /// PDF�t�@�C���̑��݂��m�F
        /// </summary>
        /// <param name="fileInfo">�t�@�C�����</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="baseFolderPath">��t�H���_�p�X</param>
        /// <param name="includeSubfolders">�T�u�t�H���_���܂ނ��ǂ���</param>
        /// <returns>PDF�t�@�C�������݂��邩�ǂ���</returns>
        private static bool CheckPdfExists(FileInfo fileInfo, string pdfOutputFolder, string baseFolderPath, bool includeSubfolders)
        {
            if (fileInfo.Extension.ToLower() == ".pdf") return true;

            var pdfPath = GetPdfPath(fileInfo.FullName, pdfOutputFolder, baseFolderPath, includeSubfolders);
            return File.Exists(pdfPath);
        }

        /// <summary>
        /// PDF�t�@�C���̃p�X���擾
        /// </summary>
        /// <param name="originalFilePath">���̃t�@�C���p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="baseFolderPath">��t�H���_�p�X</param>
        /// <param name="includeSubfolders">�T�u�t�H���_���܂ނ��ǂ���</param>
        /// <returns>PDF�t�@�C���̃p�X</returns>
        private static string GetPdfPath(string originalFilePath, string pdfOutputFolder, string baseFolderPath, bool includeSubfolders)
        {
            var fileInfo = new FileInfo(originalFilePath);
            var fileName = Path.GetFileNameWithoutExtension(fileInfo.Name) + ".pdf";

            if (includeSubfolders)
            {
                // �T�u�t�H���_�\�����ێ�
                var relativePath = GetRelativePath(baseFolderPath, fileInfo.DirectoryName!);
                var outputDir = Path.Combine(pdfOutputFolder, relativePath);
                return Path.Combine(outputDir, fileName);
            }
            else
            {
                // ���ׂẴt�@�C���𓯂��t�H���_�ɏo��
                return Path.Combine(pdfOutputFolder, fileName);
            }
        }

        /// <summary>
        /// ���΃p�X���擾
        /// </summary>
        /// <param name="basePath">��p�X</param>
        /// <param name="fullPath">���S�p�X</param>
        /// <returns>���΃p�X</returns>
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
        /// �g���q�Ɋ�Â��ăf�t�H���g�̑Ώۃy�[�W���擾
        /// </summary>
        /// <param name="extension">�g���q</param>
        /// <returns>�f�t�H���g�̑Ώۃy�[�W</returns>
        private static string GetDefaultTargetPages(string extension)
        {
            return extension switch
            {
                "XLS" or "XLSX" or "XLSM" => "1-1",
                _ => ""
            };
        }

        /// <summary>
        /// �T�u�t�H���_�p��PDF�o�̓f�B���N�g�����쐬
        /// </summary>
        /// <param name="filePath">�t�@�C���p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="baseFolderPath">��t�H���_�p�X</param>
        /// <param name="includeSubfolders">�T�u�t�H���_���܂ނ��ǂ���</param>
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