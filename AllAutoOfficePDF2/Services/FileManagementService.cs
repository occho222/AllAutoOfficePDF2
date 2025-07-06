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
        /// <returns>�t�@�C���A�C�e�����X�g</returns>
        public static List<FileItem> LoadFilesFromFolder(string folderPath, string pdfOutputFolder)
        {
            var fileItems = new List<FileItem>();
            var extensions = new[] { "*.xls", "*.xlsx", "*.xlsm", "*.doc", "*.docx", "*.ppt", "*.pptx", "*.pdf" };

            foreach (var ext in extensions)
            {
                var files = Directory.GetFiles(folderPath, ext);
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
                        PdfStatus = CheckPdfExists(fileInfo, pdfOutputFolder) ? "�ϊ���" : "���ϊ�",
                        TargetPages = GetDefaultTargetPages(extensionUpper)
                    };
                    fileItems.Add(item);
                }
            }

            return fileItems.OrderBy(f => f.FileName).ToList();
        }

        /// <summary>
        /// �t�@�C���̍X�V���`�F�b�N
        /// </summary>
        /// <param name="folderPath">�t�H���_�p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="currentFileItems">���݂̃t�@�C���A�C�e�����X�g</param>
        /// <returns>�X�V���ꂽ�t�@�C���A�C�e�����X�g</returns>
        public static (List<FileItem> UpdatedItems, List<string> ChangedFiles, List<string> AddedFiles, List<string> DeletedFiles) 
            UpdateFiles(string folderPath, string pdfOutputFolder, List<FileItem> currentFileItems)
        {
            var previousFiles = currentFileItems.ToDictionary(f => f.FilePath, f => f);
            var newFileItems = new List<FileItem>();
            var changedFiles = new List<string>();
            var addedFiles = new List<string>();
            var extensions = new[] { "*.xls", "*.xlsx", "*.xlsm", "*.doc", "*.docx", "*.ppt", "*.pptx", "*.pdf" };

            foreach (var ext in extensions)
            {
                var files = Directory.GetFiles(folderPath, ext);
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
                        PdfStatus = CheckPdfExists(fileInfo, pdfOutputFolder) ? "�ϊ���" : "���ϊ�",
                        TargetPages = targetPages,
                        DisplayOrder = displayOrder
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
                        var pdfPath = Path.Combine(pdfOutputFolder, 
                            Path.GetFileNameWithoutExtension(previousFile.FileName) + ".pdf");
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
            var orderedItems = newFileItems.OrderBy(f => f.DisplayOrder).ThenBy(f => f.FileName).ToList();
            
            // �ԍ����Đݒ�
            for (int i = 0; i < orderedItems.Count; i++)
            {
                orderedItems[i].Number = i + 1;
                orderedItems[i].DisplayOrder = i;
            }

            return (orderedItems, changedFiles, addedFiles, deletedFiles);
        }

        /// <summary>
        /// PDF�t�@�C���̑��݂��m�F
        /// </summary>
        /// <param name="fileInfo">�t�@�C�����</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <returns>PDF�t�@�C�������݂��邩�ǂ���</returns>
        private static bool CheckPdfExists(FileInfo fileInfo, string pdfOutputFolder)
        {
            if (fileInfo.Extension.ToLower() == ".pdf") return true;

            var pdfPath = Path.Combine(pdfOutputFolder, 
                Path.GetFileNameWithoutExtension(fileInfo.Name) + ".pdf");
            return File.Exists(pdfPath);
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
    }
}