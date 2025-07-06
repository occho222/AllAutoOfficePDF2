using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using AllAutoOfficePDF2.Models;
using System.Runtime.InteropServices;

namespace AllAutoOfficePDF2.Services
{
    /// <summary>
    /// PDF�ϊ��T�[�r�X
    /// </summary>
    public class PdfConversionService
    {
        /// <summary>
        /// Office������PDF�ɕϊ�
        /// </summary>
        /// <param name="filePath">�ϊ����t�@�C���p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="targetPages">�Ώۃy�[�W</param>
        /// <param name="baseFolderPath">��t�H���_�p�X�i�T�u�t�H���_�\���ێ��p�j</param>
        /// <param name="maintainSubfolderStructure">�T�u�t�H���_�\�����ێ����邩�ǂ���</param>
        public static void ConvertToPdf(string filePath, string pdfOutputFolder, string targetPages = "", 
            string baseFolderPath = "", bool maintainSubfolderStructure = false)
        {
            var extension = Path.GetExtension(filePath).ToLower();
            
            string outputPath;
            if (maintainSubfolderStructure && !string.IsNullOrEmpty(baseFolderPath))
            {
                // �T�u�t�H���_�\�����ێ�
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
                // �]���ʂ�A���ׂē����t�H���_�ɏo��
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
                    // PDF�t�@�C���̏ꍇ�̓R�s�[
                    if (!File.Exists(outputPath))
                        File.Copy(filePath, outputPath, overwrite: false);
                    break;
                default:
                    throw new NotSupportedException($"�Ή����Ă��Ȃ��t�@�C���`��: {extension}");
            }
        }

        /// <summary>
        /// Office������PDF�ɕϊ��i�]���̌݊������\�b�h�j
        /// </summary>
        /// <param name="filePath">�ϊ����t�@�C���p�X</param>
        /// <param name="pdfOutputFolder">PDF�o�̓t�H���_</param>
        /// <param name="targetPages">�Ώۃy�[�W</param>
        public static void ConvertToPdf(string filePath, string pdfOutputFolder, string targetPages = "")
        {
            ConvertToPdf(filePath, pdfOutputFolder, targetPages, "", false);
        }

        /// <summary>
        /// Excel��PDF�ϊ�
        /// </summary>
        /// <param name="inputPath">���̓t�@�C���p�X</param>
        /// <param name="outputPath">�o�̓t�@�C���p�X</param>
        /// <param name="targetPages">�Ώۃy�[�W</param>
        private static void ConvertExcelToPdf(string inputPath, string outputPath, string targetPages = "")
        {
            dynamic? excelApp = null;
            dynamic? workbook = null;

            try
            {
                // ������Excel�v���Z�X�������I��
                KillExistingExcelProcesses();

                // Excel�A�v���P�[�V�����𓮓I�ɍ쐬
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    throw new InvalidOperationException("Excel Application��������܂���B");
                }

                excelApp = Activator.CreateInstance(excelType);
                if (excelApp == null)
                {
                    throw new InvalidOperationException("Excel Application�̋N�����ł��܂���ł����B");
                }

                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;
                excelApp.ScreenUpdating = false;
                excelApp.EnableEvents = false;
                
                workbook = excelApp.Workbooks.Open(inputPath);

                var targetSheets = ParsePageRange(targetPages);

                if (targetSheets.Any())
                {
                    // �w��V�[�g�̂ݕϊ�
                    var totalSheets = workbook.Worksheets.Count;

                    // ���݂��Ȃ��V�[�g�̃`�F�b�N
                    var invalidSheets = targetSheets.Where(s => s > totalSheets).ToList();
                    if (invalidSheets.Any())
                    {
                        throw new ArgumentException($"���݂��Ȃ��V�[�g�ԍ����w�肳��Ă��܂�: {string.Join(", ", invalidSheets)} (���V�[�g��: {totalSheets})");
                    }

                    // �w��V�[�g��I��
                    for (int i = 0; i < targetSheets.Count; i++)
                    {
                        var sheet = workbook.Worksheets[targetSheets[i]];
                        
                        if (i == 0)
                        {
                            sheet.Select();
                        }
                        else
                        {
                            sheet.Select(false); // �ǉ��I��
                        }
                    }

                    // �I�����ꂽ�V�[�g��PDF�ϊ� (xlTypePDF = 0)
                    excelApp.ActiveSheet.ExportAsFixedFormat(0, outputPath);
                }
                else
                {
                    // �S�V�[�g�ϊ� (xlTypePDF = 0)
                    workbook.ExportAsFixedFormat(0, outputPath);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Excel�ϊ����ɃG���[���������܂���: {ex.Message}", ex);
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
                    Debug.WriteLine($"Excel�I�������ŃG���[: {ex.Message}");
                }

                if (workbook != null) ReleaseComObject(workbook);
                if (excelApp != null) ReleaseComObject(excelApp);

                // �����K�x�[�W�R���N�V����
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // �c���v���Z�X�������I��
                KillExistingExcelProcesses();
            }
        }

        /// <summary>
        /// Word��PDF�ϊ�
        /// </summary>
        /// <param name="inputPath">���̓t�@�C���p�X</param>
        /// <param name="outputPath">�o�̓t�@�C���p�X</param>
        /// <param name="targetPages">�Ώۃy�[�W</param>
        private static void ConvertWordToPdf(string inputPath, string outputPath, string targetPages = "")
        {
            Word.Application? wordApp = null;
            Word.Document? document = null;

            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;
                document = wordApp.Documents.Open(inputPath);

                var targetPageList = ParsePageRange(targetPages);

                if (targetPageList.Any())
                {
                    // �w��y�[�W�̂ݕϊ�
                    var totalPages = document.ComputeStatistics(Word.WdStatistic.wdStatisticPages);

                    // ���݂��Ȃ��y�[�W�̃`�F�b�N
                    var invalidPages = targetPageList.Where(p => p > totalPages).ToList();
                    if (invalidPages.Any())
                    {
                        throw new ArgumentException($"���݂��Ȃ��y�[�W�ԍ����w�肳��Ă��܂�: {string.Join(", ", invalidPages)} (���y�[�W��: {totalPages})");
                    }

                    // �y�[�W�͈͎w���PDF�o��
                    document.ExportAsFixedFormat(outputPath, Word.WdExportFormat.wdExportFormatPDF,
                        Range: Word.WdExportRange.wdExportFromTo,
                        From: targetPageList.Min(),
                        To: targetPageList.Max());
                }
                else
                {
                    // �S�y�[�W�ϊ�
                    document.ExportAsFixedFormat(outputPath, Word.WdExportFormat.wdExportFormatPDF);
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

        /// <summary>
        /// PowerPoint��PDF�ϊ�
        /// </summary>
        /// <param name="inputPath">���̓t�@�C���p�X</param>
        /// <param name="outputPath">�o�̓t�@�C���p�X</param>
        /// <param name="targetPages">�Ώۃy�[�W</param>
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
                    throw new InvalidOperationException("PowerPoint�A�v���P�[�V������������܂���B");
                }
                
                pptApp = Activator.CreateInstance(pptType);
                if (pptApp == null)
                {
                    throw new InvalidOperationException("PowerPoint�A�v���P�[�V�����̋N�����ł��܂���ł����B");
                }
                
                presentation = pptApp.Presentations.Open(inputPath);

                var targetSlides = ParsePageRange(targetPages);
                var totalSlides = presentation.Slides.Count;

                if (targetSlides.Any())
                {
                    // ���݂��Ȃ��X���C�h�̃`�F�b�N
                    var invalidSlides = targetSlides.Where(s => s > totalSlides).ToList();
                    if (invalidSlides.Any())
                    {
                        throw new ArgumentException($"���݂��Ȃ��X���C�h�ԍ����w�肳��Ă��܂�: {string.Join(", ", invalidSlides)} (���X���C�h��: {totalSlides})");
                    }

                    // �ꎞ�I�ɑS�X���C�h��PDF�ɕϊ�
                    tempPdfPath = Path.Combine(Path.GetTempPath(), $"temp_ppt_{Guid.NewGuid()}.pdf");
                    presentation.SaveAs(tempPdfPath, 32); // 32 = ppSaveAsPDF

                    // PowerPoint�����
                    presentation.Close();
                    pptApp.Quit();
                    ReleaseComObject(presentation);
                    ReleaseComObject(pptApp);
                    presentation = null;
                    pptApp = null;

                    // �ꎞ�I��GC���s
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    // �w�肳�ꂽ�y�[�W�݂̂𒊏o����PDF���쐬
                    ExtractPdfPages(tempPdfPath, outputPath, targetSlides);
                }
                else
                {
                    // �S�X���C�h�ϊ�
                    presentation.SaveAs(outputPath, 32); // 32 = ppSaveAsPDF
                }

                presentation?.Close();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"PowerPoint�ϊ����ɃG���[���������܂���: {ex.Message}", ex);
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

                // �ꎞ�t�@�C�����폜
                if (!string.IsNullOrEmpty(tempPdfPath) && File.Exists(tempPdfPath))
                {
                    try
                    {
                        File.Delete(tempPdfPath);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"�ꎞ�t�@�C���폜�Ɏ��s: {ex.Message}");
                    }
                }

                // �O�̂���GC��2��
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();

                // PowerPoint�v���Z�X���c���Ă���΋����I���i�ŏI��i�j
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
                    Debug.WriteLine($"PowerPoint�v���Z�X�I�������ŃG���[: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// ������Excel�v���Z�X�������I��
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
                
                // �����ҋ@
                System.Threading.Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Excel�v���Z�X�I�������ŃG���[: {ex.Message}");
            }
        }

        /// <summary>
        /// �y�[�W�͈͂����
        /// </summary>
        /// <param name="pageRange">�y�[�W�͈͕�����</param>
        /// <returns>�y�[�W�ԍ����X�g</returns>
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
                throw new ArgumentException($"�����ȃy�[�W�͈͎w��: {pageRange}");
            }

            return pages.OrderBy(p => p).ToList();
        }

        /// <summary>
        /// PDF����w�肳�ꂽ�y�[�W�݂̂𒊏o
        /// </summary>
        /// <param name="inputPdfPath">����PDF�t�@�C���p�X</param>
        /// <param name="outputPdfPath">�o��PDF�t�@�C���p�X</param>
        /// <param name="pageNumbers">�y�[�W�ԍ����X�g</param>
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
        /// COM�I�u�W�F�N�g�����
        /// </summary>
        /// <param name="obj">����ΏۃI�u�W�F�N�g</param>
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
                    Debug.WriteLine($"COM�I�u�W�F�N�g����G���[: {ex.Message}");
                }
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
    }
}