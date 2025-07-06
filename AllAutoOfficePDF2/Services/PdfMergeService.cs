using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Threading.Tasks;
using AllAutoOfficePDF2.Models;

namespace AllAutoOfficePDF2.Services
{
    /// <summary>
    /// PDF�����T�[�r�X
    /// </summary>
    public class PdfMergeService
    {
        /// <summary>
        /// PDF�t�@�C��������
        /// </summary>
        /// <param name="pdfFilePaths">��������PDF�t�@�C���p�X���X�g</param>
        /// <param name="outputPath">�o�̓t�@�C���p�X</param>
        /// <param name="addPageNumber">�y�[�W�ԍ��ǉ��t���O</param>
        public static void MergePdfFiles(List<string> pdfFilePaths, string outputPath, bool addPageNumber = false)
        {
            using (var document = new Document())
            using (var copy = new PdfCopy(document, new FileStream(outputPath, FileMode.Create)))
            {
                document.Open();

                foreach (var pdfPath in pdfFilePaths)
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

            // �y�[�W�ԍ��ǉ��i�I�v�V�����j
            if (addPageNumber)
            {
                AddPageNumbers(outputPath);
            }
        }

        /// <summary>
        /// PDF�Ƀy�[�W�ԍ���ǉ�
        /// </summary>
        /// <param name="pdfPath">PDF�t�@�C���p�X</param>
        private static void AddPageNumbers(string pdfPath)
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
                    // �y�[�W�ԍ����E��ɔz�u�ix���W�F�E�[����20pt�Ay���W�F��[����20pt�j
                    cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT,
                        $"{i} / {totalPages}",
                        pageSize.Width - 20, pageSize.Height - 20, 0);
                    cb.EndText();
                }
            }

            File.Delete(pdfPath);
            File.Move(tempPath, pdfPath);
        }
    }
}