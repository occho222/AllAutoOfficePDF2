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
    /// PDF結合サービス
    /// </summary>
    public class PdfMergeService
    {
        /// <summary>
        /// PDFファイルを結合
        /// </summary>
        /// <param name="pdfFilePaths">結合するPDFファイルパスリスト</param>
        /// <param name="outputPath">出力ファイルパス</param>
        /// <param name="addPageNumber">ページ番号追加フラグ</param>
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

            // ページ番号追加（オプション）
            if (addPageNumber)
            {
                AddPageNumbers(outputPath);
            }
        }

        /// <summary>
        /// PDFにページ番号を追加
        /// </summary>
        /// <param name="pdfPath">PDFファイルパス</param>
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
                    // ページ番号を右上に配置（x座標：右端から20pt、y座標：上端から20pt）
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