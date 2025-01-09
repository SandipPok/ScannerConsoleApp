using DocumentFormat.OpenXml.Packaging;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System.Drawing.Imaging;
using System.Drawing;

namespace ScannerApp
{
    internal class DocumentExporter
    {
        private List<Bitmap> ScannedImages { get; }

        public DocumentExporter(List<Bitmap> scannedImages)
        {
            ScannedImages = scannedImages;
        }

        private string GetFileSavePath(string extension)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string folderPath = Path.Combine(desktopPath, "ScannedDocuments");

            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            string uniqueFileName = DateTime.Now.ToString("yyyyMMdd_HHmmss") + extension;
            return Path.Combine(folderPath, uniqueFileName);
        }

        public void ExportAsImage()
        {
            string savePath = GetFileSavePath(".jpg");

            Console.WriteLine($"Saving image to: {savePath}");

            foreach (var image in ScannedImages)
            {
                image.Save(savePath, ImageFormat.Jpeg);
                Console.WriteLine($"Image saved to {savePath}");
            }
        }

        public void ExportAsPdf()
        {
            string savePath = GetFileSavePath(".pdf");

            Console.WriteLine($"Saving PDF to: {savePath}");

            PdfDocument pdf = new PdfDocument();
            pdf.Info.Title = "Scanned Document";

            foreach (var image in ScannedImages)
            {
                PdfPage page = pdf.AddPage();
                XGraphics gfx = XGraphics.FromPdfPage(page);

                string tempFilePath = Path.GetTempFileName();
                image.Save(tempFilePath, ImageFormat.Jpeg);
                XImage xImage = XImage.FromFile(tempFilePath);

                double aspectRatio = (double)xImage.PixelWidth / xImage.PixelHeight;
                double pageWidth = page.Width;
                double pageHeight = page.Height;
                double imageWidth = pageWidth;
                double imageHeight = pageWidth / aspectRatio;

                if (imageHeight > pageHeight)
                {
                    imageHeight = pageHeight;
                    imageWidth = pageHeight * aspectRatio;
                }

                gfx.DrawImage(xImage, 0, 0, imageWidth, imageHeight);
                File.Delete(tempFilePath);
            }

            pdf.Save(savePath);
            Console.WriteLine($"PDF saved to {savePath}");
        }


        public void ExportAsWord()
        {
            string savePath = GetFileSavePath(".docx");

            Console.WriteLine($"Saving Word document to: {savePath}");

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(savePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(new DocumentFormat.OpenXml.Wordprocessing.Body());

                foreach (var image in ScannedImages)
                {
                    mainPart.Document.Body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text("Scanned Document Image"))));
                }

                wordDocument.Save();
            }

            Console.WriteLine($"Word document saved to {savePath}");
        }
    }
}
