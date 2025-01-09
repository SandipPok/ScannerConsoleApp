using System.Drawing;
using System.Drawing.Imaging;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using DocumentFormat.OpenXml.Packaging;
using WIA;
using ScannerApp;

namespace ScanExportApp
{
    //class Program
    //{
    //    public static List<Bitmap> ScannedImages = new List<Bitmap>();

    //    static void Main(string[] args)
    //    {
    //        Console.WriteLine("Scanner App - .NET 8");

    //        while (true)
    //        {
    //            Console.WriteLine("Scanning Document...");
    //            ScanDocument();

    //            Console.WriteLine("Would you like to scan another document? (y/n): ");
    //            string userInput = Console.ReadLine();
    //            if (userInput?.ToLower() != "y")
    //            {
    //                break;
    //            }
    //        }

    //        Console.WriteLine("What would you like to do with the scanned documents?");
    //        Console.WriteLine("1. Export as Photo (JPG/PNG)");
    //        Console.WriteLine("2. Export as PDF");
    //        Console.WriteLine("3. Export as Word");

    //        string choice = Console.ReadLine();

    //        switch (choice)
    //        {
    //            case "1":
    //                ExportAsImage();
    //                break;
    //            case "2":
    //                ExportAsPdf();
    //                break;
    //            case "3":
    //                ExportAsWord();
    //                break;
    //            default:
    //                Console.WriteLine("Invalid choice.");
    //                break;
    //        }

    //        Console.WriteLine("Operation completed.");
    //    }

    //    // Simulate scanning a document (replace with actual scanning logic using WIA/TWAIN)
    //    static void ScanDocument()
    //    {
    //        var deviceManager = new DeviceManager();

    //        if (deviceManager.DeviceInfos == null) 
    //        { 
    //            Console.WriteLine("DeviceInfos is null."); 
    //            return; 
    //        }
    //        Console.WriteLine($"DeviceInfos count: {deviceManager.DeviceInfos.Count}"); 
    //        if (deviceManager.DeviceInfos.Count == 0) 
    //        { 
    //            Console.WriteLine("No devices found. Please check if your scanner is connected and powered on, and that the WIA service is running."); 
    //            return; 
    //        }

    //        List<DeviceInfo> scannerDevices = new List<DeviceInfo>();

    //        // Collect all available scanner devices
    //        for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++)
    //        {
    //            if (deviceManager.DeviceInfos[i].Type == WiaDeviceType.ScannerDeviceType)
    //            {
    //                scannerDevices.Add(deviceManager.DeviceInfos[i]);
    //            }
    //        }

    //        if (scannerDevices.Count == 0)
    //        {
    //            Console.WriteLine("No scanner devices found.");
    //            return;
    //        }

    //        Console.WriteLine("Available scanners: ");
    //        for (int i = 0; i < scannerDevices.Count; i++)
    //        {
    //            Console.WriteLine($"{i + 1}. {scannerDevices[i].Properties["Name"].get_Value()}");
    //        }

    //        Console.Write("Select a scanner (enter the number): ");

    //        if (int.TryParse(Console.ReadLine(), out int selectedScannerIndex) && selectedScannerIndex > 0 && selectedScannerIndex <= scannerDevices.Count)
    //        {
    //            var deviceInfo = scannerDevices[selectedScannerIndex - 1];
    //            var device = deviceInfo.Connect();
    //            var item = device.Items[1];

    //            // Ask for DPI values
    //            Console.Write("Enter horizontal DPI value (e.g., 300): ");
    //            int horizontalDpi = int.TryParse(Console.ReadLine(), out horizontalDpi) ? horizontalDpi : 300;

    //            Console.Write("Enter vertical DPI value (e.g., 300): ");
    //            int verticalDpi = int.TryParse(Console.ReadLine(), out verticalDpi) ? verticalDpi : 300;

    //            SetDpi(item, horizontalDpi, verticalDpi);

    //            // Scan the document
    //            var imgFile = (ImageFile)item.Transfer("{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}");

    //            // Convert the scanned image to a Bitmap and add to the list
    //            var imageBytes = (byte[])imgFile.FileData.get_BinaryData();
    //            using (var ms = new MemoryStream(imageBytes))
    //            {
    //                Bitmap scannedImage = new Bitmap(ms);
    //                ScannedImages.Add(scannedImage);
    //                Console.WriteLine("Document scanned successfully");
    //            }
    //        }
    //        else
    //        {
    //            Console.WriteLine("Invalid selection");
    //        }
    //    }

    //    // Export the scanned images as a single photo (JPG/PNG)
    //    static void ExportAsImage()
    //    {
    //        string savePath = GetFileSavePath(".jpg");

    //        if (string.IsNullOrEmpty(savePath))
    //        {
    //            Console.WriteLine("Invalid path.");
    //            return;
    //        }

    //        foreach (var image in ScannedImages)
    //        {
    //            image.Save(savePath, ImageFormat.Jpeg);  // Or PNG, BMP, etc.
    //            Console.WriteLine($"Image saved to {savePath}");
    //        }
    //    }

    //    // Export the scanned images as a PDF
    //    static void ExportAsPdf()
    //    {
    //        string savePath = GetFileSavePath(".pdf");

    //        //Create a new PDF document
    //        PdfDocument pdf = new PdfDocument();
    //        pdf.Info.Title = "Scanned Document";

    //        foreach (var image in ScannedImages)
    //        {
    //            // Add a page to the PDF for each scnned image
    //            PdfPage page = pdf.AddPage();
    //            XGraphics gfx = XGraphics.FromPdfPage(page);

    //            // Save the image to a temporary file and then load it with XImage
    //            string tempFilePath = Path.GetTempFileName();
    //            image.Save(tempFilePath, System.Drawing.Imaging.ImageFormat.Jpeg); // Save the image temporarily
    //            XImage xImage = XImage.FromFile(tempFilePath);

    //            // Scale the image to fit the page size, keeping the aspect ratio
    //            double aspectRatio = (double)xImage.PixelWidth / xImage.PixelHeight;
    //            double pageWidth = page.Width;
    //            double pageHeight = page.Height;
    //            double imageWidth = pageWidth;
    //            double imageHeight = pageWidth / aspectRatio;

    //            if (imageHeight > pageHeight)
    //            {
    //                imageHeight = pageHeight;
    //                imageWidth = pageHeight * aspectRatio;
    //            }

    //            // Draw the image onto the PDF page
    //            gfx.DrawImage(xImage, 0, 0, imageWidth, imageHeight);

    //            // Clean up the temporary file
    //            File.Delete(tempFilePath);
    //        }
    //    }

    //    // Export the scanned images as a Word document
    //    static void ExportAsWord()
    //    {
    //        string savePath = GetFileSavePath(".docx");

    //        using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(savePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
    //        {
    //            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
    //            mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(new DocumentFormat.OpenXml.Wordprocessing.Body());

    //            foreach (var image in ScannedImages)
    //            {
    //                // Convert the image to base64 or handle embedding as a picture in Word (requires more logic)
    //                // Here we simply add a dummy paragraph for the sake of example
    //                mainPart.Document.Body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text("Scanned Document Image"))));
    //            }

    //            wordDocument.Save();
    //        }

    //        Console.WriteLine($"Word document saved to {savePath}");
    //    }

    //    static void SetDpi(IItem item, int horizontalDpi, int verticalDpi)
    //    {
    //        const string HorizontalResolution = "6147"; // WIA_HORIZONTAL_SCAN_RESOLUTION_DPI
    //        const string VerticalResolution = "6148";   // WIA_VERTICAL_SCAN_RESOLUTION_DPI

    //        SetItemProperty(item, HorizontalResolution, horizontalDpi);
    //        SetItemProperty(item, VerticalResolution, verticalDpi);
    //    }

    //    static void SetItemProperty(IItem item, string propertyId, int value)
    //    {
    //        var property = item.Properties[propertyId];
    //        if (property != null)
    //        {
    //            property.set_Value(value);
    //        }
    //        else
    //        {
    //            Console.WriteLine($"Property {propertyId} not found");
    //        }
    //    }

    //    static string GetFileSavePath(string extension)
    //    {
    //        // Define the folder path on the Desktop
    //        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
    //        string folderPath = Path.Combine(desktopPath, "ScannedDocuments");

    //        // Create the folder if it doesn't exist
    //        if (!Directory.Exists(folderPath))
    //        {
    //            Directory.CreateDirectory(folderPath);
    //        }

    //        // Generate a unique filename using the current date and time
    //        string uniqueFileName = DateTime.Now.ToString("yyyyMMdd_HHmmss") + extension;
    //        return Path.Combine(folderPath, uniqueFileName);
    //    }
    //}

    class Program
    {
        static List<Bitmap> ScannedImages = new List<Bitmap>();

        static void Main(string[] args)
        {
            Console.WriteLine("Scanner App - .NET 8");

            Console.WriteLine("\nPress any key to scan:");

            Console.ReadKey();

            while (true)
            {
                Console.WriteLine("Scanning Document...");
                ScanDocument();

                Console.WriteLine("Would you like to scan another document? (y/n): ");
                string userInput = Console.ReadLine();
                if (userInput?.ToLower() != "y")
                {
                    break;
                }
            }

            Console.WriteLine("What would you like to do with the scanned documents?");
            Console.WriteLine("1. Export as Photo (JPG/PNG)");
            Console.WriteLine("2. Export as PDF");
            Console.WriteLine("3. Export as Word");

            string choice = Console.ReadLine();

            try
            {
                DocumentExporter exporter = new DocumentExporter(ScannedImages);

                switch (choice)
                {
                    case "1":
                        exporter.ExportAsImage();
                        break;
                    case "2":
                        exporter.ExportAsPdf();
                        break;
                    case "3":
                        exporter.ExportAsWord();
                        break;
                    default:
                        Console.WriteLine("Invalid choice.");
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            Console.WriteLine("Operation completed.");
        }

        static void ScanDocument()
        {
            var deviceManager = new DeviceManager();

            if (deviceManager.DeviceInfos == null)
            {
                Console.WriteLine("DeviceInfos is null.");
                return;
            }
            Console.WriteLine($"DeviceInfos count: {deviceManager.DeviceInfos.Count}");
            if (deviceManager.DeviceInfos.Count == 0)
            {
                Console.WriteLine("No devices found. Please check if your scanner is connected and powered on, and that the WIA service is running.");
                Console.WriteLine("-----------------------------------");
                Console.WriteLine("Press any key if your are ready.");
                Console.WriteLine("-----------------------------------");

                Console.ReadKey();

                return;
            }

            List<DeviceInfo> scannerDevices = new List<DeviceInfo>();

            for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++)
            {
                if (deviceManager.DeviceInfos[i].Type == WiaDeviceType.ScannerDeviceType)
                {
                    scannerDevices.Add(deviceManager.DeviceInfos[i]);
                }
            }

            if (scannerDevices.Count == 0)
            {
                Console.WriteLine("No scanner devices found.");
                return;
            }

            Console.WriteLine("Available scanners: ");
            for (int i = 0; i < scannerDevices.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {scannerDevices[i].Properties["Name"].get_Value()}");
            }

            Console.Write("Select a scanner (enter the number): ");

            if (int.TryParse(Console.ReadLine(), out int selectedScannerIndex) && selectedScannerIndex > 0 && selectedScannerIndex <= scannerDevices.Count)
            {
                var deviceInfo = scannerDevices[selectedScannerIndex - 1];
                var device = deviceInfo.Connect();
                var item = device.Items[1];

                Console.Write("Enter horizontal DPI value (e.g., 300): ");
                int horizontalDpi = int.TryParse(Console.ReadLine(), out horizontalDpi) ? horizontalDpi : 300;

                Console.Write("Enter vertical DPI value (e.g., 300): ");
                int verticalDpi = int.TryParse(Console.ReadLine(), out verticalDpi) ? verticalDpi : 300;

                SetDpi(item, horizontalDpi, verticalDpi);

                var imgFile = (ImageFile)item.Transfer("{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}");

                var imageBytes = (byte[])imgFile.FileData.get_BinaryData();
                using (var ms = new MemoryStream(imageBytes))
                {
                    Bitmap scannedImage = new Bitmap(ms);
                    ScannedImages.Add(scannedImage);
                    Console.WriteLine("Document scanned successfully");
                }
            }
            else
            {
                Console.WriteLine("Invalid selection");
            }
        }

        static void SetDpi(IItem item, int horizontalDpi, int verticalDpi)
        {
            const string HorizontalResolution = "6147";
            const string VerticalResolution = "6148";

            SetItemProperty(item, HorizontalResolution, horizontalDpi);
            SetItemProperty(item, VerticalResolution, verticalDpi);
        }

        static void SetItemProperty(IItem item, string propertyId, int value)
        {
            var property = item.Properties[propertyId];
            if (property != null)
            {
                property.set_Value(value);
            }
            else
            {
                Console.WriteLine($"Property {propertyId} not found");
            }
        }
    }
}
