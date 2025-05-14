using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using ZXing;
using ZXing.Common;
using ZXing.Multi;
using ZXing.Windows.Compatibility;
// Add ClosedXML namespace
using ClosedXML.Excel;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("PDF Barcode Scanner");
        Console.WriteLine("------------------");

        // Get directory from user or use default
        Console.Write("Enter the directory path containing PDF files (or press Enter for current directory): ");
        string directoryPath = Console.ReadLine();
        Console.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " start time");
        if (string.IsNullOrWhiteSpace(directoryPath))
            directoryPath = Directory.GetCurrentDirectory();

        if (!Directory.Exists(directoryPath))
        {
            Console.WriteLine($"Directory not found: {directoryPath}");
            return;
        }

        // Get all PDF files
        string[] pdfFiles = Directory.GetFiles(directoryPath, "*.pdf");

        if (pdfFiles.Length == 0)
        {
            Console.WriteLine("No PDF files found in the directory.");
            return;
        }

        Console.WriteLine($"Found {pdfFiles.Length} PDF files. Processing...");

        // Create a list to hold all barcode results with file information
        var allResults = new List<(string FileName, DateTime FileDate, BarcodeResult Result)>();

        // Process each PDF file
        for (int fileIndex = 0; fileIndex < pdfFiles.Length; fileIndex++)
        {
            string pdfFile = pdfFiles[fileIndex];
            string fileName = Path.GetFileName(pdfFile);
            DateTime fileDate = File.GetLastWriteTime(pdfFile);

            // Add file counter to console output (current/total)
            Console.WriteLine($"\nScanning file [{fileIndex + 1}/{pdfFiles.Length}]: {fileName}");
            var results = ScanPdfForBarcodes(pdfFile);

            if (results.Count == 0)
            {
                Console.WriteLine("  No barcodes found.");
            }
            else
            {
                Console.WriteLine($"  Found {results.Count} barcodes:");
                foreach (var result in results)
                {
                    Console.WriteLine($"  - Page {result.Page}: {result.BarcodeFormat} = {result.Text}");
                    // Add to the consolidated results with file information
                    allResults.Add((fileName, fileDate, result));
                }
            }
        }

        // Export results to Excel if any barcodes were found
        if (allResults.Count > 0)
        {
            string excelFilePath = ExportToExcel(allResults, directoryPath);
            Console.WriteLine($"\nResults exported to Excel file: {excelFilePath}");

            // Open the file location in Explorer
            OpenFileLocationInExplorer(excelFilePath);
        }
        else
        {
            Console.WriteLine("\nNo barcodes found in any files. Excel export skipped.");
        }
        Console.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " end time");

        Console.WriteLine("\nProcessing complete.");
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }

    static string ExportToExcel(List<(string FileName, DateTime FileDate, BarcodeResult Result)> results, string directoryPath)
    {
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        string fileName = $"BarcodeResults_{timestamp}.xlsx";
        string filePath = Path.Combine(directoryPath, fileName);

        // Create a new Excel workbook
        using (var workbook = new XLWorkbook())
        {
            // Add a worksheet
            var worksheet = workbook.Worksheets.Add("Barcode Results");

            // Add headers
            worksheet.Cell(1, 1).Value = "Barcode Text";
            worksheet.Cell(1, 2).Value = "File Date";
            worksheet.Cell(1, 3).Value = "File Name";
            worksheet.Cell(1, 4).Value = "Page Number";
            worksheet.Cell(1, 5).Value = "Barcode Format";

            // Style the header row
            var headerRow = worksheet.Row(1);
            headerRow.Style.Font.Bold = true;
            headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;
            headerRow.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Add data rows
            int row = 2;
            foreach (var item in results)
            {
                worksheet.Cell(row, 1).Value = item.Result.Text;
                worksheet.Cell(row, 2).Value = item.FileDate;
                worksheet.Cell(row, 2).Style.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
                worksheet.Cell(row, 3).Value = item.FileName;
                worksheet.Cell(row, 4).Value = item.Result.Page;
                worksheet.Cell(row, 5).Value = item.Result.BarcodeFormat.ToString();
                row++;
            }

            // Auto-adjust column widths
            worksheet.Columns().AdjustToContents();

            // Save the workbook
            workbook.SaveAs(filePath);
        }

        return filePath;
    }

    static void OpenFileLocationInExplorer(string filePath)
    {
        try
        {
            // Open the folder and select the file
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"/select,\"{filePath}\"",
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error opening file location: {ex.Message}");
        }
    }

    static List<BarcodeResult> ScanPdfForBarcodes(string pdfPath)
    {
        List<BarcodeResult> results = new List<BarcodeResult>();

        try
        {
            // Fix: Use the proper reader that can handle binary bitmaps
            var reader = new MultiFormatReader();

            var hints = new Dictionary<DecodeHintType, object>
            {
                { DecodeHintType.TRY_HARDER, true },
                { DecodeHintType.POSSIBLE_FORMATS, new List<BarcodeFormat>
                    {
                        BarcodeFormat.All_1D,
                        BarcodeFormat.QR_CODE,
                        BarcodeFormat.DATA_MATRIX
                    }
                }
            };

            // Open the PDF
            using (PdfDocument document = PdfDocument.Open(pdfPath))
            {
                // Process each page
                for (int i = 0; i < document.NumberOfPages; i++)
                {
                    Page page = document.GetPage(i + 1);

                    try
                    {
                        // Extract images from the page
                        var images = page.GetImages().ToList();

                        if (images.Count > 0)
                        {
                            foreach (var image in images)
                            {
                                try
                                {
                                    // Safely check if image is null or RawBytes is empty
                                    if (image == null || image.RawBytes == null || image.RawBytes.Length == 0)
                                        continue;
                                    if (image.RawBytes.ToArray().Length == 0)
                                        continue;

                                    // Convert image data to a memory stream
                                    using (var ms = new MemoryStream(image.RawBytes.ToArray()))
                                    {
                                        // Create a bitmap from the stream
                                        using (var bitmap = new Bitmap(ms))
                                        {
                                            // Fix: Use the correct conversion path
                                            var luminanceSource = new BitmapLuminanceSource(bitmap);
                                            var binaryBitmap = new BinaryBitmap(new HybridBinarizer(luminanceSource));

                                            try
                                            {
                                                // Fix: Use the MultiFormatReader directly with the binary bitmap
                                                var result = reader.decode(binaryBitmap, hints);
                                                if (result != null)
                                                    if (result.BarcodeFormat == BarcodeFormat.CODE_128)
                                                    {
                                                        {
                                                            results.Add(new BarcodeResult
                                                            {
                                                                Page = i + 1,
                                                                Text = result.Text,
                                                                BarcodeFormat = result.BarcodeFormat
                                                            });
                                                        }
                                                    }

                                                // Try to find multiple barcodes in one image
                                                try
                                                {
                                                    var multiReader = new GenericMultipleBarcodeReader(reader);
                                                    var multiResults = multiReader.decodeMultiple(binaryBitmap, hints);

                                                    if (multiResults != null && multiResults.Length > 0)
                                                    {
                                                        foreach (var barcode in multiResults)
                                                        {
                                                            // Only add if it's not a duplicate of what we already found
                                                            if (!results.Any(r => r.Page == i + 1 && r.Text == barcode.Text))
                                                            {
                                                                if (barcode.BarcodeFormat == BarcodeFormat.CODE_128)
                                                                {
                                                                    results.Add(new BarcodeResult
                                                                    {
                                                                        Page = i + 1,
                                                                        Text = barcode.Text,
                                                                        BarcodeFormat = barcode.BarcodeFormat
                                                                    });
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                catch (Exception)
                                                {
                                                    // Silently ignore errors in multiple barcode reading
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine($"  Error decoding barcode on page {i + 1}: {ex.Message}");
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    //Console.WriteLine($"  Error processing image on page first {i + 1}: {ex.Message}");
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine($"  No images found on page {i + 1}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  Error extracting images from page {i + 1}: {ex.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing PDF: {ex.Message}");
        }

        return results;
    }

    class BarcodeResult
    {
        public int Page { get; set; }
        public string Text { get; set; }
        public BarcodeFormat BarcodeFormat { get; set; }
    }
}
