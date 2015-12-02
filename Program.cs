using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Collections;
using System.Drawing;

namespace Downtime_PDF_Parser
{
    class Program
    {
        static void Main(string[] args)
        {
            // Read director with backup files and create a string array
            string[] downTimeFiles = Directory.GetFiles(@"\\UCHNETPRINT4\MEDITECH\DOWNTIME");

            // loop through the string array of directory files and find those needed
            foreach (string file in downTimeFiles)
            {
                if (file.EndsWith(".pdf"))
                {
                    //WaterMarker(file, @"c: \users\sprouse\desktop\test.pdf", "DOWNTIME USE");
                    ReadDowntimeFile(file);
                    File.Delete(file);
                }
                else
                {
                    Console.WriteLine("No downtime files to read in");
                }
            }

            GetForerunFiles(3);
            DeleteOldFile(14);         
        }

        static void WaterMarker(string inputFile, string outputFile, string waterMarkText)
        {
            PdfReader pdfReader = new PdfReader(inputFile);
            PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(outputFile, FileMode.Create));

            for (int pageIndex = 1; pageIndex <= pdfReader.NumberOfPages; pageIndex++)
            {
                //Rectangle class in iText represent geomatric representation... in this case, rectanle object would contain page geomatry
                iTextSharp.text.Rectangle pageRectangle = pdfReader.GetPageSizeWithRotation(pageIndex);
                //pdfcontentbyte object contains graphics and text content of page returned by pdfstamper
                PdfContentByte pdfData = pdfStamper.GetUnderContent(pageIndex);
                //create fontsize for watermark
                pdfData.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 60);
                //create new graphics state and assign opacity
                PdfGState graphicsState = new PdfGState();
                graphicsState.FillOpacity = 0.4F;
                //set graphics state to pdfcontentbyte
                pdfData.SetGState(graphicsState);
                //set color of watermark
                pdfData.SetColorFill(BaseColor.RED);
                //indicates start of writing of text
                pdfData.BeginText();
                //show text as per position and rotation
                pdfData.ShowTextAligned(Element.ALIGN_CENTER, waterMarkText, pageRectangle.Width / 2, pageRectangle.Height / 2, 45);
                //call endText to invalid font set
                pdfData.EndText();
            }
            //File.Delete(inputFile);
            pdfStamper.Close();
        }

        static void ReadDowntimeFile(string inputFile)
        {
            string endOfRecord = "** END OF RECORD **";

            int patient = 0;
            int startPage = 1;

            ArrayList patientList = new ArrayList();

            PdfReader reader = new PdfReader(inputFile);
            Console.WriteLine("Page count: " + reader.NumberOfPages);

            for (int page = 1; page <= reader.NumberOfPages; page++)
            {
                ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();

                string currentPageText = PdfTextExtractor.GetTextFromPage(reader, page, strategy);
                if (currentPageText.Contains(endOfRecord))
                {
                    Console.WriteLine(currentPageText);

                    var patientAccount = currentPageText.Substring(currentPageText.Length - Math.Min(12, currentPageText.Length));

                    patient++;
                    patientList.Add("Patient" + patient + "|" + startPage + "|" + page);

                    string outputFile = @"\\UCHNETPRINT4\MEDITECH\DOWNTIME\PATIENTS\" + patientAccount + "_SUMMARY.pdf";

                    ExtractPages(inputFile, outputFile, startPage, page);

                    startPage = page + 1;
                }
            }
            reader.Close();
            Console.WriteLine("number of patients: " + patient.ToString());
        }

        static void ExtractPages(string sourcePdfPath, string outputPdfPath, int startPage, int endPage)
        {
            PdfReader reader = null;
            Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage = null;

            try
            {
                // Intialize a new PdfReader instance with the contents of the source Pdf file:
                reader = new PdfReader(sourcePdfPath);

                // For simplicity, I am assuming all the pages share the same size
                // and rotation as the first page:
                sourceDocument = new Document(reader.GetPageSizeWithRotation(startPage));

                // Initialize an instance of the PdfCopyClass with the source 
                // document and an output file stream:
                pdfCopyProvider = new PdfCopy(sourceDocument, new FileStream(outputPdfPath, FileMode.Create));

                sourceDocument.Open();

                // Walk the specified range and add the page copies to the output file:
                for (int i = startPage; i <= endPage; i++)
                {
                    importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                    pdfCopyProvider.AddPage(importedPage);
                }

                sourceDocument.Close();
                reader.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        static void GetForerunFiles(int hours)
        {
            var files = new DirectoryInfo(@"\\upc-sca\FORERUNLIVEBKUP").GetFiles("*.tif");
            foreach (var file in files)
            {
                if (DateTime.UtcNow - file.CreationTimeUtc < TimeSpan.FromHours(hours))
                {
                    ConvertTifftoPdf(file.FullName);
                    Console.WriteLine(file.Name.Remove(0, 30).Remove(12));
                    //File.Copy(file.FullName, destPath + file.Name.Remove(0,30).Remove(12) + "_FORERUN.tif", true);
                }
            }
        }

        static void ConvertTifftoPdf(string fileName)
        {
            // creation of the document with a certain size and certain margins
            Document document = new Document(PageSize.A4, 0, 0, 0, 0);

            // creation of the different writers
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(@"\\uchnetprint4\meditech\downtime\patients\" + fileName.Remove(0, 56).Remove(12) + "_FORERUN.pdf", System.IO.FileMode.Create));

            // load the tiff image and count the total pages
            Bitmap bm = new Bitmap(fileName);
            int total = bm.GetFrameCount(System.Drawing.Imaging.FrameDimension.Page);

            document.Open();
            iTextSharp.text.pdf.PdfContentByte cb = writer.DirectContent;
            for (int k = 0; k < total; ++k)
            {
                bm.SelectActiveFrame(System.Drawing.Imaging.FrameDimension.Page, k);
                iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(bm, System.Drawing.Imaging.ImageFormat.Bmp);
                // scale the image to fit in the page
                img.ScalePercent(72f / img.DpiX * 100);
                img.SetAbsolutePosition(0, 0);
                cb.AddImage(img);
                document.NewPage();
            }
            document.Close();
        }

        static void DeleteOldFile(int days)
        {
            var files = new DirectoryInfo(@"\\UCHNETPRINT4\MEDITECH\DOWNTIME\PATIENTS").GetFiles("*.pdf");
            foreach (var file in files)
            {
                if (DateTime.UtcNow - file.CreationTimeUtc > TimeSpan.FromDays(days))
                {
                    Console.WriteLine(file.FullName);
                    File.Delete(file.FullName);
                }
            }
        }
    }
}
