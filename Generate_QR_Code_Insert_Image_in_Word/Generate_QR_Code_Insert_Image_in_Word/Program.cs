using Aspose.BarCode;
using Aspose.BarCode.Generation;
using Aspose.Words;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace Generate_QR_Code_Insert_Image_in_Word
{
    class Program
    {
        static void Main(string[] args)
        {
            ApplyLicenses();

            Console.WriteLine("Program started");

            #region QR Code BMP Image Generation

            // Initialize a BarcodeGenerator class object and Set CodeText & Symbology Type
            BarcodeGenerator generator = new BarcodeGenerator(EncodeTypes.QR, "12345TEX");
            // Set ForceQR (default) for standard QR and Code text
            generator.Parameters.Barcode.QR.QrEncodeMode = QREncodeMode.Auto;
            generator.Parameters.Barcode.QR.QrEncodeType = QREncodeType.ForceQR;
            generator.Parameters.Barcode.QR.QrErrorLevel = QRErrorLevel.LevelL;
            // Get barcode image Bitmap and Save QR code
            Bitmap lBmp = generator.GenerateBarCodeImage();
            // Save to Stream in BMP format
            MemoryStream memoryStream = new MemoryStream();
            lBmp.Save(memoryStream, ImageFormat.Bmp);
            memoryStream.Position = 0;

            // Or Save to BMP File on Disk
            //lBmp.Save("image.bmp", ImageFormat.Bmp);

            #endregion

            #region Insert Image in Word & Convert to PDF

            // Load DOCX file you want to insert QR Code Image into
            Document doc = new Document("input.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd(); // or move cursor to any Node position
            // Insert QR Code Image from Memory Stream
            builder.InsertImage(memoryStream);
            // Save to PDF
            // doc.Save("output.docx");
            doc.Save("output.pdf");

            #endregion

            Console.WriteLine("Program ended");
            Console.WriteLine("press any key...");
            Console.ReadLine();
        }

        static void ApplyLicenses()
        {
            try
            {
                Aspose.Words.License license_Words = new Aspose.Words.License();
                license_Words.SetLicense("Aspose.Total.Product.Family.lic");
            }
            catch (Exception)
            {
                Console.WriteLine("Aspsoe.Words' license not applied");
            }

            try
            {
                Aspose.BarCode.License license_BarCode = new Aspose.BarCode.License();
                license_BarCode.SetLicense("Aspose.Total.Product.Family.lic");
            }
            catch (Exception)
            {
                Console.WriteLine("Aspsoe.BarCode's license not applied");
            }
        }
    }
}