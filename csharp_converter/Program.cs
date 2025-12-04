using System;
using System.IO;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.Pdf;
using Syncfusion.Licensing;

namespace csharp_converter
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: csharp_converter <input_pptx> <output_pdf>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            // Register Syncfusion license
            SyncfusionLicenseProvider.RegisterLicense("NxYtFisQPR08Cit/Vkd+XU9FcVRDX3xKf0x/TGpQb19xflBPallYVBYiSV9jS3tSdkVmWXtbdXVWRWRcUE91Xg==");

            try
            {
                Console.WriteLine($"Converting {inputPath} to {outputPath}...");

                // Open the PowerPoint presentation
                using (FileStream fileStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
                {
                    using (IPresentation pptxDoc = Presentation.Open(fileStream))
                    {
                        // Initialize the PresentationToPdfConverter settings
                        PresentationToPdfConverterSettings settings = new PresentationToPdfConverterSettings();
                        settings.ShowHiddenSlides = true;

                        // Convert the PowerPoint presentation to PDF document
                        using (PdfDocument pdfDoc = PresentationToPdfConverter.Convert(pptxDoc, settings))
                        {
                            // Save the PDF document
                            using (FileStream outputStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                            {
                                pdfDoc.Save(outputStream);
                            }
                        }
                    }
                }

                Console.WriteLine("Conversion successful!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                Environment.Exit(1);
            }
        }
    }
}
