using System;
using System.IO;
using System.Windows.Controls;
using Tesseract;
using MessageBox = System.Windows.MessageBox;

namespace PngToDocxConvertera.Services
{
    public class OcrService
    {
        public static string ExtractTextFromImage(string imagePath)
        {
            try
            {
                // Fix the path so it matches the absolute setup installation folder location
                string baseDir = AppDomain.CurrentDomain.BaseDirectory; 
                string tessdataPath = Path.Combine(baseDir, "tessdata");
                string trainedDataFile = Path.Combine(tessdataPath, "eng.traineddata");

                // Fix this line to point explicitly to the x64 subfolder
                string x64Path = Path.Combine(baseDir, "x64");
                string tessDll = Path.Combine(x64Path, "tesseract50.dll");
                string leptonicaDll = Path.Combine(x64Path, "leptonica-1.82.0.dll");

                if (!Directory.Exists(tessdataPath))
                {
                    MessageBox.Show("tessdata folder not found:\n\n" + tessdataPath, "OCR Error");
                    return string.Empty;
                }

                if (!File.Exists(trainedDataFile))
                {
                    MessageBox.Show("eng.traineddata missing:\n\n" + trainedDataFile, "OCR Error");
                    return string.Empty;
                }

                if (!File.Exists(tessDll))
                {
                    MessageBox.Show("Missing Tesseract DLL:\n\n" + tessDll, "OCR Error");
                    return string.Empty;
                }

                if (!File.Exists(leptonicaDll))
                {
                    MessageBox.Show("Missing Leptonica DLL:\n\n" + leptonicaDll, "OCR Error");
                    return string.Empty;
                }                

                using var engine = new TesseractEngine(
                    tessdataPath,
                    "eng",
                    EngineMode.Default
                );

                using var img = Pix.LoadFromFile(imagePath);
                using var page = engine.Process(img);

                return page.GetText();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.ToString(),
                    "OCR Preview Failed"
                );

                return string.Empty;
            }
        }
    }
}