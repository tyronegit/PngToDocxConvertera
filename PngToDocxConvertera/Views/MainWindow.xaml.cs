using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MigraDoc.Rendering;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PngToDocConvertera;
using PngToDocxConvertera.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Tesseract;



namespace PngToDocxConvertera.Views
{
    public partial class MainWindow : Window
    {
        private string selectedImagePath = "";

        private static readonly string[] SupportedImageExtensions =
        {
            ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"
        };

        private readonly TextCleanupService textCleanupService = new TextCleanupService();

        public MainWindow()
        {
            GlobalFontSettings.FontResolver ??= new WindowsFontResolver();

            InitializeComponent();

            ApplyThemeTextColors();
        }

        private void ApplyThemeTextColors()
        {
            Brush titleBlue = new SolidColorBrush(System.Windows.Media.Color.FromRgb(74, 165, 255));

            Brush black = new SolidColorBrush(System.Windows.Media.Color.FromRgb(18, 18, 18));

            TitleTextBlock.Foreground = titleBlue;
            SubtitleTextBlock.Foreground = titleBlue;

            MainMenu.Background = black;
            MainMenu.Foreground = titleBlue;
        }

        private static bool IsSupportedImageFile(string filePath)
        {
            string extension = Path.GetExtension(filePath);

            return SupportedImageExtensions.Contains(
                extension,
                StringComparer.OrdinalIgnoreCase);
        }

        private void SetStatus(string message)
        {
            StatusTextBlock.Text = message;
        }

        private void Window_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;

            e.Handled = true;
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
                return;

            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            if (files.Length == 0)
                return;

            string filePath = files[0];

            if (IsSupportedImageFile(filePath))
            {
                LoadImageFile(filePath);
            }
            else
            {
                System.Windows.MessageBox.Show(
                    "Please drop a supported image file: PNG, JPG, JPEG, BMP, or TIFF.",
                    "Unsupported File Format",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        // Closes the application completely
        private void MenuExit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        // Simple informational pop-up
        private void MenuAbout_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("PNG to DOCX Converter\nVersion 1.0\n\nA lightweight utility utilizing Tesseract OCR and OpenXML.",
                            "About Application",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
        }

        // Placeholders so the compiler builds cleanly
        private void MenuOpen_Click(object sender, RoutedEventArgs e)
        {
            // If you have a file opening method, you can trigger it here
        }

        private void MenuSaveAs_Click(object sender, RoutedEventArgs e)
        {
        }

        private void MenuClear_Click(object sender, RoutedEventArgs e)
        {
        }
        private void ClearWorkspaceMenuItem_Click(object sender, RoutedEventArgs e)
        {
            OcrPreviewTextBox.Clear();
            PreviewImage.Source = null;
            SelectedFileTextBox.Text = "";
            selectedImagePath = "";
            DropHintText.Visibility = Visibility.Visible;
            SetStatus("Workspace cleared.");
        }

        private void LoadImageFile(string filePath)
        {
            try
            {
                selectedImagePath = filePath;
                SelectedFileTextBox.Text = filePath;

                BitmapImage bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.UriSource = new Uri(filePath);
                bitmap.EndInit();

                PreviewImage.Source = bitmap;
                DropHintText.Visibility = Visibility.Collapsed;

                SetStatus("Image loaded successfully.");
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    "Unable to load image:\n\n" + ex.Message,
                    "Image Load Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void SelectImageButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select Image File",
                Filter = "Image Files (*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff)|*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff|All Files (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                LoadImageFile(openFileDialog.FileName);
            }
        }

        private string NormalizeBullets(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;

            string[] lines = text.Replace("\r\n", "\n").Split('\n');

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].TrimStart();

                if (line.StartsWith("•") ||
                    line.StartsWith("-") ||
                    line.StartsWith("*") ||
                    line.StartsWith("·") ||
                    line.StartsWith("o "))
                {
                    lines[i] = "• " + line.TrimStart('•', '-', '*', '·', 'o', ' ');
                }
            }

            return string.Join(Environment.NewLine, lines);
        }

        private void PreviewOcrButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(selectedImagePath) || !File.Exists(selectedImagePath))
            {
                System.Windows.MessageBox.Show(
                    "Please select or drop a supported image file first.",
                    "No Image Selected",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            try
            {
                string text = RunOcr(selectedImagePath);
                text = RemoveDuplicateLines(text);
                text = NormalizeBullets(text);
                text = textCleanupService.Clean(text);
                text = FixBulletRecognition(text);
                text = RemoveDuplicateLines(text);

                OcrPreviewTextBox.Text = text;
                SetStatus("OCR preview generated successfully.");
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    "Error running OCR processing:\n\n" + ex.ToString(),
                    "OCR Engine Failure",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                SetStatus("OCR preview failed.");
            }
        }

        private string FixBulletRecognition(string text)
        {
            string[] lines = text.Replace("\r\n", "\n").Split('\n');

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].TrimStart();

                if (line.StartsWith("e ") || line.StartsWith("o "))
                {
                    int firstSpace = line.IndexOf(' ');

                    if (firstSpace > 0)
                    {
                        lines[i] = "• " + line.Substring(firstSpace + 1).Trim();
                    }
                }
            }

            return string.Join(Environment.NewLine, lines);
        }

        private void SaveTxtButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(OcrPreviewTextBox.Text))
            {
                MessageBox.Show(
                    "No OCR text to save. Click 'Preview OCR' first.",
                    "Empty Text Warning",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            Microsoft.Win32.SaveFileDialog saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Save Text File",
                Filter = "Text File (*.txt)|*.txt",
                FileName = Path.GetFileNameWithoutExtension(selectedImagePath) + ".txt"
            };

            if (saveDialog.ShowDialog() == true)
            {
                File.WriteAllText(saveDialog.FileName, OcrPreviewTextBox.Text);
                SetStatus("TXT file created successfully.");
            }
        }

        private void SavePdfButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(OcrPreviewTextBox.Text))
            {
                MessageBox.Show(
                    "No OCR text to save. Click 'Preview OCR' first.",
                    "Empty Text Warning",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            Microsoft.Win32.SaveFileDialog saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Save PDF File",
                Filter = "PDF File (*.pdf)|*.pdf",
                FileName = Path.GetFileNameWithoutExtension(selectedImagePath) + ".pdf"
            };

            if (saveDialog.ShowDialog() == true)
            {
                CreateSimplePdf(saveDialog.FileName, OcrPreviewTextBox.Text);
                SetStatus("PDF file created successfully.");
            }
        }

        private void CreateSimplePdf(string outputPath, string text)
        {
            try
            {
                MigraDoc.DocumentObjectModel.Document document =
                    new MigraDoc.DocumentObjectModel.Document();

                document.Info.Title = "Image to DOCX Converter OCR Export";
                #pragma warning disable CS8602 // Dereference of a possibly null reference.
                document.Styles["Normal"].Font.Name = "Arial";
                #pragma warning restore CS8602 // Dereference of a possibly null reference.

                MigraDoc.DocumentObjectModel.Section section =
                    document.AddSection();

                MigraDoc.DocumentObjectModel.Paragraph title =
                    section.AddParagraph();

                title.AddText("Image to DOCX Converter - OCR Export");
                title.Format.Font.Name = "Arial";
                title.Format.Font.Size = 16;
                title.Format.Font.Bold = true;
                title.Format.SpaceAfter = "0.25in";

                foreach (string line in text.Replace("\r\n", "\n").Split('\n'))
                {
                    MigraDoc.DocumentObjectModel.Paragraph paragraph =
                        section.AddParagraph();

                    paragraph.AddText(line);
                    paragraph.Format.Font.Name = "Arial";
                    paragraph.Format.Font.Size = 11;
                    paragraph.Format.SpaceAfter = "0.05in";
                }

                MigraDoc.Rendering.PdfDocumentRenderer renderer = new MigraDoc.Rendering.PdfDocumentRenderer();

                renderer.Document = document;
                renderer.RenderDocument();
                renderer.PdfDocument.Save(outputPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "PDF creation failed:\n\n" + ex.Message,
                    "PDF Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(OcrPreviewTextBox.Text))
            {
                System.Windows.MessageBox.Show(
                    "No OCR text to save. Click 'Preview OCR' first to extract your text layout.",
                    "Empty Text Warning",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            Microsoft.Win32.SaveFileDialog saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Save Word Document",
                Filter = "Word Document (*.docx)|*.docx",
                FileName = Path.GetFileNameWithoutExtension(selectedImagePath) + ".docx"
            };

            if (saveDialog.ShowDialog() == true)
            {
                CreateDocx(saveDialog.FileName, OcrPreviewTextBox.Text);
                SetStatus("Word document created successfully.");
            }
        }

        private void BatchConvertButton_Click(object sender, RoutedEventArgs e)
        {
            using var folderDialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select a folder containing image files"
            };

            if (folderDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return;

            string folderPath = folderDialog.SelectedPath;

            var imageFiles = Directory
                .EnumerateFiles(folderPath)
                .Where(IsSupportedImageFile)
                .ToList();

            if (imageFiles.Count == 0)
            {
                System.Windows.MessageBox.Show(
                    "No supported image files were found in the selected folder.",
                    "No Images",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            string outputFolder = Path.Combine(folderPath, "Converted DOCX Files");
            Directory.CreateDirectory(outputFolder);

            int successCount = 0;
            int failCount = 0;

            foreach (string imageFile in imageFiles)
            {
                try
                {
                    string text = RunOcr(imageFile);
                    text = textCleanupService.Clean(text);

                    string outputPath = Path.Combine(
                        outputFolder,
                        Path.GetFileNameWithoutExtension(imageFile) + ".docx");

                    CreateDocx(outputPath, text);
                    successCount++;
                }
                catch
                {
                    failCount++;
                }
            }

            SetStatus($"Converted {successCount} image file(s). Failed: {failCount}. Output: {outputFolder}");
        }


        private string RunOcr(string imagePath)
        {
            string tessdataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");

            if (!Directory.Exists(tessdataPath))
                throw new DirectoryNotFoundException("Missing tessdata folder: " + tessdataPath);

            string trainedData = Path.Combine(tessdataPath, "eng.traineddata");

            if (!File.Exists(trainedData))
                throw new FileNotFoundException("Missing eng.traineddata file.", trainedData);

            using TesseractEngine engine = new TesseractEngine(tessdataPath, "eng", EngineMode.Default);
            using Pix img = Pix.LoadFromFile(imagePath);
            using Page page = engine.Process(img);

            return page.GetText();
        }

        private static bool LooksLikeBulletLine(string line)
        {
            string trimmed = line.TrimStart();

            return trimmed.StartsWith("• ") ||
                   trimmed.StartsWith("- ") ||
                   trimmed.StartsWith("* ") ||
                   trimmed.StartsWith("o ") ||
                   trimmed.StartsWith("e ");
        }

        private string RemoveDuplicateLines(string text)
        {
            var result = new List<string>();
            string previous = "";

            foreach (string line in text.Replace("\r\n", "\n").Split('\n'))
            {
                string current = line.Trim();

                if (!string.Equals(current, previous, StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(line);
                }

                previous = current;
            }

            return string.Join(Environment.NewLine, result);
        }


        private string RemoveDuplicateLinesAggressive(string text)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var result = new List<string>();

            foreach (string line in text.Replace("\r\n", "\n").Split('\n'))
            {
                string key = line.Trim();

                if (string.IsNullOrWhiteSpace(key))
                {
                    result.Add(line);
                    continue;
                }

                if (seen.Add(key))
                {
                    result.Add(line);
                }
            }

            return string.Join(Environment.NewLine, result);
        }

        private void AddImageToDocx(MainDocumentPart mainPart, Body body, string imagePath)
        {
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            string relationshipId = mainPart.GetIdOfPart(imagePart);

            DocumentFormat.OpenXml.Wordprocessing.Drawing drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent
                    {
                        Cx = 5000000,
                        Cy = 6500000
                    },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties
                    {
                        Id = 1U,
                        Name = "Source Image"
                    },
                    new DocumentFormat.OpenXml.Drawing.Graphic(
                        new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties
                                    {
                                        Id = 0U,
                                        Name = Path.GetFileName(imagePath)
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                                new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                    new DocumentFormat.OpenXml.Drawing.Blip
                                    {
                                        Embed = relationshipId
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Stretch(
                                        new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                                new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties()
                            )
                        )
                        {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                        }
                    )
                )
            );

            body.AppendChild(new Paragraph(new Run(drawing)));
        }

        private void CreateDocx(string outputPath, string text)
        {
            text = RemoveDuplicateLines(text);

            string templatePath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Templates",
                "blank Central.docx");

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException(
                    "Template file was not found.",
                    templatePath);
            }

            File.Copy(templatePath, outputPath, true);

            using WordprocessingDocument wordDocument =
                WordprocessingDocument.Open(outputPath, true);

            MainDocumentPart mainPart = wordDocument.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart is missing.");

            mainPart.Document ??= new Document();
            mainPart.Document.Body ??= new Body();

            Body body = mainPart.Document.Body;

            body.RemoveAllChildren<Paragraph>();            

            string[] lines = text.Replace("\r\n", "\n").Split('\n');

            foreach (string line in lines)
            {
                string trimmed = line.TrimStart();

                if (LooksLikeBulletLine(trimmed))
                {
                    trimmed = trimmed.TrimStart('•', '-', '*', '·', 'o', 'O', ' ');

                    Paragraph paragraph = new Paragraph(
                        new Run(
                            new Text("• " + trimmed)
                        )
                    );

                    body.Append(paragraph);
                }
                else
                {
                    Paragraph paragraph = new Paragraph(
                        new Run(
                            new Text(line)
                            {
                                Space = SpaceProcessingModeValues.Preserve
                            }
                        )
                    );

                    body.Append(paragraph);
                }
            }

            mainPart.Document.Save();
        }
    }
 }