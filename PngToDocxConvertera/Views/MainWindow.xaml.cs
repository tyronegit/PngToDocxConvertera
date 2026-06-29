using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MigraDoc.Rendering;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PngToDocConvertera;
using PngToDocxConvertera.Services;
using SixLabors.ImageSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Tesseract;



namespace PngToDocxConvertera.Views
{
    public partial class MainWindow : Window
    {
        private string selectedImagePath = "";
        private string selectedOutputBaseName = "ocr-export";
        private string pendingBatchFolderPath = "";
        private bool isBatchRunning;
        private CancellationTokenSource? batchCancellationTokenSource;

        private static readonly string[] SupportedImageExtensions =
        {
            ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"
        };

        private static readonly Dictionary<string, string> OcrLanguageNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["eng"] = "English",
            ["spa"] = "Spanish",
            ["fra"] = "French",
            ["deu"] = "German",
            ["por"] = "Portuguese",
            ["ita"] = "Italian"
        };

        private enum BatchOutputType
        {
            Docx,
            Txt,
            Pdf
        }

        private readonly TextCleanupService textCleanupService = new TextCleanupService();
        private readonly ImagePreprocessingService imagePreprocessingService = new ImagePreprocessingService();

        private sealed class BatchOptions
        {
            public required string InputFolder { get; init; }
            public required string OutputFolder { get; init; }
            public required string Extension { get; init; }
            public required string OcrLanguage { get; init; }
            public required BatchOutputType OutputType { get; init; }
            public required SearchOption SearchOption { get; init; }
            public required bool AutoFixImages { get; init; }
            public required bool SearchablePdf { get; init; }
            public required bool CombinePdf { get; init; }
        }

        private sealed class BatchProgressInfo
        {
            public required int CurrentFile { get; init; }
            public required int TotalFiles { get; init; }
            public required string FileName { get; init; }
            public TimeSpan? Elapsed { get; init; }
        }

        private sealed class BatchResult
        {
            public required int SuccessCount { get; init; }
            public required int FailCount { get; init; }
            public required int SkippedCount { get; init; }
            public required bool WasCanceled { get; init; }
            public required string OutputFolder { get; init; }
            public string ErrorLogPath { get; init; } = "";
        }

        public MainWindow()
        {
            GlobalFontSettings.FontResolver ??= new WindowsFontResolver();

            InitializeComponent();

            LoadAvailableOcrLanguages();

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

        private void LoadAvailableOcrLanguages()
        {
            OcrLanguageComboBox.Items.Clear();

            string tessdataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");

            if (!Directory.Exists(tessdataPath))
            {
                AddOcrLanguageItem("eng");
                OcrLanguageComboBox.SelectedIndex = 0;
                return;
            }

            var languageCodes = Directory
                .EnumerateFiles(tessdataPath, "*.traineddata")
                .Select(Path.GetFileNameWithoutExtension)
                .Where(code => !string.IsNullOrWhiteSpace(code))
                .Select(code => code!)
                .OrderBy(code => GetOcrLanguageDisplayName(code), StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (string languageCode in languageCodes)
            {
                AddOcrLanguageItem(languageCode);
            }

            if (OcrLanguageComboBox.Items.Count == 0)
                AddOcrLanguageItem("eng");

            SelectOcrLanguage("eng");
        }

        private void AddOcrLanguageItem(string languageCode)
        {
            string displayName = GetOcrLanguageDisplayName(languageCode);

            OcrLanguageComboBox.Items.Add(new ComboBoxItem
            {
                Content = $"{displayName} ({languageCode})",
                Tag = languageCode,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(18, 18, 18)),
                Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(74, 165, 255)),
                FontWeight = FontWeights.SemiBold
            });
        }

        private static string GetOcrLanguageDisplayName(string languageCode)
        {
            return OcrLanguageNames.TryGetValue(languageCode, out string? displayName)
                ? displayName
                : languageCode;
        }

        private void SelectOcrLanguage(string languageCode)
        {
            foreach (object item in OcrLanguageComboBox.Items)
            {
                if (item is ComboBoxItem comboBoxItem &&
                    comboBoxItem.Tag is string itemLanguageCode &&
                    string.Equals(itemLanguageCode, languageCode, StringComparison.OrdinalIgnoreCase))
                {
                    OcrLanguageComboBox.SelectedItem = comboBoxItem;
                    return;
                }
            }

            OcrLanguageComboBox.SelectedIndex = 0;
        }

        private string GetSelectedOcrLanguage()
        {
            if (OcrLanguageComboBox.SelectedItem is ComboBoxItem item &&
                item.Tag is string language &&
                !string.IsNullOrWhiteSpace(language))
            {
                return language;
            }

            return "eng";
        }

        private string GetSelectedOcrLanguageDisplayName()
        {
            return GetOcrLanguageDisplayName(GetSelectedOcrLanguage());
        }

        private bool IsAutoFixEnabled()
        {
            return AutoFixCheckBox.IsChecked == true;
        }

        private bool IsSearchablePdfEnabled()
        {
            return SearchablePdfCheckBox.IsChecked == true;
        }

        private bool IsCombinePdfEnabled()
        {
            return CombinePdfCheckBox.IsChecked == true;
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

            if (Directory.Exists(filePath))
            {
                pendingBatchFolderPath = filePath;
                SelectedFileTextBox.Text = filePath;
                DropHintText.Visibility = Visibility.Visible;
                SetStatus("Folder ready for batch conversion. Choose DOCX, TXT, or PDF.");
            }
            else if (IsSupportedImageFile(filePath))
            {
                LoadImageFile(filePath);
            }
            else
            {
                System.Windows.MessageBox.Show(
                    "Please drop a supported image file or a folder containing images.",
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
            MessageBox.Show("Image to DOCX Converter Pro\nVersion 2.2\n\nOffline OCR for DOCX, TXT, and searchable PDF exports.\nPowered by Tesseract OCR and OpenXML.",
                            "About Application",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
        }

        private void MenuUserGuide_Click(object sender, RoutedEventArgs e)
        {
            string guideText =
                "Image to DOCX Converter Pro - User Guide" + Environment.NewLine +
                Environment.NewLine +
                "Single image workflow:" + Environment.NewLine +
                "1. Click Select Image File, or drag an image into the preview area." + Environment.NewLine +
                "2. Choose the OCR Language on Image that matches the text in the image. This recognizes text; it does not translate it." + Environment.NewLine +
                "3. Leave Auto-fix image for OCR checked for scans, screenshots, and low-contrast images." + Environment.NewLine +
                "4. Click Preview OCR, edit the text if needed, then save as DOCX, TXT, or PDF." + Environment.NewLine +
                "5. Keep Create searchable image PDFs checked when you want a PDF that looks like the original image but can be searched and copied." + Environment.NewLine +
                Environment.NewLine +
                "Batch workflow:" + Environment.NewLine +
                "1. Click Convert Image Folder to DOCX, TXT, or PDF." + Environment.NewLine +
                "2. Select the folder that contains image files." + Environment.NewLine +
                "3. Check Include subfolders in batch when images are inside nested folders." + Environment.NewLine +
                "4. Check Choose output folder for batches if you want to pick a custom destination." + Environment.NewLine +
                "5. Check Combine batch PDFs into one file when you want a single searchable PDF package instead of one PDF per image." + Environment.NewLine +
                "6. Watch the Progress bar and Batch counter at the bottom of the window." + Environment.NewLine +
                "7. Click Cancel Batch to stop after the current file finishes." + Environment.NewLine +
                Environment.NewLine +
                "PDF export modes:" + Environment.NewLine +
                "- Searchable image PDF preserves the source image and adds a hidden OCR text layer for search/copy workflows." + Environment.NewLine +
                "- Text PDF creates a clean text-only PDF from the OCR result." + Environment.NewLine +
                "- Combined batch PDF creates one multi-page PDF in the output folder." + Environment.NewLine +
                Environment.NewLine +
                "Output and errors:" + Environment.NewLine +
                "- If an output file already exists, the app creates a safe name such as file_1.txt." + Environment.NewLine +
                "- Failed batch files are listed in conversion-errors.txt inside the output folder." + Environment.NewLine +
                "- When a batch finishes, the summary can open the output folder for you." + Environment.NewLine +
                Environment.NewLine +
                "OCR languages:" + Environment.NewLine +
                "- The selected language is the language printed in the image, not the output language." + Environment.NewLine +
                "- The selected language requires a matching tessdata file in the app's tessdata folder." + Environment.NewLine +
                "- Example: Spanish requires spa.traineddata, French requires fra.traineddata." + Environment.NewLine +
                "- This build supports English, Spanish, French, German, Portuguese, and Italian when the matching files are included." + Environment.NewLine +
                Environment.NewLine +
                "Supported language files:" + Environment.NewLine +
                "- English: eng.traineddata" + Environment.NewLine +
                "- Spanish: spa.traineddata" + Environment.NewLine +
                "- French: fra.traineddata" + Environment.NewLine +
                "- German: deu.traineddata" + Environment.NewLine +
                "- Portuguese: por.traineddata" + Environment.NewLine +
                "- Italian: ita.traineddata";

            Window guideWindow = new Window
            {
                Title = "User Guide",
                Owner = this,
                Width = 680,
                Height = 560,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(18, 18, 18))
            };

            TextBox guideTextBox = new TextBox
            {
                Text = guideText,
                IsReadOnly = true,
                TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(18, 18, 18)),
                Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(230, 230, 230)),
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(74, 85, 104)),
                Padding = new Thickness(16),
                FontSize = 14
            };

            guideWindow.Content = guideTextBox;
            guideWindow.ShowDialog();
        }

        private void MenuOpen_Click(object sender, RoutedEventArgs e)
        {
            SelectImageButton_Click(sender, e);
        }

        private void MenuSaveAs_Click(object sender, RoutedEventArgs e)
        {
            ConvertButton_Click(sender, e);
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
            selectedOutputBaseName = "ocr-export";
            pendingBatchFolderPath = "";
            DropHintText.Visibility = Visibility.Visible;
            BatchProgressBar.Value = 0;
            BatchCounterTextBlock.Text = "Batch: idle";
            CancelBatchButton.Visibility = Visibility.Collapsed;
            SetStatus("Workspace cleared.");
        }

        private void LoadImageFile(string filePath)
        {
            try
            {
                selectedImagePath = filePath;
                selectedOutputBaseName = GetStableOutputBaseName(filePath);
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

        private static string GetStableOutputBaseName(string sourcePath)
        {
            string baseName = Path.GetFileNameWithoutExtension(sourcePath);

            if (string.IsNullOrWhiteSpace(baseName))
                return "ocr-export";

            foreach (char invalidChar in Path.GetInvalidFileNameChars())
            {
                baseName = baseName.Replace(invalidChar, '_');
            }

            return string.IsNullOrWhiteSpace(baseName)
                ? "ocr-export"
                : baseName.Trim();
        }

        private Microsoft.Win32.SaveFileDialog CreateStableSaveDialog(string title, string filter, string extension)
        {
            string normalizedExtension = extension.StartsWith(".", StringComparison.Ordinal)
                ? extension
                : "." + extension;

            var saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = title,
                Filter = filter,
                DefaultExt = normalizedExtension,
                AddExtension = true,
                OverwritePrompt = true,
                FileName = selectedOutputBaseName + normalizedExtension
            };

            saveDialog.FileOk += (_, _) =>
            {
                saveDialog.FileName = GetStableSavePath(saveDialog.FileName, normalizedExtension);
            };

            if (!string.IsNullOrWhiteSpace(selectedImagePath))
            {
                string? sourceFolder = Path.GetDirectoryName(selectedImagePath);
                if (!string.IsNullOrWhiteSpace(sourceFolder) && Directory.Exists(sourceFolder))
                    saveDialog.InitialDirectory = sourceFolder;
            }

            return saveDialog;
        }

        private string GetStableSavePath(string selectedPath, string extension)
        {
            string normalizedExtension = extension.StartsWith(".", StringComparison.Ordinal)
                ? extension
                : "." + extension;

            string targetFolder;

            if (!string.IsNullOrWhiteSpace(selectedPath) && Directory.Exists(selectedPath))
            {
                targetFolder = selectedPath;
            }
            else
            {
                targetFolder = Path.GetDirectoryName(selectedPath) ?? "";

                if (string.IsNullOrWhiteSpace(targetFolder) || !Directory.Exists(targetFolder))
                {
                    targetFolder = !string.IsNullOrWhiteSpace(selectedImagePath)
                        ? Path.GetDirectoryName(selectedImagePath) ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                        : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                }
            }

            return Path.Combine(targetFolder, selectedOutputBaseName + normalizedExtension);
        }

        private string? ChooseStableSingleOutputPath(string dialogTitle, string extension)
        {
            string normalizedExtension = extension.StartsWith(".", StringComparison.Ordinal)
                ? extension
                : "." + extension;

            string initialFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (!string.IsNullOrWhiteSpace(selectedImagePath))
            {
                string? sourceFolder = Path.GetDirectoryName(selectedImagePath);
                if (!string.IsNullOrWhiteSpace(sourceFolder) && Directory.Exists(sourceFolder))
                    initialFolder = sourceFolder;
            }

            using var folderDialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = dialogTitle,
                SelectedPath = initialFolder,
                UseDescriptionForTitle = true
            };

            if (folderDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return null;

            string outputPath = Path.Combine(
                folderDialog.SelectedPath,
                selectedOutputBaseName + normalizedExtension);

            if (File.Exists(outputPath))
            {
                MessageBoxResult overwrite = MessageBox.Show(
                    "This file already exists:\n\n" + outputPath + "\n\nReplace it?",
                    "Replace Existing File",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (overwrite != MessageBoxResult.Yes)
                    return null;
            }

            return outputPath;
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
                text = CleanOcrText(text);

                OcrPreviewTextBox.Text = text;
                SetStatus($"OCR preview generated using {GetSelectedOcrLanguageDisplayName()} recognition.");
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

            string? outputPath = ChooseStableSingleOutputPath(
                "Select a folder for the TXT file",
                ".txt");

            if (!string.IsNullOrWhiteSpace(outputPath))
            {
                File.WriteAllText(outputPath, OcrPreviewTextBox.Text);
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

            string? outputPath = ChooseStableSingleOutputPath(
                "Select a folder for the PDF file",
                ".pdf");

            if (!string.IsNullOrWhiteSpace(outputPath))
            {
                if (IsSearchablePdfEnabled())
                {
                    CreateSearchableImagePdf(outputPath, selectedImagePath, OcrPreviewTextBox.Text);
                    SetStatus("Searchable image PDF created successfully.");
                }
                else
                {
                    CreateSimplePdf(outputPath, OcrPreviewTextBox.Text);
                    SetStatus("Text PDF file created successfully.");
                }
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

        private void CreateSearchableImagePdf(string outputPath, string imagePath, string text)
        {
            try
            {
                SaveSearchableImagePdf(outputPath, imagePath, text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Searchable PDF creation failed:\n\n" + ex.Message,
                    "PDF Error",
                    MessageBoxButton.OK,
                MessageBoxImage.Error);
            }
        }

        private void SaveSearchableImagePdf(string outputPath, string imagePath, string text)
        {
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Searchable OCR PDF";
            AddSearchableImagePdfPage(document, imagePath, text);
            document.Save(outputPath);
            document.Close();
        }

        private void AddSearchableImagePdfPage(PdfDocument document, string imagePath, string text)
        {
            string pdfImagePath = CreatePdfCompatibleImageCopy(imagePath);

            try
            {
                using XImage sourceImage = XImage.FromFile(pdfImagePath);

                PdfPage page = document.AddPage();
                page.Width = XUnit.FromPoint(sourceImage.PointWidth);
                page.Height = XUnit.FromPoint(sourceImage.PointHeight);

                using XGraphics graphics = XGraphics.FromPdfPage(page);
                graphics.DrawImage(
                    sourceImage,
                    0,
                    0,
                    page.Width.Point,
                    page.Height.Point);

                AddSearchTextLayer(graphics, text, page.Width.Point, page.Height.Point);
            }
            finally
            {
                TryDeleteTemporaryFile(pdfImagePath);
            }
        }

        private static void AddTextPdfPage(PdfDocument document, string title, string text)
        {
            PdfPage page = document.AddPage();
            page.Size = PdfSharp.PageSize.Letter;

            XGraphics graphics = XGraphics.FromPdfPage(page);
            XFont titleFont = new XFont("Arial", 14, XFontStyleEx.Bold);
            XFont bodyFont = new XFont("Arial", 10, XFontStyleEx.Regular);
            XBrush brush = XBrushes.Black;

            const double margin = 42;
            double y = margin;
            double contentWidth = page.Width.Point - (margin * 2);

            try
            {
                graphics.DrawString(title, titleFont, brush, new XRect(margin, y, contentWidth, 24), XStringFormats.TopLeft);
                y += 32;

                foreach (string line in WrapPdfText(text, 95))
                {
                    if (y > page.Height.Point - margin)
                    {
                        graphics.Dispose();
                        page = document.AddPage();
                        page.Size = PdfSharp.PageSize.Letter;
                        graphics = XGraphics.FromPdfPage(page);
                        y = margin;
                        contentWidth = page.Width.Point - (margin * 2);
                    }

                    graphics.DrawString(line, bodyFont, brush, new XRect(margin, y, contentWidth, 14), XStringFormats.TopLeft);
                    y += 14;
                }
            }
            finally
            {
                graphics.Dispose();
            }
        }

        private static void AddSearchTextLayer(XGraphics graphics, string text, double pageWidth, double pageHeight)
        {
            XFont searchFont = new XFont("Arial", 7, XFontStyleEx.Regular);
            XBrush nearlyInvisibleBrush = new XSolidBrush(XColor.FromArgb(1, 0, 0, 0));
            const double margin = 24;
            double y = margin;
            double contentWidth = Math.Max(50, pageWidth - (margin * 2));

            foreach (string line in WrapPdfText(text, 115))
            {
                if (y > pageHeight - margin)
                    break;

                graphics.DrawString(
                    line,
                    searchFont,
                    nearlyInvisibleBrush,
                    new XRect(margin, y, contentWidth, 10),
                    XStringFormats.TopLeft);

                y += 10;
            }
        }

        private static IEnumerable<string> WrapPdfText(string text, int maxCharactersPerLine)
        {
            foreach (string sourceLine in text.Replace("\r\n", "\n").Split('\n'))
            {
                string remaining = sourceLine.TrimEnd();

                if (remaining.Length == 0)
                {
                    yield return "";
                    continue;
                }

                while (remaining.Length > maxCharactersPerLine)
                {
                    int splitAt = remaining.LastIndexOf(' ', maxCharactersPerLine);
                    if (splitAt < 25)
                        splitAt = maxCharactersPerLine;

                    yield return remaining.Substring(0, splitAt).TrimEnd();
                    remaining = remaining.Substring(splitAt).TrimStart();
                }

                yield return remaining;
            }
        }

        private static string CreatePdfCompatibleImageCopy(string imagePath)
        {
            string tempPath = Path.Combine(
                Path.GetTempPath(),
                Path.GetFileNameWithoutExtension(imagePath) + "_pdf_" + Guid.NewGuid().ToString("N") + ".png");

            using SixLabors.ImageSharp.Image image = SixLabors.ImageSharp.Image.Load(imagePath);
            image.SaveAsPng(tempPath);

            return tempPath;
        }

        private static void TryDeleteTemporaryFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                    File.Delete(filePath);
            }
            catch
            {
                // Temporary export files can be cleaned up by Windows later if still locked.
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

            string? outputPath = ChooseStableSingleOutputPath(
                "Select a folder for the Word document",
                ".docx");

            if (!string.IsNullOrWhiteSpace(outputPath))
            {
                CreateDocx(outputPath, OcrPreviewTextBox.Text);
                SetStatus("Word document created successfully.");
            }
        }

        private void BatchDocxButton_Click(object sender, RoutedEventArgs e)
        {
            ConvertImageFolder(BatchOutputType.Docx);
        }

        private void BatchTxtButton_Click(object sender, RoutedEventArgs e)
        {
            ConvertImageFolder(BatchOutputType.Txt);
        }

        private void BatchPdfButton_Click(object sender, RoutedEventArgs e)
        {
            ConvertImageFolder(BatchOutputType.Pdf);
        }

        private void CancelBatchButton_Click(object sender, RoutedEventArgs e)
        {
            batchCancellationTokenSource?.Cancel();
            SetStatus("Cancel requested. Finishing the current file...");
        }

        private async void ConvertImageFolder(BatchOutputType outputType)
        {
            if (isBatchRunning)
            {
                SetStatus("A batch conversion is already running.");
                return;
            }

            string? folderPath = ChooseBatchInputFolder();
            if (string.IsNullOrWhiteSpace(folderPath))
                return;

            SearchOption searchOption = IncludeSubfoldersCheckBox.IsChecked == true
                ? SearchOption.AllDirectories
                : SearchOption.TopDirectoryOnly;

            var imageFiles = Directory
                .EnumerateFiles(folderPath, "*.*", searchOption)
                .Where(IsSupportedImageFile)
                .OrderBy(file => file, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (imageFiles.Count == 0)
            {
                System.Windows.MessageBox.Show("No supported image files were found.");
                return;
            }

            string extension = outputType switch
            {
                BatchOutputType.Docx => ".docx",
                BatchOutputType.Txt => ".txt",
                BatchOutputType.Pdf => ".pdf",
                _ => ".docx"
            };

            string? outputFolder = ChooseBatchOutputFolder(folderPath, extension);
            if (string.IsNullOrWhiteSpace(outputFolder))
                return;

            Directory.CreateDirectory(outputFolder);

            BatchOptions options = new BatchOptions
            {
                InputFolder = folderPath,
                OutputFolder = outputFolder,
                Extension = extension,
                OcrLanguage = GetSelectedOcrLanguage(),
                OutputType = outputType,
                SearchOption = searchOption,
                AutoFixImages = IsAutoFixEnabled(),
                SearchablePdf = IsSearchablePdfEnabled(),
                CombinePdf = outputType == BatchOutputType.Pdf && IsCombinePdfEnabled()
            };

            var progress = new Progress<BatchProgressInfo>(UpdateBatchProgress);
            batchCancellationTokenSource = new CancellationTokenSource();

            isBatchRunning = true;
            SetBatchControlsEnabled(false);
            BatchProgressBar.Maximum = 100;
            BatchProgressBar.Value = 0;
            BatchCounterTextBlock.Text = $"Batch: 0 of {imageFiles.Count} (0%)";
            string pdfMode = options.OutputType == BatchOutputType.Pdf && options.SearchablePdf
                ? " searchable image PDF"
                : "";

            SetStatus($"Starting {imageFiles.Count} file(s) to{pdfMode} {extension} using {GetOcrLanguageDisplayName(options.OcrLanguage)} recognition...");

            try
            {
                BatchResult result = await Task.Run(() =>
                    ConvertImageFolderCore(
                        imageFiles,
                        options,
                        progress,
                        batchCancellationTokenSource.Token));

                ShowBatchSummary(result);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Batch conversion failed:\n\n" + ex.Message,
                    "Batch Conversion Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                SetStatus("Batch conversion failed.");
            }
            finally
            {
                SetBatchControlsEnabled(true);
                isBatchRunning = false;
                batchCancellationTokenSource.Dispose();
                batchCancellationTokenSource = null;
            }
        }

        private string? ChooseBatchInputFolder()
        {
            if (!string.IsNullOrWhiteSpace(pendingBatchFolderPath) &&
                Directory.Exists(pendingBatchFolderPath))
            {
                MessageBoxResult useDroppedFolder = MessageBox.Show(
                    "Use the dropped folder for this batch?\n\n" + pendingBatchFolderPath,
                    "Use Dropped Folder",
                    MessageBoxButton.YesNoCancel,
                    MessageBoxImage.Question);

                if (useDroppedFolder == MessageBoxResult.Yes)
                    return pendingBatchFolderPath;

                if (useDroppedFolder == MessageBoxResult.Cancel)
                    return null;
            }

            using var folderDialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select a folder containing image files"
            };

            if (folderDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return null;

            pendingBatchFolderPath = folderDialog.SelectedPath;
            return folderDialog.SelectedPath;
        }

        private string? ChooseBatchOutputFolder(string inputFolder, string extension)
        {
            string defaultOutputFolder = Path.Combine(
                inputFolder,
                $"Converted {extension.TrimStart('.').ToUpper()} Files");

            if (AskOutputFolderCheckBox.IsChecked != true)
                return defaultOutputFolder;

            Directory.CreateDirectory(defaultOutputFolder);

            using var folderDialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select a folder for converted files",
                SelectedPath = defaultOutputFolder
            };

            if (folderDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return null;

            return folderDialog.SelectedPath;
        }

        private BatchResult ConvertImageFolderCore(
            IReadOnlyList<string> imageFiles,
            BatchOptions options,
            IProgress<BatchProgressInfo> progress,
            CancellationToken cancellationToken)
        {
            int successCount = 0;
            int failCount = 0;
            int skippedCount = 0;
            var errors = new List<string>();

            using TesseractEngine engine = CreateOcrEngine(options.OcrLanguage);
            PdfDocument? combinedPdf = null;
            string combinedPdfPath = "";

            if (options.OutputType == BatchOutputType.Pdf && options.CombinePdf)
            {
                combinedPdf = new PdfDocument();
                combinedPdf.Info.Title = options.SearchablePdf
                    ? "Combined Searchable OCR PDF"
                    : "Combined OCR Text PDF";

                combinedPdfPath = GetUniqueOutputPath(
                    options.OutputFolder,
                    options.SearchablePdf ? "Combined_Searchable_OCR" : "Combined_OCR_Text",
                    ".pdf");
            }

            foreach (string imageFile in imageFiles)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    skippedCount = imageFiles.Count - successCount - failCount;
                    break;
                }

                int currentFile = successCount + failCount + skippedCount + 1;
                Stopwatch stopwatch = Stopwatch.StartNew();

                progress.Report(new BatchProgressInfo
                {
                    CurrentFile = currentFile,
                    TotalFiles = imageFiles.Count,
                    FileName = Path.GetFileName(imageFile)
                });

                try
                {
                    string text = RunOcr(imageFile, engine, options.AutoFixImages);
                    text = CleanOcrText(text);

                    string outputPath = GetUniqueOutputPath(
                        options.OutputFolder,
                        Path.GetFileNameWithoutExtension(imageFile),
                        options.Extension);

                    if (cancellationToken.IsCancellationRequested)
                    {
                        skippedCount = imageFiles.Count - successCount - failCount;
                        break;
                    }

                    if (options.OutputType == BatchOutputType.Docx)
                        CreateDocx(outputPath, text);
                    else if (options.OutputType == BatchOutputType.Txt)
                        File.WriteAllText(outputPath, text);
                    else if (options.OutputType == BatchOutputType.Pdf)
                    {
                        if (combinedPdf != null)
                        {
                            if (options.SearchablePdf)
                                AddSearchableImagePdfPage(combinedPdf, imageFile, text);
                            else
                                AddTextPdfPage(combinedPdf, Path.GetFileName(imageFile), text);
                        }
                        else if (options.SearchablePdf)
                        {
                            SaveSearchableImagePdf(outputPath, imageFile, text);
                        }
                        else
                        {
                            CreateSimplePdf(outputPath, text);
                        }
                    }

                    successCount++;
                }
                catch (Exception ex)
                {
                    failCount++;
                    errors.Add($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | FAILED | {imageFile}");
                    errors.Add("Reason: " + ex.Message);
                    errors.Add("");
                }
                finally
                {
                    stopwatch.Stop();
                    progress.Report(new BatchProgressInfo
                    {
                        CurrentFile = currentFile,
                        TotalFiles = imageFiles.Count,
                        FileName = Path.GetFileName(imageFile),
                        Elapsed = stopwatch.Elapsed
                    });
                }
            }

            if (combinedPdf != null && successCount > 0)
            {
                combinedPdf.Save(combinedPdfPath);
                combinedPdf.Close();
            }

            string errorLogPath = WriteBatchErrorLog(options.OutputFolder, errors);

            return new BatchResult
            {
                SuccessCount = successCount,
                FailCount = failCount,
                SkippedCount = skippedCount,
                WasCanceled = cancellationToken.IsCancellationRequested,
                OutputFolder = options.OutputFolder,
                ErrorLogPath = errorLogPath
            };
        }

        private void UpdateBatchProgress(BatchProgressInfo progress)
        {
            double percent = progress.TotalFiles == 0
                ? 0
                : Math.Min(100, Math.Max(0, progress.CurrentFile * 100.0 / progress.TotalFiles));

            BatchProgressBar.Maximum = 100;
            BatchProgressBar.Value = percent;
            BatchCounterTextBlock.Text = $"Batch: {progress.CurrentFile} of {progress.TotalFiles} ({percent:0}%)";

            string elapsedText = progress.Elapsed.HasValue
                ? $" ({progress.Elapsed.Value.TotalSeconds:0.0}s)"
                : "";

            SetStatus($"Converting {progress.CurrentFile} of {progress.TotalFiles}: {progress.FileName}{elapsedText}");
        }

        private void ShowBatchSummary(BatchResult result)
        {
            string statusPrefix = result.WasCanceled ? "Batch canceled." : "Batch completed.";
            SetStatus($"{statusPrefix} Converted: {result.SuccessCount}. Failed: {result.FailCount}. Skipped: {result.SkippedCount}. Output: {result.OutputFolder}");
            BatchCounterTextBlock.Text = result.WasCanceled ? "Batch: canceled" : "Batch: complete";
            BatchProgressBar.Value = result.WasCanceled ? BatchProgressBar.Value : 100;

            string errorLogText = string.IsNullOrWhiteSpace(result.ErrorLogPath)
                ? ""
                : "\n\nError log:\n" + result.ErrorLogPath;

            MessageBoxResult openFolder = MessageBox.Show(
                $"{statusPrefix}\n\nConverted: {result.SuccessCount}\nFailed: {result.FailCount}\nSkipped: {result.SkippedCount}\n\nOutput folder:\n{result.OutputFolder}{errorLogText}\n\nOpen output folder?",
                "Batch Conversion Summary",
                MessageBoxButton.YesNo,
                MessageBoxImage.Information);

            if (openFolder == MessageBoxResult.Yes)
                OpenFolder(result.OutputFolder);
        }

        private static string GetUniqueOutputPath(string outputFolder, string baseFileName, string extension)
        {
            string outputPath = Path.Combine(outputFolder, baseFileName + extension);

            if (!File.Exists(outputPath))
                return outputPath;

            int suffix = 1;
            string candidatePath;

            do
            {
                candidatePath = Path.Combine(outputFolder, $"{baseFileName}_{suffix}{extension}");
                suffix++;
            }
            while (File.Exists(candidatePath));

            return candidatePath;
        }

        private static string WriteBatchErrorLog(string outputFolder, IReadOnlyList<string> errors)
        {
            if (errors.Count == 0)
                return "";

            string errorLogPath = Path.Combine(outputFolder, "conversion-errors.txt");

            var logBuilder = new StringBuilder();
            logBuilder.AppendLine("Image to DOCX Converter - Batch Error Log");
            logBuilder.AppendLine("Created: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            logBuilder.AppendLine();

            foreach (string error in errors)
                logBuilder.AppendLine(error);

            File.WriteAllText(errorLogPath, logBuilder.ToString());

            return errorLogPath;
        }

        private static void OpenFolder(string folderPath)
        {
            if (!Directory.Exists(folderPath))
                return;

            Process.Start(new ProcessStartInfo
            {
                FileName = folderPath,
                UseShellExecute = true
            });
        }

        private void SetBatchControlsEnabled(bool isEnabled)
        {
            BatchDocxButton.IsEnabled = isEnabled;
            BatchTxtButton.IsEnabled = isEnabled;
            BatchPdfButton.IsEnabled = isEnabled;
            CancelBatchButton.IsEnabled = !isEnabled;
            CancelBatchButton.Visibility = isEnabled ? Visibility.Collapsed : Visibility.Visible;
            SelectImageButton.IsEnabled = isEnabled;
            PreviewOcrButton.IsEnabled = isEnabled;
            ConvertSingleButton.IsEnabled = isEnabled;
            SaveTxtButton.IsEnabled = isEnabled;
            SavePdfButton.IsEnabled = isEnabled;
            AutoFixCheckBox.IsEnabled = isEnabled;
            SearchablePdfCheckBox.IsEnabled = isEnabled;
            CombinePdfCheckBox.IsEnabled = isEnabled;
            AskOutputFolderCheckBox.IsEnabled = isEnabled;
            IncludeSubfoldersCheckBox.IsEnabled = isEnabled;
            OcrLanguageComboBox.IsEnabled = isEnabled;
        }


        private string RunOcr(string imagePath)
        {
            using TesseractEngine engine = CreateOcrEngine(GetSelectedOcrLanguage());

            return RunOcr(imagePath, engine, IsAutoFixEnabled());
        }

        private TesseractEngine CreateOcrEngine(string language)
        {
            string tessdataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");

            if (!Directory.Exists(tessdataPath))
                throw new DirectoryNotFoundException("Missing tessdata folder: " + tessdataPath);

            string trainedData = Path.Combine(tessdataPath, language + ".traineddata");

            if (!File.Exists(trainedData))
                throw new FileNotFoundException("Missing OCR language data file.", trainedData);

            return new TesseractEngine(tessdataPath, language, EngineMode.Default);
        }

        private string RunOcr(string imagePath, TesseractEngine engine, bool autoFixImage)
        {
            if (!autoFixImage)
                return RunOcrWithEngine(imagePath, engine);

            var preparedImages = imagePreprocessingService
                .PrepareMultipleForOcr(imagePath)
                .Where(File.Exists)
                .ToList();

            if (preparedImages.Count == 0)
                return RunOcrWithEngine(imagePath, engine);

            try
            {
                var textBuilder = new StringBuilder();

                foreach (string preparedImage in preparedImages)
                {
                    textBuilder.AppendLine(RunOcrWithEngine(preparedImage, engine));
                }

                return textBuilder.ToString();
            }
            finally
            {
                foreach (string preparedImage in preparedImages)
                {
                    TryDeletePreparedOcrImage(preparedImage, imagePath);
                }
            }
        }

        private static string RunOcrWithEngine(string imagePath, TesseractEngine engine)
        {
            using Pix img = Pix.LoadFromFile(imagePath);
            using Tesseract.Page page = engine.Process(img);

            return page.GetText();
        }

        private static void TryDeletePreparedOcrImage(string preparedImagePath, string originalImagePath)
        {
            string preparedFileName = Path.GetFileName(preparedImagePath);
            string expectedPrefix = Path.GetFileNameWithoutExtension(originalImagePath) + "_prepared_ocr";

            if (!preparedFileName.StartsWith(expectedPrefix, StringComparison.OrdinalIgnoreCase))
                return;

            try
            {
                File.Delete(preparedImagePath);
            }
            catch
            {
                // Leaving a temporary OCR image behind is less harmful than failing the conversion.
            }
        }

        private string CleanOcrText(string text)
        {
            text = RemoveDuplicateLines(text);
            text = NormalizeBullets(text);
            text = textCleanupService.Clean(text);
            text = FixBulletRecognition(text);
            text = RemoveDuplicateLines(text);

            return text;
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
