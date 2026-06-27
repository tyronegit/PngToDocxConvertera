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
        private string pendingBatchFolderPath = "";
        private bool isBatchRunning;
        private CancellationTokenSource? batchCancellationTokenSource;

        private static readonly string[] SupportedImageExtensions =
        {
            ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"
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

        private bool IsAutoFixEnabled()
        {
            return AutoFixCheckBox.IsChecked == true;
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
            MessageBox.Show("PNG to DOCX Converter\nVersion 2.1\n\nA lightweight utility utilizing Tesseract OCR and OpenXML.",
                            "About Application",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
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
            pendingBatchFolderPath = "";
            DropHintText.Visibility = Visibility.Visible;
            BatchProgressBar.Value = 0;
            BatchCounterTextBlock.Text = "Batch: idle";
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
                text = CleanOcrText(text);

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
                AutoFixImages = IsAutoFixEnabled()
            };

            var progress = new Progress<BatchProgressInfo>(UpdateBatchProgress);
            batchCancellationTokenSource = new CancellationTokenSource();

            isBatchRunning = true;
            SetBatchControlsEnabled(false);
            BatchProgressBar.Maximum = 100;
            BatchProgressBar.Value = 0;
            BatchCounterTextBlock.Text = $"Batch: 0 of {imageFiles.Count} (0%)";
            SetStatus($"Starting {imageFiles.Count} file(s) to {extension}...");

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
                        CreateSimplePdf(outputPath, text);

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
            SelectImageButton.IsEnabled = isEnabled;
            PreviewOcrButton.IsEnabled = isEnabled;
            ConvertSingleButton.IsEnabled = isEnabled;
            SaveTxtButton.IsEnabled = isEnabled;
            SavePdfButton.IsEnabled = isEnabled;
            AutoFixCheckBox.IsEnabled = isEnabled;
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
