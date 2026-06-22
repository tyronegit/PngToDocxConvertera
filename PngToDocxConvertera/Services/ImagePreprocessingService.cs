using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
using System.Collections.Generic;
using System.IO;

namespace PngToDocxConvertera.Services
{
    public class ImagePreprocessingService
    {
        public IEnumerable<string> PrepareMultipleForOcr(string imagePath)
        {
            List<string> outputImages = new();

            if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
                return outputImages;

            string folder = Path.GetDirectoryName(imagePath)!;
            string fileName = Path.GetFileNameWithoutExtension(imagePath);
            string preparedPath = Path.Combine(folder, fileName + "_prepared_ocr.png");

            using SixLabors.ImageSharp.Image<Rgba32> image =
                SixLabors.ImageSharp.Image.Load<Rgba32>(imagePath);

            bool shouldInvert = IsDarkImage(image);

            image.Mutate(ctx =>
            {
                ctx.Grayscale();

                ctx.Contrast(1.15f);

                if (shouldInvert)
                {
                    ctx.BinaryThreshold(0.62f);

                    ctx.Invert();
                }
                else
                {
                    // Better for presentation slides
                    ctx.Resize(image.Width * 2, image.Height * 2);

                    ctx.GaussianSharpen(1.2f);
                }
            });

            image.Save(preparedPath);

            outputImages.Add(preparedPath);

            return outputImages;
        }

        private bool IsDarkImage(SixLabors.ImageSharp.Image<Rgba32> image)
        {
            double totalBrightness = 0;
            int sampleCount = 0;

            for (int y = 0; y < image.Height; y += 10)
            {
                for (int x = 0; x < image.Width; x += 10)
                {
                    Rgba32 pixel = image[x, y];

                    double brightness =
                        (pixel.R + pixel.G + pixel.B) / 3.0;

                    totalBrightness += brightness;
                    sampleCount++;
                }
            }

            if (sampleCount == 0)
                return false;

            double averageBrightness = totalBrightness / sampleCount;

            return averageBrightness < 110;
        }
    }
}