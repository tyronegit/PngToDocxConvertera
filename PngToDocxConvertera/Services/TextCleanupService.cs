using System.Linq;
using System.Text.RegularExpressions;

namespace PngToDocxConvertera.Services
{
    public class TextCleanupService
    {
        public string Clean(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            // Normalize line endings
            text = text.Replace("\r\n", "\n");

            // Remove OCR garbage symbols
            text = Regex.Replace(
                text,
                @"[|~<>`]",
                ""
            );

            // Remove random leading e before capitalized words
            // Example: eAddictions -> Addictions
            text = Regex.Replace(
                text,
                @"\be([A-Z])",
                "$1"
            );

            // Fix multiple spaces
            text = Regex.Replace(
                text,
                @"[ ]{2,}",
                " "
            );

            // Fix repeated blank lines
            text = Regex.Replace(
                text,
                @"\n{3,}",
                "\n\n"
            );

            // Remove isolated garbage lines
            text = Regex.Replace(
                text,
                @"^\W{1,5}$",
                "",
                RegexOptions.Multiline
            );

            // Trim document
            text = text.Trim();

            return text;
        }
    

    private static int ScoreOcrResult(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return 0;

            int score = 0;

            score += text.Where(char.IsLetter).Count();

            score -= text.Count(c =>
                !char.IsLetterOrDigit(c) &&
                !char.IsWhiteSpace(c) &&
                c != '.' &&
                c != ',' &&
                c != ':' &&
                c != ';' &&
                c != '-' &&
                c != '•');

            return score;
        }

    }
}