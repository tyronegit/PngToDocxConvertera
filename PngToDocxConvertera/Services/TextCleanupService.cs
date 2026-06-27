using System;
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
            text = text.Replace("\r\n", "\n").Replace("\r", "\n");

            // Common OCR bullet/symbol mistakes
            text = text.Replace("¢", "•");
            text = text.Replace("·", "•");
            text = text.Replace("●", "•");
            text = text.Replace("○", "•");
            text = text.Replace("◦", "•");

            // Remove degree symbol when OCR places it at line start
            text = Regex.Replace(text, @"^°\s*", "", RegexOptions.Multiline);

            // Normalize quotes and apostrophes
            text = text.Replace("“", "\"");
            text = text.Replace("”", "\"");
            text = text.Replace("‘", "'");
            text = text.Replace("’", "'");

            // Normalize dashes
            text = text.Replace("—", "-");
            text = text.Replace("–", "-");

            // Remove random OCR artifacts
            text = text.Replace("™", "");
            text = text.Replace("©", "");

            // Fix known OCR/header mistakes
            text = Regex.Replace(text, @"&\s*ECUREMEDY", "SECUREMEDY", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"\bECUREMEDY\b", "SECUREMEDY", RegexOptions.IgnoreCase);

            // Fix known document title mistakes
            text = Regex.Replace(
                text,
                @"\bfir\s+watch\s+and\s+responsibilities\b",
                "Fire Watch and Fire Response Responsibilities",
                RegexOptions.IgnoreCase);

            // Fix common missing foot mark after 35
            text = Regex.Replace(
                text,
                @"\bMore\s+than\s+35\s+away\b",
                "More than 35' away",
                RegexOptions.IgnoreCase);

            // Remove duplicate punctuation
            text = Regex.Replace(text, @"([,.;:]){2,}", "$1");

            // Remove duplicate punctuation
            text = Regex.Replace(text, @"([.,;:]){2,}", "$1");

            // Fix extra spaces
            text = Regex.Replace(text, @"[ \t]{2,}", " ");

            // Remove spaces before punctuation
            text = Regex.Replace(text, @"\s+([.,;:!?])", "$1");

            // Limit too many blank lines
            text = Regex.Replace(text, @"\n{3,}", "\n\n");

            // Trim right side of each line
            string[] lines = text.Split('\n');

            for (int i = 0; i < lines.Length; i++)
            {
                lines[i] = lines[i].TrimEnd();
            }

            return string.Join(Environment.NewLine, lines).Trim();
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