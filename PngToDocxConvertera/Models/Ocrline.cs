using System;
using System.Collections.Generic;
using System.Text;

// ================================================================
// File: Models/OcrLine.cs
// ================================================================

namespace PngToDocxConvertera.Models
{
    public sealed class OcrLine
    {
        public string Text { get; set; } = string.Empty;
        public float Confidence { get; set; }
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public bool IsHeading { get; set; }
        public bool IsBullet { get; set; }
        public bool IsLikelyTableRow { get; set; }
    }
}
