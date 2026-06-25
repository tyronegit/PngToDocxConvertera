using PdfSharp.Fonts;
using System.IO;

namespace PngToDocConvertera
{
    public class WindowsFontResolver : IFontResolver
    {
        public byte[] GetFont(string faceName)
        {
            string fontPath = faceName switch
            {
                "Arial#Bold" => @"C:\Windows\Fonts\arialbd.ttf",
                _ => @"C:\Windows\Fonts\arial.ttf"
            };

            return File.ReadAllBytes(fontPath);
        }

        public FontResolverInfo ResolveTypeface(
            string familyName,
            bool isBold,
            bool isItalic)
        {
            return new FontResolverInfo(
                isBold ? "Arial#Bold" : "Arial#Regular");
        }
    }
}