namespace PngToDocxConvertera.Models
{
    public class OcrPipelineResult
    {
        public string Text { get; set; } = string.Empty;

        public int Score { get; set; }

        public OcrProfile Profile { get; set; }

        public string ImageVersion { get; set; } = string.Empty;
    }
}