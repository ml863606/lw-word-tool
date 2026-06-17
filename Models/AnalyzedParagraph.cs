using Word = Microsoft.Office.Interop.Word;

namespace WordTool.Models
{
    public class AnalyzedParagraph
    {
        public Word.Paragraph Paragraph { get; set; }
        public ParagraphRole Role { get; set; }
        public string TextContent { get; set; }
        public float Confidence { get; set; } // 置信度 0-1

        public AnalyzedParagraph(Word.Paragraph paragraph, ParagraphRole role, string text, float confidence = 1.0f)
        {
            Paragraph = paragraph;
            Role = role;
            TextContent = text;
            Confidence = confidence;
        }
    }
}
