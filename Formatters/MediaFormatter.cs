using Word = Microsoft.Office.Interop.Word;

namespace WordTool.Formatters
{
    public class MediaFormatter
    {
        public void FormatImages(Word.Document doc, System.Action<string> logger = null)
        {
            logger?.Invoke($"正在扫描并格式化图片对象，共发现 {doc.InlineShapes.Count} 个...");
            foreach (Word.InlineShape shape in doc.InlineShapes)
            {
                if (shape.Type == Word.WdInlineShapeType.wdInlineShapePicture)
                {
                    if (shape.Range.Paragraphs.Count > 0)
                    {
                        Word.Paragraph para = shape.Range.Paragraphs[1];
                        para.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        para.Format.FirstLineIndent = 0;
                        para.Format.SpaceBefore = 6f;
                        para.Format.SpaceAfter = 6f;
                    }
                }
            }
        }

        public void FormatTables(Word.Document doc, System.Action<string> logger = null)
        {
            logger?.Invoke($"正在扫描并调整表格宽度，共发现 {doc.Tables.Count} 个...");
            foreach (Word.Table table in doc.Tables)
            {
                // 设置自动调整
                table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
                table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;

                // 遍历表格内容，设置为单倍行距，取消缩进
                foreach (Word.Row row in table.Rows)
                {
                    foreach (Word.Cell cell in row.Cells)
                    {
                        cell.Range.ParagraphFormat.FirstLineIndent = 0;
                        cell.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                        cell.Range.Font.Size = 10.5f; // 五号字
                    }
                }
            }
        }
    }
}
