using Word = Microsoft.Office.Interop.Word;
using WordTool.Models;

namespace WordTool.Formatters
{
    public class MediaFormatter
    {
        public void FormatImages(Word.Document doc, FormattingTemplate template, System.Action<string> logger = null, System.Action checkStatus = null)
        {
            if (template == null) return;

            logger?.Invoke($"正在扫描并格式化图片对象，共发现 {doc.InlineShapes.Count} 个...");
            foreach (Word.InlineShape shape in doc.InlineShapes)
            {
                checkStatus?.Invoke();
                if (shape.Type == Word.WdInlineShapeType.wdInlineShapePicture)
                {
                    if (shape.Range.Paragraphs.Count > 0)
                    {
                        Word.Paragraph para = shape.Range.Paragraphs[1];
                        para.Format.Alignment = (Word.WdParagraphAlignment)template.ImageAlignment;
                        para.Format.FirstLineIndent = 0;
                        para.Format.SpaceBefore = template.ImageSpaceBefore;
                        para.Format.SpaceAfter = template.ImageSpaceAfter;
                    }
                }
            }
        }

        public void FormatTables(Word.Document doc, FormattingTemplate template, System.Action<string> logger = null, System.Action checkStatus = null)
        {
            if (template == null) return;

            logger?.Invoke($"正在扫描并调整表格，共发现 {doc.Tables.Count} 个...");
            foreach (Word.Table table in doc.Tables)
            {
                checkStatus?.Invoke();
                // 设置自动调整
                if (template.TableAutoFit)
                {
                    table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
                }
                table.Rows.Alignment = (Word.WdRowAlignment)template.TableAlignment;

                // 统计最大行索引以处理合并单元格情况
                int maxRowIndex = 0;
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    if (cell.RowIndex > maxRowIndex)
                    {
                        maxRowIndex = cell.RowIndex;
                    }
                }

                // 如果启用三线表，先清除表格级别边框
                if (template.TableThreeLine)
                {
                    table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                    table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                }

                // 遍历单元格，设置缩进、行距、字体、字号、加粗及三线表边框
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    checkStatus?.Invoke();
                    cell.Range.ParagraphFormat.FirstLineIndent = doc.Application.CentimetersToPoints(template.TableTextFirstLineIndentCm);
                    cell.Range.ParagraphFormat.LineSpacingRule = (Word.WdLineSpacing)template.TableTextLineSpacingRule;
                    
                    cell.Range.Font.NameFarEast = template.TableTextFontName;
                    cell.Range.Font.NameAscii = template.NormalFontNameAscii;
                    cell.Range.Font.Name = template.NormalFontNameAscii;
                    cell.Range.Font.Size = template.TableTextSize;

                    // 标题行设置加粗，其他行重置加粗（防止残留干扰）
                    if (cell.RowIndex == 1)
                    {
                        cell.Range.Font.Bold = template.TableHeaderBold ? 1 : 0;
                    }
                    else
                    {
                        cell.Range.Font.Bold = 0;
                    }

                    // 设置三线表样式
                    if (template.TableThreeLine)
                    {
                        // 清除当前单元格自身四周可能遗留的框线
                        cell.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                        cell.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                        cell.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                        cell.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;

                        if (cell.RowIndex == 1)
                        {
                            // 标题行上边框
                            cell.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                            cell.Borders[Word.WdBorderType.wdBorderTop].LineWidth = GetWdLineWidth(template.TableTopBottomBorderWidth);
                            
                            // 标题行下边框（中间线）
                            cell.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                            cell.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = GetWdLineWidth(template.TableHeaderBottomBorderWidth);
                        }
                        else if (cell.RowIndex == maxRowIndex)
                        {
                            // 最底行下边框
                            cell.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                            cell.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = GetWdLineWidth(template.TableTopBottomBorderWidth);
                        }
                    }
                }
            }
        }

        private Word.WdLineWidth GetWdLineWidth(float width)
        {
            if (width <= 0.25f) return Word.WdLineWidth.wdLineWidth025pt;
            if (width <= 0.50f) return Word.WdLineWidth.wdLineWidth050pt;
            if (width <= 0.75f) return Word.WdLineWidth.wdLineWidth075pt;
            if (width <= 1.00f) return Word.WdLineWidth.wdLineWidth100pt;
            if (width <= 1.50f) return Word.WdLineWidth.wdLineWidth150pt;
            if (width <= 2.25f) return Word.WdLineWidth.wdLineWidth225pt;
            if (width <= 3.00f) return Word.WdLineWidth.wdLineWidth300pt;
            if (width <= 4.50f) return Word.WdLineWidth.wdLineWidth450pt;
            return Word.WdLineWidth.wdLineWidth600pt;
        }
    }
}
