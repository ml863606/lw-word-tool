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
            int imageIndex = 0;
            foreach (Word.InlineShape shape in doc.InlineShapes)
            {
                checkStatus?.Invoke();
                if (shape.Type == Word.WdInlineShapeType.wdInlineShapePicture)
                {
                    imageIndex++;
                    if (shape.Range.Paragraphs.Count > 0)
                    {
                        Word.Paragraph para = shape.Range.Paragraphs[1];
                        para.Format.Alignment = (Word.WdParagraphAlignment)template.ImageAlignment;
                        para.Format.FirstLineIndent = 0;
                        para.Format.SpaceBefore = template.ImageSpaceBefore;
                        para.Format.SpaceAfter = template.ImageSpaceAfter;
                        logger?.Invoke($"【图片排版】第 {imageIndex} 张图片 -> {template.ImageAlignment}，首行缩进 0，段前 {template.ImageSpaceBefore:0.#}pt，段后 {template.ImageSpaceAfter:0.#}pt");
                    }
                    else
                    {
                        logger?.Invoke($"【图片排版】第 {imageIndex} 张图片未找到所在段落，已跳过段落对齐设置");
                    }
                }
            }
        }

        public void FormatTables(Word.Document doc, FormattingTemplate template, System.Action<string> logger = null, System.Action checkStatus = null)
        {
            if (template == null) return;

            logger?.Invoke($"正在扫描并调整表格，共发现 {doc.Tables.Count} 个...");
            int tableIndex = 0;
            foreach (Word.Table table in doc.Tables)
            {
                checkStatus?.Invoke();
                tableIndex++;
                int cellCount = table.Range.Cells.Count;

                // 设置自动调整
                if (template.TableAutoFit)
                {
                    table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
                }
                table.Rows.Alignment = (Word.WdRowAlignment)template.TableAlignment;
                logger?.Invoke($"【表格排版】第 {tableIndex} 个表格 -> {template.TableAlignment}，自动适应窗口: {YesNo(template.TableAutoFit)}，单元格数: {cellCount}");

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
                    logger?.Invoke($"【表格排版】第 {tableIndex} 个表格 -> 已清除原边框，准备套用三线表");
                }

                // 遍历单元格，设置缩进、行距、字体、字号、加粗及三线表边框
                int formattedCells = 0;
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    checkStatus?.Invoke();
                    formattedCells++;
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

                logger?.Invoke($"【表格排版】第 {tableIndex} 个表格 -> 已格式化 {formattedCells} 个单元格；正文 {template.TableTextFontName}/{template.NormalFontNameAscii} {template.TableTextSize:0.#}pt，{template.TableTextLineSpacingRule}，首行缩进 {template.TableTextFirstLineIndentCm:0.##}cm，标题行{BoldText(template.TableHeaderBold)}");

                if (template.TableThreeLine)
                {
                    logger?.Invoke($"【表格排版】第 {tableIndex} 个表格 -> 三线表完成：上下边框 {template.TableTopBottomBorderWidth:0.##}pt，标题行下边框 {template.TableHeaderBottomBorderWidth:0.##}pt");
                }
            }
        }

        private string YesNo(bool value)
        {
            return value ? "是" : "否";
        }

        private string BoldText(bool isBold)
        {
            return isBold ? "加粗" : "不加粗";
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
