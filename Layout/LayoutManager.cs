using Word = Microsoft.Office.Interop.Word;
using WordTool.Models;

namespace WordTool.Layout
{
    public class LayoutManager
    {
        public void UpdateOrInsertTOC(Word.Document doc, FormattingTemplate template, System.Action<string> logger = null, System.Action checkStatus = null)
        {
            if (template == null) return;
            checkStatus?.Invoke();

            if (doc.TablesOfContents.Count > 0)
            {
                if (template.TocAutoUpdate)
                {
                    logger?.Invoke($"发现已有目录，正在更新页码...");
                    foreach (Word.TableOfContents toc in doc.TablesOfContents)
                    {
                        toc.Update();
                    }
                }
            }
            else
            {
                logger?.Invoke($"文档中未发现目录，正在文档开头自动插入标准目录...");
                // 如果没有目录，在文档最前面（或第一节之后）插入目录
                Word.Range range = doc.Range(0, 0);
                if (doc.Sections.Count > 1)
                {
                    range = doc.Sections[2].Range;
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                }
                
                doc.TablesOfContents.Add(
                    range,
                    UseHeadingStyles: true,
                    UpperHeadingLevel: 1,
                    LowerHeadingLevel: template.TocLevels,
                    UseHyperlinks: true,
                    HidePageNumbersInWeb: true,
                    UseOutlineLevels: true);
            }
        }
    }
}
