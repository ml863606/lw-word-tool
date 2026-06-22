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
                    logger?.Invoke($"发现已有目录 {doc.TablesOfContents.Count} 个，正在更新页码...");
                    int index = 0;
                    foreach (Word.TableOfContents toc in doc.TablesOfContents)
                    {
                        index++;
                        toc.Update();
                        logger?.Invoke($"【目录】第 {index} 个目录已更新页码与标题引用");
                    }
                }
                else
                {
                    logger?.Invoke($"【目录】发现已有目录 {doc.TablesOfContents.Count} 个；模板设置为不自动更新，已跳过");
                }
            }
            else
            {
                logger?.Invoke($"文档中未发现目录，正在文档开头自动插入标准目录...");
                // 如果没有目录，在文档最前面（或第一节之后）插入目录
                Word.Range range = doc.Range(0, 0);
                string insertPosition = "文档开头";
                if (doc.Sections.Count > 1)
                {
                    range = doc.Sections[2].Range;
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    insertPosition = "第 2 节开头";
                }
                
                doc.TablesOfContents.Add(
                    range,
                    UseHeadingStyles: true,
                    UpperHeadingLevel: 1,
                    LowerHeadingLevel: template.TocLevels,
                    UseHyperlinks: true,
                    HidePageNumbersInWeb: true,
                    UseOutlineLevels: true);
                logger?.Invoke($"【目录】已在{insertPosition}插入目录；包含 1-{template.TocLevels} 级标题，启用超链接与大纲级别");
            }
        }
    }
}
