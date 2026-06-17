using System;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordTool.Cleaners
{
    public class DocumentCleaner
    {
        public void Clean(Word.Document doc, Action<string> logger = null, int startSectionIndex = 2)
        {
            logger?.Invoke("正在扫描并清理文档中连续的空行...");
            RemoveMultipleEmptyLines(doc);
            
            logger?.Invoke($"正在扫描并清理第 {startSectionIndex} 节及以后的段首冗余空格...");
            RemoveLeadingSpaces(doc, startSectionIndex);
        }

        private void RemoveMultipleEmptyLines(Word.Document doc)
        {
            Word.Find find = doc.Content.Find;
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            find.Text = "^p^p";
            find.Replacement.Text = "^p";
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindContinue;
            find.Format = false;
            find.MatchCase = false;
            find.MatchWholeWord = false;
            find.MatchWildcards = false;
            find.MatchSoundsLike = false;
            find.MatchAllWordForms = false;

            while (find.Execute(Replace: Word.WdReplace.wdReplaceAll))
            {
                // 循环替换，直到没有连续的空行
            }
        }

        private void RemoveLeadingSpaces(Word.Document doc, int startSectionIndex)
        {
            // 对于非第一节（或者指定的起始节），清除段落首部的空格
            for (int i = startSectionIndex; i <= doc.Sections.Count; i++)
            {
                Word.Section section = doc.Sections[i];
                foreach (Word.Paragraph para in section.Range.Paragraphs)
                {
                    string text = para.Range.Text;
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    if (text.StartsWith(" ") || text.StartsWith("\t") || text.StartsWith("　"))
                    {
                        // 计算前导空白的数量并删除
                        int spaceCount = 0;
                        while (spaceCount < text.Length && (text[spaceCount] == ' ' || text[spaceCount] == '\t' || text[spaceCount] == '　'))
                        {
                            spaceCount++;
                        }
                        
                        if(spaceCount > 0)
                        {
                            Word.Range startRange = para.Range.Duplicate;
                            startRange.End = startRange.Start + spaceCount;
                            startRange.Text = "";
                        }
                    }
                }
            }
        }
    }
}
