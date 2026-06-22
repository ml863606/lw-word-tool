using System;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordTool.Cleaners
{
    public class DocumentCleaner
    {
        public void Clean(Word.Document doc, Action<string> logger = null, int startSectionIndex = 2, Action checkStatus = null)
        {
            logger?.Invoke("正在扫描并清理文档中连续的空行...");
            RemoveMultipleEmptyLines(doc, checkStatus);
            
            logger?.Invoke($"正在以毫秒级通配符模式清理第 {startSectionIndex} 节及以后的段首冗余空格...");
            RemoveLeadingSpacesFast(doc, startSectionIndex, checkStatus);
        }

        private void RemoveMultipleEmptyLines(Word.Document doc, Action checkStatus)
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

            // 限制最大执行 8 次，防止因尾部段落标记或表格内空行无法被 Word 物理删除而陷入死循环挂起
            for (int k = 0; k < 8; k++)
            {
                checkStatus?.Invoke();
                bool found = find.Execute(Replace: Word.WdReplace.wdReplaceAll);
                if (!found)
                {
                    break;
                }
            }
        }

        private void RemoveLeadingSpacesFast(Word.Document doc, int startSectionIndex, Action checkStatus)
        {
            // 采用 Word 原生通配符批量替换，避免逐行遍历，提速 100 倍以上
            for (int i = startSectionIndex; i <= doc.Sections.Count; i++)
            {
                checkStatus?.Invoke();
                Word.Section section = doc.Sections[i];
                
                Word.Find find = section.Range.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                
                // 匹配段落标记后的任意空格、全角空格或制表符
                // Word 通配符中：^13 匹配段落标记，^t 匹配制表符，@ 匹配一个或多个
                find.Text = "^13[ 　^t]@";
                find.Replacement.Text = "^p";
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop; // 仅限于该节范围
                find.Format = false;
                find.MatchWildcards = true;

                find.Execute(Replace: Word.WdReplace.wdReplaceAll);

                // 特殊处理本节第一段，因为第一段前可能不带回车符 (^13)
                try
                {
                    if (section.Range.Paragraphs.Count > 0)
                    {
                        Word.Paragraph firstPara = section.Range.Paragraphs[1];
                        string text = firstPara.Range.Text;
                        if (!string.IsNullOrEmpty(text))
                        {
                            int spaceCount = 0;
                            while (spaceCount < text.Length && (text[spaceCount] == ' ' || text[spaceCount] == '\t' || text[spaceCount] == '　'))
                            {
                                spaceCount++;
                            }
                            
                            if (spaceCount > 0)
                            {
                                Word.Range startRange = firstPara.Range.Duplicate;
                                startRange.End = startRange.Start + spaceCount;
                                startRange.Text = "";
                            }
                        }
                    }
                }
                catch { }
            }
        }
    }
}
