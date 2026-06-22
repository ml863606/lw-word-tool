using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordTool.Cleaners
{
    public class DocumentCleaner
    {
        public void Clean(Word.Document doc, Action<string> logger = null, int startSectionIndex = 2, Action checkStatus = null)
        {
            logger?.Invoke($"正在以毫秒级通配符模式清理第 {startSectionIndex} 节及以后的段首冗余空格...");
            RemoveLeadingSpacesFast(doc, startSectionIndex, logger, checkStatus);

            logger?.Invoke("正在以单次通配符模式压缩文档中连续的空行...");
            RemoveMultipleEmptyLines(doc, logger, checkStatus);
        }

        private void RemoveMultipleEmptyLines(Word.Document doc, Action<string> logger, Action checkStatus)
        {
            checkStatus?.Invoke();
            int beforeParagraphCount = 0;
            try
            {
                beforeParagraphCount = doc.Paragraphs.Count;
            }
            catch { }

            try
            {
                // Word 通配符一次性压缩连续段落标记，避免 ^p^p -> ^p 反复全篇扫描。
                ExecuteReplaceAll(
                    doc.Content.Duplicate,
                    "^13{2,}",
                    "^p",
                    matchWildcards: true);
                LogEmptyLineResult(doc, beforeParagraphCount, logger);
            }
            catch
            {
                // 极少数 Word 版本/文档结构下通配符替换可能失败，回退到旧策略但减少扫描次数。
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

                for (int k = 0; k < 4; k++)
                {
                    checkStatus?.Invoke();
                    bool found = find.Execute(Replace: Word.WdReplace.wdReplaceAll);
                    if (!found)
                    {
                        break;
                    }
                }

                LogEmptyLineResult(doc, beforeParagraphCount, logger);
            }
        }

        private void RemoveLeadingSpacesFast(Word.Document doc, int startSectionIndex, Action<string> logger, Action checkStatus)
        {
            // 采用 Word 原生通配符批量替换，避免逐行遍历，提速 100 倍以上
            for (int i = startSectionIndex; i <= doc.Sections.Count; i++)
            {
                checkStatus?.Invoke();
                Word.Section section = doc.Sections[i];
                int paragraphCount = 0;
                try
                {
                    paragraphCount = section.Range.Paragraphs.Count;
                }
                catch { }
                
                // 匹配段落标记后的任意空格、全角空格或制表符
                // Word 通配符中：^13 匹配段落标记，^t 匹配制表符，@ 匹配一个或多个
                ExecuteReplaceAll(
                    section.Range.Duplicate,
                    "^13[ 　^t]@",
                    "^p",
                    matchWildcards: true);
                logger?.Invoke($"【清理脏数据】第 {i} 节 -> 已批量清理段首半角空格、全角空格和制表符；扫描段落数: {paragraphCount}");

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
                                logger?.Invoke($"【清理脏数据】第 {i} 节第 1 段 -> 删除开头空白字符 {spaceCount} 个");
                            }
                        }
                    }
                }
                catch { }
            }
        }

        private void ExecuteReplaceAll(Word.Range range, string findText, string replacementText, bool matchWildcards)
        {
            Word.Find find = range.Find;
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            find.Text = findText;
            find.Replacement.Text = replacementText;
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            find.Format = false;
            find.MatchCase = false;
            find.MatchWholeWord = false;
            find.MatchWildcards = matchWildcards;
            find.MatchSoundsLike = false;
            find.MatchAllWordForms = false;

            find.Execute(Replace: Word.WdReplace.wdReplaceAll);
        }

        private void LogEmptyLineResult(Word.Document doc, int beforeParagraphCount, Action<string> logger)
        {
            if (logger == null) return;

            try
            {
                int afterParagraphCount = doc.Paragraphs.Count;
                if (beforeParagraphCount > 0)
                {
                    int removed = Math.Max(0, beforeParagraphCount - afterParagraphCount);
                    logger.Invoke($"【清理脏数据】连续空行压缩完成 -> 段落数 {beforeParagraphCount} -> {afterParagraphCount}，减少 {removed} 个空段落");
                }
                else
                {
                    logger.Invoke("【清理脏数据】连续空行压缩完成");
                }
            }
            catch
            {
                logger.Invoke("【清理脏数据】连续空行压缩完成");
            }
        }
    }
}
