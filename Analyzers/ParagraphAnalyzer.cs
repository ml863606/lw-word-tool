using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using WordTool.Models;
using Word = Microsoft.Office.Interop.Word;

namespace WordTool.Analyzers
{
    public class ParagraphAnalyzer
    {
        public List<AnalyzedParagraph> Analyze(Word.Document doc, int startSectionIndex = 2, Action checkStatus = null)
        {
            List<AnalyzedParagraph> results = new List<AnalyzedParagraph>();

            for (int i = startSectionIndex; i <= doc.Sections.Count; i++)
            {
                checkStatus?.Invoke();
                Word.Section section = doc.Sections[i];
                foreach (Word.Paragraph para in section.Range.Paragraphs)
                {
                    checkStatus?.Invoke();
                    AnalyzedParagraph analyzed = AnalyzeSingleParagraph(para);
                    if (analyzed != null)
                    {
                        results.Add(analyzed);
                    }
                }
            }
            return results;
        }

        private AnalyzedParagraph AnalyzeSingleParagraph(Word.Paragraph para)
        {
            string text = para.Range.Text;
            if (string.IsNullOrWhiteSpace(text) || text == "\r" || text == "\a") return null;
            text = text.Trim();

            // 1. 判断是否为空行或特殊元素
            if (text.Length == 0) return null;

            // 2. 特征提取
            bool isBold = para.Range.Font.Bold != 0;
            float fontSize = para.Range.Font.Size;
            int length = text.Length;
            Word.WdParagraphAlignment alignment = para.Format.Alignment;

            // 3. 规则判断
            ParagraphRole role = ParagraphRole.Normal;
            float confidence = 0.5f;

            // 题注判断: 图1-1，表1-1
            if ((text.StartsWith("图") || text.StartsWith("表")) && text.Length < 30 && Regex.IsMatch(text, @"^(图|表)\s*\d+[-_.]\d+"))
            {
                role = ParagraphRole.Caption;
                confidence = 0.9f;
            }
            // 表注判断: 注：，说明：
            else if ((text.StartsWith("注") || text.StartsWith("说明") || text.StartsWith("Note") || text.StartsWith("note")) && text.Length < 100 && Regex.IsMatch(text, @"^(注(意|释)?|说明|Note|note)\s*[:：]"))
            {
                role = ParagraphRole.TableNote;
                confidence = 0.9f;
            }
            // 参考文献判断
            else if (Regex.IsMatch(text, @"^\[\d+\]"))
            {
                role = ParagraphRole.Reference;
                confidence = 0.95f;
            }
            // 标题判断 (字数不能太多)
            else if (length < 50)
            {
                if (Regex.IsMatch(text, @"^第[一二三四五六七八九十百]+章") || Regex.IsMatch(text, @"^[一二三四五六七八九十]+、"))
                {
                    role = ParagraphRole.Heading1;
                    confidence = 0.9f;
                }
                else if (Regex.IsMatch(text, @"^\d+\.\s") || Regex.IsMatch(text, @"^（[一二三四五六七八九十]+）"))
                {
                    // 1. 简介 或 (一) 简介
                    role = ParagraphRole.Heading2;
                    confidence = 0.8f;
                }
                else if (Regex.IsMatch(text, @"^\d+\.\d+"))
                {
                    // 1.1 背景
                    // 判断是否有三级: 1.1.1
                    if (Regex.IsMatch(text, @"^\d+\.\d+\.\d+"))
                    {
                        role = ParagraphRole.Heading3;
                    }
                    else
                    {
                        role = ParagraphRole.Heading2;
                    }
                    confidence = 0.8f;
                }
                // 辅助判断：如果字号大于12或者加粗，且较短，很可能是标题
                else if ((fontSize > 12 || isBold) && length < 30)
                {
                    role = ParagraphRole.Heading2; // 默认给个二级，交给用户在界面确认
                    confidence = 0.6f;
                }
            }

            return new AnalyzedParagraph(para, role, text, confidence);
        }
    }
}
