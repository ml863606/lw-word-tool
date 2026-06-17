using System.Collections.Generic;
using WordTool.Models;
using Word = Microsoft.Office.Interop.Word;

namespace WordTool.Formatters
{
    public class DocumentFormatter
    {
        public void ApplyFormatting(List<AnalyzedParagraph> paragraphs, System.Action<string> logger = null)
        {
            logger?.Invoke($"正在清除 {paragraphs.Count} 个段落的残留直接格式，并重新赋予规范样式...");
            foreach (var item in paragraphs)
            {
                Word.Paragraph para = item.Paragraph;
                
                // 1. 清除直接格式
                para.Range.Font.Reset();
                para.Range.ParagraphFormat.Reset();

                // 2. 根据角色应用样式
                switch (item.Role)
                {
                    case ParagraphRole.Heading1:
                        para.set_Style(Word.WdBuiltinStyle.wdStyleHeading1);
                        break;
                    case ParagraphRole.Heading2:
                        para.set_Style(Word.WdBuiltinStyle.wdStyleHeading2);
                        break;
                    case ParagraphRole.Heading3:
                        para.set_Style(Word.WdBuiltinStyle.wdStyleHeading3);
                        break;
                    case ParagraphRole.Reference:
                        try { para.set_Style("参考文献_Auto"); } catch { para.set_Style(Word.WdBuiltinStyle.wdStyleNormal); }
                        break;
                    case ParagraphRole.Caption:
                        para.set_Style(Word.WdBuiltinStyle.wdStyleCaption);
                        break;
                    case ParagraphRole.Normal:
                    default:
                        para.set_Style(Word.WdBuiltinStyle.wdStyleNormal);
                        break;
                }
            }
        }
    }
}
