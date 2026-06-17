using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordTool.Styles
{
    public class StyleManager
    {
        public void RebuildStyles(Word.Document doc, Action<string> logger = null)
        {
            logger?.Invoke("正在设置纸张大小和全局页边距...");
            SetupPage(doc);
            
            logger?.Invoke("正在重塑全局【正文】基础样式...");
            RebuildNormalStyle(doc);
            
            logger?.Invoke("正在重新定义各级标题的大纲格式...");
            RebuildHeadingStyles(doc);
            
            logger?.Invoke("正在配置【题注】和【参考文献】特殊样式...");
            RebuildSpecialStyles(doc);
        }

        private void SetupPage(Word.Document doc)
        {
            Word.PageSetup ps = doc.PageSetup;
            ps.PaperSize = Word.WdPaperSize.wdPaperA4;
            ps.TopMargin = doc.Application.CentimetersToPoints(3.0f);
            ps.BottomMargin = doc.Application.CentimetersToPoints(2.5f);
            ps.LeftMargin = doc.Application.CentimetersToPoints(2.5f);
            ps.RightMargin = doc.Application.CentimetersToPoints(2.5f);
        }

        private void RebuildNormalStyle(Word.Document doc)
        {
            Word.Style normalStyle = doc.Styles[Word.WdBuiltinStyle.wdStyleNormal];
            normalStyle.Font.NameFarEast = "宋体";
            normalStyle.Font.NameAscii = "Times New Roman";
            normalStyle.Font.Name = "Times New Roman";
            normalStyle.Font.Size = 12f; // 小四
            
            normalStyle.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
            normalStyle.ParagraphFormat.FirstLineIndent = doc.Application.CentimetersToPoints(0.74f); // 大约2个中文字符宽度
            normalStyle.ParagraphFormat.SpaceBefore = 0;
            normalStyle.ParagraphFormat.SpaceAfter = 0;
            normalStyle.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
        }

        private void RebuildHeadingStyles(Word.Document doc)
        {
            // Heading 1
            Word.Style h1 = doc.Styles[Word.WdBuiltinStyle.wdStyleHeading1];
            h1.Font.NameFarEast = "黑体";
            h1.Font.Size = 16f; // 三号
            h1.Font.Bold = 1;
            h1.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            h1.ParagraphFormat.FirstLineIndent = 0;
            h1.ParagraphFormat.KeepWithNext = -1; // -1 means True in Word VBA/Interop
            h1.ParagraphFormat.SpaceBefore = 12f;
            h1.ParagraphFormat.SpaceAfter = 12f;

            // Heading 2
            Word.Style h2 = doc.Styles[Word.WdBuiltinStyle.wdStyleHeading2];
            h2.Font.NameFarEast = "黑体";
            h2.Font.Size = 14f; // 四号
            h2.Font.Bold = 1;
            h2.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            h2.ParagraphFormat.FirstLineIndent = 0;
            h2.ParagraphFormat.KeepWithNext = -1;
            h2.ParagraphFormat.SpaceBefore = 6f;
            h2.ParagraphFormat.SpaceAfter = 6f;

            // Heading 3
            Word.Style h3 = doc.Styles[Word.WdBuiltinStyle.wdStyleHeading3];
            h3.Font.NameFarEast = "黑体";
            h3.Font.Size = 12f; // 小四
            h3.Font.Bold = 1;
            h3.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            h3.ParagraphFormat.FirstLineIndent = 0;
            h3.ParagraphFormat.KeepWithNext = -1;
        }

        private void RebuildSpecialStyles(Word.Document doc)
        {
            // Reference style (can be created if not exists, or just use normal but customize it in Formatter)
            // It's safer to create a custom style.
            Word.Style refStyle;
            try
            {
                refStyle = doc.Styles["参考文献_Auto"];
            }
            catch
            {
                refStyle = doc.Styles.Add("参考文献_Auto", Word.WdStyleType.wdStyleTypeParagraph);
            }
            
            object baseStyleName = Word.WdBuiltinStyle.wdStyleNormal;
            refStyle.set_BaseStyle(ref baseStyleName);
            refStyle.Font.Size = 10.5f; // 五号
            refStyle.ParagraphFormat.FirstLineIndent = doc.Application.CentimetersToPoints(-0.74f); // 悬挂缩进
            refStyle.ParagraphFormat.LeftIndent = doc.Application.CentimetersToPoints(0.74f);
            
            // Caption (题注)
            Word.Style captionStyle = doc.Styles[Word.WdBuiltinStyle.wdStyleCaption];
            captionStyle.Font.Size = 10.5f; // 五号
            captionStyle.Font.Bold = 0;
            captionStyle.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            captionStyle.ParagraphFormat.FirstLineIndent = 0;
        }
    }
}
