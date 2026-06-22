using System;
using Word = Microsoft.Office.Interop.Word;
using WordTool.Models;

namespace WordTool.Styles
{
    public class StyleManager
    {
        public void RebuildStyles(Word.Document doc, FormattingTemplate template, Action<string> logger = null, Action checkStatus = null)
        {
            if (template == null) return;

            checkStatus?.Invoke();
            logger?.Invoke("正在设置纸张大小和全局页边距...");
            SetupPage(doc, template);
            logger?.Invoke($"【页面设置】纸张: {template.PaperSize}；页边距: 上 {template.TopMarginCm:0.##}cm / 下 {template.BottomMarginCm:0.##}cm / 左 {template.LeftMarginCm:0.##}cm / 右 {template.RightMarginCm:0.##}cm");
            
            checkStatus?.Invoke();
            logger?.Invoke("正在重塑全局【正文】基础样式...");
            RebuildNormalStyle(doc, template);
            logger?.Invoke($"【样式重建】正文 -> {template.NormalFontNameFarEast}/{template.NormalFontNameAscii} {template.NormalFontSize:0.#}pt，{template.NormalLineSpacingRule}，首行缩进 {template.NormalFirstLineIndentChars:0.#} 字符，两端对齐，段前/段后 0pt");
            
            checkStatus?.Invoke();
            logger?.Invoke("正在重新定义各级标题的大纲格式...");
            RebuildHeadingStyles(doc, template);
            logger?.Invoke($"【样式重建】标题 1 -> {template.H1FontNameFarEast}/{template.NormalFontNameAscii} {template.H1FontSize:0.#}pt，{BoldText(template.H1Bold)}，{template.H1Alignment}，段前 {template.H1SpaceBefore:0.#}pt，段后 {template.H1SpaceAfter:0.#}pt，保持与下段同页");
            logger?.Invoke($"【样式重建】标题 2 -> {template.H2FontNameFarEast}/{template.NormalFontNameAscii} {template.H2FontSize:0.#}pt，{BoldText(template.H2Bold)}，{template.H2Alignment}，段前 {template.H2SpaceBefore:0.#}pt，段后 {template.H2SpaceAfter:0.#}pt，保持与下段同页");
            logger?.Invoke($"【样式重建】标题 3 -> {template.H3FontNameFarEast}/{template.NormalFontNameAscii} {template.H3FontSize:0.#}pt，{BoldText(template.H3Bold)}，{template.H3Alignment}，保持与下段同页");
            
            checkStatus?.Invoke();
            logger?.Invoke("正在配置【题注】和【参考文献】特殊样式...");
            RebuildSpecialStyles(doc, template);
            logger?.Invoke($"【样式重建】参考文献_Auto -> {template.RefFontSize:0.#}pt，首行缩进 {template.RefFirstLineIndentCm:0.##}cm，左缩进 {template.RefLeftIndentCm:0.##}cm");
            logger?.Invoke($"【样式重建】题注 -> 宋体/Times New Roman {template.CaptionFontSize:0.#}pt，{BoldText(template.CaptionBold)}，{template.CaptionAlignment}");
            logger?.Invoke($"【样式重建】表注_Auto -> {template.NoteFontNameFarEast}/{template.NormalFontNameAscii} {template.NoteFontSize:0.#}pt，{template.NoteAlignment}");
        }

        private void SetupPage(Word.Document doc, FormattingTemplate template)
        {
            Word.PageSetup ps = doc.PageSetup;
            if (template.PaperSize == PaperSizeOption.A4)
            {
                ps.PaperSize = Word.WdPaperSize.wdPaperA4;
            }
            else if (template.PaperSize == PaperSizeOption.Letter)
            {
                ps.PaperSize = Word.WdPaperSize.wdPaperLetter;
            }
            else
            {
                ps.PaperSize = Word.WdPaperSize.wdPaperA4;
            }

            ps.TopMargin = doc.Application.CentimetersToPoints(template.TopMarginCm);
            ps.BottomMargin = doc.Application.CentimetersToPoints(template.BottomMarginCm);
            ps.LeftMargin = doc.Application.CentimetersToPoints(template.LeftMarginCm);
            ps.RightMargin = doc.Application.CentimetersToPoints(template.RightMarginCm);
        }

        private void RebuildNormalStyle(Word.Document doc, FormattingTemplate template)
        {
            Word.Style normalStyle = doc.Styles[Word.WdBuiltinStyle.wdStyleNormal];
            normalStyle.Font.NameFarEast = template.NormalFontNameFarEast;
            normalStyle.Font.NameAscii = template.NormalFontNameAscii;
            normalStyle.Font.Name = template.NormalFontNameAscii;
            normalStyle.Font.Size = template.NormalFontSize;
            
            normalStyle.ParagraphFormat.LineSpacingRule = (Word.WdLineSpacing)template.NormalLineSpacingRule;
            normalStyle.ParagraphFormat.CharacterUnitFirstLineIndent = template.NormalFirstLineIndentChars;
            normalStyle.ParagraphFormat.SpaceBefore = 0;
            normalStyle.ParagraphFormat.SpaceAfter = 0;
            normalStyle.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
        }

        private void RebuildHeadingStyles(Word.Document doc, FormattingTemplate template)
        {
            // Heading 1
            Word.Style h1 = doc.Styles[Word.WdBuiltinStyle.wdStyleHeading1];
            h1.Font.NameFarEast = template.H1FontNameFarEast;
            h1.Font.NameAscii = template.NormalFontNameAscii;
            h1.Font.Name = template.NormalFontNameAscii;
            h1.Font.Size = template.H1FontSize;
            h1.Font.Bold = template.H1Bold ? 1 : 0;
            h1.ParagraphFormat.Alignment = (Word.WdParagraphAlignment)template.H1Alignment;
            h1.ParagraphFormat.FirstLineIndent = 0;
            h1.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            h1.ParagraphFormat.KeepWithNext = -1; // -1 means True in Word VBA/Interop
            h1.ParagraphFormat.SpaceBefore = template.H1SpaceBefore;
            h1.ParagraphFormat.SpaceAfter = template.H1SpaceAfter;

            // Heading 2
            Word.Style h2 = doc.Styles[Word.WdBuiltinStyle.wdStyleHeading2];
            h2.Font.NameFarEast = template.H2FontNameFarEast;
            h2.Font.NameAscii = template.NormalFontNameAscii;
            h2.Font.Name = template.NormalFontNameAscii;
            h2.Font.Size = template.H2FontSize;
            h2.Font.Bold = template.H2Bold ? 1 : 0;
            h2.ParagraphFormat.Alignment = (Word.WdParagraphAlignment)template.H2Alignment;
            h2.ParagraphFormat.FirstLineIndent = 0;
            h2.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            h2.ParagraphFormat.KeepWithNext = -1;
            h2.ParagraphFormat.SpaceBefore = template.H2SpaceBefore;
            h2.ParagraphFormat.SpaceAfter = template.H2SpaceAfter;

            // Heading 3
            Word.Style h3 = doc.Styles[Word.WdBuiltinStyle.wdStyleHeading3];
            h3.Font.NameFarEast = template.H3FontNameFarEast;
            h3.Font.NameAscii = template.NormalFontNameAscii;
            h3.Font.Name = template.NormalFontNameAscii;
            h3.Font.Size = template.H3FontSize;
            h3.Font.Bold = template.H3Bold ? 1 : 0;
            h3.ParagraphFormat.Alignment = (Word.WdParagraphAlignment)template.H3Alignment;
            h3.ParagraphFormat.FirstLineIndent = 0;
            h3.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            h3.ParagraphFormat.KeepWithNext = -1;
        }

        private void RebuildSpecialStyles(Word.Document doc, FormattingTemplate template)
        {
            // Reference style
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
            refStyle.Font.Size = template.RefFontSize;
            refStyle.ParagraphFormat.FirstLineIndent = doc.Application.CentimetersToPoints(template.RefFirstLineIndentCm);
            refStyle.ParagraphFormat.LeftIndent = doc.Application.CentimetersToPoints(template.RefLeftIndentCm);
            
            // Caption (题注)
            Word.Style captionStyle = doc.Styles[Word.WdBuiltinStyle.wdStyleCaption];
            captionStyle.Font.NameFarEast = "宋体";
            captionStyle.Font.NameAscii = "Times New Roman";
            captionStyle.Font.Name = "Times New Roman";
            captionStyle.Font.Size = template.CaptionFontSize;
            captionStyle.Font.Bold = template.CaptionBold ? 1 : 0;
            captionStyle.ParagraphFormat.Alignment = (Word.WdParagraphAlignment)template.CaptionAlignment;
            captionStyle.ParagraphFormat.FirstLineIndent = 0;
            captionStyle.ParagraphFormat.CharacterUnitFirstLineIndent = 0;

            // Note (表注)
            Word.Style noteStyle;
            try
            {
                noteStyle = doc.Styles["表注_Auto"];
            }
            catch
            {
                noteStyle = doc.Styles.Add("表注_Auto", Word.WdStyleType.wdStyleTypeParagraph);
            }
            object baseStyleNote = Word.WdBuiltinStyle.wdStyleNormal;
            noteStyle.set_BaseStyle(ref baseStyleNote);
            noteStyle.Font.NameFarEast = template.NoteFontNameFarEast;
            noteStyle.Font.NameAscii = template.NormalFontNameAscii;
            noteStyle.Font.Name = template.NormalFontNameAscii;
            noteStyle.Font.Size = template.NoteFontSize;
            noteStyle.ParagraphFormat.Alignment = (Word.WdParagraphAlignment)template.NoteAlignment;
            noteStyle.ParagraphFormat.FirstLineIndent = 0;
            noteStyle.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
        }

        private string BoldText(bool isBold)
        {
            return isBold ? "加粗" : "不加粗";
        }
    }
}
