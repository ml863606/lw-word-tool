using System.Collections.Generic;
using WordTool.Models;
using Word = Microsoft.Office.Interop.Word;

namespace WordTool.Formatters
{
    public class DocumentFormatter
    {
        public void ApplyFormatting(List<AnalyzedParagraph> paragraphs, System.Action<string> logger = null, System.Action checkStatus = null)
        {
            ApplyFormatting(paragraphs, null, logger, checkStatus);
        }

        public void ApplyFormatting(List<AnalyzedParagraph> paragraphs, FormattingTemplate template, System.Action<string> logger = null, System.Action checkStatus = null)
        {
            logger?.Invoke($"正在清除 {paragraphs.Count} 个段落的残留直接格式，并重新赋予规范样式...");
            int index = 0;
            foreach (var item in paragraphs)
            {
                checkStatus?.Invoke();
                index++;
                Word.Paragraph para = item.Paragraph;
                string originalStyle = GetStyleName(para);
                string text = TrimForLog(item.TextContent, 80);
                string styleName = "正文";
                
                // 1. 清除直接格式
                para.Range.Font.Reset();
                para.Range.ParagraphFormat.Reset();

                // 2. 根据角色应用样式
                switch (item.Role)
                {
                    case ParagraphRole.Heading1:
                        para.set_Style(Word.WdBuiltinStyle.wdStyleHeading1);
                        styleName = "一级标题 (标题 1)";
                        break;
                    case ParagraphRole.Heading2:
                        para.set_Style(Word.WdBuiltinStyle.wdStyleHeading2);
                        styleName = "二级标题 (标题 2)";
                        break;
                    case ParagraphRole.Heading3:
                        para.set_Style(Word.WdBuiltinStyle.wdStyleHeading3);
                        styleName = "三级标题 (标题 3)";
                        break;
                    case ParagraphRole.Reference:
                        try { para.set_Style("参考文献_Auto"); } catch { para.set_Style(Word.WdBuiltinStyle.wdStyleNormal); }
                        styleName = "参考文献 (参考文献_Auto)";
                        break;
                    case ParagraphRole.Caption:
                        para.set_Style(Word.WdBuiltinStyle.wdStyleCaption);
                        styleName = "图表题注 (题注)";
                        break;
                    case ParagraphRole.TableNote:
                        try { para.set_Style("表注_Auto"); } catch { para.set_Style(Word.WdBuiltinStyle.wdStyleNormal); }
                        styleName = "表注 (表注_Auto)";
                        break;
                    case ParagraphRole.Normal:
                    default:
                        para.set_Style(Word.WdBuiltinStyle.wdStyleNormal);
                        styleName = "正文";
                        break;
                }

                logger?.Invoke($"【正文排版】第 {index} 段 “{text}”：{RoleDisplayName(item.Role)} -> {styleName}；原样式: {originalStyle}；{BuildStyleSummary(item.Role, template)}");
            }
        }

        private string GetStyleName(Word.Paragraph para)
        {
            try
            {
                Word.Style style = para.get_Style() as Word.Style;
                if (style != null) return style.NameLocal;
            }
            catch { }

            return "未知";
        }

        private string RoleDisplayName(ParagraphRole role)
        {
            switch (role)
            {
                case ParagraphRole.Heading1: return "识别为一级标题";
                case ParagraphRole.Heading2: return "识别为二级标题";
                case ParagraphRole.Heading3: return "识别为三级标题";
                case ParagraphRole.Reference: return "识别为参考文献";
                case ParagraphRole.Caption: return "识别为图表题注";
                case ParagraphRole.TableNote: return "识别为表注";
                case ParagraphRole.Normal: return "识别为正文";
                default: return "识别为未知类型";
            }
        }

        private string BuildStyleSummary(ParagraphRole role, FormattingTemplate template)
        {
            if (template == null)
            {
                return "已清除直接格式并套用目标样式";
            }

            switch (role)
            {
                case ParagraphRole.Heading1:
                    return $"样式: {template.H1FontNameFarEast} {template.H1FontSize:0.#}pt，{BoldText(template.H1Bold)}，{template.H1Alignment}，段前 {template.H1SpaceBefore:0.#}pt，段后 {template.H1SpaceAfter:0.#}pt";
                case ParagraphRole.Heading2:
                    return $"样式: {template.H2FontNameFarEast} {template.H2FontSize:0.#}pt，{BoldText(template.H2Bold)}，{template.H2Alignment}，段前 {template.H2SpaceBefore:0.#}pt，段后 {template.H2SpaceAfter:0.#}pt";
                case ParagraphRole.Heading3:
                    return $"样式: {template.H3FontNameFarEast} {template.H3FontSize:0.#}pt，{BoldText(template.H3Bold)}，{template.H3Alignment}";
                case ParagraphRole.Reference:
                    return $"样式: 参考文献 {template.RefFontSize:0.#}pt，首行缩进 {template.RefFirstLineIndentCm:0.##}cm，左缩进 {template.RefLeftIndentCm:0.##}cm";
                case ParagraphRole.Caption:
                    return $"样式: 题注 {template.CaptionFontSize:0.#}pt，{BoldText(template.CaptionBold)}，{template.CaptionAlignment}";
                case ParagraphRole.TableNote:
                    return $"样式: 表注 {template.NoteFontNameFarEast} {template.NoteFontSize:0.#}pt，{template.NoteAlignment}";
                case ParagraphRole.Normal:
                default:
                    return $"样式: {template.NormalFontNameFarEast}/{template.NormalFontNameAscii} {template.NormalFontSize:0.#}pt，{template.NormalLineSpacingRule}，首行缩进 {template.NormalFirstLineIndentChars:0.#} 字符";
            }
        }

        private string BoldText(bool isBold)
        {
            return isBold ? "加粗" : "不加粗";
        }

        private string TrimForLog(string text, int maxLength)
        {
            if (string.IsNullOrWhiteSpace(text)) return "(空文本)";
            string normalized = text.Replace("\r", "").Replace("\n", "").Replace("\t", " ").Trim();
            if (normalized.Length <= maxLength) return normalized;
            return normalized.Substring(0, maxLength) + "...";
        }
    }
}
