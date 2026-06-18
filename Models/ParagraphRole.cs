namespace WordTool.Models
{
    public enum ParagraphRole
    {
        Unknown,
        Normal,
        Heading1,
        Heading2,
        Heading3,
        Heading4,
        Reference,
        Caption,    // 图表题注
        TableText,  // 表格内正文
        TableNote   // 表注
    }
}
