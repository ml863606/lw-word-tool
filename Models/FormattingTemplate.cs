using System;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Xml.Serialization;

namespace WordTool.Models
{
    public enum PaperSizeOption
    {
        A4,
        Letter
    }

    public enum ParagraphAlignmentOption
    {
        左对齐 = 0,
        居中对齐 = 1,
        右对齐 = 2
    }

    public enum RowAlignmentOption
    {
        左对齐 = 0,
        居中对齐 = 1,
        右对齐 = 2
    }

    public enum LineSpacingOption
    {
        单倍行距 = 0,
        一点五倍行距 = 3,
        多倍行距 = 5
    }

    public class FontSizeTypeConverter : TypeConverter
    {
        private static readonly string[] Names = { "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五", "六号", "小六", "七号", "八号" };
        private static readonly float[] Points = { 42.0f, 36.0f, 26.0f, 24.0f, 22.0f, 18.0f, 16.0f, 15.0f, 14.0f, 12.0f, 10.5f, 9.0f, 7.5f, 6.5f, 5.5f, 5.0f };

        public override bool GetStandardValuesSupported(ITypeDescriptorContext context)
        {
            return true;
        }

        public override StandardValuesCollection GetStandardValues(ITypeDescriptorContext context)
        {
            return new StandardValuesCollection(Names);
        }

        public override bool GetStandardValuesExclusive(ITypeDescriptorContext context)
        {
            return false;
        }

        public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType)
        {
            if (sourceType == typeof(string)) return true;
            return base.CanConvertFrom(context, sourceType);
        }

        public override bool CanConvertTo(ITypeDescriptorContext context, Type destinationType)
        {
            if (destinationType == typeof(string)) return true;
            return base.CanConvertTo(context, destinationType);
        }

        public override object ConvertFrom(ITypeDescriptorContext context, CultureInfo culture, object value)
        {
            var strValue = value as string;
            if (strValue != null)
            {
                int idx = Array.IndexOf(Names, strValue.Trim());
                if (idx >= 0) return Points[idx];

                if (float.TryParse(strValue, out float pts)) return pts;
            }
            return base.ConvertFrom(context, culture, value);
        }

        public override object ConvertTo(ITypeDescriptorContext context, CultureInfo culture, object value, Type destinationType)
        {
            if (destinationType == typeof(string) && value is float)
            {
                float pts = (float)value;
                int idx = Array.IndexOf(Points, pts);
                if (idx >= 0) return Names[idx];
                return pts.ToString("0.0");
            }
            return base.ConvertTo(context, culture, value, destinationType);
        }
    }

    [Serializable]
    public class FormattingTemplate
    {
        [Browsable(false)]
        public string Name { get; set; }

        // 1. 页面设置
        [Category("1. 页面与版面设置")]
        [DisplayName("纸张大小")]
        [Description("支持 A4 或 Letter")]
        public PaperSizeOption PaperSize { get; set; }

        [Category("1. 页面与版面设置")]
        [DisplayName("上边距 (厘米)")]
        [Description("页面上边缘的距离")]
        public float TopMarginCm { get; set; }

        [Category("1. 页面与版面设置")]
        [DisplayName("下边距 (厘米)")]
        [Description("页面下边缘的距离")]
        public float BottomMarginCm { get; set; }

        [Category("1. 页面与版面设置")]
        [DisplayName("左边距 (厘米)")]
        [Description("页面左边缘的距离")]
        public float LeftMarginCm { get; set; }

        [Category("1. 页面与版面设置")]
        [DisplayName("右边距 (厘米)")]
        [Description("页面右边缘的距离")]
        public float RightMarginCm { get; set; }


        // 2. 正文样式
        [Category("2. 正文基础样式")]
        [DisplayName("中文字体")]
        [Description("正文中文字体名称，如 宋体")]
        public string NormalFontNameFarEast { get; set; }

        [Category("2. 正文基础样式")]
        [DisplayName("西文字体")]
        [Description("正文西文/数字字体名称，如 Times New Roman")]
        public string NormalFontNameAscii { get; set; }

        [Category("2. 正文基础样式")]
        [DisplayName("字号")]
        [Description("支持下拉选择字号，或手动输入数字字号")]
        [TypeConverter(typeof(FontSizeTypeConverter))]
        public float NormalFontSize { get; set; }

        [Category("2. 正文基础样式")]
        [DisplayName("行间距类型")]
        [Description("单倍行距、一点五倍行距或多倍行距")]
        public LineSpacingOption NormalLineSpacingRule { get; set; }

        [Category("2. 正文基础样式")]
        [DisplayName("首行缩进 (字符)")]
        [Description("首行缩进字符数，如 2.0 表示首行空两格")]
        public float NormalFirstLineIndentChars { get; set; }


        // 3. 一级标题
        [Category("3. 一级标题样式")]
        [DisplayName("中文字体")]
        public string H1FontNameFarEast { get; set; }

        [Category("3. 一级标题样式")]
        [DisplayName("字号")]
        [TypeConverter(typeof(FontSizeTypeConverter))]
        public float H1FontSize { get; set; }

        [Category("3. 一级标题样式")]
        [DisplayName("是否加粗")]
        public bool H1Bold { get; set; }

        [Category("3. 一级标题样式")]
        [DisplayName("对齐方式")]
        public ParagraphAlignmentOption H1Alignment { get; set; }

        [Category("3. 一级标题样式")]
        [DisplayName("段前间距 (磅)")]
        public float H1SpaceBefore { get; set; }

        [Category("3. 一级标题样式")]
        [DisplayName("段后间距 (磅)")]
        public float H1SpaceAfter { get; set; }


        // 二级标题
        [Category("4. 二级标题样式")]
        [DisplayName("中文字体")]
        public string H2FontNameFarEast { get; set; }

        [Category("4. 二级标题样式")]
        [DisplayName("字号")]
        [TypeConverter(typeof(FontSizeTypeConverter))]
        public float H2FontSize { get; set; }

        [Category("4. 二级标题样式")]
        [DisplayName("是否加粗")]
        public bool H2Bold { get; set; }

        [Category("4. 二级标题样式")]
        [DisplayName("对齐方式")]
        public ParagraphAlignmentOption H2Alignment { get; set; }

        [Category("4. 二级标题样式")]
        [DisplayName("段前间距 (磅)")]
        public float H2SpaceBefore { get; set; }

        [Category("4. 二级标题样式")]
        [DisplayName("段后间距 (磅)")]
        public float H2SpaceAfter { get; set; }


        // 三级标题
        [Category("5. 三级标题样式")]
        [DisplayName("中文字体")]
        public string H3FontNameFarEast { get; set; }

        [Category("5. 三级标题样式")]
        [DisplayName("字号")]
        [TypeConverter(typeof(FontSizeTypeConverter))]
        public float H3FontSize { get; set; }

        [Category("5. 三级标题样式")]
        [DisplayName("是否加粗")]
        public bool H3Bold { get; set; }

        [Category("5. 三级标题样式")]
        [DisplayName("对齐方式")]
        public ParagraphAlignmentOption H3Alignment { get; set; }


        // 特殊样式
        [Category("6. 参考文献与题注")]
        [DisplayName("参考文献字号")]
        [TypeConverter(typeof(FontSizeTypeConverter))]
        public float RefFontSize { get; set; }

        [Category("6. 参考文献与题注")]
        [DisplayName("参考文献悬挂缩进 (厘米)")]
        [Description("通常设为负值（如 -0.74），表示首行外凸，第二行起缩进")]
        public float RefFirstLineIndentCm { get; set; }

        [Category("6. 参考文献与题注")]
        [DisplayName("参考文献左侧缩进 (厘米)")]
        [Description("配合悬挂缩进使用，如 0.74")]
        public float RefLeftIndentCm { get; set; }

        [Category("6. 参考文献与题注")]
        [DisplayName("图表题注字号")]
        [TypeConverter(typeof(FontSizeTypeConverter))]
        public float CaptionFontSize { get; set; }

        [Category("6. 参考文献与题注")]
        [DisplayName("图表题注是否加粗")]
        public bool CaptionBold { get; set; }

        [Category("6. 参考文献与题注")]
        [DisplayName("图表题注对齐方式")]
        public ParagraphAlignmentOption CaptionAlignment { get; set; }

        [Category("6. 参考文献与题注")]
        [DisplayName("表注中文字体")]
        public string NoteFontNameFarEast { get; set; }

        [Category("6. 参考文献与题注")]
        [DisplayName("表注字号")]
        [TypeConverter(typeof(FontSizeTypeConverter))]
        public float NoteFontSize { get; set; }

        [Category("6. 参考文献与题注")]
        [DisplayName("表注对齐方式")]
        public ParagraphAlignmentOption NoteAlignment { get; set; }


        // 图片样式
        [Category("7. 图样式")]
        [DisplayName("图片对齐方式")]
        public ParagraphAlignmentOption ImageAlignment { get; set; }

        [Category("7. 图样式")]
        [DisplayName("图片段前间距 (磅)")]
        public float ImageSpaceBefore { get; set; }

        [Category("7. 图样式")]
        [DisplayName("图片段后间距 (磅)")]
        public float ImageSpaceAfter { get; set; }

        // 表格样式
        [Category("8. 表格样式")]
        [DisplayName("表格宽度 (%)")]
        [Description("表格占页面可用宽度的百分比，默认 100 表示满宽")]
        public float TableWidthPercent { get; set; }

        [Category("8. 表格样式")]
        [DisplayName("表格对齐方式")]
        public RowAlignmentOption TableAlignment { get; set; }

        [Browsable(false)]
        [DisplayName("表格是否宽度自动适应窗口")]
        public bool TableAutoFit { get; set; }

        [Category("8. 表格样式")]
        [DisplayName("表格内容中文字体")]
        public string TableTextFontName { get; set; }

        [Category("8. 表格样式")]
        [DisplayName("表格内容字号")]
        [TypeConverter(typeof(FontSizeTypeConverter))]
        public float TableTextSize { get; set; }

        [Category("8. 表格样式")]
        [DisplayName("表格内容对齐方式")]
        [Description("表格单元格内文字的段落对齐方式，默认居中")]
        public ParagraphAlignmentOption TableTextAlignment { get; set; }

        [Category("8. 表格样式")]
        [DisplayName("表格内容行间距类型")]
        public LineSpacingOption TableTextLineSpacingRule { get; set; }

        [Category("8. 表格样式")]
        [DisplayName("表格内首行缩进 (厘米)")]
        [Description("表内文本一般不需要缩进，建议为 0.0")]
        public float TableTextFirstLineIndentCm { get; set; }

        [Category("8. 表格样式")]
        [DisplayName("是否启用三线表")]
        [Description("启用后表格将应用三线表格式（无竖线，上下边框较粗，标题行下方边框较细）")]
        public bool TableThreeLine { get; set; }

        [Browsable(false)]
        public float TableTopBottomBorderWidth { get; set; }

        [Category("8. 表格样式")]
        [DisplayName("三线表上边线宽度 (px)")]
        [Description("三线表第一条线，默认 1.5")]
        public float TableTopBorderWidth { get; set; }

        [Category("8. 表格样式")]
        [DisplayName("三线表中间线宽度 (px)")]
        [Description("标题行下方第二条线，默认 0.75")]
        public float TableHeaderBottomBorderWidth { get; set; }

        [Category("8. 表格样式")]
        [DisplayName("三线表下边线宽度 (px)")]
        [Description("三线表第三条线，默认 1.5")]
        public float TableBottomBorderWidth { get; set; }

        [Category("8. 表格样式")]
        [DisplayName("表格标题行是否加粗")]
        [Description("表格第一行（标题行）内容是否加粗")]
        public bool TableHeaderBold { get; set; }


        // 目录
        [Category("9. 目录辅助参数")]
        [DisplayName("目录包含级数")]
        [Description("支持生成 1 - 3 级大纲目录")]
        public int TocLevels { get; set; }

        [Category("9. 目录辅助参数")]
        [DisplayName("自动更新页码")]
        [Description("排版时若已有目录是否自动更新")]
        public bool TocAutoUpdate { get; set; }

        public FormattingTemplate()
        {
            TableWidthPercent = 100.0f;
            TableTextAlignment = ParagraphAlignmentOption.居中对齐;
            TableTopBorderWidth = 1.5f;
            TableHeaderBottomBorderWidth = 0.75f;
            TableBottomBorderWidth = 1.5f;
        }

        // 获取国内标准的“学位论文”默认模板 (GB/T 7713.1)
        public static FormattingTemplate GetDefaultThesisTemplate()
        {
            return new FormattingTemplate
            {
                Name = "学位论文模板 (GB/T 7713.1)",
                PaperSize = PaperSizeOption.A4,
                TopMarginCm = 3.0f,
                BottomMarginCm = 2.5f,
                LeftMarginCm = 2.5f,
                RightMarginCm = 2.5f,

                NormalFontNameFarEast = "宋体",
                NormalFontNameAscii = "Times New Roman",
                NormalFontSize = 12.0f, // 小四
                NormalLineSpacingRule = LineSpacingOption.一点五倍行距,
                NormalFirstLineIndentChars = 2.0f, // 首行缩进 2 字符

                H1FontNameFarEast = "黑体",
                H1FontSize = 16.0f, // 三号
                H1Bold = true,
                H1Alignment = ParagraphAlignmentOption.左对齐, // 大标题左对齐
                H1SpaceBefore = 12.0f,
                H1SpaceAfter = 12.0f,

                H2FontNameFarEast = "黑体",
                H2FontSize = 14.0f, // 四号
                H2Bold = true,
                H2Alignment = ParagraphAlignmentOption.左对齐,
                H2SpaceBefore = 6.0f,
                H2SpaceAfter = 6.0f,

                H3FontNameFarEast = "黑体",
                H3FontSize = 12.0f, // 小四
                H3Bold = true,
                H3Alignment = ParagraphAlignmentOption.左对齐,

                RefFontSize = 10.5f, // 五号
                RefFirstLineIndentCm = -0.74f, // 悬挂缩进
                RefLeftIndentCm = 0.74f,

                CaptionFontSize = 10.5f, // 五号
                CaptionBold = false,
                CaptionAlignment = ParagraphAlignmentOption.居中对齐,

                NoteFontNameFarEast = "宋体",
                NoteFontSize = 10.5f, // 五号
                NoteAlignment = ParagraphAlignmentOption.居中对齐,

                ImageAlignment = ParagraphAlignmentOption.居中对齐,
                ImageSpaceBefore = 6.0f,
                ImageSpaceAfter = 6.0f,

                TableWidthPercent = 100.0f,
                TableAlignment = RowAlignmentOption.居中对齐,
                TableAutoFit = true,
                TableThreeLine = true,
                TableTopBottomBorderWidth = 1.5f,
                TableTopBorderWidth = 1.5f,
                TableHeaderBottomBorderWidth = 0.75f,
                TableBottomBorderWidth = 1.5f,
                TableHeaderBold = true,
                TableTextFontName = "宋体",
                TableTextSize = 10.5f, // 五号
                TableTextAlignment = ParagraphAlignmentOption.居中对齐,
                TableTextLineSpacingRule = LineSpacingOption.单倍行距,
                TableTextFirstLineIndentCm = 0.0f,

                TocLevels = 3,
                TocAutoUpdate = true
            };
        }

        // 获取公文规范默认模板 (GB/T 9704-2012)
        public static FormattingTemplate GetDefaultOfficialDocumentTemplate()
        {
            return new FormattingTemplate
            {
                Name = "党政公文模板 (GB/T 9704)",
                PaperSize = PaperSizeOption.A4,
                TopMarginCm = 3.7f,
                BottomMarginCm = 3.5f,
                LeftMarginCm = 2.8f,
                RightMarginCm = 2.6f,

                NormalFontNameFarEast = "仿宋_GB2312",
                NormalFontNameAscii = "Times New Roman",
                NormalFontSize = 16.0f, // 三号
                NormalLineSpacingRule = LineSpacingOption.多倍行距,
                NormalFirstLineIndentChars = 2.0f, // 首行缩进 2 字符

                H1FontNameFarEast = "方正小标宋简体",
                H1FontSize = 22.0f, // 二号
                H1Bold = true,
                H1Alignment = ParagraphAlignmentOption.居中对齐,
                H1SpaceBefore = 0.0f,
                H1SpaceAfter = 0.0f,

                H2FontNameFarEast = "楷体_GB2312",
                H2FontSize = 16.0f, // 三号
                H2Bold = true,
                H2Alignment = ParagraphAlignmentOption.左对齐,
                H2SpaceBefore = 0.0f,
                H2SpaceAfter = 0.0f,

                H3FontNameFarEast = "仿宋_GB2312",
                H3FontSize = 16.0f, // 三号
                H3Bold = true,
                H3Alignment = ParagraphAlignmentOption.左对齐,

                RefFontSize = 12.0f, // 小四
                RefFirstLineIndentCm = -0.74f,
                RefLeftIndentCm = 0.74f,

                CaptionFontSize = 12.0f,
                CaptionBold = false,
                CaptionAlignment = ParagraphAlignmentOption.居中对齐,

                NoteFontNameFarEast = "仿宋_GB2312",
                NoteFontSize = 12.0f, // 小四
                NoteAlignment = ParagraphAlignmentOption.居中对齐,

                ImageAlignment = ParagraphAlignmentOption.居中对齐,
                ImageSpaceBefore = 6.0f,
                ImageSpaceAfter = 6.0f,

                TableWidthPercent = 100.0f,
                TableAlignment = RowAlignmentOption.居中对齐,
                TableAutoFit = true,
                TableThreeLine = false,
                TableTopBottomBorderWidth = 1.5f,
                TableTopBorderWidth = 1.5f,
                TableHeaderBottomBorderWidth = 0.75f,
                TableBottomBorderWidth = 1.5f,
                TableHeaderBold = false,
                TableTextFontName = "仿宋_GB2312",
                TableTextSize = 12.0f, // 小四
                TableTextAlignment = ParagraphAlignmentOption.居中对齐,
                TableTextLineSpacingRule = LineSpacingOption.单倍行距,
                TableTextFirstLineIndentCm = 0.0f,

                TocLevels = 3,
                TocAutoUpdate = true
            };
        }

        // 保存模板为 XML
        public void Save(string filePath)
        {
            var serializer = new XmlSerializer(typeof(FormattingTemplate));
            using (var writer = new StreamWriter(filePath))
            {
                serializer.Serialize(writer, this);
            }
        }

        // 从 XML 加载
        public static FormattingTemplate Load(string filePath)
        {
            var serializer = new XmlSerializer(typeof(FormattingTemplate));
            using (var reader = new StreamReader(filePath))
            {
                return (FormattingTemplate)serializer.Deserialize(reader);
            }
        }
    }
}
