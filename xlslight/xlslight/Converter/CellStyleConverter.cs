using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;

namespace xlslight.Converter
{
    internal class CellStyleConverter : ConverterBase
    {
        protected override void PreConvertXToL_Implement(XSSFWorkbook xlsx, XLSLightWorkbook xlslight)
        {
            var cellstyles = new List<XLSLightCellStyle>();
            for (short i = 0; i < xlsx.NumCellStyles; i++)
            {
                XSSFCellStyle cellstyle = (XSSFCellStyle)xlsx.GetCellStyleAt(i);
                if(cellstyle != null)
                {
                    XLSLightCellStyle xlslightCellStyle = new XLSLightCellStyle();
                    xlslightCellStyle.IsHidden = cellstyle.IsHidden;
                    xlslightCellStyle.IsLocked = cellstyle.IsLocked;
                    xlslightCellStyle.FillPattern = (int)cellstyle.FillPattern;
                    xlslightCellStyle.HorizontalAlignment = (int)cellstyle.Alignment;
                    xlslightCellStyle.VerticalAlignment = (int)cellstyle.VerticalAlignment;
                    xlslightCellStyle.FillForegroundColor = cellstyle.FillForegroundColor;
                    xlslightCellStyle.FillBackgroundColor = cellstyle.FillBackgroundColor;
                    xlslightCellStyle.FontIndex = cellstyle.FontIndex;
                    cellstyles.Add(xlslightCellStyle);
                }
            }

            if (cellstyles.Count > 0)
            {
                xlslight.cellStyles = cellstyles;
            }

            var fonts = new List<XLSLightFont>();
            for (short i = 0; i < xlsx.NumberOfFonts; i++)
            {
                XSSFFont font = (XSSFFont)xlsx.GetFontAt(i);
                if (font != null)
                {
                    XLSLightFont xlsLightFont = new XLSLightFont();
                    xlsLightFont.FontName = font.FontName;
                    xlsLightFont.FontHeightInPoints = font.FontHeightInPoints;
                    xlsLightFont.IsStrickout = font.IsStrikeout;
                    xlsLightFont.IsItalic = font.IsItalic;
                    xlsLightFont.IsBold = font.IsBold;
                    xlsLightFont.FontColor = font.Color;
                    fonts.Add(xlsLightFont);
                }
            }

            if (fonts.Count > 0)
            {
                xlslight.fonts = fonts;
            }
        }

        protected override void PreConvertLToX_Implement(XLSLightWorkbook xlslight, XSSFWorkbook xlsx)
        {
            short fontCount = xlsx.NumberOfFonts;
            for (short i = 0; i < xlslight.fonts.Count; i++)
            {
                XSSFFont font = (XSSFFont)(fontCount - 1 < i ?
                    xlsx.CreateFont() : xlsx.GetFontAt(i));

                var xlslightFont = xlslight.fonts[i];

                font.FontName = xlslightFont.FontName;
                font.FontHeightInPoints = xlslightFont.FontHeightInPoints;
                font.IsStrikeout = xlslightFont.IsStrickout;
                font.IsItalic = xlslightFont.IsItalic;
                font.IsBold = xlslightFont.IsBold;
                font.Color = xlslightFont.FontColor;
            }
            
            int cellStyleCount = xlsx.NumCellStyles;
            for (short i = 0; i < xlslight.cellStyles.Count; i++)
            {
                XSSFCellStyle cellStyle = (XSSFCellStyle)(cellStyleCount - 1 < i ?
                    xlsx.CreateCellStyle() : xlsx.GetCellStyleAt(i));

                var xlslightCellStyle = xlslight.cellStyles[i];

                cellStyle.IsHidden = xlslightCellStyle.IsHidden;
                cellStyle.IsLocked = xlslightCellStyle.IsLocked;
                cellStyle.Alignment = (HorizontalAlignment)xlslightCellStyle.HorizontalAlignment;
                cellStyle.FillPattern = (FillPattern)xlslightCellStyle.FillPattern;
                cellStyle.VerticalAlignment = (VerticalAlignment)xlslightCellStyle.VerticalAlignment;
                cellStyle.FillForegroundColor = xlslightCellStyle.FillForegroundColor;
                cellStyle.FillBackgroundColor = xlslightCellStyle.FillBackgroundColor;
                var font = xlsx.GetFontAt(xlslightCellStyle.FontIndex);
                cellStyle.SetFont(font);
            }
        }

        protected override void ConvertXToL_Implement(ICell xlsx, XLSLightCell xlslight) 
        {
            if(xlsx.CellStyle != null)
            {
                xlslight.SetCellStyleIndex(xlsx.CellStyle.Index);
            }
        }
                
        protected override void ConvertLToX_Implement(XLSLightCell xlslight, ICell xlsx) 
        {
            short cellStyleIndex = xlslight.GetCellStyleIndex();
            if (cellStyleIndex >= 0)
            {
                var cellStyle = xlsx.Sheet.Workbook.GetCellStyleAt(cellStyleIndex);
                xlsx.CellStyle = cellStyle;
            }
        }
    }
}
