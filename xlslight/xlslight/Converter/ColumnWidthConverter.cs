using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace xlslight.Converter
{
    class ColumnWidthConverter : ConverterBase
    {
        protected override void ConvertXToL_Implement(ISheet xlsx, XLSLightSheet xlslight, ref ConvertContext convertContext)
        {
            int maxColumnIndex = 0;
            for (int rowCount = xlsx.FirstRowNum; rowCount <= xlsx.LastRowNum; rowCount++)
            {
                IRow row = xlsx.GetRow(rowCount);
                if (row.LastCellNum > maxColumnIndex)
                {
                    maxColumnIndex = row.LastCellNum;
                }
            }

            for (int columnCount = 0; columnCount <= maxColumnIndex; columnCount++)
            {
                int columnWidth = xlsx.GetColumnWidth(columnCount);
                xlslight.SetColumnWidth(columnCount, columnWidth);
            }
        }
        protected override void ConvertLToX_Implement(XLSLightSheet xlslight, ISheet xlsx, ref ConvertContext convertContext)
        {
            foreach (var columnWidthIter in xlslight.RowHeight)
            {
                int columnIndex = columnWidthIter.Key;
                int width = columnWidthIter.Value;
                xlsx.SetColumnWidth(columnIndex, width);
            }
        }
        protected override void ConvertLToX_Implement(XLSLightWorkbook xlslight, XSSFWorkbook xlsx, ref ConvertContext convertContext) { }
        protected override void ConvertXToL_Implement(XSSFWorkbook xlsx, XLSLightWorkbook xlslight, ref ConvertContext convertContext) { }
        protected override void ConvertLToX_Implement(XLSLightCell xlslight, ICell xlsx, ref ConvertContext convertContext) { }
        protected override void ConvertXToL_Implement(ICell xlsx, XLSLightCell xlslight, ref ConvertContext convertContext) { }
    }
}
