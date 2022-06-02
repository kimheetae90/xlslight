using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace xlslight.Converter
{
    class ColumnWidthConverter : ConverterBase
    {
        protected override void ConvertXToL_Implement(ISheet xlsx, XLSLightSheet xlslight)
        {
            int maxColumnIndex = 0;
            for (int rowCount = xlsx.FirstRowNum; rowCount <= xlsx.LastRowNum; rowCount++)
            {
                IRow row = xlsx.GetRow(rowCount);
                if (row != null &&
                    row.LastCellNum > maxColumnIndex)
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
        protected override void ConvertLToX_Implement(XLSLightSheet xlslight, ISheet xlsx)
        {
            foreach (var columnWidthIter in xlslight.columnWidth)
            {
                int columnIndex = (int)columnWidthIter.Key;
                int width = (int)columnWidthIter.Value;
                xlsx.SetColumnWidth(columnIndex, width);
            }
        }
    }
}
