using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace xlslight.Converter
{
    class RowHeightConverter : ConverterBase
    {
        protected override void ConvertXToL_Implement(ISheet xlsx, XLSLightSheet xlslight)
        {
            for (int rowCount = xlsx.FirstRowNum; rowCount <= xlsx.LastRowNum; rowCount++)
            {
                IRow row = xlsx.GetRow(rowCount);
                if (row != null)
                {
                    xlslight.SetRowHeight(rowCount, row.Height);
                }
            }
        }
        protected override void ConvertLToX_Implement(XLSLightSheet xlslight, ISheet xlsx)
        {
            foreach (var rowHeightIter in xlslight.RowHeight)
            {
                int rowIndex = rowHeightIter.Key;
                short height = rowHeightIter.Value;
                IRow row = xlsx.GetRow(rowIndex);
                if (row != null)
                {
                    row.Height = height;
                }
            }
        }
    }
}
