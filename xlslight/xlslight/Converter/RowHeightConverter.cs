using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace xlslight.Converter
{
    class RowHeightConverter : ConverterBase
    {
        protected override void ConvertXToL_Implement(ISheet xlsx, XLSLightSheet xlslight, ref ConvertContext convertContext)
        {
            for(int rowCount = xlsx.FirstRowNum ; rowCount <= xlsx.LastRowNum; rowCount++)
            {
                IRow row = xlsx.GetRow(rowCount);
                xlslight.SetRowHeight(rowCount, row.Height);
            }
        }
        protected override void ConvertLToX_Implement(XLSLightSheet xlslight, ISheet xlsx, ref ConvertContext convertContext)
        {
            foreach(var rowHeightIter in xlslight.RowHeight)
            {
                int rowIndex = rowHeightIter.Key;
                short height = rowHeightIter.Value;
                IRow row = xlsx.GetRow(rowIndex);
                row.Height = height;
            }
        }
        protected override void ConvertLToX_Implement(XLSLightWorkbook xlslight, XSSFWorkbook xlsx, ref ConvertContext convertContext) { }
        protected override void ConvertXToL_Implement(XSSFWorkbook xlsx, XLSLightWorkbook xlslight, ref ConvertContext convertContext) { }
        protected override void ConvertLToX_Implement(XLSLightCell xlslight, ICell xlsx, ref ConvertContext convertContext) {}
        protected override void ConvertXToL_Implement(ICell xlsx, XLSLightCell xlslight, ref ConvertContext convertContext) {}
    }
}
