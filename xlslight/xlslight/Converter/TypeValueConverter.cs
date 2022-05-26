using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlslight.Converter
{
    class TypeValueConverter : ConverterBase
    {
        protected override void ConvertLToX_Implement(XLSLightCell xlslight, ICell xlsx)
        {
            var type = xlslight.GetCellType();
            var value = xlslight.GetValue();

            CellType cellType = (CellType)type;
            xlsx.SetCellType(cellType);
            switch (cellType)
            {
                case CellType.Formula:
                    xlsx.SetCellFormula(value);
                    break;
                case CellType.Error:
                    byte err = 0;
                    byte.TryParse(value, out err);
                    xlsx.SetCellErrorValue(err);
                    break;
                case CellType.Numeric:
                    double numb = 0;
                    double.TryParse(value, out numb);
                    xlsx.SetCellValue(numb);
                    break;
                case CellType.Boolean:
                    if (value == "TRUE")
                    {
                        xlsx.SetCellValue(true);
                    }
                    else if (value == "FALSE")
                    {
                        xlsx.SetCellValue(false);
                    }
                    break;
                case CellType.String:
                    xlsx.SetCellValue(value);
                    break;
            }
        }
        protected override void ConvertXToL_Implement(ICell xlsx, XLSLightCell xlslight) 
        {
            xlslight.SetCellType((int)xlsx.CellType);
            switch (xlsx.CellType)
            {
                case CellType.Formula:
                    xlslight.SetValue(xlsx.CellFormula);
                    break;
                case CellType.Error:
                    xlslight.SetValue(xlsx.ErrorCellValue.ToString());
                    break;
                case CellType.Numeric:
                    xlslight.SetValue(xlsx.NumericCellValue.ToString());
                    break;
                case CellType.Boolean:
                    xlslight.SetValue(xlsx.BooleanCellValue ? "TRUE" : "FALSE");
                    break;
                case CellType.String:
                    xlslight.SetValue(xlsx.StringCellValue);
                    break;
            }
        }

        protected override void ConvertLToX_Implement(XLSLightWorkbook xlslight, XSSFWorkbook xlsx) {}
        protected override void ConvertXToL_Implement(ISheet xlsx, XLSLightSheet xlslight) {}
        protected override void ConvertXToL_Implement(XSSFWorkbook xlsx, XLSLightWorkbook xlslight) { }
        protected override void ConvertLToX_Implement(XLSLightSheet xlslight, ISheet xlsx) { }
    }
}
