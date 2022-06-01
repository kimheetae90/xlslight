using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlslight.Converter
{
    class SheetNameConverter : ConverterBase
    {
        //ISheet에 SheetName의 Setter가 제공되지 않아서 Workbook레벨에서 해결
        protected override void ConvertLToX_Implement(XLSLightWorkbook xlslight, XSSFWorkbook xlsx)
        {
            for (int index = 0; index < xlslight.sheets.Length; index++)
            {
                XLSLightSheet xlslightSheet = xlslight.GetSheet(index);
                if (xlslightSheet == null)
                {
                    continue;
                }

                xlsx.SetSheetName(index, xlslightSheet.name);
            }
        }

        protected override void ConvertXToL_Implement(ISheet xlsx, XLSLightSheet xlslight)
        {
            xlslight.name = xlsx.SheetName;
        }
    }
}
