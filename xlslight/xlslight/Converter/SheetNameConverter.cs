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
        protected override void ConvertLToX_Implement(XLSLightWorkbook xlslight, XSSFWorkbook xlsx, ref ConvertContext convertContext)
        {
            for (int index = 0; index < xlslight.sheets.Length; index++)
            {
                XLSLightSheet xlslightSheet = xlslight.GetSheet(index);
                if(xlslightSheet == null)
                {
                    continue;
                }

                xlsx.SetSheetName(index, xlslightSheet.name);
            }
        }

        protected override void ConvertXToL_Implement(ISheet xlsx, XLSLightSheet xlslight, ref ConvertContext convertContext)
        {
            xlslight.name = xlsx.SheetName;
        }

        protected override void ConvertXToL_Implement(XSSFWorkbook xlsx, XLSLightWorkbook xlslight, ref ConvertContext convertContext) {}
        protected override void ConvertLToX_Implement(XLSLightSheet xlslight, ISheet xlsx, ref ConvertContext convertContext) {}
        protected override void ConvertLToX_Implement(XLSLightCell xlslight, ICell xlsx, ref ConvertContext convertContext) {}
        protected override void ConvertXToL_Implement(ICell xlsx, XLSLightCell xlslight, ref ConvertContext convertContext) {}
    }
}
