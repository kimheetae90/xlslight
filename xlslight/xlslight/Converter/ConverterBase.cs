using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlslight.Converter
{
    public class ConverterContainer
    {
        public List<ConverterBase> converters = new List<ConverterBase>();

        public void ConvertWorkBook_XToL(XSSFWorkbook xlsx, XLSLightWorkbook xlslight)
        {
            foreach (var converter in converters)
            {
                converter.ConvertXToL(xlsx, xlslight);
            }
        }
        public void ConvertWorkBook_LToX(XLSLightWorkbook xlslight, XSSFWorkbook xlsx)
        {
            foreach (var converter in converters)
            {
                converter.ConvertLToX(xlslight, xlsx);
            }
        }

        public void ConvertSheet_XToL(ISheet xlsx, XLSLightSheet xlslight)
        {
            foreach (var converter in converters)
            {
                converter.ConvertXToL(xlsx, xlslight);
            }
        }
        public void ConvertSheet_LToX(XLSLightSheet xlslight, ISheet xlsx)
        {
            foreach (var converter in converters)
            {
                converter.ConvertLToX(xlslight, xlsx);
            }
        }

        public void ConvertCell_XToL(ICell xlsx, XLSLightCell xlslight)
        {
            foreach (var converter in converters)
            {
                converter.ConvertXToL(xlsx, xlslight);
            }
        }
        public void ConvertCell_LToX(XLSLightCell xlslight, ICell xlsx)
        {
            foreach (var converter in converters)
            {
                converter.ConvertLToX(xlslight, xlsx);
            }
        }
    }

    public abstract class ConverterBase
    {
        protected abstract void ConvertXToL_Implement(XSSFWorkbook xlsx, XLSLightWorkbook xlslight);
        public void ConvertXToL(XSSFWorkbook xlsx, XLSLightWorkbook xlslight)
        {
            if (xlsx != null && xlslight != null)
            {
                ConvertXToL_Implement(xlsx, xlslight);
            }
        }
        protected abstract void ConvertLToX_Implement(XLSLightWorkbook xlslight, XSSFWorkbook xlsx);
        public void ConvertLToX(XLSLightWorkbook xlslight, XSSFWorkbook xlsx)
        {
            if (xlsx != null && xlslight != null)
            {
                ConvertLToX_Implement(xlslight, xlsx);
            }
        }

        protected abstract void ConvertXToL_Implement(ISheet xlsx, XLSLightSheet xlslight);
        public void ConvertXToL(ISheet xlsx, XLSLightSheet xlslight)
        {
            if (xlsx != null && xlslight != null)
            {
                ConvertXToL_Implement(xlsx, xlslight);
            }
        }
        protected abstract void ConvertLToX_Implement(XLSLightSheet xlslight, ISheet xlsx);
        public void ConvertLToX(XLSLightSheet xlslight, ISheet xlsx)
        {
            if (xlsx != null && xlslight != null)
            {
                ConvertLToX_Implement(xlslight, xlsx);
            }
        }

        protected abstract void ConvertXToL_Implement(ICell xlsx, XLSLightCell xlslight);
        public void ConvertXToL(ICell xlsx, XLSLightCell xlslight)
        {
            if (xlsx != null && xlslight != null)
            {
                ConvertXToL_Implement(xlsx, xlslight);
            }
        }
        protected abstract void ConvertLToX_Implement(XLSLightCell xlslight, ICell xlsx);
        public void ConvertLToX(XLSLightCell xlslight, ICell xlsx)
        {
            if (xlsx != null && xlslight != null)
            {
                ConvertLToX_Implement(xlslight, xlsx);
            }
        }
    }
}
