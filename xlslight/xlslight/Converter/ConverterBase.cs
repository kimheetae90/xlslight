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

        public void PreConvertXToL(XSSFWorkbook xlsx, XLSLightWorkbook xlslight)
        {
            foreach (var converter in converters)
            {
                converter.PreConvertXToL(xlsx, xlslight);
            }
        }
        public void PreConvertLToX(XLSLightWorkbook xlslight, XSSFWorkbook xlsx)
        {
            foreach (var converter in converters)
            {
                converter.PreConvertLToX(xlslight, xlsx);
            }
        }

        public void PreConvertXToL(ISheet xlsx, XLSLightSheet xlslight)
        {
            foreach (var converter in converters)
            {
                converter.PreConvertXToL(xlsx, xlslight);
            }
        }
        public void PreConvertLToX(XLSLightSheet xlslight, ISheet xlsx)
        {
            foreach (var converter in converters)
            {
                converter.PreConvertLToX(xlslight, xlsx);
            }
        }

        public void ConvertXToL(XSSFWorkbook xlsx, XLSLightWorkbook xlslight)
        {
            foreach (var converter in converters)
            {
                converter.ConvertXToL(xlsx, xlslight);
            }
        }
        public void ConvertLToX(XLSLightWorkbook xlslight, XSSFWorkbook xlsx)
        {
            foreach (var converter in converters)
            {
                converter.ConvertLToX(xlslight, xlsx);
            }
        }

        public void ConvertXToL(ISheet xlsx, XLSLightSheet xlslight)
        {
            foreach (var converter in converters)
            {
                converter.ConvertXToL(xlsx, xlslight);
            }
        }
        public void ConvertLToX(XLSLightSheet xlslight, ISheet xlsx)
        {
            foreach (var converter in converters)
            {
                converter.ConvertLToX(xlslight, xlsx);
            }
        }

        public void ConvertXToL(ICell xlsx, XLSLightCell xlslight)
        {
            foreach (var converter in converters)
            {
                converter.ConvertXToL(xlsx, xlslight);
            }
        }
        public void ConvertLToX(XLSLightCell xlslight, ICell xlsx)
        {
            foreach (var converter in converters)
            {
                converter.ConvertLToX(xlslight, xlsx);
            }
        }
    }

    public abstract class ConverterBase
    {
        protected virtual void PreConvertXToL_Implement(XSSFWorkbook xlsx, XLSLightWorkbook xlslight) { }
        protected virtual void PreConvertLToX_Implement(XLSLightWorkbook xlslight, XSSFWorkbook xlsx) { }
        protected virtual void PreConvertXToL_Implement(ISheet xlsx, XLSLightSheet xlslight) { }
        protected virtual void PreConvertLToX_Implement(XLSLightSheet xlslight, ISheet xlsx) { }

        protected virtual void ConvertXToL_Implement(XSSFWorkbook xlsx, XLSLightWorkbook xlslight) { }
        protected virtual void ConvertLToX_Implement(XLSLightWorkbook xlslight, XSSFWorkbook xlsx) { }
        protected virtual void ConvertXToL_Implement(ISheet xlsx, XLSLightSheet xlslight) { }
        protected virtual void ConvertLToX_Implement(XLSLightSheet xlslight, ISheet xlsx) { }
        protected virtual void ConvertXToL_Implement(ICell xlsx, XLSLightCell xlslight) { }
        protected virtual void ConvertLToX_Implement(XLSLightCell xlslight, ICell xlsx) { }

        public void PreConvertXToL(XSSFWorkbook xlsx, XLSLightWorkbook xlslight)
        {
            if (xlsx != null && xlslight != null)
            {
                PreConvertXToL_Implement(xlsx, xlslight);
            }
        }
        public void PreConvertLToX(XLSLightWorkbook xlslight, XSSFWorkbook xlsx)
        {
            if (xlsx != null && xlslight != null)
            {
                PreConvertLToX_Implement(xlslight, xlsx);
            }
        }

        public void PreConvertXToL(ISheet xlsx, XLSLightSheet xlslight)
        {
            if (xlsx != null && xlslight != null)
            {
                PreConvertXToL_Implement(xlsx, xlslight);
            }
        }
        public void PreConvertLToX(XLSLightSheet xlslight, ISheet xlsx)
        {
            if (xlsx != null && xlslight != null)
            {
                PreConvertLToX_Implement(xlslight, xlsx);
            }
        }

        public void ConvertXToL(XSSFWorkbook xlsx, XLSLightWorkbook xlslight)
        {
            if (xlsx != null && xlslight != null)
            {
                ConvertXToL_Implement(xlsx, xlslight);
            }
        }
        public void ConvertLToX(XLSLightWorkbook xlslight, XSSFWorkbook xlsx)
        {
            if (xlsx != null && xlslight != null)
            {
                ConvertLToX_Implement(xlslight, xlsx);
            }
        }
        
        public void ConvertXToL(ISheet xlsx, XLSLightSheet xlslight)
        {
            if (xlsx != null && xlslight != null)
            {
                ConvertXToL_Implement(xlsx, xlslight);
            }
        }
        public void ConvertLToX(XLSLightSheet xlslight, ISheet xlsx)
        {
            if (xlsx != null && xlslight != null)
            {
                ConvertLToX_Implement(xlslight, xlsx);
            }
        }

        public void ConvertXToL(ICell xlsx, XLSLightCell xlslight)
        {
            if (xlsx != null && xlslight != null)
            {
                ConvertXToL_Implement(xlsx, xlslight);
            }
        }
        public void ConvertLToX(XLSLightCell xlslight, ICell xlsx)
        {
            if (xlsx != null && xlslight != null)
            {
                ConvertLToX_Implement(xlslight, xlsx);
            }
        }
    }
}
