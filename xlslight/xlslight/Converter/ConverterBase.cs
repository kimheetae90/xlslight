using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlslight.Converter
{
    public struct ConvertContext
    {
        public int sheetIndex;
        public int xOffset;
        public int yOffset;
    }

    public class ConverterContainer
    {
        private List<ConverterBase> converters;

        public void ConvertWorkBook_XToL(XSSFWorkbook xlsx, XLSLightWorkbook xlslight, ref ConvertContext convertContext)
        {
            foreach (var converter in converters)
            {
                converter.ConvertXToL(xlsx, xlslight, ref convertContext);
            }
        }
        public void ConvertWorkBook_LToX(XLSLightWorkbook xlslight, XSSFWorkbook xlsx, ref ConvertContext convertContext)
        {
            foreach (var converter in converters)
            {
                converter.ConvertLToX(xlslight, xlsx, ref convertContext);
            }
        }

        public void ConvertSheet_XToL(ISheet xlsx, XLSLightSheet xlslight, ref ConvertContext convertContext)
        {
            foreach (var converter in converters)
            {
                converter.ConvertXToL(xlsx, xlslight, ref convertContext);
            }
        }
        public void ConvertSheet_LToX(XLSLightSheet xlslight, ISheet xlsx, ref ConvertContext convertContext)
        {
            foreach (var converter in converters)
            {
                converter.ConvertLToX(xlslight, xlsx, ref convertContext);
            }
        }

        public void ConvertCell_XToL(ICell xlsx, XLSLightCell xlslight, ref ConvertContext convertContext)
        {
            foreach (var converter in converters)
            {
                converter.ConvertXToL(xlsx, xlslight, ref convertContext);
            }
        }
        public void ConvertCell_LToX(XLSLightCell xlslight, ICell xlsx, ref ConvertContext convertContext)
        {
            foreach (var converter in converters)
            {
                converter.ConvertLToX(xlslight, xlsx, ref convertContext);
            }
        }
    }

    public abstract class ConverterBase
    {
        protected abstract void ConvertXToL_Implement(XSSFWorkbook xlsx, XLSLightWorkbook xlslight, ref ConvertContext convertContext);
        public void ConvertXToL(XSSFWorkbook xlsx, XLSLightWorkbook xlslight, ref ConvertContext convertContext) 
        {
            if(xlsx != null && xlslight != null)
            {
                ConvertXToL_Implement(xlsx, xlslight, ref convertContext);
            }
        }
        protected abstract void ConvertLToX_Implement(XLSLightWorkbook xlslight, XSSFWorkbook xlsx, ref ConvertContext convertContext);
        public void ConvertLToX(XLSLightWorkbook xlslight, XSSFWorkbook xlsx, ref ConvertContext convertContext)
        {
            if (xlsx != null && xlslight != null)
            {
                ConvertLToX_Implement(xlslight, xlsx, ref convertContext);
            }
        }

        protected abstract void ConvertXToL_Implement(ISheet xlsx, XLSLightSheet xlslight, ref ConvertContext convertContext);
        public void ConvertXToL(ISheet xlsx, XLSLightSheet xlslight, ref ConvertContext convertContext)
        {
            if (xlsx != null && xlslight != null)
            {
                ConvertXToL_Implement(xlsx, xlslight, ref convertContext);
            }
        }
        protected abstract void ConvertLToX_Implement(XLSLightSheet xlslight, ISheet xlsx, ref ConvertContext convertContext);
        public void ConvertLToX(XLSLightSheet xlslight, ISheet xlsx, ref ConvertContext convertContext)
        {
            if (xlsx != null && xlslight != null)
            {
                ConvertLToX_Implement(xlslight, xlsx, ref convertContext);
            }
        }

        protected abstract void ConvertXToL_Implement(ICell xlsx, XLSLightCell xlslight, ref ConvertContext convertContext);
        public void ConvertXToL(ICell xlsx, XLSLightCell xlslight, ref ConvertContext convertContext)
        {
            if (xlsx != null && xlslight != null)
            {
                ConvertXToL_Implement(xlsx, xlslight, ref convertContext);
            }
        }
        protected abstract void ConvertLToX_Implement(XLSLightCell xlslight, ICell xlsx, ref ConvertContext convertContext);
        public void ConvertLToX(XLSLightCell xlslight, ICell xlsx, ref ConvertContext convertContext)
        {
            if (xlsx != null && xlslight != null)
            {
                ConvertLToX_Implement(xlslight, xlsx, ref convertContext);
            }
        }
    }
}
