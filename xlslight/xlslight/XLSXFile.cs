using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace xlslight
{
    static class XLSXFile
    {
        public static void Write(string path, XSSFWorkbook input)
        {
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                if (fs != null)
                {
                    input.Write(fs);
                }
            }   
        }

        public static XSSFWorkbook Load(string path)
        {
            XSSFWorkbook result = null;
            using (var stream = new FileStream(path, FileMode.Open))
            {
                if (stream != null)
                {
                    stream.Position = 0;
                    result = new XSSFWorkbook(stream);
                }
            }

            return result;
        }
    }
}
