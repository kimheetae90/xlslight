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
                input.Write(fs);
            }   
        }

        public static XSSFWorkbook Load(string path)
        {
            XSSFWorkbook result = new XSSFWorkbook();
            using (var stream = new FileStream(path, FileMode.Open))
            {
                stream.Position = 0;
                result = new XSSFWorkbook(stream);
            }

            return result;
        }
    }
}
