using System.Collections.Generic;

namespace xlslight
{
    class Program
    {
        static void Main(string[] args)
        {
            string rpath = System.IO.Directory.GetCurrentDirectory() + "\\testRead.xlsx";
            string wpath = System.IO.Directory.GetCurrentDirectory() + "\\testwrite.yaml";

            //var input = YamlFile.Load(rpath);

            var input = XLSXFile.Load(rpath);
            XLSLightWorkbook yaml = XLSConverter.ConvertXLSXToYaml(input);
            System.Threading.Tasks.Task task = XLSLightFile.WriteAsync(wpath, yaml);   
        }
    }
}
