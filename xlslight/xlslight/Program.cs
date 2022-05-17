using System.Collections.Generic;

namespace xlslight
{
    class Program
    {
        static void Main(string[] args)
        {
            string rxlspath = System.IO.Directory.GetCurrentDirectory() + "\\testRead.xlsx";
            string ryamlpath = System.IO.Directory.GetCurrentDirectory() + "\\testRead.yaml";
            string wxlspath = System.IO.Directory.GetCurrentDirectory() + "\\testWrite.xlsx";
            string wyamlpath = System.IO.Directory.GetCurrentDirectory() + "\\testWrite.yaml";

            var input = XLSXFile.Load(rxlspath);
            var yaml = ConvertController.ConvertXLSXToXLSLight(input);
            System.Threading.Tasks.Task task = XLSLightFile.WriteAsync(wyamlpath, yaml);

            var xlsx = ConvertController.ConvertXLSLightToXLSX(yaml);
            XLSXFile.Write(wxlspath, xlsx);
        }
    }
}
