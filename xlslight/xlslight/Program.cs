using System.Collections.Generic;
using xlslight.Converter;

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

            ConvertController.converterContainer.converters.Add(new SheetNameConverter());
            ConvertController.converterContainer.converters.Add(new RowHeightConverter());
            ConvertController.converterContainer.converters.Add(new ColumnWidthConverter());
            ConvertController.converterContainer.converters.Add(new TypeValueConverter());

            var input = XLSXFile.Load(rxlspath);
            var yaml = ConvertController.ConvertXLSXToXLSLight(input);
            System.Threading.Tasks.Task task = XLSLightFile.WriteAsync(wyamlpath, yaml);

            var xlsx = ConvertController.ConvertXLSLightToXLSX(yaml);
            XLSXFile.Write(wxlspath, xlsx);
        }
    }
}
