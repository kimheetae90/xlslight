using System.Diagnostics;
using System.IO;
using System.Threading;
using xlslight.Converter;

namespace xlslight
{
    class XLSLight
    {
        private static string directory;
        private static string xlslightFileName;
        private static string xlslightPath;
        private static string fileName;
        private static string xlsxFileName;
        private static string xlsxPath;

        static void Main(string[] args)
        {
            Initialize();
            var xlsx = XLSXFile.Load(Directory.GetCurrentDirectory() + "\\xlsxTest.xlsx");
            var xlslight = ConvertController.ConvertXLSXToXLSLight(xlsx);            
            xlsx = ConvertController.ConvertXLSLightToXLSX(xlslight);
            XLSXFile.Write(Directory.GetCurrentDirectory() + "\\xlsxConvertTest.xlsx", xlsx);


            return;

            directory = Directory.GetCurrentDirectory();
            xlslightFileName = "xlslightTest.yaml";
            xlslightPath = directory + "\\" + xlslightFileName;            
            fileName = Path.GetFileNameWithoutExtension(xlslightPath);
            xlsxFileName = fileName + ".xlsx";
            xlsxPath = directory + "\\" + xlsxFileName;

            //if (args.Length > 0)
            {
                Thread startThread = new Thread(Start);
                startThread.Start();
                startThread.Join();
            }
        }

        private static void Start()
        {
            Initialize();
            if (CreateXLSX())
            {
                var p = CreateExcelProcess();
                p.Start();
                p.WaitForExit();

                SaveXLSLight();

                try
                {
                    File.Delete(xlsxPath);
                }
                catch (IOException)
                {
                    return;
                }
            }
        }

        private static void SaveXLSLight()
        {
            var xlsx = XLSXFile.Load(xlsxPath);
            if (xlsx == null)
                return;

            var xlsxlight = ConvertController.ConvertXLSXToXLSLight(xlsx);
            XLSLightFile.Write(xlslightPath, xlsxlight);
        }

        private static void Initialize()
        {
            //Todo : 옵션화
            ConvertController.converterContainer.converters.Add(new SheetNameConverter());
            ConvertController.converterContainer.converters.Add(new RowHeightConverter());
            ConvertController.converterContainer.converters.Add(new ColumnWidthConverter());
            ConvertController.converterContainer.converters.Add(new TypeValueConverter());
            ConvertController.converterContainer.converters.Add(new CellStyleConverter());
        }

        private static bool CreateXLSX()
        {
            var xlslightWorkbook = XLSLightFile.Load(xlslightPath);
            if (xlslightWorkbook == null)
                return false;

            var xlsx = ConvertController.ConvertXLSLightToXLSX(xlslightWorkbook);
            if (xlsx == null)
                return false;

            XLSXFile.Write(xlsxPath, xlsx);
            return true;
        }

        private static Process CreateExcelProcess()
        {
            Process p = new Process();
            p.StartInfo.UseShellExecute = true;
            p.StartInfo.FileName = xlsxPath;

            return p;
        }
    }
}
