using System.Collections.Generic;
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
                var w = CreateFileSystemWatcher();
                var p = CreateExcelProcess();
                p.Start();
                p.WaitForExit();

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

        private static void OnChanged(object sender, FileSystemEventArgs e)
        {
        }

        private static void Initialize()
        {
            ConvertController.converterContainer.converters.Add(new SheetNameConverter());
            ConvertController.converterContainer.converters.Add(new RowHeightConverter());
            ConvertController.converterContainer.converters.Add(new ColumnWidthConverter());
            ConvertController.converterContainer.converters.Add(new TypeValueConverter());
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

        private static FileSystemWatcher CreateFileSystemWatcher()
        {
            var watcher = new FileSystemWatcher();
            watcher.Path = directory;
            watcher.NotifyFilter = NotifyFilters.CreationTime
                                 | NotifyFilters.FileName
                                 | NotifyFilters.LastAccess
                                 | NotifyFilters.LastWrite
                                 | NotifyFilters.Size;

            watcher.Filter = "*.xlsx";
            watcher.IncludeSubdirectories = true;
            watcher.Changed += OnChanged;
            watcher.EnableRaisingEvents = true;

            return watcher;
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
