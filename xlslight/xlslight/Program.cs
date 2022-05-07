using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO;

namespace xlslight
{
    class Program
    {
        static void Main(string[] args)
        {
            string rpath = System.IO.Directory.GetCurrentDirectory() + "\\testRead.xlsx";
            string wpath = System.IO.Directory.GetCurrentDirectory() + "\\testwrite.xlsx";

            var input = Load(rpath);
            Write(wpath, input);
        }

        static void Write(string path, List<List<string>> input)
        {
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet newSheet = workbook.CreateSheet("Sheet1");

                for (int rowCount = 0; rowCount < input.Count; rowCount++)
                {
                    List<string> currentRow = input[rowCount];
                    IRow newRow = newSheet.CreateRow(rowCount);
                    for (int columnCount = 0; columnCount < currentRow.Count; columnCount++)
                    {
                        string data = currentRow[columnCount];
                        double value = 0;
                        if (data.StartsWith("="))
                        {
                            data = data.TrimStart('=');
                            newRow.CreateCell(columnCount).SetCellFormula(data);
                        }
                        else if(double.TryParse(data, out value))
                        {
                            newRow.CreateCell(columnCount).SetCellValue(value);
                        }
                        else
                        {
                            newRow.CreateCell(columnCount).SetCellValue(data);
                        }
                    }
                }

                workbook.Write(fs);
            }
        }

        static List<List<string>> Load(string path)
        {
            List<List<string>> result = new List<List<string>>();
            ISheet sheet;

            using (var stream = new FileStream(path, FileMode.Open))
            {
                stream.Position = 0;
                XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
                sheet = xssWorkbook.GetSheetAt(0);

                for (int rowCount = 0; rowCount <= sheet.LastRowNum; rowCount++)
                {
                    IRow row = sheet.GetRow(rowCount);
                    List<string> rowResult = new List<string>();
                    for (int columnCount = 0; columnCount < row.LastCellNum; columnCount++)
                    {
                        ICell cell = row.GetCell(columnCount);
                        switch (cell.CellType)
                        {
                            case CellType.Formula:
                                rowResult.Add("=" + cell.CellFormula);
                                break;
                            case CellType.Numeric:
                                rowResult.Add(cell.NumericCellValue.ToString());
                                break;
                            case CellType.String:
                                rowResult.Add(cell.StringCellValue);
                                break;
                            case CellType.Boolean:
                                rowResult.Add(cell.BooleanCellValue ? "TRUE" : "FALSE");
                                break;
                            case CellType.Blank:
                                rowResult.Add(string.Empty);
                                break;
                            case CellType.Error:
                                rowResult.Add(cell.ErrorCellValue.ToString());
                                break;
                        }
                    }
                    result.Add(rowResult);
                }
            }
            return result;
        }
    }
}
