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
            List<List<string>> result = new List<List<string>>();
            ISheet sheet;

            string path = System.IO.Directory.GetCurrentDirectory() + "\\test.xlsx";

            using (var stream = new FileStream(path, FileMode.Open))
            {
                stream.Position = 0;
                XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
                sheet = xssWorkbook.GetSheetAt(0);

                for (int rowCount = 0; rowCount < sheet.LastRowNum; rowCount++)
                {
                    IRow row = sheet.GetRow(rowCount);
                    List<string> rowResult = new List<string>();
                    for (int columnCount = 0; columnCount < row.LastCellNum; columnCount++)
                    {
                        ICell cell = row.GetCell(columnCount);
                        switch (cell.CellType)
                        {
                            case CellType.Formula:
                                rowResult.Add(cell.CellFormula);
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
                        }
                    }
                    result.Add(rowResult);
                }
            }
        }
    }
}
