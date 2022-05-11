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
            //using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            //{
            //    IWorkbook workbook = new XSSFWorkbook();
            //    ISheet newSheet = workbook.CreateSheet("Sheet1");

            //    for (int sheetCount = 0; sheetCount < input.NumberOfSheets; sheetCount++)
            //    {
            //        ISheet sheet = input.GetSheetAt(sheetCount);                    
            //        for (int rowCount = 0; rowCount < input.Count; rowCount++)
            //        {
            //            List<ICell> currentRow = input[rowCount];
            //            IRow newRow = newSheet.CreateRow(rowCount);
            //            for (int columnCount = 0; columnCount < currentRow.Count; columnCount++)
            //            {
            //                ICell originCell = currentRow[columnCount];
            //                ICell newCell = null;
            //                if (originCell == null)
            //                {
            //                    newCell = newRow.CreateCell(columnCount, CellType.Blank);
            //                    newCell.SetBlank();
            //                    continue;
            //                }

            //                newCell = newRow.CreateCell(columnCount, originCell.CellType);
            //                switch (originCell.CellType)
            //                {
            //                    case CellType.Blank:
            //                        newCell.SetBlank();
            //                        break;
            //                    case CellType.Formula:
            //                        newCell.SetCellFormula(originCell.CellFormula);
            //                        break;
            //                    case CellType.Error:
            //                        newCell.SetCellErrorValue(originCell.ErrorCellValue);
            //                        break;
            //                    case CellType.Numeric:
            //                        newCell.SetCellValue(originCell.NumericCellValue);
            //                        break;
            //                    case CellType.Boolean:
            //                        newCell.SetCellValue(originCell.BooleanCellValue);
            //                        break;
            //                    case CellType.String:
            //                        newCell.SetCellValue(originCell.StringCellValue);
            //                        break;
            //                }
            //            }
            //        }
            //    }

            //    workbook.Write(fs);
            //}
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
