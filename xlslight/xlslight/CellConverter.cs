using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlslight
{
    public class XLSConverter
    {
        public static XLSLightWorkbook ConvertXLSXToXLSLight(XSSFWorkbook xlsx)
        {
            var xlslight = new XLSLightWorkbook();
            int xOffset = -1, yOffset = -1;
            var xlslightSheets = new List<XLSLightSheet>();

            for (int sheetCount = 0; sheetCount < xlsx.NumberOfSheets; sheetCount++)
            {
                var xlslightSheet = new XLSLightSheet();
                var xlsxSheet = xlsx.GetSheetAt(sheetCount);
                xlslightSheet.name = xlsxSheet.SheetName;
                var xlslightCells = new List<XLSLightCell>();

                xOffset = yOffset = 0;
                for (int rowCount = 0; rowCount <= xlsxSheet.LastRowNum; rowCount++)
                {
                    var xlsxRow = xlsxSheet.GetRow(rowCount);

                    for (int columnCount = 0; columnCount < xlsxRow.LastCellNum; columnCount++)
                    {
                        var xlsxCell = xlsxRow.GetCell(columnCount);

                        if (xlsxCell == null || xlsxCell.CellType == CellType.Blank)
                        {
                            xOffset++;
                            continue;
                        }

                        var xlslightCell = new XLSLightCell();
                        xlslightCell.SetOffset(xOffset, yOffset);
                        xOffset = 0;
                        yOffset = 0;
                        xlslightCell.SetCellType((int)xlsxCell.CellType);                        
                        switch (xlsxCell.CellType)
                        {
                            case CellType.Formula:
                                xlslightCell.SetValue(xlsxCell.CellFormula);
                                break;
                            case CellType.Error:
                                xlslightCell.SetValue(xlsxCell.ErrorCellValue.ToString());
                                break;
                            case CellType.Numeric:
                                xlslightCell.SetValue(xlsxCell.NumericCellValue.ToString());
                                break;
                            case CellType.Boolean:
                                xlslightCell.SetValue(xlsxCell.BooleanCellValue ? "TRUE" : "FALSE");
                                break;
                            case CellType.String:
                                xlslightCell.SetValue(xlsxCell.StringCellValue);
                                break;
                        }

                        xlslightCells.Add(xlslightCell);
                        xOffset++;
                    }

                    xOffset = 0;
                    yOffset++;
                }
                xlslightSheet.cells = xlslightCells.ToArray();
                xlslightSheets.Add(xlslightSheet);
            }

            xlslight.sheets = xlslightSheets.ToArray();

            return xlslight;
        }

        public static XSSFWorkbook ConvertXLSLightToXLSX(XLSLightWorkbook xlslight)
        {
            var workbook = new XSSFWorkbook();

            foreach (var originSheet in xlslight.sheets)
            {
                var sheet = workbook.CreateSheet(originSheet.name);
                IRow row = null;
                int rowIter = 0, columnIter = 0;

                foreach (var originCell in originSheet.cells)
                {
                    var offset = originCell.GetOffset();
                    var value = originCell.GetValue();
                    var type = originCell.GetCellType();

                    int offsetX = offset.Key;
                    int offsetY = offset.Value;

                    if(offsetY > 0)
                    {
                        columnIter = offsetX;
                        rowIter += offsetY;
                        row = sheet.CreateRow(rowIter);
                    }
                    else if(offsetX <= 1)
                    {
                        columnIter += 1;
                    }
                    else
                    {
                        columnIter += offsetX;
                    }

                    ICell cell = row.CreateCell(columnIter);
                    CellType cellType = (CellType)type;
                    cell.SetCellType(cellType);
                    switch (cellType)
                    {
                        case CellType.Formula:
                            cell.SetCellFormula(value);
                            break;
                        case CellType.Error:
                            byte err = 0;
                            byte.TryParse(value, out err);
                            cell.SetCellErrorValue(err);
                            break;
                        case CellType.Numeric:
                            double numb = 0;
                            double.TryParse(value, out numb);
                            cell.SetCellValue(numb);
                            break;
                        case CellType.Boolean:
                            if(value == "TRUE")
                            {
                                cell.SetCellValue(true);
                            }
                            else if(value == "FALSE")
                            {
                                cell.SetCellValue(false);
                            }
                            break;
                        case CellType.String:
                            cell.SetCellValue(value);
                            break;
                    }
                }
            }

            return workbook;
        }
    }
}
