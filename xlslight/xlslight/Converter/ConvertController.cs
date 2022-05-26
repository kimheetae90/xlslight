using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xlslight.Converter;

namespace xlslight
{
    public static class ConvertController
    {
        public static ConverterContainer converterContainer = new ConverterContainer();

        public static XLSLightWorkbook ConvertXLSXToXLSLight(XSSFWorkbook xlsx)
        {
            var xlslight = new XLSLightWorkbook();
            var xlslightSheets = new List<XLSLightSheet>();

            for (int sheetCount = 0; sheetCount < xlsx.NumberOfSheets; sheetCount++)
            {
                var xlslightSheet = new XLSLightSheet(xlslight);
                var xlsxSheet = xlsx.GetSheetAt(sheetCount);                
                var xlslightCells = new List<XLSLightCell>();
                int xOffset = 0, yOffset = 0;

                for (int rowCount = 0; rowCount <= xlsxSheet.LastRowNum; rowCount++)
                {
                    var xlsxRow = xlsxSheet.GetRow(rowCount);

                    if (xlsxRow == null)
                        continue;

                    for (int columnCount = 0; columnCount < xlsxRow.LastCellNum; columnCount++)
                    {
                        var xlsxCell = xlsxRow.GetCell(columnCount);
                        if (xlsxCell == null || xlsxCell.CellType == CellType.Blank)
                        {
                            xOffset++;
                            continue;
                        }

                        var xlslightCell = new XLSLightCell(xlslightSheet);
                        converterContainer.ConvertCell_XToL(xlsxCell, xlslightCell);
                        xlslightCell.SetOffset(xOffset, yOffset);

                        xOffset = 1;
                        yOffset = 0;
                        xlslightCells.Add(xlslightCell);
                    }

                    xOffset = 0;
                    yOffset++;
                }

                converterContainer.ConvertSheet_XToL(xlsxSheet, xlslightSheet);
                xlslightSheet.cells = xlslightCells.ToArray();
                xlslightSheets.Add(xlslightSheet);
            }

            converterContainer.ConvertWorkBook_XToL(xlsx, xlslight);
            xlslight.sheets = xlslightSheets.ToArray();

            return xlslight;
        }

        public static XSSFWorkbook ConvertXLSLightToXLSX(XLSLightWorkbook xlslight)
        {
            var workbook = new XSSFWorkbook();

            foreach(var xlsLightSheet in xlslight.sheets)
            {
                var xlsxSheet = workbook.CreateSheet();
                IRow row = null;
                int rowIter = 0, columnIter = -1;

                foreach (var xlsxLightCell in xlsLightSheet.cells)
                {
                    var offset = xlsxLightCell.GetOffset();
                    int offsetX = offset.Key;
                    int offsetY = offset.Value;

                    if (offsetY > 0)
                    {
                        columnIter = offsetX;
                        rowIter += offsetY;
                        row = xlsxSheet.CreateRow(rowIter);
                    }
                    else if (offsetX <= 1)
                    {
                        columnIter += 1;
                    }
                    else
                    {
                        columnIter += offsetX;
                    }

                    if (row == null)
                    {
                        row = xlsxSheet.CreateRow(rowIter);
                    }

                    ICell xlsxCell = row.CreateCell(columnIter);
                    converterContainer.ConvertCell_LToX(xlsxLightCell, xlsxCell);
                }
                converterContainer.ConvertSheet_LToX(xlsLightSheet, xlsxSheet);
            }
            converterContainer.ConvertWorkBook_LToX(xlslight, workbook);

            return workbook;
        }
    }
}
