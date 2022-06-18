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

            var prevOffset = new Offset(-1, 0);
            var currOffset = new Offset(0, 0);

            converterContainer.PreConvertXToL(xlsx, xlslight);

            for (int sheetCount = 0; sheetCount < xlsx.NumberOfSheets; sheetCount++)
            {
                var xlslightSheet = new XLSLightSheet();
                var xlsxSheet = xlsx.GetSheetAt(sheetCount);                
                var xlslightCells = new List<XLSLightCell>();

                xlslightSheet.Workbook = xlslight;

                converterContainer.PreConvertXToL(xlsxSheet, xlslightSheet);
                
                for (int rowCount = 0; rowCount <= xlsxSheet.LastRowNum; rowCount++)
                {
                    var xlsxRow = xlsxSheet.GetRow(rowCount);

                    if (xlsxRow != null)
                    {
                        for (int columnCount = 0; columnCount < xlsxRow.LastCellNum; columnCount++)
                        {
                            var xlsxCell = xlsxRow.GetCell(columnCount);
                            if (xlsxCell == null || xlsxCell.CellType == CellType.Blank)
                            {
                                currOffset.x++;
                                continue;
                            }

                            var xlslightCell = new XLSLightCell();
                            xlslightCell.Sheet = xlslightSheet;
                            converterContainer.ConvertXToL(xlsxCell, xlslightCell);
                            xlslightCell.SetOffset(currOffset - prevOffset);

                            prevOffset = currOffset;
                            currOffset = new Offset(0,0);
                            xlslightCells.Add(xlslightCell);
                        }
                    }

                    prevOffset.x = 0;
                    currOffset.x = 0;
                    currOffset.y++;
                }

                converterContainer.ConvertXToL(xlsxSheet, xlslightSheet);
                xlslightSheet.cells = xlslightCells.ToArray();
                xlslightSheets.Add(xlslightSheet);
            }

            converterContainer.ConvertXToL(xlsx, xlslight);
            xlslight.sheets = xlslightSheets;

            return xlslight;
        }

        public static XSSFWorkbook ConvertXLSLightToXLSX(XLSLightWorkbook xlslight)
        {
            var workbook = new XSSFWorkbook();            

            converterContainer.PreConvertLToX(xlslight, workbook);
            foreach (var xlsLightSheet in xlslight.sheets)
            {
                var xlsxSheet = workbook.CreateSheet();
                IRow row = null;
                int rowIter = 0, columnIter = -1;

                converterContainer.PreConvertLToX(xlsLightSheet, xlsxSheet);
                foreach (var xlsxLightCell in xlsLightSheet.cells)
                {
                    var offset = xlsxLightCell.GetOffset();

                    if (offset.y > 0)
                    {
                        columnIter = offset.x;
                        rowIter += offset.y;
                        row = xlsxSheet.CreateRow(rowIter);
                    }
                    else if (offset.x <= 1)
                    {
                        columnIter += 1;
                    }
                    else
                    {
                        columnIter += offset.x;
                    }

                    if (row == null)
                    {
                        row = xlsxSheet.CreateRow(rowIter);
                    }

                    ICell xlsxCell = row.CreateCell(columnIter);
                    converterContainer.ConvertLToX(xlsxLightCell, xlsxCell);
                }
                converterContainer.ConvertLToX(xlsLightSheet, xlsxSheet);
            }
            converterContainer.ConvertLToX(xlslight, workbook);

            return workbook;
        }
    }
}
