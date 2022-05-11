using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlslight
{
    public interface ICellInfoConvert
    {

    }

    public class XLSConverter
    {
        public static XLSLightWorkbook ConvertXLSXToYaml(XSSFWorkbook xlsx)
        {
            XLSLightWorkbook yaml = new XLSLightWorkbook();
            int xOffset = -1, yOffset = -1;
            List<XLSLightSheet> yamlSheets = new List<XLSLightSheet>();

            for (int sheetCount = 0; sheetCount < xlsx.NumberOfSheets; sheetCount++)
            {
                XLSLightSheet yamlSheet = new XLSLightSheet();
                ISheet xlsxSheet = xlsx.GetSheetAt(sheetCount);
                yamlSheet.name = xlsxSheet.SheetName;
                List<XLSLightCell> yamlCells = new List<XLSLightCell>();

                xOffset = yOffset = 0;

                for (int rowCount = 0; rowCount <= xlsxSheet.LastRowNum; rowCount++)
                {
                    IRow xlsxRow = xlsxSheet.GetRow(rowCount);

                    for (int columnCount = 0; columnCount < xlsxRow.LastCellNum; columnCount++)
                    {
                        ICell xlsxCell = xlsxRow.GetCell(columnCount);

                        if (xlsxCell == null || xlsxCell.CellType == CellType.Blank)
                        {
                            xOffset++;
                            continue;
                        }

                        XLSLightCell yamlCell = new XLSLightCell();
                        yamlCell.SetOffset(xOffset, yOffset);
                        xOffset = 0;
                        yOffset = 0;
                        yamlCell.SetType((int)xlsxCell.CellType);                        
                        switch (xlsxCell.CellType)
                        {
                            case CellType.Formula:
                                yamlCell.SetValue(xlsxCell.CellFormula);
                                break;
                            case CellType.Error:
                                yamlCell.SetValue(xlsxCell.ErrorCellValue.ToString());
                                break;
                            case CellType.Numeric:
                                yamlCell.SetValue(xlsxCell.NumericCellValue.ToString());
                                break;
                            case CellType.Boolean:
                                yamlCell.SetValue(xlsxCell.BooleanCellValue.ToString());
                                break;
                            case CellType.String:
                                yamlCell.SetValue(xlsxCell.StringCellValue);
                                break;
                        }

                        yamlCells.Add(yamlCell);
                        xOffset++;
                    }

                    xOffset = 0;
                    yOffset++;
                }
                yamlSheet.cells = yamlCells.ToArray();
                yamlSheets.Add(yamlSheet);
            }

            yaml.sheets = yamlSheets.ToArray();

            return yaml;
        }

        public static XSSFWorkbook ConvertXLSXToYaml(XLSLightWorkbook yaml)
        {
            XSSFWorkbook workbook = new XSSFWorkbook();

            

            return workbook;
        }
    }
}
