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

    public class XLSLightConverter
    {
        public static YamlWorkbook ConvertXLSXToYaml(XSSFWorkbook xlsx)
        {
            YamlWorkbook yaml = new YamlWorkbook();
            int xOffset = -1, yOffset = -1;
            List<YamlSheet> yamlSheets = new List<YamlSheet>();

            for (int sheetCount = 0; sheetCount < xlsx.NumberOfSheets; sheetCount++)
            {
                YamlSheet yamlSheet = new YamlSheet();
                ISheet xlsxSheet = xlsx.GetSheetAt(sheetCount);
                yamlSheet.name = xlsxSheet.SheetName;
                List<YamlCell> yamlCells = new List<YamlCell>();

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

                        YamlCell yamlCell = new YamlCell();
                        yamlCell.SetOffset(xOffset, yOffset);
                        xOffset = 0;
                        yOffset = 0;
                        yamlCell.type = (int)xlsxCell.CellType;
                        switch (xlsxCell.CellType)
                        {
                            case CellType.Formula:
                                yamlCell.value = xlsxCell.CellFormula;
                                break;
                            case CellType.Error:
                                yamlCell.value = xlsxCell.ErrorCellValue.ToString();
                                break;
                            case CellType.Numeric:
                                yamlCell.value = xlsxCell.NumericCellValue.ToString();
                                break;
                            case CellType.Boolean:
                                yamlCell.value = xlsxCell.BooleanCellValue.ToString();
                                break;
                            case CellType.String:
                                yamlCell.value = xlsxCell.StringCellValue;
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

        public static XSSFWorkbook ConvertXLSXToYaml(YamlWorkbook yaml)
        {
            XSSFWorkbook workbook = new XSSFWorkbook();

            

            return workbook;
        }
    }
}
