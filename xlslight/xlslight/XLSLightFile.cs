using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace xlslight
{
    public enum XLSLightProperty
    {
        Offset,
        Type,
        Value,
        CellStyleIndex,
    }

    public class XLSLightCellStyle
    {
        public bool IsHidden { get; set; }
        public bool IsLocked { get; set; }
        public int HorizontalAlignment { get; set; }
        public int VerticalAlignment { get; set; }
        public int FillPattern { get; set; }
        public short FillBackgroundColor { get; set; }
        public short FillForegroundColor { get; set; }
        public short FontIndex { get; set; }
    }

    public class XLSLightFont
    {
        public string FontName { get; set; }
        public double FontHeightInPoints { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsStrickout { get; set; }
        public short FontColor { get; set; }
    }

    public class XLSLightWorkbook
    {
        public List<XLSLightSheet> sheets { get; set; }
        public List<XLSLightCellStyle> cellStyles { get; set; }
        public List<XLSLightFont> fonts { get; set; }

        public XLSLightSheet GetSheet(int index)
        {
            if(sheets.Count < index && index >= 0)
            {
                return sheets[index];
            }

            return null;
        }

        public XLSLightCellStyle GetCellStyle(int index)
        {
            if (cellStyles.Count < index && index >= 0)
            {
                return cellStyles[index];
            }

            return null;
        }

        public XLSLightFont GetFont(int index)
        {
            if (fonts.Count < index && index >= 0)
            {
                return fonts[index];
            }

            return null;
        }
    }

    public class XLSLightSheet
    {
        [YamlIgnore]
        public XLSLightWorkbook Workbook { get; set; }

        public string name { get; set; }
        public Dictionary<int, int> columnWidth { get; set; }
        public Dictionary<int, short> rowHeight { get; set; }
        public XLSLightCell[] cells { get; set; }

        public int GetColumnWidth(int column)
        {
            int width = 1;
            if (columnWidth != null)
            {
                columnWidth.TryGetValue(column, out width);
            }

            return width;
        }

        public short GetRowHeight(int row)
        {
            short height = 1;
            if (rowHeight != null)
            {
                rowHeight.TryGetValue(row, out height);
            }

            return height;
        }

        public void SetColumnWidth(int column, int width)
        {
            if (columnWidth == null)
            {
                columnWidth = new Dictionary<int, int>();
            }

            if(columnWidth.ContainsKey(column))
            {
                columnWidth[column] = width;
            }
            else
            {
                columnWidth.Add(column, width);
            }
        }

        public void SetRowHeight(int row, short height)
        {
            if (rowHeight == null)
            {
                rowHeight = new Dictionary<int, short>();
            }

            if (rowHeight.ContainsKey(row))
            {
                rowHeight[row] = height;
            }
            else
            {
                rowHeight.Add(row, height);
            }
        }
    }

    public class XLSLightCell : Dictionary<XLSLightProperty, string>
    {
        [YamlIgnore]
        public XLSLightSheet Sheet { get; set; }
        
        public string GetValue()
        {
            return GetProperty(XLSLightProperty.Value);
        }

        public int GetCellType()
        {
            int typeInt = 0;
            int.TryParse(GetProperty(XLSLightProperty.Type), out typeInt);
            return typeInt;
        }

        public short GetCellStyleIndex()
        {
            short cellStyleIndex = 0;
            short.TryParse(GetProperty(XLSLightProperty.CellStyleIndex), out cellStyleIndex);
            return cellStyleIndex;
        }

        public Offset GetOffset()
        {
            var offsetString = GetProperty(XLSLightProperty.Offset);
            Offset parsedOffset = Offset.Parse(offsetString);
            return parsedOffset;
        }

        public void SetValue(string value)
        {
            SetProperty(XLSLightProperty.Value, value);
        }

        public void SetCellType(int type)
        {
            SetProperty(XLSLightProperty.Type, type.ToString());
        }
        
        public void SetCellStyleIndex(short cellStyleIndex)
        {
            SetProperty(XLSLightProperty.CellStyleIndex, cellStyleIndex.ToString());
        }

        public void SetOffset(Offset offset)
        {
            string offsetString = offset.ToString();

            if (offsetString != string.Empty)
            {
                SetProperty(XLSLightProperty.Offset, offsetString);
            }
        }

        private void SetProperty(XLSLightProperty propertyType, string propertyValue)
        {
            Add(propertyType, propertyValue);
        }

        private string GetProperty(XLSLightProperty propertyType)
        {
            string propertyValue = string.Empty;
            TryGetValue(propertyType, out propertyValue);

            return propertyValue;
        }
    }

    static class XLSLightFile
    {
        public static void Write(string path, XLSLightWorkbook workbook)
        {
            var serializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();

            if (serializer != null)
            {
                var xlslight = serializer.Serialize(workbook);
                File.WriteAllText(path, xlslight);
            }
        }

        public static XLSLightWorkbook Load(string path)
        {
            XLSLightWorkbook workbook = new XLSLightWorkbook();

            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();

            if (deserializer != null)
            {
                string text = File.ReadAllText(path);
                workbook = deserializer.Deserialize<XLSLightWorkbook>(text);
                return workbook;
            }
            else
            {
                return null;
            }
        }
    }
}
