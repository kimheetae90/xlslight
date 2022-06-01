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
    }

    public class XLSLightWorkbook
    {
        public XLSLightSheet[] sheets { get; set; }

        public XLSLightSheet GetSheet(int index)
        {
            if(sheets.Length < index && index >= 0)
            {
                return sheets[index];
            }

            return null;
        }
    }

    public class XLSLightSheet
    {
        [YamlIgnore]
        public XLSLightWorkbook Workbook { get; private set; }

        public string name { get; set; }
        public Dictionary<int, int> ColumnWidth { get; set; }
        public Dictionary<int, short> RowHeight { get; set; }
        public XLSLightCell[] cells { get; set; }

        public XLSLightSheet(XLSLightWorkbook parent)
        {
            Workbook = parent;
        }

        public int GetColumnWidth(int column)
        {
            int width = 1;
            if (ColumnWidth != null)
            {
                ColumnWidth.TryGetValue(column, out width);
            }

            return width;
        }

        public short GetRowHeight(int row)
        {
            short height = 1;
            if (RowHeight != null)
            {
                RowHeight.TryGetValue(row, out height);
            }

            return height;
        }

        public void SetColumnWidth(int column, int width)
        {
            if (ColumnWidth == null)
            {
                ColumnWidth = new Dictionary<int, int>();
            }

            if(ColumnWidth.ContainsKey(column))
            {
                ColumnWidth[column] = width;
            }
            else
            {
                ColumnWidth.Add(column, width);
            }
        }

        public void SetRowHeight(int row, short height)
        {
            if (RowHeight == null)
            {
                RowHeight = new Dictionary<int, short>();
            }

            if (RowHeight.ContainsKey(row))
            {
                RowHeight[row] = height;
            }
            else
            {
                RowHeight.Add(row, height);
            }
        }
    }

    public class XLSLightCell : Dictionary<XLSLightProperty, string>
    {
        [YamlIgnore]
        public XLSLightSheet Sheet { get; private set; }

        public XLSLightCell(XLSLightSheet parent)
        {
            Sheet = parent;
        }

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
        public static async Task WriteAsync(string path, XLSLightWorkbook workbook)
        {
            var serializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();

            var xlslight = serializer.Serialize(workbook);
            await File.WriteAllTextAsync(path, xlslight);
        }

        public static XLSLightWorkbook Load(string path)
        {
            XLSLightWorkbook result = new XLSLightWorkbook();

            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(UnderscoredNamingConvention.Instance)
                .Build();

            string text = File.ReadAllText(path);
            result = deserializer.Deserialize<XLSLightWorkbook>(text);

            return result;
        }
    }
}
