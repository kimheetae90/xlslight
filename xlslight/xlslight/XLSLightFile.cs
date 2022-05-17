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
    }

    public struct XLSLightSheet
    {
        public string name { get; set; }
        public Dictionary<int, int> ColumnWidth { get; set; }
        public Dictionary<int, short> RowHeight { get; set; }
        public XLSLightCell[] cells { get; set; }

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

        public KeyValuePair<int, int> GetOffset()
        {
            var offsetString = GetProperty(XLSLightProperty.Offset);
            char[] delimiterChars = { ',' };
            int offsetX = 0, offsetY = 0;
            if (offsetString != null && offsetString.Length != 0)
            {
                string[] offsetSplitedString = offsetString.Split(delimiterChars);
                if (offsetSplitedString.Length > 0)
                {
                    int.TryParse(offsetSplitedString[0], out offsetX);
                }

                if (offsetSplitedString.Length > 1)
                {
                    int.TryParse(offsetSplitedString[1], out offsetY);
                }
            }

            return new KeyValuePair<int, int>(offsetX, offsetY);
        }

        public void SetValue(string value)
        {
            SetProperty(XLSLightProperty.Value, value);
        }

        public void SetCellType(int type)
        {
            SetProperty(XLSLightProperty.Type, type.ToString());
        }

        public void SetOffset(int xOffset, int yOffset)
        {
            string offset = string.Empty;

            if (yOffset > 0)
            {
                offset += xOffset.ToString();
                offset += ",";
                offset += yOffset.ToString();
            }
            else
            {
                if (xOffset > 1)
                {
                    offset += xOffset.ToString();
                }
            }

            if (offset != string.Empty)
            {
                SetProperty(XLSLightProperty.Offset, offset);
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
