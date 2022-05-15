﻿using System;
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
        CellType,
        Value,
    }

    public class XLSLightWorkbook
    {
        public XLSLightSheet[] sheets { get; set; }
    }

    public struct XLSLightSheet
    {
        public string name { get; set; }
        public XLSLightCell[] cells { get; set; }
    }

    public struct XLSLightCell
    {
        public Dictionary<XLSLightProperty, string> property { get; set; }

        public string GetValue()
        {
            return GetProperty(XLSLightProperty.Value);
        }

        public int GetCellType()
        {
            int typeInt = 0;
            int.TryParse(GetProperty(XLSLightProperty.CellType), out typeInt);
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
            SetProperty(XLSLightProperty.CellType, type.ToString());
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
            if (property == null)
            {
                property = new Dictionary<XLSLightProperty, string>();
            }
            property.Add(propertyType, propertyValue);
        }

        private string GetProperty(XLSLightProperty propertyType)
        {
            if (property == null)
            {
                return string.Empty;
            }

            string propertyValue = string.Empty;
            property.TryGetValue(propertyType, out propertyValue);

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
