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
        public XLSLightCell[] cells { get; set; }
    }

    public struct XLSLightCell
    {
        public Dictionary<XLSLightProperty, string> property { get; set; }

        public void SetValue(string value)
        {
            SetProperty(XLSLightProperty.Value, value);
        }

        public void SetType(int type)
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
            if (property == null)
            {
                property = new Dictionary<XLSLightProperty, string>();
            }
            property.Add(propertyType, propertyValue);
        }
    }

    static class XLSLightFile
    {
        public static async Task WriteAsync(string path, XLSLightWorkbook workbook)
        {
            var serializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();

            var yaml = serializer.Serialize(workbook);
            await File.WriteAllTextAsync(path, yaml);
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
