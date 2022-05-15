using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace xlslight
{
    public enum YamlProperty
    {
        Offset,
        Type,
        Value,
    }

    public class YamlWorkbook
    {
        public YamlSheet[] sheets { get; set; }
    }

    public struct YamlSheet
    {
        public string name { get; set; }
        public YamlCell[] cells { get; set; }
    }

    public struct YamlCell
    {
        public int type { get; set; }
        public string value { get; set; }
        public Dictionary<YamlProperty, string> property { get; set; }

        public void SetOffset(int xOffset, int yOffset )
        {
            string offset = string.Empty;

            if(yOffset > 0)
            {
                offset += xOffset.ToString();
                offset += ",";
                offset += yOffset.ToString();
            }
            else
            {
                if(xOffset > 1)
                {
                    offset += xOffset.ToString();
                }
            }

            if (offset != string.Empty)
            {
                if(property == null)
                {
                    property = new Dictionary<YamlProperty, string>();
                }
                property.Add(YamlProperty.Offset, offset);
            }
        }
    }

    static class YamlFile
    {
        public static async Task WriteAsync(string path, YamlWorkbook workbook)
        {
            var serializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();

            var yaml = serializer.Serialize(workbook);
            await File.WriteAllTextAsync(path, yaml);
        }

        public static YamlWorkbook Load(string path)
        {
            YamlWorkbook result = new YamlWorkbook();

            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(UnderscoredNamingConvention.Instance)
                .Build();

            string text = File.ReadAllText(path);
            result = deserializer.Deserialize<YamlWorkbook>(text);

            return result;
        }
    }
}
