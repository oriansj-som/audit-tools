using Newtonsoft.Json;
using OfficeOpenXml.Style;

namespace common
{
    internal class ugly
    {
        internal static UglyBlob unpack_pretty(string value)
        {
            Require.True(null != value, "impossible null passed to unpack_pretty");
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
#pragma warning disable CS8604 // Possible null reference argument.
            UglyBlob a = JsonConvert.DeserializeObject<UglyBlob>(value);
#pragma warning disable CS8603 // Possible null reference return.
            return a;
#pragma warning restore CS8603 // Possible null reference return.
#pragma warning restore CS8604 // Possible null reference argument.
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
        }
        public static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }
        public static string Base64Encode(string plainText)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return Convert.ToBase64String(plainTextBytes);
        }
        public static string make_ugly(ExcelRichTextCollection richText)
        {
            if (0 >= richText.Count) return richText.Text;
            if (Match.CaseInsensitive("", richText.Text))
            {
                return string.Empty;
            }

            string r = "{\"richtext\": [";
            int i = 0;
            List<string> set = new List<string>();
            while (i < richText.Count)
            {
                ExcelRichText h = richText[i];
                string s = "{\"bold\": " + h.Bold.ToString().ToLower() + ","
                          + "\"color\": \"" + h.Color.Name.ToString() + "\","
                          + "\"italic\":" + h.Italic.ToString().ToLower() + ","
                          + "\"preserve_space\":" + h.PreserveSpace.ToString().ToLower() + ","
                          + "\"strike\":" + h.Strike.ToString().ToLower() + ","
                          + "\"underline\":" + h.UnderLine.ToString().ToLower() + ","
                          + "\"base64\": \"" + Base64Encode(h.Text) + "\"}";
                set.Add(s);
                i += 1;
            }
            r += string.Join(",", set);
            r += "]}";
            return r;
        }
    }

    public class UglyBlob
    {
        [JsonProperty(PropertyName = "richtext")]
        public List<UglySegment> segments;
    }
    public class UglySegment
    {
        [JsonProperty(PropertyName = "bold")]
        public bool bold { get; set; }
        [JsonProperty(PropertyName = "color")]
        public string color { get; set; }
        [JsonProperty(PropertyName = "italic")]
        public bool italic { get; set; }
        [JsonProperty(PropertyName = "preserve_space")]
        public bool preserve_space { get; set; }
        [JsonProperty(PropertyName = "strike")]
        public bool strike { get; set; }
        [JsonProperty(PropertyName = "underline")]
        public bool underline { get; set; }
        [JsonProperty(PropertyName = "base64")]
        public string encoded_text { get; set; }
    }
 }