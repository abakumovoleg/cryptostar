using Newtonsoft.Json;
using System;
using System.Reflection;

namespace Cryptostar
{
    public static class AttributeExtensions
    {
        public static FieldFormatAttribute GetFieldAttribute(this PropertyInfo prop)             
        {
            var attr = prop.GetCustomAttribute<FieldFormatAttribute>();

            return attr;          
        }
    }

    public enum FieldType
    {
        Text,
        Number,
        UnixDate
    }

    public class FieldFormatAttribute : Attribute
    {
        public string Format { get; set; }
        public FieldType Type { get; set; }
    }

 

    public class Ticker
    {
        public string id { get; set; }
        public string name { get; set; }
        public string symbol { get; set; }        
        public decimal? rank { get; set; }
        [FieldFormat(Format = @"_-[$$-409]* #,##0.00_ ;_-[$$-409]* -#,##0.00 ;_-[$$-409]* "" - ""??_ ;_-@_ ", Type = FieldType.Number)]
        public string price_usd { get; set; }
        [FieldFormat(Format = "0.000000", Type = FieldType.Number)]
        public decimal? price_btc { get; set; }
        [JsonProperty(PropertyName = "24h_volume_usd")]       
        [FieldFormat(Format = @"_-[$$-409]* #,##0_ ;_-[$$-409]* -#,##0 ;_-[$$-409]* "" - ""_ ;_-@_ ", Type = FieldType.Number)]
        public string p24h_volume_usd { get; set; }      
        [FieldFormat(Format = @"_-[$$-409]* #,##0_ ;_-[$$-409]* -#,##0 ;_-[$$-409]* ""-""_ ;_-@_ ", Type = FieldType.Number)]
        public string market_cap_usd { get; set; }
        [FieldFormat(Format = @"_ * #,##0_ ;_ * -#,##0_ ;_ * "" - ""??_ ;_ @_ ", Type = FieldType.Number)]
        public decimal? available_supply { get; set; }
        [FieldFormat(Format = @"_ * #,##0_ ;_ * -#,##0_ ;_ * "" - ""??_ ;_ @_ ", Type = FieldType.Number)]
        public decimal? total_supply { get; set; }
        [FieldFormat(Format = "#%", Type = FieldType.Number)]
        public string percent_change_1h { get; set; }
        [FieldFormat(Format = "#%", Type = FieldType.Number)]
        public string percent_change_24h { get; set; }
        [FieldFormat(Format = "#%", Type = FieldType.Number)]
        public string percent_change_7d { get; set; }
        [FieldFormat(Format = "dd-mm-yyyy", Type = FieldType.UnixDate)]
        public long last_updated { get; set; }
    }

}
