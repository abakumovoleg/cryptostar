using ExcelDna.Integration; 
using System.Linq;

namespace Cryptostar
{
    public class DataRender
    {
        public void RenderData(Ticker[] tickers)
        {
            dynamic xlApp =  ExcelDnaUtil.Application;
            
            var ws = xlApp.ActiveSheet;

            if (ws == null)
                return;
            
            ws.Cells.Clear();

            var props = typeof(Ticker).GetProperties();

            for (var j = 0; j < props.Length; j++)
            {
                var prop = props[j];
                var cell = ws.Cells[1, j + 1];
                cell.Value2 = prop.Name;
                cell.Font.Bold = true;
            }

            object[,] data = new object[tickers.Length, props.Length];

            var attributes = props.Select(x => x.GetFieldAttribute()).ToArray();

            for (var i = 0; i < tickers.Length; i++)
            {
                for (var j = 0; j < props.Length; j++)
                {
                    var val = props[j].GetValue(tickers[i], null);
                    var attr = attributes[j];
                    if (attr != null && attr.Type == FieldType.UnixDate)
                    {
                        data[i, j] = DateTimeExtensions.FromUnixTime((long)val).ToOADate();
                    }
                    else
                        data[i, j] = val;
                }
            }

            var startCell = ws.Cells[2, 1];
            var endCell = ws.Cells[1 + tickers.Length, props.Length];

            var range = ws.Range[startCell, endCell];
            range.Value2 = data; 

            var firstCell = ws.Cells[1, 1];
            ws.Range[firstCell, endCell].Columns.AutoFit();

            for (var j = 0; j < props.Length; j++)
            {
                var prop = props[j];
                var cell = ws.Cells[1, j + 1];

                var attr = prop.GetFieldAttribute();
                if (attr != null)
                {
                    cell.EntireColumn.NumberFormat
                        = UnifyFormatString(attr.Format, xlApp);
                }
            }
        }


        private string UnifyFormatString(string format, dynamic xlApp)
        {
            var yearCode = xlApp.International[Constants.xlYearCode];
            var monthCode = xlApp.International[Constants.xlMonthCode];
            var dayCode = xlApp.International[Constants.xlDayCode];

            var hourCode = xlApp.International[Constants.xlHourCode];
            var minuteCode = xlApp.International[Constants.xlMinuteCode];
            var secondCode = xlApp.International[Constants.xlSecondCode];
             
            return format.Replace("M", monthCode)
                            .Replace("y", yearCode)
                            .Replace("d", dayCode)
                            .Replace("s", secondCode)
                            .Replace("m", minuteCode)
                            .Replace("H", hourCode);
        }
    }

    public class Constants
    {
        public const int xlYearCode = 19;
        public const int xlMonthCode = 20;
        public const int xlDayCode = 21;
        public const int xlMinuteCode = 23;
        public const int xlSecondCode = 24;
        public const int xlHourCode = 22;
    }
}
