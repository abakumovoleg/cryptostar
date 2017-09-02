using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cryptostar
{
    public class DataRender
    {
        public void RenderData(Ticker[] tickers)
        {
            Application xlApp =  (Application)ExcelDnaUtil.Application;
            
            Worksheet ws = xlApp.ActiveSheet as Worksheet;

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


        private string UnifyFormatString(string format, Application xlApp)
        {
            var yearCode = xlApp.International[XlApplicationInternational.xlYearCode];
            var monthCode = xlApp.International[XlApplicationInternational.xlMonthCode];
            var dayCode = xlApp.International[XlApplicationInternational.xlDayCode];

            var hourCode = xlApp.International[XlApplicationInternational.xlHourCode];
            var minuteCode = xlApp.International[XlApplicationInternational.xlMinuteCode];
            var secondCode = xlApp.International[XlApplicationInternational.xlSecondCode];
             
            return format.Replace("M", monthCode)
                            .Replace("y", yearCode)
                            .Replace("d", dayCode)
                            .Replace("s", secondCode)
                            .Replace("m", minuteCode)
                            .Replace("H", hourCode);
        }
    }
}
