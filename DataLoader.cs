using System.Net;
using System.IO;
using Newtonsoft.Json;

namespace Cryptostar
{

    public class DataLoader
    {
        public Ticker[] LoadTickers()
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://api.coinmarketcap.com/v1/ticker/");
            request.Method = "GET";
            request.ContentType = "application/json";

            using (var response = request.GetResponse())
            using (var stream = response.GetResponseStream())
            using (StreamReader responseReader = new StreamReader(stream))
            {
                string data = responseReader.ReadToEnd();

                using (var sr = new StringReader(data))
                using (var jsonReader = new JsonTextReader(sr))
                {
                    var items = JsonSerializer.CreateDefault()
                            .Deserialize<Ticker[]>(jsonReader);

                    return items;
                }
            }
        }
    }
}
