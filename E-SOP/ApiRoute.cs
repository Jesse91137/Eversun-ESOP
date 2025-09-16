using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace E_SOP
{
    public class WipAtt
    {
        public string wipNO { get; set; }
        public string itemNO { get; set; }
        public string wipProcess { get; set; }
        public string ecn { get; set; }
    }
    class ApiRoute
    {
        private static HttpClient NewClient()
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.Timeout = TimeSpan.FromHours(1);//1小時,可自訂義
            //client.DefaultRequestHeaders.Add("x-api-key", System.Configuration.ConfigurationManager.AppSettings["x-api-key"]);
            return client;
        }
        public static string GetMethod(string apiUrl)
        {
            var client = NewClient();
            try
            {
                apiUrl = "http://192.168.4.109:5088/" + apiUrl;
                var aa = client.GetAsync(apiUrl);
                HttpResponseMessage response = client.GetAsync(apiUrl).Result;

                if (response.IsSuccessStatusCode)
                {
                    var jsonString = response.Content.ReadAsStringAsync();
                    jsonString.Wait();
                    //List<TResult> data = JsonConvert.DeserializeObject<List<TResult>>(jsonString.Result);
                    return jsonString.Result;
                }
                return "error";
            }
            catch
            {
                //Console.WriteLine(ex.Message);
                return "無法連線WebAPI";
            }
        }
    }
}
