using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace FileViewer.Services
{
    public class QueryService
    {
        private static readonly HttpClient client = new HttpClient();

        public static async Task<T> GetRequestAsync<T>(string url, Dictionary<string, string> queryParams, Dictionary<string, string> headers = null)
        {
            headers = headers ?? new Dictionary<string, string>() { };
            queryParams = queryParams ?? new Dictionary<string, string>() { };
            string urlWithParams = url;
            if (queryParams.Count > 0)
            {
                string paramsStr = String.Join('&', queryParams.Select(s => $"{s.Key}={s.Value}"));
                urlWithParams += $"?{paramsStr}";
            }
            
            using (var requestMessage = new HttpRequestMessage(HttpMethod.Get, urlWithParams))
            {
                

                foreach(var header in headers)
                {
                    requestMessage.Headers.Add(header.Key, header.Value);
                }

                using HttpResponseMessage response = await client.SendAsync(requestMessage).ConfigureAwait(false);
                var responseContent = await response.Content.ReadAsStringAsync();
                return JsonConvert.DeserializeObject<T>(responseContent);
            }
        }

        public static async Task<T> PostRequestAsync<T>(string url, Dictionary<string,string> queryParams)
        {
            var json = MyDictionaryToJson(queryParams);
            using HttpContent content = new StringContent(json, Encoding.UTF8, "application/json");
            using HttpResponseMessage response = await client.PostAsync(url, content).ConfigureAwait(false);
            var responseContent = await response.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<T>(responseContent);
        }


        private static string MyDictionaryToJson(Dictionary<string, string> dict)
        {
            var entries = dict.Select(d =>
                string.Format("\"{0}\": {1}", d.Key, d.Value));
            return "{" + string.Join(",", entries) + "}";
        }
    }
}
