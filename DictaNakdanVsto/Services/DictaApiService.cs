using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Newtonsoft.Json;
using DictaNakdanVsto.Models;

namespace DictaNakdanVsto.Services
{
    public class DictaApiService
    {
        private static readonly HttpClient _httpClient = new HttpClient();
        private const string API_URL = "https://nakdan-u1-0.loadbalancer.dicta.org.il/api";

        public async Task<List<DictaToken>> GetNakdanAsync(string text, DictaRequest settings)
        {
            settings.Data = text;
            var json = JsonConvert.SerializeObject(settings);

            using (var request = new HttpRequestMessage(HttpMethod.Post, API_URL))
            {
                request.Content = new StringContent(json, Encoding.UTF8, "text/plain");

                var response = await _httpClient.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var responseJson = await response.Content.ReadAsStringAsync();

                // כאן התיקון: קוראים לתוך DictaResponse ושולפים את Data
                var dictaResponse = JsonConvert.DeserializeObject<DictaResponse>(responseJson);

                // אם Data ריק או NULL, נחזיר רשימה ריקה כדי לא לקרוס
                return dictaResponse?.Data ?? new List<DictaToken>();
            }
        }

        public List<string> SplitToSentences(string text)
        {
            string pattern = @"(?<=[\.!\?\n\r]+)(?<!\b[א-ת]{1,3}[""'][א-ת]?[\.!\?])\s+";
            var sentences = Regex.Split(text, pattern);

            var result = new List<string>();
            foreach (var sentence in sentences)
            {
                if (!string.IsNullOrWhiteSpace(sentence))
                {
                    result.Add(sentence.Trim());
                }
            }
            return result;
        }
    }
}