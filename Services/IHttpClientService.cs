using Newtonsoft.Json;

using System.Text;

using WinFormsAutoFiller.Models.KrsEntity;

namespace WinFormsAutoFiller.Services
{
    public interface IHttpClientService
    {
        Task<Root> GetAsync(string url);
        Task<string> PostKrsAsync(string url, string key, string krs);
    }

    public class HttpClientService : IHttpClientService
    {
        private static readonly HttpClient HttpClient = new HttpClient();

        public async Task<Root> GetAsync(string url)
        {
            try
            {
                var result = await HttpClient.GetAsync(url);

                if (result.IsSuccessStatusCode)
                {
                    var json = await result.Content.ReadAsStringAsync();
                    var odpis = JsonConvert.DeserializeObject<Root>(json);
                    return odpis;
                }
                else
                {
                    Console.WriteLine($"Request failed with status code: {result.StatusCode}");
                    return null; // or throw an exception if desired
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return null;
            }
        }

        public async Task<string> PostKrsAsync(string url, string key, string krs)
        {
            HttpClient.DefaultRequestHeaders.Add("X-api-key", "TopSecretApiKey");
            var jsonPayload = $@"
                {{
                    ""krs"": ""{key}"",
                    ""register"": ""P"",
                    ""format"": ""PDF""
                }}";
            var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
            var result = await HttpClient.PostAsync(url, content);
            var path = Path.Combine(Path.GetTempPath(), $"Odpis_Pełny_KRS_{krs}" + ".pdf");
            using var file = File.Create(path);
            var stream = await result.Content.ReadAsStreamAsync();
            await stream.CopyToAsync(file);

            return path;
        }
    }
}
