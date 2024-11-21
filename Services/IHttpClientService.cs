using System.Text;

namespace WinFormsAutoFiller.Services
{
    public interface IHttpClientService
    {
        Task<Stream> GetAsync(string url);
        Task<string> PostKrsAsync(string url, string key, string krs);
    }

    public class HttpClientService : IHttpClientService
    {
        private static readonly HttpClient HttpClient = new HttpClient();

        public async Task<Stream> GetAsync(string url)
        {
            try
            {
                var result = await HttpClient.GetAsync(url);
                return await result.Content.ReadAsStreamAsync();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return Stream.Null;
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
            var path = Path.Combine(Path.GetTempPath(), $"Odpis_Pełny_KRS_{krs}");
            using var file = File.Create(path);
            var stream = await result.Content.ReadAsStreamAsync();
            await stream.CopyToAsync(file);

            return path;
        }
    }
}
