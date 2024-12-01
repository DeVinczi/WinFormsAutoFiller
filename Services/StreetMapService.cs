using System.Text.Json;

namespace WinFormsAutoFiller.Services;

public interface IStreetMapService
{
    Task<StreetMapApiModel> GetCityDetailsByName(string city, string postalcode);
}

internal class StreetMapService : IStreetMapService
{
    private static readonly HttpClient _httpClient = new();

    public async Task<StreetMapApiModel> GetCityDetailsByName(string city, string postalcode)
    {
        try
        {
            _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("YourAppName/1.0 (dawid8sly@gmail.com)");
            var result = await _httpClient.GetAsync($"https://nominatim.openstreetmap.org/search?city={city}&postalcode={postalcode}&format=jsonv2");

            var jsonStream = await result.Content.ReadAsStreamAsync();
            var jsonDocument = JsonDocument.Parse(jsonStream);

            var firstResult = jsonDocument.RootElement[0]; // Get the first element of the array
            var s = firstResult.GetProperty("display_name").GetString();
            var cityType = firstResult.GetProperty("addresstype").GetString();

            var isVillage = IsVillage(cityType);

            var info = s.Split(',');

            var model = new StreetMapApiModel();

            if (info.Length <= 3)
            {
                model.Miasto = info[0];
                model.Województwo = info[1];
                model.CzyWieś = false;
                return model;
            }

            if (cityType == "hamlet")
            {
                model.Miasto = info[0];
                model.Gmina = info[2];
                model.Powiat = info[3];
                model.Województwo = info[4];
                model.CzyWieś = true;
                return model;
            }

            model.Miasto = info[0];
            model.Gmina = info[1];
            model.Powiat = info[2];
            model.Województwo = info[3];
            model.CzyWieś = isVillage;

            return model;
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Uwaga", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return null;
        }
    }

    public static bool IsVillage(string cityType)
    {
        const string Wieś = "village";

        var isVillage = cityType switch
        {
            Wieś => true,
            _ => false
        };
        return isVillage;
    }
}

public class StreetMapApiModel
{
    public string Miasto { get; set; }
    public string Gmina { get; set; }
    public string Powiat { get; set; }
    public string Województwo { get; set; }
    public bool CzyWieś { get; set; }
}
