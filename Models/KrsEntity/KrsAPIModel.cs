using Newtonsoft.Json;

namespace WinFormsAutoFiller.Models.KrsEntity
{
    public class Root
    {
        [JsonProperty("odpis")]
        public Odpis Odpis { get; set; }
    }

    public class Odpis
    {
        [JsonProperty("rodzaj")]
        public string Rodzaj { get; set; }

        [JsonProperty("naglowekA")]
        public NaglowekA NaglowekA { get; set; }

        [JsonProperty("dane")]
        public Dane Dane { get; set; }
    }

    public class NaglowekA
    {
        [JsonProperty("rejestr")]
        public string Rejestr { get; set; }

        [JsonProperty("numerKRS")]
        public string NumerKRS { get; set; }

        [JsonProperty("dataCzasOdpisu")]
        public string DataCzasOdpisu { get; set; }

        [JsonProperty("stanZDnia")]
        public string StanZDnia { get; set; }

        [JsonProperty("dataRejestracjiWKRS")]
        public string DataRejestracjiWKRS { get; set; }

        [JsonProperty("numerOstatniegoWpisu")]
        public int NumerOstatniegoWpisu { get; set; }

        [JsonProperty("dataOstatniegoWpisu")]
        public string DataOstatniegoWpisu { get; set; }

        [JsonProperty("sygnaturaAktSprawyDotyczacejOstatniegoWpisu")]
        public string SygnaturaAktSprawyDotyczacejOstatniegoWpisu { get; set; }

        [JsonProperty("oznaczenieSaduDokonujacegoOstatniegoWpisu")]
        public string OznaczenieSaduDokonujacegoOstatniegoWpisu { get; set; }

        [JsonProperty("stanPozycji")]
        public int StanPozycji { get; set; }
    }

    public class Dane
    {
        [JsonProperty("dzial1")]
        public Dzial1 Dzial1 { get; set; }

        // You can add other sections like "dzial2", "dzial3" if needed
    }

    public class Dzial1
    {
        [JsonProperty("danePodmiotu")]
        public DanePodmiotu DanePodmiotu { get; set; }

        [JsonProperty("siedzibaIAdres")]
        public SiedzibaIAdres SiedzibaIAdres { get; set; }
    }

    public class DanePodmiotu
    {
        [JsonProperty("formaPrawna")]
        public string FormaPrawna { get; set; }

        [JsonProperty("identyfikatory")]
        public Identifikatory Identyfikatory { get; set; }

        [JsonProperty("nazwa")]
        public string Nazwa { get; set; }

        [JsonProperty("czyProwadziDzialalnoscZInnymiPodmiotami")]
        public bool CzyProwadziDzialalnoscZInnymiPodmiotami { get; set; }

        [JsonProperty("czyPosiadaStatusOPP")]
        public bool CzyPosiadaStatusOPP { get; set; }
    }

    public class Identifikatory
    {
        [JsonProperty("regon")]
        public string Regon { get; set; }

        [JsonProperty("nip")]
        public string Nip { get; set; }
    }

    public class SiedzibaIAdres
    {
        [JsonProperty("siedziba")]
        public Siedziba Siedziba { get; set; }

        [JsonProperty("adres")]
        public Adres Adres { get; set; }

        [JsonProperty("adresPocztyElektronicznej")]
        public string AdresPocztyElektronicznej { get; set; }

        [JsonProperty("adresStronyInternetowej")]
        public string AdresStronyInternetowej { get; set; }
    }

    public class Siedziba
    {
        [JsonProperty("kraj")]
        public string Kraj { get; set; }

        [JsonProperty("wojewodztwo")]
        public string Wojewodztwo { get; set; }

        [JsonProperty("powiat")]
        public string Powiat { get; set; }

        [JsonProperty("gmina")]
        public string Gmina { get; set; }

        [JsonProperty("miejscowosc")]
        public string Miejscowosc { get; set; }
    }

    public class Adres
    {
        [JsonProperty("ulica")]
        public string Ulica { get; set; }

        [JsonProperty("nrDomu")]
        public string NrDomu { get; set; }

        [JsonProperty("nrLokalu")]
        public string NrLokalu { get; set; }

        [JsonProperty("miejscowosc")]
        public string Miejscowosc { get; set; }

        [JsonProperty("kodPocztowy")]
        public string KodPocztowy { get; set; }

        [JsonProperty("poczta")]
        public string Poczta { get; set; }

        [JsonProperty("kraj")]
        public string Kraj { get; set; }
    }
}
