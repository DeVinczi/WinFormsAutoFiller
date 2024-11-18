using System.Text.RegularExpressions;

namespace FormFiller.Constants;

public static class WorkersFormExcel
{
    public static readonly string[] Headers = new string[]
    {
        WorkersFormKeys.Number,
        WorkersFormKeys.NazwiskoImie,
        WorkersFormKeys.Miasto,
        WorkersFormKeys.Pesel,
        WorkersFormKeys.Priorytet,
        WorkersFormKeys.Nieobecny,
        WorkersFormKeys.Plec,
        WorkersFormKeys.Wiek,
        WorkersFormKeys.DeficytWieku,
        WorkersFormKeys.GrupaWiekowa,
        WorkersFormKeys.Stanowisko,
        WorkersFormKeys.KodZawodu,
        WorkersFormKeys.GrupaDeficytowa,
        WorkersFormKeys.DeficytGdansk,
        WorkersFormKeys.DeficytPomorskie,
        WorkersFormKeys.PoziomWyksztalcenia,
        WorkersFormKeys.FormaZatrudnienia,
        WorkersFormKeys.Etat,
        WorkersFormKeys.OkresZatrudnieniaOd,
        WorkersFormKeys.OkresZatrudnieniaDo,
        WorkersFormKeys.KwotaOtrzymanegoDofinansowania
    };
}

public static class WorkersFormKeys
{
    public const string Number = "Numer";
    public const string NazwiskoImie = "Nazwisko i imię";
    public const string Miasto = "Miasto";
    public const string Pesel = "PESEL";
    public const string Priorytet = "Priorytet";
    public const string Nieobecny = "Nieobecny (L4/MACIERZYŃSTWO)";
    public const string Plec = "Płeć";
    public const string Wiek = "Wiek";
    public const string DeficytWieku = "DeficytWieku";
    public const string GrupaWiekowa = "Grupa Wiekowa";
    public const string Stanowisko = "Stanowisko";
    public const string KodZawodu = "Kod Zawodu";
    public const string GrupaDeficytowa = "Grupa Deficytowa";
    public const string DeficytGdansk = "Deficyt Gdańsk";
    public const string DeficytPomorskie = "Deficyt Pomorskie";
    public const string PoziomWyksztalcenia = "Poziom Wykształcenia";
    public const string FormaZatrudnienia = "Forma Zatrudnienia";
    public const string Etat = "Etat";
    public const string OkresZatrudnieniaOd = "Okres zatrudnienia z";
    public const string OkresZatrudnieniaDo = "Okres zatrudnienia do";
    public const string KwotaOtrzymanegoDofinansowania = "Kwota otrzymanego dofinansowania";
}

public static class WorkerFormPatterns
{
    public static readonly Dictionary<string, Regex> Patterns = new Dictionary<string, Regex>
    {
        { WorkersFormKeys.Number, new Regex(@"^.*$")},
        { WorkersFormKeys.NazwiskoImie, new Regex(@"^(?:[Nn]azwisko\s*i\s*[Ii]mi[ęe]|[Ii]mi[ęe]\s*i\s*[Nn]azwisko)$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.Miasto, new Regex(@"^[Mm]iasto$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.Pesel, new Regex(@"^PESEL$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.Priorytet, new Regex(@"^[Pp]riorytet", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.Nieobecny, new Regex(@"Nieobecny \(([\w\d/ĄŚŻąćęłńóśźż]+(?:\/[\w\d/ĄŚŻąćęłńóśźż]+)*)\)", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.Plec, new Regex(@"^P[łl]e[ćc]$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.Wiek, new Regex(@"^[Ww]iek$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.DeficytWieku, new Regex(@"^[Dd]eficyt\s*[Ww]iek(?:u)?$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.GrupaWiekowa, new Regex(@"^[Gg]rupa\s*[Ww]iekowa$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.Stanowisko, new Regex(@"^[Ss]tanowisko$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.KodZawodu, new Regex(@"^[Kk]od\s*[Zz]awodu$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.GrupaDeficytowa, new Regex(@"^[Gg]rupa\s*[Dd]eficytowa$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.DeficytGdansk, new Regex(@"^[Dd]eficyt\s*[Gg]da[ńn]sk$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.DeficytPomorskie, new Regex(@"^[Dd]eficyt\s*[Pp]omorskie$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.PoziomWyksztalcenia, new Regex(@"^POZIOM\s*WYKSZTA[ŁL]CENIA$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.FormaZatrudnienia, new Regex(@"^FORMA\s*ZATRUDNIENIA$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.Etat, new Regex(@"^ETAT$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.OkresZatrudnieniaOd, new Regex(@"^OKRES\s*ZATR\.?\s*OD$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.OkresZatrudnieniaDo, new Regex(@"^OKRES\s*ZATR\.?\s*DO$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { WorkersFormKeys.KwotaOtrzymanegoDofinansowania, new Regex(@"(?i)\s*kwota\s+otrzymanego\s+dofinansowania\s*[\.,]?", RegexOptions.Compiled|RegexOptions.IgnoreCase) }
    };

    public class Worker
    {
        public string Number { get; set; }
        public string NazwiskoImie { get; set; }
        public string Miasto { get; set; }
        public string Pesel { get; set; }
        public string Plec { get; set; }
        public string Priorytet { get; set; }
        public int Wiek { get; set; }
        public string DeficytWieku { get; set; }
        public string GrupaWiekowa { get; set; }
        public string Stanowisko { get; set; }
        public string KodZawodu { get; set; }
        public string GrupaDeficytowa { get; set; }
        public string DeficytGdansk { get; set; }
        public string DeficytPomorskie { get; set; }
        public string PoziomWyksztalcenia { get; set; }
        public string FormaZatrudnienia { get; set; }
        public string Etat { get; set; }
        public string OkresZatrudnieniaOd { get; set; }
        public string OkresZatrudnieniaDo { get; set; }
        
        public Worker(
            string number,
            string nazwiskoImie,
            string miasto,
            string pesel,
            string priorytet,
            string plec,
            int wiek,
            string deficytWieku,
            string grupaWiekowa,
            string stanowisko,
            string kodZawodu,
            string grupaDeficytowa,
            string deficytGdansk,
            string deficytPomorskie,
            string poziomWyksztalcenia,
            string formaZatrudnienia ,
            string etat,
            string okresZatrudnieniaOd,
            string okresZatrudnieniaDo)
        {
            Number = number;
            NazwiskoImie = nazwiskoImie;
            Miasto = miasto;
            Pesel = pesel;
            Priorytet = priorytet;
            Plec = plec;
            Wiek = wiek;
            DeficytWieku = deficytWieku;
            GrupaWiekowa = grupaWiekowa;
            Stanowisko = stanowisko;
            KodZawodu = kodZawodu;
            GrupaDeficytowa = grupaDeficytowa;
            DeficytGdansk = deficytGdansk;
            DeficytPomorskie = deficytPomorskie;
            PoziomWyksztalcenia = poziomWyksztalcenia;
            FormaZatrudnienia = formaZatrudnienia;
            Etat = etat;
            OkresZatrudnieniaOd = okresZatrudnieniaOd;
            OkresZatrudnieniaDo = okresZatrudnieniaDo;
        }
    }
}
