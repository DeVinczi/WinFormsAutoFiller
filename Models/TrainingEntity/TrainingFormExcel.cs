using System.Text.RegularExpressions;

namespace FormFiller.Models.TrainingEntity;

public static class TrainingFormExcel
{
    public static readonly string[] Headers = new string[]
    {
        TrainingFormKeys.TematSzkolenia,
        TrainingFormKeys.Trener,
        TrainingFormKeys.IleOsob,
        TrainingFormKeys.IleDni,
        TrainingFormKeys.DniNaWniosku,
        TrainingFormKeys.DzienTrenera,
        TrainingFormKeys.WartośćSzkolenia,
        TrainingFormKeys.WartośćSzkoleniaNaOsobe,
        TrainingFormKeys.OfertaKonkurencji,
        TrainingFormKeys.SrednieDofinansowanieOsobe,
        TrainingFormKeys.Dofinansowanie,
        TrainingFormKeys.VAT
    };
}

public static class TrainingFormKeys
{
    public const string TematSzkolenia = "Temat szkolenia";
    public const string Trener = "Trener";
    public const string IleOsob = "ile osób";
    public const string IleDni = "ile dni";
    public const string DniNaWniosku = "dni na wniosku";
    public const string DzienTrenera = "dzień trenera";
    public const string WartośćSzkolenia = "wartość szkolenia";
    public const string WartośćSzkoleniaNaOsobe = "wartość szkolenia na osobę";
    public const string OfertaKonkurencji = "oferta konkurencji";
    public const string SrednieDofinansowanieOsobe = "średnie dofinansowanie na osobę";
    public const string Dofinansowanie = "dofinansowanie";
    public const string VAT = "VAT";
}

public static class TrainingFormPatterns
{
    public static readonly Dictionary<string, Regex> Patterns = new Dictionary<string, Regex>
    {
        { TrainingFormKeys.TematSzkolenia, new Regex(@"^[Tt]emat\s*[Ss]zkolenia$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { TrainingFormKeys.Trener, new Regex(@"^[Tt]rener$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { TrainingFormKeys.IleOsob, new Regex(@"^[Ii]le\s*[Oo]sób$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { TrainingFormKeys.IleDni, new Regex(@"^[Ii]le\s*[Dd]ni$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { TrainingFormKeys.DniNaWniosku, new Regex(@"^[Dd]ni\s*[Nn]a\s*[Ww]niosku$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { TrainingFormKeys.WartośćSzkolenia, new Regex(@"^[Ww]arto[śs][ćc]\s*[Ss]zkolenia$", RegexOptions.Compiled | RegexOptions.IgnoreCase) },
        { TrainingFormKeys.WartośćSzkoleniaNaOsobe, new Regex(@"^[Ww]arto[śs][ćc]\s*[Ss]zkolenia\s*na\s*[Oo]sob[ęe]$", RegexOptions.Compiled | RegexOptions.IgnoreCase) },
        { TrainingFormKeys.Dofinansowanie, new Regex(@"^[Dd]ofinansowanie$", RegexOptions.Compiled | RegexOptions.IgnoreCase)},
        { TrainingFormKeys.DzienTrenera, new Regex(@"^[Dd]zie[ńn]\s*[Tt]renera$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { TrainingFormKeys.OfertaKonkurencji, new Regex(@"^[Oo]ferta\s*[Kk]onkurencji$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { TrainingFormKeys.SrednieDofinansowanieOsobe, new Regex(@"^[Śś]rednie\s*[Dd]ofinansowanie\s*na\s*[Oo]sobę$", RegexOptions.Compiled|RegexOptions.IgnoreCase) },
        { TrainingFormKeys.VAT, new Regex(@"^VAT$", RegexOptions.Compiled|RegexOptions.IgnoreCase) }
    };
}
