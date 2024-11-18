using System.Text.RegularExpressions;

namespace FormFiller.Models.WorkerEntity;

public static partial class Education
{
    public const string Wyzsze = "wyższe";
    public const string Policealne = "policealne i średnie zawodowe/branżowe";
    public const string SrednieOgolnoksztalcace = "średnie ogólnokształcące";
    public const string ZasadniczeZawodowe = "zasadnicze zawodowe/branżowe";
    public const string Gimnazjalne = "gimnazjalne/podstawowe i poniżej";

    public static string MapEducationLevel(string input)
    {
        input = input.ToLower().Trim();

        // "Wyzsze" - dla fraz zawierających słowa związane z wykształceniem wyższym
        if (MyRegex().IsMatch(input))
            return Wyzsze;

        // "Policealne" - dla fraz zawierających słowa związane z wykształceniem policealnym i średnim zawodowym/branżowym
        else if (MyRegex1().IsMatch(input))
            return Policealne;

        // "SrednieOgolnoksztalcace" - dla średniego ogólnokształcącego lub technicznego z maturą
        else if (MyRegex2().IsMatch(input))
            return SrednieOgolnoksztalcace;

        // "ZasadniczeZawodowe" - dla fraz zawierających wykształcenie zawodowe lub branżowe bez matury
        else if (MyRegex3().IsMatch(input))
            return ZasadniczeZawodowe;

        // "Gimnazjalne" - dla fraz związanych z gimnazjum lub szkołą podstawową
        else if (MyRegex4().IsMatch(input))
            return Gimnazjalne;

        return Gimnazjalne;
    }

    [GeneratedRegex(@"\bwyższe\b|\buniwersytet\b|\bmagister\b|\binżynier\b")]
    private static partial Regex MyRegex();
    [GeneratedRegex(@"\bpolicealne\b|\bśrednie\s+zawodowe\b|\bbranżowe\b")]
    private static partial Regex MyRegex1();
    [GeneratedRegex(@"\bśrednie\s+ogólnokształcące\b|\bśrednie\s+techniczne\b|\bmaturalne\b")]
    private static partial Regex MyRegex2();
    [GeneratedRegex(@"\bzawodowe\b|\bbranżowe\b")]
    private static partial Regex MyRegex3();
    [GeneratedRegex(@"\bpodstawowe\b|\bgimnazjalne\b")]
    private static partial Regex MyRegex4();
}