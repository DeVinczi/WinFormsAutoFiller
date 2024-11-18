using System.Text.RegularExpressions;

namespace FormFiller.Models.WorkerEntity;

public static class KFSPriority
{
    public const string FirstValue = " Wsparcie kształcenia ustawicznego w związku z zastosowaniem w firmach nowych procesów, technologii i narzędzi pracy";
    public const string SecondValue = " Wsparcie kształcenia ustawicznego w zidentyfikowanych w danym powiecie lub województwie zawodach deficytowych";
    public const string ThirdValue = " Wsparcie kształcenia ustawicznego osób powracających na rynek pracy po przerwie związanej ze sprawowaniem opieki nad dzieckiem oraz osób będących członkami rodzin wielodzietnych";
    public const string FourthValue = " Wsparcie kształcenia ustawicznego w zakresie umiejętności cyfrowych";
    public const string FifthValue = " Wsparcie kształcenia ustawicznego osób pracujących w branży motoryzacyjnej";
    public const string SixthValue = " Wsparcie kształcenia ustawicznego osób po 45 roku życia";
    public const string SeventhValue = " Wsparcie kształcenia ustawicznego skierowane do pracodawców zatrudniających cudzoziemców";
    public const string EighthValue = " Wsparcie kształcenia ustawicznego w zakresie zarządzania finansami i zapobieganie sytuacjom kryzysowym w przedsiębiorstwach";

    public static string MapPriority(string priority)
    {
        if (string.IsNullOrWhiteSpace(priority))
        {
            return string.Empty;
        }
        
        var reg = new Regex(@"(?<=Priorytet\s)\d+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        var match = reg.Match(priority);

        if (!match.Success)
        {
            return string.Empty;
        }
        
        return match.Value switch
        {
            "1" => FirstValue,
            "2" => SecondValue,
            "3" => ThirdValue,
            "4" => FourthValue,
            "5" => FifthValue,
            "6" => SixthValue,
            "7" => SeventhValue,
            "8" => EighthValue,
            _ => throw new ArgumentOutOfRangeException(nameof(priority), "Priority number is out of range.")
        };
    }
}