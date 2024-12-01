using System.Text.RegularExpressions;

namespace FormFiller.Models.WorkerEntity;

public static class KFSPriority
{
    public const string FirstValue = "1\\";
    public const string SecondValue = "2\\";
    public const string ThirdValue = "3\\";
    public const string FourthValue = "4\\";
    public const string FifthValue = "5\\";
    public const string SixthValue = "6\\";
    public const string SeventhValue = "7\\";
    public const string EighthValue = "8\\";
    public const string NinthValue = "9\\";
    public const string TenthValue = "10\\";
    public const string EleventhValue = "11\\";
    public const string TwelveValue = "12\\";
    public const string ThirteenthValue = "13\\";

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
            "9" => NinthValue,
            //Rezerwa
            "10" => TenthValue,
            "11" => EleventhValue,
            "12" => TwelveValue,
            "13" => ThirteenthValue,
            _ => throw new ArgumentOutOfRangeException(nameof(priority), "Priority number is out of range.")
        };
    }
}