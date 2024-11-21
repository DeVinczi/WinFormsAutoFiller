using System.Text.RegularExpressions;

namespace WinFormsAutoFiller.Infrastructure
{
    internal static partial class RegexPatterns
    {
        [GeneratedRegex(@"([\p{L}]+)_(\d{4}\.\d{2}\.\d{2})_.*\.(xlsx|xls|xlsm|xlsb|xltx|xlw)", RegexOptions.IgnoreCase, 3000)]
        internal static partial Regex CityAndDateRegex();

        [GeneratedRegex(@"([A-Za-z0-9\s]+)_([A-Za-z0-9\s]+)_KFS\s(\d{4})", RegexOptions.IgnoreCase, 3000)]
        internal static partial Regex CheckDaneOgolneRegex();

        [GeneratedRegex(@"(\d+(?:,\d+)?)\s*(godzin?)(?:\([^\)]*\))?", RegexOptions.IgnoreCase, 3000)]
        internal static partial Regex FindOutHoursRegex();


        [GeneratedRegex(@"(?i)\\(\bprogram\b)\s*-\s*[^\\]+(?:\.docx?|\.dotx?|\.docm?|\.dotm?|\.rtf|\.wps)$", RegexOptions.IgnoreCase, 3000)]
        internal static partial Regex GetProgramRegex();
    }
}
