using System.Text.RegularExpressions;

using WinFormsAutoFiller.Infrastructure;
using WinFormsAutoFiller.Models.RegexEntity;
using WinFormsAutoFiller.Utilis;

namespace WinFormsAutoFiller.Helpers
{
    internal static partial class RegexHelpers
    {
        private static readonly Regex DateRangeRegexs = DateRangeRegex();

        public static (int RangeCount, List<DateTime> Dates) ExtractDatesWithRangeInfo(string input)
        {
            var matches = DateRangeRegexs.Matches(input);
            var result = new List<DateTime>();
            int rangeCount = 0;

            foreach (Match match in matches)
            {
                if (match.Groups[1].Success) // Date range (e.g., 15-17.11.2024)
                {
                    int startDay = int.Parse(match.Groups[1].Value);
                    int endDay = int.Parse(match.Groups[2].Value);
                    int month = int.Parse(match.Groups[3].Value);
                    int year = int.Parse(match.Groups[4].Value);

                    result.Add(new DateTime(year, month, startDay));
                    result.Add(new DateTime(year, month, endDay));

                    rangeCount++; // Increment range counter
                }
                else if (match.Groups[5].Success) // Single date (e.g., 16.11.2024)
                {
                    int day = int.Parse(match.Groups[5].Value);
                    int month = int.Parse(match.Groups[6].Value);
                    int year = int.Parse(match.Groups[7].Value);

                    result.Add(new DateTime(year, month, day));
                }
            }

            return (rangeCount, result);
        }

        public static Result<CityAndDateModel, Error> GetName(string input)
        {
            try
            {
                var fileName = input[(input.LastIndexOf('\\') + 1)..];
                var containsCityAndDate = RegexPatterns.CityAndDateRegex().Match(fileName);
                if (containsCityAndDate.Success)
                {
                    var city = containsCityAndDate.Groups[1].Value;
                    var date = containsCityAndDate.Groups[2].Value;
                    if (DateTime.TryParse(date, out var dateParsed))
                    {
                        return new CityAndDateModel
                        {
                            City = city,
                            StartDate = dateParsed
                        };
                    }
                }

                return Errors.IncorrectCityOrDate;
            }
            catch
            {
                return Errors.IncorrectCityOrDate;
            }
        }

        public static CityAndDateModel GetNameFromExcel(string input)
        {
            try
            {
                var cityAndDate = RegexPatterns.GetCityAndDateExcel().Match(input);
                if (cityAndDate.Success)
                {
                    var city = cityAndDate.Groups[1].Value;
                    var date = cityAndDate.Groups[2].Value;
                    if (DateTime.TryParse(date, out var dateParsed))
                    {
                        return new CityAndDateModel
                        {
                            City = city,
                            StartDate = dateParsed
                        };
                    }
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        public static Result<Success, Error> CheckDaneOgolneRegex(string input)
        {
            try
            {
                var fileName = input[(input.LastIndexOf('\\') + 1)..];
                var containsDaneOgolne = RegexPatterns.CheckDaneOgolneRegex().Match(fileName);
                if (containsDaneOgolne.Success)
                {
                    return new Success();
                }

                return Errors.IncorrectNameFile;
            }
            catch
            {
                return Errors.IncorrectNameFile;
            }
        }

        public static Dictionary<string, string> GetProgram(string[] inputs)
        {
            Dictionary<string, string> list = [];
            try
            {
                foreach (var input in inputs)
                {
                    var match = RegexPatterns.GetProgramRegex().Match(input);
                    if (match.Success)
                    {
                        list.Add(input, match.Value);
                    }
                }

                return list;
            }
            catch
            {
                return [];
            }
        }

        public static string ExtractMatch(string input, string pattern)
        {
            Match match = Regex.Match(input, pattern, RegexOptions.IgnoreCase);
            return match.Success ? match.Value : null;
        }

        public static string FindMatchInWorksheets(string target, List<string> worksheets)
        {
            foreach (var sheet in worksheets)
            {
                if (IsMatch(target, sheet))
                {
                    return sheet; // Exact match
                }
            }
            return null; // No match found
        }

        public static bool IsMatch(string str1, string str2)
        {
            string normalized1 = NormalizeString(str1);
            string normalized2 = NormalizeString(str2);

            return normalized1 == normalized2;
        }

        public static string NormalizeString(string input)
        {
            string withoutSeparators = Regex.Replace(input, @"[_\.\s]", "");
            return withoutSeparators.ToLower();
        }

        [GeneratedRegex(@"(?:(\d{1,2})-(\d{1,2})\.(\d{1,2})\.(\d{4})|(\d{1,2})\.(\d{1,2})\.(\d{4}))", RegexOptions.Compiled)]
        private static partial Regex DateRangeRegex();
    }
}
