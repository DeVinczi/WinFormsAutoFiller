using System.Text.RegularExpressions;

using WinFormsAutoFiller.Infrastructure;
using WinFormsAutoFiller.Models.RegexEntity;
using WinFormsAutoFiller.Utilis;

namespace WinFormsAutoFiller.Helpers
{
    internal static class RegexHelpers
    {
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

        public static string GetProgram(string[] inputs)
        {
            try
            {
                foreach (var input in inputs)
                {
                    var match = RegexPatterns.GetProgramRegex().Match(input);
                    if (match.Success)
                    {
                        return input;
                    }
                }
                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public static string GetPUPCityName(string input)
        {
            try
            {
                var match = RegexPatterns.GetPUPCityName().Match(input);
                if (match.Success)
                {
                    return match.Groups[1].Value;
                }

                return string.Empty;
            }
            catch
            {
                return string.Empty;
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
    }
}
