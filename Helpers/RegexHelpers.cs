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
    }
}
