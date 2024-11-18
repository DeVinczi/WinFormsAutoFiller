using System.Globalization;
using System.Text.RegularExpressions;

namespace FormFiller.Models.WorkerEntity;

public static class PrzyznanaKwota
{
    public static string EnsureTwoDecimalPlaces(string input)
    {
        if (string.IsNullOrEmpty(input))
        {
            return "0,00";
        }
        
        // Try parsing the string as a decimal
        if (decimal.TryParse(input.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
        {
            // Format with two decimal places and replace '.' with ','
            return number.ToString("F2", CultureInfo.InvariantCulture).Replace(".", ",");
        }
        
        // If parsing fails, return the original input or handle error
        return input;
    }
}