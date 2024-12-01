using System.Text.RegularExpressions;

namespace WinFormsAutoFiller.Models
{
    public class Address
    {
        public string PostalCode { get; set; }
        public string City { get; set; }
        public string Street { get; set; }
        public string HouseNumber { get; set; }
        public string FlatNumber { get; set; }
    }

    public static class AddressHelpers
    {
        public static List<Address> GetAddresses(string input)
        {
            try
            {
                var address = input
                .Split('\n')
                .Select(line =>
                {
                    // Split line into parts
                    var parts = line.Split(new[] { ' ' }, 3);
                    var code = parts[0];
                    var city = parts[1].TrimEnd(',');

                    string addressPart = parts.Length > 2 ? parts[2] : "";
                    string street = ExtractStreet(addressPart);
                    string houseNumber = ExtractHouseNumber(addressPart);
                    string flatNumber = ExtractFlatNumber(addressPart);

                    return new Address
                    {
                        PostalCode = code,
                        City = city,
                        Street = street,
                        HouseNumber = houseNumber,
                        FlatNumber = flatNumber
                    };
                }).ToList();

                return address;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ostrzeżenie", MessageBoxButtons.CancelTryContinue, MessageBoxIcon.Warning);
                return [];
            }
        }

        public static string ExtractStreet(string address)
        {
            var match = Regex.Match(address, @"^\s*(?:ul\.?\s+)?(\D.*?)(?:\s+\d.*)?$", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value.Trim() : "";
        }

        static string ExtractHouseNumber(string address)
        {
            var match = Regex.Match(address, @"\b\d+");
            return match.Success ? match.Value : "";
        }

        static string ExtractFlatNumber(string address)
        {
            var match = Regex.Match(address, @"/(\d+)");
            return match.Success ? match.Groups[1].Value : "";
        }
    }
}
