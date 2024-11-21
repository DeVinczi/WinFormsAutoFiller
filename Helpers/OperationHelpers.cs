namespace WinFormsAutoFiller.Helpers
{
    internal class OperationHelpers
    {
        public static bool ValidateKRS(string krs)
        {
            // Check if KRS is numeric and has exactly 10 digits
            if (string.IsNullOrEmpty(krs) || krs.Length != 10 || !krs.All(char.IsDigit))
            {
                return false;
            }

            // Parse the first 9 digits
            int[] weights = { 9, 7, 3, 1, 9, 7, 3, 1, 9 }; // Weights for checksum calculation
            int checksum = 0;

            for (int i = 0; i < 9; i++)
            {
                checksum += (krs[i] - '0') * weights[i];
            }

            // Calculate modulo 10
            checksum %= 10;

            // Compare with the 10th digit
            int lastDigit = krs[9] - '0';
            return checksum == lastDigit;
        }
    }
}
