using System.Security.Cryptography;
using System.Text;

namespace WinFormsAutoFiller.Encryption
{
    public class AESBruteForceDecryption
    {
        public const string Key = "TopSecretApiKey1";

        public static string Decrypt(string cipherText)
        {
            try
            {
                using (var aesAlg = Aes.Create())
                {
                    aesAlg.Mode = CipherMode.CBC;
                    aesAlg.Padding = PaddingMode.PKCS7;

                    aesAlg.Key = Encoding.UTF8.GetBytes(Key.PadRight(16));
                    aesAlg.IV = Encoding.UTF8.GetBytes(Key.PadRight(16));

                    using (var decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV))
                    {
                        var cipherBytes = Convert.FromBase64String(cipherText);
                        byte[] resultBytes = decryptor.TransformFinalBlock(cipherBytes, 0, cipherBytes.Length);
                        return Encoding.UTF8.GetString(resultBytes);
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        public static string Encrypt(string cipherText)
        {
            try
            {
                using var aesAlg = Aes.Create();
                aesAlg.Mode = CipherMode.CBC;
                aesAlg.Padding = PaddingMode.PKCS7;

                aesAlg.Key = Encoding.UTF8.GetBytes(Key.PadRight(16));
                aesAlg.IV = Encoding.UTF8.GetBytes(Key.PadRight(16));

                using var encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
                var bytes = Encoding.UTF8.GetBytes(cipherText);
                byte[] resultBytes = encryptor.TransformFinalBlock(bytes, 0, cipherText.Length);
                var result = Convert.ToBase64String(resultBytes);

                return result;
            }
            catch (Exception ex)
            {
                return string.Empty;
            }
        }

        public static string BruteForceDecrypt(string cipherText)
        {
            string decryptedText = Decrypt(cipherText);
            return decryptedText;
        }
    }
}
