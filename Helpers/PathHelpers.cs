using WinFormsAutoFiller.Utilis;

namespace WinFormsAutoFiller.Helpers
{
    internal static class PathHelpers
    {
        public const string DE_MINIMIS = "DE MINIMIS";
        public const string CIT = "CIT";
        public const string NIP8 = "NIP8";
        public const string OŚWIADCZENIE_DE_MINIMIS = "OŚWIADCZENIE DE MINIMIS";
        public const string PROGRAM = "Program";
        public const string PEŁNOMOCNICTWO = "Pełnomocnictwo";
        public const string WZÓR_CERTYFIKATU = "WZÓR CERTYFIKATU";
        public const string UMOWA_NAJMU_LOKALIZACJA = "Umowa Najmu_";
        public const string OŚWIADCZENIE_CERTYFIKAT_KOMPLET = "OŚWIADCZENIE + CERTYFIKAT KOMPLET";
        public const string ZAKRES_EGZAMINU = "Zakres_";

        public static Result<Dictionary<string, List<string>>, Error> GetRightPath(string[] paths)
        {
            try
            {
                var dict = new Dictionary<string, List<string>>();

                foreach (var path in paths)
                {
                    var file = Path.GetFileName(path);

                    if (file.StartsWith(DE_MINIMIS, StringComparison.InvariantCultureIgnoreCase))
                    {
                        AddToDictionary(dict, DE_MINIMIS, path);
                        continue;
                    }
                    if (file.StartsWith(CIT, StringComparison.InvariantCultureIgnoreCase))
                    {
                        AddToDictionary(dict, CIT, path);
                        continue;
                    }
                    if (file.StartsWith(NIP8, StringComparison.InvariantCultureIgnoreCase))
                    {
                        AddToDictionary(dict, NIP8, path);
                        continue;
                    }
                    if (file.StartsWith(OŚWIADCZENIE_DE_MINIMIS, StringComparison.InvariantCultureIgnoreCase))
                    {
                        AddToDictionary(dict, OŚWIADCZENIE_DE_MINIMIS, path);
                        continue;
                    }
                    if (file.StartsWith(PROGRAM, StringComparison.InvariantCultureIgnoreCase))
                    {
                        AddToDictionary(dict, PROGRAM, path);
                        continue;
                    }
                    if (file.StartsWith(WZÓR_CERTYFIKATU, StringComparison.InvariantCultureIgnoreCase))
                    {
                        AddToDictionary(dict, WZÓR_CERTYFIKATU, path);
                        continue;
                    }
                    if (file.StartsWith(UMOWA_NAJMU_LOKALIZACJA, StringComparison.InvariantCultureIgnoreCase))
                    {
                        AddToDictionary(dict, UMOWA_NAJMU_LOKALIZACJA, path);
                        continue;
                    }
                    if (file.StartsWith(OŚWIADCZENIE_CERTYFIKAT_KOMPLET, StringComparison.InvariantCultureIgnoreCase))
                    {
                        AddToDictionary(dict, OŚWIADCZENIE_CERTYFIKAT_KOMPLET, path);
                        continue;
                    }
                    if (file.StartsWith(ZAKRES_EGZAMINU, StringComparison.InvariantCultureIgnoreCase))
                    {
                        AddToDictionary(dict, ZAKRES_EGZAMINU, path);
                        continue;
                    }
                    if (file.StartsWith(PEŁNOMOCNICTWO, StringComparison.OrdinalIgnoreCase))
                    {
                        AddToDictionary(dict, PEŁNOMOCNICTWO, path);
                        continue;
                    }
                }

                //if (dict.Count < 3)
                //{
                //    return Errors.PlikZZałącznikaNieWystepujeError;
                //}

                return dict;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return new Error("Error", "");
            }
        }

        private static void AddToDictionary(Dictionary<string, List<string>> dict, string key, string path)
        {
            if (!dict.ContainsKey(key))
            {
                dict[key] = new List<string>();
            }
            dict[key].Add(path);
        }
    }
}
