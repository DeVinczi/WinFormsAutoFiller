namespace WinFormsAutoFiller.Utilis
{
    public sealed record Error(string Code, string? Message = null);

    public static class Errors
    {
        public static readonly Error IncorrectCityOrDate =
            new(ErrorMessage.NiepoprawnaNazwaPliku, "Niepoprawna nazwa pliku z wyceną.\n" +
                "Miejscowość lub data nierozpoznana.\n" +
                "Poprawny schemat: KFS_NAZWA KLIENTA - MIASTO_DATA_v1");

        public static readonly Error IncorrectNameFile =
            new(ErrorMessage.NiepoprawnaNazwaPliku, "Niepoprawna nazwa pliku.\n" +
                "Poprawny schemat: Dane ogólne i pracowników_NAZWA KLIENTA_KFS 2025.xlsx");

        public static readonly Error WorkHoursAreEmpty =
            new(ErrorMessage.BrakIlosciGodzin, "Nie znaleziono godzin w pliku z programem szkolenia.");

        public static readonly Error ChromeProccessingError = new
            (ErrorMessage.ChromeError, "Wystapił bład podczas dodawania informacji do formularza.");

        public static Error PytaniaDoWnioskuError = new
            (ErrorMessage.PytaniaDoWniosku, "Zakładka 'Pytania do wniosku' nie znalezione.");

        public static readonly Error BrakPlikuNipError = new(ErrorMessage.BrakPlikuNip, "Wystąpił bład. Nie ma plików w folderze z NIP.");
        public static readonly Error PlikZZałącznikaNieWystepujeError = new(ErrorMessage.PlikZZałącznikaNieWystepujeError, "Plik z załącznika nie wystepuje. Sprawdź nazwy!");
    }

    public static class ErrorMessage
    {
        public const string NiepoprawnaNazwaPliku = "Niepoprawna nazwa pliku";
        public const string BrakIlosciGodzin = "Brak ilosci godzin w pliku word";
        public const string ChromeError = "Wystapił bład podczas dodawania informacji";
        public const string PytaniaDoWniosku = "Wystąpił bład z danymi do wniosku.";
        public const string BrakPlikuNip = "Wystąpił bład. Nie ma plików w folderze z NIP.";
        public const string PlikZZałącznikaNieWystepujeError = "Plik z załącznika nie wystepuje. Sprawdź nazwy!";
    }

    public static class GUIMessage
    {
        public const string BrakWybranegoPliku = "Brak wybranego pliku.";
    }
}
