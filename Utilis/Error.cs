﻿namespace WinFormsAutoFiller.Utilis
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
    }

    public static class ErrorMessage
    {
        public const string NiepoprawnaNazwaPliku = "Niepoprawna nazwa pliku";
        public const string BrakIlosciGodzin = "Brak ilosci godzin w pliku word";
    }

    public static class GUIMessage
    {
        public const string BrakWybranegoPliku = "Brak wybranego pliku.";
    }
}