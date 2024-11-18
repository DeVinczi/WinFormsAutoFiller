namespace FormFiller.Models.WorkerEntity;

public static class Gender
{
    public const string Kobieta = "kobieta";
    public const string Mezczyzna = "mężczyzna";

    public static string MapGender(string gender)
    {
        if (gender == "K")
        {
            return Kobieta;
        }
        else
        {
            return Mezczyzna;   
        }
    }
}