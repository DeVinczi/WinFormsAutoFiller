namespace FormFiller.Models.WorkerEntity;

public static class Age
{
    public const string A15B24 = "15-24";
    public const string A25B34 = "25-34";
    public const string A35B44 = "35-44";
    public const string A45Plus = "45 i więcej";

    public static string GetAge(int age)
    {
        if (15 <= age && age <= 24)
        {
            return A15B24;
        }

        if (25 <= age && age <= 34)
        {
            return A25B34;
        }

        if (35 <= age && age <= 44)
        {
            return A35B44;
        }

        if (45 <= age)
        {
            return A45Plus;
        }
        
        return A15B24;
    }
}