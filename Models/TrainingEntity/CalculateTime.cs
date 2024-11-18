namespace FormFiller.Models.TrainingEntity;

public static class CalculateTime
{
    public static string Calculate(string days)
    {
        if (!string.IsNullOrWhiteSpace(days))
        {
            int.TryParse(days, out var day);
            return (day * 8).ToString();
        }
        return string.Empty;
    }
}