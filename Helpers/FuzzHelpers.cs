using FuzzySharp;

namespace WinFormsAutoFiller.Helpers
{
    internal class FuzzHelpers
    {
        public static string GetAddress(string inputs, string city)
        {
            Dictionary<int, string> dict = [];
            var list = inputs.Split('\n');

            for (int i = 0; i < list.Length; i++)
            {
                int similarity = Fuzz.Ratio(list[i], city);
                dict.Add(i, list[i]);
            }

            var ordered = dict.OrderByDescending(x => x.Key);
            var x = ordered.First();
            return x.Value;
        }
    }
}
