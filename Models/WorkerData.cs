namespace AutoFiller.Models
{
    public class WorkerData
    {
        public string Name { get; set; }
        public Dictionary<string, bool> ColumnData { get; set; } = [];
    }
}

