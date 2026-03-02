namespace Drops_Tracker.Models
{
    public class TrackerData
    {
        public List<Character> Characters { get; set; } = new();
        public List<LootItem> Items { get; set; } = new();
        public List<WeeklyDrop> Drops { get; set; } = new();
    }
}
