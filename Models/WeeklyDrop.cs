namespace Drops_Tracker.Models
{
    public class WeeklyDrop
    {
        public string CharacterId { get; set; } = string.Empty;
        public string ItemId { get; set; } = string.Empty;
        public string WeekKey { get; set; } = string.Empty;
        public int Quantity { get; set; }
        public string Notes { get; set; } = string.Empty;
        public bool IncludeInSummary { get; set; } = true;
    }
}
