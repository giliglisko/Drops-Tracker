namespace Drops_Tracker.Models
{
    [Serializable]
    public class LootItem
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Name { get; set; } = string.Empty;
        public string ImageFileName { get; set; } = string.Empty;
    }
}
