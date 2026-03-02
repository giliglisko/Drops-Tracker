namespace Drops_Tracker.Models
{
    public class Character
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Name { get; set; } = string.Empty;
        public string Class { get; set; } = string.Empty;
        public int Level { get; set; }
        public string ImageUrl { get; set; } = string.Empty;
        public string LocalImagePath { get; set; } = string.Empty;
    }
}
