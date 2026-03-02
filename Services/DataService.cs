using System.IO;
using Drops_Tracker.Models;

namespace Drops_Tracker.Services
{
    public class DataService
    {
        private readonly IniService _iniService;
        private TrackerData _data;
        private readonly string _assetsPath;

        public DataService(string assetsPath)
        {
            _assetsPath = assetsPath;
            _iniService = new IniService();
            _data = _iniService.LoadData();
        }

        public List<Character> Characters => _data.Characters;
        public List<WeeklyDrop> Drops => _data.Drops;

        // Items are loaded dynamically from assets folder
        public List<LootItem> GetItemsFromAssets()
        {
            var items = new List<LootItem>();

            if (!Directory.Exists(_assetsPath))
            {
                Directory.CreateDirectory(_assetsPath);
                return items;
            }

            var imageExtensions = new[] { ".png", ".jpg", ".jpeg", ".gif", ".bmp" };
            var imageFiles = Directory.GetFiles(_assetsPath)
                .Where(f => imageExtensions.Contains(Path.GetExtension(f).ToLower()))
                .OrderBy(f => Path.GetFileName(f))
                .ToList();

            foreach (var file in imageFiles)
            {
                var fileName = Path.GetFileName(file);
                var itemName = Path.GetFileNameWithoutExtension(file);

                items.Add(new LootItem
                {
                    Id = fileName, // Use filename as ID for consistency
                    Name = itemName,
                    ImageFileName = fileName
                });
            }

            return items;
        }

        public void SaveData()
        {
            _iniService.SaveData(_data);
        }

        public void AddCharacter(string name, string charClass, int level = 0)
        {
            _data.Characters.Add(new Character
            {
                Name = name,
                Class = charClass,
                Level = level
            });
            SaveData();
        }

        public bool UpdateCharacter(string id, string name, string charClass, int level = 0)
        {
            var character = _data.Characters.FirstOrDefault(c => c.Id == id);
            if (character == null)
            {
                return false;
            }

            character.Name = name;
            character.Class = charClass;
            character.Level = level;
            SaveData();
            return true;
        }

        public void DeleteCharacter(string id)
        {
            _data.Characters.RemoveAll(c => c.Id == id);
            _data.Drops.RemoveAll(d => d.CharacterId == id);
            SaveData();
        }

        public bool SwapCharacterOrder(string sourceCharacterId, string targetCharacterId)
        {
            var sourceIndex = _data.Characters.FindIndex(c => c.Id == sourceCharacterId);
            var targetIndex = _data.Characters.FindIndex(c => c.Id == targetCharacterId);

            if (sourceIndex < 0 || targetIndex < 0 || sourceIndex == targetIndex)
            {
                return false;
            }

            (_data.Characters[sourceIndex], _data.Characters[targetIndex]) =
                (_data.Characters[targetIndex], _data.Characters[sourceIndex]);

            SaveData();
            return true;
        }

        public WeeklyDrop? GetDrop(string characterId, string itemId, string weekKey)
        {
            return _data.Drops.FirstOrDefault(d =>
                d.CharacterId == characterId &&
                d.ItemId == itemId &&
                d.WeekKey == weekKey);
        }

        public List<WeeklyDrop> GetDropsForCharacterWeek(string characterId, string weekKey)
        {
            return _data.Drops.Where(d =>
                d.CharacterId == characterId &&
                d.WeekKey == weekKey).ToList();
        }

        public void AddDropToCell(string characterId, string weekKey, string itemId)
        {
            // Prevent duplicate items for the same character in the same week
            var existingDrop = GetDrop(characterId, itemId, weekKey);
            if (existingDrop != null)
            {
                // Item already exists for this character this week - do not add duplicate
                return;
            }

            _data.Drops.Add(new WeeklyDrop
            {
                CharacterId = characterId,
                ItemId = itemId,
                WeekKey = weekKey,
                Quantity = 1,
                Notes = string.Empty,
                IncludeInSummary = true
            });
            SaveData();
        }

        public bool SetDropIncludeInSummary(string characterId, string weekKey, string itemId, bool includeInSummary)
        {
            var existingDrop = GetDrop(characterId, itemId, weekKey);
            if (existingDrop == null)
            {
                return false;
            }

            if (existingDrop.IncludeInSummary == includeInSummary)
            {
                return true;
            }

            existingDrop.IncludeInSummary = includeInSummary;
            SaveData();
            return true;
        }

        public void RemoveDropFromCell(string characterId, string weekKey, string itemId)
        {
            var existingDrop = GetDrop(characterId, itemId, weekKey);
            if (existingDrop != null)
            {
                if (existingDrop.Quantity > 1)
                {
                    existingDrop.Quantity--;
                }
                else
                {
                    _data.Drops.Remove(existingDrop);
                }
                SaveData();
            }
        }

        public static string GetWeekKey(DateTime date)
        {
            var monday = GetWeekStart(date);
            return monday.ToString("yyyy-MM-dd");
        }

        public static DateTime GetWeekStart(DateTime date)
        {
            // Weekly reset is on Thursday
            int diff = (7 + (date.DayOfWeek - DayOfWeek.Thursday)) % 7;
            return date.Date.AddDays(-diff);
        }

        public static string FormatWeekDisplay(DateTime weekStart)
        {
            return $"Week of {weekStart:MMM dd, yyyy}";
        }

        public bool GetDarkModeSetting()
        {
            return _iniService.GetDarkModeSetting();
        }

        public void SetDarkModeSetting(bool isDarkMode)
        {
            _iniService.SetDarkModeSetting(isDarkMode);
        }
    }
}
