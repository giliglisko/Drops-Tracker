using System.IO;
using System.Text;
using Drops_Tracker.Models;

namespace Drops_Tracker.Services
{
    public class IniService
    {
        private readonly string _iniFilePath;
        private readonly string _settingsFilePath;

        public IniService()
        {
            var appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MapleDropTracker"
            );
            Directory.CreateDirectory(appDataPath);
            _iniFilePath = Path.Combine(appDataPath, "tracker_data.ini");
            _settingsFilePath = Path.Combine(appDataPath, "settings.ini");
        }

        public TrackerData LoadData()
        {
            var data = new TrackerData();
            var characterOrder = new Dictionary<string, int>();

            if (!File.Exists(_iniFilePath))
                return data;

            var lines = File.ReadAllLines(_iniFilePath);
            string currentSection = "";
            Character? currentCharacter = null;

            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith(";"))
                    continue;

                // Section header
                if (trimmed.StartsWith("[") && trimmed.EndsWith("]"))
                {
                    currentSection = trimmed.Substring(1, trimmed.Length - 2);

                    // Check if it's a character section
                    if (currentSection.StartsWith("Character."))
                    {
                        var charId = currentSection.Substring("Character.".Length);
                        currentCharacter = new Character { Id = charId };
                        data.Characters.Add(currentCharacter);
                    }
                    else
                    {
                        currentCharacter = null;
                    }
                    continue;
                }

                // Key=Value pair
                var eqIndex = trimmed.IndexOf('=');
                if (eqIndex > 0)
                {
                    var key = trimmed.Substring(0, eqIndex).Trim();
                    var value = trimmed.Substring(eqIndex + 1).Trim();

                    if (currentCharacter != null)
                    {
                        switch (key)
                        {
                            case "Name":
                                currentCharacter.Name = value;
                                break;
                            case "Class":
                                currentCharacter.Class = value;
                                break;
                            case "Level":
                                if (int.TryParse(value, out int level))
                                    currentCharacter.Level = level;
                                break;
                            case "ImagePath":
                                currentCharacter.LocalImagePath = value;
                                break;
                        }
                    }
                    else if (currentSection == "Drops")
                    {
                        // Format: CharacterId|WeekKey|ItemId=Quantity|Notes|IncludeInSummary
                        var keyParts = key.Split('|');
                        if (keyParts.Length == 3)
                        {
                            var valueParts = value.Split('|');
                            var drop = new WeeklyDrop
                            {
                                CharacterId = keyParts[0],
                                WeekKey = keyParts[1],
                                ItemId = keyParts[2],
                                Quantity = int.TryParse(valueParts[0], out int qty) ? qty : 1,
                                Notes = valueParts.Length > 1 ? valueParts[1] : "",
                                IncludeInSummary = valueParts.Length > 2
                                    ? valueParts[2].Equals("true", StringComparison.OrdinalIgnoreCase) || valueParts[2] == "1"
                                    : true
                            };
                            data.Drops.Add(drop);
                        }
                    }
                    else if (currentSection == "CharacterOrder")
                    {
                        if (int.TryParse(key, out var sortIndex) && !string.IsNullOrWhiteSpace(value))
                        {
                            characterOrder[value] = sortIndex;
                        }
                    }
                }
            }

            if (characterOrder.Count > 0)
            {
                data.Characters = data.Characters
                    .OrderBy(c => characterOrder.TryGetValue(c.Id, out var sortIndex) ? sortIndex : int.MaxValue)
                    .ToList();
            }

            return data;
        }

        public void SaveData(TrackerData data)
        {
            var sb = new StringBuilder();

            sb.AppendLine("; MapleStory Drop Tracker Data");
            sb.AppendLine($"; Last saved: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine();

            // Save characters
            foreach (var character in data.Characters)
            {
                sb.AppendLine($"[Character.{character.Id}]");
                sb.AppendLine($"Name={character.Name}");
                sb.AppendLine($"Class={character.Class}");
                sb.AppendLine($"Level={character.Level}");
                sb.AppendLine($"ImagePath={character.LocalImagePath}");
                sb.AppendLine();
            }

            sb.AppendLine("[CharacterOrder]");
            for (int i = 0; i < data.Characters.Count; i++)
            {
                sb.AppendLine($"{i}={data.Characters[i].Id}");
            }
            sb.AppendLine();

            // Save drops
            sb.AppendLine("[Drops]");
            foreach (var drop in data.Drops)
            {
                var key = $"{drop.CharacterId}|{drop.WeekKey}|{drop.ItemId}";
                var value = $"{drop.Quantity}|{drop.Notes}|{(drop.IncludeInSummary ? 1 : 0)}";
                sb.AppendLine($"{key}={value}");
            }

            File.WriteAllText(_iniFilePath, sb.ToString());
        }

        public bool GetDarkModeSetting()
        {
            if (!File.Exists(_settingsFilePath))
                return false;

            var lines = File.ReadAllLines(_settingsFilePath);
            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (trimmed.StartsWith("DarkMode=", StringComparison.OrdinalIgnoreCase))
                {
                    var value = trimmed.Substring("DarkMode=".Length).Trim();
                    return value.Equals("true", StringComparison.OrdinalIgnoreCase) || value == "1";
                }
            }
            return false;
        }

        public void SetDarkModeSetting(bool isDarkMode)
        {
            var sb = new StringBuilder();
            sb.AppendLine("; MapleStory Drop Tracker Settings");
            sb.AppendLine($"; Last saved: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine();
            sb.AppendLine("[Settings]");
            sb.AppendLine($"DarkMode={isDarkMode.ToString().ToLower()}");

            File.WriteAllText(_settingsFilePath, sb.ToString());
        }
    }
}
