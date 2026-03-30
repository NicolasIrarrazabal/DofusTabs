using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace DofusMiniTabber
{
    public class WindowPositionManager
    {
        private static readonly string ConfigPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "DofusMiniTabber",
            "window_positions.json"
        );

        public class WindowPosition
        {
            public string WindowName { get; set; } = string.Empty;
            public int Position { get; set; }
        }

        public class WindowConfiguration
        {
            public string Name { get; set; } = string.Empty;
            public List<WindowPosition> Positions { get; set; } = new();
            public DateTime CreatedAt { get; set; }
            public string Description { get; set; } = string.Empty;
        }

        public static void SaveConfiguration(string configName, List<WindowPosition> positions, string description = "")
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(ConfigPath)!);

                var configurations = LoadAllConfigurations();

                var newConfig = new WindowConfiguration
                {
                    Name        = configName,
                    Positions   = positions,
                    CreatedAt   = DateTime.Now,
                    Description = description
                };

                configurations[configName] = newConfig;

                var json = JsonSerializer.Serialize(configurations, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigPath, json);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error saving configuration: {ex.Message}");
            }
        }

        public static WindowConfiguration? LoadConfiguration(string configName)
        {
            try
            {
                var configurations = LoadAllConfigurations();
                return configurations.TryGetValue(configName, out var config) ? config : null;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error loading configuration: {ex.Message}");
            }
        }

        public static Dictionary<string, WindowConfiguration> LoadAllConfigurations()
        {
            try
            {
                if (!File.Exists(ConfigPath))
                    return new Dictionary<string, WindowConfiguration>();

                var json = File.ReadAllText(ConfigPath);
                return JsonSerializer.Deserialize<Dictionary<string, WindowConfiguration>>(json)
                       ?? new Dictionary<string, WindowConfiguration>();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error loading configurations: {ex.Message}");
            }
        }

        public static List<string> GetConfigurationNames()
        {
            try
            {
                var configurations = LoadAllConfigurations();
                return new List<string>(configurations.Keys);
            }
            catch
            {
                return new List<string>();
            }
        }

        public static List<WindowConfiguration> GetAllConfigurations()
        {
            try
            {
                var configurations = LoadAllConfigurations();
                return configurations.Values.ToList();
            }
            catch
            {
                return new List<WindowConfiguration>();
            }
        }

        public static void DeleteConfiguration(string configName)
        {
            try
            {
                var configurations = LoadAllConfigurations();
                configurations.Remove(configName);

                var json = JsonSerializer.Serialize(configurations, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigPath, json);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error deleting configuration: {ex.Message}");
            }
        }

        public static bool ConfigurationExists(string configName)
        {
            try
            {
                var configurations = LoadAllConfigurations();
                return configurations.ContainsKey(configName);
            }
            catch
            {
                return false;
            }
        }
    }
}
