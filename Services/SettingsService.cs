using System;
using System.IO;
using System.Text.Json;

namespace TrayFolder.Services
{
    public class AppSettings
    {
        public string RootFolderPath { get; set; } = "C:\\";
    }

    public static class SettingsService
    {
        private static readonly string SettingsFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "TaskbarFolderShortcut",
            "settings.json");

        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(SettingsFile))
                {
                    var json = File.ReadAllText(SettingsFile);
                    return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                }
            }
            catch { }
            return new AppSettings();
        }

        public static void Save(AppSettings settings)
        {
            try
            {
                var dir = Path.GetDirectoryName(SettingsFile);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                
                var json = JsonSerializer.Serialize(settings);
                File.WriteAllText(SettingsFile, json);
            }
            catch { }
        }
    }
}
