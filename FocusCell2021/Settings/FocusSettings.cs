using System;
using System.Globalization;
using System.IO;
using FocusCell2021.Localization;

namespace FocusCell2021.Settings
{
    public enum FocusHighlightMode
    {
        RowAndColumn,
        RowOnly,
        ColumnOnly
    }

    public sealed class FocusSettings
    {
        private static readonly string SettingsDirectory =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), AppText.SettingsFolderName);

        private static readonly string SettingsPath = Path.Combine(SettingsDirectory, "settings.ini");

        public bool Enabled { get; set; } = true;
        public string HighlightColor { get; set; } = "#FFD54F";
        public double Opacity { get; set; } = 0.18;
        public bool ShowCellBorder { get; set; } = true;
        public double BorderOpacity { get; set; } = 0.95;
        public double BorderThickness { get; set; } = 2.0;
        public FocusHighlightMode Mode { get; set; } = FocusHighlightMode.RowAndColumn;
        public int RefreshIntervalMs { get; set; } = 55;
        public bool HideWhileMoving { get; set; } = true;

        public static FocusSettings Load()
        {
            var settings = new FocusSettings();
            try
            {
                if (!File.Exists(SettingsPath))
                    return settings;

                foreach (var rawLine in File.ReadAllLines(SettingsPath))
                {
                    var line = rawLine.Trim();
                    if (line.Length == 0 || line.StartsWith("#")) continue;

                    var p = line.IndexOf('=');
                    if (p < 1) continue;

                    var key = line.Substring(0, p).Trim();
                    var value = line.Substring(p + 1).Trim();

                    switch (key)
                    {
                        case "Enabled":
                            if (bool.TryParse(value, out var enabled)) settings.Enabled = enabled;
                            break;
                        case "HighlightColor":
                            if (!string.IsNullOrWhiteSpace(value)) settings.HighlightColor = value;
                            break;
                        case "Opacity":
                            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var opacity))
                                settings.Opacity = Clamp(opacity, 0.03, 0.75);
                            break;
                        case "ShowCellBorder":
                            if (bool.TryParse(value, out var showBorder)) settings.ShowCellBorder = showBorder;
                            break;
                        case "BorderOpacity":
                            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var borderOpacity))
                                settings.BorderOpacity = Clamp(borderOpacity, 0.1, 1.0);
                            break;
                        case "BorderThickness":
                            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var thickness))
                                settings.BorderThickness = Clamp(thickness, 1.0, 6.0);
                            break;
                        case "Mode":
                            if (Enum.TryParse(value, true, out FocusHighlightMode mode)) settings.Mode = mode;
                            break;
                        case "RefreshIntervalMs":
                            if (int.TryParse(value, out var interval)) settings.RefreshIntervalMs = Math.Max(30, Math.Min(250, interval));
                            break;
                        case "HideWhileMoving":
                            if (bool.TryParse(value, out var hideWhileMoving)) settings.HideWhileMoving = hideWhileMoving;
                            break;
                    }
                }
            }
            catch
            {
                // Corrupt settings should never prevent Excel from starting.
            }

            return settings;
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(SettingsDirectory);
                File.WriteAllLines(SettingsPath, new[]
                {
                    AppText.SettingsFileHeader,
                    $"Enabled={Enabled}",
                    $"HighlightColor={HighlightColor}",
                    $"Opacity={Opacity.ToString(CultureInfo.InvariantCulture)}",
                    $"ShowCellBorder={ShowCellBorder}",
                    $"BorderOpacity={BorderOpacity.ToString(CultureInfo.InvariantCulture)}",
                    $"BorderThickness={BorderThickness.ToString(CultureInfo.InvariantCulture)}",
                    $"Mode={Mode}",
                    $"RefreshIntervalMs={RefreshIntervalMs}",
                    $"HideWhileMoving={HideWhileMoving}"
                });
            }
            catch
            {
                // Settings persistence is optional.
            }
        }

        public void ResetVisualDefaults()
        {
            HighlightColor = "#FFD54F";
            Opacity = 0.18;
            ShowCellBorder = true;
            BorderOpacity = 0.95;
            BorderThickness = 2.0;
            Mode = FocusHighlightMode.RowAndColumn;
            RefreshIntervalMs = 55;
            HideWhileMoving = true;
        }

        private static double Clamp(double value, double min, double max)
            => Math.Max(min, Math.Min(max, value));
    }
}
