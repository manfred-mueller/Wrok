using System.Diagnostics;
using Microsoft.Win32;

namespace Wrok
{
    /// <summary>
    /// Verwaltet den Windows-Autostart über den HKCU\...\Run-Registry-Schlüssel.
    /// Kein Admin-Recht nötig (nur HKEY_CURRENT_USER).
    /// </summary>
    internal static class AutostartManager
    {
        private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string ValueName  = "Wrok";

        /// <summary>Pfad zur aktuellen ausführbaren Datei (single-file-tauglich).</summary>
        private static string ExecutablePath =>
            Environment.ProcessPath ?? Application.ExecutablePath;

        /// <summary>
        /// Gibt true zurück, wenn der Autostart-Eintrag existiert UND auf den
        /// aktuellen Programmpfad zeigt (verhindert veraltete Einträge nach Umzug).
        /// </summary>
        public static bool IsEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: false);
                var stored = key?.GetValue(ValueName) as string;
                if (string.IsNullOrWhiteSpace(stored)) return false;

                var current = $"\"{ExecutablePath}\"";
                return string.Equals(stored.Trim(), current, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [AutostartManager] IsEnabled fehlgeschlagen: {ex}");
                return false;
            }
        }

        /// <summary>Aktiviert oder deaktiviert den Autostart. Gibt true bei Erfolg zurück.</summary>
        public static bool SetEnabled(bool enabled)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: true)
                                ?? Registry.CurrentUser.CreateSubKey(RunKeyPath);
                if (key == null) return false;

                if (enabled)
                    key.SetValue(ValueName, $"\"{ExecutablePath}\"");
                else if (key.GetValue(ValueName) != null)
                    key.DeleteValue(ValueName, throwOnMissingValue: false);

                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [AutostartManager] SetEnabled({enabled}) fehlgeschlagen: {ex}");
                return false;
            }
        }
    }
}
