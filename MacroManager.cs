using System.Diagnostics;
using System.Text.Json;
using System.Windows.Forms;

namespace Wrok
{
    // ---------------------------------------------------------------------------
    // Makro-Datenmodell
    // ---------------------------------------------------------------------------

    /// <summary>
    /// Ein einzelner Makroeintrag mit optionalem Anzeigenamen und Text.
    /// Name leer → UI zeigt "Macro #n" als Fallback.
    /// </summary>
    internal record MacroEntry(string Name, string Text)
    {
        public string DisplayName(int oneBasedIndex) =>
            string.IsNullOrWhiteSpace(Name)
                ? string.Format(Properties.Resources.Macro0, oneBasedIndex)
                : Name;
    }

    // ---------------------------------------------------------------------------
    // MacroManager
    // ---------------------------------------------------------------------------

    /// <summary>
    /// Verwaltet das Laden, Cachen, Speichern und Ausführen von Makros.
    /// Makros werden beim ersten Zugriff einmalig deserialisiert und danach
    /// im Speicher gehalten – kein JSON-Roundtrip bei jedem Hotkey-Drücken.
    /// </summary>
    internal sealed class MacroManager
    {
        // Anzahl unterstützter Makros; Hotkeys Ctrl+1 .. Ctrl+MacroCount.
        public const int MacroCount = 10;

        // Basis-ID für Makro-Hotkeys (muss mit MainForm übereinstimmen).
        public const int HotkeyBase = 0x9100;

        private List<MacroEntry>? _cache;
        private readonly object _lock = new();

        // -------------------------------------------------------------------
        // Öffentliche API
        // -------------------------------------------------------------------

        /// <summary>Gibt die gecachte Makroliste zurück (lazy load).</summary>
        public List<MacroEntry> GetMacros()
        {
            lock (_lock)
            {
                if (_cache == null)
                    _cache = LoadFromSettings();
                return _cache;
            }
        }

        /// <summary>Aktualisiert einen Eintrag und persistiert sofort.</summary>
        public void UpdateMacro(int index, MacroEntry entry)
        {
            lock (_lock)
            {
                var macros = GetMacros();
                if (index >= 0 && index < macros.Count)
                    macros[index] = entry;
                PersistToSettings(macros);
            }
        }

        /// <summary>Ersetzt die gesamte Liste und persistiert sofort.</summary>
        public void SaveAll(List<MacroEntry> macros)
        {
            lock (_lock)
            {
                _cache = macros;
                PersistToSettings(macros);
            }
        }

        /// <summary>
        /// Gibt den Text des per Hotkey-ID adressierten Makros zurück,
        /// mit aufgelösten Variablen. Gibt null zurück wenn leer oder ungültig.
        /// </summary>
        public string? GetTextForHotkey(int hotkeyId)
        {
            int macroIndex = hotkeyId - HotkeyBase;
            if (macroIndex < 0 || macroIndex >= MacroCount)
                return null;

            var macros = GetMacros();
            string text = macroIndex < macros.Count ? macros[macroIndex].Text : string.Empty;
            if (string.IsNullOrWhiteSpace(text)) return null;
            return ExpandVariables(text);
        }

        /// <summary>
        /// Gibt den ROHEN (nicht expandierten) Text des per Hotkey-ID adressierten
        /// Makros zurück, oder null wenn leer/ungültig. Wird vom Host genutzt, der
        /// Variablen inkl. {input} selbst auflöst (für die UI-Eingabe).
        /// </summary>
        public string? GetRawTextForHotkey(int hotkeyId)
        {
            int macroIndex = hotkeyId - HotkeyBase;
            if (macroIndex < 0 || macroIndex >= MacroCount)
                return null;

            var macros = GetMacros();
            string text = macroIndex < macros.Count ? macros[macroIndex].Text : string.Empty;
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }

        /// <summary>
        /// Ersetzt nicht-interaktive Variablen im Makrotext:
        ///   {date}      → aktuelles Datum (dd.MM.yyyy)
        ///   {time}      → aktuelle Uhrzeit (HH:mm)
        ///   {clipboard} → aktueller Zwischenablage-Inhalt (Text)
        ///   {username}  → Windows-Benutzername
        /// Hinweis: {input} bzw. {input:Label} wird NICHT hier aufgelöst, sondern
        /// vom Host (MainForm) per Eingabedialog zur Sendezeit.
        /// </summary>
        public static string ExpandVariables(string text)
        {
            var now = DateTime.Now;
            text = text.Replace("{date}",     now.ToString("dd.MM.yyyy"),           StringComparison.OrdinalIgnoreCase);
            text = text.Replace("{time}",     now.ToString("HH:mm"),                StringComparison.OrdinalIgnoreCase);
            text = text.Replace("{username}", Environment.UserName,                 StringComparison.OrdinalIgnoreCase);
            text = text.Replace("{clipboard}", GetClipboardText(),                  StringComparison.OrdinalIgnoreCase);
            return text;
        }

        private static string GetClipboardText()
        {
            try
            {
                return Clipboard.ContainsText() ? (Clipboard.GetText() ?? string.Empty) : string.Empty;
            }
            catch { return string.Empty; }
        }

        // -------------------------------------------------------------------
        // Import / Export
        // -------------------------------------------------------------------

        /// <summary>
        /// Exportiert die aktuelle Makroliste als JSON-Datei.
        /// Gibt true zurück wenn erfolgreich.
        /// </summary>
        public bool Export(string filePath)
        {
            try
            {
                var macros = GetMacros();
                var json   = JsonSerializer.Serialize(macros, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(filePath, json, System.Text.Encoding.UTF8);
                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroManager] Export fehlgeschlagen: {ex}");
                return false;
            }
        }

        /// <summary>
        /// Importiert Makros aus einer JSON-Datei und überschreibt die aktuelle Liste.
        /// Gibt true zurück wenn erfolgreich.
        /// </summary>
        public bool Import(string filePath)
        {
            try
            {
                var json = File.ReadAllText(filePath, System.Text.Encoding.UTF8);
                var list = JsonSerializer.Deserialize<List<MacroEntry>>(json);
                if (list == null) return false;

                while (list.Count < MacroCount)
                    list.Add(new MacroEntry(string.Empty, string.Empty));
                if (list.Count > MacroCount)
                    list = list.Take(MacroCount).ToList();

                SaveAll(list);
                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroManager] Import fehlgeschlagen: {ex}");
                return false;
            }
        }

        // -------------------------------------------------------------------
        // Private Helpers
        // -------------------------------------------------------------------

        private List<MacroEntry> LoadFromSettings()
        {
            // 1. Neues Format: MacrosJson (JSON-String)
            try
            {
                var json = Properties.Settings.Default.MacrosJson;
                if (!string.IsNullOrWhiteSpace(json))
                {
                    var list = JsonSerializer.Deserialize<List<MacroEntry>>(json);
                    if (list != null)
                    {
                        while (list.Count < MacroCount)
                            list.Add(new MacroEntry(string.Empty, string.Empty));
                        return list;
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroManager] LoadFromSettings (JSON) failed: {ex}");
            }

            // 2. Migration: altes Format Macros (StringCollection) → MacrosJson
            try
            {
                var legacy = Properties.Settings.Default.Macros;
                if (legacy != null && legacy.Count > 0)
                {
                    Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroManager] Migriere {legacy.Count} Makros aus altem Format.");

                    var migrated = new List<MacroEntry>();
                    foreach (var item in legacy)
                        migrated.Add(new MacroEntry(string.Empty, item ?? string.Empty));

                    while (migrated.Count < MacroCount)
                        migrated.Add(new MacroEntry(string.Empty, string.Empty));

                    // Sofort ins neue Format persistieren
                    PersistToSettings(migrated);

                    // Altes Feld leeren damit nicht doppelt migriert wird
                    Properties.Settings.Default.Macros = new System.Collections.Specialized.StringCollection();
                    Properties.Settings.Default.Save();

                    return migrated;
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroManager] Migration fehlgeschlagen: {ex}");
            }

            // 3. Erster Start oder korrupte Daten → leere Standardeinträge.
            return Enumerable.Range(0, MacroCount)
                             .Select(_ => new MacroEntry(string.Empty, string.Empty))
                             .ToList();
        }

        private static void PersistToSettings(List<MacroEntry> macros)
        {
            try
            {
                Properties.Settings.Default.MacrosJson = JsonSerializer.Serialize(macros);
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroManager] PersistToSettings failed: {ex}");
            }
        }
    }
}
