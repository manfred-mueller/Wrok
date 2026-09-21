using System.Diagnostics;
using System.Globalization;
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

    /// <summary>
    /// Ein Satz Makros unter einem Namen – etwa je Figur oder Szenario. Jedes
    /// Profil hat einen eigenen, immer aktiven Hotkey (Strg+F1..Strg+F{n}, siehe
    /// MainForm.ProfileMenuKeys), der ein Auswahl-Popup mit den Makrotiteln
    /// dieses Profils öffnet – unabhängig davon, welches Profil gerade im
    /// Tray-Menü zum Bearbeiten angezeigt wird ("aktives" Profil, siehe
    /// <see cref="MacroManager.ActiveProfile"/>).
    /// </summary>
    internal record MacroProfile(string Name, List<MacroEntry> Macros)
    {
        public string DisplayName(int oneBasedIndex) =>
            string.IsNullOrWhiteSpace(Name)
                ? string.Format(Properties.Resources.MacroProfileDefault, oneBasedIndex)
                : Name;
    }

    /// <summary>Was eine Importdatei enthält.</summary>
    internal enum MacroImportKind
    {
        Invalid,
        /// <summary>Schlichte Makroliste – muss einem Zielprofil zugeordnet werden.</summary>
        SingleProfile,
        /// <summary>Vollständiger Satz inklusive Profilnamen.</summary>
        AllProfiles
    }

    /// <summary>Dateiformat für den Export aller Profile.</summary>
    internal record MacroProfileExport(int WrokMacroExport, List<MacroProfile> Profiles);

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
        // Makros pro Profil, erreichbar über das Auswahl-Popup (Ziffer 0-9) nach
        // Öffnen des jeweiligen Profil-Menüs per Strg+F-Taste (siehe MainForm).
        public const int MacroCount = 10;

        /// <summary>Anzahl der Makro-Profile (z. B. je Figur oder Szenario).</summary>
        public const int ProfileCount = 3;

        /// <summary>Version des Exportformats für alle Profile.</summary>
        private const int ExportFormatVersion = 2;

        private List<MacroProfile>? _profiles;
        private int _activeIndex = -1;
        private readonly object _lock = new();

        // -------------------------------------------------------------------
        // Profile
        // -------------------------------------------------------------------

        /// <summary>
        /// Index des Profils, das im Tray-Menü angezeigt/bearbeitet wird. Hat
        /// KEINEN Einfluss darauf, welche Makros per Hotkey ausgelöst werden
        /// können – alle Profile sind dafür immer gleichzeitig aktiv.
        /// </summary>
        public int ActiveProfile
        {
            get { lock (_lock) { EnsureLoaded(); return _activeIndex; } }
        }

        /// <summary>Namen aller Profile, in Reihenfolge.</summary>
        public List<string> GetProfileNames()
        {
            lock (_lock)
            {
                EnsureLoaded();
                return _profiles!.Select(p => p.Name).ToList();
            }
        }

        /// <summary>
        /// Wechselt das im Tray-Menü angezeigte/bearbeitete Profil. Betrifft nur
        /// die Anzeige – die Hotkeys aller drei Profile bleiben unverändert aktiv.
        /// </summary>
        public void SwitchProfile(int index)
        {
            lock (_lock)
            {
                EnsureLoaded();
                if (index < 0 || index >= _profiles!.Count || index == _activeIndex) return;

                _activeIndex = index;
                Properties.Settings.Default.ActiveMacroProfile = index;
                Properties.Settings.Default.Save();
            }
        }

        /// <summary>
        /// Leert ein Profil vollstaendig: Name und alle Makros. Das Profil selbst
        /// bleibt bestehen – es gibt immer genau ProfileCount Stueck.
        /// </summary>
        public void ResetProfile(int index)
        {
            lock (_lock)
            {
                EnsureLoaded();
                if (index < 0 || index >= _profiles!.Count) return;

                _profiles[index] = new MacroProfile(string.Empty, EmptyMacros());
                Persist();
            }
        }

        /// <summary>Benennt ein Profil um und persistiert sofort.</summary>
        public void RenameProfile(int index, string name)
        {
            lock (_lock)
            {
                EnsureLoaded();
                if (index < 0 || index >= _profiles!.Count) return;

                _profiles[index] = _profiles[index] with { Name = name.Trim() };
                Persist();
            }
        }

        // -------------------------------------------------------------------
        // Öffentliche API
        // -------------------------------------------------------------------

        /// <summary>Gibt die Makroliste des aktiven Profils zurück (lazy load).</summary>
        public List<MacroEntry> GetMacros()
        {
            lock (_lock)
            {
                EnsureLoaded();
                return _profiles![_activeIndex].Macros;
            }
        }

        /// <summary>Aktualisiert einen Eintrag im aktiven Profil und persistiert sofort.</summary>
        public void UpdateMacro(int index, MacroEntry entry)
        {
            lock (_lock)
            {
                var macros = GetMacros();
                if (index >= 0 && index < macros.Count)
                    macros[index] = entry;
                Persist();
            }
        }

        /// <summary>Ersetzt die Makros des aktiven Profils und persistiert sofort.</summary>
        public void SaveAll(List<MacroEntry> macros)
        {
            lock (_lock)
            {
                EnsureLoaded();
                _profiles![_activeIndex] = _profiles[_activeIndex] with { Macros = macros };
                Persist();
            }
        }

        /// <summary>
        /// Makros eines bestimmten Profils (nicht zwingend das im Tray "aktive") –
        /// für die Anzeige im Auswahl-Popup nach Strg+F-Taste.
        /// </summary>
        public List<MacroEntry> GetMacrosForProfile(int profileIndex)
        {
            lock (_lock)
            {
                EnsureLoaded();
                if (profileIndex < 0 || profileIndex >= _profiles!.Count) return new List<MacroEntry>();
                return _profiles[profileIndex].Macros;
            }
        }

        /// <summary>
        /// Roher (nicht expandierter) Text eines einzelnen Makros eines bestimmten
        /// Profils, oder null wenn leer/ungültig. Variablen inkl. {input} löst der
        /// Host (MainForm.SendMacroTextAsync) selbst auf.
        /// </summary>
        public string? GetRawText(int profileIndex, int macroIndex)
        {
            lock (_lock)
            {
                EnsureLoaded();
                if (profileIndex < 0 || profileIndex >= _profiles!.Count) return null;

                var macros = _profiles[profileIndex].Macros;
                if (macroIndex < 0 || macroIndex >= macros.Count) return null;

                string text = macros[macroIndex].Text;
                return string.IsNullOrWhiteSpace(text) ? null : text;
            }
        }

        // Feste deutsche Kultur für {day}: die UI/Zielgruppe ist durchgehend
        // deutsch (siehe Resources.de.resx), unabhängig von der Systemkultur
        // soll {day} deshalb immer "Montag".."Sonntag" liefern statt vom
        // zufälligen Thread-CurrentCulture des OS abzuhängen.
        private static readonly CultureInfo GermanCulture = new("de-DE");

        /// <summary>
        /// Ersetzt nicht-interaktive Variablen im Makrotext:
        ///   {date}      → aktuelles Datum (dd.MM.yyyy)
        ///   {time}      → aktuelle Uhrzeit (HH:mm)
        ///   {day}       → voller Wochentagsname, deutsch (z. B. "Montag")
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
            text = text.Replace("{day}",      now.ToString("dddd", GermanCulture),  StringComparison.OrdinalIgnoreCase);
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

        private static readonly JsonSerializerOptions ExportOptions =
            new() { WriteIndented = true, PropertyNameCaseInsensitive = true };

        /// <summary>
        /// Exportiert nur das aktive Profil als schlichte Makroliste.
        /// Dieses Format ist absichtlich identisch mit dem früherer Versionen,
        /// damit alte Exportdateien weiterhin eingelesen werden können.
        /// </summary>
        public bool ExportActiveProfile(string filePath)
        {
            try
            {
                var json = JsonSerializer.Serialize(GetMacros(), ExportOptions);
                File.WriteAllText(filePath, json, System.Text.Encoding.UTF8);
                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroManager] Export (Profil) fehlgeschlagen: {ex}");
                return false;
            }
        }

        /// <summary>Exportiert alle Profile samt Namen in eine Datei.</summary>
        public bool ExportAllProfiles(string filePath)
        {
            try
            {
                lock (_lock)
                {
                    EnsureLoaded();
                    var payload = new MacroProfileExport(ExportFormatVersion, _profiles!);
                    var json    = JsonSerializer.Serialize(payload, ExportOptions);
                    File.WriteAllText(filePath, json, System.Text.Encoding.UTF8);
                }
                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroManager] Export (alle) fehlgeschlagen: {ex}");
                return false;
            }
        }

        /// <summary>
        /// Stellt fest, was in einer Datei steckt, ohne etwas zu verändern.
        /// So kann der Aufrufer vorher fragen, wohin importiert werden soll.
        /// </summary>
        public static MacroImportKind Inspect(string filePath)
        {
            try
            {
                var json = File.ReadAllText(filePath, System.Text.Encoding.UTF8).TrimStart();
                if (json.StartsWith("[", StringComparison.Ordinal))
                    return JsonSerializer.Deserialize<List<MacroEntry>>(json, ExportOptions) != null
                        ? MacroImportKind.SingleProfile : MacroImportKind.Invalid;

                var all = JsonSerializer.Deserialize<MacroProfileExport>(json, ExportOptions);
                return all?.Profiles is { Count: > 0 }
                    ? MacroImportKind.AllProfiles : MacroImportKind.Invalid;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroManager] Inspect fehlgeschlagen: {ex}");
                return MacroImportKind.Invalid;
            }
        }

        /// <summary>Importiert eine Makroliste in das angegebene Profil.</summary>
        public bool ImportIntoProfile(string filePath, int profileIndex)
        {
            try
            {
                var json = File.ReadAllText(filePath, System.Text.Encoding.UTF8);
                var list = JsonSerializer.Deserialize<List<MacroEntry>>(json, ExportOptions);
                if (list == null) return false;

                lock (_lock)
                {
                    EnsureLoaded();
                    if (profileIndex < 0 || profileIndex >= _profiles!.Count) return false;

                    _profiles[profileIndex] = _profiles[profileIndex] with { Macros = Normalize(list) };
                    Persist();
                }
                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroManager] Import (Profil) fehlgeschlagen: {ex}");
                return false;
            }
        }

        /// <summary>Ersetzt sämtliche Profile durch die Datei.</summary>
        public bool ImportAllProfiles(string filePath)
        {
            try
            {
                var json = File.ReadAllText(filePath, System.Text.Encoding.UTF8);
                var all  = JsonSerializer.Deserialize<MacroProfileExport>(json, ExportOptions);
                if (all?.Profiles is not { Count: > 0 }) return false;

                lock (_lock)
                {
                    var imported = all.Profiles
                                      .Select(p => p with { Macros = Normalize(p.Macros) })
                                      .ToList();

                    while (imported.Count < ProfileCount)
                        imported.Add(new MacroProfile(string.Empty, EmptyMacros()));

                    _profiles = imported.Take(ProfileCount).ToList();
                    if (_activeIndex < 0 || _activeIndex >= _profiles.Count) _activeIndex = 0;
                    Persist();
                }
                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroManager] Import (alle) fehlgeschlagen: {ex}");
                return false;
            }
        }

        // -------------------------------------------------------------------
        // Private Helpers
        // -------------------------------------------------------------------

        /// <summary>Lädt Profile beim ersten Zugriff. Aufrufer hält bereits _lock.</summary>
        private void EnsureLoaded()
        {
            if (_profiles != null) return;

            _profiles = LoadProfiles();

            int saved = Properties.Settings.Default.ActiveMacroProfile;
            _activeIndex = (saved >= 0 && saved < _profiles.Count) ? saved : 0;
        }

        private static List<MacroEntry> EmptyMacros() =>
            Enumerable.Range(0, MacroCount)
                      .Select(_ => new MacroEntry(string.Empty, string.Empty))
                      .ToList();

        private static List<MacroEntry> Normalize(List<MacroEntry>? list)
        {
            list ??= new List<MacroEntry>();
            while (list.Count < MacroCount)
                list.Add(new MacroEntry(string.Empty, string.Empty));
            if (list.Count > MacroCount)
                list = list.Take(MacroCount).ToList();
            return list;
        }

        private List<MacroProfile> LoadProfiles()
        {
            // 1. Aktuelles Format: alle Profile in einem JSON.
            try
            {
                var json = Properties.Settings.Default.MacroProfilesJson;
                if (!string.IsNullOrWhiteSpace(json))
                {
                    var list = JsonSerializer.Deserialize<List<MacroProfile>>(json);
                    if (list != null && list.Count > 0)
                    {
                        var result = list.Select(p => p with { Macros = Normalize(p.Macros) }).ToList();
                        while (result.Count < ProfileCount)
                            result.Add(new MacroProfile(string.Empty, EmptyMacros()));
                        return result.Take(ProfileCount).ToList();
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroManager] LoadProfiles fehlgeschlagen: {ex}");
            }

            // 2. Migration: bisherige Makros werden zu Profil 1, der Rest bleibt leer.
            // Ein Profil fasst weiterhin MacroCount (= 10) Makros, also passt der
            // gesamte Alt-Satz verlustfrei hinein.
            var profiles = new List<MacroProfile>();
            var existing = Normalize(LoadLegacyMacros());

            profiles.Add(new MacroProfile(string.Empty, existing));
            while (profiles.Count < ProfileCount)
                profiles.Add(new MacroProfile(string.Empty, EmptyMacros()));

            PersistProfiles(profiles);
            return profiles;
        }

        /// <summary>
        /// Liest die Makros aus den beiden Vorgängerformaten (vor den Profilen).
        /// Gibt die Rohliste zurück; der Aufrufer normalisiert sie auf MacroCount.
        /// </summary>
        private static List<MacroEntry> LoadLegacyMacros()
        {
            // a) MacrosJson (eine einzelne Liste)
            try
            {
                var json = Properties.Settings.Default.MacrosJson;
                if (!string.IsNullOrWhiteSpace(json))
                {
                    var list = JsonSerializer.Deserialize<List<MacroEntry>>(json);
                    if (list != null)
                    {
                        Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroManager] Übernehme {list.Count} bisherige Makros und verteile sie auf die Profile.");
                        return list;
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroManager] Legacy-JSON fehlgeschlagen: {ex}");
            }

            // b) Noch älter: StringCollection
            try
            {
                var legacy = Properties.Settings.Default.Macros;
                if (legacy != null && legacy.Count > 0)
                {
                    Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroManager] Migriere {legacy.Count} Makros aus altem Format.");

                    var migrated = new List<MacroEntry>();
                    foreach (var item in legacy)
                        migrated.Add(new MacroEntry(string.Empty, item ?? string.Empty));

                    Properties.Settings.Default.Macros = new System.Collections.Specialized.StringCollection();
                    Properties.Settings.Default.Save();

                    return migrated;
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroManager] Migration fehlgeschlagen: {ex}");
            }

            // c) Erster Start: nichts zu übernehmen.
            return new List<MacroEntry>();
        }

        /// <summary>Persistiert den aktuellen Profilstand. Aufrufer hält bereits _lock.</summary>
        private void Persist()
        {
            if (_profiles != null) PersistProfiles(_profiles);
        }

        private static void PersistProfiles(List<MacroProfile> profiles)
        {
            try
            {
                Properties.Settings.Default.MacroProfilesJson = JsonSerializer.Serialize(profiles);
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroManager] PersistProfiles fehlgeschlagen: {ex}");
            }
        }
    }
}
