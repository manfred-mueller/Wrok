using System.Configuration;
using System.Diagnostics;
using System.Text.Json;

namespace Wrok
{
    /// <summary>
    /// Ersetzt den Standard-Settings-Provider (LocalFileSettingsProvider), der
    /// Benutzereinstellungen in einem Ordner ablegt, dessen Name sowohl von der
    /// AssemblyVersion als auch von einem Hash des Pfads der ausführenden .exe
    /// abhängt ("Wrok_Url_&lt;hash&gt;"). Bei Wrok als Single-File-Publish
    /// (PublishSingleFile) liefert .NET für die Assembly keinen stabilen Pfad mehr,
    /// wodurch praktisch bei jedem Start ein neuer, leerer Einstellungsordner
    /// entsteht - Makros, Fenstergröße usw. gingen dadurch verloren, nicht nur bei
    /// Versionswechseln, sondern quasi bei jedem Neustart.
    ///
    /// Dieser Provider speichert stattdessen alle Einstellungen in einer einzigen,
    /// festen Datei unter %LocalAppData%\Wrok\settings.json - unabhängig von
    /// AssemblyVersion und Startpfad. Registriert über
    /// [SettingsProvider(typeof(WrokSettingsProvider))] in Settings.Provider.cs.
    ///
    /// Die eigentliche Wert-Serialisierung (String/Xml je nach Eigenschaftstyp,
    /// z. B. für die alte StringCollection-Einstellung "Macros") überlässt diese
    /// Klasse bewusst SettingsPropertyValue.SerializedValue - genau das, was auch
    /// LocalFileSettingsProvider intern nutzt. Hier wird nur noch entschieden, WO
    /// diese bereits fertig serialisierten Werte landen, nicht WIE sie entstehen.
    /// </summary>
    internal sealed class WrokSettingsProvider : SettingsProvider
    {
        private static readonly string StoreDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Wrok");

        private static readonly string StoreFile = Path.Combine(StoreDirectory, "settings.json");

        private static readonly object FileLock = new();

        public override string ApplicationName { get; set; } = "Wrok";

        public override string Name => nameof(WrokSettingsProvider);

        public override void Initialize(string? name, System.Collections.Specialized.NameValueCollection? config)
        {
            base.Initialize(string.IsNullOrEmpty(name) ? Name : name,
                config ?? new System.Collections.Specialized.NameValueCollection());
        }

        public override SettingsPropertyValueCollection GetPropertyValues(
            SettingsContext context, SettingsPropertyCollection properties)
        {
            var stored = Load();
            var result = new SettingsPropertyValueCollection();

            foreach (SettingsProperty prop in properties)
            {
                var value = new SettingsPropertyValue(prop) { IsDirty = false };

                // Kein gespeicherter Eintrag -> SerializedValue bleibt unangetastet,
                // SettingsPropertyValue liefert dann beim ersten Lesen automatisch
                // Property.DefaultValue (aus Settings.settings) zurück.
                if (stored.TryGetValue(prop.Name, out var serialized) && serialized != null)
                    value.SerializedValue = serialized;

                result.Add(value);
            }

            return result;
        }

        public override void SetPropertyValues(SettingsContext context, SettingsPropertyValueCollection values)
        {
            var stored = Load();

            foreach (SettingsPropertyValue value in values)
                stored[value.Name] = value.SerializedValue as string ?? string.Empty;

            Save(stored);
        }

        private static Dictionary<string, string> Load()
        {
            lock (FileLock)
            {
                try
                {
                    if (File.Exists(StoreFile))
                    {
                        var json = File.ReadAllText(StoreFile, System.Text.Encoding.UTF8);
                        var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                        if (dict != null) return dict;
                    }
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [WrokSettingsProvider] Laden fehlgeschlagen: {ex}");
                }
                return new Dictionary<string, string>();
            }
        }

        private static void Save(Dictionary<string, string> store)
        {
            lock (FileLock)
            {
                try
                {
                    Directory.CreateDirectory(StoreDirectory);
                    var json = JsonSerializer.Serialize(store, new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(StoreFile, json, System.Text.Encoding.UTF8);
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [WrokSettingsProvider] Speichern fehlgeschlagen: {ex}");
                }
            }
        }
    }
}
