using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text.Json;

namespace Wrok
{
    /// <summary>Ergebnis einer erfolgreichen Update-Prüfung.</summary>
    internal sealed record UpdateInfo(
        Version LatestVersion, string TagName, string ReleaseUrl, string? InstallerUrl);

    /// <summary>Ausgang einer manuellen Prüfung (für den About-Dialog).</summary>
    internal enum UpdateOutcome { UpToDate, UpdateAvailable, Failed }

    /// <summary>Ergebnis einer manuellen Prüfung: Ausgang plus ggf. die Update-Info.</summary>
    internal sealed record UpdateCheckResult(UpdateOutcome Outcome, UpdateInfo? Info);

    /// <summary>
    /// Prüft beim Start, ob auf GitHub eine neuere Version vorliegt.
    ///
    /// Bewusst nur ein Hinweis, kein Selbst-Updater: Der eigentliche Download läuft
    /// weiter über WinGet bzw. den signierten Installer. Die Prüfung ist ein
    /// einzelner, lesender API-Aufruf – sie lädt nichts herunter und führt nichts aus.
    ///
    /// Datenschutz: Der Aufruf verrät GitHub die IP-Adresse des Nutzers. Deshalb
    /// abschaltbar (Einstellung <c>CheckForUpdates</c>) und auf einmal täglich
    /// gedrosselt (<c>LastUpdateCheckUtc</c>).
    /// </summary>
    internal static class UpdateChecker
    {
        private const string LatestReleaseApi =
            "https://api.github.com/repos/manfred-mueller/Wrok/releases/latest";

        // GitHub verlangt einen User-Agent; ohne wird der Aufruf mit 403 abgewiesen.
        private static readonly HttpClient _http = CreateClient();

        // Nicht öfter als einmal in diesem Abstand tatsächlich anfragen.
        private static readonly TimeSpan MinInterval = TimeSpan.FromHours(20);

        private static HttpClient CreateClient()
        {
            var client = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
            client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("Wrok-UpdateCheck", CurrentVersion()?.ToString() ?? "1.0"));
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
            return client;
        }

        /// <summary>Aktuelle Programmversion (Major.Minor.Build), oder null.</summary>
        public static Version? CurrentVersion()
        {
            var v = Assembly.GetExecutingAssembly().GetName().Version;
            return v == null ? null : new Version(v.Major, v.Minor, v.Build);
        }

        /// <summary>
        /// Führt die Prüfung aus, sofern eingeschaltet und nicht kürzlich schon
        /// geschehen. Gibt die Info nur zurück, wenn tatsächlich eine neuere Version
        /// vorliegt; in allen anderen Fällen (aus, gedrosselt, aktuell, Fehler) null.
        /// </summary>
        public static async Task<UpdateInfo?> CheckAsync()
        {
            try
            {
                if (!Properties.Settings.Default.CheckForUpdates) return null;
                if (!DueForCheck()) return null;

                // Zeitstempel VOR dem Aufruf setzen: Auch wenn die Anfrage scheitert,
                // soll nicht bei jedem Start erneut gegen GitHub gelaufen werden.
                StampChecked();

                var (ok, info) = await FetchLatestAsync();
                return ok ? info : null;   // "aktuell" (info == null) hier wie "kein Hinweis"
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [UpdateChecker] Prüfung fehlgeschlagen: {ex}");
                return null;
            }
        }

        /// <summary>
        /// Manuelle Prüfung für den About-Dialog: läuft IMMER (ohne Rücksicht auf den
        /// Schalter oder die Tagesdrossel) und unterscheidet die drei Ausgänge, damit
        /// der Nutzer auch ein „schon aktuell" oder einen Fehler gemeldet bekommt.
        /// </summary>
        public static async Task<UpdateCheckResult> CheckNowAsync()
        {
            try
            {
                StampChecked();
                var (ok, info) = await FetchLatestAsync();
                if (!ok) return new UpdateCheckResult(UpdateOutcome.Failed, null);
                return info != null
                    ? new UpdateCheckResult(UpdateOutcome.UpdateAvailable, info)
                    : new UpdateCheckResult(UpdateOutcome.UpToDate, null);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [UpdateChecker] Manuelle Prüfung fehlgeschlagen: {ex}");
                return new UpdateCheckResult(UpdateOutcome.Failed, null);
            }
        }

        /// <summary>
        /// Holt das neueste Release und vergleicht mit der laufenden Version.
        /// Rückgabe: <c>ok</c> = Abfrage erfolgreich; <c>info</c> = neuere Version
        /// (oder null, wenn bereits aktuell). Bei <c>ok == false</c> ist etwas
        /// schiefgegangen (Netz, HTTP, Parsen).
        /// </summary>
        private static async Task<(bool ok, UpdateInfo? info)> FetchLatestAsync()
        {
            var current = CurrentVersion();
            if (current == null) return (false, null);

            using var resp = await _http.GetAsync(LatestReleaseApi);
            if (!resp.IsSuccessStatusCode)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [UpdateChecker] HTTP {(int)resp.StatusCode} von GitHub.");
                return (false, null);
            }

            var json = await resp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            // Vorabversionen zählen nicht als Update.
            if (root.TryGetProperty("prerelease", out var pre) && pre.ValueKind == JsonValueKind.True)
                return (true, null);

            string tag = root.TryGetProperty("tag_name", out var t) ? (t.GetString() ?? string.Empty) : string.Empty;
            string url = root.TryGetProperty("html_url", out var u) ? (u.GetString() ?? string.Empty) : string.Empty;

            if (!TryParseTag(tag, out var latest)) return (false, null);
            if (latest <= current) return (true, null);

            string? installerUrl = FindInstallerAssetUrl(root);
            return (true, new UpdateInfo(latest, tag, url, installerUrl));
        }

        /// <summary>
        /// Sucht im Release-JSON das Setup-Asset (Name wie "Wrok-Setup-1.6.0.exe",
        /// siehe OutputBaseFilename in InstallScript.iss) und liefert dessen direkten
        /// Download-Link. Null, wenn kein passendes Asset im Release liegt (z. B. bei
        /// einem Release ohne angehängte Binärdatei) - der Aufrufer fällt dann auf das
        /// bisherige Verhalten (Release-Seite im Browser öffnen) zurück.
        /// </summary>
        private static string? FindInstallerAssetUrl(JsonElement root)
        {
            if (!root.TryGetProperty("assets", out var assets) || assets.ValueKind != JsonValueKind.Array)
                return null;

            foreach (var asset in assets.EnumerateArray())
            {
                var name = asset.TryGetProperty("name", out var n) ? n.GetString() : null;
                if (name == null) continue;
                if (!name.StartsWith("Wrok-Setup-", StringComparison.OrdinalIgnoreCase)) continue;
                if (!name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)) continue;

                return asset.TryGetProperty("browser_download_url", out var dl) ? dl.GetString() : null;
            }

            return null;
        }

        private static void StampChecked()
        {
            try
            {
                Properties.Settings.Default.LastUpdateCheckUtc = DateTime.UtcNow.ToString("o");
                Properties.Settings.Default.Save();
            }
            catch { /* nicht kritisch */ }
        }

        private static bool DueForCheck()
        {
            var last = Properties.Settings.Default.LastUpdateCheckUtc;
            if (string.IsNullOrWhiteSpace(last)) return true;

            // Nicht parsebar → als „lange her" behandeln und prüfen.
            if (!DateTime.TryParse(last, null, System.Globalization.DateTimeStyles.RoundtripKind, out var when))
                return true;

            return DateTime.UtcNow - when.ToUniversalTime() >= MinInterval;
        }

        /// <summary>Wandelt ein Release-Tag wie „v1.3.1" oder „1.3.1" in eine Version.</summary>
        private static bool TryParseTag(string? tag, out Version version)
        {
            version = new Version(0, 0);
            if (string.IsNullOrWhiteSpace(tag)) return false;

            string cleaned = tag.Trim().TrimStart('v', 'V');
            if (Version.TryParse(cleaned, out var parsed) && parsed != null)
            {
                // Auf Major.Minor.Build normalisieren, damit der Vergleich nicht an
                // einer Revision hängenbleibt (ein 4-teiliges Tag hätte Revision 0,
                // CurrentVersion dagegen -1 – und läge damit scheinbar zurück).
                version = new Version(parsed.Major, parsed.Minor, parsed.Build < 0 ? 0 : parsed.Build);
                return true;
            }
            return false;
        }
    }
}
