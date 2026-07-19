using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Wrok
{
    // -------------------------------------------------------------------------
    // Datenmodell
    // -------------------------------------------------------------------------

    /// <summary>Wie knapp das Kontingent ist – steuert die Farbe im Tray-Menue.</summary>
    internal enum RateLimitSeverity
    {
        Unknown,
        Normal,
        /// <summary>Hoechstens 25 % uebrig.</summary>
        Warning,
        /// <summary>Hoechstens 10 % uebrig.</summary>
        Critical
    }

    internal record RateLimitEntry(
        [property: JsonPropertyName("remainingQueries")]  int?  RemainingQueries,
        [property: JsonPropertyName("totalQueries")]      int?  TotalQueries,
        [property: JsonPropertyName("windowSizeSeconds")] long? WindowSizeSeconds,
        [property: JsonPropertyName("error")]             string? Error
    )
    {
        public bool   IsError    => !string.IsNullOrWhiteSpace(Error);
        public bool   IsUnknown  => !IsError && RemainingQueries == null;

        /// <summary>
        /// Restkontingent als Anteil (0.0 bis 1.0) oder null, wenn unbekannt.
        /// </summary>
        public double? RemainingFraction =>
            (RemainingQueries is int r && TotalQueries is int t && t > 0)
                ? r / (double)t
                : null;

        /// <summary>
        /// Dringlichkeitsstufe fuer die farbliche Hervorhebung im Menue.
        /// Eine Farbe erfasst man im Vorbeigehen, eine Zahl muss man lesen.
        /// </summary>
        public RateLimitSeverity Severity => RemainingFraction switch
        {
            null      => RateLimitSeverity.Unknown,
            <= 0.10   => RateLimitSeverity.Critical,
            <= 0.25   => RateLimitSeverity.Warning,
            _         => RateLimitSeverity.Normal
        };
        public string ResetInfo
        {
            get
            {
                if (WindowSizeSeconds == null) return string.Empty;
                var ts = TimeSpan.FromSeconds(WindowSizeSeconds.Value);
                return ts.TotalHours >= 1
                    ? string.Format(Properties.Resources.RateLimitResetsIn0h1m, (int)ts.TotalHours, ts.Minutes)
                    : string.Format(Properties.Resources.RateLimitResetsIn0m,   (int)ts.TotalMinutes);
            }
        }
    }

    internal record RateLimitResult(
        RateLimitEntry? Grok3,
        RateLimitEntry? Grok4Heavy,
        DateTime        FetchedAt,
        string?         FetchError = null
    );

    // -------------------------------------------------------------------------
    // RateLimitManager
    // -------------------------------------------------------------------------

    internal sealed class RateLimitManager
    {
        // Automatischer Refresh-Intervall
        private static readonly TimeSpan AutoRefreshInterval = TimeSpan.FromMinutes(5);

        private readonly WebViewManager              _webViewManager;
        private readonly Action                      _onUpdated;
        private System.Windows.Forms.Timer?          _refreshTimer;
        private RateLimitResult?                     _lastResult;
        private bool                                 _fetching;
        private readonly object                      _lock = new();

        public RateLimitResult? LastResult
        {
            get { lock (_lock) { return _lastResult; } }
        }

        public RateLimitManager(WebViewManager webViewManager, Action onUpdated)
        {
            _webViewManager = webViewManager;
            _onUpdated      = onUpdated;
        }

        // ------------------------------------------------------------------
        // Öffentliche API
        // ------------------------------------------------------------------

        public void StartAutoRefresh()
        {
            _refreshTimer = new System.Windows.Forms.Timer
            {
                Interval = (int)AutoRefreshInterval.TotalMilliseconds
            };
            _refreshTimer.Tick += async (s, e) => await RefreshAsync();
            _refreshTimer.Start();
        }

        public void StopAutoRefresh()
        {
            _refreshTimer?.Stop();
            _refreshTimer?.Dispose();
            _refreshTimer = null;
        }

        public async Task RefreshAsync()
        {
            lock (_lock)
            {
                if (_fetching) return;
                _fetching = true;
            }

            try
            {
                var result = await FetchRateLimitsAsync();
                lock (_lock)
                {
                    _lastResult = result;
                }
                _onUpdated();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [RateLimitManager] RefreshAsync failed: {ex}");
            }
            finally
            {
                lock (_lock) { _fetching = false; }
            }
        }

        // ------------------------------------------------------------------
        // Fetch via WebView JS (nutzt die eingeloggte Session)
        // ------------------------------------------------------------------

        // Abgefragte Modelle. Anpassen, wenn xAI die Modellnamen ändert.
        private const string ModelGrok3     = "grok-3";
        private const string ModelGrok4Heavy = "grok-4-heavy";

        private async Task<RateLimitResult> FetchRateLimitsAsync()
        {
            // Über die Async-Brücke, NICHT über ExecuteScriptAsync: dessen Rückgabewert
            // wäre bei einem Promise nur "{}" – die Abfrage lieferte deshalb früher nie Daten.
            string? raw;
            try
            {
                raw = await _webViewManager.FetchRateLimitsJsonAsync(
                    new[] { ModelGrok3, ModelGrok4Heavy },
                    TimeSpan.FromSeconds(15));
            }
            catch (Exception ex)
            {
                return new RateLimitResult(null, null, DateTime.Now, ex.Message);
            }

            if (string.IsNullOrWhiteSpace(raw))
                return new RateLimitResult(null, null, DateTime.Now, "Empty response");

            try
            {
                // Die Brücke liefert bereits rohes JSON – kein doppeltes Unquoting nötig.
                using var doc = JsonDocument.Parse(raw);
                var root = doc.RootElement;

                var grok3      = ParseEntry(root, ModelGrok3);
                var grok4heavy = ParseEntry(root, ModelGrok4Heavy);

                return new RateLimitResult(grok3, grok4heavy, DateTime.Now);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [RateLimitManager] Parse failed: {ex.Message}, raw={raw}");
                return new RateLimitResult(null, null, DateTime.Now, $"Parse error: {ex.Message}");
            }
        }

        private static RateLimitEntry? ParseEntry(JsonElement root, string key)
        {
            try
            {
                if (!root.TryGetProperty(key, out var el)) return null;
                return JsonSerializer.Deserialize<RateLimitEntry>(el.GetRawText());
            }
            catch { return null; }
        }
    }
}
