using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Wrok
{
    /// <summary>
    /// Kapselt die gesamte WebView2-Interaktion:
    ///   â¢ Initialisierung inkl. User-Data-Pfad und User-Agent-Spoofing
    ///   â¢ Laden des JS-Helpers aus eingebetteten Ressourcen
    ///   â¢ Senden von Text (mit optionalem Enter) Ã¼ber JS + NativeInput-Fallback
    ///   â¢ Anzeige der Offline-Seite
    ///   â¢ Internet-KonnektivitÃ¤tsprÃ¼fung
    /// </summary>
    internal sealed class WebViewManager
    {
        // ------------------------------------------------------------------
        // Timing-Konstanten fÃ¼r Focus/UI-Settle
        // ------------------------------------------------------------------

        private const int FocusSettleDelayMs      = 120;
        private const int ButtonEnableDelayMs      = 140;
        private const int NativeInputFocusDelayMs  = 80;
        private const int NativeInputTextDelayMs   = 150;
        private const int NativeInputAfterTextDelayMs = 40;

        // Versionsnummer des JS-Helpers â erhÃ¶hen, sobald __wrokSend oder
        // __wrokEnsureFocus geÃ¤ndert wird.
        private const int WrokHelperVersion = 7;

        // ------------------------------------------------------------------
        // Felder
        // ------------------------------------------------------------------

        private readonly WebView2 _webView;
        private readonly MainForm _owner;
        private readonly Action _onActivityReset;
        private NativeInput? _input;

        // Shared HttpClient â einmal pro Prozess.
        private static readonly HttpClient _httpClient = new();

        /// <summary>
        /// Die WebView2-Umgebung. Wird vom Video-Fenster wiederverwendet, damit es
        /// denselben User-Data-Ordner – und damit dieselbe Login-Session – nutzt.
        /// </summary>
        public CoreWebView2Environment? CoreEnvironment { get; private set; }

        // Offene Anfragen an die JS-Async-Brücke (Token -> Ergebnis).
        private readonly System.Collections.Concurrent.ConcurrentDictionary<string, TaskCompletionSource<string?>>
            _pendingResults = new();

        // ------------------------------------------------------------------
        // Konstruktor
        // ------------------------------------------------------------------

        /// <param name="webView">Das bereits dem Form hinzugefÃ¼gte WebView2-Control.</param>
        /// <param name="owner">EigentÃ¼mer-Form (fÃ¼r BeginInvoke und SetForegroundWindow).</param>
        /// <param name="onActivityReset">Callback, der bei NutzeraktivitÃ¤t im WebView aufgerufen wird.</param>
        public WebViewManager(WebView2 webView, MainForm owner, Action onActivityReset)
        {
            _webView = webView;
            _owner   = owner;
            _onActivityReset = onActivityReset;
        }

        // ------------------------------------------------------------------
        // Initialisierung
        // ------------------------------------------------------------------

        /// <summary>
        /// Erstellt die CoreWebView2-Umgebung, richtet User-Agent und Header ein
        /// und injiziert den JS-Helper in alle kÃ¼nftigen Dokumente.
        /// </summary>
        public async Task InitializeAsync()
        {
            string userDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Wrok", "WebView2Data");

            try { Directory.CreateDirectory(userDataPath); }
            catch { /* nicht kritisch */ }

            _webView.CoreWebView2InitializationCompleted += OnCoreWebView2InitializationCompleted;

            try
            {
                var env = await CoreWebView2Environment.CreateAsync(userDataFolder: userDataPath);
                CoreEnvironment = env;   // für weitere WebViews (Video-Fenster) wiederverwenden
                await _webView.EnsureCoreWebView2Async(env);

                if (_webView.CoreWebView2 != null)
                    ConfigureCore(_webView.CoreWebView2);
            }
            catch (Exception ex)
            {
                Log(ex, "InitializeAsync: CoreWebView2-Umgebung konnte nicht erstellt werden");
            }
        }

        /// <summary>
        /// Erzeugt den NativeInput-Simulator (muss nach OnHandleCreated aufgerufen werden).
        /// </summary>
        public void CreateInputSimulator() => _input = new NativeInput();

        /// <summary>
        /// Gibt den NativeInput-Simulator frei (bei OnHandleDestroyed aufrufen).
        /// </summary>
        public void DisposeInputSimulator() => _input = null;

        /// <summary>
        /// FÃ¼hrt ein JavaScript-Script im WebView aus und gibt das Ergebnis zurÃ¼ck.
        /// Wirft eine Exception wenn CoreWebView2 nicht initialisiert ist.
        /// </summary>
        public async Task<string?> ExecuteScriptAsync(string script)
        {
            if (_webView.CoreWebView2 == null)
                throw new InvalidOperationException("CoreWebView2 not initialized.");
            return await _webView.CoreWebView2.ExecuteScriptAsync(script).ConfigureAwait(false);
        }

        // ------------------------------------------------------------------
        // Navigation
        // ------------------------------------------------------------------

        /// <summary>
        /// Navigiert zu <paramref name="url"/>, wenn eine Internetverbindung besteht;
        /// andernfalls wird die Offline-Seite angezeigt.
        /// </summary>
        public async Task NavigateAsync(string url)
        {
            if (_webView.CoreWebView2 == null)
            {
                try { await _webView.EnsureCoreWebView2Async(null); }
                catch { }
            }

            var core = _webView.CoreWebView2;
            if (core == null) { await ShowOfflinePageAsync(); return; }

            bool online = await HasInternetConnectionAsync(attempts: 3, timeoutSeconds: 5);
            if (online)
            {
                try { core.Navigate(url); }
                catch (Exception ex) { Log(ex, "Navigate fehlgeschlagen"); await ShowOfflinePageAsync(); }
            }
            else
            {
                await ShowOfflinePageAsync();
            }
        }

        // ------------------------------------------------------------------
        // Text senden
        // ------------------------------------------------------------------

        /// <summary>
        /// Sendet <paramref name="text"/> in das aktive Eingabefeld des WebViews.
        /// Falls <paramref name="pressEnter"/> true ist, wird anschlieÃend Enter ausgelÃ¶st.
        /// </summary>
        public async Task SendTextAsync(string text, bool pressEnter = false)
        {
            if (_webView.CoreWebView2 == null)
                return;

            try { await EnsureFocusOnUiThreadAsync(); }
            catch { }

            await Task.Delay(FocusSettleDelayMs).ConfigureAwait(false);

            var payload    = System.Text.Json.JsonSerializer.Serialize(text);
            var enterArg   = pressEnter ? "true" : "false";
            var callScript = $"(function(){{ try {{ if (window.__wrokEnsureFocus) window.__wrokEnsureFocus(); " +
                             $"return window.__wrokSend ? window.__wrokSend({payload}, {enterArg}) : false; }} " +
                             $"catch(e) {{ return false; }} }})();";

            string? rawResult = null;
            try
            {
                rawResult = await _webView.CoreWebView2.ExecuteScriptAsync(callScript).ConfigureAwait(false);
                Trace.WriteLine($"SendTextAsync JS result: {rawResult}");
            }
            catch (Exception ex)
            {
                Trace.WriteLine(string.Format(Properties.Resources.ParsingClickScriptResultFailed0, ex));
            }

            bool jsSucceeded = ParseJsBoolResult(rawResult);

            if (pressEnter)
            {
                await Task.Delay(ButtonEnableDelayMs).ConfigureAwait(false);

                bool clickSucceeded = await TryClickSendButtonAsync();
                if (clickSucceeded) return;

                if (jsSucceeded)
                {
                    await SendEnterViaNativeInputAsync();
                    return;
                }
            }
            else if (jsSucceeded)
            {
                return;
            }

            // VollstÃ¤ndiger Fallback Ã¼ber NativeInput.
            await FallbackSendViaNativeInputAsync(text, pressEnter);
        }

        // ------------------------------------------------------------------
        // KonnektivitÃ¤t
        // ------------------------------------------------------------------

        public async Task<bool> HasInternetConnectionAsync(int attempts = 2, int timeoutSeconds = 4)
        {
            try
            {
                if (!System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable())
                    return false;
            }
            catch { /* SchnellprÃ¼fung nicht kritisch */ }

            var urls = new[]
            {
                "https://clients3.google.com/generate_204",
                "http://detectportal.firefox.com/success.txt",
                "https://www.bing.com/"
            };

            for (int attempt = 0; attempt < Math.Max(1, attempts); attempt++)
            {
                foreach (var u in urls)
                {
                    try
                    {
                        using var req = new HttpRequestMessage(HttpMethod.Head, u);
                        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(timeoutSeconds));
                        try
                        {
                            using var resp = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, cts.Token);
                            if (resp.IsSuccessStatusCode || resp.StatusCode == System.Net.HttpStatusCode.NoContent)
                                return true;
                        }
                        catch
                        {
                            using var req2  = new HttpRequestMessage(HttpMethod.Get, u);
                            using var resp2 = await _httpClient.SendAsync(req2, HttpCompletionOption.ResponseHeadersRead, cts.Token);
                            if (resp2.IsSuccessStatusCode || resp2.StatusCode == System.Net.HttpStatusCode.NoContent)
                                return true;
                        }
                    }
                    catch { /* nÃ¤chste URL probieren */ }
                }

                if (attempt + 1 < attempts)
                    await Task.Delay(300 + attempt * 200);
            }
            return false;
        }

        // ------------------------------------------------------------------
        // Private Helpers
        // ------------------------------------------------------------------

        private async void OnCoreWebView2InitializationCompleted(object? sender, CoreWebView2InitializationCompletedEventArgs e)
        {
            if (_webView.CoreWebView2 == null) return;

            string helperScript = BuildHelperScript();

            try { await _webView.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync(helperScript); }
            catch (Exception ex) { Log(ex, "AddScriptToExecuteOnDocumentCreated fehlgeschlagen"); }

            try { await _webView.CoreWebView2.ExecuteScriptAsync(helperScript); }
            catch (Exception ex) { Log(ex, "Sofortiges Injizieren des Helpers fehlgeschlagen"); }

            _webView.CoreWebView2.WebMessageReceived += (s, args) =>
            {
                var msg = args.TryGetWebMessageAsString();
                if (msg == "resetActivity")
                    _onActivityReset();
                else if (msg != null && msg.StartsWith("openImage:", StringComparison.Ordinal))
                {
                    var imgUrl = msg["openImage:".Length..];
                    Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [WebViewManager] openImage empfangen: {imgUrl}");
                    _ = ShowImageViewerAsync(imgUrl);
                }
                else if (msg != null && msg.StartsWith("openVideo:", StringComparison.Ordinal))
                {
                    var vidUrl = msg["openVideo:".Length..];
                    Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [WebViewManager] openVideo empfangen: {vidUrl}");
                    _owner.ShowVideoViewer(vidUrl);
                }
                else if (msg != null && msg.StartsWith(ResultPrefix, StringComparison.Ordinal))
                    HandleBridgeResult(msg[ResultPrefix.Length..]);
            };
        }

        // ------------------------------------------------------------------
        // Async-Brücke zum JS
        //
        // CoreWebView2.ExecuteScriptAsync löst KEINE Promises auf (liefert "{}"),
        // deshalb liefert das JS asynchrone Ergebnisse per postMessage zurück.
        // ------------------------------------------------------------------

        private const string ResultPrefix = "wrokResult:";

        /// <summary>Ordnet eine eingehende Brücken-Antwort ihrem wartenden Task zu.</summary>
        private void HandleBridgeResult(string rest)
        {
            int sep = rest.IndexOf(':');
            if (sep <= 0) return;

            string token   = rest[..sep];
            string payload = rest[(sep + 1)..];

            if (_pendingResults.TryRemove(token, out var tcs))
                tcs.TrySetResult(payload.Length == 0 ? null : payload);
        }

        /// <summary>
        /// Ruft eine JS-Funktion auf, die ihr Ergebnis per postMessage zurückschickt,
        /// und wartet darauf. Gibt null zurück bei Fehler oder Zeitüberschreitung.
        /// </summary>
        private async Task<string?> CallBridgeAsync(Func<string, string> buildScript, TimeSpan timeout)
        {
            string token = Guid.NewGuid().ToString("N");
            var tcs = new TaskCompletionSource<string?>(TaskCreationOptions.RunContinuationsAsynchronously);
            _pendingResults[token] = tcs;

            try
            {
                await ExecuteOnUiThreadAsync(buildScript(token)).ConfigureAwait(false);

                using var cts = new CancellationTokenSource(timeout);
                using var reg = cts.Token.Register(() => tcs.TrySetResult(null));
                return await tcs.Task.ConfigureAwait(false);
            }
            finally
            {
                _pendingResults.TryRemove(token, out _);
            }
        }

        /// <summary>
        /// Führt ein Skript auf dem UI-Thread aus (CoreWebView2 ist threadgebunden).
        /// Der Rückgabewert des Skripts wird bewusst ignoriert – das Ergebnis kommt
        /// über die Brücke.
        /// </summary>
        private Task ExecuteOnUiThreadAsync(string script)
        {
            var done = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

            void Run()
            {
                try
                {
                    if (_webView.CoreWebView2 != null)
                        _ = _webView.CoreWebView2.ExecuteScriptAsync(script);
                }
                catch (Exception ex) { Log(ex, "ExecuteOnUiThreadAsync fehlgeschlagen"); }
                finally { done.TrySetResult(true); }
            }

            try
            {
                if (_owner.IsDisposed || !_owner.IsHandleCreated) { done.TrySetResult(true); }
                else if (_owner.InvokeRequired) _owner.BeginInvoke((System.Windows.Forms.MethodInvoker)Run);
                else Run();
            }
            catch (Exception ex)
            {
                Log(ex, "ExecuteOnUiThreadAsync: Marshalling fehlgeschlagen");
                done.TrySetResult(true);
            }

            return done.Task;
        }

        /// <summary>
        /// Fragt die Rate-Limits der angegebenen Modelle ab und gibt das rohe JSON zurück
        /// (Objekt mit Modellnamen als Schlüssel), oder null bei Fehler/Timeout.
        /// </summary>
        /// <summary>
        /// Ermittelt den Content-Type einer URL über die eingeloggte Session.
        /// Nötig, weil Grok-URLs keine Dateiendung haben und man ihnen nicht
        /// ansieht, ob dahinter ein Bild oder ein Video steckt.
        /// </summary>
        public Task<string?> ProbeContentTypeAsync(string url, TimeSpan timeout) =>
            CallBridgeAsync(
                token => $"window.__wrokProbeType(" +
                         $"{System.Text.Json.JsonSerializer.Serialize(url)}, " +
                         $"{System.Text.Json.JsonSerializer.Serialize(token)});",
                timeout);

        public Task<string?> FetchRateLimitsJsonAsync(IEnumerable<string> models, TimeSpan timeout)
        {
            var json = System.Text.Json.JsonSerializer.Serialize(models);
            return CallBridgeAsync(
                token => $"window.__wrokRateLimits({System.Text.Json.JsonSerializer.Serialize(token)}, {json});",
                timeout);
        }

        private static void ConfigureCore(CoreWebView2 core)
        {
            core.Settings.UserAgent =
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " +
                "(KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36 Edg/146.0.0.0";

            // NUR Dokument-Anfragen, nicht "All".
            //
            // Zuvor wurden diese Header auf JEDE Anfrage gesetzt – auch auf Bilder,
            // API-Aufrufe und Datei-Uploads. Einem Upload-POST wurde damit
            // "Sec-Fetch-Dest: document" und "Accept: text/html" verpasst, was für
            // eine XHR-Anfrage schlicht falsch ist und serverseitig zu unerwartetem
            // Verhalten führen kann. Nebenbei kostete der Filter bei jeder einzelnen
            // Anfrage einen Sprung in verwalteten Code.
            //
            // Für eine echte Seitennavigation sind die Werte dagegen korrekt –
            // Chromium sendet sie für alle anderen Anfragetypen ohnehin selbst
            // und jeweils passend.
            core.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.Document);
            core.WebResourceRequested += (sender, args) =>
            {
                var h = args.Request.Headers;
                h.SetHeader("Accept",          "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7");
                h.SetHeader("Accept-Language", "de-DE,de;q=0.9,en-US;q=0.8,en;q=0.7");
                h.SetHeader("Sec-Fetch-Dest",  "document");
                h.SetHeader("Sec-Fetch-Mode",  "navigate");
                // "Accept-Encoding" nicht setzen: Das handelt Chromium selbst aus und
                // kennt die tatsaechlich unterstuetzten Verfahren besser.
                // "Sec-Fetch-Site" ebenfalls nicht: Der Wert haengt davon ab, woher
                // die Navigation kommt - fest "same-origin" waere beim Erstaufruf falsch.
            };
        }

        /// <summary>
        /// LÃ¤dt WrokHelper.js aus den eingebetteten Ressourcen und setzt die Versionsnummer ein.
        /// </summary>
        private static string BuildHelperScript()
        {
            string template = LoadEmbeddedJs("WrokHelper.js");
            return template.Replace("__WROK_VERSION__", WrokHelperVersion.ToString());
        }

        /// <summary>Liest eine eingebettete JS-Ressource aus der Assembly.</summary>
        private static string LoadEmbeddedJs(string filename)
        {
            var asm  = Assembly.GetExecutingAssembly();
            // Ressourcenname: <DefaultNamespace>.<Dateiname>
            string resourceName = $"Wrok.{filename}";
            using var stream = asm.GetManifestResourceStream(resourceName)
                               ?? throw new InvalidOperationException($"Embedded resource '{resourceName}' not found.");
            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }

        private Task EnsureFocusOnUiThreadAsync()
        {
            if (_owner.IsDisposed || !_owner.IsHandleCreated)
                return Task.CompletedTask;

            var tcs = new TaskCompletionSource<bool>();
            try
            {
                _owner.BeginInvoke((System.Windows.Forms.MethodInvoker)(() =>
                {
                    try
                    {
                        _webView.Focus();
                        SetForegroundWindow(_owner.Handle);
                    }
                    catch { }
                    finally { tcs.TrySetResult(true); }
                }));
            }
            catch
            {
                tcs.TrySetResult(false);
            }
            return tcs.Task;
        }

        private async Task<bool> TryClickSendButtonAsync()
        {
            var clickScript = @"
(function(){
  try {
    var sel = 'button[type=submit], button[aria-label*=""send"" i], button[aria-label*=""submit"" i], button[aria-label*=""absenden"" i], button[class*=""send"" i], [role=button][aria-label*=""send"" i]';
    var btn = document.querySelector(sel);
    if (!btn) {
      var candidates = Array.from(document.querySelectorAll('button, [role=button]'));
      for (var i=0;i<candidates.length;i++){
        try {
          var txt = ((candidates[i].innerText || candidates[i].getAttribute('aria-label') || candidates[i].title) + '').toLowerCase();
          if (txt.indexOf('absend') !== -1 || txt.indexOf('send') !== -1 || txt.indexOf('submit') !== -1) { btn = candidates[i]; break; }
        } catch(e){}
      }
    }
    if (!btn) return false;
    try {
      var wasDisabled = !!btn.disabled;
      if (wasDisabled) { btn.disabled = false; btn.removeAttribute('disabled'); }
      btn.click();
      if (wasDisabled) { setTimeout(function(){ try { btn.disabled = true; btn.setAttribute('disabled',''); } catch(e){} }, 200); }
      return true;
    } catch(e){ return false; }
  } catch(e){ return false; }
})();";

            string? result = null;
            try
            {
                result = await _webView.CoreWebView2!.ExecuteScriptAsync(clickScript).ConfigureAwait(false);
                Trace.WriteLine(string.Format(Properties.Resources.SendTextToWebViewAsyncClickScriptResult0, result));
            }
            catch (Exception ex)
            {
                Trace.WriteLine(string.Format(Properties.Resources.ParsingClickScriptResultFailed0, ex));
            }
            return ParseJsBoolResult(result);
        }

        private async Task SendEnterViaNativeInputAsync()
        {
            try
            {
                await EnsureFocusOnUiThreadAsync().ConfigureAwait(false);
                await Task.Delay(NativeInputFocusDelayMs).ConfigureAwait(false);
                _input?.PressEnter();
                Trace.WriteLine("SendEnterViaNativeInputAsync: Enter gesendet.");
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"SendEnterViaNativeInputAsync fehlgeschlagen: {ex}");
            }
        }

        private async Task FallbackSendViaNativeInputAsync(string text, bool pressEnter)
        {
            try
            {
                await EnsureFocusOnUiThreadAsync().ConfigureAwait(false);
                await Task.Delay(NativeInputTextDelayMs).ConfigureAwait(false);
                if (!string.IsNullOrEmpty(text))
                {
                    _input?.TypeText(text);
                    await Task.Delay(NativeInputAfterTextDelayMs).ConfigureAwait(false);
                }
                if (pressEnter)
                {
                    _input?.PressEnter();
                    Trace.WriteLine("FallbackSendViaNativeInputAsync: Enter gesendet.");
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"FallbackSendViaNativeInputAsync fehlgeschlagen: {ex}");
            }
        }

        private async Task ShowOfflinePageAsync()
        {
            if (_webView.CoreWebView2 == null)
            {
                try { await _webView.EnsureCoreWebView2Async(null); }
                catch { }
            }
            try
            {
                string base64 = BitmapToBase64(Properties.Resources.nonet);
                var html = $@"
<html>
  <head>
    <meta charset=""utf-8"" />
    <title>offline / keine Verbindung</title>
    <style>
      body {{
        margin: 0; padding: 0; background: #202020; color: #ffffff;
        font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        display: flex; align-items: center; justify-content: center; height: 100vh;
      }}
      .wrapper {{ text-align: center; }}
      img {{ max-width: 256px; height: auto; margin-bottom: 1rem; }}
      h1 {{ margin: 0 0 0.5rem 0; font-size: 1.2rem; }}
      p {{ margin: 0; opacity: 0.8; }}
    </style>
  </head>
  <body>
    <div class=""wrapper"">
      <img src=""data:image/png;base64,{base64}"" alt=""offline"" />
      <h1>offline / keine Verbindung</h1>
      <p>Bitte Ã¼berprÃ¼fen Sie Ihre Internetverbindung und versuchen Sie es erneut.</p>
    </div>
  </body>
</html>";
                _webView.CoreWebView2?.NavigateToString(html);
            }
            catch { }
        }

        private static string BitmapToBase64(System.Drawing.Bitmap bmp)
        {
            try
            {
                using var ms = new MemoryStream();
                bmp.Save(ms, ImageFormat.Png);
                return Convert.ToBase64String(ms.ToArray());
            }
            catch { return string.Empty; }
        }

        private static bool ParseJsBoolResult(string? rawResult)
        {
            if (string.IsNullOrWhiteSpace(rawResult)) return false;
            try
            {
                var trimmed = rawResult.Trim();
                if (trimmed.Length >= 2 && trimmed[0] == '"' && trimmed[^1] == '"')
                    trimmed = trimmed[1..^1];
                if (string.Equals(trimmed, "true", StringComparison.OrdinalIgnoreCase)) return true;
                var el = System.Text.Json.JsonSerializer.Deserialize<System.Text.Json.JsonElement>(rawResult);
                if (el.ValueKind == System.Text.Json.JsonValueKind.True) return true;
                if (el.ValueKind == System.Text.Json.JsonValueKind.Object &&
                    el.TryGetProperty("ok", out var p) &&
                    p.ValueKind == System.Text.Json.JsonValueKind.True) return true;
            }
            catch { }
            return false;
        }


        // ------------------------------------------------------------------
        // Bild-Viewer
        // ------------------------------------------------------------------
        /// <summary>
        /// Lädt das Bild über die eingeloggte WebView-Session und zeigt es im
        /// eigenständigen Viewer-Fenster an. Wird sowohl vom Bildklick als auch
        /// vom Tray-Eintrag „Bild aus Zwischenablage öffnen" genutzt.
        /// Gibt false zurück, wenn das Bild nicht geladen werden konnte – der
        /// Aufrufer kann das dem Nutzer melden (beim Bildklick bewusst ignoriert).
        /// </summary>
        public async Task<bool> ShowImageViewerAsync(string imageUrl)
        {
            try
            {
                // Grok-CDN: Thumbnail-URL auf Vollbild-URL umschreiben
                // .../preview-image  -->  .../content
                string fullUrl = imageUrl.EndsWith("/preview-image", StringComparison.Ordinal)
                    ? imageUrl[..^"preview-image".Length] + "content"
                    : imageUrl;

                // Bild über den WebView laden (hat die nötigen Session-Cookies) und
                // als Base64 über die Async-Brücke zurückholen. Ein direktes
                // ExecuteScriptAsync funktioniert hier NICHT, weil dessen Rückgabewert
                // bei einem Promise nur "{}" wäre.
                string? b64 = await CallBridgeAsync(
                    token => $"window.__wrokFetchImage(" +
                             $"{System.Text.Json.JsonSerializer.Serialize(fullUrl)}, " +
                             $"{System.Text.Json.JsonSerializer.Serialize(token)});",
                    TimeSpan.FromSeconds(30)).ConfigureAwait(false);

                if (string.IsNullOrWhiteSpace(b64))
                {
                    Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [WebViewManager] " +
                                    $"ShowImageViewerAsync: kein Bilddatenstrom erhalten für {fullUrl}");
                    return false;
                }

                var bytes2 = Convert.FromBase64String(b64);

                // WICHTIG: Ein Bitmap, das aus einem Stream erzeugt wird, behält eine
                // Referenz auf diesen Stream und benötigt ihn über seine gesamte
                // Lebensdauer. Da das Bild erst später auf dem UI-Thread gezeichnet
                // wird (BeginInvoke), kopieren wir die Pixeldaten in ein eigenständiges
                // Bitmap – sonst drohen sporadische GDI+-Fehler beim Zeichnen/Speichern.
                System.Drawing.Bitmap bmp;
                using (var ms = new MemoryStream(bytes2))
                using (var decoded = new System.Drawing.Bitmap(ms))
                {
                    bmp = new System.Drawing.Bitmap(decoded);
                }

                if (_owner.IsHandleCreated && !_owner.IsDisposed)
                    _owner.BeginInvoke((System.Windows.Forms.MethodInvoker)(() =>
                    {
                        // Bewusst OHNE Owner-Fenster (siehe ShowImageViewer): sonst läge
                        // das Fenster in Windows immer über dem Hauptfenster.
                        _owner.ShowImageViewer(bmp, fullUrl);

                        // Als „letztes Bild" ablegen – die Daten liegen ohnehin schon
                        // dekodiert vor, kostet also keinen zweiten Download.
                        SaveAsLastImage(bmp, fullUrl);
                    }));
                else
                {
                    bmp.Dispose(); // kein Fenster mehr da → nicht lecken
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Log(ex, "ShowImageViewerAsync fehlgeschlagen");
                return false;
            }
        }

        /// <summary>
        /// Legt das gerade geöffnete Bild als „letztes Bild" ab (immer dieselbe Datei,
        /// wird überschrieben) und merkt Pfad und Quell-URL in den Settings.
        /// Als PNG gespeichert, damit es später verlustfrei und sicher ladbar ist.
        /// </summary>
        private static void SaveAsLastImage(System.Drawing.Bitmap bmp, string url)
        {
            try
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Wrok", "images");
                Directory.CreateDirectory(dir);

                var path = Path.Combine(dir, "last-image.png");
                var tmp  = path + ".tmp";

                // Erst temporär schreiben, dann ersetzen – ein Abbruch mittendrin
                // zerstört so nicht das zuvor gespeicherte Bild.
                bmp.Save(tmp, ImageFormat.Png);
                File.Copy(tmp, path, overwrite: true);
                try { File.Delete(tmp); } catch { /* unkritisch */ }

                Properties.Settings.Default.LastImagePath = path;
                Properties.Settings.Default.LastImageUrl  = url;
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                Log(ex, "Letztes Bild konnte nicht gespeichert werden");
            }
        }

        private static void Log(Exception ex, string message) =>
            Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [WebViewManager] {message}: {ex}");

        // P/Invoke
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}
