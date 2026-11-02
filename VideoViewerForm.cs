using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Wrok
{
    /// <summary>
    /// Zeigt ein Video im Endlos-Loop in einem eigenständigen Fenster.
    ///
    /// Umgesetzt mit einem zweiten WebView2, das dieselbe CoreWebView2Environment
    /// (und damit denselben User-Data-Ordner und dieselbe Login-Session) nutzt wie
    /// das Hauptfenster. Dadurch kann die Grok-URL direkt gestreamt werden – kein
    /// Download, kein Base64-Transport wie beim Bild-Viewer.
    ///
    /// Tastenkürzel:
    ///   Esc      schließen
    ///   Strg+P   anheften (immer im Vordergrund)
    ///   F11      pausieren
    ///   F12      fortsetzen
    /// </summary>
    internal sealed class VideoViewerForm : Form
    {
        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

        // Interne Adresse für die Player-Seite. Wird nie wirklich angefragt, sondern
        // von WebResourceRequested abgefangen – dient nur dazu, dem Dokument den
        // Ursprung grok.com zu geben, damit die Session-Cookies greifen.
        private const string PlayerUrl = "https://grok.com/__wrok_player";

        // Virtueller Host für lokale Videodateien: WebView2 darf aus einer
        // https-Seite nicht auf file:// zugreifen, deshalb wird der Ordner der
        // Datei per SetVirtualHostNameToFolderMapping unter diesem Namen
        // eingeblendet (inkl. Range-Requests, d. h. Spulen funktioniert).
        private const string LocalMediaHost = "wrok-media.local";

        private readonly WebView2 _web;
        private readonly string   _url;
        private          string   _playUrl;   // _url oder die virtuelle Adresse einer lokalen Datei
        private readonly CoreWebView2Environment? _env;

        /// <summary>
        /// Wird gemeldet, sobald die echten Videoabmessungen bekannt sind.
        /// Vorher lässt sich das Seitenverhältnis nicht kennen.
        /// </summary>
        public event Action<Size>? VideoSizeKnown;

        /// <summary>Wird gemeldet, wenn Strg+1 bis Strg+9 im Viewer gedrückt wird (1 bis 9).</summary>
        public event Action<int>? VideoSlotRequested;

        public VideoViewerForm(string url, CoreWebView2Environment? env)
        {
            _url = url ?? throw new ArgumentNullException(nameof(url));
            _playUrl = _url;
            _env = env;

            Text            = Properties.Resources.VideoViewerTitle;
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition   = FormStartPosition.CenterScreen;
            BackColor       = Color.FromArgb(28, 28, 28);
            ShowInTaskbar   = true;
            KeyPreview      = true;
            MinimumSize     = new Size(240, 180);
            ClientSize      = new Size(540, 960);   // vorläufig, wird nach Metadaten korrigiert

            TryApplyDarkTitleBar();

            _web = new WebView2 { Dock = DockStyle.Fill };
            Controls.Add(_web);

            KeyDown += OnKeyDown;
            _ = InitializeAsync();
        }

        private async Task InitializeAsync()
        {
            try
            {
                // Dieselbe Umgebung wie das Hauptfenster → gleiche Session/Cookies.
                await _web.EnsureCoreWebView2Async(_env);

                var core = _web.CoreWebView2;
                if (core == null) return;

                core.Settings.AreDefaultContextMenusEnabled = true;
                core.Settings.IsStatusBarEnabled            = false;

                // Browser-Accelerators aus, damit F11/F12 die Player-Seite erreichen
                // (Pause/Fortsetzen) statt Chromiums Vollbild bzw. DevTools auszulösen.
                core.Settings.AreBrowserAcceleratorKeysEnabled = false;

                // Lokale Datei? Dann ihren Ordner über einen virtuellen Host freigeben.
                if (!_url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                    !_url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) &&
                    File.Exists(_url))
                {
                    var dir = Path.GetDirectoryName(Path.GetFullPath(_url));
                    if (!string.IsNullOrEmpty(dir))
                    {
                        core.SetVirtualHostNameToFolderMapping(
                            LocalMediaHost, dir, CoreWebView2HostResourceAccessKind.Allow);
                        _playUrl = $"https://{LocalMediaHost}/{Uri.EscapeDataString(Path.GetFileName(_url))}";
                    }
                }

                core.WebMessageReceived += (s, args) =>
                {
                    var msg = args.TryGetWebMessageAsString();
                    if (msg != null && msg.StartsWith("videoSlot:", StringComparison.Ordinal))
                    {
                        if (int.TryParse(msg["videoSlot:".Length..], out int slotNo))
                            VideoSlotRequested?.Invoke(slotNo);
                        return;
                    }
                    if (msg == null || !msg.StartsWith("videoSize:", StringComparison.Ordinal)) return;

                    var parts = msg["videoSize:".Length..].Split('x');
                    if (parts.Length == 2 &&
                        int.TryParse(parts[0], out int vw) &&
                        int.TryParse(parts[1], out int vh) &&
                        vw > 0 && vh > 0)
                    {
                        VideoSizeKnown?.Invoke(new Size(vw, vh));
                    }
                };

                // Die Player-Seite wird unter einer grok.com-URL ausgeliefert, statt
                // sie per NavigateToString zu laden. Sonst hätte das Dokument einen
                // undurchsichtigen Ursprung, und SameSite-geschützte Session-Cookies
                // würden beim Laden des Videos nicht mitgeschickt.
                core.AddWebResourceRequestedFilter(PlayerUrl + "*", CoreWebView2WebResourceContext.All);
                core.WebResourceRequested += (s, args) =>
                {
                    try
                    {
                        if (!args.Request.Uri.StartsWith(PlayerUrl, StringComparison.OrdinalIgnoreCase)) return;
                        if (_env == null) return;

                        var html  = BuildPlayerHtml(_playUrl);
                        var bytes = System.Text.Encoding.UTF8.GetBytes(html);
                        var ms    = new MemoryStream(bytes);

                        args.Response = _env.CreateWebResourceResponse(
                            ms, 200, "OK",
                            "Content-Type: text/html; charset=utf-8\r\nCache-Control: no-store");
                    }
                    catch (Exception ex)
                    {
                        Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [VideoViewerForm] Player-Seite konnte nicht ausgeliefert werden: {ex}");
                    }
                };

                core.Navigate(PlayerUrl);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [VideoViewerForm] Initialisierung fehlgeschlagen: {ex}");
            }
        }

        /// <summary>
        /// Minimale Player-Seite: Endlos-Loop, startet mit Ton (Browser-Flag
        /// --autoplay-policy, Fallback stumm), mit Bedienelementen.
        /// </summary>
        private static string BuildPlayerHtml(string url)
        {
            string safeUrl = System.Text.Json.JsonSerializer.Serialize(url);

            return $@"<!doctype html>
<html><head><meta charset=""utf-8"">
<style>
  html, body {{ margin:0; height:100%; background:#1c1c1c; overflow:hidden; }}
  video {{ width:100%; height:100%; object-fit:contain; display:block; background:#1c1c1c; }}
</style></head>
<body>
<video id=""v"" loop autoplay playsinline controls></video>
<script>
  (function () {{
    var v = document.getElementById('v');
    v.src = {safeUrl};
    v.addEventListener('loadedmetadata', function () {{
      try {{
        window.chrome.webview.postMessage('videoSize:' + v.videoWidth + 'x' + v.videoHeight);
      }} catch (e) {{}}
    }});
    // Start mit Ton. Verweigert der Browser das, startet das Video stumm.
    var p = v.play();
    if (p && p.catch) p.catch(function () {{
      v.muted = true;
      v.play();
    }});
    // F11 pausiert, F12 setzt fort.
    document.addEventListener('keydown', function (e) {{
      if (e.key === 'F11') {{ e.preventDefault(); v.pause(); }}
      else if (e.key === 'F12') {{ e.preventDefault(); v.play().catch(function () {{}}); }}
    }});
    // Strg+1 bis Strg+9: Video-Slot im Hauptfenster abspielen lassen.
    document.addEventListener('keydown', function (e) {{
      if (e.ctrlKey && !e.altKey && !e.shiftKey && e.keyCode >= 49 && e.keyCode <= 57) {{
        e.preventDefault();
        try {{ window.chrome.webview.postMessage('videoSlot:' + (e.keyCode - 48)); }} catch (err) {{}}
      }}
    }});
  }})();
</script>
</body></html>";
        }

        private void OnKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Control && !e.Shift && !e.Alt && e.KeyCode >= Keys.D1 && e.KeyCode <= Keys.D9)
            {
                e.SuppressKeyPress = true;
                VideoSlotRequested?.Invoke((int)e.KeyCode - (int)Keys.D0);
                return;
            }

            switch (e.KeyCode)
            {
                case Keys.Escape:
                    Close();
                    break;

                // Hat das Formular statt des Videoframes den Fokus, sieht es F11/F12.
                // Dann den Frame fokussieren und dort pausieren bzw. starten.
                case Keys.F11:
                    e.SuppressKeyPress = true;
                    FocusPlayerAndRun("v.pause();");
                    break;

                case Keys.F12:
                    e.SuppressKeyPress = true;
                    FocusPlayerAndRun("v.play().catch(function () {});");
                    break;

                case Keys.P when e.Control:
                    TopMost = !TopMost;
                    Text = Properties.Resources.VideoViewerTitle + (TopMost ? "   📌" : string.Empty);
                    break;
            }
        }

        /// <summary>
        /// Setzt den Fokus auf den Videoframe und führt danach ein Kommando auf dem
        /// Video-Element aus. Ohne Fokus auf dem Frame erreicht JavaScript die Tasten nicht.
        /// </summary>
        private void FocusPlayerAndRun(string js)
        {
            try
            {
                _web.Focus();
                RunOnPlayer(js);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [VideoViewerForm] Steuerung fehlgeschlagen: {ex}");
            }
        }

        /// <summary>Pausiert bzw. startet das Video, ohne den Fokus zu verschieben.</summary>
        public void SetPlayback(bool play)
        {
            RunOnPlayer(play ? "v.play().catch(function () {});" : "v.pause();");
        }

        private void RunOnPlayer(string js)
        {
            _ = _web.CoreWebView2?.ExecuteScriptAsync(
                "(function () { var v = document.getElementById('v'); if (v) { " + js + " } })();");
        }

        private void TryApplyDarkTitleBar()
        {
            try
            {
                int val = 1;
                if (DwmSetWindowAttribute(Handle, 20, ref val, sizeof(int)) != 0)
                    DwmSetWindowAttribute(Handle, 19, ref val, sizeof(int));
            }
            catch { }
        }

        /// <summary>
        /// Gesamtbreite (inkl. Rahmen), die das Fenster bei der angegebenen
        /// Gesamthöhe braucht, damit das Video seitenverhältnisgerecht hineinpasst.
        /// </summary>
        public int WidthForHeight(int totalHeight, Size videoSize)
        {
            if (videoSize.Width <= 0 || videoSize.Height <= 0) return Width;

            int extraH = Height - ClientSize.Height;
            int extraW = Width  - ClientSize.Width;

            int clientH = Math.Max(1, totalHeight - extraH);
            int clientW = (int)Math.Round(videoSize.Width * (clientH / (double)videoSize.Height));

            return clientW + extraW;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                try { _web.Dispose(); } catch { }
            }
            base.Dispose(disposing);
        }
    }
}
