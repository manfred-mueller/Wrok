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
    /// </summary>
    internal sealed class VideoViewerForm : Form
    {
        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

        // Interne Adresse für die Player-Seite. Wird nie wirklich angefragt, sondern
        // von WebResourceRequested abgefangen – dient nur dazu, dem Dokument den
        // Ursprung grok.com zu geben, damit die Session-Cookies greifen.
        private const string PlayerUrl = "https://grok.com/__wrok_player";

        private readonly WebView2 _web;
        private readonly string   _url;
        private readonly CoreWebView2Environment? _env;

        /// <summary>
        /// Wird gemeldet, sobald die echten Videoabmessungen bekannt sind.
        /// Vorher lässt sich das Seitenverhältnis nicht kennen.
        /// </summary>
        public event Action<Size>? VideoSizeKnown;

        public VideoViewerForm(string url, CoreWebView2Environment? env)
        {
            _url = url ?? throw new ArgumentNullException(nameof(url));
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

                core.WebMessageReceived += (s, args) =>
                {
                    var msg = args.TryGetWebMessageAsString();
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

                        var html  = BuildPlayerHtml(_url);
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
        /// Minimale Player-Seite: Endlos-Loop, stumm gestartet (sonst greift die
        /// Autoplay-Sperre), mit Bedienelementen zum Aufdrehen.
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
<video id=""v"" loop autoplay muted playsinline controls></video>
<script>
  (function () {{
    var v = document.getElementById('v');
    v.src = {safeUrl};
    v.addEventListener('loadedmetadata', function () {{
      try {{
        window.chrome.webview.postMessage('videoSize:' + v.videoWidth + 'x' + v.videoHeight);
      }} catch (e) {{}}
    }});
    // Autoplay kann trotz muted abgelehnt werden – dann erneut versuchen.
    var p = v.play();
    if (p && p.catch) p.catch(function () {{ setTimeout(function () {{ v.play(); }}, 300); }});
  }})();
</script>
</body></html>";
        }

        private void OnKeyDown(object? sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Escape:
                    Close();
                    break;

                case Keys.P when e.Control:
                    TopMost = !TopMost;
                    Text = Properties.Resources.VideoViewerTitle + (TopMost ? "   📌" : string.Empty);
                    break;
            }
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
