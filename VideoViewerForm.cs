using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Wrok
{
    /// <summary>
    /// Medienfenster für Videos und Bilder. Ein Fenster zeigt immer ein Medium; ein neues
    /// Medium ersetzt den Inhalt, ohne das Fenster zu schließen.
    ///
    /// Umgesetzt mit einem WebView2, das dieselbe CoreWebView2Environment nutzt wie das
    /// Hauptfenster (gleiche Session/Cookies). Videos werden direkt gestreamt, Bilder
    /// liegen als PNG im lokalen Ordner und werden über einen virtuellen Host geladen.
    ///
    /// Tastenkürzel (Video und Bild):
    ///   Esc      schließen
    ///   Strg+P   anheften (immer im Vordergrund)
    ///   Strg+1-9 Video-Slot abspielen
    /// Video:
    ///   F11      Pause / Wiedergabe umschalten
    /// Bild:
    ///   + / -    zoomen
    ///   0        100 %
    ///   Strg+F   einpassen
    ///   Ziehen   verschiebt das Bild
    /// </summary>
    internal sealed class VideoViewerForm : Form
    {
        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

        // Interne Adresse für die Player-Seite. Wird nie wirklich angefragt, sondern
        // von WebResourceRequested abgefangen – dient nur dazu, dem Dokument den
        // Ursprung grok.com zu geben, damit die Session-Cookies greifen.
        private const string PlayerUrl = "https://grok.com/__wrok_player";

        // Virtueller Host für lokale Dateien (Video und Bild). WebView2 darf aus einer
        // https-Seite nicht auf file:// zugreifen, deshalb wird der Ordner der Datei
        // unter diesem Namen eingeblendet (inkl. Range-Requests, Spulen funktioniert).
        private const string LocalMediaHost = "wrok-media.local";

        private const string VideoPageTemplate = @"<!doctype html>
<html><head><meta charset='utf-8'>
<style>
  html, body { margin:0; height:100%; background:#1c1c1c; overflow:hidden; }
  video { width:100%; height:100%; object-fit:contain; display:block; background:#1c1c1c; }
</style></head>
<body>
<video id='v' loop autoplay playsinline controls></video>
<script>
(function () {
  var v = document.getElementById('v');
  var noop = function () {};
  function post(m) { try { window.chrome.webview.postMessage(m); } catch (e) {} }
  v.src = __SRC__;
  // Haltepunkte (Sekunden, aufsteigend). Erreicht die Wiedergabe einen Punkt, wird pausiert.
  // Weiter geht es mit F11 bis zum nächsten Punkt.
  var stops = __STOPS__.slice().sort(function (a, b) { return a - b; });
  var next = 0, running = false, lastT = 0;
  function rearm() {
    next = 0;
    while (next < stops.length && stops[next] <= v.currentTime + 0.001) next++;
    lastT = v.currentTime;
  }
  function check() {
    // Schleife oder Sprung zurück an den Anfang: Haltepunkte wieder scharf schalten.
    if (v.currentTime < lastT - 0.5) rearm();
    lastT = v.currentTime;
    if (next < stops.length && !v.paused && v.currentTime >= stops[next]) {
      next++;
      v.pause();
    }
  }
  // requestAnimationFrame pausiert bei minimiertem Fenster, timeupdate läuft weiter.
  function loop() {
    if (v.paused || v.ended) { running = false; return; }
    check();
    requestAnimationFrame(loop);
  }
  v.addEventListener('play', function () {
    if (!running) { running = true; requestAnimationFrame(loop); }
  });
  v.addEventListener('timeupdate', check);
  v.addEventListener('seeked', rearm);
  rearm();
  v.addEventListener('loadedmetadata', function () {
    post('videoSize:' + v.videoWidth + 'x' + v.videoHeight);
  });
  // Start mit Ton. Verweigert der Browser das, startet das Video stumm.
  var p = v.play();
  if (p && p.catch) p.catch(function () { v.muted = true; v.play().catch(noop); });
  window.wrokToggle = function () { if (v.paused) { v.play().catch(noop); } else { v.pause(); } };
  window.wrokSet = function (play) { if (play) { v.play().catch(noop); } else { v.pause(); } };
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') { e.preventDefault(); post('close'); return; }
    if (e.ctrlKey && !e.altKey && !e.shiftKey && (e.key === 'p' || e.key === 'P')) { e.preventDefault(); post('pin'); return; }
    if (e.key === 'F11') { e.preventDefault(); window.wrokToggle(); return; }
    if (e.ctrlKey && !e.altKey && !e.shiftKey && e.keyCode >= 49 && e.keyCode <= 57) {
      e.preventDefault();
      post('videoSlot:' + (e.keyCode - 48));
    }
  });
})();
</script>
</body></html>";

        private const string ImagePageTemplate = @"<!doctype html>
<html><head><meta charset='utf-8'>
<style>
  html, body { margin:0; height:100%; background:#1c1c1c; overflow:hidden; }
  #wrap { position:fixed; left:0; top:0; right:0; bottom:0; overflow:hidden; cursor:grab; }
  #img { position:absolute; left:0; top:0; transform-origin:0 0; max-width:none; user-select:none; -webkit-user-drag:none; }
</style></head>
<body>
<div id='wrap'><img id='img' alt=''></div>
<script>
(function () {
  var wrap = document.getElementById('wrap');
  var img = document.getElementById('img');
  var s = 1, tx = 0, ty = 0, fit = true, nw = 0, nh = 0, drag = null;
  var MIN = 0.05, MAX = 16, STEP = 1.15;
  function post(m) { try { window.chrome.webview.postMessage(m); } catch (e) {} }
  function apply() { img.style.transform = 'translate(' + tx + 'px,' + ty + 'px) scale(' + s + ')'; }
  function fitView() {
    if (!nw) return;
    fit = true;
    s = Math.min(wrap.clientWidth / nw, wrap.clientHeight / nh);
    tx = (wrap.clientWidth - nw * s) / 2;
    ty = (wrap.clientHeight - nh * s) / 2;
    apply();
  }
  function zoomAt(ns, cx, cy) {
    if (!nw) return;
    ns = Math.max(MIN, Math.min(MAX, ns));
    tx = cx - (cx - tx) * (ns / s);
    ty = cy - (cy - ty) * (ns / s);
    s = ns; fit = false; apply();
  }
  function centre() { return [wrap.clientWidth / 2, wrap.clientHeight / 2]; }
  window.wrokZoomIn = function () { var c = centre(); zoomAt(s * STEP, c[0], c[1]); };
  window.wrokZoomOut = function () { var c = centre(); zoomAt(s / STEP, c[0], c[1]); };
  window.wrokZoomReset = function () { var c = centre(); zoomAt(1, c[0], c[1]); };
  window.wrokFit = fitView;
  img.addEventListener('load', function () {
    nw = img.naturalWidth; nh = img.naturalHeight;
    img.style.width = nw + 'px'; img.style.height = nh + 'px';
    post('videoSize:' + nw + 'x' + nh);
    fitView();
  });
  window.addEventListener('resize', function () { if (fit) fitView(); });
  wrap.addEventListener('mousedown', function (e) {
    if (e.button !== 0) return;
    drag = { x: e.clientX, y: e.clientY, tx: tx, ty: ty };
    wrap.style.cursor = 'grabbing';
  });
  window.addEventListener('mousemove', function (e) {
    if (!drag) return;
    tx = drag.tx + (e.clientX - drag.x);
    ty = drag.ty + (e.clientY - drag.y);
    fit = false; apply();
  });
  window.addEventListener('mouseup', function () { drag = null; wrap.style.cursor = 'grab'; });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') { e.preventDefault(); post('close'); return; }
    if (e.ctrlKey && !e.altKey && !e.shiftKey && (e.key === 'p' || e.key === 'P')) { e.preventDefault(); post('pin'); return; }
    if (e.ctrlKey && (e.key === 'f' || e.key === 'F')) { e.preventDefault(); fitView(); return; }
    if (e.ctrlKey || e.altKey) return;
    if (e.key === '+' || e.key === '=' || e.key === 'Add') { e.preventDefault(); window.wrokZoomIn(); }
    else if (e.key === '-' || e.key === 'Subtract') { e.preventDefault(); window.wrokZoomOut(); }
    else if (e.key === '0') { e.preventDefault(); window.wrokZoomReset(); }
  });
  img.src = __SRC__;
})();
</script>
</body></html>";

        private enum MediaKind { None, Video, Image }

        private readonly WebView2 _web;
        private readonly CoreWebView2Environment? _env;

        private bool _coreReady;
        private MediaKind _kind = MediaKind.None;
        private string _target = string.Empty;    // URL oder lokaler Pfad, wie übergeben
        private string _mediaUrl = string.Empty;  // Adresse, die die Seite tatsächlich lädt
        private double[] _stops = Array.Empty<double>();  // Haltepunkte in Sekunden (nur Video)
        private Size _imageSize = Size.Empty;
        private int _navCounter;

        /// <summary>
        /// Wird gemeldet, sobald die echten Abmessungen des Mediums bekannt sind.
        /// Vorher lässt sich das Seitenverhältnis nicht kennen.
        /// </summary>
        public event Action<Size>? VideoSizeKnown;

        /// <summary>Wird gemeldet, wenn Strg+1 bis Strg+9 im Fenster gedrückt wird (1 bis 9).</summary>
        public event Action<int>? VideoSlotRequested;

        public bool IsVideo => _kind == MediaKind.Video;
        private bool IsImage => _kind == MediaKind.Image;

        public VideoViewerForm(CoreWebView2Environment? env)
        {
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

        /// <summary>Zeigt ein Video (URL oder lokaler Pfad) im Fenster.</summary>
        public void ShowVideo(string url)
        {
            if (url == null) throw new ArgumentNullException(nameof(url));
            _kind   = MediaKind.Video;
            _target = SplitStops(url, out _stops);
            UpdateTitle();
            if (_coreReady) Navigate();
        }

        /// <summary>Zeigt ein Bild (lokale PNG-Datei) im Fenster.</summary>
        public void ShowImage(string localPath, Size size)
        {
            _kind      = MediaKind.Image;
            _target    = localPath ?? throw new ArgumentNullException(nameof(localPath));
            _stops     = Array.Empty<double>();
            _imageSize = size;
            UpdateTitle();
            if (_coreReady) Navigate();
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

                // Browser-Accelerators aus, damit F11 und Strg+F die Seite erreichen,
                // statt Chromiums Vollbild bzw. Suche auszulösen.
                core.Settings.AreBrowserAcceleratorKeysEnabled = false;

                core.WebMessageReceived += OnWebMessage;

                // Die Player-Seite wird unter einer grok.com-URL ausgeliefert, statt
                // sie per NavigateToString zu laden. Sonst hätte das Dokument einen
                // undurchsichtigen Ursprung, und SameSite-geschützte Session-Cookies
                // würden beim Laden nicht mitgeschickt.
                core.AddWebResourceRequestedFilter(PlayerUrl + "*", CoreWebView2WebResourceContext.All);
                core.WebResourceRequested += OnPlayerRequested;

                _coreReady = true;
                if (_kind != MediaKind.None) Navigate();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [VideoViewerForm] Initialisierung fehlgeschlagen: {ex}");
            }
        }

        private void OnWebMessage(object? sender, CoreWebView2WebMessageReceivedEventArgs args)
        {
            var msg = args.TryGetWebMessageAsString();
            if (string.IsNullOrEmpty(msg)) return;

            if (msg == "close")
            {
                Close();
                return;
            }

            if (msg == "pin")
            {
                TogglePin();
                return;
            }

            if (msg.StartsWith("videoSlot:", StringComparison.Ordinal))
            {
                if (int.TryParse(msg["videoSlot:".Length..], out int slotNo))
                    VideoSlotRequested?.Invoke(slotNo);
                return;
            }

            if (msg.StartsWith("videoSize:", StringComparison.Ordinal))
            {
                var parts = msg["videoSize:".Length..].Split('x');
                if (parts.Length == 2 &&
                    int.TryParse(parts[0], out int w) &&
                    int.TryParse(parts[1], out int h) &&
                    w > 0 && h > 0)
                {
                    if (IsImage)
                    {
                        _imageSize = new Size(w, h);
                        UpdateTitle();
                    }
                    VideoSizeKnown?.Invoke(new Size(w, h));
                }
            }
        }

        private void OnPlayerRequested(object? sender, CoreWebView2WebResourceRequestedEventArgs args)
        {
            try
            {
                if (!args.Request.Uri.StartsWith(PlayerUrl, StringComparison.OrdinalIgnoreCase)) return;
                if (_env == null) return;

                var bytes = Encoding.UTF8.GetBytes(BuildPageHtml());
                args.Response = _env.CreateWebResourceResponse(
                    new MemoryStream(bytes), 200, "OK",
                    "Content-Type: text/html; charset=utf-8\r\nCache-Control: no-store");
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [VideoViewerForm] Player-Seite konnte nicht ausgeliefert werden: {ex}");
            }
        }

        /// <summary>
        /// Lädt das aktuelle Medium. Jede Navigation bekommt eine eigene Abfrage-Nummer,
        /// damit WebView2 die Seite nicht aus dem Cache holt, wenn das Medium wechselt.
        /// </summary>
        private void Navigate()
        {
            var core = _web.CoreWebView2;
            if (core == null || _kind == MediaKind.None) return;

            _mediaUrl = ResolveMediaUrl(core, _target);
            core.Navigate(PlayerUrl + "?n=" + (++_navCounter).ToString());
        }

        /// <summary>
        /// Liefert die Adresse für das Medium. Lokale Dateien bekommen ihren Ordner als
        /// virtuellen Host; andere Adressen bleiben unverändert.
        /// </summary>
        private static string ResolveMediaUrl(CoreWebView2 core, string target)
        {
            if (target.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                target.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                return target;

            if (!File.Exists(target)) return target;

            var full = Path.GetFullPath(target);
            var dir  = Path.GetDirectoryName(full);
            if (!string.IsNullOrEmpty(dir))
            {
                core.SetVirtualHostNameToFolderMapping(
                    LocalMediaHost, dir, CoreWebView2HostResourceAccessKind.Allow);
            }
            return $"https://{LocalMediaHost}/{Uri.EscapeDataString(Path.GetFileName(full))}";
        }

        private string BuildPageHtml()
        {
            string src      = JsonSerializer.Serialize(_mediaUrl);
            string stops    = JsonSerializer.Serialize(_stops);
            string template = IsImage ? ImagePageTemplate : VideoPageTemplate;
            return template.Replace("__SRC__", src).Replace("__STOPS__", stops);
        }

        /// <summary>
        /// Trennt eine Haltepunkt-Liste am Ende einer Video-Adresse ab, z. B.
        /// "https://example.com/v.mp4 [5, 35, 1:41]". Ohne Klammerliste bleibt die Adresse
        /// unverändert. Zeiten in Sekunden oder als m:ss, sortiert und ohne Dubletten.
        /// </summary>
        internal static string SplitStops(string input, out double[] stops)
        {
            stops = Array.Empty<double>();
            if (string.IsNullOrEmpty(input)) return input ?? string.Empty;

            var m = Regex.Match(input, @"^(.*\S)\s+\[([^\[\]]*)\]\s*$", RegexOptions.Singleline);
            if (!m.Success) return input;

            var list = new List<double>();
            foreach (var raw in m.Groups[2].Value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                if (TryParseStop(raw, out double t) && t >= 0) list.Add(t);
            }
            stops = list.Distinct().OrderBy(x => x).ToArray();
            return m.Groups[1].Value;
        }

        private static bool TryParseStop(string raw, out double seconds)
        {
            seconds = 0;
            int colon = raw.IndexOf(':');
            if (colon < 0)
                return double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out seconds);

            if (double.TryParse(raw[..colon], NumberStyles.Float, CultureInfo.InvariantCulture, out double min) &&
                double.TryParse(raw[(colon + 1)..], NumberStyles.Float, CultureInfo.InvariantCulture, out double sec))
            {
                seconds = min * 60 + sec;
                return true;
            }
            return false;
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

                case Keys.P when e.Control:
                    TogglePin();
                    break;

                // Hat das Formular statt der Seite den Fokus, sieht es die Tasten.
                // Dann den Frame fokussieren und dort ausführen.
                case Keys.F11 when IsVideo:
                    e.SuppressKeyPress = true;
                    FocusAndRun("if (window.wrokToggle) { window.wrokToggle(); }");
                    break;

                case Keys.Oemplus when IsImage:
                case Keys.Add when IsImage:
                    e.SuppressKeyPress = true;
                    FocusAndRun("if (window.wrokZoomIn) { window.wrokZoomIn(); }");
                    break;

                case Keys.OemMinus when IsImage:
                case Keys.Subtract when IsImage:
                    e.SuppressKeyPress = true;
                    FocusAndRun("if (window.wrokZoomOut) { window.wrokZoomOut(); }");
                    break;

                case Keys.D0 when IsImage && !e.Control:
                case Keys.NumPad0 when IsImage:
                    e.SuppressKeyPress = true;
                    FocusAndRun("if (window.wrokZoomReset) { window.wrokZoomReset(); }");
                    break;

                case Keys.F when e.Control && IsImage:
                    e.SuppressKeyPress = true;
                    FocusAndRun("if (window.wrokFit) { window.wrokFit(); }");
                    break;
            }
        }

        private void TogglePin()
        {
            TopMost = !TopMost;
            UpdateTitle();
        }

        private void UpdateTitle()
        {
            string baseTitle = IsImage && !_imageSize.IsEmpty
                ? string.Format(Properties.Resources.ImageViewerTitle, _imageSize.Width, _imageSize.Height)
                : Properties.Resources.VideoViewerTitle;
            Text = baseTitle + (TopMost ? "   📌" : string.Empty);
        }

        /// <summary>Pausiert bzw. startet das Video, ohne den Fokus zu verschieben.</summary>
        public void SetPlayback(bool play)
        {
            RunScript("if (window.wrokSet) { window.wrokSet(" + (play ? "true" : "false") + "); }");
        }

        /// <summary>Wechselt zwischen Pause und Wiedergabe, je nach aktuellem Zustand.</summary>
        public void TogglePlayback()
        {
            RunScript("if (window.wrokToggle) { window.wrokToggle(); }");
        }

        /// <summary>Setzt den Fokus auf die Seite und führt danach das Skript aus.</summary>
        private void FocusAndRun(string js)
        {
            try
            {
                _web.Focus();
                RunScript(js);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [VideoViewerForm] Steuerung fehlgeschlagen: {ex}");
            }
        }

        private void RunScript(string js)
        {
            try
            {
                _ = _web.CoreWebView2?.ExecuteScriptAsync(js);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [VideoViewerForm] Skript fehlgeschlagen: {ex}");
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
        /// Gesamthöhe braucht, damit das Medium seitenverhältnisgerecht hineinpasst.
        /// </summary>
        public int WidthForHeight(int totalHeight, Size mediaSize)
        {
            if (mediaSize.Width <= 0 || mediaSize.Height <= 0) return Width;

            int extraH = Height - ClientSize.Height;
            int extraW = Width  - ClientSize.Width;

            int clientH = Math.Max(1, totalHeight - extraH);
            int clientW = (int)Math.Round(mediaSize.Width * (clientH / (double)mediaSize.Height));

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
