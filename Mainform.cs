using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;
using System;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsInput;
using WindowsInput.Native;

namespace Wrok
{
    public partial class MainForm : Form
    {
        // Wiederverwendbarer HttpClient
        private static readonly HttpClient _httpClient = new HttpClient();

        // Steuerelemente / Ressourcen
        private WebView2? webView;
        private NotifyIcon? trayIcon;
        private ContextMenuStrip? trayMenu;
        private ToolStripMenuItem? macrosMenu;

        // Basis-URL und Menüeinträge für das Tray-Menü
        private readonly string baseUrl = "https://grok.com/";
        private readonly (string name, string url)[] menuPages = new[]
        {
            (Properties.Resources.Settings, "?_s=home"),
        };

        // --- Globaler Hotkey ---
        private const int HOTKEY_ID = 0x9000;
        private const int WM_HOTKEY = 0x0312;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        // --- Inaktivitätstimer (automatisches Minimieren) ---
        private System.Windows.Forms.Timer? inactivityTimer;
        private TimeSpan inactivityTimeout = TimeSpan.FromSeconds(30);
        private ActivityMessageFilter? activityFilter;
        private bool inactivityEnabled = true;
        private readonly int[] inactivityOptions = new[] { 0, 30, 60, 90 };

        // Ergänze Felder für Aktivitäts-Tracking
        private DateTime _lastActivity = DateTime.UtcNow;
        private readonly object _activityLock = new object();

        // Neue Felder / Hotkey-IDs
        private InputSimulator? _inputSimulator;

        // === Hotkey-IDs: Basis separat halten (keine Kollision mit anderem HOTKEY_ID) ===
        private const int HOTKEY_BASE = 0x9100;
        private const int HOTKEY_MACRO_1 = HOTKEY_BASE + 0;   // Strg+1
        private const int HOTKEY_MACRO_2 = HOTKEY_BASE + 1;   // Strg+2
        private const int HOTKEY_MACRO_3 = HOTKEY_BASE + 2;   // Strg+3
        private const int HOTKEY_MACRO_4 = HOTKEY_BASE + 3;   // Strg+4
        private const int HOTKEY_MACRO_5 = HOTKEY_BASE + 4;   // Strg+5
        private const int HOTKEY_MACRO_6 = HOTKEY_BASE + 5;   // Strg+^ (dynamisch ermittelt)

        // P/Invoke: SetForegroundWindow, damit Zielfenster Fokus bekommt (wenn nötig)
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern short VkKeyScan(char ch);

        public MainForm()
        {
            InitializeComponent();

            // Globalen Nachrichtenfilter installieren, um Aktivität zu erkennen (WeakReference im Filter verwenden)
            activityFilter = new ActivityMessageFilter(this);
            try
            {
                Application.AddMessageFilter(activityFilter);
            }
            catch
            {
                activityFilter = null; // OK, falls Hinzufügen aus irgendeinem Grund fehlschlägt
            }

            // Tray initialisieren und Icon dem aktuellen Theme anpassen
            InitializeTrayIcon();
            RefreshTheme(); // Zentrale Theme-Aktualisierung
            LoadWindowSettings();

            // Prüfen, ob die Anwendung zum ersten Mal gestartet wird (Marker-Datei in LocalAppData)
            var markerDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Wrok");
            var firstRunMarker = Path.Combine(markerDir, "firstrun.marker");
            var savedSeconds = Properties.Settings.Default.InactivityTimeoutSeconds;
            bool isFirstRun = false;

            try
            {
                if (!Directory.Exists(markerDir))
                    Directory.CreateDirectory(markerDir);

                isFirstRun = !File.Exists(firstRunMarker);
            }
            catch
            {
                isFirstRun = false;
            }

            if (isFirstRun || savedSeconds <= 0)
            {
                inactivityTimeout = TimeSpan.Zero;
                inactivityEnabled = false;
                Properties.Settings.Default.InactivityTimeoutSeconds = 0;
                Properties.Settings.Default.Save();

                try
                {
                    File.WriteAllText(firstRunMarker, DateTime.UtcNow.ToString("o"));
                }
                catch
                {
                    // Keine Aktion erforderlich
                }
            }
            else
            {
                inactivityTimeout = TimeSpan.FromSeconds(savedSeconds);
                inactivityEnabled = savedSeconds > 0;
            }

            _ = InitializeWebViewAsync();
            InitializeInactivityTimer();

            // Seite im Hintergrund laden
            try
            {
                _ = LoadUrlAsync(baseUrl, bringToFront: false);
            }
            catch
            {
                // Keine Aktion erforderlich
            }

            // Form-Events für WindowState/Position speichern
            this.Resize += MainForm_Resize;
            this.ResizeEnd += MainForm_ResizeEnd;
            this.Move += MainForm_Move;

            // Inaktivitätszustand im Tray-Menü anzeigen
            UpdateTrayMenuInactivityState();
        }

        private void InitializeComponent()
        {
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = version != null ? $"Wrok {version.Major}.{version.Minor}.{version.Build}" : "Wrok";
            this.WindowState = FormWindowState.Normal;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.ShowInTaskbar = false;
            this.Visible = false;
        }

        // Gespeicherte Fensterposition und -größe laden
        private void LoadWindowSettings()
        {
            var s = Properties.Settings.Default;

            if (s.WindowWidth > 0 && s.WindowHeight > 0)
            {
                this.StartPosition = FormStartPosition.Manual;
                var desired = new Rectangle(
                    s.WindowLeft,
                    s.WindowTop,
                    s.WindowWidth,
                    s.WindowHeight);

                bool intersects = false;
                foreach (var scr in Screen.AllScreens)
                {
                    if (scr.WorkingArea.IntersectsWith(desired))
                    {
                        intersects = true;
                        break;
                    }
                }

                if (intersects)
                {
                    this.Bounds = desired;
                }
                else
                {
                    this.StartPosition = FormStartPosition.CenterScreen;
                }
            }

            if (s.IsMaximized)
            {
                this.WindowState = FormWindowState.Maximized;
            }
        }

        // Aktuelle Fenstergeometrie speichern
        private void SaveWindowSettings()
        {
            try
            {
                var s = Properties.Settings.Default;
                Rectangle bounds;

                if (this.WindowState == FormWindowState.Maximized || this.WindowState == FormWindowState.Minimized)
                    bounds = this.RestoreBounds;
                else
                    bounds = this.Bounds;

                s.WindowLeft = bounds.Left;
                s.WindowTop = bounds.Top;
                s.WindowWidth = Math.Max(100, bounds.Width);
                s.WindowHeight = Math.Max(100, bounds.Height);
                s.IsMaximized = (this.WindowState == FormWindowState.Maximized);
                s.Save();
            }
            catch
            {
                // Keine Aktion erforderlich
            }
        }

        private void MainForm_Resize(object? sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized || this.WindowState == FormWindowState.Maximized)
            {
                SaveWindowSettings();
            }
        }

        private void MainForm_ResizeEnd(object? sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Normal)
            {
                SaveWindowSettings();
            }
        }

        private void MainForm_Move(object? sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Normal)
            {
                SaveWindowSettings();
            }
        }

        /// <summary>
        /// WebView2 initialisieren und feste Runtime nutzen.
        /// </summary>
        private async Task InitializeWebViewAsync()
        {
            webView = new WebView2
            {
                Dock = DockStyle.Fill
            };
            this.Controls.Add(webView);

            string userDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Wrok",
                "WebView2Data");

            try
            {
                Directory.CreateDirectory(userDataPath);
            }
            catch
            {
                // Keine Aktion erforderlich
            }

            webView.CoreWebView2InitializationCompleted += async (s, e) =>
            {
                if (webView.CoreWebView2 != null)
                {
                    // helperScript als Verbatim-String ohne C#-Escape-Sequenzen wie \".
                    var helperScript = @"
(function() {
  const resetActivity = function() { window.chrome.webview.postMessage('resetActivity'); };
  ['mousemove','mousedown','keydown','scroll','touchstart'].forEach(function(ev){ window.addEventListener(ev, resetActivity, { passive: true }); });

  window.__wrokSend = function(text, pressEnter) {
    try {
      if (typeof text !== 'string') text = String(text || '');
      var target = document.activeElement;
      if (!target || target === document.body) {
        target = document.querySelector('[contenteditable], textarea, input[type=text], input[type=search], [role=textbox]');
      }
      if (!target) return false;
      try { target.focus(); } catch (e) {}

      var tag = (target.tagName || '').toUpperCase();
      if (tag === 'INPUT' || tag === 'TEXTAREA' || 'value' in target) {
        var start = typeof target.selectionStart === 'number' ? target.selectionStart : (target.value || '').length;
        var end = typeof target.selectionEnd === 'number' ? target.selectionEnd : start;
        var val = target.value || '';
        var prefix = (start > 0 && val.charAt(start - 1) !== ' ') ? ' ' : '';
        var newVal = val.slice(0, start) + prefix + text + val.slice(end);
        target.value = newVal;
        var newPos = start + prefix.length + text.length;
        try { target.setSelectionRange(newPos, newPos); } catch (e) {}
        target.dispatchEvent(new Event('input', { bubbles: true }));
        target.dispatchEvent(new Event('change', { bubbles: true }));
        if (pressEnter) {
          try {
            if (target.form) {
              if (typeof target.form.requestSubmit === 'function') target.form.requestSubmit();
              else target.form.submit();
            } else {
              target.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
              target.dispatchEvent(new KeyboardEvent('keyup',   { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
            }
          } catch (e) {}
        }
        return true;
      }

      var sel = window.getSelection();
      var range = sel && sel.rangeCount ? sel.getRangeAt(0) : null;
      if (!range) {
        var prefix = (target.innerText && target.innerText.slice(-1) !== ' ') ? ' ' : '';
        target.innerText = (target.innerText || '') + prefix + text;
        var r2 = document.createRange();
        r2.selectNodeContents(target);
        r2.collapse(false);
        sel.removeAllRanges();
        sel.addRange(r2);
        try { target.dispatchEvent(new InputEvent('input', { bubbles: true })); } catch (e) {}
        if (pressEnter) {
          var btn = document.querySelector('button[type=submit], button[aria-label*=""send"" i], button[class*=""send"" i], [role=button][aria-label*=""send"" i]');
          if (btn) { try { btn.click(); } catch (e) {} }
          else {
            try { target.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true })); } catch (e) {}
            try { target.dispatchEvent(new KeyboardEvent('keyup',   { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true })); } catch (e) {}
          }
        }
        return true;
      }

      var prefix = '';
      var sc = range.startContainer;
      var off = range.startOffset;
      var prevChar = '';
      if (sc.nodeType === Node.TEXT_NODE) {
        if (off > 0) prevChar = sc.textContent.charAt(off - 1) || '';
        else {
          var prev = sc.previousSibling;
          if (prev && prev.nodeType === Node.TEXT_NODE) prevChar = prev.textContent.charAt(prev.textContent.length - 1) || '';
        }
      } else {
        var prevNode = range.startContainer.childNodes[off - 1];
        if (prevNode && prevNode.nodeType === Node.TEXT_NODE) prevChar = prevNode.textContent.charAt(prevNode.textContent.length - 1) || '';
      }
      if (prevChar && prevChar !== ' ') prefix = ' ';
      var node = document.createTextNode(prefix + text);
      range.insertNode(node);
      range.setStartAfter(node);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
      try { target.dispatchEvent(new InputEvent('input', { bubbles: true })); } catch (e) {}
      if (pressEnter) {
        var btn2 = document.querySelector('button[type=submit], button[aria-label*=""send"" i], button[class*=""send"" i], [role=button][aria-label*=""send"" i]');
        if (btn2) { try { btn2.click(); } catch (e) {} }
        else {
          try { range.insertNode(document.createElement('br')); } catch (e) {}
          try { range.setStartAfter(node.nextSibling || node); } catch (e) {}
          try { range.collapse(true); } catch (e) {}
          try { sel.removeAllRanges(); sel.addRange(range); } catch (e) {}
          try { target.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true })); } catch (e) {}
          try { target.dispatchEvent(new KeyboardEvent('keyup',   { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true })); } catch (e) {}
        }
      }
      return true;
    } catch (e) {
      return false;
    }
  };
})();";
                    try
                    {
                        await webView.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync(helperScript);
                    }
                    catch
                    {
                        // ignore
                    }

                    // Wenn die Seite schon geladen ist, helperScript sofort in das aktuelle Dokument injizieren (verhindert "erstes Mal fehlt helper")
                    try
                    {
                        await webView.CoreWebView2.ExecuteScriptAsync(helperScript);
                    }
                    catch (Exception ex)
                    {
                        Trace.WriteLine($"Inject helperScript to current document failed: {ex}");
                    }

                    webView.CoreWebView2.WebMessageReceived += (sender, args) =>
                    {
                        if (args.TryGetWebMessageAsString() == "resetActivity")
                        {
                            ResetInactivityTimer();
                        }
                    };
                }
            };

            try
            {
                var env = await CoreWebView2Environment.CreateAsync(userDataFolder: userDataPath);
                await webView.EnsureCoreWebView2Async(env);
            }
            catch
            {
                // Keine Aktion erforderlich
            }
        }

        // Sehr kurzer, performanter per-call Aufruf der einmal registrierten helper-Funktion.
        // Das erzeugt minimalen JS-Overhead (nur zwei Argumente, kein lange Script-Parsing pro Aufruf).
        private async Task SendTextToWebViewAsync(string text, bool pressEnter = false)
        {
            if (webView?.CoreWebView2 == null)
                return;

            // Fokus sicherstellen (UI-Thread)
            try
            {
                if (!this.IsDisposed && this.IsHandleCreated)
                {
                    var tcs = new TaskCompletionSource<bool>();
                    this.BeginInvoke((MethodInvoker)(() =>
                    {
                        try
                        {
                            webView?.Focus();
                            SetForegroundWindow(this.Handle);
                        }
                        catch { }
                        finally { tcs.TrySetResult(true); }
                    }));
                    await tcs.Task.ConfigureAwait(false);
                }
            }
            catch { }

            // Kurzes Timing-Window
            await Task.Delay(120).ConfigureAwait(false);

            var payload = System.Text.Json.JsonSerializer.Serialize(text);
            var callScript = $"(function(){{ try {{ return window.__wrokSend ? window.__wrokSend({payload}, {(pressEnter ? "true" : "false")}) : false; }} catch(e) {{ return false; }} }})();";

            string? rawResult = null;
            try
            {
                rawResult = await webView.CoreWebView2.ExecuteScriptAsync(callScript).ConfigureAwait(false);
                Trace.WriteLine($"SendTextToWebViewAsync: first ExecuteScriptAsync result={rawResult}");
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"ExecuteScriptAsync failed (first attempt): {ex}");
            }

            bool jsSucceeded = false;
            try
            {
                if (!string.IsNullOrWhiteSpace(rawResult))
                {
                    var trimmed = rawResult.Trim();
                    if (trimmed.Length >= 2 && trimmed[0] == '"' && trimmed[^1] == '"')
                        trimmed = trimmed.Substring(1, trimmed.Length - 2);

                    if (string.Equals(trimmed, "true", StringComparison.OrdinalIgnoreCase))
                        jsSucceeded = true;
                    else
                    {
                        // ggf. als JsonElement parsen
                        try
                        {
                            var el = System.Text.Json.JsonSerializer.Deserialize<System.Text.Json.JsonElement>(rawResult);
                            if (el.ValueKind == System.Text.Json.JsonValueKind.True) jsSucceeded = true;
                            else if (el.ValueKind == System.Text.Json.JsonValueKind.Object && el.TryGetProperty("ok", out var p) && p.ValueKind == System.Text.Json.JsonValueKind.True) jsSucceeded = true;
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Parsing ExecuteScriptAsync result failed: {ex}");
            }

            // Wenn Enter gewünscht, versuche zusätzlich verzögert, gezielt den Send-Button zu klicken.
            if (pressEnter)
            {
                // Warte kurz, damit die Seite Zeit hat, UI (z.B. Button enable) zu aktualisieren
                await Task.Delay(140).ConfigureAwait(false);

                var clickScript = @"
(function(){
  try {
    var sel = 'button[type=submit], button[aria-label*=""send"" i], button[aria-label*=""submit"" i], button[aria-label*=""absenden"" i], button[class*=""send"" i], [role=button][aria-label*=""send"" i]';
    var btn = document.querySelector(sel);
    if (!btn) {
      // fallback: suche sichtbaren Button mit Text 'Absenden'/'Senden'/'Submit'
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

                string? clickResult = null;
                try
                {
                    clickResult = await webView.CoreWebView2.ExecuteScriptAsync(clickScript).ConfigureAwait(false);
                    Trace.WriteLine($"SendTextToWebViewAsync: clickScript result={clickResult}");
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"ExecuteScriptAsync(clickScript) failed: {ex}");
                }

                bool clickSucceeded = false;
                try
                {
                    if (!string.IsNullOrWhiteSpace(clickResult))
                    {
                        var t = clickResult.Trim();
                        if (t.Length >= 2 && t[0] == '"' && t[^1] == '"') t = t.Substring(1, t.Length - 2);
                        clickSucceeded = string.Equals(t, "true", StringComparison.OrdinalIgnoreCase);
                    }
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"Parsing clickScript result failed: {ex}");
                }

                if (clickSucceeded)
                {
                    Trace.WriteLine("SendTextToWebViewAsync: click succeeded, returning.");
                    return;
                }

                // Falls JS bereits erfolgreiches Insert/submit meldete, return
                if (jsSucceeded)
                {
                    Trace.WriteLine("SendTextToWebViewAsync: jsSucceeded true, returning (no click).");
                    return;
                }

                // Wenn alles JS fehlschlug -> InputSimulator Fallback unten
            }
            else
            {
                if (jsSucceeded)
                    return;
            }

            // Fallback: InputSimulator (TextEntry + optional Enter)
            try
            {
                // nochmals Fokus auf UI-Thread setzen
                try
                {
                    if (!this.IsDisposed && this.IsHandleCreated)
                    {
                        this.BeginInvoke((MethodInvoker)(() =>
                        {
                            try
                            {
                                webView?.Focus();
                                SetForegroundWindow(this.Handle);
                            }
                            catch { }
                        }));
                    }
                }
                catch { }

                await Task.Delay(150).ConfigureAwait(false);

                if (!string.IsNullOrEmpty(text))
                {
                    _inputSimulator?.Keyboard.TextEntry(text);
                    await Task.Delay(40).ConfigureAwait(false);
                }

                if (pressEnter)
                {
                    _inputSimulator?.Keyboard.KeyPress(VirtualKeyCode.RETURN);
                    Trace.WriteLine("SendTextToWebViewAsync: fallback Enter sent via InputSimulator.");
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Fallback via InputSimulator failed: {ex}");
            }
        }

        // Tray-Icon + Kontextmenü
        private void InitializeTrayIcon()
        {
            trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add(Properties.Resources.ShowWindow, null, (s, e) => Reactivate());
            trayMenu.Items.Add(Properties.Resources.Reload, null, async (s, e) =>
            {
                try
                {
                    if (webView?.CoreWebView2 != null)
                        webView.CoreWebView2.Reload();
                    else
                        await LoadUrlAsync(baseUrl, bringToFront: false);
                }
                catch
                {
                    // Keine Aktion erforderlich
                }
            });

            trayMenu.Items.Add(new ToolStripSeparator());

            // Inaktivitäts-Untermenü
            var inactivityMenu = new ToolStripMenuItem(Properties.Resources.Inaktivity);
            int current = Properties.Settings.Default.InactivityTimeoutSeconds;

            foreach (var sec in inactivityOptions)
            {
                var item = new ToolStripMenuItem(string.Format(Properties.Resources._0Seconds, sec))
                {
                    Tag = sec,
                    CheckOnClick = false,
                    Checked = (current == sec)
                };

                if (sec == 0)
                {
                    item.Text = string.Format(Properties.Resources._0Deactivated, sec);
                }

                item.Click += InactivityMenuItem_Click;
                inactivityMenu.DropDownItems.Add(item);
            }

            trayMenu.Items.Add(inactivityMenu);
            trayMenu.Items.Add(new ToolStripSeparator());

            // Makros-Menü initialisieren
            InitializeMacrosMenu();
            if (macrosMenu != null) trayMenu.Items.Add(macrosMenu);
            trayMenu.Items.Add(new ToolStripSeparator());

            foreach (var page in menuPages)
            {
                var item = trayMenu.Items.Add(page.name);
                item.Click += async (s, e) => await LoadUrlAsync(baseUrl + page.url);
            }

            // Neuer Menüpunkt: Cache löschen
            trayMenu.Items.Add(Properties.Resources.ClearCache, null, async (s, e) => await ClearCacheAsync());
            trayMenu.Items.Add(new ToolStripSeparator());

            // About-Eintrag
            trayMenu.Items.Add(Properties.Resources.AboutWrok, null, (s, e) =>
            {
                using (var dlg = new AboutForm())
                {
                    dlg.ShowDialog(this);
                }
            });

            trayMenu.Items.Add(new ToolStripSeparator()); trayMenu.Items.Add(Properties.Resources.Exit, null, (s, e) => Application.Exit());

            var initialIcon = IsDarkMode() ? Properties.Resources.wrok_white : Properties.Resources.wrok_black;

            try
            {
                trayIcon = new NotifyIcon
                {
                    Text = Properties.Resources.WrokClickToOpen,
                    ContextMenuStrip = trayMenu,
                    Visible = true,
                    Icon = (Icon)initialIcon.Clone()
                };
            }
            catch
            {
                try
                {
                    trayIcon = new NotifyIcon
                    {
                        Text = Properties.Resources.WrokClickToOpen,
                        ContextMenuStrip = trayMenu,
                        Visible = true,
                        Icon = initialIcon
                    };
                }
                catch
                {
                    // Keine Aktion erforderlich
                }
            }

            try
            {
                this.Icon = (Icon)initialIcon.Clone();
            }
            catch
            {
                this.Icon = initialIcon;
            }

            if (trayIcon != null)
            {
                trayIcon.MouseClick += (s, e) =>
                {
                    if (e.Button == MouseButtons.Left)
                        Reactivate();
                };
            }

            UpdateTrayMenuInactivityState();
        }

        // Fenster sichtbar machen und ggf. laden
        private void Reactivate()
        {
            LoadWindowSettings();

            this.Show();
            this.WindowState = Properties.Settings.Default.IsMaximized ? FormWindowState.Maximized : FormWindowState.Normal;
            this.Opacity = 1.0;
            this.ShowInTaskbar = true;
            this.BringToFront();
            this.Activate();
            ResetInactivityTimer();

            try
            {
                if (webView != null)
                {
                    bool needLoad = false;

                    if (webView.CoreWebView2 == null)
                    {
                        needLoad = true;
                    }
                    else
                    {
                        try
                        {
                            var src = webView.CoreWebView2.Source?.ToString() ?? string.Empty;
                            if (!src.Contains(baseUrl, StringComparison.OrdinalIgnoreCase))
                                needLoad = true;
                        }
                        catch
                        {
                            needLoad = true;
                        }
                    }

                    if (needLoad)
                        _ = LoadUrlAsync(baseUrl, bringToFront: true);
                }
            }
            catch
            {
                // Ignorieren
            }
        }

        // URL laden
        private async Task LoadUrlAsync(string url, bool bringToFront = true)
        {
            try
            {
                if (webView == null)
                    return;

                if (webView.CoreWebView2 == null)
                {
                    try
                    {
                        await webView.EnsureCoreWebView2Async(null);
                    }
                    catch
                    {
                    }
                }

                var core = webView.CoreWebView2;

                if (core != null)
                {
                    bool online = await HasInternetConnectionAsync(attempts: 3, timeoutSeconds: 5);

                    if (online)
                    {
                        try
                        {
                            core.Navigate(url);
                        }
                        catch (Exception ex)
                        {
                            Log(ex, "core.Navigate fehlgeschlagen");
                            await ShowNoNetImageAsync();
                        }
                    }
                    else
                    {
                        await ShowNoNetImageAsync();
                    }
                }
                else
                {
                    await ShowNoNetImageAsync();
                }
            }
            catch
            {
                try
                {
                    await ShowNoNetImageAsync();
                }
                catch
                {
                }
            }

            if (!bringToFront)
                return;

            this.Show();
            this.WindowState = Properties.Settings.Default.IsMaximized ? FormWindowState.Maximized : FormWindowState.Normal;
            this.Opacity = 1.0;
            this.ShowInTaskbar = true;
            this.BringToFront();
            this.Activate();
            ResetInactivityTimer();
        }

        // Internetverbindung prüfen
        private async Task<bool> HasInternetConnectionAsync(int attempts = 2, int timeoutSeconds = 4)
        {
            try
            {
                if (!System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable())
                    return false;
            }
            catch
            {
                // Wenn der Check fehlschlägt, testen wir trotzdem weiter.
            }

            var urls = new[]
            {
                "https://clients3.google.com/generate_204",
                "http://detectportal.firefox.com/success.txt",
                "https://www.bing.com/"
            };

            var client = _httpClient;

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
                            using var resp = await client.SendAsync(
                                req,
                                HttpCompletionOption.ResponseHeadersRead,
                                cts.Token);

                            if (resp.IsSuccessStatusCode ||
                                resp.StatusCode == System.Net.HttpStatusCode.NoContent)
                                return true;
                        }
                        catch
                        {
                            using var req2 = new HttpRequestMessage(HttpMethod.Get, u);
                            using var resp2 = await client.SendAsync(
                                req2,
                                HttpCompletionOption.ResponseHeadersRead,
                                cts.Token);

                            if (resp2.IsSuccessStatusCode ||
                                resp2.StatusCode == System.Net.HttpStatusCode.NoContent)
                                return true;
                        }
                    }
                    catch
                    {
                        // Ignorieren
                    }
                }

                if (attempt + 1 < attempts)
                    await Task.Delay(300 + attempt * 200);
            }

            return false;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                SaveWindowSettings();
                e.Cancel = true;
                this.WindowState = FormWindowState.Minimized;
                this.Opacity = 0;
                this.ShowInTaskbar = false;
            }

            base.OnFormClosing(e);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            SystemEvents.UserPreferenceChanged += SystemEvents_UserPreferenceChanged;
            base.OnHandleCreated(e);

            _inputSimulator = new InputSimulator();

            try
            {
                bool ok;

                // Minimize / Toggle Hotkey: Strg+Shift+Space
                ok = RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, (uint)Keys.Space);
                if (!ok) Debug.WriteLine($"RegisterHotKey failed id={HOTKEY_ID} err={Marshal.GetLastWin32Error()}");

                ok = RegisterHotKey(this.Handle, HOTKEY_MACRO_1, MOD_CONTROL, (uint)Keys.D1);
                if (!ok) Debug.WriteLine($"RegisterHotKey failed id={HOTKEY_MACRO_1} err={Marshal.GetLastWin32Error()}");

                ok = RegisterHotKey(this.Handle, HOTKEY_MACRO_2, MOD_CONTROL, (uint)Keys.D2);
                if (!ok) Debug.WriteLine($"RegisterHotKey failed id={HOTKEY_MACRO_2} err={Marshal.GetLastWin32Error()}");

                ok = RegisterHotKey(this.Handle, HOTKEY_MACRO_3, MOD_CONTROL, (uint)Keys.D3);
                if (!ok) Debug.WriteLine($"RegisterHotKey failed id={HOTKEY_MACRO_3} err={Marshal.GetLastWin32Error()}");

                ok = RegisterHotKey(this.Handle, HOTKEY_MACRO_4, MOD_CONTROL, (uint)Keys.D4);
                if (!ok) Debug.WriteLine($"RegisterHotKey failed id={HOTKEY_MACRO_4} err={Marshal.GetLastWin32Error()}");

                ok = RegisterHotKey(this.Handle, HOTKEY_MACRO_5, MOD_CONTROL, (uint)Keys.D5);
                if (!ok) Debug.WriteLine($"RegisterHotKey failed id={HOTKEY_MACRO_5} err={Marshal.GetLastWin32Error()}");

                // Ctrl + '^' dynamisch ermitteln (VkKeyScan bereits vorhanden)
                short scan = VkKeyScan('^');
                if (scan != -1)
                {
                    byte vk = (byte)(scan & 0xFF);
                    byte shiftState = (byte)((scan >> 8) & 0xFF);
                    uint mods = MOD_CONTROL;
                    if ((shiftState & 0x01) != 0) mods |= MOD_SHIFT;
                    ok = RegisterHotKey(this.Handle, HOTKEY_MACRO_6, mods, vk);
                    if (!ok) Debug.WriteLine($"RegisterHotKey failed id={HOTKEY_MACRO_6} vk={vk} err={Marshal.GetLastWin32Error()}");
                }
                else
                {
                    ok = RegisterHotKey(this.Handle, HOTKEY_MACRO_6, MOD_CONTROL, (uint)Keys.D6);
                    if (!ok) Debug.WriteLine($"RegisterHotKey fallback failed id={HOTKEY_MACRO_6} err={Marshal.GetLastWin32Error()}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"OnHandleCreated Hotkey registration exception: {ex}");
            }

            RefreshTheme();
            EnsureTrayIconVisible();
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            // Nachrichtenfilter entfernen, falls vorhanden
            if (activityFilter != null)
            {
                try
                {
                    Application.RemoveMessageFilter(activityFilter);
                }
                catch
                {
                    // Ignorieren
                }

                activityFilter = null;
            }

            try
            {
                UnregisterHotKey(this.Handle, HOTKEY_ID);
            }
            catch
            {
            }

            // alle Hotkeys wieder abmelden
            try { UnregisterHotKey(this.Handle, HOTKEY_MACRO_1); } catch { }
            try { UnregisterHotKey(this.Handle, HOTKEY_MACRO_2); } catch { }
            try { UnregisterHotKey(this.Handle, HOTKEY_MACRO_3); } catch { }
            try { UnregisterHotKey(this.Handle, HOTKEY_MACRO_4); } catch { }
            try { UnregisterHotKey(this.Handle, HOTKEY_MACRO_5); } catch { }
            try { UnregisterHotKey(this.Handle, HOTKEY_MACRO_6); } catch { }

            _inputSimulator = null;

            SystemEvents.UserPreferenceChanged -= SystemEvents_UserPreferenceChanged;

            if (!this.RecreatingHandle)
            {
                DisposeTrayIcon();
            }

            base.OnHandleDestroyed(e);
        }

        // Theme/Settings-Nachrichten
        private const int WM_THEMECHANGED = 0x031A;
        private const int WM_SETTINGCHANGE = 0x001A;
        private const int WM_SHOWWINDOW = 0x0018;

        // === WndProc: auf UI-Thread die Makro-Ausführung starten (damit Fokus gesetzt werden kann) ===
        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case WM_HOTKEY:
                    try
                    {
                        int id = m.WParam.ToInt32();

                        if (id == HOTKEY_ID)
                        {
                            // Toggle: bei Minimiert -> reaktivieren, sonst minimieren in Tray
                            this.BeginInvoke((MethodInvoker)(() =>
                            {
                                try
                                {
                                    if (this.WindowState == FormWindowState.Minimized || this.Opacity == 0.0)
                                    {
                                        Reactivate();
                                    }
                                    else
                                    {
                                        MinimizeToTray();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Trace.WriteLine($"Hotkey toggle failed: {ex}");
                                }
                            }));
                        }
                        else
                        {
                            // Makro-Hotkey
                            this.BeginInvoke((MethodInvoker)(() => _ = PerformMacroAsync(id)));
                        }
                    }
                    catch
                    {
                        // Ignorieren
                    }
                    break;

                case WM_THEMECHANGED:
                case WM_SETTINGCHANGE:
                    try { RefreshTheme(); } catch { }
                    break;

                case WM_SHOWWINDOW:
                    try
                    {
                        this.BeginInvoke((MethodInvoker)(() =>
                        {
                            if (this.WindowState == FormWindowState.Minimized)
                                this.WindowState = FormWindowState.Normal;

                            this.Show();
                            this.BringToFront();
                            this.Activate();
                        }));
                    }
                    catch { }
                    break;
            }

            base.WndProc(ref m);
        }

        private void MainForm_Load(object? sender, EventArgs e)
        {
        }

        // Inaktivitätstimer
        private void InitializeInactivityTimer()
        {
            if (inactivityTimer != null)
            {
                try
                {
                    inactivityTimer.Stop();
                    inactivityTimer.Tick -= InactivityTimer_Tick;
                }
                catch
                {
                }

                inactivityTimer = null;
            }

            inactivityTimer = new System.Windows.Forms.Timer();
            var intervalMs = inactivityTimeout.TotalMilliseconds > 0
                ? (int)Math.Min(inactivityTimeout.TotalMilliseconds, int.MaxValue)
                : 60_000;

            inactivityTimer.Interval = intervalMs;
            inactivityTimer.Tick += InactivityTimer_Tick;

            if (inactivityEnabled && inactivityTimeout.TotalMilliseconds > 0)
            {
                lock (_activityLock)
                {
                    _lastActivity = DateTime.UtcNow;
                }

                inactivityTimer.Start();
            }
            else
            {
                inactivityTimer.Stop();
            }
        }

        private void InactivityTimer_Tick(object? sender, EventArgs e)
        {
            if (inactivityTimer == null)
                return;

            if (!inactivityEnabled)
                return;

            if (inactivityTimeout.TotalMilliseconds <= 0)
                return;

            TimeSpan elapsed;
            lock (_activityLock)
            {
                elapsed = DateTime.UtcNow - _lastActivity;
            }

            if (elapsed < inactivityTimeout)
            {
                try
                {
                    inactivityTimer.Stop();
                    inactivityTimer.Start();
                }
                catch
                {
                }

                return;
            }

            try
            {
                inactivityTimer.Stop();
            }
            catch
            {
            }

            MinimizeToTray();
        }

        private void ResetInactivityTimer()
        {
            if (!inactivityEnabled)
                return;

            if (inactivityTimer == null)
                return;

            lock (_activityLock)
            {
                _lastActivity = DateTime.UtcNow;

                try
                {
                    inactivityTimer.Stop();
                    inactivityTimer.Start();
                }
                catch
                {
                }
            }
        }

        public void SetInactivityTimeout(TimeSpan timeout)
        {
            inactivityTimeout = timeout;

            if (inactivityTimer != null)
            {
                inactivityTimer.Interval = (int)Math.Min(
                    Math.Max(1, inactivityTimeout.TotalMilliseconds),
                    int.MaxValue);

                lock (_activityLock)
                {
                    _lastActivity = DateTime.UtcNow;
                }

                ResetInactivityTimer();
            }
        }

        public void EnableInactivityTimer(bool enabled)
        {
            inactivityEnabled = enabled;

            if (inactivityTimer == null)
                return;

            if (enabled)
            {
                lock (_activityLock)
                {
                    _lastActivity = DateTime.UtcNow;
                }

                inactivityTimer.Start();
            }
            else
            {
                try
                {
                    inactivityTimer.Stop();
                }
                catch
                {
                }
            }
        }

        // Nachrichtenfilter für Aktivität
        private class ActivityMessageFilter : IMessageFilter
        {
            private readonly WeakReference<MainForm> _formRef;

            public ActivityMessageFilter(MainForm form)
            {
                _formRef = new WeakReference<MainForm>(form);
            }

            public bool PreFilterMessage(ref Message m)
            {
                const int WM_MOUSEMOVE = 0x0200;
                const int WM_LBUTTONDOWN = 0x0201;
                const int WM_RBUTTONDOWN = 0x0204;
                const int WM_MBUTTONDOWN = 0x0207;
                const int WM_MOUSEWHEEL = 0x020A;
                const int WM_KEYDOWN = 0x0100;
                const int WM_SYSKEYDOWN = 0x0104;

                if (m.Msg == WM_MOUSEMOVE ||
                    m.Msg == WM_LBUTTONDOWN ||
                    m.Msg == WM_RBUTTONDOWN ||
                    m.Msg == WM_MBUTTONDOWN ||
                    m.Msg == WM_MOUSEWHEEL ||
                    m.Msg == WM_KEYDOWN ||
                    m.Msg == WM_SYSKEYDOWN)
                {
                    if (_formRef.TryGetTarget(out var target))
                    {
                        target.ResetInactivityTimer();
                    }
                }

                return false;
            }
        }

        // Inaktivität aus Tray-Menü heraus einstellen
        private void InactivityMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is not ToolStripMenuItem clicked)
                return;

            int seconds = Convert.ToInt32(clicked.Tag ?? 0);
            Properties.Settings.Default.InactivityTimeoutSeconds = seconds;
            Properties.Settings.Default.Save();

            if (seconds > 0)
            {
                SetInactivityTimeout(TimeSpan.FromSeconds(seconds));
                EnableInactivityTimer(true);
            }
            else
            {
                EnableInactivityTimer(false);
            }

            if (clicked.OwnerItem is ToolStripMenuItem parent)
            {
                foreach (ToolStripItem item in parent.DropDownItems)
                {
                    if (item is ToolStripMenuItem menuItem)
                    {
                        menuItem.Checked = (menuItem == clicked);
                    }
                }
            }
        }

        // Offline-Seite anzeigen
        private async Task ShowNoNetImageAsync()
        {
            if (webView == null)
                return;

            if (webView.CoreWebView2 == null)
            {
                try
                {
                    await webView.EnsureCoreWebView2Async(null);
                }
                catch
                {
                }
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
        margin: 0;
        padding: 0;
        background: #202020;
        color: #ffffff;
        font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        display: flex;
        align-items: center;
        justify-content: center;
        height: 100vh;
      }}
      .wrapper {{
        text-align: center;
      }}
      img {{
        max-width: 256px;
        height: auto;
        margin-bottom: 1rem;
      }}
      h1 {{
        margin: 0 0 0.5rem 0;
        font-size: 1.2rem;
      }}
      p {{
        margin: 0;
        opacity: 0.8;
      }}
    </style>
  </head>
  <body>
    <div class=""wrapper"">
      <img src=""data:image/png;base64,{base64}"" alt=""offline"" />
      <h1>offline / keine Verbindung</h1>
      <p>Bitte überprüfe deine Internetverbindung und versuche es erneut.</p>
    </div>
  </body>
</html>
";

                var core = webView.CoreWebView2;
                core?.NavigateToString(html);
            }
            catch
            {
            }
        }

        private string BitmapToBase64(Bitmap bmp)
        {
            try
            {
                using var ms = new MemoryStream();
                bmp.Save(ms, ImageFormat.Png);
                return Convert.ToBase64String(ms.ToArray());
            }
            catch
            {
                return string.Empty;
            }
        }

        // Light/Dark-Thema
        public static bool IsDarkMode()
        {
            try
            {
                var key = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
                var value = key?.GetValue("AppsUseLightTheme");
                return value is int i && i == 0; // 0 = Dark Mode
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"IsDarkMode Fallback: {ex}");
                return true;
            }
        }

        // Zentrale Theme-Aktualisierung
        private void RefreshTheme()
        {
            try
            {
                bool dark = IsDarkMode();
                ApplyThemeIcon();      // Tray- und Fenster-Icon anpassen
                SetTitleBarDarkMode(dark); // Titelleiste umschalten
            }
            catch
            {
                // Bewusst stumm
            }
        }

        private void SystemEvents_UserPreferenceChanged(object? sender, UserPreferenceChangedEventArgs e)
        {
            // Windows feuert beim Theme-Wechsel je nach Version unterschiedliche Kategorien
            if (e.Category == UserPreferenceCategory.Color ||
                e.Category == UserPreferenceCategory.General ||
                e.Category == UserPreferenceCategory.VisualStyle)
            {
                try
                {
                    if (!this.IsDisposed)
                    {
                        this.BeginInvoke((MethodInvoker)(RefreshTheme));
                    }
                }
                catch
                {
                }
            }
        }

        // Icons dem Theme anpassen
        private void ApplyThemeIcon()
        {
            var sourceIcon = IsDarkMode() ? Properties.Resources.wrok_white : Properties.Resources.wrok_black;
            Icon newIcon;

            try
            {
                newIcon = (Icon)sourceIcon.Clone();
            }
            catch
            {
                newIcon = sourceIcon;
            }

            if (trayIcon != null)
            {
                try
                {
                    var old = trayIcon.Icon;
                    trayIcon.Visible = false;
                    trayIcon.Icon = newIcon;
                    trayIcon.Visible = true;

                    if (old != null && !ReferenceEquals(old, sourceIcon))
                    {
                        try
                        {
                            old.Dispose();
                        }
                        catch
                        {
                        }
                    }
                }
                catch
                {
                    try
                    {
                        trayIcon.Icon = newIcon;
                    }
                    catch
                    {
                    }
                }
            }

            try
            {
                this.Icon = (Icon)newIcon.Clone();
            }
            catch
            {
                this.Icon = newIcon;
            }
        }

        // DWM für Dark Titlebar
        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(
            IntPtr hwnd,
            int attr,
            ref int attrValue,
            int attrSize);

        private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;

        private void SetTitleBarDarkMode(bool enabled)
        {
            try
            {
                int val = enabled ? 1 : 0;
                int hr = DwmSetWindowAttribute(
                    this.Handle,
                    DWMWA_USE_IMMERSIVE_DARK_MODE,
                    ref val,
                    Marshal.SizeOf<int>());

                if (hr != 0)
                {
                    try
                    {
                        DwmSetWindowAttribute(
                            this.Handle,
                            DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1,
                            ref val,
                            Marshal.SizeOf<int>());
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        private void MinimizeToTray()
        {
            this.WindowState = FormWindowState.Minimized;
            this.Opacity = 0;
            this.ShowInTaskbar = false;
        }

        private void DisposeTrayIcon()
        {
            if (trayIcon == null)
                return;

            try
            {
                trayIcon.Visible = false;
            }
            catch
            {
            }

            try
            {
                var ico = trayIcon.Icon;
                trayIcon.Dispose();
                trayIcon = null;

                if (ico != null)
                {
                    try
                    {
                        ico.Dispose();
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
                trayIcon = null;
            }
        }

        private void EnsureTrayIconVisible()
        {
            try
            {
                if (trayIcon == null)
                {
                    InitializeTrayIcon();
                    ApplyThemeIcon();
                }

                if (trayIcon != null && !trayIcon.Visible)
                    trayIcon.Visible = true;
            }
            catch
            {
            }
        }

        private void UpdateTrayMenuInactivityState()
        {
            if (trayMenu == null)
                return;

            ToolStripMenuItem? inactivityMenu = null;

            foreach (ToolStripItem item in trayMenu.Items)
            {
                if (item is ToolStripMenuItem mi &&
                    mi.DropDownItems.Count > 0 &&
                    mi.Text == Properties.Resources.Inaktivity)
                {
                    inactivityMenu = mi;
                    break;
                }
            }

            if (inactivityMenu == null)
                return;

            int current = Properties.Settings.Default.InactivityTimeoutSeconds;

            foreach (ToolStripItem item in inactivityMenu.DropDownItems)
            {
                if (item is ToolStripMenuItem mi && mi.Tag is int sec)
                {
                    mi.Checked = (sec == current);
                }
            }
        }

        private void InitializeMacrosMenu()
        {
            try
            {
                if (macrosMenu == null)
                {
                    macrosMenu = new ToolStripMenuItem(Properties.Resources.Macros);
                }

                RefreshMacrosMenu();
            }
            catch
            {
                // still silent per project guidelines
            }
        }

        private void RefreshMacrosMenu()
        {
            if (macrosMenu == null)
                return;

            // Remove existing macro items (we rebuild all)
            macrosMenu.DropDownItems.Clear();

            var col = Properties.Settings.Default.Macros;
            if (col == null)
            {
                col = new StringCollection();
                Properties.Settings.Default.Macros = col;
            }

            const int requiredMacros = 6;
            bool addedDefaults = false;
            for (int i = 0; i < requiredMacros; i++)
            {
                if (i >= col.Count)
                {
                    col.Add($"Macro {i + 1}");
                    addedDefaults = true;
                }
            }

            if (addedDefaults)
                SaveMacros();

            for (int i = 0; i < requiredMacros; i++)
            {
                string text = col[i] ?? string.Empty;
                var display = string.IsNullOrWhiteSpace(text) ? $"Makro {i + 1}" : text;
                var macroItem = new ToolStripMenuItem(display)
                {
                    Tag = i
                };

                // Single left-click sends text; Shift/Ctrl+click sends +Enter; right-click edits.
                macroItem.MouseDown += async (sender, me) =>
                {
                    try
                    {
                        if (!(sender is ToolStripMenuItem tsi)) return;
                        int idx = tsi.Tag is int ii ? ii : -1;
                        if (idx < 0) return;

                        if (me.Button == MouseButtons.Left)
                        {
                            bool pressEnter = (Control.ModifierKeys & Keys.Alt) == Keys.Alt
                                              || (Control.ModifierKeys & Keys.Control) == Keys.Control;
                            try
                            {
                                var macros = Properties.Settings.Default.Macros;
                                var mtext = (macros != null && idx < macros.Count) ? (macros[idx] ?? string.Empty) : string.Empty;
                                await SendTextToWebViewAsync(mtext, pressEnter).ConfigureAwait(false);
                            }
                            catch (Exception ex)
                            {
                                Trace.WriteLine($"Macro click send failed: {ex}");
                            }
                        }
                        else if (me.Button == MouseButtons.Right)
                        {
                            try
                            {
                                EditMacroAndSave(idx);
                            }
                            catch (Exception ex)
                            {
                                Trace.WriteLine($"EditMacroAndSave failed: {ex}");
                            }
                        }
                    }
                    catch
                    {
                        // ignore per guidelines
                    }
                };

                macrosMenu.DropDownItems.Add(macroItem);
            }
        }

        private void EditMacroAndSave(int index)
        {
            try
            {
                var col = Properties.Settings.Default.Macros;
                if (col == null)
                {
                    col = new StringCollection();
                    Properties.Settings.Default.Macros = col;
                }

                string initial = string.Empty;
                if (index >= 0 && index < col.Count)
                    initial = col[index] ?? string.Empty;

                string edited = initial;
                bool ok = ShowEditMacroDialog(index >= 0 ? (Properties.Resources.EditMacro) : (Properties.Resources.NewMacro), ref edited);

                if (!ok)
                    return;

                edited = edited?.Trim() ?? string.Empty;

                if (index >= 0)
                {
                    // replace
                    col[index] = edited;
                }
                else
                {
                    // add
                    col.Add(edited);
                }

                SaveMacros();
                RefreshMacrosMenu();
            }
            catch
            {
                // ignore per guidelines
            }
        }

        private bool ShowEditMacroDialog(string title, ref string text)
        {
            using var dlg = new Form()
            {
                Text = title,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                MinimizeBox = false,
                MaximizeBox = false,
                ClientSize = new Size(480, 120)
            };

            var tb = new TextBox()
            {
                Left = 10,
                Top = 10,
                Width = dlg.ClientSize.Width - 20,
                Text = text ?? string.Empty
            };

            var btnOk = new Button()
            {
                Text = Properties.Resources.OK,
                DialogResult = DialogResult.OK,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            var btnCancel = new Button()
            {
                Text = Properties.Resources.Cancel,
                DialogResult = DialogResult.Cancel,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };

            btnOk.SetBounds(dlg.ClientSize.Width - 180, 60, 80, 26);
            btnCancel.SetBounds(dlg.ClientSize.Width - 90, 60, 80, 26);

            dlg.Controls.Add(tb);
            dlg.Controls.Add(btnOk);
            dlg.Controls.Add(btnCancel);
            dlg.AcceptButton = btnOk;
            dlg.CancelButton = btnCancel;

            var res = dlg.ShowDialog(this);
            if (res == DialogResult.OK)
            {
                text = tb.Text;
                return true;
            }

            return false;
        }

        private void SaveMacros()
        {
            try
            {
                if (Properties.Settings.Default.Macros == null)
                    Properties.Settings.Default.Macros = new StringCollection();

                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"SaveMacros failed: {ex}");
            }
        }

        // Ergänze asynchrone Methode zur Ausführung des Makros anhand der Hotkey-ID
        private async Task PerformMacroAsync(int hotkeyId)
        {
            try
            {
                var col = Properties.Settings.Default.Macros;
                if (col == null || col.Count == 0)
                    return;

                int macroIndex = -1;
                bool pressEnter = false;

                switch (hotkeyId)
                {
                    case HOTKEY_MACRO_1: macroIndex = 0; break;
                    case HOTKEY_MACRO_2: macroIndex = 1; break;
                    case HOTKEY_MACRO_3: macroIndex = 2; break;
                    case HOTKEY_MACRO_4: macroIndex = 3; break;
                    case HOTKEY_MACRO_5: macroIndex = 4; break;
                    case HOTKEY_MACRO_6: macroIndex = 5; pressEnter = true; break;
                    default: return;
                }

                if (macroIndex < 0 || macroIndex >= col.Count)
                    return;

                string text = col[macroIndex] ?? string.Empty;
                if (string.IsNullOrWhiteSpace(text))
                    return;

                await SendTextToWebViewAsync(text, pressEnter);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"PerformMacroAsync failed: {ex}");
            }
        }

        private void Log(Exception ex, string message)
        {
            try
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [Wrok] {message}: {ex}");
            }
            catch
            {
                // Logging darf keine Ausnahme werfen
            }
        }

        // Ergänze asynchrone Methode zum Löschen des WebView2-Caches gemäß Projektkonventionen
        private async Task ClearCacheAsync()
        {
            try
            {
                if (webView?.CoreWebView2 == null)
                {
                    MessageBox.Show(Properties.Resources.WebView2IsNotInitializedYet, Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Alle Browserdaten löschen (Cache, Cookies etc.)
                await webView.CoreWebView2.Profile.ClearBrowsingDataAsync();

                MessageBox.Show(Properties.Resources.CacheHasBeenDeleted, Properties.Resources.ClearCache, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Trace.WriteLine(String.Format(Properties.Resources.ClearCacheFailed + "{0}", ex));
                MessageBox.Show(
                    string.Format(Properties.Resources.ErrorWhileDeletingCache + "\n{0}", ex.Message),
                    Properties.Resources.Error,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}