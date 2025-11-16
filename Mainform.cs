using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Win32;
using System;
using System.Drawing.Imaging;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading;

namespace Wrok
{
    public partial class MainForm : Form
    {
        // Controls / resources
        private WebView2? webView;
        private NotifyIcon? trayIcon;
        private ContextMenuStrip? trayMenu;

        // Basis-URL und Menüeinträge für das Tray-Menü
        private readonly string baseUrl = "https://grok.com";
        private readonly (string name, string url)[] menuPages = new[]
        {
            (Properties.Resources.Settings, "?_s=home"),
        };

        // --- Hotkey (global) ---
        // HOTKEY_ID identifiziert die Registrierung, WM_HOTKEY ist die Window-Message.
        private const int HOTKEY_ID = 0x9000;
        private const int WM_HOTKEY = 0x0312;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;

        // P/Invoke: Registrierung von globalen Hotkeys
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        // --- Inactivity timer (automatisches Minimieren) ---
        private System.Windows.Forms.Timer? inactivityTimer;
        private TimeSpan inactivityTimeout = TimeSpan.FromSeconds(30);
        private ActivityMessageFilter? activityFilter;
        private bool inactivityEnabled = true;
        private readonly int[] inactivityOptions = new[] { 0, 30, 60, 90 };
        // ------------------------

        public MainForm()
        {
            InitializeComponent();

            // Tray initialisieren und Icon dem aktuellen Theme anpassen
            InitializeTrayIcon();
            ApplyThemeIcon();

            LoadWindowSettings();

            // Inaktivitäts-Einstellung aus den Einstellungen lesen
            var savedSeconds = Properties.Settings.Default.InactivityTimeoutSeconds;
            if (savedSeconds > 0)
            {
                inactivityTimeout = TimeSpan.FromSeconds(savedSeconds);
                inactivityEnabled = true;
            }
            else
            {
                inactivityEnabled = false;
            }

            InitializeWebView();
            InitializeInactivityTimer();

            // Lade die Seite im Hintergrund beim Start, aber ohne das Fenster sichtbar zu machen
            try { LoadUrl(baseUrl, bringToFront: false); } catch { }

            // Form-Events für WindowState/Position speichern
            this.Resize += MainForm_Resize;
            this.ResizeEnd += MainForm_ResizeEnd;
            this.Move += MainForm_Move;
        }

        private void InitializeComponent()
        {
            // ApplyThemeIcon(); // entfernt, wird bereits in Konstruktor nach InitializeTrayIcon aufgerufen

            this.Text = "Wrok";
            this.WindowState = FormWindowState.Normal;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.ShowInTaskbar = false;
            this.Visible = false;
        }

        // Lade gespeicherte Fensterposition/-größe und stelle sicher, dass die Position auf einem Bildschirm liegt.
        private void LoadWindowSettings()
        {
            var s = Properties.Settings.Default;

            if (s.WindowWidth > 0 && s.WindowHeight > 0)
            {
                this.StartPosition = FormStartPosition.Manual;
                var desired = new Rectangle(s.WindowLeft, s.WindowTop, s.WindowWidth, s.WindowHeight);

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

        // Speichert aktuelle Fenstergeometrie in den Settings (mit Schutz vor ungültigen Werten)
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
                // Absichtlich still: Settings-Write-Fehler sollen die Benutzer-Interaktion nicht blockieren.
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
        /// - WebView wird dem Form hinzugefügt.
        /// - CoreWebView2Environment mit lokalem Runtime-Ordner wird versucht, ansonsten Standard-Environment.
        /// - NavigationCompleted wird benutzt, um nach erfolgreichem Laden Aktionen auszuführen.
        /// </summary>
        private async void InitializeWebView()
        {
            webView = new WebView2
            {
                Dock = DockStyle.Fill
            };
            this.Controls.Add(webView);

            // Kurze Hotkey-ähnliche Aktion im WebView (Strg+Shift minimiert)
            webView.PreviewKeyDown += (s, e) =>
            {
                if ((ModifierKeys & (Keys.Control | Keys.Shift)) == (Keys.Control | Keys.Shift))
                {
                    this.WindowState = FormWindowState.Minimized;
                    this.ShowInTaskbar = false;
                    this.Opacity = 0;
                    e.IsInputKey = true;
                }

                ResetInactivityTimer();
            };

            // feste Runtime angeben (wenn im App-Ordner "WebView2Runtime" liegt)
            string appBase = AppDomain.CurrentDomain.BaseDirectory;
            string runtimePath = Path.Combine(appBase, "WebView2Runtime");

            string userDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Wrok",
                "WebView2Data"
            );
            Directory.CreateDirectory(userDataPath);

            CoreWebView2Environment? env = null;
            try
            {
                env = await CoreWebView2Environment.CreateAsync(
                    browserExecutableFolder: runtimePath,
                    userDataFolder: userDataPath,
                    options: null
                );
            }
            catch
            {
                // Fallback: Default-Environment benutzen
                try
                {
                    env = await CoreWebView2Environment.CreateAsync();
                }
                catch
                {
                    env = null;
                }
            }

            if (env != null)
            {
                try
                {
                    await webView.EnsureCoreWebView2Async(env);
                }
                catch
                {
                    // EnsureCoreWebView2 schlug fehl — weiter versuchen mit null
                }
            }
            else
            {
                try { await webView.EnsureCoreWebView2Async(null); } catch { }
            }

            if (webView.CoreWebView2 != null)
            {
                webView.CoreWebView2.NavigationCompleted += async (sender, args) =>
                {
                    // HINWEIS:
                    // Dieser Block ist jetzt wieder in deinem Originalzustand.
                    // D.h. das Script enthält Properties.Resources.* als JS-Code.
                    // Das compiliert in C#, weil es einfach nur Text in einem @""-String ist.

                    if (webView.CoreWebView2.Source.ToString().Contains("grok.com"))
                    {
                        string script = @"
                            // Suche nach Voice-Button und klicke ihn an
                            const voiceBtn = document.querySelector('[data-testid=""voice-mode-button""]') ||
                                             document.querySelector('button[aria-label*=""Sprachmodus""]') ||
                                             Array.from(document.querySelectorAll('button')).find(b => 
                                                 b.textContent.includes('Sprachmodus') || 
                                                 b.getAttribute('aria-label')?.includes('voice'));
                            if (voiceBtn) {
                                voiceBtn.click();
                                console.log(Properties.Resources.SpeechModeActivated);
                            } else {
                                console.warn(Properties.Resources.SpeechModeButtonNotFound);
                            }
                        ";
                        try { await webView.CoreWebView2.ExecuteScriptAsync(script); } catch { }
                    }

                    // Aktivität nach Navigation zurücksetzen
                    ResetInactivityTimer();
                };
            }
        }

        // Erzeugt das Tray-Icon + Kontextmenü
        private void InitializeTrayIcon()
        {
            trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add(Properties.Resources.ShowWindow, null, (s, e) => Reactivate());

            trayMenu.Items.Add(Properties.Resources.Reload, null, (s, e) =>
            {
                try
                {
                    webView?.CoreWebView2?.Reload();
                }
                catch { }
            });

            trayMenu.Items.Add(new ToolStripSeparator());

            var inactivityMenu = new ToolStripMenuItem(Properties.Resources.Inaktivity);
            int current = Properties.Settings.Default.InactivityTimeoutSeconds;

            foreach (var sec in inactivityOptions)
            {
                var item = new ToolStripMenuItem(String.Format(Properties.Resources._0Seconds, sec))
                {
                    Tag = sec,
                    CheckOnClick = false,
                    Checked = (current == sec && current > 0)
                };
                if (sec == 0)
                {
                    item.Text = String.Format(Properties.Resources._0Deactivated, sec);
                }
                item.Click += InactivityMenuItem_Click;
                inactivityMenu.DropDownItems.Add(item);
            }

            trayMenu.Items.Add(inactivityMenu);

            trayMenu.Items.Add(new ToolStripSeparator());

            foreach (var page in menuPages)
            {
                var item = trayMenu.Items.Add(page.name);
                item.Click += (s, e) => LoadUrl(baseUrl + page.url);
            }

            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add(Properties.Resources.Exit, null, (s, e) => Application.Exit());

            // Setze das Tray-Icon direkt bei Erstellung
            var initialIcon = IsDarkMode() ? Properties.Resources.wrok_white : Properties.Resources.wrok_black;
            trayIcon = new NotifyIcon
            {
                Text = Properties.Resources.WrokClickToOpen,
                ContextMenuStrip = trayMenu,
                Visible = true,
                Icon = (System.Drawing.Icon)initialIcon.Clone()
            };

            // Form-Icon synchronisieren
            try { this.Icon = (System.Drawing.Icon)initialIcon.Clone(); } catch { this.Icon = initialIcon; }

            trayIcon.MouseClick += (s, e) =>
            {
                if (e.Button == MouseButtons.Left)
                    Reactivate();
            };
        }

        // Macht das Fenster sichtbar und ggf. lädt den Web-Inhalt
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

            // Bei Aktivierung: lade oder bring die bereits geladene Seite in den Vordergrund.
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
                        LoadUrl(baseUrl, bringToFront: true);
                }
            }
            catch
            {
                // ignore
            }
        }

        // Lädt die URL und zeigt optional das Fenster (bringToFront).
        private async void LoadUrl(string url, bool bringToFront = true)
        {
            try
            {
                if (webView == null) return;

                if (webView.CoreWebView2 == null)
                {
                    try { await webView.EnsureCoreWebView2Async(null); }
                    catch { }
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
                        catch
                        {
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
                try { await ShowNoNetImageAsync(); } catch { }
            }

            if (!bringToFront) return;

            // Nur wenn explizit erwünscht: Fenster sichtbar machen / in den Vordergrund bringen
            this.Show();
            this.WindowState = Properties.Settings.Default.IsMaximized ? FormWindowState.Maximized : FormWindowState.Normal;
            this.Opacity = 1.0;
            this.ShowInTaskbar = true;
            this.BringToFront();
            this.Activate();

            ResetInactivityTimer();
        }

        // Prüft Internetverbindung durch Requests an bekannte Endpoints.
        private async Task<bool> HasInternetConnectionAsync(int attempts = 2, int timeoutSeconds = 4)
        {
            try
            {
                if (!System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable())
                    return false;
            }
            catch
            {
            }

            var urls = new[]
            {
                "https://clients3.google.com/generate_204",
                "http://detectportal.firefox.com/success.txt",
                "https://www.bing.com/"
            };

            using var client = new HttpClient();
            client.Timeout = TimeSpan.FromSeconds(timeoutSeconds);

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
                            using var resp = await client.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, cts.Token);
                            if (resp.IsSuccessStatusCode || resp.StatusCode == System.Net.HttpStatusCode.NoContent)
                                return true;
                        }
                        catch
                        {
                            using var req2 = new HttpRequestMessage(HttpMethod.Get, u);
                            using var resp2 = await client.SendAsync(req2, HttpCompletionOption.ResponseHeadersRead, cts.Token);
                            if (resp2.IsSuccessStatusCode || resp2.StatusCode == System.Net.HttpStatusCode.NoContent)
                                return true;
                        }
                    }
                    catch
                    {
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

            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, (uint)Keys.Space);

            if (activityFilter == null)
            {
                activityFilter = new ActivityMessageFilter(this);
                Application.AddMessageFilter(activityFilter);
            }

            // Titelleiste initial an das aktuelle Theme anpassen
            SetTitleBarDarkMode(IsDarkMode());
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID);

            if (activityFilter != null)
            {
                try { Application.RemoveMessageFilter(activityFilter); } catch { }
                activityFilter = null;
            }

            SystemEvents.UserPreferenceChanged -= SystemEvents_UserPreferenceChanged;
            base.OnHandleDestroyed(e);
        }

        // Zusätzliche Windows-Messages, die Theme-Änderungen signalisieren können.
        private const int WM_THEMECHANGED = 0x031A;
        private const int WM_SETTINGCHANGE = 0x001A;

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                MinimizeToTray();
            }
            else if (m.Msg == WM_THEMECHANGED || m.Msg == WM_SETTINGCHANGE)
            {
                // Theme/Settings changed -> Update Titelbar + Tray-Icon asynchron auf UI-Thread
                try
                {
                    this.BeginInvoke((MethodInvoker)(() =>
                    {
                        SetTitleBarDarkMode(IsDarkMode());
                        ApplyThemeIcon();
                    }));
                }
                catch
                {
                    // Fallback synchron
                    try { SetTitleBarDarkMode(IsDarkMode()); ApplyThemeIcon(); } catch { }
                }
            }

            base.WndProc(ref m);
        }

        private void MainForm_Load(object? sender, EventArgs e) { }

        // Inactivity timer: startet/stoppt und reagiert auf Tick
        private void InitializeInactivityTimer()
        {
            if (inactivityTimer != null)
            {
                try
                {
                    inactivityTimer.Stop();
                    inactivityTimer.Tick -= InactivityTimer_Tick;
                }
                catch { }
                inactivityTimer = null;
            }

            inactivityTimer = new System.Windows.Forms.Timer();
            inactivityTimer.Interval = (int)Math.Min(inactivityTimeout.TotalMilliseconds, int.MaxValue);
            inactivityTimer.Tick += InactivityTimer_Tick;

            if (inactivityEnabled && inactivityTimeout.TotalMilliseconds > 0)
            {
                inactivityTimer.Start();
            }
            else
            {
                inactivityTimer.Stop();
            }
        }

        private void InactivityTimer_Tick(object? sender, EventArgs e)
        {
            inactivityTimer?.Stop();
            MinimizeToTray();
        }

        private void ResetInactivityTimer()
        {
            if (!inactivityEnabled) return;
            if (inactivityTimer == null) return;

            inactivityTimer.Stop();
            inactivityTimer.Start();
        }

        // Minimiert das Fenster in die Tray-Leiste (sichtbar=false, Opacity=0)
        private void MinimizeToTray()
        {
            SaveWindowSettings();

            if (this.WindowState == FormWindowState.Minimized) return;

            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Opacity = 0;
        }

        public void SetInactivityTimeout(TimeSpan timeout)
        {
            inactivityTimeout = timeout;
            if (inactivityTimer != null)
            {
                inactivityTimer.Interval = (int)Math.Min(inactivityTimeout.TotalMilliseconds, int.MaxValue);
                ResetInactivityTimer();
            }
        }

        public void EnableInactivityTimer(bool enabled)
        {
            inactivityEnabled = enabled;
            if (inactivityTimer == null) return;
            if (enabled) inactivityTimer.Start(); else inactivityTimer.Stop();
        }

        // Message-Filter, der UI-Aktivität erkennt und damit den Inactivity-Timer zurücksetzt.
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
                    if (_formRef.TryGetTarget(out var form))
                    {
                        form.ResetInactivityTimer();
                    }
                }

                return false;
            }
        }

        // Menü-Handler: Inaktivitätsdauer einstellen und Settings speichern
        private void InactivityMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is not ToolStripMenuItem clicked) return;
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
                foreach (ToolStripItem it in parent.DropDownItems)
                {
                    if (it is ToolStripMenuItem mi)
                        mi.Checked = (mi == clicked);
                }
            }
        }

        // Zeigt eine einfache Offline-Seite mit eingebettetem Base64-Bitmap
        private async Task ShowNoNetImageAsync()
        {
            if (webView == null) return;

            if (webView.CoreWebView2 == null)
            {
                try { await webView.EnsureCoreWebView2Async(null); } catch { }
            }

            try
            {
                string base64 = BitmapToBase64(Properties.Resources.nonet);
                var html = $@"
<!doctype html>
<html>
<head>
<meta charset='utf-8'/>
<meta name='viewport' content='width=device-width,initial-scale=1'/>
<style>
  html, body {{ height:100%; margin:0; background:#ffffff; }}
  body {{ display:flex; align-items:center; justify-content:center; }}
  img {{ max-width:100%; max-height:100%; object-fit:contain; }}
  .msg {{
      position: absolute;
      bottom: 2rem;
      left: 0;
      right: 0;
      text-align: center;
      font-family: sans-serif;
      font-size: 0.9rem;
      color: #444;
      opacity: 0.7;
  }}
</style>
</head>
<body>
  <img src='data:image/png;base64,{base64}' alt='no network' />
  <div class='msg'>offline / keine Verbindung</div>
</body>
</html>";
                var core = webView.CoreWebView2;
                if (core != null)
                    core.NavigateToString(html);
            }
            catch
            {
            }
        }

        // Hilfsfunktion: Bitmap in Base64 kodieren (für Embedded-HTML)
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

        // Liefert true, wenn Windows auf Apps "Light Theme" verwendet, false = Dark
        public static bool IsDarkMode()
        {
            try
            {
                var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
                var value = key?.GetValue("AppsUseLightTheme");
                return value is int i && i == 0; // 0 = Dark Mode, 1 = Light Mode
            }
            catch
            {
                return true; // Fallback: Dark Mode (häufiger Standard)
            }
        }

        private void SystemEvents_UserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            if (e.Category == UserPreferenceCategory.Color)
            {
                this.Invoke((MethodInvoker)ApplyThemeIcon); // UI-Thread
            }
        }

        // Aktualisiert Tray-Icon und Form-Icon passend zum aktuellen Theme.
        // Wichtig: NotifyIcon kann Icon-Caching haben, deshalb wird ein Clone verwendet.
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

                    // Windows dazu bringen, die Änderung zu übernehmen
                    trayIcon.Visible = false;
                    trayIcon.Icon = newIcon;
                    trayIcon.Visible = true;

                    // Altes Icon freigeben, wenn es nicht das Ressourcen-Icon ist
                    if (old != null && !ReferenceEquals(old, sourceIcon))
                    {
                        try { old.Dispose(); } catch { }
                    }
                }
                catch
                {
                    // Fallback: wenigstens das Icon setzen
                    try { trayIcon.Icon = newIcon; } catch { }
                }
            }

            // Form-Icon setzen (eigene Instanz)
            try
            {
                this.Icon = (Icon)newIcon.Clone();
            }
            catch
            {
                this.Icon = newIcon;
            }
        }

        // Ergänze neben den anderen DllImport-Deklarationen
        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        // DWM-Attribute für immersive dark titlebar (verschiedene Windows-Builds)
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20; // neuer Windows 10/11 Wert
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19; // älterer Fall

        // Setzt/entfernt den "immersive dark mode" für die native Titelleiste (wenn möglich).
        private void SetTitleBarDarkMode(bool enabled)
        {
            try
            {
                int val = enabled ? 1 : 0;
                // Versuch mit neuem Attribut, falls Fehler dann mit älterem
                int hr = DwmSetWindowAttribute(this.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref val, Marshal.SizeOf<int>());
                if (hr != 0)
                {
                    try { DwmSetWindowAttribute(this.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1, ref val, Marshal.SizeOf<int>()); } catch { }
                }
            }
            catch { }
        }
    }
}
