using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

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

        // Ergänzte Felder für Aktivitäts-Tracking
        private DateTime _lastActivity = DateTime.UtcNow;
        private readonly object _activityLock = new object();

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
            RefreshTheme();   // Zentrale Theme-Aktualisierung

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

            InitializeWebView();
            InitializeInactivityTimer();

            // Seite im Hintergrund laden
            try
            {
                LoadUrl(baseUrl, bringToFront: false);
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

                if (this.WindowState == FormWindowState.Maximized ||
                    this.WindowState == FormWindowState.Minimized)
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
            if (this.WindowState == FormWindowState.Minimized ||
                this.WindowState == FormWindowState.Maximized)
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
        private async void InitializeWebView()
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
                    string script = @"
                        (function() {
                            const resetActivity = () => {
                                window.chrome.webview.postMessage('resetActivity');
                            };

                            ['mousemove', 'mousedown', 'keydown', 'scroll', 'touchstart'].forEach(event => {
                                window.addEventListener(event, resetActivity, { passive: true });
                            });
                        })();
                    ";
                    try
                    {
                        await webView.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync(script);
                    }
                    catch
                    {
                        // Keine Aktion erforderlich
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

        // Tray-Icon + Kontextmenü
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

            foreach (var page in menuPages)
            {
                var item = trayMenu.Items.Add(page.name);
                item.Click += (s, e) => LoadUrl(baseUrl + page.url);
            }
            // Neuer Menüpunkt: Cache löschen
            trayMenu.Items.Add(Properties.Resources.ClearCache, null, async (s, e) => await ClearCacheAsync());
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add(Properties.Resources.Exit, null, (s, e) => Application.Exit());

            var initialIcon = IsDarkMode()
                ? Properties.Resources.wrok_white
                : Properties.Resources.wrok_black;

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
            this.WindowState = Properties.Settings.Default.IsMaximized
                ? FormWindowState.Maximized
                : FormWindowState.Normal;
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
                        catch (Exception)
                        {
                            needLoad = true;
                        }
                    }
                    if (needLoad)
                        LoadUrl(baseUrl, bringToFront: true);
                }
            }
            catch (Exception)
            {
            }
        }

        // URL laden
        private async void LoadUrl(string url, bool bringToFront = true)
        {
            try
            {
                if (webView == null) return;

                if (webView.CoreWebView2 == null)
                {
                    try
                    {
                        await webView.EnsureCoreWebView2Async(null);
                    }
                    catch (Exception)
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
            catch (Exception)
            {
                try
                {
                    await ShowNoNetImageAsync();
                }
                catch (Exception)
                {
                }
            }

            if (!bringToFront) return;

            this.Show();
            this.WindowState = Properties.Settings.Default.IsMaximized
                ? FormWindowState.Maximized
                : FormWindowState.Normal;
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
            catch (Exception)
            {
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
                    catch (Exception)
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

            try
            {
                bool registered = RegisterHotKey(
                    this.Handle,
                    HOTKEY_ID,
                    MOD_CONTROL | MOD_SHIFT,
                    (uint)Keys.Space);
                if (!registered)
                {
                    int err = Marshal.GetLastWin32Error();
                }
            }
            catch (Exception)
            {
            }

            RefreshTheme(); // Theme auch nach Handle-Erzeugung anpassen
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
            catch (Exception)
            {
            }

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

        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case WM_THEMECHANGED:
                case WM_SETTINGCHANGE:
                    try
                    {
                        RefreshTheme();
                    }
                    catch
                    {
                    }
                    break;
                case WM_SHOWWINDOW:
                    try
                    {
                        this.BeginInvoke((MethodInvoker)(() =>
                        {
                            if (this.WindowState == FormWindowState.Minimized)
                            {
                                this.WindowState = FormWindowState.Normal;
                            }
                            this.Show();
                            this.BringToFront();
                            this.Activate();
                        }));
                    }
                    catch
                    {
                    }
                    break;
            }
            base.WndProc(ref m);
        }

        private void MainForm_Load(object? sender, EventArgs e) { }

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
                lock (_activityLock) { _lastActivity = DateTime.UtcNow; }
                inactivityTimer.Start();
            }
            else
            {
                inactivityTimer.Stop();
            }
        }

        private void InactivityTimer_Tick(object? sender, EventArgs e)
        {
            if (inactivityTimer == null) return;
            if (!inactivityEnabled) return;
            if (inactivityTimeout.TotalMilliseconds <= 0) return;

            lock (_activityLock)
            {
                var elapsed = DateTime.UtcNow - _lastActivity;
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
            }

            lock (_activityLock)
            {
                var elapsed = DateTime.UtcNow - _lastActivity;
                if (elapsed >= inactivityTimeout)
                {
                    try
                    {
                        inactivityTimer.Stop();
                    }
                    catch
                    {
                    }
                    MinimizeToTray();
                }
            }
        }

        private void ResetInactivityTimer()
        {
            if (!inactivityEnabled) return;
            if (inactivityTimer == null) return;

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
                lock (_activityLock) { _lastActivity = DateTime.UtcNow; }
                ResetInactivityTimer();
            }
        }

        public void EnableInactivityTimer(bool enabled)
        {
            inactivityEnabled = enabled;
            if (inactivityTimer == null) return;

            if (enabled)
            {
                lock (_activityLock) { _lastActivity = DateTime.UtcNow; }
                inactivityTimer.Start();
            }
            else
            {
                try
                {
                    inactivityTimer.Stop();
                }
                catch (Exception)
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
                    if (_formRef.TryGetTarget(out var targetInstanceInner))
                    {
                        targetInstanceInner.ResetInactivityTimer();
                    }
                }

                return false;
            }
        }

        // Inaktivität aus Tray-Menü heraus einstellen
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
            if (webView == null) return;

            if (webView.CoreWebView2 == null)
            {
                try
                {
                    await webView.EnsureCoreWebView2Async(null);
                }
                catch (Exception)
                {
                }
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
            catch (Exception)
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
            catch (Exception)
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
                ApplyThemeIcon();          // Tray- und Fenster-Icon anpassen
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
            var sourceIcon = IsDarkMode()
                ? Properties.Resources.wrok_white
                : Properties.Resources.wrok_black;

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
                        try { old.Dispose(); }
                        catch (Exception) { }
                    }
                }
                catch (Exception)
                {
                    try { trayIcon.Icon = newIcon; }
                    catch (Exception) { }
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
                    catch (Exception)
                    {
                    }
                }
            }
            catch (Exception)
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
            if (trayIcon == null) return;

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
                    try { ico.Dispose(); }
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
            catch (Exception)
            {
            }
        }

        private void UpdateTrayMenuInactivityState()
        {
            if (trayMenu == null) return;

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

            if (inactivityMenu == null) return;

            int current = Properties.Settings.Default.InactivityTimeoutSeconds;

            foreach (ToolStripItem item in inactivityMenu.DropDownItems)
            {
                if (item is ToolStripMenuItem mi && mi.Tag is int sec)
                {
                    mi.Checked = (sec == current);
                }
            }
        }

        private async Task ClearCacheAsync()
        {
            if (webView?.CoreWebView2?.Profile == null)
            {
                MessageBox.Show(Properties.Resources.WebView2NotInitializedYet, Properties.Resources.DeleteCache,
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // TaskDialog für Windows 10+
            var page = new TaskDialogPage
            {
                Caption = "Wrok",
                Heading = Properties.Resources.ClearBrowsingData,
                Text = Properties.Resources.ChooseWhatToBeDeleted,
                Icon = TaskDialogIcon.Information,
                AllowCancel = true,
                DefaultButton = TaskDialogButton.Yes
            };

            var btnCache = new TaskDialogButton(Properties.Resources.ClearCacheOnly)
            {
                Tag = "cache"
            };
            var btnAll = new TaskDialogButton(Properties.Resources.DeleteAll)
            {
                Tag = "all"
            };
            var btnCancel = TaskDialogButton.Cancel;

            page.Buttons.Add(btnCache);
            page.Buttons.Add(btnAll);
            page.Buttons.Add(btnCancel);

            page.DefaultButton = btnCache;  // Fokus auf "Nur Cache löschen"

            var result = TaskDialog.ShowDialog(this.Handle, page);

            if (result == btnCancel || result == TaskDialogButton.Cancel)
                return;

            try
            {
                if (result.Tag?.ToString() == "cache")
                {
                    var cacheOnly = CoreWebView2BrowsingDataKinds.DiskCache |
                                    CoreWebView2BrowsingDataKinds.CacheStorage;

                    await webView.CoreWebView2.Profile.ClearBrowsingDataAsync(cacheOnly);

                    TaskDialog.ShowDialog(this.Handle, new TaskDialogPage
                    {
                        Caption = "Wrok",
                        Heading = Properties.Resources.CacheCleared,
                        Text = Properties.Resources.PicturesScriptsAndOtherDataDeletedNYouAreStillLoggedIn,
                        Icon = TaskDialogIcon.Information,
                        Buttons = { TaskDialogButton.OK }
                    });
                }
                else // Alles löschen
                {
                    await webView.CoreWebView2.Profile.ClearBrowsingDataAsync();

                    TaskDialog.ShowDialog(this.Handle, new TaskDialogPage
                    {
                        Caption = "Wrok",
                        Heading = Properties.Resources.AllDataDeleted,
                        Text = Properties.Resources.CookiesLoginDataAndSettingsDeletedNYouAreLoggedOut,
                        Icon = TaskDialogIcon.Warning,
                        Buttons = { TaskDialogButton.OK }
                    });
                }

                webView.CoreWebView2?.Reload();
            }
            catch (Exception ex)
            {
                TaskDialog.ShowDialog(this.Handle, new TaskDialogPage
                {
                    Caption = "Wrok",
                    Heading = Properties.Resources.Error,
                    Text = String.Format(Properties.Resources.ErrorWhileDeletingCacheN0, ex.Message),
                    Icon = TaskDialogIcon.Error,
                    Buttons = { TaskDialogButton.OK }
                });
                Log(ex, Properties.Resources.ClearCacheAsyncFailed);
            }
        }
        private void Log(Exception ex, string message)
        {
            try
            {
                Trace.WriteLine($"{message}: {ex}");
            }
            catch
            {
            }
        }
    }
}