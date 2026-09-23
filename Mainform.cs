using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Wrok
{
    public partial class MainForm : Form
    {
        // ------------------------------------------------------------------
        // Konstanten
        // ------------------------------------------------------------------

        private const int HOTKEY_ID       = 0x9000;
        private const int HOTKEY_ID_IMAGE = 0x9001;   // Strg+Ö
        private const int WM_HOTKEY    = 0x0312;
        private const uint MOD_CONTROL = 0x0002;

        // F1–F10 (ohne Modifier) feuern per lokaler Tastenbehandlung (siehe
        // MainForm.Macros.HandleFKeyOverride) die Makros des aktiven Profils,
        // Strg+Umschalt+F1..F{ProfileCount} wechselt das aktive Profil. Kein
        // globaler Hotkey mehr nötig - alles greift nur, wenn Wrok den
        // Tastaturfokus hat (siehe ProcessCmdKey unten und
        // WebViewManager.OnCoreWebView2InitializationCompleted für den Fall,
        // dass das WebView2-Control selbst fokussiert ist).

        private const int WM_THEMECHANGED    = 0x031A;
        private const int WM_SETTINGCHANGE   = 0x001A;
        private const int WM_SHOWWINDOW      = 0x0018;
        private const int WM_QUERYENDSESSION = 0x0011;
        private const int WM_ENDSESSION      = 0x0016;

        private readonly string baseUrl = "https://grok.com/";
        private readonly (string name, string url)[] menuPages = new[]
        {
            (Properties.Resources.MenuGrok, "?_s=home"),
        };

        private readonly int[] inactivityOptions = new[] { 0, 30, 60, 90 };

        // ------------------------------------------------------------------
        // Felder
        // ------------------------------------------------------------------

        private WebView2?          _webView;
        private WebViewManager?    _webViewManager;
        private MacroManager       _macroManager = new();

        private NotifyIcon?        trayIcon;
        private ContextMenuStrip?  trayMenu;
        private ToolStripMenuItem? macrosMenu;
        private ToolStripMenuItem? _inactivityMenu;

        // Nebeneinander-Anordnung (Medium links, Wrok rechts)
        private bool  _suppressWindowSave;
        private Form? _mediaViewer;   // es gibt immer höchstens einen (Bild ODER Video)
        private bool  _mediaViewerWasPinned;   // TopMost-Zustand über ein Verstecken hinweg

        private InactivityWatcher? _inactivity;
        private ThemeManager?      _theme;

        // Windows meldet Abmelden/Herunterfahren und der Installer-Restart-Manager
        // ein bevorstehendes Sitzungsende ueber WM_QUERYENDSESSION - immer VOR dem
        // eigentlichen Schliessen. In diesem Fall soll Wrok wirklich beenden statt
        // in den Tray zu minimieren (sonst blockiert es ein stilles Upgrade).
        private bool _sessionEnding;

        // URL der neueren Version, sobald die Update-Prüfung fündig wird.
        private string? _pendingUpdateUrl;
        private ToolStripItem? _updateMenuItem;

        private RateLimitManager?      _rateLimitManager;
        private ToolStripMenuItem?     _rateLimitMenu;

        // ------------------------------------------------------------------
        // P/Invoke
        // ------------------------------------------------------------------

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        // ------------------------------------------------------------------
        // Konstruktor
        // ------------------------------------------------------------------

        public MainForm()
        {
            InitializeComponent();

            _inactivity = new InactivityWatcher(this, MinimizeToTray, () => _mediaViewer);
            _theme      = new ThemeManager(this, () => trayIcon);

            InitializeTrayIcon();
            _theme.Refresh();
            LoadWindowSettings();
            _inactivity.LoadSettings();

            _ = InitializeWebViewAsync();
            _inactivity.Start();
            _ = InitializeRateLimitManagerAsync();

            try { _ = LoadUrlAsync(baseUrl, bringToFront: false); }
            catch (Exception ex) { Log(ex, "Initialer LoadUrlAsync fehlgeschlagen"); }

            this.Resize   += (s, e) => { if (this.WindowState != FormWindowState.Normal) SaveWindowSettings(); };
            this.ResizeEnd += (s, e) => { if (this.WindowState == FormWindowState.Normal) SaveWindowSettings(); };
            this.Move      += (s, e) => { if (this.WindowState == FormWindowState.Normal) SaveWindowSettings(); };

            UpdateTrayMenuInactivityState();

            // Nach dem Aufbau: still im Hintergrund nach einer neueren Version sehen.
            StartUpdateCheck();
        }

        // ------------------------------------------------------------------
        // InitializeComponent
        // ------------------------------------------------------------------

        private void InitializeComponent()
        {
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            this.Text             = version != null ? $"Wrok {version.Major}.{version.Minor}.{version.Build}" : "Wrok";
            this.WindowState      = FormWindowState.Normal;
            this.StartPosition    = FormStartPosition.CenterScreen;
            this.FormBorderStyle  = FormBorderStyle.Sizable;
            this.ShowInTaskbar    = false;
            this.Visible          = false;
        }

        // ------------------------------------------------------------------
        // WebView-Initialisierung
        // ------------------------------------------------------------------

        private async Task InitializeWebViewAsync()
        {
            _webView = new WebView2 { Dock = DockStyle.Fill };
            this.Controls.Add(_webView);

            _webViewManager = new WebViewManager(_webView, this, () => _inactivity?.Reset("WebView-Ereignis"));
            await _webViewManager.InitializeAsync();
        }

        // ------------------------------------------------------------------
        // Navigation
        // ------------------------------------------------------------------

        private async Task LoadUrlAsync(string url, bool bringToFront = true)
        {
            try
            {
                if (_webViewManager != null)
                    await _webViewManager.NavigateAsync(url);
            }
            catch (Exception ex)
            {
                Log(ex, "LoadUrlAsync fehlgeschlagen");
            }

            if (!bringToFront) return;

            this.Show();
            this.WindowState = Properties.Settings.Default.IsMaximized
                ? FormWindowState.Maximized : FormWindowState.Normal;
            this.Opacity      = 1.0;
            this.ShowInTaskbar = true;
            this.BringToFront();
            this.Activate();
            _inactivity?.Reset("Fenster wieder angezeigt");
        }

        // ------------------------------------------------------------------
        // F-Tasten-Override (Makros) - lokale Tastenbehandlung
        // ------------------------------------------------------------------

        /// <summary>
        /// Greift, wenn ein normales WinForms-Control (nicht der WebView2-Inhalt
        /// selbst) den Fokus hat. Für den Fall, dass die Grok-Seite im WebView2
        /// fokussiert ist, ruft WebViewManager denselben Weg über
        /// HandleFKeyOverride direkt auf (CoreWebView2.AcceleratorKeyPressed).
        /// </summary>
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            Keys keyCode = keyData & Keys.KeyCode;
            if (keyCode >= Keys.F1 && keyCode <= Keys.F12)
            {
                bool ctrl  = (keyData & Keys.Control) == Keys.Control;
                bool shift = (keyData & Keys.Shift)   == Keys.Shift;
                bool alt   = (keyData & Keys.Alt)     == Keys.Alt;
                if (HandleFKeyOverride(keyCode, ctrl, shift, alt))
                    return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // ------------------------------------------------------------------
        // Logging
        // ------------------------------------------------------------------

        private static void Log(Exception ex, string message) =>
            Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MainForm] {message}: {ex}");
    }
}
