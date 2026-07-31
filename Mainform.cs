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

        /// <summary>
        /// Strg-Makro-Tasten des AKTIVEN Profils: Strg+1..Strg+9, Strg+0. Index i
        /// (0-basiert) → Makro i des aktiven Profils, registriert unter
        /// MacroManager.CtrlMacroBase+i. Immer aktiv, auf jeder Tastatur bedienbar.
        /// </summary>
        private static readonly Keys[] CtrlMacroKeys =
        {
            Keys.D1, Keys.D2, Keys.D3, Keys.D4, Keys.D5,
            Keys.D6, Keys.D7, Keys.D8, Keys.D9, Keys.D0
        };

        /// <summary>
        /// Feste Tastenzuordnung für den Nummernblock-Modus: Index entspricht dem
        /// Numpad-Slot (siehe MacroManager.NumpadMacroBase). Abgebildet werden nur
        /// die ersten fünf Makros jedes Profils: Profil 1 = NumPad0-4, Profil 2 =
        /// NumPad5-9, Profil 3 = die vier Rechenzeichen + Dezimalpunkt. Alle ohne
        /// Modifier, damit sie einhändig gehen – optional, siehe Numpad-Makro-Modus.
        /// </summary>
        private static readonly Keys[] NumpadMacroKeys =
        {
            // Profil 1: erste 5 Makros → Ziffern 0-4
            Keys.NumPad0, Keys.NumPad1, Keys.NumPad2, Keys.NumPad3, Keys.NumPad4,
            // Profil 2: erste 5 Makros → Ziffern 5-9
            Keys.NumPad5, Keys.NumPad6, Keys.NumPad7, Keys.NumPad8, Keys.NumPad9,
            // Profil 3: erste 5 Makros → / * - + ,
            Keys.Divide, Keys.Multiply, Keys.Subtract, Keys.Add, Keys.Decimal
        };

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
        private ToolStripMenuItem? grokAccountsMenu;
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

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
        private const byte VK_NUMLOCK        = 0x90;
        private const uint KEYEVENTF_KEYUP   = 0x0002;

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

            // Fuer den E-Mail-Prefill beim Grok-Kontowechsel: direkter Zugriff auf
            // CoreWebView2 wie beim bestehenden Reload-Tray-Eintrag, WebViewManager
            // kapselt keine Navigations-Events nach aussen.
            if (_webView.CoreWebView2 != null)
                _webView.CoreWebView2.NavigationCompleted += OnGrokSignInNavigationCompleted;
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
        // Fenster-Zustand
        // ------------------------------------------------------------------

        private void LoadWindowSettings()
        {
            var s = Properties.Settings.Default;
            bool hasValidSize      = s.WindowWidth > 0 && s.WindowHeight > 0;
            bool hasExplicitPos    = s.WindowLeft != 0 || s.WindowTop != 0;

            if (hasValidSize && hasExplicitPos)
            {
                var desired = new Rectangle(s.WindowLeft, s.WindowTop, s.WindowWidth, s.WindowHeight);
                bool onScreen = Screen.AllScreens.Any(scr => scr.WorkingArea.IntersectsWith(desired));
                if (onScreen)
                {
                    this.StartPosition = FormStartPosition.Manual;
                    this.Bounds = desired;
                }
                else
                {
                    this.Size = new Size(s.WindowWidth, s.WindowHeight);
                    this.StartPosition = FormStartPosition.CenterScreen;
                    try { this.CenterToScreen(); } catch { }
                }
            }
            else if (hasValidSize)
            {
                this.Size = new Size(s.WindowWidth, s.WindowHeight);
                this.StartPosition = FormStartPosition.CenterScreen;
                try { this.CenterToScreen(); } catch { }
            }

            if (s.IsMaximized)
                this.WindowState = FormWindowState.Maximized;
        }

        private void SaveWindowSettings()
        {
            // Während der Nebeneinander-Anordnung nicht speichern – sonst würde
            // die vom Nutzer gewählte Fenstergeometrie dauerhaft überschrieben.
            if (_suppressWindowSave) return;

            try
            {
                var s = Properties.Settings.Default;
                var bounds = (this.WindowState == FormWindowState.Maximized ||
                              this.WindowState == FormWindowState.Minimized)
                    ? this.RestoreBounds : this.Bounds;

                s.WindowLeft   = bounds.Left;
                s.WindowTop    = bounds.Top;
                s.WindowWidth  = Math.Max(100, bounds.Width);
                s.WindowHeight = Math.Max(100, bounds.Height);
                s.IsMaximized  = this.WindowState == FormWindowState.Maximized;
                s.Save();
            }
            catch { }
        }

        private void Reactivate()
        {
            LoadWindowSettings();
            if (this.StartPosition == FormStartPosition.CenterScreen)
                try { this.CenterToScreen(); } catch { }

            this.Show();
            this.WindowState   = Properties.Settings.Default.IsMaximized
                ? FormWindowState.Maximized : FormWindowState.Normal;
            this.Opacity       = 1.0;
            this.ShowInTaskbar = true;
            this.BringToFront();
            this.Activate();
            RestoreMediaViewer();
            _inactivity?.Reset("Fenster reaktiviert");

            try
            {
                if (_webView != null)
                {
                    bool needLoad = _webView.CoreWebView2 == null ||
                                   !(_webView.CoreWebView2.Source?.Contains(baseUrl, StringComparison.OrdinalIgnoreCase) ?? false);
                    if (needLoad) _ = LoadUrlAsync(baseUrl, bringToFront: true);
                }
            }
            catch { }
        }

        private void MinimizeToTray()
        {
            HideMediaViewer();
            this.WindowState   = FormWindowState.Minimized;
            this.Opacity       = 0;
            this.ShowInTaskbar = false;
        }

        /// <summary>
        /// Blendet ein offenes Medienfenster mit aus.
        ///
        /// Wichtig fuer die Boss-Taste: Der Betrachter ist ein eigenstaendiges
        /// Top-Level-Fenster. Ohne diesen Schritt verschwaende nur das Hauptfenster,
        /// waehrend das Bild sichtbar stehen bliebe – angeheftet sogar ueber allem
        /// anderen. Genau dann soll die Funktion aber greifen.
        ///
        /// TopMost wird dabei aufgehoben und gemerkt, weil ein verstecktes
        /// TopMost-Fenster beim Wiederauftauchen sonst unvermittelt vor allen
        /// anderen Fenstern erschiene.
        /// </summary>
        private void HideMediaViewer()
        {
            try
            {
                if (_mediaViewer is { IsDisposed: false, Visible: true } viewer)
                {
                    _mediaViewerWasPinned = viewer.TopMost;
                    viewer.TopMost = false;
                    // Hide() nimmt das Fenster zugleich aus Taskleiste und Alt+Tab.
                    // ShowInTaskbar wird bewusst nicht angefasst: Das erzwingt eine
                    // Neuerzeugung des Fensterhandles und wuerde beim Video-Betrachter
                    // das eingebettete WebView2 mitreissen.
                    viewer.Hide();
                }
            }
            catch (Exception ex) { Log(ex, "Medienfenster ausblenden fehlgeschlagen"); }
        }

        /// <summary>Holt ein zuvor ausgeblendetes Medienfenster samt Anheftung zurueck.</summary>
        private void RestoreMediaViewer()
        {
            try
            {
                if (_mediaViewer is { IsDisposed: false, Visible: false } viewer)
                {
                    viewer.Show();
                    viewer.TopMost = _mediaViewerWasPinned;
                }
            }
            catch (Exception ex) { Log(ex, "Medienfenster wiederherstellen fehlgeschlagen"); }
        }

        // ------------------------------------------------------------------
        // Tray-Icon
        // ------------------------------------------------------------------

        private void InitializeTrayIcon()
        {
            trayMenu = new ContextMenuStrip { ShowItemToolTips = true };

            trayMenu.Items.Add(Properties.Resources.ShowWindow, null, (s, e) => Reactivate());
            trayMenu.Items.Add(Properties.Resources.Reload, null, async (s, e) =>
            {
                try
                {
                    if (_webView?.CoreWebView2 != null) _webView.CoreWebView2.Reload();
                    else await LoadUrlAsync(baseUrl, bringToFront: false);
                }
                catch { }
            });
            trayMenu.Items.Add(new ToolStripSeparator());

            // ---------- Makros (oberste Ebene: meistgenutzte Funktion) ----------
            InitializeMacrosMenu();
            if (macrosMenu != null) trayMenu.Items.Add(macrosMenu);

            // ---------- Grok-Konto wechseln (ebenfalls oberste Ebene) ----------
            InitializeGrokAccountsMenu();
            if (grokAccountsMenu != null) trayMenu.Items.Add(grokAccountsMenu);

            trayMenu.Items.Add(new ToolStripSeparator());

            // ---------- Einstellungen ----------
            var settingsMenu = new ToolStripMenuItem(Properties.Resources.Settings);

            //   Einstellungen → App
            var appMenu = new ToolStripMenuItem(Properties.Resources.MenuApp);

            var inactivityMenu = new ToolStripMenuItem(Properties.Resources.Inaktivity);
            _inactivityMenu = inactivityMenu;   // Referenz merken: liegt jetzt verschachtelt
            int current = Properties.Settings.Default.InactivityTimeoutSeconds;
            foreach (var sec in inactivityOptions)
            {
                var item = new ToolStripMenuItem(sec == 0
                    ? string.Format(Properties.Resources._0Deactivated, sec)
                    : string.Format(Properties.Resources._0Seconds, sec))
                {
                    Tag           = sec,
                    CheckOnClick  = false,
                    Checked       = current == sec
                };
                item.Click += InactivityMenuItem_Click;
                inactivityMenu.DropDownItems.Add(item);
            }
            appMenu.DropDownItems.Add(inactivityMenu);

            var macroIoMenu = new ToolStripMenuItem(Properties.Resources.MacrosImportExport);
            macroIoMenu.DropDownItems.Add(Properties.Resources.MacrosExportActive, null, (s, e) => ExportMacros(allProfiles: false));
            macroIoMenu.DropDownItems.Add(Properties.Resources.MacrosExportAll,    null, (s, e) => ExportMacros(allProfiles: true));
            macroIoMenu.DropDownItems.Add(new ToolStripSeparator());
            macroIoMenu.DropDownItems.Add(Properties.Resources.MacrosImport,       null, (s, e) => ImportMacros());
            appMenu.DropDownItems.Add(macroIoMenu);

            var autostartItem = new ToolStripMenuItem(Properties.Resources.StartWithWindows)
            {
                CheckOnClick = false,
                Checked      = AutostartManager.IsEnabled()
            };
            autostartItem.Click += (s, e) =>
            {
                bool desired = !autostartItem.Checked;
                if (AutostartManager.SetEnabled(desired))
                    autostartItem.Checked = AutostartManager.IsEnabled();
            };
            appMenu.DropDownItems.Add(autostartItem);

            // Numpad-Makro-Modus: alle 15 Makros ohne Modifier per Nummernblock,
            // dafür ist der Block waehrend des Modus fuer normale Eingaben blockiert.
            var numpadMacroItem = new ToolStripMenuItem(Properties.Resources.NumpadMacroMode)
            {
                CheckOnClick = false,
                Checked      = Properties.Settings.Default.NumpadMacroModeEnabled
            };
            numpadMacroItem.Click += (s, e) =>
            {
                bool desired = !numpadMacroItem.Checked;
                SetNumpadMacroMode(desired);
                numpadMacroItem.Checked = desired;
                RefreshMacrosMenu();
            };
            appMenu.DropDownItems.Add(numpadMacroItem);

            // Chromiums eigener Passwort-Manager fuers Grok-Login (siehe
            // WebViewManager.ConfigureCore) - wirkt sofort, kein Neustart noetig.
            var passwordAutosaveItem = new ToolStripMenuItem(Properties.Resources.PasswordAutosaveMenuItem)
            {
                CheckOnClick = false,
                Checked      = Properties.Settings.Default.PasswordAutosaveEnabled
            };
            passwordAutosaveItem.Click += (s, e) =>
            {
                bool desired = !passwordAutosaveItem.Checked;
                Properties.Settings.Default.PasswordAutosaveEnabled = desired;
                Properties.Settings.Default.Save();
                if (_webView?.CoreWebView2 != null)
                    _webView.CoreWebView2.Settings.IsPasswordAutosaveEnabled = desired;
                passwordAutosaveItem.Checked = desired;
            };
            appMenu.DropDownItems.Add(passwordAutosaveItem);

            var proxyItem = new ToolStripMenuItem(Properties.Resources.ProxyMenuItem);
            proxyItem.Click += (s, e) => ShowProxyDialog();
            appMenu.DropDownItems.Add(proxyItem);

            // Update-Prüfung: kontaktiert beim Start GitHub, deshalb abschaltbar.
            var updateCheckItem = new ToolStripMenuItem(Properties.Resources.CheckForUpdatesMenuItem)
            {
                CheckOnClick = false,
                Checked      = Properties.Settings.Default.CheckForUpdates
            };
            updateCheckItem.ToolTipText = Properties.Resources.CheckForUpdatesHint;
            updateCheckItem.Click += (s, e) =>
            {
                bool desired = !updateCheckItem.Checked;
                Properties.Settings.Default.CheckForUpdates = desired;
                Properties.Settings.Default.Save();
                updateCheckItem.Checked = desired;
            };
            appMenu.DropDownItems.Add(updateCheckItem);

            settingsMenu.DropDownItems.Add(appMenu);

            //   Einstellungen → Grok (öffnet Groks eigene Einstellungsseite)
            foreach (var page in menuPages)
            {
                var item = new ToolStripMenuItem(page.name);
                item.Click += async (s, e) => await LoadUrlAsync(baseUrl + page.url);
                settingsMenu.DropDownItems.Add(item);
            }

            trayMenu.Items.Add(settingsMenu);

            // ---------- Werkzeuge ----------
            var toolsMenu = new ToolStripMenuItem(Properties.Resources.MenuTools);

            _rateLimitMenu = new ToolStripMenuItem(Properties.Resources.RateLimits);
            _rateLimitMenu.DropDownItems.Add(new ToolStripMenuItem(Properties.Resources.RateLimitLoading) { Enabled = false });
            var refreshItem = new ToolStripMenuItem(Properties.Resources.RateLimitRefresh);
            refreshItem.Click += async (s, e) =>
            {
                if (_rateLimitManager != null)
                    await _rateLimitManager.RefreshAsync();
            };
            _rateLimitMenu.DropDownItems.Add(new ToolStripSeparator());
            _rateLimitMenu.DropDownItems.Add(refreshItem);
            toolsMenu.DropDownItems.Add(_rateLimitMenu);

            toolsMenu.DropDownItems.Add(new ToolStripSeparator());
            toolsMenu.DropDownItems.Add(Properties.Resources.OpenImageFromClipboard, null,
                async (s, e) => await OpenImageFromClipboardAsync());
            toolsMenu.DropDownItems.Add(Properties.Resources.OpenLastImage, null,
                (s, e) => OpenLastImage());

            toolsMenu.DropDownItems.Add(new ToolStripSeparator());
            toolsMenu.DropDownItems.Add(Properties.Resources.ClearCache, null,
                async (s, e) => await ClearCacheAsync());

            trayMenu.Items.Add(toolsMenu);

            // ---------- Über / Beenden ----------
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add(Properties.Resources.AboutWrok, null, (s, e) =>
            {
                using var dlg = new AboutForm();
                // Nicht direkt ShowDialog(this): ist Wrok gerade in die Tray
                // minimiert, verschiebt Windows das Hauptfenster intern weit aus
                // dem Bildschirm - CenterParent würde den Dialog dann relativ zu
                // dieser verschobenen Position zeigen statt zentriert.
                Dialogs.ShowCentered(dlg, this);
            });
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add(Properties.Resources.Exit, null, (s, e) => Application.Exit());

            var initialIcon = ThemeManager.CurrentIcon();
            try
            {
                trayIcon = new NotifyIcon
                {
                    Text             = Properties.Resources.WrokClickToOpen,
                    ContextMenuStrip = trayMenu,
                    Visible          = true,
                    Icon             = (System.Drawing.Icon)initialIcon.Clone()
                };
            }
            catch
            {
                trayIcon = new NotifyIcon
                {
                    Text             = Properties.Resources.WrokClickToOpen,
                    ContextMenuStrip = trayMenu,
                    Visible          = true,
                    Icon             = initialIcon
                };
            }

            try { this.Icon = (System.Drawing.Icon)initialIcon.Clone(); }
            catch { this.Icon = initialIcon; }

            if (trayIcon != null)
            {
                trayIcon.MouseClick += (s, e) => { if (e.Button == MouseButtons.Left) Reactivate(); };
                // Klick auf den Update-Ballon öffnet die Release-Seite.
                trayIcon.BalloonTipClicked += (s, e) => OpenPendingUpdateUrl();
            }
        }

        private void EnsureTrayIconVisible()
        {
            try
            {
                if (trayIcon == null) { InitializeTrayIcon(); _theme?.ApplyIcon(); }
                if (trayIcon != null && !trayIcon.Visible) trayIcon.Visible = true;
            }
            catch { }
        }

        private void DisposeTrayIcon()
        {
            if (trayIcon == null) return;
            try { trayIcon.Visible = false; } catch { }
            try
            {
                var ico = trayIcon.Icon;
                trayIcon.Dispose();
                trayIcon = null;
                try { ico?.Dispose(); } catch { }
            }
            catch { trayIcon = null; }
        }

        // ------------------------------------------------------------------
        // Update-Prüfung
        // ------------------------------------------------------------------

        /// <summary>
        /// Stößt die Update-Prüfung an (best-effort, ohne den Start zu blockieren).
        /// Bei einer neueren Version erscheint ein Tray-Ballon und ein Menüeintrag.
        /// </summary>
        private void StartUpdateCheck()
        {
            _ = Task.Run(async () =>
            {
                try
                {
                    var info = await UpdateChecker.CheckAsync();
                    if (info == null || this.IsDisposed) return;

                    this.BeginInvoke((MethodInvoker)(() => ShowUpdateNotice(info)));
                }
                catch (Exception ex) { Log(ex, "Update-Prüfung fehlgeschlagen"); }
            });
        }

        private void ShowUpdateNotice(UpdateInfo info)
        {
            try
            {
                _pendingUpdateUrl = info.ReleaseUrl;

                // Menüeintrag ganz oben – bleibt sichtbar, auch wenn der Ballon weg ist.
                if (_updateMenuItem == null && trayMenu != null)
                {
                    var item = new ToolStripMenuItem(
                        string.Format(Properties.Resources.UpdateAvailableMenuItem, info.TagName))
                    {
                        Font = new Font(trayMenu.Font, FontStyle.Bold)
                    };
                    item.Click += (s, e) => OpenPendingUpdateUrl();

                    trayMenu.Items.Insert(0, new ToolStripSeparator());
                    trayMenu.Items.Insert(0, item);
                    _updateMenuItem = item;
                }

                trayIcon?.ShowBalloonTip(
                    8000,
                    Properties.Resources.UpdateAvailableTitle,
                    string.Format(Properties.Resources.UpdateAvailableBalloon, info.TagName),
                    ToolTipIcon.Info);
            }
            catch (Exception ex) { Log(ex, "Update-Hinweis anzeigen fehlgeschlagen"); }
        }

        private void OpenPendingUpdateUrl()
        {
            var url = _pendingUpdateUrl;
            if (string.IsNullOrWhiteSpace(url)) return;

            try
            {
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            }
            catch (Exception ex) { Log(ex, "Release-Seite öffnen fehlgeschlagen"); }
        }

        private void UpdateTrayMenuInactivityState()
        {
            // Direkte Referenz statt Suche: das Untermenü hängt seit der Menü-
            // Umstrukturierung unter „Einstellungen → App" und wäre auf der
            // obersten Ebene nicht mehr auffindbar.
            if (_inactivityMenu == null) return;

            int current = Properties.Settings.Default.InactivityTimeoutSeconds;
            foreach (var item in _inactivityMenu.DropDownItems.OfType<ToolStripMenuItem>())
                if (item.Tag is int sec)
                    item.Checked = sec == current;
        }

        // ------------------------------------------------------------------
        // Makros-Menü
        // ------------------------------------------------------------------

        private void InitializeMacrosMenu()
        {
            macrosMenu ??= new ToolStripMenuItem(Properties.Resources.Macros);
            RefreshMacrosMenu();
        }

        /// <summary>
        /// Untermenü zum Wechseln zwischen den Makro-Profilen.
        /// Linksklick wechselt, Rechtsklick benennt um – dieselbe Logik wie bei
        /// den Makros selbst, damit man sich nur ein Muster merken muss.
        /// </summary>
        /// <summary>
        /// Anzeigename eines Profils – bei leerem Namen die Ersatzbezeichnung
        /// „Profil N". An mehreren Stellen gebraucht, deshalb einmal hier.
        /// </summary>
        private string ProfileLabel(int index)
        {
            var names = _macroManager.GetProfileNames();
            if (index < 0 || index >= names.Count)
                return string.Format(Properties.Resources.MacroProfileDefault, index + 1);

            return string.IsNullOrWhiteSpace(names[index])
                ? string.Format(Properties.Resources.MacroProfileDefault, index + 1)
                : names[index];
        }

        private string ActiveProfileLabel() => ProfileLabel(_macroManager.ActiveProfile);

        /// <summary>
        /// Baut den ersten Eintrag des Makro-Menues: Er traegt den Namen des
        /// AKTIVEN Profils und klappt die uebrigen auf. So sieht man den Kontext
        /// der darunter stehenden Makros und wechselt ihn an derselben Stelle.
        /// </summary>
        private ToolStripMenuItem BuildActiveProfileMenu()
        {
            var root = new ToolStripMenuItem(ActiveProfileLabel());

            int active = _macroManager.ActiveProfile;
            int count  = _macroManager.GetProfileNames().Count;

            for (int i = 0; i < count; i++)
            {
                if (i == active) continue;   // steht bereits im Titel des Eintrags

                int idx = i;   // fuer die Closure festhalten
                root.DropDownItems.Add(ProfileLabel(i), null, (s, e) =>
                {
                    _macroManager.SwitchProfile(idx);
                    RefreshMacrosMenu();   // Beschriftung und Makros ziehen mit
                });
            }

            // Verwaltung betrifft immer das aktive Profil - also das, dessen Name
            // ueber diesem Untermenue steht. Dadurch braucht es keinen versteckten
            // Rechtsklick auf die Profilnamen.
            root.DropDownItems.Add(new ToolStripSeparator());
            root.DropDownItems.Add(Properties.Resources.MacroProfileRename, null,
                (s, e) => RenameProfileAndSave(_macroManager.ActiveProfile));
            root.DropDownItems.Add(Properties.Resources.MacroProfileReset, null,
                (s, e) => ResetActiveProfile());

            return root;
        }

        private void ResetActiveProfile()
        {
            try
            {
                int active = _macroManager.ActiveProfile;
                string label = ActiveProfileLabel();

                if (MessageBox.Show(this,
                        string.Format(Properties.Resources.MacroProfileResetConfirm, label),
                        Properties.Resources.MacroProfileReset,
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                    return;

                _macroManager.ResetProfile(active);
                RefreshMacrosMenu();
            }
            catch (Exception ex)
            {
                Log(ex, "ResetActiveProfile fehlgeschlagen");
            }
        }

        private void RenameProfileAndSave(int index)
        {
            try
            {
                var names = _macroManager.GetProfileNames();
                if (index < 0 || index >= names.Count) return;

                string current = names[index];
                if (!Dialogs.ShowSingleLine(this, Properties.Resources.MacroProfileRename,
                                            Properties.Resources.MacroProfileName,
                                            current, out string entered)) return;

                _macroManager.RenameProfile(index, entered);
                RefreshMacrosMenu();
            }
            catch (Exception ex)
            {
                Log(ex, "RenameProfileAndSave fehlgeschlagen");
            }
        }

        private void RefreshMacrosMenu()
        {
            if (macrosMenu == null) return;
            macrosMenu.DropDownItems.Clear();
            macrosMenu.DropDown.ShowItemToolTips = true;   // Tooltips vererben sich nicht auf Untermenues

            // Erster Eintrag: der aktive Profilname, aufklappbar zu den uebrigen.
            macrosMenu.DropDownItems.Add(BuildActiveProfileMenu());
            macrosMenu.DropDownItems.Add(new ToolStripSeparator());


            bool numpadModeOn = Properties.Settings.Default.NumpadMacroModeEnabled;
            int  active       = _macroManager.ActiveProfile;

            var macros = _macroManager.GetMacros();
            for (int i = 1; i <= MacroManager.MacroCount; i++)
            {
                var entry = macros[i - 1];

                var macroItem = new ToolStripMenuItem(entry.DisplayName(i)) { Tag = i - 1 };

                // Strg+Ziffer gilt immer (Makro 10 → Strg+0). Ist der Nummernblock-
                // Modus an, kommt für die ersten fünf Makros die Numpad-Taste dazu.
                string ctrlDigit     = i <= 9 ? i.ToString() : "0";
                string hotkeyDisplay = string.Format(Properties.Resources.CtrlMacroLabel, ctrlDigit);
                if (numpadModeOn && (i - 1) < MacroManager.NumpadPerProfile)
                    hotkeyDisplay += " / " + NumpadKeyLabel(active * MacroManager.NumpadPerProfile + (i - 1));
                string preview = string.IsNullOrWhiteSpace(entry.Text)
                    ? Properties.Resources.EmptyMacro
                    : (entry.Text.Length > 80 ? entry.Text[..80] + "…" : entry.Text);
                macroItem.ToolTipText = string.Format(Properties.Resources.EditWithRightClickRunWith0, hotkeyDisplay) + "\n" + preview;

                macroItem.MouseDown += async (sender, me) =>
                {
                    try
                    {
                        if (sender is not ToolStripMenuItem tsi) return;
                        int idx = tsi.Tag is int ii ? ii : -1;
                        if (idx < 0) return;

                        if (me.Button == MouseButtons.Left)
                        {
                            bool pressEnter = !((ModifierKeys & Keys.Shift) == Keys.Shift ||
                                                (ModifierKeys & Keys.Alt)   == Keys.Alt);
                            await SendMacroTextAsync(_macroManager.GetMacros()[idx].Text, pressEnter);
                        }
                        else if (me.Button == MouseButtons.Right)
                        {
                            EditMacroAndSave(idx);
                        }
                    }
                    catch { }
                };
                macrosMenu.DropDownItems.Add(macroItem);
            }

        }

        private void EditMacroAndSave(int index)
        {
            try
            {
                var macros  = _macroManager.GetMacros();
                var current = index >= 0 && index < macros.Count
                    ? macros[index] : new MacroEntry(string.Empty, string.Empty);

                string name = current.Name;
                string text = current.Text;
                string dlgTitle = index >= 0
                    ? string.Format(Properties.Resources.EditMacro + " #{0}", index + 1)
                    : Properties.Resources.NewMacro;

                if (!Dialogs.ShowEditMacro(this, dlgTitle, ref name, ref text)) return;

                _macroManager.UpdateMacro(index, new MacroEntry(name.Trim(), text.Trim()));
                RefreshMacrosMenu();
            }
            catch (Exception ex)
            {
                Log(ex, "EditMacroAndSave fehlgeschlagen");
            }
        }

        // ------------------------------------------------------------------
        // Grok-Konto wechseln
        // ------------------------------------------------------------------

        private const string GrokSignInHost = "accounts.x.ai";
        private const string GrokHost       = "grok.com";

        // Merkt sich die E-Mail, die nach der naechsten abgeschlossenen Navigation
        // zur Sign-in-Seite per JS eingetragen werden soll. Wird nach einmaliger
        // Verwendung sofort geloescht, damit Folge-Navigationen (z. B. nach dem
        // Login weiter zu grok.com) nicht erneut befuellt werden.
        private string? _pendingGrokLoginEmail;

        // Wird beim Kontowechsel gesetzt: da SwitchGrokAccountAsync die Cookies
        // fuer grok.com UND accounts.x.ai loescht, geht dabei auch das
        // OneTrust-Cookie-Consent fuer grok.com verloren - der Banner dort taucht
        // dann (live vom Nutzer bestaetigt) erst NACH dem erfolgreichen Login und
        // der Weiterleitung zurueck zu grok.com auf, nicht schon auf der
        // Sign-in-Seite. Diese Flag sorgt dafuer, dass der naechste
        // NavigationCompleted-Event auf grok.com (also nach abgeschlossenem
        // Login) den Banner dort ebenfalls automatisch wegklickt - einmalig, um
        // normale grok.com-Navigationen ausserhalb eines Kontowechsels nicht zu
        // beeinflussen.
        private bool _pendingGrokCookieDismiss;

        private void InitializeGrokAccountsMenu()
        {
            grokAccountsMenu ??= new ToolStripMenuItem(Properties.Resources.GrokAccountsMenu);
            RefreshGrokAccountsMenu();
        }

        private void RefreshGrokAccountsMenu()
        {
            if (grokAccountsMenu == null) return;
            grokAccountsMenu.DropDownItems.Clear();

            var accounts = GrokAccountManager.LoadAccounts();

            if (accounts.Count == 0)
            {
                grokAccountsMenu.DropDownItems.Add(new ToolStripMenuItem(Properties.Resources.GrokAccountsEmpty) { Enabled = false });
            }
            else
            {
                foreach (var acc in accounts)
                {
                    var item = new ToolStripMenuItem(acc.DisplayName);
                    item.Click += async (s, e) => await SwitchGrokAccountAsync(acc.Email);
                    grokAccountsMenu.DropDownItems.Add(item);
                }
            }

            grokAccountsMenu.DropDownItems.Add(new ToolStripSeparator());
            grokAccountsMenu.DropDownItems.Add(Properties.Resources.GrokAccountAdd, null, (s, e) => AddGrokAccountAndSave());

            if (accounts.Count > 0)
                grokAccountsMenu.DropDownItems.Add(Properties.Resources.GrokAccountRemove, null, (s, e) => RemoveGrokAccountAndSave(accounts));
        }

        private void AddGrokAccountAndSave()
        {
            try
            {
                if (!Dialogs.ShowAddGrokAccount(this, out string label, out string email)) return;

                GrokAccountManager.AddAccount(label, email);
                RefreshGrokAccountsMenu();
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(this, ex.Message, Properties.Resources.GrokAccountAddTitle,
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                Log(ex, "AddGrokAccountAndSave fehlgeschlagen");
            }
        }

        private void RemoveGrokAccountAndSave(List<GrokAccount> accounts)
        {
            try
            {
                if (!Dialogs.ShowRemoveGrokAccount(this, accounts, out int index)) return;

                GrokAccountManager.RemoveAccount(accounts[index].Email);
                RefreshGrokAccountsMenu();
            }
            catch (Exception ex)
            {
                Log(ex, "RemoveGrokAccountAndSave fehlgeschlagen");
            }
        }

        /// <summary>
        /// Wechselt zum angegebenen Grok-Konto: loescht die Session-Cookies fuer
        /// grok.com/accounts.x.ai (sonst landet man in der alten Sitzung statt im
        /// Login-Formular), navigiert zur Sign-in-Seite und traegt danach die
        /// E-Mail-Adresse automatisch ein. Das Passwort bleibt manuell bzw. Sache
        /// des Chromium-eigenen Passwort-Managers (siehe WebViewManager.ConfigureCore).
        /// </summary>
        private async Task SwitchGrokAccountAsync(string email)
        {
            if (_webView?.CoreWebView2 == null || _webViewManager == null) return;

            try
            {
                _pendingGrokLoginEmail = email;
                _pendingGrokCookieDismiss = true;

                var cookieManager = _webView.CoreWebView2.CookieManager;
                foreach (var host in new[] { GrokHost, GrokSignInHost })
                {
                    var cookies = await cookieManager.GetCookiesAsync($"https://{host}");
                    foreach (var cookie in cookies)
                        cookieManager.DeleteCookie(cookie);
                }

                await _webViewManager.NavigateAsync($"https://{GrokSignInHost}/sign-in?redirect=grok-com&email=true");
            }
            catch (Exception ex)
            {
                _pendingGrokLoginEmail = null;
                _pendingGrokCookieDismiss = false;
                Log(ex, "SwitchGrokAccountAsync fehlgeschlagen");
            }
        }

        private async void OnGrokSignInNavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (_webView?.CoreWebView2 == null) return;
            string? source = _webView.CoreWebView2.Source;

            if (_pendingGrokLoginEmail != null &&
                (source?.Contains(GrokSignInHost, StringComparison.OrdinalIgnoreCase) ?? false))
            {
                string email = _pendingGrokLoginEmail;
                _pendingGrokLoginEmail = null;   // nur einmal verwenden, sonst Re-Fill bei Folge-Navigationen

                await PrefillGrokEmailAsync(email);
                return;
            }

            // Nach abgeschlossenem Login landet man per redirect=grok-com wieder
            // auf grok.com - dort ggf. erneut auftauchendes Cookie-Consent-Banner
            // (siehe Kommentar bei _pendingGrokCookieDismiss) einmalig wegklicken.
            if (_pendingGrokCookieDismiss &&
                (source?.Contains(GrokHost, StringComparison.OrdinalIgnoreCase) ?? false) &&
                !(source?.Contains(GrokSignInHost, StringComparison.OrdinalIgnoreCase) ?? false))
            {
                _pendingGrokCookieDismiss = false;
                await TryDismissCookieBannerAsync();
            }
        }

        // JS-Helper zum Finden eines Buttons: erst gezielte Selektoren, dann
        // Text-Scan (DE+EN) über alle Buttons/role=button-Elemente. Gleiches
        // Muster wie WebViewManager.TryClickSendButtonAsync. Wird in mehrere
        // der folgenden Skripte eingebettet.
        private const string FindByHeuristicsJs = @"
            function findByHeuristics(selectors, textNeedles) {
                var btn = document.querySelector(selectors.join(', '));
                if (btn) return btn;
                var candidates = Array.from(document.querySelectorAll('button, [role=button]'));
                for (var i = 0; i < candidates.length; i++) {
                    var txt = ((candidates[i].innerText || candidates[i].getAttribute('aria-label') || candidates[i].title) + '').toLowerCase();
                    for (var j = 0; j < textNeedles.length; j++) {
                        if (txt.indexOf(textNeedles[j]) !== -1) return candidates[i];
                    }
                }
                return null;
            }";

        /// <summary>
        /// Führt <paramref name="script"/> (muss ein IIFE sein, das true/false
        /// zurückgibt) wiederholt aus, bis es true liefert oder die Versuche
        /// aufgebraucht sind. Bewusst von der C#-Seite aus gepollt statt mit
        /// einem einzigen in-page setInterval: Falls die Seite zwischen zwei
        /// Schritten des Sign-in-Flows tatsächlich neu navigiert (statt nur
        /// React-intern den Zustand zu wechseln), würde ein laufendes
        /// setInterval mit der ganzen Seite zerstört, bevor es fertig ist - ein
        /// C#-Loop dagegen fragt bei jedem Versuch einfach den gerade aktuellen
        /// Seitenzustand neu ab, unabhängig davon, ob zwischendurch navigiert wurde.
        /// </summary>
        private async Task<bool> RunJsRetryAsync(string script, int attempts, int delayMs)
        {
            if (_webViewManager == null) return false;

            for (int attempt = 0; attempt < attempts; attempt++)
            {
                try
                {
                    string? result = await _webViewManager.ExecuteScriptAsync(script);
                    if (result != null && result.Trim().Trim('"').Equals("true", StringComparison.OrdinalIgnoreCase))
                        return true;
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"RunJsRetryAsync fehlgeschlagen: {ex}");
                }
                await Task.Delay(delayMs);
            }
            return false;
        }

        /// <summary>
        /// Schliesst das OneTrust-Cookie-Banner (falls vorhanden) über "Alle
        /// ablehnen" - datensparsamste Standardoption. Best-effort: läuft nur
        /// kurz mit und blockiert den restlichen Flow nicht, falls kein Banner
        /// auftaucht (z. B. weil die Zustimmung schon per Cookie vorliegt).
        /// </summary>
        private Task<bool> TryDismissCookieBannerAsync()
        {
            const string js = @"
                (function() {
                    var btn = document.querySelector('#onetrust-reject-all-handler');
                    if (!btn) {
                        var candidates = Array.from(document.querySelectorAll('button'));
                        for (var i = 0; i < candidates.length; i++) {
                            var txt = (candidates[i].innerText || '').toLowerCase();
                            if (txt.indexOf('ablehnen') !== -1 || txt.indexOf('reject') !== -1 || txt.indexOf('decline') !== -1) { btn = candidates[i]; break; }
                        }
                    }
                    if (!btn) return false;
                    try { btn.click(); return true; } catch (e) { return false; }
                })();";

            // Etwas grosszuegigeres Fenster als bei den anderen Retry-Aufrufen:
            // OneTrust-Consent-Skripte binden sich haeufig erst nach dem eigentlichen
            // Seitenladen (NavigationCompleted) asynchron ein, u. a. wegen einer
            // Geo-IP-Abfrage im Hintergrund. Nicht live nachgestellt (Banner liess
            // sich im Test-Browser nicht reproduzieren), daher bewusst grosszuegig
            // statt knapp bemessen.
            return RunJsRetryAsync(js, attempts: 25, delayMs: 300);
        }

        private Task<bool> TrySetEmailAsync(string email)
        {
            string payload = System.Text.Json.JsonSerializer.Serialize(email);
            string js = $@"
                (function() {{
                    var input = document.querySelector('input[data-testid=""email""]');
                    if (!input) return false;
                    var nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                    nativeSetter.call(input, {payload});
                    input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    input.dispatchEvent(new Event('change', {{ bubbles: true }}));
                    return true;
                }})();";

            return RunJsRetryAsync(js, attempts: 20, delayMs: 100);
        }

        private Task<bool> TryClickContinueButtonAsync()
        {
            string js = $@"
                (function() {{
                    {FindByHeuristicsJs}
                    var btn = findByHeuristics(
                        ['button[type=""submit""]', 'button[data-testid*=""continue"" i]', 'button[data-testid*=""submit"" i]',
                         'button[aria-label*=""weiter"" i]', 'button[aria-label*=""continue"" i]', 'button[aria-label*=""next"" i]'],
                        ['weiter', 'continue', 'next']);
                    if (!btn || btn.disabled) return false;
                    try {{ btn.click(); return true; }} catch (e) {{ return false; }}
                }})();";

            return RunJsRetryAsync(js, attempts: 15, delayMs: 150);
        }

        /// <summary>
        /// Wartet auf befülltes Passwortfeld, abgeschlossene Cloudflare-Turnstile-
        /// Prüfung UND aktivierten Anmelden-Button (`data-testid="sign-in-submit"`,
        /// live verifiziert), dann Klick. Der Button ist waehrend der laufenden
        /// Cloudflare-Pruefung NICHT disabled (live verifiziert - vermutlich der
        /// Grund fuer bisherige Fehlklicks: der Button wurde geklickt, bevor
        /// Cloudflare fertig war). Verlaessliches Signal fuer "fertig" ist das von
        /// Cloudflare selbst befuellte versteckte Feld
        /// input[name="cf-turnstile-response"] (live verifiziert: leer waehrend
        /// der Pruefung, befuellt bei Erfolg). Ist auf der Seite kein
        /// Turnstile-Feld vorhanden, wird diese Bedingung uebersprungen, um nicht
        /// grundlos zu blockieren, falls Grok die Pruefung fuer ein Konto mal
        /// nicht anzeigt. Passwort selbst wird NICHT automatisiert - bleibt
        /// manuell bzw. Sache von Chromiums eigenem Passwort-Manager.
        /// Grosszuegiges Zeitfenster (Standard: 2 Minuten), da hier auf eine
        /// menschliche Eingabe gewartet wird.
        /// </summary>
        private Task<bool> TryClickSignInButtonAsync(int attempts = 400, int delayMs = 300)
        {
            string js = $@"
                (function() {{
                    {FindByHeuristicsJs}
                    var pwField = document.querySelector('input[type=""password""]');
                    if (!pwField || !pwField.value) return false;

                    var turnstileField = document.querySelector('input[name*=""turnstile"" i]');
                    if (turnstileField && !turnstileField.value) return false;

                    var btn = document.querySelector('button[data-testid=""sign-in-submit""]') ||
                              findByHeuristics(['button[type=""submit""]'], ['anmelden', 'sign in', 'log in', 'login']);
                    if (!btn || btn.disabled) return false;
                    try {{ btn.click(); return true; }} catch (e) {{ return false; }}
                }})();";

            return RunJsRetryAsync(js, attempts, delayMs);
        }

        /// <summary>
        /// Steuert den kompletten Sign-in-Flow: Cookie-Banner wegklicken → E-Mail
        /// eintragen → "Weiter" klicken → warten bis Passwort da + Anmelden aktiv
        /// → "Anmelden" klicken. Jeder Schritt ist ein eigener, von C# aus
        /// wiederholter Aufruf statt eines einzelnen JS-Blocks mit
        /// verschachtelten setInterval-Timern (siehe RunJsRetryAsync).
        /// </summary>
        private async Task PrefillGrokEmailAsync(string email)
        {
            if (_webViewManager == null) return;

            await TryDismissCookieBannerAsync();   // best effort, blockiert den Rest nicht

            if (!await TrySetEmailAsync(email))
            {
                Trace.WriteLine("PrefillGrokEmailAsync: E-Mail-Feld nicht gefunden.");
                return;
            }

            if (!await TryClickContinueButtonAsync())
            {
                Trace.WriteLine("PrefillGrokEmailAsync: Weiter-Button nicht gefunden/geklickt.");
                return;
            }

            if (!await TryClickSignInButtonAsync())
                Trace.WriteLine("PrefillGrokEmailAsync: Anmelden-Button nicht innerhalb des Zeitfensters geklickt.");
        }

        // ------------------------------------------------------------------
        // Inaktivitäts-Timer
        // ------------------------------------------------------------------

        private void InactivityMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is not ToolStripMenuItem clicked) return;
            int seconds = Convert.ToInt32(clicked.Tag ?? 0);

            _inactivity?.SetTimeout(seconds);

            if (clicked.OwnerItem is ToolStripMenuItem parent)
                foreach (var item in parent.DropDownItems.OfType<ToolStripMenuItem>())
                    item.Checked = item == clicked;
        }

        // ------------------------------------------------------------------
        // Thema / Dark Mode
        // ------------------------------------------------------------------

        // ------------------------------------------------------------------
        // Rate Limits
        // ------------------------------------------------------------------

        private async Task InitializeRateLimitManagerAsync()
        {
            // Kurz warten bis WebView initialisiert ist
            await Task.Delay(3000);
            if (_webViewManager == null) return;

            _rateLimitManager = new RateLimitManager(_webViewManager, UpdateRateLimitMenu);
            _rateLimitManager.StartAutoRefresh();
            await _rateLimitManager.RefreshAsync();
        }

        private void UpdateRateLimitMenu()
        {
            if (_rateLimitMenu == null) return;
            if (this.IsDisposed || !this.IsHandleCreated) return;

            try
            {
                this.BeginInvoke((System.Windows.Forms.MethodInvoker)(() =>
                {
                    try
                    {
                        var result = _rateLimitManager?.LastResult;

                        // Alle Items außer Separator und Refresh-Button entfernen
                        var toRemove = _rateLimitMenu.DropDownItems
                            .OfType<ToolStripItem>()
                            .Where(i => i is not ToolStripSeparator &&
                                        i.Text != Properties.Resources.RateLimitRefresh)
                            .ToList();
                        foreach (var item in toRemove)
                            _rateLimitMenu.DropDownItems.Remove(item);

                        if (result == null)
                        {
                            _rateLimitMenu.DropDownItems.Insert(0,
                                new ToolStripMenuItem(Properties.Resources.RateLimitLoading) { Enabled = false });
                            return;
                        }

                        if (result.FetchError != null)
                        {
                            _rateLimitMenu.DropDownItems.Insert(0,
                                new ToolStripMenuItem(
                                    string.Format(Properties.Resources.RateLimitError0, result.FetchError))
                                { Enabled = false });
                            return;
                        }

                        int insertAt = 0;

                        // Grok 3
                        insertAt = InsertRateLimitItems(_rateLimitMenu, insertAt,
                            "Grok 3", result.Grok3);

                        // Grok 4 Heavy
                        insertAt = InsertRateLimitItems(_rateLimitMenu, insertAt,
                            "Grok 4 Heavy", result.Grok4Heavy);

                        // Zeitstempel
                        var tsItem = new ToolStripMenuItem(
                            string.Format(Properties.Resources.RateLimitUpdatedAt0,
                                result.FetchedAt.ToString("HH:mm:ss")))
                        { Enabled = false, Font = new Font("Segoe UI", 7.5f) };
                        _rateLimitMenu.DropDownItems.Insert(insertAt, tsItem);
                    }
                    catch (Exception ex)
                    {
                        Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MainForm] UpdateRateLimitMenu failed: {ex}");
                    }
                }));
            }
            catch { }
        }

        /// <summary>
        /// Farbe fuer knappe Kontingente, sonst null (dann bleibt die Standardfarbe).
        /// Die Toene sind so gewaehlt, dass sie auf hellem wie dunklem Menuehintergrund
        /// lesbar bleiben – reines Rot waere auf Dunkel schlecht erkennbar.
        /// </summary>
        private static Color? SeverityColor(RateLimitEntry entry) => entry.Severity switch
        {
            RateLimitSeverity.Critical => Color.FromArgb(220, 60, 60),    // rot
            RateLimitSeverity.Warning  => Color.FromArgb(200, 130, 0),    // orange
            _                          => null
        };

        private static int InsertRateLimitItems(ToolStripMenuItem menu, int insertAt, string label, RateLimitEntry? entry)
        {
            if (entry == null) return insertAt;

            string text;
            if (entry.IsError)
            {
                text = entry.Error == "UNAUTHORIZED"
                    ? string.Format(Properties.Resources.RateLimitModel0NotLoggedIn, label)
                    : string.Format(Properties.Resources.RateLimitModel0Error1, label, entry.Error);
            }
            else if (entry.IsUnknown)
            {
                text = string.Format(Properties.Resources.RateLimitModel0Unknown, label);
            }
            else
            {
                text = string.Format(Properties.Resources.RateLimitModel0Remaining1Of2,
                    label, entry.RemainingQueries, entry.TotalQueries);
            }

            var item = new ToolStripMenuItem(text) { Enabled = false };

            // Farbe erst ab "knapp": Solange reichlich Kontingent da ist, bleibt der
            // Eintrag unauffaellig grau wie die uebrigen deaktivierten Menuepunkte.
            // Deaktivierte Items zeichnet WinForms sonst grundsaetzlich ausgegraut,
            // deshalb muss die Farbe ueber ForeColor gesetzt werden.
            var accent = SeverityColor(entry);
            if (accent.HasValue)
            {
                item.ForeColor = accent.Value;
                item.Font      = new Font(item.Font ?? SystemFonts.MenuFont!, FontStyle.Bold);
            }

            if (!entry.IsError && !entry.IsUnknown && !string.IsNullOrEmpty(entry.ResetInfo))
            {
                var resetItem = new ToolStripMenuItem($"  {entry.ResetInfo}") { Enabled = false,
                    Font = new Font("Segoe UI", 7.5f) };
                menu.DropDownItems.Insert(insertAt++, item);
                menu.DropDownItems.Insert(insertAt++, resetItem);
            }
            else
            {
                menu.DropDownItems.Insert(insertAt++, item);
            }

            return insertAt;
        }

        // ------------------------------------------------------------------
        // Makros Import / Export
        // ------------------------------------------------------------------

        /// <param name="allProfiles">
        /// true = alle Profile samt Namen, false = nur das aktive Profil.
        /// </param>
        private void ExportMacros(bool allProfiles)
        {
            // Dateiname macht sichtbar, was drinsteckt.
            string profile  = ActiveProfileLabel();
            string suggested = allProfiles
                ? string.Format(Properties.Resources.ExportFileNameAll, $"{DateTime.Now:yyyyMMdd}")
                : string.Format(Properties.Resources.ExportFileNameOne, profile, $"{DateTime.Now:yyyyMMdd}");

            using var dlg = new SaveFileDialog
            {
                Title           = Properties.Resources.ExportMacrosTitle,
                Filter          = Properties.Resources.ExportMacrosFilter,
                FileName        = SanitizeFileName(suggested),
                DefaultExt      = "json",
                OverwritePrompt = true
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            bool ok = allProfiles
                ? _macroManager.ExportAllProfiles(dlg.FileName)
                : _macroManager.ExportActiveProfile(dlg.FileName);

            if (ok)
                MessageBox.Show(this, string.Format(Properties.Resources.ExportSuccessMsg, dlg.FileName),
                    Properties.Resources.ExportSuccessTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show(this, Properties.Resources.ExportFailedMsg,
                    Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static string SanitizeFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }

        private void ImportMacros()
        {
            using var dlg = new OpenFileDialog
            {
                Title  = Properties.Resources.ImportMacrosTitle,
                Filter = Properties.Resources.ExportMacrosFilter,
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            // Erst nachsehen, was in der Datei steckt – davon haengt ab, was zu
            // fragen ist. Bis hierhin wurde nichts veraendert.
            var kind = MacroManager.Inspect(dlg.FileName);
            bool ok;

            switch (kind)
            {
                case MacroImportKind.AllProfiles:
                    if (MessageBox.Show(this, Properties.Resources.ImportAllConfirmMsg,
                            Properties.Resources.ImportConfirmTitle,
                            MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                        return;
                    ok = _macroManager.ImportAllProfiles(dlg.FileName);
                    break;

                case MacroImportKind.SingleProfile:
                    if (!Dialogs.ShowProfileChooser(this, _macroManager.GetProfileNames(),
                                                    _macroManager.ActiveProfile, out int target)) return;
                    ok = _macroManager.ImportIntoProfile(dlg.FileName, target);
                    break;

                default:
                    MessageBox.Show(this, Properties.Resources.ImportInvalidMsg,
                        Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
            }

            if (ok)
            {
                RefreshMacrosMenu();
                MessageBox.Show(this, Properties.Resources.ImportSuccessMsg,
                    Properties.Resources.ImportSuccessTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show(this, Properties.Resources.ImportFailedMsg,
                    Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        // ------------------------------------------------------------------
        // Proxy-Einstellungen
        // ------------------------------------------------------------------

        /// <summary>
        /// Zeigt den Proxy-Dialog und speichert das Ergebnis. Chromium liest
        /// den Proxy nur beim Erzeugen der CoreWebView2Environment (siehe
        /// WebViewManager.InitializeAsync) – eine Änderung wirkt daher erst
        /// nach einem Neustart, den diese Methode bei Bedarf direkt anbietet.
        /// </summary>
        private void ShowProxyDialog()
        {
            bool   enabled      = Properties.Settings.Default.ProxyEnabled;
            string server       = Properties.Settings.Default.ProxyServer ?? string.Empty;
            bool   requiresAuth = Properties.Settings.Default.ProxyAuthEnabled;
            string username     = Properties.Settings.Default.ProxyUsername ?? string.Empty;
            // Nur zum Anzeigen/Bearbeiten im Dialog entschlüsselt, danach sofort
            // wieder verworfen (liegt nur lokal in dieser Methode im Klartext).
            string password     = ProxyCredentialProtector.Unprotect(Properties.Settings.Default.ProxyPasswordProtected);

            if (!Dialogs.ShowProxySettings(this, ref enabled, ref server, ref requiresAuth, ref username, ref password))
                return;

            bool changed = enabled      != Properties.Settings.Default.ProxyEnabled
                        || server       != (Properties.Settings.Default.ProxyServer   ?? string.Empty)
                        || requiresAuth != Properties.Settings.Default.ProxyAuthEnabled
                        || username     != (Properties.Settings.Default.ProxyUsername ?? string.Empty)
                        || password     != ProxyCredentialProtector.Unprotect(Properties.Settings.Default.ProxyPasswordProtected);
            if (!changed) return;

            Properties.Settings.Default.ProxyEnabled          = enabled;
            Properties.Settings.Default.ProxyServer           = server;
            Properties.Settings.Default.ProxyAuthEnabled      = requiresAuth;
            // Ohne aktive Anmeldung keine Reste stehen lassen, statt nur die
            // Checkbox umzuschalten und ein altes Passwort verschlüsselt liegen
            // zu lassen.
            Properties.Settings.Default.ProxyUsername         = requiresAuth ? username : string.Empty;
            Properties.Settings.Default.ProxyPasswordProtected = requiresAuth
                ? ProxyCredentialProtector.Protect(password)
                : string.Empty;
            Properties.Settings.Default.Save();

            var result = MessageBox.Show(this, Properties.Resources.ProxyRestartRequiredMessage,
                Properties.Resources.ProxyRestartRequiredTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Information);

            if (result == DialogResult.Yes)
                Program.RestartApplication();
        }

        // ------------------------------------------------------------------
        // Cache leeren
        // ------------------------------------------------------------------

        private async Task ClearCacheAsync()
        {
            try
            {
                if (_webView?.CoreWebView2 == null)
                { MessageBox.Show(Properties.Resources.WebView2IsNotInitializedYet, Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var choice = Dialogs.ShowClearCacheChoice(this, out bool includeMacros);
                if (choice == ClearCacheChoice.Cancel) return;

                if (choice == ClearCacheChoice.CacheOnly)
                {
                    // Nur Cache (Bilder, Skripte, …) – Cookies und Login bleiben erhalten.
                    await _webView.CoreWebView2.Profile.ClearBrowsingDataAsync(
                        CoreWebView2BrowsingDataKinds.DiskCache |
                        CoreWebView2BrowsingDataKinds.CacheStorage);
                    MessageBox.Show(
                        Properties.Resources.PicturesScriptsAndOtherDataDeletedNYouAreStillLoggedIn,
                        Properties.Resources.CacheCleared, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // Alles: erst die WebView2-Seite (Cookies, Login, lokaler Speicher),
                    // dann die Spuren, die Wrok selbst ausserhalb des Browserprofils
                    // hinterlaesst. Ohne den zweiten Schritt blieben das zuletzt
                    // geoeffnete Bild und dessen URL zurueck.
                    await _webView.CoreWebView2.Profile.ClearBrowsingDataAsync();
                    ClearWrokPrivateData(includeMacros);

                    MessageBox.Show(
                        includeMacros
                            ? Properties.Resources.AllDataDeletedInclMacrosMsg
                            : Properties.Resources.AllDataDeletedMsg,
                        Properties.Resources.AllDataDeleted, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MainForm] ClearCacheAsync fehlgeschlagen: {ex}");
                MessageBox.Show(string.Format(Properties.Resources.ErrorWhileDeletingCache + "\n{0}", ex.Message), Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Entfernt die Daten, die Wrok ausserhalb des WebView2-Profils ablegt.
        ///
        /// <see cref="CoreWebView2Profile.ClearBrowsingDataAsync()"/> raeumt nur den
        /// Browser auf. Das zuletzt geoeffnete Bild liegt aber als PNG unter
        /// %LOCALAPPDATA%\Wrok\images, und seine Quell-URL steht im Klartext in den
        /// Einstellungen – beides ueberlebte das Loeschen bisher unbemerkt.
        /// </summary>
        private void ClearWrokPrivateData(bool includeMacros)
        {
            // Ein offenes Medienfenster wuerde sonst genau das weiter anzeigen,
            // was gerade geloescht wird.
            CloseCurrentMedia();

            try
            {
                var path = Properties.Settings.Default.LastImagePath;
                if (!string.IsNullOrWhiteSpace(path))
                {
                    // Auch die temporaere Datei aus SaveAsLastImage, die bei einem
                    // Abbruch liegen bleiben kann.
                    foreach (var candidate in new[] { path, path + ".tmp" })
                        if (File.Exists(candidate)) File.Delete(candidate);
                }
            }
            catch (Exception ex) { Log(ex, "Letztes Bild loeschen fehlgeschlagen"); }

            try
            {
                Properties.Settings.Default.LastImagePath = string.Empty;
                Properties.Settings.Default.LastImageUrl  = string.Empty;
                Properties.Settings.Default.Save();
            }
            catch (Exception ex) { Log(ex, "Bildverweise zuruecksetzen fehlgeschlagen"); }

            if (!includeMacros) return;

            try
            {
                for (int i = 0; i < MacroManager.ProfileCount; i++)
                    _macroManager.ResetProfile(i);

                RefreshMacrosMenu();
            }
            catch (Exception ex) { Log(ex, "Makroprofile loeschen fehlgeschlagen"); }
        }

        // ------------------------------------------------------------------
        // Form-Events
        // ------------------------------------------------------------------

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Nur der bewusste Klick aufs X minimiert in den Tray. Ein Sitzungsende
            // (Abmelden, Herunterfahren, Installer-Restart-Manager) beendet dagegen
            // wirklich - sonst haelt Wrok die laufende Exe fest und ein stilles
            // Upgrade schlaegt fehl. _sessionEnding faengt auch den Fall ab, dass der
            // Restart-Manager das Schliessen als gewoehnliches WM_CLOSE schickt.
            if (e.CloseReason == CloseReason.UserClosing && !_sessionEnding)
            {
                SaveWindowSettings();
                e.Cancel = true;
                // Derselbe Weg wie die Boss-Taste – damit auch hier kein
                // Medienfenster sichtbar zurueckbleibt.
                MinimizeToTray();
            }
            else
            {
                // Echtes Beenden: Fensterlage noch sichern, dann durchlassen.
                SaveWindowSettings();
            }
            base.OnFormClosing(e);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            _theme?.StartListening();
            base.OnHandleCreated(e);

            _webViewManager?.CreateInputSimulator();

            bool ok = RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, (uint)Keys.Space);
            if (!ok) Debug.WriteLine($"RegisterHotKey fehlgeschlagen id={HOTKEY_ID} err={Marshal.GetLastWin32Error()}");

            // Strg+Ö → Bild aus Zwischenablage öffnen
            ok = RegisterHotKey(this.Handle, HOTKEY_ID_IMAGE, MOD_CONTROL, (uint)Keys.Oem1);
            if (!ok) Trace.WriteLine($"RegisterHotKey fehlgeschlagen id={HOTKEY_ID_IMAGE} (Strg+Ö) err={Marshal.GetLastWin32Error()}");

            // Strg+1..Strg+0 für das aktive Profil – immer aktiv, auf jeder Tastatur.
            RegisterCtrlMacroHotkeys();

            // Nummernblock zusätzlich, aber nur wenn eingeschaltet.
            if (Properties.Settings.Default.NumpadMacroModeEnabled)
                RegisterNumpadMacroHotkeys();

            _theme?.Refresh();
            EnsureTrayIconVisible();
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            _rateLimitManager?.StopAutoRefresh();
            try { UnregisterHotKey(this.Handle, HOTKEY_ID); } catch { }
            try { UnregisterHotKey(this.Handle, HOTKEY_ID_IMAGE); } catch { }
            UnregisterCtrlMacroHotkeys();
            UnregisterNumpadMacroHotkeys();

            _webViewManager?.DisposeInputSimulator();
            _theme?.StopListening();   // wird in OnHandleCreated wieder angemeldet

            // ACHTUNG: Diese Methode laeuft auch bei jeder Neuerzeugung des
            // Fensterhandles, und die erzwingt WinForms bereits beim Aendern von
            // ShowInTaskbar - was Wrok beim Anzeigen und Minimieren staendig tut.
            // Der Inaktivitaets-Waechter darf deshalb nur beim echten Beenden
            // abgeraeumt werden, sonst ist er nach dem ersten Anzeigen tot.
            if (!this.RecreatingHandle)
            {
                _inactivity?.Dispose();
                DisposeTrayIcon();
            }
            base.OnHandleDestroyed(e);
        }

        // ------------------------------------------------------------------
        // Numpad-Makro-Modus
        // ------------------------------------------------------------------

        /// <summary>
        /// Registriert alle 15 Numpad-Makro-Hotkeys ohne Modifier. Da diese Tasten
        /// dabei systemweit für JEDE Anwendung abgefangen werden (keine normale
        /// Zahlen-/Rechenzeicheneingabe über den Nummernblock mehr möglich, solange
        /// der Modus an ist), ist das an die Einstellung NumpadMacroModeEnabled
        /// gekoppelt statt immer aktiv zu sein.
        /// </summary>
        /// <summary>Registriert Strg+1..Strg+0 für das aktive Profil.</summary>
        private void RegisterCtrlMacroHotkeys()
        {
            for (int i = 0; i < CtrlMacroKeys.Length; i++)
            {
                bool ok = RegisterHotKey(this.Handle, MacroManager.CtrlMacroBase + i, MOD_CONTROL, (uint)CtrlMacroKeys[i]);
                if (!ok) Debug.WriteLine($"RegisterHotKey fehlgeschlagen strg-makro={i} err={Marshal.GetLastWin32Error()}");
            }
        }

        private void UnregisterCtrlMacroHotkeys()
        {
            for (int i = 0; i < CtrlMacroKeys.Length; i++)
                try { UnregisterHotKey(this.Handle, MacroManager.CtrlMacroBase + i); } catch { }
        }

        private void RegisterNumpadMacroHotkeys()
        {
            EnsureNumLockOn();
            for (int i = 0; i < NumpadMacroKeys.Length; i++)
            {
                bool ok = RegisterHotKey(this.Handle, MacroManager.NumpadMacroBase + i, 0 /* kein Modifier */, (uint)NumpadMacroKeys[i]);
                if (!ok) Debug.WriteLine($"RegisterHotKey fehlgeschlagen numpad-makro={i} err={Marshal.GetLastWin32Error()}");
            }
        }

        private void UnregisterNumpadMacroHotkeys()
        {
            for (int i = 0; i < NumpadMacroKeys.Length; i++)
                try { UnregisterHotKey(this.Handle, MacroManager.NumpadMacroBase + i); } catch { }
        }

        /// <summary>Schaltet den Numpad-Makro-Modus um, inkl. Persistierung.</summary>
        private void SetNumpadMacroMode(bool enabled)
        {
            if (enabled) RegisterNumpadMacroHotkeys();
            else UnregisterNumpadMacroHotkeys();

            Properties.Settings.Default.NumpadMacroModeEnabled = enabled;
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Stellt sicher, dass NumLock aktiv ist – die virtuellen Codes der
        /// Numpad-Ziffern (und des Dezimalpunkts) werden nur dann gesendet,
        /// sonst liefern dieselben Tasten Navigationscodes (Pfeile, Bild↑/↓ usw.).
        /// Betrifft nicht die vier Rechenzeichen, die sind NumLock-unabhängig.
        /// </summary>
        private static void EnsureNumLockOn()
        {
            try
            {
                if (Control.IsKeyLocked(Keys.NumLock)) return;
                keybd_event(VK_NUMLOCK, 0, 0, UIntPtr.Zero);
                keybd_event(VK_NUMLOCK, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            }
            catch { }
        }

        /// <summary>Anzeigename der physischen Nummernblock-Taste für einen Numpad-Slot (0-14).</summary>
        private static string NumpadKeyLabel(int numpadSlot) => numpadSlot switch
        {
            >= 0 and <= 9 => string.Format(Properties.Resources.NumpadDigitLabel0, numpadSlot),
            10 => Properties.Resources.NumpadDivideLabel,
            11 => Properties.Resources.NumpadMultiplyLabel,
            12 => Properties.Resources.NumpadSubtractLabel,
            13 => Properties.Resources.NumpadAddLabel,
            14 => Properties.Resources.NumpadDecimalLabel,
            _  => "?"
        };

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            var s = Properties.Settings.Default;
            if (s.WindowWidth <= 0 || s.WindowHeight <= 0)
                try { this.CenterToScreen(); } catch { }
        }

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
                            this.BeginInvoke((System.Windows.Forms.MethodInvoker)(() =>
                            {
                                try
                                {
                                    if (this.WindowState == FormWindowState.Minimized || this.Opacity == 0.0)
                                        Reactivate();
                                    else
                                        MinimizeToTray();
                                }
                                catch (Exception ex) { Trace.WriteLine(String.Format(Wrok.Properties.Resources.HotkeyToggleFailed, ex)); }
                            }));
                        }
                        else if (id == HOTKEY_ID_IMAGE)
                        {
                            this.BeginInvoke((System.Windows.Forms.MethodInvoker)(() => _ = OpenImageFromClipboardAsync()));
                        }
                        else
                        {
                            this.BeginInvoke((System.Windows.Forms.MethodInvoker)(() => _ = PerformMacroAsync(id)));
                        }
                    }
                    catch { }
                    break;

                case WM_THEMECHANGED:
                case WM_SETTINGCHANGE:
                    try { _theme?.Refresh(); } catch { }
                    break;

                case WM_QUERYENDSESSION:
                    // Sitzungsende kuendigt sich an - merken, damit das folgende
                    // Schliessen wirklich beendet und nicht in den Tray minimiert.
                    _sessionEnding = true;
                    break;

                case WM_ENDSESSION:
                    // wParam == 0 heisst: ein zuvor angekuendigtes Sitzungsende
                    // wurde doch abgeblasen. Dann wieder normal in den Tray gehen.
                    if (m.WParam == IntPtr.Zero) _sessionEnding = false;
                    break;

                case WM_SHOWWINDOW:
                    try
                    {
                        this.BeginInvoke((System.Windows.Forms.MethodInvoker)(() =>
                        {
                            if (this.WindowState == FormWindowState.Minimized) this.WindowState = FormWindowState.Normal;
                            this.Show(); this.BringToFront(); this.Activate();
                        }));
                    }
                    catch { }
                    break;
            }
            base.WndProc(ref m);
        }

        // ------------------------------------------------------------------
        // Makro senden (zentral: Variablen + {input}-Auflösung)
        // ------------------------------------------------------------------

        // {input} oder {input:Eigene Frage}
        private static readonly System.Text.RegularExpressions.Regex InputTokenRegex =
            new(@"\{input(?::([^}]*))?\}",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase |
                System.Text.RegularExpressions.RegexOptions.Compiled);

        /// <summary>
        /// Löst Variablen und {input}-Platzhalter auf und sendet den Text ins WebView.
        /// Bricht ab (sendet nichts), wenn der Nutzer eine {input}-Abfrage abbricht.
        /// </summary>
        private async Task SendMacroTextAsync(string rawText, bool pressEnter)
        {
            if (_webViewManager == null || string.IsNullOrWhiteSpace(rawText)) return;

            string text = MacroManager.ExpandVariables(rawText);
            if (!TryResolveInputs(ref text)) return; // Nutzer hat abgebrochen

            await _webViewManager.SendTextAsync(text, pressEnter);
        }

        /// <summary>
        /// Ersetzt alle {input}/{input:Label}-Platzhalter durch Nutzereingaben.
        /// Gleiche Tokens werden nur einmal abgefragt. Gibt false zurück, wenn der
        /// Nutzer eine Abfrage abbricht (dann soll nichts gesendet werden).
        /// </summary>
        private bool TryResolveInputs(ref string text)
        {
            var matches = InputTokenRegex.Matches(text);
            if (matches.Count == 0) return true;

            var answers = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (System.Text.RegularExpressions.Match m in matches)
            {
                string token = m.Value;
                if (answers.ContainsKey(token)) continue;

                string label = m.Groups[1].Success && !string.IsNullOrWhiteSpace(m.Groups[1].Value)
                    ? m.Groups[1].Value.Trim()
                    : Properties.Resources.MacroInputPrompt;

                if (!Dialogs.ShowMultiLineInput(this, label, out string value)) return false;
                answers[token] = value;
            }

            foreach (var kv in answers)
                text = text.Replace(kv.Key, kv.Value);
            return true;
        }

        // ------------------------------------------------------------------
        // Bild aus Zwischenablage öffnen
        // ------------------------------------------------------------------

        /// <summary>
        /// Öffnet die Bild-URL aus der Zwischenablage (Rechtsklick im WebView →
        /// „Bildadresse kopieren") ohne Rückfrage direkt im Viewer-Fenster.
        /// Enthält die Zwischenablage keine gültige URL, wird auf das zuletzt
        /// gespeicherte Bild zurückgefallen – so macht Strg+Ö immer etwas Sinnvolles.
        /// </summary>
        private async Task OpenImageFromClipboardAsync()
        {
            if (_webViewManager == null) return;

            string url = string.Empty;
            try
            {
                if (Clipboard.ContainsText())
                    url = (Clipboard.GetText() ?? string.Empty).Trim();
            }
            catch (Exception ex) { Log(ex, "Zwischenablage konnte nicht gelesen werden"); }

            if (!IsHttpUrl(url))
            {
                OpenLastImage();
                return;
            }

            // Der Download kann einen Moment dauern – Wartecursor als Rückmeldung.
            try { Cursor.Current = Cursors.AppStarting; } catch { }
            try
            {
                // Grok-URLs haben keine Dateiendung – Typ über den Content-Type klären.
                var contentType = await _webViewManager.ProbeContentTypeAsync(url, TimeSpan.FromSeconds(10));

                if (contentType != null &&
                    contentType.Contains("video", StringComparison.OrdinalIgnoreCase))
                {
                    ShowVideoViewer(url);
                    return;
                }

                if (!await _webViewManager.ShowImageViewerAsync(url))
                    MessageBox.Show(this, Properties.Resources.ImageLoadFailed,
                        Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                try { Cursor.Current = Cursors.Default; } catch { }
            }
        }

        // ------------------------------------------------------------------
        // Medien anzeigen und nebeneinander anordnen
        // ------------------------------------------------------------------

        /// <summary>
        /// Zeigt ein Bild im Viewer-Fenster und ordnet es links auf voller
        /// Bildschirmhöhe an; Wrok füllt den Rest rechts daneben. Es gibt immer
        /// höchstens ein Medienfenster – ein neues Medium ersetzt das bisherige.
        /// Wrok bleibt nach dem Schließen in der Anordnung stehen: Medien haben
        /// meist dieselbe Größe, das nächste öffnet dadurch ohne Sprung.
        /// </summary>
        internal void ShowImageViewer(Bitmap bmp, string sourceUrl)
        {
            CloseCurrentMedia();

            var viewer = new ImageViewerForm(bmp, sourceUrl);
            RegisterMediaViewer(viewer);

            ArrangeSideBySide(viewer, viewer.WidthForHeight(WorkArea.Height));
            viewer.Show();
        }

        /// <summary>
        /// Zeigt ein Video im Endlos-Loop. Die echten Abmessungen sind erst nach
        /// dem Laden der Metadaten bekannt – das Fenster wird dann nachjustiert.
        /// </summary>
        internal void ShowVideoViewer(string url)
        {
            if (_webViewManager == null) return;

            CloseCurrentMedia();

            var viewer = new VideoViewerForm(url, _webViewManager.CoreEnvironment);
            RegisterMediaViewer(viewer);

            // Vorläufig hochkant anordnen; sobald die Metadaten da sind, korrigieren.
            ArrangeSideBySide(viewer, (int)(WorkArea.Height * 9.0 / 16.0));

            viewer.VideoSizeKnown += size =>
            {
                if (viewer.IsDisposed) return;
                try
                {
                    viewer.BeginInvoke((MethodInvoker)(() =>
                    {
                        if (!viewer.IsDisposed)
                            ArrangeSideBySide(viewer, viewer.WidthForHeight(WorkArea.Height, size));
                    }));
                }
                catch (Exception ex) { Log(ex, "Video-Anordnung nach Metadaten fehlgeschlagen"); }
            };

            viewer.Show();

            Properties.Settings.Default.LastImageUrl  = url;
            Properties.Settings.Default.LastImagePath = string.Empty;   // Video: nur URL
            Properties.Settings.Default.Save();
        }

        private Rectangle WorkArea => Screen.FromControl(this).WorkingArea;

        private void CloseCurrentMedia()
        {
            if (_mediaViewer is { IsDisposed: false } previous)
            {
                try { previous.Close(); }
                catch (Exception ex) { Log(ex, "Vorheriges Medienfenster schließen fehlgeschlagen"); }
            }
        }

        private void RegisterMediaViewer(Form viewer)
        {
            _mediaViewer = viewer;
            viewer.FormClosed += (_, _) =>
            {
                if (ReferenceEquals(_mediaViewer, viewer)) _mediaViewer = null;
            };
        }

        private void ArrangeSideBySide(Form viewer, int desiredWidth)
        {
            try
            {
                var wa = WorkArea;

                // Breite aus dem Seitenverhältnis bei voller Bildschirmhöhe –
                // aber gedeckelt, damit rechts genug für Wrok übrig bleibt.
                int viewerWidth = Math.Clamp(desiredWidth, 240, (int)(wa.Width * 0.6));

                viewer.StartPosition = FormStartPosition.Manual;
                viewer.Bounds = new Rectangle(wa.Left, wa.Top, viewerWidth, wa.Height);

                // Ist Wrok gar nicht sichtbar (Tray/minimiert), nur das Medium setzen.
                if (!this.Visible || this.WindowState == FormWindowState.Minimized)
                    return;

                try
                {
                    _suppressWindowSave = true;
                    if (this.WindowState != FormWindowState.Normal)
                        this.WindowState = FormWindowState.Normal;

                    this.Bounds = new Rectangle(wa.Left + viewerWidth, wa.Top,
                                                wa.Width - viewerWidth, wa.Height);
                }
                finally { _suppressWindowSave = false; }
            }
            catch (Exception ex)
            {
                Log(ex, "ArrangeSideBySide fehlgeschlagen");
            }
        }

        /// <summary>
        /// Öffnet das zuletzt angezeigte Bild aus dem lokalen Zwischenspeicher.
        /// Funktioniert ohne Netzwerk und ohne gültige Session, da die Bilddatei
        /// beim letzten Öffnen mitgespeichert wurde.
        /// </summary>
        private void OpenLastImage()
        {
            var path = Properties.Settings.Default.LastImagePath;
            var lastUrl = Properties.Settings.Default.LastImageUrl;

            // Kein lokaler Pfad, aber eine URL gemerkt → das war ein Video.
            // Videos werden nicht zwischengespeichert, sondern neu gestreamt.
            if (string.IsNullOrWhiteSpace(path) && IsHttpUrl(lastUrl))
            {
                ShowVideoViewer(lastUrl);
                return;
            }

            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                MessageBox.Show(this, Properties.Resources.NoLastImage,
                    Properties.Resources.OpenLastImage, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Datei-Stream sofort wieder freigeben: Bitmap aus einem Stream behält
                // sonst eine Referenz darauf und sperrt die Datei fürs Überschreiben.
                Bitmap bmp;
                using (var fs = File.OpenRead(path))
                using (var decoded = new Bitmap(fs))
                    bmp = new Bitmap(decoded);

                ShowImageViewer(bmp, string.IsNullOrWhiteSpace(lastUrl) ? path : lastUrl);
            }
            catch (Exception ex)
            {
                Log(ex, "OpenLastImage fehlgeschlagen");
                MessageBox.Show(this, Properties.Resources.ImageLoadFailed,
                    Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static bool IsHttpUrl(string? s) =>
            Uri.TryCreate(s, UriKind.Absolute, out var u) &&
            (u.Scheme == Uri.UriSchemeHttp || u.Scheme == Uri.UriSchemeHttps);

        // ------------------------------------------------------------------
        // Makro-Hotkey ausführen
        // ------------------------------------------------------------------

        private async Task PerformMacroAsync(int hotkeyId)
        {
            try
            {
                string? raw = _macroManager.GetRawTextForHotkey(hotkeyId);
                if (raw == null || _webViewManager == null) return;

                bool pressEnter = true;
                try
                {
                    var mods = ModifierKeys;
                    if ((mods & Keys.Shift) == Keys.Shift || (mods & Keys.Alt) == Keys.Alt)
                        pressEnter = false;
                }
                catch { }

                await SendMacroTextAsync(raw, pressEnter);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"PerformMacroAsync fehlgeschlagen: {ex}");
            }
        }
        // ------------------------------------------------------------------
        // Logging
        // ------------------------------------------------------------------

        private static void Log(Exception ex, string message) =>
            Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MainForm] {message}: {ex}");
    }
}
