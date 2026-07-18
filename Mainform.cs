using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Win32;
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
        private const int HOTKEY_ID_IMAGE = 0x9001;   // Strg+Shift+P
        private const int WM_HOTKEY    = 0x0312;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT   = 0x0004;

        private const int WM_THEMECHANGED  = 0x031A;
        private const int WM_SETTINGCHANGE = 0x001A;
        private const int WM_SHOWWINDOW    = 0x0018;

        private const int DWMWA_USE_IMMERSIVE_DARK_MODE          = 20;
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;

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

        // Inaktivitäts-Timer
        private System.Windows.Forms.Timer? inactivityTimer;
        private TimeSpan  inactivityTimeout = TimeSpan.FromSeconds(30);
        private bool      inactivityEnabled = true;
        private DateTime  _lastActivity     = DateTime.UtcNow;
        private readonly object _activityLock = new();

        private ActivityMessageFilter? activityFilter;
        private RateLimitManager?      _rateLimitManager;
        private ToolStripMenuItem?     _rateLimitMenu;

        // ------------------------------------------------------------------
        // P/Invoke
        // ------------------------------------------------------------------

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        // ------------------------------------------------------------------
        // Konstruktor
        // ------------------------------------------------------------------

        public MainForm()
        {
            InitializeComponent();

            activityFilter = new ActivityMessageFilter(this);
            try { Application.AddMessageFilter(activityFilter); }
            catch { activityFilter = null; }

            InitializeTrayIcon();
            RefreshTheme();
            LoadWindowSettings();
            ApplyInactivitySettings();

            _ = InitializeWebViewAsync();
            InitializeInactivityTimer();
            _ = InitializeRateLimitManagerAsync();

            try { _ = LoadUrlAsync(baseUrl, bringToFront: false); }
            catch (Exception ex) { Log(ex, "Initialer LoadUrlAsync fehlgeschlagen"); }

            this.Resize   += (s, e) => { if (this.WindowState != FormWindowState.Normal) SaveWindowSettings(); };
            this.ResizeEnd += (s, e) => { if (this.WindowState == FormWindowState.Normal) SaveWindowSettings(); };
            this.Move      += (s, e) => { if (this.WindowState == FormWindowState.Normal) SaveWindowSettings(); };

            UpdateTrayMenuInactivityState();
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

            _webViewManager = new WebViewManager(_webView, this, ResetInactivityTimer);
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
            ResetInactivityTimer();
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
            ResetInactivityTimer();

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
            this.WindowState   = FormWindowState.Minimized;
            this.Opacity       = 0;
            this.ShowInTaskbar = false;
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
            macroIoMenu.DropDownItems.Add(Properties.Resources.MacrosExport, null, (s, e) => ExportMacros());
            macroIoMenu.DropDownItems.Add(Properties.Resources.MacrosImport, null, (s, e) => ImportMacros());
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
                dlg.ShowDialog(this);
            });
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add(Properties.Resources.Exit, null, (s, e) => Application.Exit());

            var initialIcon = IsDarkMode() ? Properties.Resources.wrok_white : Properties.Resources.wrok_black;
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
                trayIcon.MouseClick += (s, e) => { if (e.Button == MouseButtons.Left) Reactivate(); };
        }

        private void EnsureTrayIconVisible()
        {
            try
            {
                if (trayIcon == null) { InitializeTrayIcon(); ApplyThemeIcon(); }
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

        private void RefreshMacrosMenu()
        {
            if (macrosMenu == null) return;
            macrosMenu.DropDownItems.Clear();

            var macros = _macroManager.GetMacros();
            for (int i = 1; i <= MacroManager.MacroCount; i++)
            {
                var entry      = macros[i - 1];
                int displayNum = i == 10 ? 0 : i;

                var macroItem = new ToolStripMenuItem(entry.DisplayName(displayNum)) { Tag = i - 1 };
                string hotkeyDisplay = string.Format(Properties.Resources.Ctrl0, displayNum);
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

                if (!ShowEditMacroDialog(dlgTitle, ref name, ref text)) return;

                _macroManager.UpdateMacro(index, new MacroEntry(name.Trim(), text.Trim()));
                RefreshMacrosMenu();
            }
            catch (Exception ex)
            {
                Log(ex, "EditMacroAndSave fehlgeschlagen");
            }
        }

        private bool ShowEditMacroDialog(string title, ref string name, ref string text)
        {
            using var dlg = new Form
            {
                Text             = title,
                FormBorderStyle  = FormBorderStyle.SizableToolWindow,
                StartPosition    = FormStartPosition.CenterParent,
                MinimizeBox      = false,
                MaximizeBox      = false,
                MinimumSize      = new Size(360, 260),
                ClientSize       = new Size(520, 320)
            };

            var lblName = new Label { Text = Properties.Resources.MacroName + ":", AutoSize = true, Location = new Point(10, 14), Font = new Font("Segoe UI", 9F) };
            var tbName  = new TextBox
            {
                Left            = lblName.Right + 6,
                Top             = 10,
                Width           = dlg.ClientSize.Width - lblName.Right - 16,
                Anchor          = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Text            = name ?? string.Empty,
                Font            = new Font("Segoe UI", 9F),
                PlaceholderText = string.Format(Properties.Resources.Macro0, "?")
            };
            int textTop = tbName.Bottom + 10;
            var tb = new TextBox
            {
                Multiline    = true,
                ScrollBars   = ScrollBars.Vertical,
                AcceptsReturn = false,
                AcceptsTab   = false,
                WordWrap     = true,
                Anchor       = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                Left         = 10,
                Top          = textTop,
                Width        = dlg.ClientSize.Width - 20,
                Height       = dlg.ClientSize.Height - textTop - 46,
                Text         = text ?? string.Empty,
                Font         = new Font("Segoe UI", 10F)
            };
            var lblCount  = new Label { AutoSize = true, Font = new Font("Segoe UI", 8F), ForeColor = SystemColors.GrayText, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Text = $"{tb.Text.Length} Zeichen" };
            var btnOk     = new Button { Text = Properties.Resources.OK,     DialogResult = DialogResult.OK,     Size = new Size(80, 26), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            var btnCancel = new Button { Text = Properties.Resources.Cancel, DialogResult = DialogResult.Cancel, Size = new Size(80, 26), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };

            var varTip = new ToolTip { AutoPopDelay = 15000, InitialDelay = 400, ReshowDelay = 200 };
            varTip.SetToolTip(tb, Properties.Resources.MacroVariablesHint);

            tb.TextChanged += (s, e) => lblCount.Text = $"{tb.Text.Length} Zeichen";

            void LayoutBottom()
            {
                int y = dlg.ClientSize.Height - btnOk.Height - 8;
                tb.Height       = y - tb.Top - 6;
                btnCancel.Location = new Point(dlg.ClientSize.Width - btnCancel.Width - 10, y);
                btnOk.Location     = new Point(btnCancel.Left - btnOk.Width - 6, y);
                lblCount.Location  = new Point(10, y + (btnOk.Height - lblCount.Height) / 2);
                tbName.Width       = dlg.ClientSize.Width - lblName.Right - 16;
            }

            dlg.Resize += (s, e) => LayoutBottom();
            dlg.Controls.AddRange(new Control[] { lblName, tbName, tb, lblCount, btnOk, btnCancel });

            tb.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter && !e.Shift) { e.SuppressKeyPress = true; dlg.DialogResult = DialogResult.OK; dlg.Close(); }
            };
            tbName.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; tb.Focus(); } };

            dlg.CancelButton = btnCancel;
            dlg.Shown += (s, e) =>
            {
                LayoutBottom();
                if (string.IsNullOrWhiteSpace(tbName.Text)) tbName.Focus();
                else { tb.SelectionStart = tb.Text.Length; tb.Focus(); }
            };
            btnOk.Click += (s, e) => dlg.Close();

            if (dlg.ShowDialog(this) == DialogResult.OK) { name = tbName.Text; text = tb.Text; return true; }
            return false;
        }

        // ------------------------------------------------------------------
        // Inaktivitäts-Timer
        // ------------------------------------------------------------------

        private void ApplyInactivitySettings()
        {
            if (!Properties.Settings.Default.InactivityConfigured)
            {
                inactivityTimeout = TimeSpan.Zero;
                inactivityEnabled = false;
                Properties.Settings.Default.InactivityTimeoutSeconds = 0;
                Properties.Settings.Default.InactivityConfigured     = true;
                Properties.Settings.Default.Save();
            }
            else
            {
                var savedSeconds  = Properties.Settings.Default.InactivityTimeoutSeconds;
                inactivityTimeout = TimeSpan.FromSeconds(savedSeconds);
                inactivityEnabled = savedSeconds > 0;
            }
        }

        private void InitializeInactivityTimer()
        {
            if (inactivityTimer != null)
            {
                try { inactivityTimer.Stop(); inactivityTimer.Tick -= InactivityTimer_Tick; }
                catch { }
                inactivityTimer = null;
            }
            inactivityTimer = new System.Windows.Forms.Timer
            {
                Interval = inactivityTimeout.TotalMilliseconds > 0
                    ? (int)Math.Min(1000, inactivityTimeout.TotalMilliseconds)
                    : 60_000
            };
            inactivityTimer.Tick += InactivityTimer_Tick;
            if (inactivityEnabled && inactivityTimeout.TotalMilliseconds > 0)
            {
                lock (_activityLock) { _lastActivity = DateTime.UtcNow; }
                inactivityTimer.Start();
            }
        }

        private void InactivityTimer_Tick(object? sender, EventArgs e)
        {
            if (!inactivityEnabled || inactivityTimeout.TotalMilliseconds <= 0) return;

            TimeSpan elapsed;
            lock (_activityLock) { elapsed = DateTime.UtcNow - _lastActivity; }

            try
            {
                if (this.Visible && (this.Focused || this.Bounds.Contains(Cursor.Position)))
                { lock (_activityLock) { _lastActivity = DateTime.UtcNow; } return; }
            }
            catch { }

            if (elapsed >= inactivityTimeout)
            {
                try { inactivityTimer?.Stop(); } catch { }
                MinimizeToTray();
            }
        }

        public void ResetInactivityTimer()
        {
            if (!inactivityEnabled || inactivityTimer == null) return;
            lock (_activityLock) { _lastActivity = DateTime.UtcNow; }
            try { if (!inactivityTimer.Enabled) inactivityTimer.Start(); } catch { }
        }

        private void InactivityMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is not ToolStripMenuItem clicked) return;
            int seconds = Convert.ToInt32(clicked.Tag ?? 0);

            Properties.Settings.Default.InactivityTimeoutSeconds = seconds;
            Properties.Settings.Default.Save();

            inactivityTimeout = TimeSpan.FromSeconds(seconds);
            inactivityEnabled = seconds > 0;
            if (inactivityEnabled) { lock (_activityLock) { _lastActivity = DateTime.UtcNow; } inactivityTimer?.Start(); }
            else try { inactivityTimer?.Stop(); } catch { }

            if (clicked.OwnerItem is ToolStripMenuItem parent)
                foreach (var item in parent.DropDownItems.OfType<ToolStripMenuItem>())
                    item.Checked = item == clicked;
        }

        // ------------------------------------------------------------------
        // Thema / Dark Mode
        // ------------------------------------------------------------------

        public static bool IsDarkMode()
        {
            try
            {
                var key   = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
                var value = key?.GetValue("AppsUseLightTheme");
                return value is int i && i == 0;
            }
            catch (Exception ex) { Trace.WriteLine($"IsDarkMode fallback: {ex}"); return true; }
        }

        private void RefreshTheme()
        {
            try { bool dark = IsDarkMode(); ApplyThemeIcon(); SetTitleBarDarkMode(dark); }
            catch { }
        }

        private void ApplyThemeIcon()
        {
            var sourceIcon = IsDarkMode() ? Properties.Resources.wrok_white : Properties.Resources.wrok_black;
            System.Drawing.Icon newIcon;
            try { newIcon = (System.Drawing.Icon)sourceIcon.Clone(); } catch { newIcon = sourceIcon; }

            if (trayIcon != null)
            {
                try
                {
                    var old = trayIcon.Icon;
                    trayIcon.Visible = false;
                    trayIcon.Icon    = newIcon;
                    trayIcon.Visible = true;
                    if (old != null && !ReferenceEquals(old, sourceIcon)) try { old.Dispose(); } catch { }
                }
                catch { try { trayIcon.Icon = newIcon; } catch { } }
            }
            try { this.Icon = (System.Drawing.Icon)newIcon.Clone(); } catch { this.Icon = newIcon; }
        }

        private void SetTitleBarDarkMode(bool enabled)
        {
            try
            {
                int val = enabled ? 1 : 0;
                int hr  = DwmSetWindowAttribute(this.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref val, Marshal.SizeOf<int>());
                if (hr != 0)
                    try { DwmSetWindowAttribute(this.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1, ref val, Marshal.SizeOf<int>()); } catch { }
            }
            catch { }
        }

        private void SystemEvents_UserPreferenceChanged(object? sender, UserPreferenceChangedEventArgs e)
        {
            if (e.Category == UserPreferenceCategory.Color ||
                e.Category == UserPreferenceCategory.General ||
                e.Category == UserPreferenceCategory.VisualStyle)
                try { if (!this.IsDisposed) this.BeginInvoke((MethodInvoker)RefreshTheme); } catch { }
        }

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

        private void ExportMacros()
        {
            using var dlg = new SaveFileDialog
            {
                Title            = "Makros exportieren",
                Filter           = "JSON-Datei (*.json)|*.json",
                FileName         = $"Wrok-Makros_{DateTime.Now:yyyyMMdd}.json",
                DefaultExt       = "json",
                OverwritePrompt  = true
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            bool ok = _macroManager.Export(dlg.FileName);
            if (ok)
                MessageBox.Show($"Makros erfolgreich exportiert nach:\n{dlg.FileName}",
                    "Export erfolgreich", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show("Export fehlgeschlagen. Bitte Pfad und Berechtigungen prüfen.",
                    "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void ImportMacros()
        {
            using var dlg = new OpenFileDialog
            {
                Title  = "Makros importieren",
                Filter = "JSON-Datei (*.json)|*.json",
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            var confirm = MessageBox.Show(
                "Die aktuellen Makros werden durch die importierten ersetzt.\nFortfahren?",
                "Makros importieren", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes) return;

            bool ok = _macroManager.Import(dlg.FileName);
            if (ok)
            {
                RefreshMacrosMenu();
                MessageBox.Show("Makros erfolgreich importiert.",
                    "Import erfolgreich", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show("Import fehlgeschlagen. Bitte prüfen ob die Datei ein gültiges Wrok-Makro-Format hat.",
                    "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        // ------------------------------------------------------------------
        // Cache leeren
        // ------------------------------------------------------------------

        private enum ClearCacheChoice { Cancel, CacheOnly, All }

        private async Task ClearCacheAsync()
        {
            try
            {
                if (_webView?.CoreWebView2 == null)
                { MessageBox.Show(Properties.Resources.WebView2IsNotInitializedYet, Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var choice = ShowClearCacheChoiceDialog();
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
                    // Alles (inkl. Cookies, Login-Daten, lokalem Speicher).
                    await _webView.CoreWebView2.Profile.ClearBrowsingDataAsync();
                    MessageBox.Show(
                        Properties.Resources.CookiesLoginDataAndSettingsDeletedNYouAreLoggedOut,
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
        /// Zeigt einen kleinen Dialog mit drei Schaltflächen
        /// (Nur Cache / Alles löschen / Abbrechen) und gibt die Auswahl zurück.
        /// </summary>
        private ClearCacheChoice ShowClearCacheChoiceDialog()
        {
            using var dlg = new Form
            {
                Text            = Properties.Resources.ClearBrowsingData,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition   = FormStartPosition.CenterParent,
                MinimizeBox     = false,
                MaximizeBox     = false,
                ShowInTaskbar   = false,
                ClientSize      = new Size(420, 120)
            };

            var lbl = new Label
            {
                Text     = Properties.Resources.ChooseWhatToBeDeleted,
                AutoSize = false,
                Left     = 16,
                Top      = 16,
                Width    = dlg.ClientSize.Width - 32,
                Height   = 40,
                Font     = new Font("Segoe UI", 9.5F)
            };

            var btnCacheOnly = new Button { Text = Properties.Resources.ClearCacheOnly, DialogResult = DialogResult.Yes,    Size = new Size(120, 30), Top = 70 };
            var btnAll       = new Button { Text = Properties.Resources.DeleteAll,      DialogResult = DialogResult.No,     Size = new Size(120, 30), Top = 70 };
            var btnCancel    = new Button { Text = Properties.Resources.Cancel,         DialogResult = DialogResult.Cancel, Size = new Size(90,  30), Top = 70 };

            btnCacheOnly.Left = 16;
            btnAll.Left       = btnCacheOnly.Right + 8;
            btnCancel.Left    = dlg.ClientSize.Width - btnCancel.Width - 16;

            dlg.Controls.AddRange(new Control[] { lbl, btnCacheOnly, btnAll, btnCancel });
            dlg.AcceptButton = btnCacheOnly;
            dlg.CancelButton = btnCancel;

            return dlg.ShowDialog(this) switch
            {
                DialogResult.Yes => ClearCacheChoice.CacheOnly,
                DialogResult.No  => ClearCacheChoice.All,
                _                => ClearCacheChoice.Cancel
            };
        }

        // ------------------------------------------------------------------
        // Form-Events
        // ------------------------------------------------------------------

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                SaveWindowSettings();
                e.Cancel           = true;
                this.WindowState   = FormWindowState.Minimized;
                this.Opacity       = 0;
                this.ShowInTaskbar = false;
            }
            base.OnFormClosing(e);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            SystemEvents.UserPreferenceChanged += SystemEvents_UserPreferenceChanged;
            base.OnHandleCreated(e);

            _webViewManager?.CreateInputSimulator();

            bool ok = RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, (uint)Keys.Space);
            if (!ok) Debug.WriteLine($"RegisterHotKey fehlgeschlagen id={HOTKEY_ID} err={Marshal.GetLastWin32Error()}");

            // Strg+Shift+P → Bild aus Zwischenablage öffnen
            ok = RegisterHotKey(this.Handle, HOTKEY_ID_IMAGE, MOD_CONTROL | MOD_SHIFT, (uint)Keys.P);
            if (!ok) Trace.WriteLine($"RegisterHotKey fehlgeschlagen id={HOTKEY_ID_IMAGE} (Strg+Shift+P) err={Marshal.GetLastWin32Error()}");

            // Makro 0 (intern) → Strg+1, Makro 1 → Strg+2, ..., Makro 8 → Strg+9, Makro 9 → Strg+0
            var macroKeys = new uint[]
            {
                (uint)Keys.D1, (uint)Keys.D2, (uint)Keys.D3, (uint)Keys.D4, (uint)Keys.D5,
                (uint)Keys.D6, (uint)Keys.D7, (uint)Keys.D8, (uint)Keys.D9, (uint)Keys.D0
            };
            for (int i = 0; i < MacroManager.MacroCount; i++)
            {
                ok = RegisterHotKey(this.Handle, MacroManager.HotkeyBase + i, MOD_CONTROL, macroKeys[i]);
                if (!ok) Debug.WriteLine($"RegisterHotKey fehlgeschlagen macro={i + 1} err={Marshal.GetLastWin32Error()}");
            }

            RefreshTheme();
            EnsureTrayIconVisible();
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            _rateLimitManager?.StopAutoRefresh();
            if (activityFilter != null)
            {
                try { Application.RemoveMessageFilter(activityFilter); } catch { }
                activityFilter = null;
            }
            try { UnregisterHotKey(this.Handle, HOTKEY_ID); } catch { }
            try { UnregisterHotKey(this.Handle, HOTKEY_ID_IMAGE); } catch { }
            for (int i = 0; i < MacroManager.MacroCount; i++)
                try { UnregisterHotKey(this.Handle, MacroManager.HotkeyBase + i); } catch { }

            _webViewManager?.DisposeInputSimulator();
            SystemEvents.UserPreferenceChanged -= SystemEvents_UserPreferenceChanged;
            if (!this.RecreatingHandle) DisposeTrayIcon();
            base.OnHandleDestroyed(e);
        }

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
                    try { RefreshTheme(); } catch { }
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

                if (!ShowInputDialog(label, out string value)) return false;
                answers[token] = value;
            }

            foreach (var kv in answers)
                text = text.Replace(kv.Key, kv.Value);
            return true;
        }

        /// <summary>Kleiner einzeiliger Eingabedialog. Enter = OK, Esc = Abbrechen.</summary>
        private bool ShowInputDialog(string prompt, out string value)
        {
            value = string.Empty;

            // Wird das Makro per globalem Hotkey ausgelöst, ist Wrok evtl. minimiert
            // oder im Hintergrund. Dann muss der Dialog selbst nach vorn kommen –
            // sonst erscheint er hinter dem aktiven Fenster und bekommt keinen Fokus.
            bool ownerUsable = this.Visible && this.WindowState != FormWindowState.Minimized;

            using var dlg = new Form
            {
                Text            = Properties.Resources.MacroInputTitle,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition   = ownerUsable ? FormStartPosition.CenterParent
                                              : FormStartPosition.CenterScreen,
                MinimizeBox     = false,
                MaximizeBox     = false,
                ShowInTaskbar   = false,
                TopMost         = true,
                ClientSize      = new Size(420, 110)
            };

            var lbl = new Label
            {
                Text     = prompt,
                AutoSize = false,
                Left     = 16,
                Top      = 14,
                Width    = dlg.ClientSize.Width - 32,
                Height   = 20,
                Font     = new Font("Segoe UI", 9.5F)
            };
            var tb = new TextBox
            {
                Left   = 16,
                Top    = 40,
                Width  = dlg.ClientSize.Width - 32,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font   = new Font("Segoe UI", 10F)
            };
            var btnOk     = new Button { Text = Properties.Resources.OK,     DialogResult = DialogResult.OK,     Size = new Size(80, 28), Top = 72 };
            var btnCancel = new Button { Text = Properties.Resources.Cancel, DialogResult = DialogResult.Cancel, Size = new Size(80, 28), Top = 72 };
            btnCancel.Left = dlg.ClientSize.Width - btnCancel.Width - 16;
            btnOk.Left     = btnCancel.Left - btnOk.Width - 8;

            dlg.Controls.AddRange(new Control[] { lbl, tb, btnOk, btnCancel });
            dlg.AcceptButton = btnOk;
            dlg.CancelButton = btnCancel;
            dlg.Shown += (s, e) =>
            {
                try
                {
                    dlg.Activate();
                    SetForegroundWindow(dlg.Handle);
                }
                catch (Exception ex) { Trace.WriteLine($"ShowInputDialog: Fokus fehlgeschlagen: {ex}"); }
                tb.Focus();
            };

            // Ohne sichtbaren Owner ohne Owner-Fenster anzeigen, sonst hängt der
            // modale Dialog an einem minimierten/unsichtbaren Fenster.
            var result = ownerUsable ? dlg.ShowDialog(this) : dlg.ShowDialog();
            if (result != DialogResult.OK) return false;
            value = tb.Text;
            return true;
        }

        // ------------------------------------------------------------------
        // Bild aus Zwischenablage öffnen
        // ------------------------------------------------------------------

        /// <summary>
        /// Öffnet die Bild-URL aus der Zwischenablage (Rechtsklick im WebView →
        /// „Bildadresse kopieren") ohne Rückfrage direkt im Viewer-Fenster.
        /// Enthält die Zwischenablage keine gültige URL, wird auf das zuletzt
        /// gespeicherte Bild zurückgefallen – so macht Strg+Shift+P immer etwas Sinnvolles.
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
        // ActivityMessageFilter (innere Klasse)
        // ------------------------------------------------------------------

        private class ActivityMessageFilter : IMessageFilter
        {
            private readonly WeakReference<MainForm> _formRef;
            public ActivityMessageFilter(MainForm form) => _formRef = new(form);

            public bool PreFilterMessage(ref Message m)
            {
                const int WM_MOUSEMOVE   = 0x0200;
                const int WM_LBUTTONDOWN = 0x0201;
                const int WM_RBUTTONDOWN = 0x0204;
                const int WM_MBUTTONDOWN = 0x0207;
                const int WM_MOUSEWHEEL  = 0x020A;
                const int WM_KEYDOWN     = 0x0100;
                const int WM_SYSKEYDOWN  = 0x0104;

                if (m.Msg is WM_MOUSEMOVE or WM_LBUTTONDOWN or WM_RBUTTONDOWN or
                             WM_MBUTTONDOWN or WM_MOUSEWHEEL or WM_KEYDOWN or WM_SYSKEYDOWN)
                    if (_formRef.TryGetTarget(out var target))
                        target.ResetInactivityTimer();

                return false;
            }
        }

        // ------------------------------------------------------------------
        // Logging
        // ------------------------------------------------------------------

        private static void Log(Exception ex, string message) =>
            Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MainForm] {message}: {ex}");
    }
}
