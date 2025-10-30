using Microsoft.Web.WebView2.WinForms;
using System.Runtime.InteropServices;

namespace Wrok
{
    public partial class MainForm : Form
    {
        // Felder als nullable deklarieren, um CS8618 zu beheben
        private WebView2? webView;
        private NotifyIcon? trayIcon;
        private ContextMenuStrip? trayMenu;

        private readonly string baseUrl = "https://grok.com";
        private readonly (string name, string url)[] menuPages = new[]
        {
            (Properties.Resources.Account,     "?_s=account"),
            (Properties.Resources.Appearance,  "?_s=appearance"),
            (Properties.Resources.Behavior,    "?_s=behavior"),
            (Properties.Resources.Personality, "?_s=personality"),
            (Properties.Resources.Data,        "?_s=data"),
            (Properties.Resources.Billing,     "?_s=billing")
        };

        // --- Hotkey: Win32 RegisterHotKey ---
        private const int HOTKEY_ID = 0x9000;
        private const int WM_HOTKEY = 0x0312;
        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        // ---------------------------------------

        public MainForm()
        {
            InitializeComponent();
            InitializeTrayIcon();
            InitializeWebView();
            LoadUrl(baseUrl);
        }

        private void InitializeComponent()
        {
            this.Text = "Wrok";
            this.WindowState = FormWindowState.Normal;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new Size(1000, 700);
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.ShowInTaskbar = false; // Nur im Tray
            this.Visible = false;       // Startet minimiert
        }

        private async void InitializeWebView()
        {
            webView = new WebView2
            {
                Dock = DockStyle.Fill
            };
            this.Controls.Add(webView);

            webView.PreviewKeyDown += (s, e) =>
            {
                // Prüfen, ob Ctrl + Shift gleichzeitig gedrückt sind
                if ((ModifierKeys & (Keys.Control | Keys.Shift)) == (Keys.Control | Keys.Shift))
                {
                    this.WindowState = FormWindowState.Minimized;

                    // Optional: Falls du es komplett "ausblenden" willst
                    this.ShowInTaskbar = false;
                    this.Opacity = 0;

                    // Damit das Event nicht weitergereicht wird
                    e.IsInputKey = true;
                }
            };
            
            await webView.EnsureCoreWebView2Async(null);
            webView.CoreWebView2.Navigate("https://grok.com");

            // Sprachmodus aktivieren (via JS-Injection nach Navigation)
            webView.CoreWebView2.NavigationCompleted += async (sender, args) =>
            {
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
                    await webView.CoreWebView2.ExecuteScriptAsync(script);
                }
            };
        }
        private void InitializeTrayIcon()
        {
            trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add(Properties.Resources.ShowWindow, null, (s, e) => Reactivate());
            
            trayMenu.Items.Add(new ToolStripSeparator());

            // Dynamische Buttons für jede URL
            foreach (var page in menuPages)
            {
                var item = trayMenu.Items.Add(page.name);
                item.Click += (s, e) => LoadUrl(baseUrl + page.url);
            }
            // Exit

            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add(Properties.Resources.Exit, null, (s, e) => Application.Exit());


            trayIcon = new NotifyIcon
            {
                Icon = Properties.Resources.wrok,
                Text = "Wrok",
                ContextMenuStrip = trayMenu,
                Visible = true
            };

            trayIcon.DoubleClick += (s, e) => Reactivate();
        }

        private void Reactivate()
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.Opacity = 1.0;
            this.ShowInTaskbar = true;
            this.BringToFront();
            this.Activate();
        }
        private void LoadUrl(string url)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.Opacity = 1.0;
            this.ShowInTaskbar = true;
            this.BringToFront();
            this.Activate();
            webView?.CoreWebView2?.Navigate(url);
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.WindowState = FormWindowState.Minimized;
                this.Opacity = 0;
                this.ShowInTaskbar = false;
            }
            base.OnFormClosing(e);
        }

        // Register Hotkey, sobald Handle vorhanden
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);


            // Beispiel: Ctrl + Shift + Space
            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, (uint)Keys.Space);
            // Bei Bedarf andere Modifier/Key verwenden (z.B. MOD_ALT | (uint)Keys.F12)
        }

        // Unregister Hotkey beim Zerstören des Handles
        protected override void OnHandleDestroyed(EventArgs e)
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID);
            base.OnHandleDestroyed(e);
        }

        // Hotkey-Ereignis abfangen
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                // Minimieren + in Tray versetzen
                this.WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = false;
                this.Opacity = 0;
            }
            base.WndProc(ref m);
        }

        private void MainForm_Load(object sender, EventArgs e) { }
    }
}