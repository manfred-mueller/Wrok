using Microsoft.Web.WebView2.WinForms;
using System;
using System.Drawing;
using System.Windows.Forms;

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
                            console.log('Sprachmodus aktiviert');
                        } else {
                            console.warn('Sprachmodus-Button nicht gefunden');
                        }
                    ";
                    await webView.CoreWebView2.ExecuteScriptAsync(script);
                }
            };
        }
        private void InitializeTrayIcon()
        {
            trayMenu = new ContextMenuStrip();

            // Exit
            trayMenu.Items.Add(Properties.Resources.Exit, null, (s, e) => Application.Exit());

            trayMenu.Items.Add(new ToolStripSeparator());

            // Dynamische Buttons für jede URL
            foreach (var page in menuPages)
            {
                var item = trayMenu.Items.Add(page.name);
                item.Click += (s, e) => LoadUrl(baseUrl + page.url);
            }

            trayIcon = new NotifyIcon
            {
                Icon = Properties.Resources.wrok,
                Text = "Wrok - Grok Einstellungen",
                ContextMenuStrip = trayMenu,
                Visible = true
            };

            trayIcon.DoubleClick += (s, e) => LoadUrl(baseUrl);
        }

        private void LoadUrl(string url)
        {
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
                this.Hide();
            }
            base.OnFormClosing(e);
        }

        private void MainForm_Load(object sender, EventArgs e) { }
    }
}