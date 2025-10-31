using Microsoft.Web.WebView2.WinForms;
using System.Runtime.InteropServices;
using System.Drawing; // für Point/Size/Rectangle/Screen
using System.Net.Http;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Drawing.Imaging;

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
            (Properties.Resources.Settings,     "?_s=home"),
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

        // --- Inactivity timer ---
        private System.Windows.Forms.Timer? inactivityTimer;
        private TimeSpan inactivityTimeout = TimeSpan.FromSeconds(30); // Standardwert (konfigurierbar)
        private ActivityMessageFilter? activityFilter;
        private bool inactivityEnabled = true;
        private readonly int[] inactivityOptions = new[] { 0, 30, 60, 90 };
        // ------------------------

        public MainForm()
        {
            InitializeComponent();
            // Fenster-Einstellungen anwenden (bevor Controls initialisiert werden, aber nach InitializeComponent)
            LoadWindowSettings();

            // Inactivity-Einstellungen aus User-Settings übernehmen (wichtig: bevor Timer initialisiert wird)
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

            InitializeTrayIcon();
            InitializeWebView();
            InitializeInactivityTimer();
            LoadUrl(baseUrl);

            // Änderungen an Fensterzustand / Größe beobachten
            this.Resize += MainForm_Resize;
            this.ResizeEnd += MainForm_ResizeEnd; // speichere nach abgeschlossenem Resize
            this.Move += MainForm_Move;
        }

        private void InitializeComponent()
        {
            this.Text = "Wrok";
            this.Icon = Properties.Resources.wrok;
            this.WindowState = FormWindowState.Normal;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.ShowInTaskbar = false; // Nur im Tray
            this.Visible = false;       // Startet minimiert
        }

        // Lade gespeicherte Fensterposition/-größe (sicher für mehrere Bildschirme)
        private void LoadWindowSettings()
        {
            var s = Properties.Settings.Default;

            // Validieren
            if (s.WindowWidth > 0 && s.WindowHeight > 0)
            {
                this.StartPosition = FormStartPosition.Manual;
                var desired = new Rectangle(s.WindowLeft, s.WindowTop, s.WindowWidth, s.WindowHeight);

                // Prüfen ob die gespeicherte Position auf einem verfügbaren Screen liegt
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
                    // Falls nicht sichtbar, zentrieren
                    this.StartPosition = FormStartPosition.CenterScreen;
                }
            }

            if (s.IsMaximized)
            {
                this.WindowState = FormWindowState.Maximized;
            }
        }

        // Speichert die aktuellen Fenster-Settings (auch beim Minimieren verwendet)
        private void SaveWindowSettings()
        {
            try
            {
                var s = Properties.Settings.Default;
                Rectangle bounds;

                // Wenn maximiert oder minimiert, verwenden wir RestoreBounds, sonst aktuelle Bounds
                if (this.WindowState == FormWindowState.Maximized || this.WindowState == FormWindowState.Minimized)
                    bounds = this.RestoreBounds;
                else
                    bounds = this.Bounds;

                // Minimum-Größe sicherstellen
                s.WindowLeft = bounds.Left;
                s.WindowTop = bounds.Top;
                s.WindowWidth = Math.Max(100, bounds.Width);
                s.WindowHeight = Math.Max(100, bounds.Height);
                s.IsMaximized = (this.WindowState == FormWindowState.Maximized);

                s.Save();
            }
            catch
            {
                // Falls Speichern fehlschlägt, nicht kritisch für App-Funktion; swallow
            }
        }

        // Ereignisse, um beim Verschieben/Resize ggf. live zu speichern (optional, hier nur sanft)
        private void MainForm_Resize(object? sender, EventArgs e)
        {
            // Bei Minimieren/Maximieren speichern
            if (this.WindowState == FormWindowState.Minimized || this.WindowState == FormWindowState.Maximized)
            {
                SaveWindowSettings();
            }
        }

        // Neuer Handler: wird aufgerufen, wenn der Benutzer das Resize beendet hat.
        private void MainForm_ResizeEnd(object? sender, EventArgs e)
        {
            // Nur speichern, wenn Fenster im Normalzustand ist (wirkliche Größenänderung)
            if (this.WindowState == FormWindowState.Normal)
            {
                SaveWindowSettings();
            }
        }

        private void MainForm_Move(object? sender, EventArgs e)
        {
            // Bei manuellem Verschieben speichern (nicht zu häufig; hier einfach)
            if (this.WindowState == FormWindowState.Normal)
            {
                SaveWindowSettings();
            }
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

                // Interaktion detected → Timer zurücksetzen
                ResetInactivityTimer();
            };

            await webView.EnsureCoreWebView2Async(null);
            // Navigation wird jetzt in LoadUrl() mit Connectivity-Prüfung durchgeführt.

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
                    try { await webView.CoreWebView2.ExecuteScriptAsync(script); } catch { }
                }

                // Navigation zählt als Aktivität
                ResetInactivityTimer();
            };
        }
        private void InitializeTrayIcon()
        {
            trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add(Properties.Resources.ShowWindow, null, (s, e) => Reactivate());

            // Neuer Reload-Eintrag
            trayMenu.Items.Add(Properties.Resources.Reload, null, (s, e) =>
            {
                try
                {
                    webView?.CoreWebView2?.Reload();
                }
                catch
                {
                    // Ignoriere Fehler, falls WebView noch nicht initialisiert ist
                }
            });

            trayMenu.Items.Add(new ToolStripSeparator());

            // Inaktivitäts-Submenu (Radio-ähnlich)
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

            // Dynamische Buttons für jede URL
            foreach (var page in menuPages)
            {
                var item = trayMenu.Items.Add(page.name);
                item.Click += (s, e) => LoadUrl(baseUrl + page.url);
            }

            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add(Properties.Resources.Exit, null, (s, e) => Application.Exit());

            trayIcon = new NotifyIcon
            {
                Icon = Properties.Resources.wrok,
                Text = "Wrok",
                ContextMenuStrip = trayMenu,
                Visible = true
            };

            // Nur auf linken Mausklick reagieren
            trayIcon.MouseClick += (s, e) =>
            {
                if (e.Button == MouseButtons.Left)
                    Reactivate();
            };
        }

        private void Reactivate()
        {
            // Vor dem Anzeigen evtl. gespeicherte Größe/Position anwenden
            LoadWindowSettings();

            this.Show();
            this.WindowState = Properties.Settings.Default.IsMaximized ? FormWindowState.Maximized : FormWindowState.Normal;
            this.Opacity = 1.0;
            this.ShowInTaskbar = true;
            this.BringToFront();
            this.Activate();

            ResetInactivityTimer();
        }

        // LoadUrl wurde erweitert: prüft Internetverbindung, zeigt bei offline das Bild an
        private async void LoadUrl(string url)
        {
            try
            {
                if (webView == null) return;

                // Sicherstellen, dass CoreWebView2 initialisiert ist (versuchsweise)
                if (webView.CoreWebView2 == null)
                {
                    try { await webView.EnsureCoreWebView2Async(null); }
                    catch { /* falls Init fehlschlägt: weiter unten Offline-Seite zeigen */ }
                }

                var core = webView.CoreWebView2;
                if (core != null)
                {
                    // Robustere Konnektivitätsprüfung: mehrere Endpunkte, mehrere Versuche, Head→Get Fallback
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
                    // CoreWebView2 konnte nicht initialisiert werden
                    await ShowNoNetImageAsync();
                }
            }
            catch
            {
                try { await ShowNoNetImageAsync(); } catch { }
            }

            // UI sichtbar machen (wie vorher)
            this.Show();
            this.WindowState = Properties.Settings.Default.IsMaximized ? FormWindowState.Maximized : FormWindowState.Normal;
            this.Opacity = 1.0;
            this.ShowInTaskbar = true;
            this.BringToFront();
            this.Activate();

            ResetInactivityTimer();
        }

        private async Task<bool> HasInternetConnectionAsync(int attempts = 2, int timeoutSeconds = 4)
        {
            // Quick check: kein Netzwerk-Interface verfügbar → sofort false
            try
            {
                if (!System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable())
                    return false;
            }
            catch
            {
                // falls die Prüfung fehlschlägt, weiter mit HTTP-Checks
            }

            // Zuverlässige, kleine Testendpunkte
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
                        // Erst HEAD (schnell), falls nicht erfolgreich oder nicht erlaubt -> GET
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
                            // HEAD evtl. nicht erlaubt; versuche GET
                            using var req2 = new HttpRequestMessage(HttpMethod.Get, u);
                            using var resp2 = await client.SendAsync(req2, HttpCompletionOption.ResponseHeadersRead, cts.Token);
                            if (resp2.IsSuccessStatusCode || resp2.StatusCode == System.Net.HttpStatusCode.NoContent)
                                return true;
                        }
                    }
                    catch
                    {
                        // ignore und nächsten URL versuchen
                    }
                }

                // kurze Wartezeit vor erneutem Versuch (exponentiell)
                if (attempt + 1 < attempts)
                    await Task.Delay(300 + attempt * 200);
            }

            return false;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                // Speichern vor Minimieren/Verbergen
                SaveWindowSettings();

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

            // MessageFilter hinzufügen, sobald Handle existiert
            if (activityFilter == null)
            {
                activityFilter = new ActivityMessageFilter(this);
                Application.AddMessageFilter(activityFilter);
            }
        }

        // Unregister Hotkey beim Zerstören des Handles
        protected override void OnHandleDestroyed(EventArgs e)
        {
            // MessageFilter entfernen
            if (activityFilter != null)
            {
                try { Application.RemoveMessageFilter(activityFilter); } catch { }
                activityFilter = null;
            }

            UnregisterHotKey(this.Handle, HOTKEY_ID);
            base.OnHandleDestroyed(e);
        }

        // Hotkey-Ereignis abfangen
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                // Minimieren + in Tray versetzen
                MinimizeToTray();
            }
            base.WndProc(ref m);
        }

        private void MainForm_Load(object sender, EventArgs e) { }

        // -------------------------
        // Inactivity timer helpers
        // -------------------------
        private void InitializeInactivityTimer()
        {
            // Falls bereits vorhanden: alte Handler entfernen
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

            // Timer nur starten, wenn aktiviert UND ein gültiger (größer 0) Timeout gesetzt ist
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
            // Timer stoppen vor Aktion, damit kein Race entsteht
            inactivityTimer?.Stop();
            MinimizeToTray();
        }

        private void ResetInactivityTimer()
        {
            if (!inactivityEnabled) return;
            if (inactivityTimer == null) return;

            // Restart Timer
            inactivityTimer.Stop();
            inactivityTimer.Start();
        }

        private void MinimizeToTray()
        {
            // Speichern bevor minimiert/versteckt wird
            SaveWindowSettings();

            if (this.WindowState == FormWindowState.Minimized) return;

            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Opacity = 0;
        }

        /// <summary>
        /// Konfiguriert den Inaktivitäts-Timeout. Aufruf z.B. aus einer Settings-UI.
        /// </summary>
        public void SetInactivityTimeout(TimeSpan timeout)
        {
            inactivityTimeout = timeout;
            if (inactivityTimer != null)
            {
                inactivityTimer.Interval = (int)Math.Min(inactivityTimeout.TotalMilliseconds, int.MaxValue);
                ResetInactivityTimer();
            }
        }

        /// <summary>
        /// Aktiviert/deaktiviert den Inactivity-Timer (z.B. Pause-Funktion).
        /// </summary>
        public void EnableInactivityTimer(bool enabled)
        {
            inactivityEnabled = enabled;
            if (inactivityTimer == null) return;
            if (enabled) inactivityTimer.Start(); else inactivityTimer.Stop();
        }

        // MessageFilter zur Erkennung von Benutzerinteraktion innerhalb der App
        private class ActivityMessageFilter : IMessageFilter
        {
            private readonly WeakReference<MainForm> _formRef;

            public ActivityMessageFilter(MainForm form)
            {
                _formRef = new WeakReference<MainForm>(form);
            }

            public bool PreFilterMessage(ref Message m)
            {
                // relevante Nachrichten (Mouse + Keyboard)
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

                // Nicht abfangen, nur beobachten
                return false;
            }
        }

        // Neuer Event-Handler (innerhalb der MainForm-Klasse)
        private void InactivityMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is not ToolStripMenuItem clicked) return;
            int seconds = Convert.ToInt32(clicked.Tag ?? 0);

            // Einstellungen speichern
            Properties.Settings.Default.InactivityTimeoutSeconds = seconds;
            Properties.Settings.Default.Save();

            // Timer konfigurieren
            if (seconds > 0)
            {
                SetInactivityTimeout(TimeSpan.FromSeconds(seconds));
                EnableInactivityTimer(true);
            }
            else
            {
                EnableInactivityTimer(false);
            }

            // Radio-ähnliches Verhalten: alle Unteritems des Parents entchecken, nur das geklickte checken
            if (clicked.OwnerItem is ToolStripMenuItem parent)
            {
                foreach (ToolStripItem it in parent.DropDownItems)
                {
                    if (it is ToolStripMenuItem mi)
                        mi.Checked = mi == clicked;
                }
            }
        }

        private async Task ShowNoNetImageAsync()
        {
            if (webView == null) return;

            if (webView.CoreWebView2 == null)
            {
                try { await webView.EnsureCoreWebView2Async(null); } catch { /* ignore */ }
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
</style>
</head>
<body>
  <img src='data:image/png;base64,{base64}' alt='no network' />
</body>
</html>";
                var core = webView.CoreWebView2;
                if (core != null)
                    core.NavigateToString(html);
                // falls core null bleibt: nichts tun (kein Null-Reference)
            }
            catch
            {
                // Ignoriere Fehler beim Anzeigen der Offline-Seite
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
        // ----- Ende Netzwerk-Prüfungen -----
    }
}