using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Wrok
{
    public partial class MainForm
    {
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
                    TryCenterToScreen();
                }
            }
            else if (hasValidSize)
            {
                this.Size = new Size(s.WindowWidth, s.WindowHeight);
                this.StartPosition = FormStartPosition.CenterScreen;
                TryCenterToScreen();
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

        /// <summary>
        /// CenterToScreen kann u. a. wegen einer inzwischen entfernten/verkleinerten
        /// Monitor-Konfiguration fehlschlagen - dann bleibt die zuvor gesetzte Size/
        /// StartPosition bestehen, das ist ein akzeptabler Fallback statt eines Absturzes.
        /// </summary>
        private void TryCenterToScreen()
        {
            try { this.CenterToScreen(); } catch { }
        }

        private void Reactivate()
        {
            LoadWindowSettings();
            if (this.StartPosition == FormStartPosition.CenterScreen)
                TryCenterToScreen();

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

            // Strg+Ö → Bild aus Zwischenablage öffnen.
            // WICHTIG: Keys.Oem1 (VK_OEM_1, Scancode 1A) ist auf der deutschen
            // Tastatur die Ü-Taste, NICHT Ö - das war der eigentliche Bug (per
            // kbdlayout.info/KBDGR verifiziert: Ö liegt auf Scancode 27 =
            // VK_OEM_3 = Keys.Oem3). Mit Keys.Oem1 registrierte sich der Hotkey
            // zwar erfolgreich, reagierte aber nur auf Strg+Ü statt Strg+Ö.
            ok = RegisterHotKey(this.Handle, HOTKEY_ID_IMAGE, MOD_CONTROL, (uint)Keys.Oem3);
            if (!ok) Trace.WriteLine($"RegisterHotKey fehlgeschlagen id={HOTKEY_ID_IMAGE} (Strg+Ö) err={Marshal.GetLastWin32Error()}");

            _theme?.Refresh();
            EnsureTrayIconVisible();
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            _rateLimitManager?.StopAutoRefresh();
            try { UnregisterHotKey(this.Handle, HOTKEY_ID); } catch { }
            try { UnregisterHotKey(this.Handle, HOTKEY_ID_IMAGE); } catch { }

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

    }
}
