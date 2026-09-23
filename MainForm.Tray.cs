using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Wrok
{
    public partial class MainForm
    {
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
            macroIoMenu.DropDownItems.Add(Properties.Resources.MacrosExportActive, null, (s, e) => ExportMacros(allProfiles: false));
            macroIoMenu.DropDownItems.Add(Properties.Resources.MacrosExportAll,    null, (s, e) => ExportMacros(allProfiles: true));
            macroIoMenu.DropDownItems.Add(new ToolStripSeparator());
            macroIoMenu.DropDownItems.Add(Properties.Resources.MacrosImport,       null, (s, e) => ImportMacros());
            appMenu.DropDownItems.Add(macroIoMenu);

            // F1-F10 direkt als Makro-Hotkeys, Strg+F5 = Reload, Strg+Umschalt+F1-3
            // wechselt das Profil - siehe MainForm.HandleFKeyOverride. Default aus.
            // Schaltet zusaetzlich live AreBrowserAcceleratorKeysEnabled um (siehe
            // WebViewManager.ConfigureCore) und synchronisiert den JS-Helper
            // (WrokHelper.js: __wrokSetFKeyEnabled), damit die Umstellung sofort
            // greift - kein Neustart/Reload noetig.
            var fKeyOverrideItem = new ToolStripMenuItem(Properties.Resources.FKeyOverrideMenuItem)
            {
                CheckOnClick = false,
                Checked      = Properties.Settings.Default.FKeyMacroOverrideEnabled
            };
            fKeyOverrideItem.ToolTipText = Properties.Resources.FKeyOverrideHint;
            fKeyOverrideItem.Click += (s, e) =>
            {
                bool desired = !fKeyOverrideItem.Checked;
                Properties.Settings.Default.FKeyMacroOverrideEnabled = desired;
                Properties.Settings.Default.Save();
                if (_webView?.CoreWebView2 != null)
                    _webView.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = !desired;
                _ = _webViewManager?.ExecuteScriptAsync(
                    $"if (window.__wrokSetFKeyEnabled) window.__wrokSetFKeyEnabled({(desired ? "true" : "false")});");
                fKeyOverrideItem.Checked = desired;
            };
            appMenu.DropDownItems.Add(fKeyOverrideItem);

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

    }
}
