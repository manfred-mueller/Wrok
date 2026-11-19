using System.Diagnostics;

namespace Wrok
{
    public partial class MainForm
    {
        // ------------------------------------------------------------------
        // Medienschutz: Sperre beim Minimieren, EFS, Slots nur aus dem Medienordner
        // (Logik der Ordner selbst: siehe MediaProtection.cs)
        // ------------------------------------------------------------------

        // null = noch nie angewendet; beim ersten Aufruf wird der Zustand in jedem
        // Fall geschrieben, damit eine nach Absturz oder Abschalten liegengebliebene
        // Sperre verschwindet.
        private bool? _mediaLocked;

        /// <summary>Beim Start: alte Ordner umziehen und den Sperrzustand herstellen.</summary>
        private void InitializeMediaProtection()
        {
            MediaProtection.MigrateLegacyFolders();
            try { MediaProtection.EnsureMediaDir(); }
            catch (Exception ex) { Log(ex, "Medienordner anlegen fehlgeschlagen"); }

            this.VisibleChanged += (s, e) => UpdateMediaLock();
            this.Resize         += (s, e) => UpdateMediaLock();
            UpdateMediaLock();
        }

        /// <summary>
        /// Sperrt den Medienordner, solange Wrok nicht sichtbar ist (Tray, minimiert,
        /// noch nicht angezeigt), und entsperrt ihn, sobald Wrok wieder da ist.
        /// </summary>
        private void UpdateMediaLock()
        {
            bool hidden = !this.Visible || this.WindowState == FormWindowState.Minimized || this.Opacity == 0;
            bool lockIt = Properties.Settings.Default.MediaLockEnabled && hidden;

            if (_mediaLocked == lockIt) return;
            if (MediaProtection.SetLocked(lockIt))
                _mediaLocked = lockIt;
        }

        /// <summary>
        /// Führt eine Aktion mit kurz entsperrtem Medienordner aus (z. B. Auflisten
        /// für EFS oder Export) und stellt danach den passenden Zustand wieder her.
        /// </summary>
        private T WithMediaUnlocked<T>(Func<T> action)
        {
            MediaProtection.SetLocked(false);
            _mediaLocked = false;
            try { return action(); }
            finally
            {
                _mediaLocked = null;
                if (InvokeRequired) Invoke((System.Windows.Forms.MethodInvoker)UpdateMediaLock);
                else UpdateMediaLock();
            }
        }

        /// <summary>Verschlüsselt beim Start die Browserdaten bzw. hebt das wieder auf.</summary>
        private static async Task ApplyWebViewDataEfsAtStartupAsync()
        {
            bool wanted = Properties.Settings.Default.WebViewDataEfsEnabled;
            string dir  = MediaProtection.WebViewDataDir;

            // Nichts zu tun: aus und nie verschlüsselt.
            if (!wanted && !MediaProtection.IsEncrypted(dir)) return;

            // Muss vor dem Start von WebView2 laufen, sonst sind die Dateien geöffnet.
            var result = await Task.Run(() => MediaProtection.ApplyEfs(dir, wanted));
            Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MainForm] EFS Browserdaten ({(wanted ? "ein" : "aus")}): " +
                            $"{result.Done} ok, {result.Failed} fehlgeschlagen, Fehler: {result.RootError ?? "-"}");
        }

        // ------------------------------------------------------------------
        // Tray-Menü
        // ------------------------------------------------------------------

        private ToolStripMenuItem BuildMediaProtectionMenu()
        {
            var menu = new ToolStripMenuItem(Properties.Resources.MediaProtectionMenu);
            menu.DropDown.ShowItemToolTips = true;

            var lockItem = new ToolStripMenuItem(Properties.Resources.MediaLockMenuItem)
            {
                Checked     = Properties.Settings.Default.MediaLockEnabled,
                ToolTipText = Properties.Resources.MediaLockHint
            };
            lockItem.Click += (s, e) =>
            {
                bool desired = !lockItem.Checked;
                Properties.Settings.Default.MediaLockEnabled = desired;
                Properties.Settings.Default.Save();
                lockItem.Checked = desired;
                UpdateMediaLock();
            };
            menu.DropDownItems.Add(lockItem);

            var efsItem = new ToolStripMenuItem(Properties.Resources.MediaEfsMenuItem)
            {
                Checked = Properties.Settings.Default.MediaEfsEnabled
            };
            efsItem.Click += async (s, e) =>
            {
                efsItem.Enabled = false;
                try
                {
                    bool desired = !efsItem.Checked;
                    if (await ApplyMediaEfsAsync(desired))
                    {
                        Properties.Settings.Default.MediaEfsEnabled = desired;
                        Properties.Settings.Default.Save();
                        efsItem.Checked = desired;
                    }
                }
                finally { efsItem.Enabled = true; }
            };
            menu.DropDownItems.Add(efsItem);

            var webEfsItem = new ToolStripMenuItem(Properties.Resources.WebViewEfsMenuItem)
            {
                Checked     = Properties.Settings.Default.WebViewDataEfsEnabled,
                ToolTipText = Properties.Resources.WebViewEfsHint
            };
            webEfsItem.Click += (s, e) =>
            {
                bool desired = !webEfsItem.Checked;
                Properties.Settings.Default.WebViewDataEfsEnabled = desired;
                Properties.Settings.Default.Save();
                webEfsItem.Checked = desired;
                MessageBox.Show(this, Properties.Resources.WebViewEfsRestartInfo, Properties.Resources.MediaProtectionMenu,
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (desired) OfferEfsKeyBackup();
            };
            menu.DropDownItems.Add(webEfsItem);

            menu.DropDownItems.Add(new ToolStripSeparator());
            menu.DropDownItems.Add(Properties.Resources.EfsBackupMenuItem, null, (s, e) => StartEfsKeyBackup());

            return menu;
        }

        /// <summary>Schaltet EFS für den Medienordner um. Liefert false bei Totalausfall.</summary>
        private async Task<bool> ApplyMediaEfsAsync(bool encrypt)
        {
            MediaProtection.EfsResult result;
            try { Cursor.Current = Cursors.WaitCursor; } catch { }
            try
            {
                result = await Task.Run(() => WithMediaUnlocked(() =>
                    MediaProtection.ApplyEfs(MediaProtection.MediaDir, encrypt)));
            }
            catch (Exception ex)
            {
                Log(ex, "EFS Medienordner fehlgeschlagen");
                result = new MediaProtection.EfsResult(0, 0, ex.Message);
            }
            finally
            {
                try { Cursor.Current = Cursors.Default; } catch { }
            }

            if (result.RootError != null)
            {
                MessageBox.Show(this, string.Format(Properties.Resources.EfsUnavailable, result.RootError),
                                Properties.Resources.MediaProtectionMenu, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            string text = result.Failed > 0
                ? string.Format(Properties.Resources.EfsPartial, result.Done, result.Failed)
                : string.Format(encrypt ? Properties.Resources.EfsEncryptDone : Properties.Resources.EfsDecryptDone, result.Done);
            MessageBox.Show(this, text, Properties.Resources.MediaProtectionMenu, MessageBoxButtons.OK,
                            result.Failed > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);

            if (encrypt) OfferEfsKeyBackup();
            return true;
        }

        private void OfferEfsKeyBackup()
        {
            if (MessageBox.Show(this, Properties.Resources.EfsBackupPrompt, Properties.Resources.MediaProtectionMenu,
                                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                StartEfsKeyBackup();
        }

        private void StartEfsKeyBackup()
        {
            if (!MediaProtection.StartKeyBackup())
                MessageBox.Show(this, Properties.Resources.EfsBackupFailed, Properties.Resources.MediaProtectionMenu,
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>
        /// Öffnet den Medienordner im Explorer. Wrok wird dafür angezeigt – der
        /// Ordner ist nur offen, solange Wrok sichtbar ist.
        /// </summary>
        private void OpenMediaFolder()
        {
            try
            {
                Reactivate();
                var dir = MediaProtection.EnsureMediaDir();
                Process.Start(new ProcessStartInfo("explorer.exe", $"\"{dir}\"") { UseShellExecute = true });
            }
            catch (Exception ex) { Log(ex, "Medienordner öffnen fehlgeschlagen"); }
        }

        /// <summary>Kopiert alle Medien unverschlüsselt in einen gewählten Ordner.</summary>
        private async void ExportMedia()
        {
            try
            {
                using var dlg = new FolderBrowserDialog
                {
                    Description            = Properties.Resources.ExportMediaDescription,
                    UseDescriptionForTitle = true,
                    ShowNewFolderButton    = true
                };
                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                string target = dlg.SelectedPath;
                var (copied, failed) = await Task.Run(() => WithMediaUnlocked(() => MediaProtection.ExportTo(target)));

                string text = failed > 0
                    ? string.Format(Properties.Resources.ExportMediaPartial, copied, failed)
                    : string.Format(Properties.Resources.ExportMediaDone, copied);
                MessageBox.Show(this, text, Properties.Resources.ExportMediaMenuItem, MessageBoxButtons.OK,
                                failed > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log(ex, "Medien exportieren fehlgeschlagen");
                MessageBox.Show(this, ex.Message, Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ------------------------------------------------------------------
        // Medien-Slots: nur Dateien aus dem Medienordner
        // ------------------------------------------------------------------

        /// <summary>
        /// Prüft eine Slot-Adresse. Links bleiben unverändert erlaubt. Eine lokale
        /// Datei außerhalb des Medienordners wird nach Rückfrage hineinkopiert; die
        /// Rückgabe zeigt dann auf die Kopie (Haltepunkte bleiben erhalten). Liegt
        /// dieselbe Datei (gleicher Inhalt) schon im Medienordner, wird sie verwendet
        /// statt eine weitere Kopie anzulegen.
        /// Liefert null, wenn der Nutzer ablehnt oder das Kopieren scheitert.
        /// </summary>
        private string? EnsureSlotInMediaFolder(string url)
        {
            if (string.IsNullOrWhiteSpace(url)) return url;

            string bare   = VideoViewerForm.SplitStops(url, out _);
            string suffix = url.Substring(bare.Length);

            if (IsHttpUrl(bare)) return url;

            var local = TryParseLocalPath(bare);
            if (local == null || MediaProtection.IsInMediaDir(local)) return url;

            if (MessageBox.Show(this, string.Format(Properties.Resources.MediaCopyIntoProtectedPrompt, local),
                                Properties.Resources.MediaProtectionMenu, MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question) != DialogResult.Yes)
                return null;

            try
            {
                var existing = WithMediaUnlocked(() => MediaProtection.FindIdenticalFile(local));
                return (existing ?? MediaProtection.CopyIntoMediaDir(local)) + suffix;
            }
            catch (Exception ex)
            {
                Log(ex, "Kopieren in den Medienordner fehlgeschlagen");
                MessageBox.Show(this, string.Format(Properties.Resources.MediaCopyFailed, ex.Message),
                                Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// Nach dem Ändern oder Leeren eines Slots: Zeigte er auf eine Datei im
        /// Medienordner, die jetzt kein Slot mehr verwendet, wird angeboten, sie
        /// endgültig zu löschen (nicht in den Papierkorb, damit sie nicht außerhalb
        /// des geschützten Ordners wieder auftaucht).
        /// </summary>
        private void OfferDeleteUnusedMedia(string oldUrl)
        {
            var path = LocalMediaPathOf(oldUrl);
            if (path == null || !MediaProtection.IsInMediaDir(path)) return;

            var full = Path.GetFullPath(path);
            bool stillUsed = VideoSlotStore.Get().Any(s =>
                LocalMediaPathOf(s.Url) is string p &&
                string.Equals(Path.GetFullPath(p), full, StringComparison.OrdinalIgnoreCase));
            if (stillUsed) return;

            if (MessageBox.Show(this, string.Format(Properties.Resources.MediaDeleteUnusedPrompt, full),
                                Properties.Resources.MediaProtectionMenu, MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes)
                return;

            try
            {
                File.Delete(full);
            }
            catch (Exception ex)
            {
                Log(ex, "Unbenutzte Mediendatei loeschen fehlgeschlagen");
                MessageBox.Show(this, string.Format(Properties.Resources.MediaDeleteFailed, ex.Message),
                                Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>Lokaler, vorhandener Dateipfad einer Slot-Adresse oder null (Link, leer, fehlt).</summary>
        private static string? LocalMediaPathOf(string? url)
        {
            if (string.IsNullOrWhiteSpace(url)) return null;
            string bare = VideoViewerForm.SplitStops(url, out _);
            return IsHttpUrl(bare) ? null : TryParseLocalPath(bare);
        }
    }
}
