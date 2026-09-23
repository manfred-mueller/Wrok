using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Wrok
{
    public partial class MainForm
    {
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

            var macros = _macroManager.GetMacros();
            for (int i = 1; i <= MacroManager.MacroCount; i++)
            {
                var entry = macros[i - 1];

                var macroItem = new ToolStripMenuItem(entry.DisplayName(i)) { Tag = i - 1 };

                // Direkt per F-Taste erreichbar (F1..F{MacroCount}), aber nur wenn
                // dieses Profil das aktive ist und der F-Tasten-Override in den
                // Einstellungen aktiviert wurde (siehe HandleFKeyOverride).
                string hotkeyDisplay = string.Format(Properties.Resources.FKeyMacroHotkeyLabel, i);
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
        // F-Tasten-Override (Makros) - Kernlogik
        // ------------------------------------------------------------------

        /// <summary>
        /// Zentrale Entscheidung für eine gedrückte F-Taste (F1-F12), egal ob sie
        /// über ProcessCmdKey (WinForms-Control fokussiert) oder über
        /// WebViewManager/CoreWebView2.AcceleratorKeyPressed (WebView2-Inhalt
        /// fokussiert) hereinkommt. Gibt true zurück, wenn die Taste verarbeitet
        /// wurde (Aufrufer soll sie dann schlucken).
        ///
        /// Bewusst nur wirksam, wenn Wrok den Tastaturfokus hat - kein globaler
        /// Hotkey, kein Eingriff in andere Anwendungen. F11/F12 bleiben in jedem
        /// Fall unangetastet (natives Vollbild/DevTools).
        /// </summary>
        internal bool HandleFKeyOverride(Keys keyCode, bool control, bool shift, bool alt)
        {
            if (alt) return false;   // Alt+F4 & Co. nie anfassen
            if (!Properties.Settings.Default.FKeyMacroOverrideEnabled) return false;
            if (keyCode == Keys.F11 || keyCode == Keys.F12) return false;

            int fIndex = keyCode switch
            {
                Keys.F1 => 0, Keys.F2 => 1, Keys.F3 => 2, Keys.F4 => 3, Keys.F5 => 4,
                Keys.F6 => 5, Keys.F7 => 6, Keys.F8 => 7, Keys.F9 => 8, Keys.F10 => 9,
                _ => -1
            };
            if (fIndex < 0) return false;

            if (control && shift)
            {
                if (fIndex >= MacroManager.ProfileCount) return false;
                SwitchProfileViaHotkey(fIndex);
                return true;
            }

            if (control)
            {
                // Einzige verbliebene "Standardfunktion": Strg+F5 = Neu laden.
                // WebView2 läuft ohne sichtbare Browser-Chrome, daher haben die
                // übrigen F-Tasten hier ohnehin keine relevante Default-Aktion.
                if (keyCode == Keys.F5)
                    try { _webView?.CoreWebView2?.Reload(); } catch { }
                return true;
            }

            _ = PerformMacroAsync(_macroManager.ActiveProfile, fIndex);
            return true;
        }

        /// <summary>
        /// Wechselt das aktive Profil per Strg+Umschalt+F-Taste und bestätigt das
        /// per kurzer Balloon-Tip-Meldung, da sonst keine Rückmeldung sichtbar
        /// wäre, ohne das Tray-Menü zu öffnen.
        /// </summary>
        private void SwitchProfileViaHotkey(int index)
        {
            try
            {
                if (index == _macroManager.ActiveProfile) return;

                _macroManager.SwitchProfile(index);
                RefreshMacrosMenu();

                trayIcon?.ShowBalloonTip(1500,
                    Properties.Resources.ProfileSwitchedBalloonTitle,
                    string.Format(Properties.Resources.ProfileSwitchedBalloonText, ProfileLabel(index)),
                    ToolTipIcon.Info);
            }
            catch (Exception ex)
            {
                Log(ex, "SwitchProfileViaHotkey fehlgeschlagen");
            }
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            var s = Properties.Settings.Default;
            if (s.WindowWidth <= 0 || s.WindowHeight <= 0)
                TryCenterToScreen();
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
        // Makro-Hotkey ausführen
        // ------------------------------------------------------------------

        private async Task PerformMacroAsync(int profileIndex, int macroIndex)
        {
            try
            {
                string? raw = _macroManager.GetRawText(profileIndex, macroIndex);
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
    }
}
