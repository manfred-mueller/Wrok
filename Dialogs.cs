using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Wrok
{
    /// <summary>Auswahl im Dialog zum Löschen der Browserdaten.</summary>
    internal enum ClearCacheChoice { Cancel, CacheOnly, All }

    /// <summary>
    /// Die modalen Dialoge der Anwendung, gebündelt an einer Stelle.
    ///
    /// Sie haben keinen Bezug zum Hauptformular außer dem Besitzerfenster und
    /// lagen dort nur historisch. Jede Methode ist für sich abgeschlossen:
    /// Eingaben rein, Ergebnis raus, kein gemeinsamer Zustand.
    /// </summary>
    internal static class Dialogs
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        // ------------------------------------------------------------------
        // Makro bearbeiten
        // ------------------------------------------------------------------

        /// <summary>
        /// Bearbeitet Name und Text eines Makros. Enter bestätigt, Umschalt+Enter
        /// erzeugt im Textfeld eine neue Zeile.
        /// </summary>
        public static bool ShowEditMacro(Form owner, string title, ref string name, ref string text)
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
            var lblCount  = new Label { AutoSize = true, Font = new Font("Segoe UI", 8F), ForeColor = SystemColors.GrayText, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Text = string.Format(Properties.Resources.CharacterCount, tb.Text.Length) };
            var btnOk     = new Button { Text = Properties.Resources.OK,     DialogResult = DialogResult.OK,     Size = new Size(80, 26), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            var btnCancel = new Button { Text = Properties.Resources.Cancel, DialogResult = DialogResult.Cancel, Size = new Size(80, 26), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };

            var varTip = new ToolTip { AutoPopDelay = 15000, InitialDelay = 400, ReshowDelay = 200 };
            varTip.SetToolTip(tb, Properties.Resources.MacroVariablesHint);

            tb.TextChanged += (s, e) => lblCount.Text = string.Format(Properties.Resources.CharacterCount, tb.Text.Length);

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

            if (dlg.ShowDialog(owner) == DialogResult.OK) { name = tbName.Text; text = tb.Text; return true; }
            return false;
        }

        // ------------------------------------------------------------------
        // Zielprofil für einen Import wählen
        // ------------------------------------------------------------------

        /// <summary>
        /// Lässt das Zielprofil für einen Import wählen. Vorausgewählt ist das
        /// aktive Profil, damit der häufigste Fall ein Enter-Druck bleibt.
        /// </summary>
        public static bool ShowProfileChooser(Form owner, IReadOnlyList<string> profileNames,
                                              int activeIndex, out int index)
        {
            index = activeIndex;

            var items = profileNames.Select((n, i) => string.IsNullOrWhiteSpace(n)
                    ? string.Format(Properties.Resources.MacroProfileDefault, i + 1)
                    : $"{i + 1} – {n}")
                .ToArray();

            using var dlg = new Form
            {
                Text            = Properties.Resources.ImportTargetTitle,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition   = FormStartPosition.CenterParent,
                MinimizeBox     = false,
                MaximizeBox     = false,
                ShowInTaskbar   = false,
                ClientSize      = new Size(400, 150)
            };

            var lbl = new Label
            {
                Text     = Properties.Resources.ImportTargetPrompt,
                AutoSize = false,
                Left     = 16,
                Top      = 14,
                Width    = dlg.ClientSize.Width - 32,
                Height   = 40,
                Font     = new Font("Segoe UI", 9.5F)
            };
            var combo = new ComboBox
            {
                Left          = 16,
                Top           = 62,
                Width         = dlg.ClientSize.Width - 32,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font          = new Font("Segoe UI", 10F)
            };
            combo.Items.AddRange(items);
            if (index >= 0 && index < items.Length) combo.SelectedIndex = index;

            var btnOk     = new Button { Text = Properties.Resources.OK,     DialogResult = DialogResult.OK,     Size = new Size(80, 28), Top = 106 };
            var btnCancel = new Button { Text = Properties.Resources.Cancel, DialogResult = DialogResult.Cancel, Size = new Size(80, 28), Top = 106 };
            btnCancel.Left = dlg.ClientSize.Width - btnCancel.Width - 16;
            btnOk.Left     = btnCancel.Left - btnOk.Width - 8;

            dlg.Controls.AddRange(new Control[] { lbl, combo, btnOk, btnCancel });
            dlg.AcceptButton = btnOk;
            dlg.CancelButton = btnCancel;

            if (dlg.ShowDialog(owner) != DialogResult.OK) return false;
            index = combo.SelectedIndex;
            return index >= 0;
        }

        // ------------------------------------------------------------------
        // Browserdaten löschen: Auswahl
        // ------------------------------------------------------------------

        /// <summary>
        /// Zeigt einen kleinen Dialog mit drei Schaltflächen
        /// (Nur Cache / Alles löschen / Abbrechen) und gibt die Auswahl zurück.
        ///
        /// <paramref name="includeMacros"/> meldet zurück, ob zusätzlich die
        /// Makroprofile geleert werden sollen. Bewusst abwählbar und standardmäßig
        /// aus: Makros sind eigene Arbeit des Nutzers, kein Browserballast – wer
        /// aufräumt, will sie in aller Regel behalten.
        /// </summary>
        public static ClearCacheChoice ShowClearCacheChoice(Form owner, out bool includeMacros)
        {
            includeMacros = false;

            using var dlg = new Form
            {
                Text            = Properties.Resources.ClearBrowsingData,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition   = FormStartPosition.CenterParent,
                MinimizeBox     = false,
                MaximizeBox     = false,
                ShowInTaskbar   = false,
                ClientSize      = new Size(440, 156)
            };

            var lbl = new Label
            {
                Text     = Properties.Resources.ChooseWhatToBeDeleted,
                AutoSize = false,
                Left     = 16,
                Top      = 16,
                Width    = dlg.ClientSize.Width - 32,
                Height   = 56,
                Font     = new Font("Segoe UI", 9.5F)
            };

            var chkMacros = new CheckBox
            {
                Text     = Properties.Resources.ClearMacrosToo,
                AutoSize = true,
                Left     = 16,
                Top      = 76,
                Checked  = false,
                Font     = new Font("Segoe UI", 9F)
            };

            var btnCacheOnly = new Button { Text = Properties.Resources.ClearCacheOnly, DialogResult = DialogResult.Yes,    Size = new Size(120, 30), Top = 110 };
            var btnAll       = new Button { Text = Properties.Resources.DeleteAll,      DialogResult = DialogResult.No,     Size = new Size(120, 30), Top = 110 };
            var btnCancel    = new Button { Text = Properties.Resources.Cancel,         DialogResult = DialogResult.Cancel, Size = new Size(90,  30), Top = 110 };

            btnCacheOnly.Left = 16;
            btnAll.Left       = btnCacheOnly.Right + 8;
            btnCancel.Left    = dlg.ClientSize.Width - btnCancel.Width - 16;

            // Die Makro-Option gehoert allein zu „Alles loeschen“. Statt sie still zu
            // ignorieren, wird sie sichtbar deaktiviert, sobald der Mauszeiger auf
            // „Nur Cache“ steht – so entsteht gar nicht erst der falsche Eindruck.
            btnCacheOnly.MouseEnter += (s, e) => chkMacros.Enabled = false;
            btnCacheOnly.MouseLeave += (s, e) => chkMacros.Enabled = true;

            dlg.Controls.AddRange(new Control[] { lbl, chkMacros, btnCacheOnly, btnAll, btnCancel });
            dlg.AcceptButton = btnCacheOnly;
            dlg.CancelButton = btnCancel;

            var result = dlg.ShowDialog(owner);
            includeMacros = result == DialogResult.No && chkMacros.Checked;

            return result switch
            {
                DialogResult.Yes => ClearCacheChoice.CacheOnly,
                DialogResult.No  => ClearCacheChoice.All,
                _                => ClearCacheChoice.Cancel
            };
        }

        // ------------------------------------------------------------------
        // Einzeilige Eingabe
        // ------------------------------------------------------------------

        /// <summary>
        /// Schlichter einzeiliger Eingabedialog (Profilnamen o. Ä.).
        /// Enter bestätigt, Esc bricht ab.
        /// </summary>
        public static bool ShowSingleLine(Form owner, string title, string prompt,
                                          string initial, out string value)
        {
            value = string.Empty;

            using var dlg = new Form
            {
                Text            = title,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition   = FormStartPosition.CenterParent,
                MinimizeBox     = false,
                MaximizeBox     = false,
                ShowInTaskbar   = false,
                ClientSize      = new Size(400, 120)
            };

            var lbl = new Label
            {
                Text     = prompt,
                AutoSize = true,
                Left     = 16,
                Top      = 16,
                Font     = new Font("Segoe UI", 9.5F)
            };
            var tb = new TextBox
            {
                Left   = 16,
                Top    = 44,
                Width  = dlg.ClientSize.Width - 32,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font   = new Font("Segoe UI", 10F),
                Text   = initial ?? string.Empty
            };
            var btnOk     = new Button { Text = Properties.Resources.OK,     DialogResult = DialogResult.OK,     Size = new Size(80, 28), Top = 80 };
            var btnCancel = new Button { Text = Properties.Resources.Cancel, DialogResult = DialogResult.Cancel, Size = new Size(80, 28), Top = 80 };
            btnCancel.Left = dlg.ClientSize.Width - btnCancel.Width - 16;
            btnOk.Left     = btnCancel.Left - btnOk.Width - 8;

            dlg.Controls.AddRange(new Control[] { lbl, tb, btnOk, btnCancel });
            dlg.AcceptButton = btnOk;
            dlg.CancelButton = btnCancel;
            dlg.Shown += (s, e) => { tb.Focus(); tb.SelectAll(); };

            if (dlg.ShowDialog(owner) != DialogResult.OK) return false;
            value = tb.Text;
            return true;
        }

        // ------------------------------------------------------------------
        // Mehrzeilige Eingabe für {input}
        // ------------------------------------------------------------------

        /// <summary>
        /// Mehrzeiliger Eingabedialog für {input}. Strg+Enter sendet, Esc bricht ab.
        /// Bewusst mehrzeilig: Im Rollenspiel sind die Einwürfe oft längere
        /// Erzählpassagen, die in einem einzeiligen Feld nicht überblickbar wären.
        /// </summary>
        public static bool ShowMultiLineInput(Form owner, string prompt, out string value)
        {
            value = string.Empty;

            // Wird das Makro per globalem Hotkey ausgelöst, ist Wrok evtl. minimiert
            // oder im Hintergrund. Dann muss der Dialog selbst nach vorn kommen –
            // sonst erscheint er hinter dem aktiven Fenster und bekommt keinen Fokus.
            bool ownerUsable = owner is { Visible: true } && owner.WindowState != FormWindowState.Minimized;

            using var dlg = new Form
            {
                Text            = Properties.Resources.MacroInputTitle,
                FormBorderStyle = FormBorderStyle.SizableToolWindow,
                StartPosition   = ownerUsable ? FormStartPosition.CenterParent
                                              : FormStartPosition.CenterScreen,
                MinimizeBox     = false,
                MaximizeBox     = false,
                ShowInTaskbar   = false,
                TopMost         = true,
                MinimumSize     = new Size(360, 220),
                ClientSize      = new Size(560, 300)
            };

            var lbl = new Label
            {
                Text     = prompt,
                AutoSize = false,
                Left     = 16,
                Top      = 12,
                Width    = dlg.ClientSize.Width - 32,
                Height   = 20,
                Anchor   = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font     = new Font("Segoe UI", 9.5F)
            };
            var tb = new TextBox
            {
                Multiline     = true,
                AcceptsReturn = true,     // Enter erzeugt einen Zeilenumbruch
                WordWrap      = true,
                ScrollBars    = ScrollBars.Vertical,
                Left          = 16,
                Top           = 36,
                Width         = dlg.ClientSize.Width - 32,
                Height        = dlg.ClientSize.Height - 36 - 46,
                Anchor        = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                Font          = new Font("Segoe UI", 10F)
            };
            var lblHint = new Label
            {
                Text      = Properties.Resources.InputDialogHint,
                AutoSize  = true,
                Anchor    = AnchorStyles.Bottom | AnchorStyles.Left,
                ForeColor = SystemColors.GrayText,
                Font      = new Font("Segoe UI", 8F)
            };
            var btnOk     = new Button { Text = Properties.Resources.OK,     DialogResult = DialogResult.OK,     Size = new Size(80, 28), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            var btnCancel = new Button { Text = Properties.Resources.Cancel, DialogResult = DialogResult.Cancel, Size = new Size(80, 28), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };

            void LayoutBottom()
            {
                int y = dlg.ClientSize.Height - btnOk.Height - 10;
                tb.Height          = y - tb.Top - 8;
                btnCancel.Location = new Point(dlg.ClientSize.Width - btnCancel.Width - 16, y);
                btnOk.Location     = new Point(btnCancel.Left - btnOk.Width - 8, y);
                lblHint.Location   = new Point(16, y + (btnOk.Height - lblHint.Height) / 2);
            }

            dlg.Controls.AddRange(new Control[] { lbl, tb, lblHint, btnOk, btnCancel });
            // KEIN AcceptButton: Enter soll eine neue Zeile erzeugen, nicht senden.
            dlg.CancelButton = btnCancel;
            dlg.Resize += (s, e) => LayoutBottom();

            tb.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter && e.Control)
                {
                    e.SuppressKeyPress = true;
                    dlg.DialogResult = DialogResult.OK;
                    dlg.Close();
                }
            };

            dlg.Shown += (s, e) =>
            {
                LayoutBottom();
                try
                {
                    dlg.Activate();
                    SetForegroundWindow(dlg.Handle);
                }
                catch (Exception ex) { Trace.WriteLine($"ShowMultiLineInput: Fokus fehlgeschlagen: {ex}"); }
                tb.Focus();
            };

            // Ohne sichtbaren Owner ohne Owner-Fenster anzeigen, sonst hängt der
            // modale Dialog an einem minimierten/unsichtbaren Fenster.
            var result = ownerUsable ? dlg.ShowDialog(owner) : dlg.ShowDialog();
            if (result != DialogResult.OK) return false;
            value = tb.Text;
            return true;
        }
    }
}
