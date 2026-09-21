using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Wrok
{
    /// <summary>
    /// Kurzes, modales Auswahl-Popup: zeigt die Makrotitel eines Profils an,
    /// eine folgende Ziffer (1-9, 0 für Makro 10) wählt aus, Escape oder
    /// Fokusverlust bricht ab. Kein globaler Hotkey nötig - die Zifferntasten
    /// werden nur erfasst, solange dieses Fenster den Fokus hat, blockieren
    /// also nirgendwo sonst die normale Zifferneingabe.
    /// </summary>
    internal sealed class MacroPickerForm : Form
    {
        // Reihenfolge wie im restlichen Programm etabliert: Makro 1..9, dann
        // Makro 10 auf der 0-Taste (siehe z. B. RefreshMacrosMenu/ctrlDigit).
        private static readonly Keys[] DigitKeys =
        {
            Keys.D1, Keys.D2, Keys.D3, Keys.D4, Keys.D5,
            Keys.D6, Keys.D7, Keys.D8, Keys.D9, Keys.D0
        };

        private int  _result = -1;
        private bool _closing;

        // Windows verweigert SetForegroundWindow oft, wenn der aufrufende Prozess
        // gerade nicht im Vordergrund ist - genau der Fall hier, da das Popup aus
        // einem globalen Hotkey heraus entsteht, während irgendeine andere
        // Anwendung den Fokus hat. Ohne den Trick erscheint das Fenster zwar
        // (TopMost), aber mit ausgegrauter, inaktiver Titelleiste, und Klicks
        // außerhalb lösen kein Deactivate aus, weil es nie wirklich aktiv war.
        // Standard-Workaround: ein kurzer simulierter Alt-Tastendruck setzt
        // intern das Flag, das SetForegroundWindow danach tatsächlich wirken lässt.
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
        private const byte VK_MENU         = 0x12;
        private const uint KEYEVENTF_KEYUP = 0x0002;

        private MacroPickerForm(string profileLabel, IReadOnlyList<MacroEntry> macros)
        {
            Text            = profileLabel;
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            StartPosition   = FormStartPosition.Manual;
            ShowInTaskbar   = false;
            TopMost         = true;
            MaximizeBox     = false;
            MinimizeBox     = false;
            AutoScaleMode   = AutoScaleMode.Font;
            KeyPreview      = true;

            // Bewusst OHNE Dock/FlowLayoutPanel-Autolayout: feste Positionierung
            // jeder Zeile, damit nichts von Docking-Reihenfolge oder AutoSize
            // abhängt (das hatte zuvor Zeile 1 unter dem Titel verschwinden lassen).
            // Kein zusätzliches Titel-Label mehr - der Profilname steht bereits in
            // der Fensterleiste (Text-Eigenschaft oben), eine zweite fette Zeile
            // darüber war redundant.
            const int rowHeight  = 22;
            const int formWidth  = 280;
            const int bottomPad  = 10;

            int count = Math.Min(macros.Count, DigitKeys.Length);
            for (int i = 0; i < count; i++)
            {
                string digit = i < 9 ? (i + 1).ToString() : "0";
                string title = macros[i].DisplayName(i + 1);
                string text  = string.IsNullOrWhiteSpace(macros[i].Text)
                    ? string.Format(Properties.Resources.MacroPickerEmptyRow, digit)
                    : string.Format(Properties.Resources.MacroPickerRow, digit, title);

                var row = new Label
                {
                    Text      = text,
                    AutoSize  = false,
                    Location  = new Point(0, i * rowHeight),
                    Size      = new Size(formWidth, rowHeight),
                    Padding   = new Padding(8, 0, 0, 0),
                    TextAlign = ContentAlignment.MiddleLeft,
                    Font      = new Font("Segoe UI", 9.5F),
                    Cursor    = Cursors.Hand,
                    Tag       = i
                };
                // Maus-Klick als Alternative zur Zifferntaste - selbes Ergebnis.
                row.Click += (s, e) => CloseWith((int)((Label)s!).Tag);
                Controls.Add(row);
            }

            ClientSize = new Size(formWidth, count * rowHeight + bottomPad);

            Deactivate += (s, e) => CloseWith(-1);
            Shown      += (s, e) => ForceForeground();
        }

        private void ForceForeground()
        {
            try
            {
                keybd_event(VK_MENU, 0, 0, UIntPtr.Zero);
                keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                SetForegroundWindow(Handle);
            }
            catch { }
        }

        /// <summary>
        /// Setzt das Ergebnis und schließt - mit Schutz gegen mehrfaches Auslösen
        /// (z. B. wenn Close() selbst ein Deactivate auslöst).
        /// </summary>
        private void CloseWith(int result)
        {
            if (_closing) return;
            _closing = true;
            _result  = result;
            Close();
        }

        /// <summary>
        /// Zeigt das Popup an und blockiert, bis eine Ziffer/Escape gedrückt
        /// wird oder der Fokus verloren geht. Gibt den gewählten 0-basierten
        /// Makroindex zurück, oder -1 bei Abbruch.
        /// </summary>
        /// <param name="onShown">
        /// Wird mit der fertig positionierten, aber noch nicht sichtbaren Form
        /// aufgerufen, BEVOR ShowDialog blockiert - so kann der Aufrufer sich die
        /// Referenz merken (z. B. um sie beim Ausblenden des Hauptfensters von
        /// außen zu schließen), ohne auf das Ende von PickOption warten zu müssen.
        /// </param>
        public static int PickOption(string profileLabel, IReadOnlyList<MacroEntry> macros,
            Action<MacroPickerForm>? onShown = null)
        {
            try
            {
                using var form = new MacroPickerForm(profileLabel, macros);

                var area = Screen.FromPoint(Cursor.Position).WorkingArea;
                int x = Math.Clamp(Cursor.Position.X + 12, area.Left, area.Right  - form.Width);
                int y = Math.Clamp(Cursor.Position.Y + 12, area.Top,  area.Bottom - form.Height);
                form.Location = new Point(x, y);

                onShown?.Invoke(form);
                form.ShowDialog();
                return form._result;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MacroPickerForm] PickOption fehlgeschlagen: {ex}");
                return -1;
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                CloseWith(-1);
                return true;
            }

            for (int i = 0; i < DigitKeys.Length; i++)
            {
                if (keyData == DigitKeys[i])
                {
                    CloseWith(i);
                    return true;
                }
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}
