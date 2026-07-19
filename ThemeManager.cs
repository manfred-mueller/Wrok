using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Wrok
{
    /// <summary>
    /// Hält Fenstersymbol, Tray-Symbol und Titelleiste im Gleichklang mit dem
    /// Dark-/Light-Modus von Windows und reagiert auf Umschaltungen zur Laufzeit.
    ///
    /// Das Tray-Symbol wird erst nach dem Formular erzeugt, deshalb wird es über
    /// einen Zugriffsdelegaten geholt statt beim Erstellen übergeben.
    /// </summary>
    internal sealed class ThemeManager : IDisposable
    {
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE            = 20;
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;

        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        private readonly Form                _form;
        private readonly Func<NotifyIcon?>   _trayIcon;
        private bool                         _subscribed;

        public ThemeManager(Form form, Func<NotifyIcon?> trayIconAccessor)
        {
            _form     = form;
            _trayIcon = trayIconAccessor;
        }

        /// <summary>Ist der helle Windows-Modus abgeschaltet?</summary>
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

        /// <summary>Symbol passend zum aktuellen Modus (hell auf dunkel, dunkel auf hell).</summary>
        public static Icon CurrentIcon() =>
            IsDarkMode() ? Properties.Resources.wrok_white : Properties.Resources.wrok_black;

        /// <summary>Beginnt, auf Moduswechsel von Windows zu reagieren.</summary>
        public void StartListening()
        {
            if (_subscribed) return;
            SystemEvents.UserPreferenceChanged += OnUserPreferenceChanged;
            _subscribed = true;
        }

        public void StopListening()
        {
            if (!_subscribed) return;
            try { SystemEvents.UserPreferenceChanged -= OnUserPreferenceChanged; } catch { }
            _subscribed = false;
        }

        /// <summary>Symbole und Titelleiste neu setzen.</summary>
        public void Refresh()
        {
            try
            {
                bool dark = IsDarkMode();
                ApplyIcon();
                SetTitleBarDarkMode(dark);
            }
            catch (Exception ex) { Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [ThemeManager] Refresh fehlgeschlagen: {ex}"); }
        }

        public void ApplyIcon()
        {
            var sourceIcon = CurrentIcon();
            Icon newIcon;
            try { newIcon = (Icon)sourceIcon.Clone(); } catch { newIcon = sourceIcon; }

            var tray = _trayIcon();
            if (tray != null)
            {
                try
                {
                    // Kurzes Aus- und Einblenden: Ohne das uebernimmt der
                    // Infobereich das neue Symbol nicht zuverlaessig.
                    var old = tray.Icon;
                    tray.Visible = false;
                    tray.Icon    = newIcon;
                    tray.Visible = true;
                    if (old != null && !ReferenceEquals(old, sourceIcon)) try { old.Dispose(); } catch { }
                }
                catch { try { tray.Icon = newIcon; } catch { } }
            }

            try { _form.Icon = (Icon)newIcon.Clone(); } catch { _form.Icon = newIcon; }
        }

        private void SetTitleBarDarkMode(bool enabled)
        {
            try
            {
                int val = enabled ? 1 : 0;
                int hr  = DwmSetWindowAttribute(_form.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref val, Marshal.SizeOf<int>());
                if (hr != 0)
                    try { DwmSetWindowAttribute(_form.Handle, DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1, ref val, Marshal.SizeOf<int>()); } catch { }
            }
            catch { }
        }

        private void OnUserPreferenceChanged(object? sender, UserPreferenceChangedEventArgs e)
        {
            if (e.Category != UserPreferenceCategory.Color &&
                e.Category != UserPreferenceCategory.General &&
                e.Category != UserPreferenceCategory.VisualStyle) return;

            try { if (!_form.IsDisposed) _form.BeginInvoke((MethodInvoker)Refresh); } catch { }
        }

        public void Dispose() => StopListening();
    }
}
