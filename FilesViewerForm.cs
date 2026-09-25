using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Wrok
{
    /// <summary>
    /// Zeigt Groks eigene Dateiverwaltung (grok.com/files – Uploads und von Grok/
    /// Imagine erzeugte Dateien, inkl. Löschen) in einem eigenständigen Fenster.
    ///
    /// Wie VideoViewerForm ein zweites WebView2 mit derselben CoreWebView2Environment
    /// wie das Hauptfenster, damit dieselbe Login-Session greift. Im Unterschied zum
    /// Video-Player ist grok.com/files eine echte, eigenständige Seite - kein
    /// synthetisches Dokument nötig, daher genügt ein einfaches Navigate().
    ///
    /// Tastenkürzel:
    ///   Esc      schließen
    ///   Strg+P   anheften (immer im Vordergrund)
    /// </summary>
    internal sealed class FilesViewerForm : Form
    {
        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

        private const string FilesUrl = "https://grok.com/files";

        private readonly WebView2 _web;
        private readonly CoreWebView2Environment? _env;

        public FilesViewerForm(CoreWebView2Environment? env)
        {
            _env = env;

            Text            = Properties.Resources.FilesViewerTitle;
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition   = FormStartPosition.CenterScreen;
            BackColor       = Color.FromArgb(28, 28, 28);
            ShowInTaskbar   = true;
            KeyPreview      = true;
            MinimumSize     = new Size(480, 360);
            ClientSize      = new Size(1000, 720);

            TryApplyDarkTitleBar();

            _web = new WebView2 { Dock = DockStyle.Fill };
            Controls.Add(_web);

            KeyDown += OnKeyDown;
            _ = InitializeAsync();
        }

        private async Task InitializeAsync()
        {
            try
            {
                // Dieselbe Umgebung wie das Hauptfenster → gleiche Session/Cookies,
                // Grok muss hier nicht erneut angemeldet werden.
                await _web.EnsureCoreWebView2Async(_env);

                var core = _web.CoreWebView2;
                if (core == null) return;

                core.Settings.AreDefaultContextMenusEnabled = true;
                core.Settings.IsStatusBarEnabled            = false;

                core.Navigate(FilesUrl);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [FilesViewerForm] Initialisierung fehlgeschlagen: {ex}");
            }
        }

        private void OnKeyDown(object? sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Escape:
                    Close();
                    break;

                case Keys.P when e.Control:
                    TopMost = !TopMost;
                    Text = Properties.Resources.FilesViewerTitle + (TopMost ? "   📌" : string.Empty);
                    break;
            }
        }

        private void TryApplyDarkTitleBar()
        {
            try
            {
                int val = 1;
                if (DwmSetWindowAttribute(Handle, 20, ref val, sizeof(int)) != 0)
                    DwmSetWindowAttribute(Handle, 19, ref val, sizeof(int));
            }
            catch { }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                try { _web.Dispose(); } catch { }
            }
            base.Dispose(disposing);
        }
    }
}
