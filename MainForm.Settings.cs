using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Wrok
{
    public partial class MainForm
    {
        // ------------------------------------------------------------------
        // Proxy-Einstellungen
        // ------------------------------------------------------------------

        /// <summary>
        /// Zeigt den Proxy-Dialog und speichert das Ergebnis. Chromium liest
        /// den Proxy nur beim Erzeugen der CoreWebView2Environment (siehe
        /// WebViewManager.InitializeAsync) – eine Änderung wirkt daher erst
        /// nach einem Neustart, den diese Methode bei Bedarf direkt anbietet.
        /// </summary>
        private void ShowProxyDialog()
        {
            bool   enabled      = Properties.Settings.Default.ProxyEnabled;
            string server       = Properties.Settings.Default.ProxyServer ?? string.Empty;
            bool   requiresAuth = Properties.Settings.Default.ProxyAuthEnabled;
            string username     = Properties.Settings.Default.ProxyUsername ?? string.Empty;
            // Nur zum Anzeigen/Bearbeiten im Dialog entschlüsselt, danach sofort
            // wieder verworfen (liegt nur lokal in dieser Methode im Klartext).
            string password     = ProxyCredentialProtector.Unprotect(Properties.Settings.Default.ProxyPasswordProtected);

            if (!Dialogs.ShowProxySettings(this, ref enabled, ref server, ref requiresAuth, ref username, ref password))
                return;

            bool changed = enabled      != Properties.Settings.Default.ProxyEnabled
                        || server       != (Properties.Settings.Default.ProxyServer   ?? string.Empty)
                        || requiresAuth != Properties.Settings.Default.ProxyAuthEnabled
                        || username     != (Properties.Settings.Default.ProxyUsername ?? string.Empty)
                        || password     != ProxyCredentialProtector.Unprotect(Properties.Settings.Default.ProxyPasswordProtected);
            if (!changed) return;

            Properties.Settings.Default.ProxyEnabled          = enabled;
            Properties.Settings.Default.ProxyServer           = server;
            Properties.Settings.Default.ProxyAuthEnabled      = requiresAuth;
            // Ohne aktive Anmeldung keine Reste stehen lassen, statt nur die
            // Checkbox umzuschalten und ein altes Passwort verschlüsselt liegen
            // zu lassen.
            Properties.Settings.Default.ProxyUsername         = requiresAuth ? username : string.Empty;
            Properties.Settings.Default.ProxyPasswordProtected = requiresAuth
                ? ProxyCredentialProtector.Protect(password)
                : string.Empty;
            Properties.Settings.Default.Save();

            var result = MessageBox.Show(this, Properties.Resources.ProxyRestartRequiredMessage,
                Properties.Resources.ProxyRestartRequiredTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Information);

            if (result == DialogResult.Yes)
                Program.RestartApplication();
        }

        // ------------------------------------------------------------------
        // Cache leeren
        // ------------------------------------------------------------------

        private async Task ClearCacheAsync()
        {
            try
            {
                if (_webView?.CoreWebView2 == null)
                { MessageBox.Show(Properties.Resources.WebView2IsNotInitializedYet, Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var choice = Dialogs.ShowClearCacheChoice(this, out bool includeMacros);
                if (choice == ClearCacheChoice.Cancel) return;

                if (choice == ClearCacheChoice.CacheOnly)
                {
                    // Nur Cache (Bilder, Skripte, …) – Cookies und Login bleiben erhalten.
                    await _webView.CoreWebView2.Profile.ClearBrowsingDataAsync(
                        CoreWebView2BrowsingDataKinds.DiskCache |
                        CoreWebView2BrowsingDataKinds.CacheStorage);
                    MessageBox.Show(
                        Properties.Resources.PicturesScriptsAndOtherDataDeletedNYouAreStillLoggedIn,
                        Properties.Resources.CacheCleared, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // Alles: erst die WebView2-Seite (Cookies, Login, lokaler Speicher),
                    // dann die Spuren, die Wrok selbst ausserhalb des Browserprofils
                    // hinterlaesst. Ohne den zweiten Schritt blieben das zuletzt
                    // geoeffnete Bild und dessen URL zurueck.
                    await _webView.CoreWebView2.Profile.ClearBrowsingDataAsync();
                    ClearWrokPrivateData(includeMacros);

                    MessageBox.Show(
                        includeMacros
                            ? Properties.Resources.AllDataDeletedInclMacrosMsg
                            : Properties.Resources.AllDataDeletedMsg,
                        Properties.Resources.AllDataDeleted, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MainForm] ClearCacheAsync fehlgeschlagen: {ex}");
                MessageBox.Show(string.Format(Properties.Resources.ErrorWhileDeletingCache + "\n{0}", ex.Message), Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Entfernt die Daten, die Wrok ausserhalb des WebView2-Profils ablegt.
        ///
        /// <see cref="CoreWebView2Profile.ClearBrowsingDataAsync()"/> raeumt nur den
        /// Browser auf. Das zuletzt geoeffnete Bild liegt aber als PNG unter
        /// %LOCALAPPDATA%\Wrok\images, und seine Quell-URL steht im Klartext in den
        /// Einstellungen – beides ueberlebte das Loeschen bisher unbemerkt.
        /// </summary>
        private void ClearWrokPrivateData(bool includeMacros)
        {
            // Ein offenes Medienfenster wuerde sonst genau das weiter anzeigen,
            // was gerade geloescht wird.
            CloseCurrentMedia();

            try
            {
                var path = Properties.Settings.Default.LastImagePath;
                if (!string.IsNullOrWhiteSpace(path))
                {
                    // Auch die temporaere Datei aus SaveAsLastImage, die bei einem
                    // Abbruch liegen bleiben kann.
                    foreach (var candidate in new[] { path, path + ".tmp" })
                        if (File.Exists(candidate)) File.Delete(candidate);
                }
            }
            catch (Exception ex) { Log(ex, "Letztes Bild loeschen fehlgeschlagen"); }

            try
            {
                Properties.Settings.Default.LastImagePath = string.Empty;
                Properties.Settings.Default.LastImageUrl  = string.Empty;
                Properties.Settings.Default.Save();
            }
            catch (Exception ex) { Log(ex, "Bildverweise zuruecksetzen fehlgeschlagen"); }

            if (!includeMacros) return;

            try
            {
                for (int i = 0; i < MacroManager.ProfileCount; i++)
                    _macroManager.ResetProfile(i);

                RefreshMacrosMenu();
            }
            catch (Exception ex) { Log(ex, "Makroprofile loeschen fehlgeschlagen"); }
        }

    }
}
