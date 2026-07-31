using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;

namespace Wrok
{
    /// <summary>
    /// Verschlüsselt/entschlüsselt das Proxy-Passwort mit der Windows Data
    /// Protection API (DPAPI), Scope CurrentUser. So landet das Passwort nie
    /// im Klartext in der Settings-Datei (user.config) – nur derselbe
    /// Windows-Benutzer auf demselben Rechner kann es wieder entschlüsseln.
    ///
    /// Kein Schutz gegen denselben Benutzer/dieselbe Sitzung – DPAPI wehrt
    /// andere Konten oder das Kopieren der Datei auf einen fremden Rechner ab,
    /// nicht einen Angreifer mit vollem Zugriff auf das eigene Benutzerprofil.
    /// Für ein Desktop-Tool ohne eigenes Master-Passwort ist das die übliche,
    /// pragmatische Grenze.
    /// </summary>
    internal static class ProxyCredentialProtector
    {
        // Bindet die Verschlüsselung an Wrok, nicht an ein anderes Programm,
        // das ebenfalls DPAPI fürs selbe Windows-Konto nutzt.
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("Wrok.ProxyPassword.v1");

        /// <summary>Verschlüsselt <paramref name="plainText"/>; leer bleibt leer.</summary>
        public static string Protect(string? plainText)
        {
            if (string.IsNullOrEmpty(plainText)) return string.Empty;

            try
            {
                var bytes = Encoding.UTF8.GetBytes(plainText);
                var protectedBytes = ProtectedData.Protect(bytes, Entropy, DataProtectionScope.CurrentUser);
                return Convert.ToBase64String(protectedBytes);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [ProxyCredentialProtector] Protect fehlgeschlagen: {ex}");
                return string.Empty;
            }
        }

        /// <summary>
        /// Entschlüsselt einen zuvor mit <see cref="Protect"/> erzeugten Wert.
        /// Liefert leer statt einer Exception, wenn nicht entschlüsselbar
        /// (z. B. nach Windows-Neuinstallation oder Kontowechsel).
        /// </summary>
        public static string Unprotect(string? protectedText)
        {
            if (string.IsNullOrEmpty(protectedText)) return string.Empty;

            try
            {
                var protectedBytes = Convert.FromBase64String(protectedText);
                var bytes = ProtectedData.Unprotect(protectedBytes, Entropy, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(bytes);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [ProxyCredentialProtector] Unprotect fehlgeschlagen: {ex}");
                return string.Empty;
            }
        }
    }
}
