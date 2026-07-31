using System.Diagnostics;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Wrok
{
    /// <summary>
    /// Ein gespeichertes Grok-Konto. Bewusst NUR Bezeichnung + E-Mail – kein
    /// Passwort. Das Passwort bleibt Sache des Chromium-eigenen Passwort-
    /// Managers (siehe WebViewManager.ConfigureCore, IsPasswordAutosaveEnabled)
    /// bzw. der manuellen Eingabe.
    /// </summary>
    internal record GrokAccount(string Label, string Email)
    {
        public string DisplayName => string.IsNullOrWhiteSpace(Label) ? Email : $"{Label} ({Email})";
    }

    /// <summary>
    /// Verwaltet die Liste gespeicherter Grok-Konten. Persistiert als JSON in
    /// Settings.settings – analog zu MacroManager, keine eigene Storage-
    /// Infrastruktur nötig.
    /// </summary>
    internal static class GrokAccountManager
    {
        private static readonly Regex EmailPattern =
            new(@"^[^@\s]+@[^@\s]+\.[^@\s]+$", RegexOptions.Compiled);

        /// <summary>Simple, ausreichende Prüfung – kein Anspruch auf RFC-Vollständigkeit.</summary>
        public static bool IsValidEmail(string? email) =>
            !string.IsNullOrWhiteSpace(email) && EmailPattern.IsMatch(email.Trim());

        public static List<GrokAccount> LoadAccounts()
        {
            try
            {
                var json = Properties.Settings.Default.GrokAccountsJson;
                if (string.IsNullOrWhiteSpace(json)) return new List<GrokAccount>();

                var list = JsonSerializer.Deserialize<List<GrokAccount>>(json);
                return list ?? new List<GrokAccount>();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [GrokAccountManager] LoadAccounts fehlgeschlagen: {ex}");
                return new List<GrokAccount>();
            }
        }

        private static void SaveAccounts(List<GrokAccount> accounts)
        {
            try
            {
                Properties.Settings.Default.GrokAccountsJson = JsonSerializer.Serialize(accounts);
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [GrokAccountManager] SaveAccounts fehlgeschlagen: {ex}");
            }
        }

        /// <summary>
        /// Fügt ein Konto hinzu. Wirft <see cref="InvalidOperationException"/> bei
        /// ungültiger oder bereits vorhandener (case-insensitive) E-Mail-Adresse –
        /// der Aufrufer zeigt die Meldung direkt an (siehe MainForm.AddGrokAccountAndSave).
        /// </summary>
        public static void AddAccount(string label, string email)
        {
            email = (email ?? string.Empty).Trim();

            if (!IsValidEmail(email))
                throw new InvalidOperationException(Properties.Resources.GrokAccountInvalidEmail);

            var accounts = LoadAccounts();
            if (accounts.Any(a => string.Equals(a.Email, email, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException(Properties.Resources.GrokAccountDuplicate);

            string trimmedLabel = (label ?? string.Empty).Trim();
            accounts.Add(new GrokAccount(string.IsNullOrWhiteSpace(trimmedLabel) ? email : trimmedLabel, email));
            SaveAccounts(accounts);
        }

        public static void RemoveAccount(string email)
        {
            var accounts = LoadAccounts();
            accounts.RemoveAll(a => string.Equals(a.Email, email, StringComparison.OrdinalIgnoreCase));
            SaveAccounts(accounts);
        }
    }
}
