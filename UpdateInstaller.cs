using System.Diagnostics;
using System.Net.Http;
using System.Runtime.InteropServices;

namespace Wrok
{
    /// <summary>
    /// Lädt das Setup-Release-Asset herunter, das <see cref="UpdateChecker"/> in
    /// <see cref="UpdateInfo.InstallerUrl"/> ermittelt hat, und prüft dessen
    /// Authenticode-Signatur, bevor <see cref="Program.RestartApplicationForUpdate"/>
    /// es ausführt.
    ///
    /// Bewusst eine eigene Klasse statt Teil von UpdateChecker: Der Kommentar dort
    /// betont ausdrücklich, dass die Prüfung "nichts herunterlädt und nichts
    /// ausführt" - das gilt für den reinen Versions-Check weiterhin. Diese Klasse
    /// ist der Teil, der tatsächlich eine Datei aus dem Netz holt und später (mit
    /// Administratorrechten) startet, und braucht deshalb eine eigene, robustere
    /// Absicherung als ein einzelner lesender API-Aufruf.
    /// </summary>
    internal static class UpdateInstaller
    {
        // Grosszuegiges Timeout: ein Setup-Download kann je nach Verbindung deutlich
        // laenger dauern als die 10 Sekunden fuer den reinen Versions-Check.
        private static readonly HttpClient _http = new() { Timeout = TimeSpan.FromMinutes(5) };

        /// <summary>
        /// Lädt das Installer-Asset aus <paramref name="info"/> nach %TEMP% herunter
        /// und verifiziert anschliessend dessen Authenticode-Signatur. Wirft bei jedem
        /// Fehlschlag (kein Asset-Link im Release, Download-Fehler, fehlende oder
        /// ungültige Signatur) - der Aufrufer zeigt in dem Fall eine Fehlermeldung.
        /// </summary>
        public static async Task<string> DownloadAsync(UpdateInfo info)
        {
            if (string.IsNullOrWhiteSpace(info.InstallerUrl))
                throw new InvalidOperationException("Kein Installer-Asset in diesem Release gefunden.");

            var targetPath = Path.Combine(Path.GetTempPath(), $"Wrok-Setup-{info.LatestVersion}.exe");

            try
            {
                using (var resp = await _http.GetAsync(info.InstallerUrl, HttpCompletionOption.ResponseHeadersRead))
                {
                    resp.EnsureSuccessStatusCode();

                    using var fileStream = new FileStream(
                        targetPath, FileMode.Create, FileAccess.Write, FileShare.None);
                    await resp.Content.CopyToAsync(fileStream);
                }
            }
            catch
            {
                TryDelete(targetPath);
                throw;
            }

            // Downloads passieren zwar per HTTPS (Integritaet/Vertraulichkeit auf dem
            // Transportweg ist damit gegeben), aber wir starten das Ergebnis gleich
            // mit Administratorrechten - deshalb zusaetzlich die Authenticode-Signatur
            // gegen den Windows-Zertifikatsspeicher pruefen (derselbe SignTool=Certum
            // wie im offiziellen Release, siehe InstallScript.iss), statt der Datei
            // allein deshalb zu vertrauen, weil sie von der erwarteten URL kam.
            if (!AuthenticodeVerifier.IsTrusted(targetPath))
            {
                TryDelete(targetPath);
                throw new InvalidOperationException(
                    "Signaturprüfung des heruntergeladenen Installers fehlgeschlagen.");
            }

            return targetPath;
        }

        private static void TryDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [UpdateInstaller] Aufräumen fehlgeschlagen: {ex}");
            }
        }
    }

    /// <summary>
    /// Dünner Wrapper um WinVerifyTrust (statt nur X509Certificate.CreateFromSignedFile,
    /// das lediglich das eingebettete Zertifikat ausliest, ohne Signatur oder
    /// Vertrauenskette zu prüfen). Bewusst eigenständig und ohne UI (WTD_UI_NONE) -
    /// läuft im Hintergrund eines Update-Downloads, nicht interaktiv.
    /// </summary>
    internal static class AuthenticodeVerifier
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WINTRUST_FILE_INFO
        {
            public uint cbStruct;
            public string pcwszFilePath;
            public IntPtr hFile;
            public IntPtr pgKnownSubject;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WINTRUST_DATA
        {
            public uint cbStruct;
            public IntPtr pPolicyCallbackData;
            public IntPtr pSIPClientData;
            public uint dwUIChoice;
            public uint fdwRevocationChecks;
            public uint dwUnionChoice;
            public IntPtr pFile;
            public uint dwStateAction;
            public IntPtr hWVTStateData;
            public string? pwszURLReference;
            public uint dwProvFlags;
            public uint dwUIContext;
            public IntPtr pSignatureSettings;
        }

        private const uint WTD_UI_NONE = 2;
        private const uint WTD_REVOKE_NONE = 0;
        private const uint WTD_CHOICE_FILE = 1;
        private const uint WTD_STATEACTION_VERIFY = 1;
        private const uint WTD_STATEACTION_CLOSE = 2;
        private const uint WTD_SAFER_FLAG = 0x100;

        private static readonly Guid WINTRUST_ACTION_GENERIC_VERIFY_V2 =
            new("00AAC56B-CD44-11d0-8CC2-00C04FC295EE");

        [DllImport("wintrust.dll", ExactSpelling = true, SetLastError = true)]
        private static extern uint WinVerifyTrust(
            IntPtr hwnd, [MarshalAs(UnmanagedType.LPStruct)] Guid pgActionID, ref WINTRUST_DATA pWVTData);

        /// <summary>
        /// True, wenn die Datei eine gültige Authenticode-Signatur trägt, deren
        /// Zertifikatskette Windows vertraut. Prüft NICHT, wer genau unterschrieben
        /// hat (kein Pinning auf einen bestimmten Aussteller) - das würde bei jeder
        /// Zertifikatserneuerung (z. B. jährlich bei Certum) brechen. Es geht hier
        /// darum, eine unsignierte, beschädigte oder nachträglich manipulierte Datei
        /// abzulehnen, bevor sie mit Administratorrechten läuft.
        /// </summary>
        public static bool IsTrusted(string filePath)
        {
            var fileInfoSize = Marshal.SizeOf<WINTRUST_FILE_INFO>();
            var fileInfoPtr = Marshal.AllocHGlobal(fileInfoSize);
            try
            {
                Marshal.StructureToPtr(new WINTRUST_FILE_INFO
                {
                    cbStruct = (uint)fileInfoSize,
                    pcwszFilePath = filePath,
                    hFile = IntPtr.Zero,
                    pgKnownSubject = IntPtr.Zero
                }, fileInfoPtr, false);

                var data = new WINTRUST_DATA
                {
                    cbStruct = (uint)Marshal.SizeOf<WINTRUST_DATA>(),
                    dwUIChoice = WTD_UI_NONE,
                    fdwRevocationChecks = WTD_REVOKE_NONE,
                    dwUnionChoice = WTD_CHOICE_FILE,
                    pFile = fileInfoPtr,
                    dwStateAction = WTD_STATEACTION_VERIFY,
                    dwProvFlags = WTD_SAFER_FLAG
                };

                var result = WinVerifyTrust(IntPtr.Zero, WINTRUST_ACTION_GENERIC_VERIFY_V2, ref data);

                data.dwStateAction = WTD_STATEACTION_CLOSE;
                WinVerifyTrust(IntPtr.Zero, WINTRUST_ACTION_GENERIC_VERIFY_V2, ref data);

                return result == 0; // ERROR_SUCCESS
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [AuthenticodeVerifier] Prüfung fehlgeschlagen: {ex}");
                return false;
            }
            finally
            {
                Marshal.FreeHGlobal(fileInfoPtr);
            }
        }
    }
}
