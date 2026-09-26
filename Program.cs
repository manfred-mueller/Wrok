using System.Diagnostics;
using System.Linq;
using System.IO.Pipes;
using System.Runtime.InteropServices;

namespace Wrok
{
    internal static class Program
    {
        private const string MutexName = "Global\\Wrok_SingleInstanceMutex";
        private const string PipeName = "Wrok_SingleInstancePipe";

        private static Mutex? _singleInstanceMutex;

        private const int HWND_BROADCAST = 0xFFFF;

        // Eigene, private Fensternachricht statt der echten Windows-Systemnachricht
        // WM_SHOWWINDOW: die wurde zuvor per HWND_BROADCAST an ALLE Top-Level-Fenster
        // auf dem Desktop geschickt - andere laufende Anwendungen mit eigener
        // WM_SHOWWINDOW-Behandlung koennten das faelschlich als echten
        // Sichtbarkeits-Wechsel interpretieren. RegisterWindowMessage liefert fuer
        // denselben String prozessuebergreifend immer denselben Wert und wird von
        // fremden Fenstern einfach ignoriert (Standard-Win32-Muster fuer private
        // IPC-Nachrichten, siehe z. B. "TaskbarCreated").
        internal const string ShowRequestMessageName = "Wrok_ShowRequest_F3B2A6D1";

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool PostMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern uint RegisterWindowMessage(string lpString);

        private static readonly int WM_WROK_SHOW = (int)RegisterWindowMessage(ShowRequestMessageName);

        [STAThread]
        private static void Main()
        {
            // --- Persistent file logging setup (writes Trace to %LOCALAPPDATA%\Wrok\logs\app.log) ---
            try
            {
                var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                var appDir = Path.Combine(localAppData, "Wrok");
                var logsDir = Path.Combine(appDir, "logs");

                Directory.CreateDirectory(logsDir);

                var logFile = Path.Combine(logsDir, "app.log");

                // Keep a rolling simple policy: rename existing file if bigger than 5 MB,
                // then prune old archives beyond a small retention count - otherwise the
                // archived app-<timestamp>.log files pile up forever for a long-running
                // "portable" tool that nobody ever manually cleans out.
                try
                {
                    const long maxSize = 5 * 1024 * 1024;
                    const int maxArchives = 5;
                    if (File.Exists(logFile))
                    {
                        var fi = new FileInfo(logFile);
                        if (fi.Length > maxSize)
                        {
                            var archived = Path.Combine(logsDir, $"app-{DateTime.UtcNow:yyyyMMddHHmmss}.log");
                            try { File.Move(logFile, archived); } catch { /* ignore */ }
                        }
                    }

                    try
                    {
                        var oldArchives = new DirectoryInfo(logsDir)
                            .GetFiles("app-*.log")
                            .OrderByDescending(f => f.LastWriteTimeUtc)
                            .Skip(maxArchives);
                        foreach (var old in oldArchives)
                        {
                            try { old.Delete(); } catch { /* ignore einzelne Datei */ }
                        }
                    }
                    catch { /* defensive - Aufraeumen ist nicht kritisch */ }
                }
                catch { /* defensive */ }

                var tw = new StreamWriter(new FileStream(logFile, FileMode.Append, FileAccess.Write, FileShare.Read))
                {
                    AutoFlush = true
                };

                Trace.Listeners.Add(new TextWriterTraceListener(tw));
                Trace.AutoFlush = true;
                Trace.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Wrok starting (pid={Process.GetCurrentProcess().Id})");
            }
            catch
            {
                // If logging setup fails, continue without file logging.
            }
            // --- end logging setup ---

            bool isNewInstance;
            _singleInstanceMutex = new Mutex(true, MutexName, out isNewInstance);

            if (!isNewInstance)
            {
                // Another instance is running � notify it to show its window and exit.
                NotifyExistingInstance();
                return;
            }

            // In the single (first) instance: start a named-pipe server to receive SHOW messages.
            StartNamedPipeServer();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());

            try
            {
                // Null-conditional: nach RestartApplication() ist das Feld bereits
                // freigegeben und genullt, dann soll hier einfach nichts passieren.
                _singleInstanceMutex?.ReleaseMutex();
                _singleInstanceMutex?.Dispose();
            }
            catch
            {
                // Swallow exceptions on shutdown; not critical.
            }
        }

        /// <summary>
        /// Startet eine neue Instanz und beendet die aktuelle geordnet. Für
        /// Einstellungen, die erst beim Neuerzeugen der CoreWebView2Environment
        /// wirken (z. B. Proxy). Mutex zuerst freigeben, dann die neue Instanz
        /// starten – sonst hält die alte Instanz den Mutex noch, wenn die neue
        /// ihn prüft, und die neue hielte sich fälschlich für einen Zweitstart.
        /// </summary>
        public static void RestartApplication()
        {
            try
            {
                _singleInstanceMutex?.ReleaseMutex();
                _singleInstanceMutex?.Dispose();
                _singleInstanceMutex = null;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"RestartApplication: Mutex-Freigabe fehlgeschlagen: {ex}");
            }

            try
            {
                Process.Start(Environment.ProcessPath ?? Application.ExecutablePath);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"RestartApplication: neue Instanz konnte nicht gestartet werden: {ex}");
            }

            Application.Exit();
        }

        /// <summary>
        /// Startet den (bereits heruntergeladenen und signaturgeprüften, siehe
        /// UpdateInstaller) Installer und beendet die laufende Instanz danach - fürs
        /// Selbst-Update aus der App heraus, analog zu RestartApplication() oben,
        /// nur eben mit einem zwischengeschalteten Setup statt einem reinen
        /// Neustart.
        ///
        /// Ablauf:
        ///  1. Installer per ShellExecute starten (nötig für die UAC-Erhöhung über
        ///     dessen eingebettetes Manifest - PrivilegesRequired=admin in
        ///     InstallScript.iss; ein normaler Process.Start ohne ShellExecute würde
        ///     nur mit "Elevation erforderlich" fehlschlagen, statt den UAC-Dialog
        ///     zu zeigen).
        ///  2. Einen von uns unabhängigen kleinen Warte-Prozess anstossen, der auf
        ///     das Ende des Installers wartet und danach Wrok neu startet. Not-
        ///     wendig, weil unser eigener Prozess selbst nicht mehr da sein wird,
        ///     wenn der Installer fertig ist: Entweder er beendet sich gleich unten
        ///     selbst, oder - falls das Timing eng wird - vorher schon durch
        ///     CloseApplications=force im Setup (das über den Windows Restart
        ///     Manager per WM_QUERYENDSESSION genau die laufende Instanz schliesst,
        ///     deren Programmdatei ersetzt werden soll - siehe auch _sessionEnding
        ///     in MainForm.cs, das für genau diesen Fall ein echtes Beenden statt
        ///     eines Minimierens in den Tray erzwingt).
        ///  3. Mutex freigeben und Application.Exit() - wie bei RestartApplication().
        /// </summary>
        public static void RestartApplicationForUpdate(string installerPath)
        {
            var exePath = Environment.ProcessPath ?? Application.ExecutablePath;

            try
            {
                var installer = Process.Start(new ProcessStartInfo
                {
                    FileName = installerPath,
                    Arguments = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /NOICONS",
                    UseShellExecute = true
                });

                if (installer != null)
                {
                    StartRelaunchWatcher(installer.Id, exePath);
                }
                else
                {
                    Trace.WriteLine("RestartApplicationForUpdate: Installer-Prozess konnte nicht ermittelt werden.");
                }
            }
            catch (Exception ex)
            {
                // Installer nicht gestartet (z. B. UAC-Dialog abgebrochen) - dann auch
                // nicht die laufende Instanz beenden, sonst bleibt der Nutzer ganz
                // ohne laufendes Wrok zurück.
                Trace.WriteLine($"RestartApplicationForUpdate: Installer konnte nicht gestartet werden: {ex}");
                return;
            }

            try
            {
                _singleInstanceMutex?.ReleaseMutex();
                _singleInstanceMutex?.Dispose();
                _singleInstanceMutex = null;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"RestartApplicationForUpdate: Mutex-Freigabe fehlgeschlagen: {ex}");
            }

            Application.Exit();
        }

        /// <summary>
        /// Wartet - unabhängig vom (gleich endenden) aktuellen Prozess - auf das Ende
        /// des Installers und startet danach Wrok neu. Läuft bewusst unerhöht
        /// (PowerShell selbst braucht keine Admin-Rechte; nur der Installer, der schon
        /// separat läuft, brauchte sie). Rein informativ: schlägt das Starten des
        /// Watchers fehl, bleibt Wrok nach dem Update einfach zu - der Nutzer kann es
        /// dann manuell wieder öffnen, das Update selbst ist davon nicht betroffen.
        /// </summary>
        private static void StartRelaunchWatcher(int installerProcessId, string exePath)
        {
            try
            {
                var script =
                    $"Wait-Process -Id {installerProcessId} -ErrorAction SilentlyContinue; " +
                    "Start-Sleep -Milliseconds 500; " +
                    $"Start-Process -FilePath '{exePath}'";

                Process.Start(new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -WindowStyle Hidden -Command \"{script}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                });
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"StartRelaunchWatcher: Warte-Prozess konnte nicht gestartet werden: {ex}");
            }
        }

        private static void NotifyExistingInstance()
        {
            try
            {
                using (var client = new NamedPipeClientStream(
                           ".",
                           PipeName,
                           PipeDirection.Out,
                           PipeOptions.Asynchronous))
                {
                    client.Connect(500); // 0.5 second timeout

                    using (var writer = new StreamWriter(client))
                    {
                        writer.AutoFlush = true;
                        writer.WriteLine("SHOW");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to notify existing instance: {ex.Message}");
            }
        }

        private static void StartNamedPipeServer()
        {
            var thread = new Thread(() =>
            {
                while (true)
                {
                    try
                    {
                        using (var server = new NamedPipeServerStream(
                                   PipeName,
                                   PipeDirection.In,
                                   1,
                                   PipeTransmissionMode.Byte,
                                   PipeOptions.Asynchronous))
                        {
                            server.WaitForConnection();

                            using (var reader = new StreamReader(server))
                            {
                                var message = reader.ReadLine();
                                if (string.Equals(message, "SHOW", StringComparison.OrdinalIgnoreCase))
                                {
                                    // Eigene Nachricht broadcasten - MainForm.WndProc reagiert darauf und holt das Fenster nach vorne.
                                    PostMessage((IntPtr)HWND_BROADCAST, WM_WROK_SHOW, IntPtr.Zero, IntPtr.Zero);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Named pipe server error: {ex.Message}");
                        // Continue loop and accept next client.
                    }
                }
            })
            {
                IsBackground = true,
                Name = "Wrok_SingleInstancePipeServer"
            };

            thread.Start();
        }
    }
}
