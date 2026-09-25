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
