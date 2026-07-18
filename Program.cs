using System.Diagnostics;
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
        private const int WM_SHOWWINDOW = 0x0018;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool PostMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

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

                // Keep a rolling simple policy: rename existing file if bigger than 5 MB
                try
                {
                    const long maxSize = 5 * 1024 * 1024;
                    if (File.Exists(logFile))
                    {
                        var fi = new FileInfo(logFile);
                        if (fi.Length > maxSize)
                        {
                            var archived = Path.Combine(logsDir, $"app-{DateTime.UtcNow:yyyyMMddHHmmss}.log");
                            try { File.Move(logFile, archived); } catch { /* ignore */ }
                        }
                    }
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
                // Another instance is running — notify it to show its window and exit.
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
                _singleInstanceMutex.ReleaseMutex();
                _singleInstanceMutex.Dispose();
            }
            catch
            {
                // Swallow exceptions on shutdown; not critical.
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
                                    // Broadcast WM_SHOWWINDOW — MainForm.WndProc will react and bring the window forward.
                                    PostMessage((IntPtr)HWND_BROADCAST, WM_SHOWWINDOW, IntPtr.Zero, IntPtr.Zero);
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
