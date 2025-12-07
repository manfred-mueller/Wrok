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
            bool isNewInstance;
            _singleInstanceMutex = new Mutex(true, MutexName, out isNewInstance);

            if (!isNewInstance)
            {
                // Es läuft schon eine Instanz -> diese bitten, sich zu zeigen
                NotifyExistingInstance();
                return;
            }

            // Nur in der ersten Instanz: Named-Pipe-Server für SHOW-Nachrichten starten
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
                // ignorieren – beim Beenden nicht kritisch
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
                    client.Connect(500); // 0,5 Sekunden Timeout

                    using (var writer = new StreamWriter(client))
                    {
                        writer.AutoFlush = true;
                        writer.WriteLine("SHOW");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Fehler beim Benachrichtigen der bestehenden Instanz: {ex.Message}");
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
                                    // Broadcast WM_SHOWWINDOW -> MainForm.WndProc reagiert und bringt sich nach vorne
                                    PostMessage((IntPtr)HWND_BROADCAST, WM_SHOWWINDOW, IntPtr.Zero, IntPtr.Zero);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Fehler im Named Pipe Server: {ex.Message}");
                        // danach weiterlaufen und nächsten Client akzeptieren
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
