using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace Wrok
{
    /// <summary>
    /// Simuliert native Tastatureingaben (Text und Enter) für das aktive Fenster.
    /// </summary>
    internal sealed class NativeInput : IDisposable
    {
        public void TypeText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return;

            foreach (char c in text)
            {
                SendChar(c);
                Thread.Sleep(2);
            }
        }

        public void PressEnter()
        {
            SendKey(VK_RETURN, true);
        }

        private static void SendChar(char c)
        {
            var input = new INPUT
            {
                type = 1, // INPUT_KEYBOARD
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = 0,
                        wScan = c,
                        dwFlags = KEYEVENTF_UNICODE,
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };
            SendInput(1, new[] { input }, Marshal.SizeOf(typeof(INPUT)));
            input.U.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
            SendInput(1, new[] { input }, Marshal.SizeOf(typeof(INPUT)));
        }

        private static void SendKey(ushort key, bool pressAndRelease)
        {
            var down = new INPUT
            {
                type = 1,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = key,
                        wScan = 0,
                        dwFlags = 0,
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };
            var up = down;
            up.U.ki.dwFlags = KEYEVENTF_KEYUP;

            SendInput(1, new[] { down }, Marshal.SizeOf(typeof(INPUT)));
            if (pressAndRelease)
                SendInput(1, new[] { up }, Marshal.SizeOf(typeof(INPUT)));
        }

        public void Dispose()
        {
            // Keine Ressourcen zu bereinigen
        }

        private const uint KEYEVENTF_UNICODE = 0x0004;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const ushort VK_RETURN = 0x0D;

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public uint type;
            public InputUnion U;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion
        {
            [FieldOffset(0)]
            public KEYBDINPUT ki;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);
    }
}