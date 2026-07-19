using System.Diagnostics;

namespace Wrok
{
    /// <summary>
    /// Überwacht Nutzerinaktivität und meldet sich, wenn die eingestellte Zeit
    /// ohne Eingabe verstrichen ist.
    ///
    /// Aktivität kommt aus zwei Quellen: Maus und Tastatur im Programm selbst
    /// (über einen Nachrichtenfilter) sowie Ereignisse aus dem WebView, die der
    /// JS-Helfer meldet. Beide rufen <see cref="Reset"/> auf.
    /// </summary>
    internal sealed class InactivityWatcher : IDisposable
    {
        private readonly Form   _form;
        private readonly Action _onIdle;

        /// <summary>
        /// Liefert das derzeit offene Medienfenster, falls es eines gibt.
        /// Der Betrachter ist ein eigenes Top-Level-Fenster; ohne ihn wuerde ein
        /// Zeiger ueber dem angezeigten Bild nicht als Aufmerksamkeit zaehlen.
        /// </summary>
        private readonly Func<Form?>? _extraWindow;

        // Zeigerposition der letzten Pruefung: Nur ein *bewegter* Zeiger gilt als
        // Aktivitaet. Ein bloss abgelegter Zeiger hielt die Ueberwachung sonst
        // endlos auf, obwohl niemand am Rechner sitzt.
        private Point _lastCursor = Cursor.Position;

        // Woher kam der letzte Reset? Wird beim Ausloesen protokolliert und dient
        // sonst der Fehlersuche.
        private string _lastSource = "-";

        // Zeitpunkt der letzten Protokollzeile (siehe <see cref="TraceEnabled"/>).
        private DateTime _lastLogged = DateTime.MinValue;

        /// <summary>
        /// Schreibt den Verlauf der Ueberwachung mit, wenn die Umgebungsvariable
        /// WROK_DEBUG_INACTIVITY auf 1 steht.
        ///
        /// Bewusst eine Umgebungsvariable und keine Konstante: So laesst sich die
        /// Diagnose ohne Neubau einschalten – Variable setzen, Wrok starten, fertig.
        /// Standardmaessig aus, denn im Sekundentakt kaemen rund 17.000 Zeilen
        /// pro Tag zusammen.
        /// </summary>
        private static readonly bool TraceEnabled =
            Environment.GetEnvironmentVariable("WROK_DEBUG_INACTIVITY") == "1";

        private System.Windows.Forms.Timer? _timer;
        private TimeSpan _timeout = TimeSpan.Zero;
        private bool     _enabled;
        private DateTime _lastActivity = DateTime.UtcNow;
        private readonly object _lock = new();

        private ActivityMessageFilter? _filter;

        public InactivityWatcher(Form form, Action onIdle, Func<Form?>? extraWindow = null)
        {
            _form         = form;
            _onIdle       = onIdle;
            _extraWindow  = extraWindow;
        }

        /// <summary>Aktuell eingestellte Zeitspanne in Sekunden (0 = abgeschaltet).</summary>
        public int TimeoutSeconds => (int)_timeout.TotalSeconds;

        /// <summary>
        /// Liest die Einstellung. Beim allerersten Start wird bewusst abgeschaltet,
        /// damit niemand von einem verschwindenden Fenster überrascht wird.
        /// </summary>
        public void LoadSettings()
        {
            if (!Properties.Settings.Default.InactivityConfigured)
            {
                _timeout = TimeSpan.Zero;
                _enabled = false;
                Properties.Settings.Default.InactivityTimeoutSeconds = 0;
                Properties.Settings.Default.InactivityConfigured     = true;
                Properties.Settings.Default.Save();
            }
            else
            {
                int saved = Properties.Settings.Default.InactivityTimeoutSeconds;
                _timeout  = TimeSpan.FromSeconds(saved);
                _enabled  = saved > 0;
            }
        }

        /// <summary>Startet die Überwachung mit der geladenen Einstellung.</summary>
        public void Start()
        {
            StopTimer();

            _timer = new System.Windows.Forms.Timer { Interval = TickInterval() };
            _timer.Tick += OnTick;

            if (_enabled && _timeout.TotalMilliseconds > 0)
            {
                lock (_lock) { _lastActivity = DateTime.UtcNow; }
                _timer.Start();
            }

            AttachMessageFilter();
        }

        /// <summary>
        /// Taktrate der Pruefung: Sekundentakt, solange eine Zeitspanne gesetzt ist.
        /// Der Zeitgeber prueft nur, er misst nicht selbst.
        /// </summary>
        private int TickInterval() =>
            _timeout.TotalMilliseconds > 0
                ? (int)Math.Min(1000, _timeout.TotalMilliseconds)
                : 60_000;

        /// <summary>Ändert die Zeitspanne und persistiert sie.</summary>
        public void SetTimeout(int seconds)
        {
            Properties.Settings.Default.InactivityTimeoutSeconds = seconds;
            Properties.Settings.Default.Save();

            _timeout = TimeSpan.FromSeconds(seconds);
            _enabled = seconds > 0;

            if (_enabled)
            {
                // Das Intervall stammt aus Start() und passte zur damaligen
                // Zeitspanne. Stand die Ueberwachung beim Programmstart auf „Keine“,
                // liefe der Zeitgeber sonst weiter im Minutentakt und die neu
                // eingestellte Zeitspanne wuerde erst mit bis zu einer Minute
                // Verspaetung bemerkt.
                if (_timer != null) _timer.Interval = TickInterval();

                lock (_lock) { _lastActivity = DateTime.UtcNow; }
                try { _timer?.Start(); } catch { }
            }
            else
            {
                try { _timer?.Stop(); } catch { }
            }
        }

        /// <summary>
        /// Meldet Nutzeraktivität. Wird auch vom WebView aus aufgerufen.
        ///
        /// <paramref name="source"/> dient allein der Fehlersuche: Er wird gemerkt
        /// und im Protokoll ausgegeben, damit sich nachvollziehen laesst, *was* die
        /// Uhr immer wieder zurueckstellt. Einzeln protokolliert wird er nicht –
        /// bei jeder Mausbewegung waere das Protokoll unbrauchbar.
        /// </summary>
        public void Reset(string source = "?")
        {
            if (!_enabled || _timer == null) return;
            lock (_lock) { _lastActivity = DateTime.UtcNow; _lastSource = source; }
            try { if (!_timer.Enabled) _timer.Start(); } catch { }
        }

        private void OnTick(object? sender, EventArgs e)
        {
            if (!_enabled || _timeout.TotalMilliseconds <= 0) return;

            TimeSpan elapsed;
            string   source;
            lock (_lock) { elapsed = DateTime.UtcNow - _lastActivity; source = _lastSource; }

            try
            {
                // Ein *bewegter* Zeiger ueber einem unserer Fenster gilt als
                // Aktivitaet, auch ohne Klick – sonst verschwaende das Fenster
                // beim blossen Lesen.
                //
                // Frueher genuegte hier, dass der Zeiger irgendwo im Fenster lag.
                // Zusammen mit dem Medienfenster, das Wrok bildschirmfuellend
                // daneben legt, bedeckten beide Fenster praktisch die ganze
                // Arbeitsflaeche: Der Zeiger lag damit fast immer „drin“, und die
                // Ueberwachung loeste nie aus. Ebenso entfaellt _form.Focused –
                // dass Wrok im Vordergrund steht, ist gerade kein Beleg dafuer,
                // dass noch jemand davor sitzt.
                var pos   = Cursor.Position;
                bool moved = pos != _lastCursor;
                bool over  = IsOverOwnWindow(pos);
                _lastCursor = pos;

                // Hoechstens alle fuenf Sekunden eine Zeile, sonst waere das
                // Protokoll bei Sekundentakt nicht mehr zu lesen.
                if (TraceEnabled && (DateTime.UtcNow - _lastLogged).TotalSeconds >= 5)
                {
                    _lastLogged = DateTime.UtcNow;
                    Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [InactivityWatcher] " +
                        $"verstrichen={elapsed.TotalSeconds:F0}s von {_timeout.TotalSeconds:F0}s, " +
                        $"Zeiger bewegt={moved}, ueber eigenem Fenster={over}, letzter Reset von={source}");
                }

                if (moved && over)
                {
                    lock (_lock) { _lastActivity = DateTime.UtcNow; _lastSource = "Zeigerbewegung"; }
                    return;
                }
            }
            catch { }

            if (elapsed >= _timeout)
            {
                // Diese eine Zeile bleibt immer: sie faellt genau einmal je
                // Minimierung an und beantwortet spaeter die Frage, warum das
                // Fenster verschwunden ist.
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [InactivityWatcher] " +
                    $"Zeitspanne erreicht ({elapsed.TotalSeconds:F0}s, letzte Aktivitaet von {source}) – minimiere.");
                try { _timer?.Stop(); } catch { }
                _onIdle();
            }
        }

        /// <summary>Liegt der Punkt ueber dem Hauptfenster oder dem Medienfenster?</summary>
        private bool IsOverOwnWindow(Point pos)
        {
            try
            {
                if (_form.Visible && _form.WindowState != FormWindowState.Minimized &&
                    _form.Bounds.Contains(pos))
                    return true;

                if (_extraWindow?.Invoke() is { IsDisposed: false, Visible: true } extra &&
                    extra.Bounds.Contains(pos))
                    return true;
            }
            catch { }

            return false;
        }

        private void StopTimer()
        {
            if (_timer == null) return;
            try { _timer.Stop(); _timer.Tick -= OnTick; _timer.Dispose(); } catch { }
            _timer = null;
        }

        // ------------------------------------------------------------------
        // Aktivitätserkennung im Programm
        // ------------------------------------------------------------------

        private void AttachMessageFilter()
        {
            if (_filter != null) return;
            _filter = new ActivityMessageFilter(this);
            try { Application.AddMessageFilter(_filter); }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [InactivityWatcher] Nachrichtenfilter fehlgeschlagen: {ex}");
                _filter = null;
            }
        }

        private void DetachMessageFilter()
        {
            if (_filter == null) return;
            try { Application.RemoveMessageFilter(_filter); } catch { }
            _filter = null;
        }

        /// <summary>
        /// Fängt Maus- und Tastaturmeldungen der Anwendung ab, um Aktivität zu
        /// erkennen. Hält den Watcher nur schwach, damit er nicht am Leben bleibt,
        /// falls das Entfernen des Filters einmal misslingt.
        /// </summary>
        private sealed class ActivityMessageFilter : IMessageFilter
        {
            private readonly WeakReference<InactivityWatcher> _ref;
            public ActivityMessageFilter(InactivityWatcher watcher) => _ref = new(watcher);

            public bool PreFilterMessage(ref Message m)
            {
                const int WM_MOUSEMOVE   = 0x0200;
                const int WM_LBUTTONDOWN = 0x0201;
                const int WM_RBUTTONDOWN = 0x0204;
                const int WM_MBUTTONDOWN = 0x0207;
                const int WM_MOUSEWHEEL  = 0x020A;
                const int WM_KEYDOWN     = 0x0100;
                const int WM_SYSKEYDOWN  = 0x0104;

                if (m.Msg is WM_MOUSEMOVE or WM_LBUTTONDOWN or WM_RBUTTONDOWN or
                             WM_MBUTTONDOWN or WM_MOUSEWHEEL or WM_KEYDOWN or WM_SYSKEYDOWN)
                    if (_ref.TryGetTarget(out var target))
                        target.Reset($"Fenstermeldung 0x{m.Msg:X4}");

                return false;   // Nachricht immer weiterreichen
            }
        }

        public void Dispose()
        {
            DetachMessageFilter();
            StopTimer();
        }
    }
}
