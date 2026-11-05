using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Wrok
{
    public partial class MainForm
    {
        // ------------------------------------------------------------------
        // Bild aus Zwischenablage öffnen
        // ------------------------------------------------------------------

        /// <summary>
        /// Öffnet die Bild-URL aus der Zwischenablage (Rechtsklick im WebView →
        /// „Bildadresse kopieren") ohne Rückfrage direkt im Viewer-Fenster.
        /// Ebenso lokale Medien: eine im Explorer kopierte Datei (Strg+C) oder ein
        /// kopierter Dateipfad („Als Pfad kopieren", auch als file:///-URI).
        /// Enthält die Zwischenablage nichts davon, wird auf das zuletzt
        /// gespeicherte Bild zurückgefallen – so macht Strg+Ö immer etwas Sinnvolles.
        /// </summary>
        private async Task OpenImageFromClipboardAsync()
        {
            Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MainForm] Strg+Ö empfangen, OpenImageFromClipboardAsync gestartet.");

            if (_webViewManager == null)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MainForm] OpenImageFromClipboardAsync: _webViewManager ist null, Abbruch.");
                return;
            }

            string url = string.Empty;
            try
            {
                if (Clipboard.ContainsText())
                    url = (Clipboard.GetText() ?? string.Empty).Trim();
            }
            catch (Exception ex) { Log(ex, "Zwischenablage konnte nicht gelesen werden"); }

            Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MainForm] OpenImageFromClipboardAsync: Zwischenablage-Inhalt='{url}', IsHttpUrl={IsHttpUrl(url)}");

            // Lokale Datei (Explorer-Kopie oder Pfad als Text) hat Vorrang vor dem
            // Rückfall auf das letzte Bild.
            var localPath = GetLocalMediaPathFromClipboard(url);
            if (localPath != null)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MainForm] OpenImageFromClipboardAsync: lokale Datei '{localPath}'");
                OpenLocalMedia(localPath);
                return;
            }

            if (!IsHttpUrl(url))
            {
                OpenLastImage();
                return;
            }

            // Der Download kann einen Moment dauern – Wartecursor als Rückmeldung.
            try { Cursor.Current = Cursors.AppStarting; } catch { }
            try
            {
                // Grok-URLs haben keine Dateiendung – Typ über den Content-Type klären.
                var contentType = await _webViewManager.ProbeContentTypeAsync(url, TimeSpan.FromSeconds(10));
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MainForm] OpenImageFromClipboardAsync: Content-Type='{contentType}'");

                if (contentType != null &&
                    contentType.Contains("video", StringComparison.OrdinalIgnoreCase))
                {
                    ShowVideoViewer(url);
                    return;
                }

                bool shown = await _webViewManager.ShowImageViewerAsync(url);
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MainForm] OpenImageFromClipboardAsync: ShowImageViewerAsync -> {shown}");
                if (!shown)
                    MessageBox.Show(this, Properties.Resources.ImageLoadFailed,
                        Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                // Vorher fehlte dieser catch: eine Exception hier (z. B. Timeout bei
                // ProbeContentTypeAsync/ShowImageViewerAsync) verschwand komplett
                // spurlos, weil die Methode "fire-and-forget" per _ = ... aufgerufen
                // wird - eine ungefangene Exception in einem verworfenen Task wird
                // von .NET standardmaessig NICHT sichtbar gemeldet (kein Absturz,
                // kein Log). Das erklaert vermutlich "Strg+Ö tut gar nichts".
                Log(ex, "OpenImageFromClipboardAsync fehlgeschlagen");
                MessageBox.Show(this, Properties.Resources.ImageLoadFailed,
                    Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                try { Cursor.Current = Cursors.Default; } catch { }
            }
        }

        // ------------------------------------------------------------------
        // Medien anzeigen und nebeneinander anordnen
        // ------------------------------------------------------------------

        /// <summary>
        /// Zeigt ein Bild im Medienfenster, links auf voller Bildschirmhöhe. Es gibt immer
        /// nur ein Medienfenster: Ein neues Bild ersetzt den Inhalt des bestehenden Fensters.
        /// </summary>
        internal void ShowImageViewer(Bitmap bmp, string sourceUrl)
        {
            try
            {
                string path = SaveViewImage(bmp);
                var viewer  = EnsureMediaViewer();
                viewer.ShowImage(path, new Size(bmp.Width, bmp.Height));
            }
            catch (Exception ex)
            {
                Log(ex, "Bild konnte nicht angezeigt werden");
            }
            finally
            {
                bmp.Dispose();
            }
        }

        /// <summary>
        /// Zeigt ein Video im Medienfenster. Die echten Abmessungen sind erst nach dem
        /// Laden der Metadaten bekannt – das Fenster wird dann nachjustiert.
        /// </summary>
        internal void ShowVideoViewer(string url)
        {
            if (_webViewManager == null) return;

            var viewer = EnsureMediaViewer();

            // Ein Bild (z. B. aus einem Video-Slot) wird im selben Fenster als Bild gezeigt.
            // Die Größe kennt das Fenster erst nach dem Laden. Haltepunkte-Anhang abtrennen,
            // damit die Endung geprüft werden kann; ShowVideo wertet ihn selbst aus.
            string bare = VideoViewerForm.SplitStops(url, out _);
            if (IsImageLocation(bare))
                viewer.ShowImage(bare, Size.Empty);
            else
                viewer.ShowVideo(url);

            Properties.Settings.Default.LastImageUrl  = url;
            Properties.Settings.Default.LastImagePath = string.Empty;   // Video: nur URL
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Erkennt anhand der Endung, ob eine Adresse oder ein lokaler Pfad ein Bild ist.
        /// Unbekannte Endungen gelten als Video, wie bisher.
        /// </summary>
        private static bool IsImageLocation(string target)
        {
            try
            {
                string path = target.StartsWith("http", StringComparison.OrdinalIgnoreCase)
                    ? new Uri(target).AbsolutePath
                    : target;
                return LocalImageExtensions.Contains(Path.GetExtension(path));
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Liefert das offene Medienfenster oder öffnet ein neues. Vorhandene Fenster
        /// behalten ihre Position und Größe; der Inhalt wechselt erst bei ShowImage/ShowVideo.
        /// </summary>
        private VideoViewerForm EnsureMediaViewer()
        {
            if (_mediaViewer is VideoViewerForm { IsDisposed: false } existing)
            {
                if (!existing.Visible) existing.Show();
                return existing;
            }

            var viewer = new VideoViewerForm(_webViewManager?.CoreEnvironment);
            RegisterMediaViewer(viewer);

            viewer.VideoSlotRequested += slot =>
                BeginInvoke((System.Windows.Forms.MethodInvoker)(() => PlayVideoSlot(slot - 1)));

            viewer.VideoSizeKnown += size =>
            {
                if (viewer.IsDisposed) return;
                try { ArrangeSideBySide(viewer, viewer.WidthForHeight(WorkArea.Height, size)); }
                catch (Exception ex) { Log(ex, "Medien-Anordnung nach Metadaten fehlgeschlagen"); }
            };

            // Vorläufig hochkant anordnen; sobald die Abmessungen da sind, korrigieren.
            ArrangeSideBySide(viewer, (int)(WorkArea.Height * 9.0 / 16.0));
            viewer.Show();

            // Das neue Fenster bekommt beim Öffnen den Fokus. Das Hauptfenster holt ihn
            // zurück, damit F11/F12 und Strg+1 bis Strg+9 dort wie gewohnt wirken.
            if (Visible && WindowState != FormWindowState.Minimized)
                Activate();

            return viewer;
        }

        /// <summary>
        /// Speichert das Bild als PNG im lokalen Viewer-Ordner. Ältere Dateien dieses
        /// Ordners werden entfernt, damit er nicht wächst.
        /// </summary>
        private static string SaveViewImage(Bitmap bmp)
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                   "Wrok", "view");
            Directory.CreateDirectory(dir);

            var path = Path.Combine(dir, $"view-{DateTime.Now:yyyyMMddHHmmssfff}.png");
            bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);

            foreach (var old in Directory.GetFiles(dir, "view-*.png"))
            {
                if (string.Equals(old, path, StringComparison.OrdinalIgnoreCase)) continue;
                try { File.Delete(old); } catch { /* wird beim nächsten Mal erneut versucht */ }
            }
            return path;
        }

        private Rectangle WorkArea => Screen.FromControl(this).WorkingArea;

        private void CloseCurrentMedia()
        {
            if (_mediaViewer is { IsDisposed: false } previous)
            {
                try { previous.Close(); }
                catch (Exception ex) { Log(ex, "Vorheriges Medienfenster schließen fehlgeschlagen"); }
            }
        }

        private void RegisterMediaViewer(Form viewer)
        {
            _mediaViewer = viewer;
            UpdateBrowserAccelerators();
            viewer.FormClosed += (_, _) =>
            {
                if (ReferenceEquals(_mediaViewer, viewer)) _mediaViewer = null;
                UpdateBrowserAccelerators();
            };
        }

        /// <summary>
        /// Browser-Tastenkürzel im Hauptfenster sind nur dann aktiv, wenn weder der
        /// F-Tasten-Makromodus noch ein offenes Video läuft. Solange ein Video offen
        /// ist, erreichen F11/F12 den Seitencode und steuern das Video.
        /// </summary>
        internal void UpdateBrowserAccelerators()
        {
            if (_webView?.CoreWebView2 == null) return;
            bool videoOpen = _mediaViewer is VideoViewerForm { IsDisposed: false } video && video.IsVideo;
            _webView.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled =
                !(Properties.Settings.Default.FKeyMacroOverrideEnabled || videoOpen);
        }

        /// <summary>Pausiert bzw. startet das Video im offenen Viewer, ohne den Fokus zu verschieben.</summary>
        internal void ControlVideoPlayback(bool? play)
        {
            if (_mediaViewer is not VideoViewerForm { IsDisposed: false } video) return;

            if (play is bool start) video.SetPlayback(start);
            else video.TogglePlayback();
        }

        private void ArrangeSideBySide(Form viewer, int desiredWidth)
        {
            try
            {
                var wa = WorkArea;

                // Breite aus dem Seitenverhältnis bei voller Bildschirmhöhe –
                // aber gedeckelt, damit rechts genug für Wrok übrig bleibt.
                int viewerWidth = Math.Clamp(desiredWidth, 240, (int)(wa.Width * 0.6));

                viewer.StartPosition = FormStartPosition.Manual;
                viewer.Bounds = new Rectangle(wa.Left, wa.Top, viewerWidth, wa.Height);

                // Ist Wrok gar nicht sichtbar (Tray/minimiert), nur das Medium setzen.
                if (!this.Visible || this.WindowState == FormWindowState.Minimized)
                    return;

                try
                {
                    _suppressWindowSave = true;
                    if (this.WindowState != FormWindowState.Normal)
                        this.WindowState = FormWindowState.Normal;

                    this.Bounds = new Rectangle(wa.Left + viewerWidth, wa.Top,
                                                wa.Width - viewerWidth, wa.Height);
                }
                finally { _suppressWindowSave = false; }
            }
            catch (Exception ex)
            {
                Log(ex, "ArrangeSideBySide fehlgeschlagen");
            }
        }

        /// <summary>
        /// Öffnet das zuletzt angezeigte Bild aus dem lokalen Zwischenspeicher.
        /// Funktioniert ohne Netzwerk und ohne gültige Session, da die Bilddatei
        /// beim letzten Öffnen mitgespeichert wurde.
        /// </summary>
        private void OpenLastImage()
        {
            var path = Properties.Settings.Default.LastImagePath;
            var lastUrl = Properties.Settings.Default.LastImageUrl;

            // Kein lokaler Pfad, aber eine URL gemerkt → das war ein Video.
            // Videos werden nicht zwischengespeichert, sondern neu gestreamt
            // bzw. – bei einem lokalen Video – direkt von der Platte abgespielt.
            if (string.IsNullOrWhiteSpace(path) &&
                (IsHttpUrl(lastUrl) || (IsLocalVideoFile(lastUrl) && File.Exists(lastUrl))))
            {
                ShowVideoViewer(lastUrl);
                return;
            }

            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                MessageBox.Show(this, Properties.Resources.NoLastImage,
                    Properties.Resources.OpenLastImage, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Datei-Stream sofort wieder freigeben: Bitmap aus einem Stream behält
                // sonst eine Referenz darauf und sperrt die Datei fürs Überschreiben.
                Bitmap bmp;
                using (var fs = File.OpenRead(path))
                using (var decoded = new Bitmap(fs))
                    bmp = new Bitmap(decoded);

                ShowImageViewer(bmp, string.IsNullOrWhiteSpace(lastUrl) ? path : lastUrl);
            }
            catch (Exception ex)
            {
                Log(ex, "OpenLastImage fehlgeschlagen");
                MessageBox.Show(this, Properties.Resources.ImageLoadFailed,
                    Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ------------------------------------------------------------------
        // Lokale Medien (Strg+Ö mit kopierter Datei bzw. kopiertem Pfad)
        // ------------------------------------------------------------------

        // Bilder: was GDI+ (System.Drawing) ohne Zusatzcodecs laden kann.
        private static readonly HashSet<string> LocalImageExtensions =
            new(StringComparer.OrdinalIgnoreCase) { ".jpg", ".jpeg", ".jpe", ".png", ".bmp", ".gif", ".tif", ".tiff", ".ico" };

        // Videos: was der WebView2-Player (<video>) abspielen kann.
        private static readonly HashSet<string> LocalVideoExtensions =
            new(StringComparer.OrdinalIgnoreCase) { ".mp4", ".m4v", ".webm", ".mov", ".ogv" };

        private static bool IsLocalVideoFile(string? path) =>
            !string.IsNullOrWhiteSpace(path) && LocalVideoExtensions.Contains(Path.GetExtension(path));

        /// <summary>
        /// Liefert den Pfad einer lokalen Datei aus der Zwischenablage oder null.
        /// Berücksichtigt eine im Explorer kopierte Datei (Dateiliste, bei mehreren
        /// die erste) sowie einen als Text kopierten Pfad.
        /// </summary>
        private string? GetLocalMediaPathFromClipboard(string clipboardText)
        {
            try
            {
                if (Clipboard.ContainsFileDropList())
                {
                    foreach (var f in Clipboard.GetFileDropList())
                        if (!string.IsNullOrWhiteSpace(f) && File.Exists(f))
                            return f;
                }
            }
            catch (Exception ex) { Log(ex, "Dateiliste aus der Zwischenablage konnte nicht gelesen werden"); }

            return TryParseLocalPath(clipboardText);
        }

        /// <summary>
        /// Wandelt einen als Text kopierten Pfad in einen vorhandenen Dateipfad um.
        /// Akzeptiert Anführungszeichen („Als Pfad kopieren") und file:///-URIs;
        /// bei mehreren Zeilen zählt nur die erste.
        /// </summary>
        private static string? TryParseLocalPath(string? text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;

            var s = text.Trim();
            int nl = s.IndexOfAny(new[] { '\r', '\n' });
            if (nl >= 0) s = s[..nl];
            s = s.Trim().Trim('"').Trim();
            if (s.Length == 0) return null;

            try
            {
                // Normaler Pfad zuerst – damit Zeichen wie # oder % im Dateinamen
                // nicht als URI-Bestandteile fehlinterpretiert werden.
                if (Path.IsPathFullyQualified(s))
                    return File.Exists(s) ? s : null;

                if (Uri.TryCreate(s, UriKind.Absolute, out var u) && u.IsFile)
                {
                    var p = u.LocalPath;
                    return File.Exists(p) ? p : null;
                }
            }
            catch { /* ungültiger Pfad → kein lokales Medium */ }

            return null;
        }

        /// <summary>
        /// Zeigt eine lokale Bild- oder Videodatei im Medienfenster an und merkt sie
        /// als „letztes Bild". Bilder werden dafür – wie Grok-Bilder – als Kopie in
        /// den Zwischenspeicher gelegt; die Originaldatei wird nie als LastImagePath
        /// eingetragen, damit „Browserdaten löschen" sie nicht anfasst.
        /// </summary>
        private void OpenLocalMedia(string path)
        {
            var ext = Path.GetExtension(path);

            if (LocalVideoExtensions.Contains(ext))
            {
                ShowVideoViewer(path);   // merkt sich den Pfad selbst als letztes Medium
                return;
            }

            if (!LocalImageExtensions.Contains(ext))
            {
                MessageBox.Show(this,
                    string.Format(Properties.Resources.UnsupportedMediaFormat,
                                  string.IsNullOrEmpty(ext) ? Path.GetFileName(path) : ext),
                    Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Datei-Stream sofort wieder freigeben (siehe OpenLastImage).
                Bitmap bmp;
                using (var fs = File.OpenRead(path))
                using (var decoded = new Bitmap(fs))
                    bmp = new Bitmap(decoded);

                ShowImageViewer(bmp, path);
                WebViewManager.SaveAsLastImage(bmp, path);
            }
            catch (Exception ex)
            {
                Log(ex, "OpenLocalMedia fehlgeschlagen");
                MessageBox.Show(this, Properties.Resources.ImageLoadFailed,
                    Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static bool IsHttpUrl(string? s) =>
            Uri.TryCreate(s, UriKind.Absolute, out var u) &&
            (u.Scheme == Uri.UriSchemeHttp || u.Scheme == Uri.UriSchemeHttps);

    }
}
