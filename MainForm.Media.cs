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
        /// Zeigt ein Bild im Viewer-Fenster und ordnet es links auf voller
        /// Bildschirmhöhe an; Wrok füllt den Rest rechts daneben. Es gibt immer
        /// höchstens ein Medienfenster – ein neues Medium ersetzt das bisherige.
        /// Wrok bleibt nach dem Schließen in der Anordnung stehen: Medien haben
        /// meist dieselbe Größe, das nächste öffnet dadurch ohne Sprung.
        /// </summary>
        internal void ShowImageViewer(Bitmap bmp, string sourceUrl)
        {
            CloseCurrentMedia();

            var viewer = new ImageViewerForm(bmp, sourceUrl);
            RegisterMediaViewer(viewer);

            ArrangeSideBySide(viewer, viewer.WidthForHeight(WorkArea.Height));
            viewer.Show();
        }

        /// <summary>
        /// Zeigt ein Video im Endlos-Loop. Die echten Abmessungen sind erst nach
        /// dem Laden der Metadaten bekannt – das Fenster wird dann nachjustiert.
        /// </summary>
        internal void ShowVideoViewer(string url)
        {
            if (_webViewManager == null) return;

            CloseCurrentMedia();

            var viewer = new VideoViewerForm(url, _webViewManager.CoreEnvironment);
            RegisterMediaViewer(viewer);

            // Vorläufig hochkant anordnen; sobald die Metadaten da sind, korrigieren.
            ArrangeSideBySide(viewer, (int)(WorkArea.Height * 9.0 / 16.0));

            viewer.VideoSizeKnown += size =>
            {
                if (viewer.IsDisposed) return;
                try
                {
                    viewer.BeginInvoke((MethodInvoker)(() =>
                    {
                        if (!viewer.IsDisposed)
                            ArrangeSideBySide(viewer, viewer.WidthForHeight(WorkArea.Height, size));
                    }));
                }
                catch (Exception ex) { Log(ex, "Video-Anordnung nach Metadaten fehlgeschlagen"); }
            };

            viewer.Show();

            Properties.Settings.Default.LastImageUrl  = url;
            Properties.Settings.Default.LastImagePath = string.Empty;   // Video: nur URL
            Properties.Settings.Default.Save();
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
            viewer.FormClosed += (_, _) =>
            {
                if (ReferenceEquals(_mediaViewer, viewer)) _mediaViewer = null;
            };
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
