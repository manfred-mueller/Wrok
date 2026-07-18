using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace Wrok
{
    /// <summary>
    /// Zeigt ein aus dem WebView abgefangenes Bild in einem eigenständigen Fenster an.
    /// Bewusst ohne Toolbar – geschlossen wird über das X der Titelleiste.
    ///
    /// Tastenkürzel:
    ///   Esc          schließen
    ///   + / -        zoomen (auch Mausrad)
    ///   Strg+F       einpassen
    ///   Strg+S       speichern
    ///   Strg+P       anheften (immer im Vordergrund)
    /// Ziehen mit der linken Maustaste verschiebt das Bild.
    /// </summary>
    internal sealed class ImageViewerForm : Form
    {
        // ------------------------------------------------------------------
        // P/Invoke
        // ------------------------------------------------------------------
        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

        // ------------------------------------------------------------------
        // Felder
        // ------------------------------------------------------------------
        private readonly Bitmap _image;
        private readonly string _sourceUrl;

        private float  _zoom = 1f;
        private PointF _pan  = PointF.Empty;
        private Point  _dragStart;
        private bool   _dragging;

        private const float ZoomStep = 0.15f;
        private const float ZoomMin  = 0.05f;
        private const float ZoomMax  = 16f;

        private readonly PictureBox _canvas;

        /// <summary>
        /// PictureBox, die den Fokus annehmen kann. Nötig, weil WM_MOUSEWHEEL nur an
        /// das fokussierte Control geht – eine normale PictureBox ist nicht
        /// fokussierbar, wodurch der Mausrad-Zoom sonst gar nicht ankommt.
        /// </summary>
        private sealed class FocusablePictureBox : PictureBox
        {
            public FocusablePictureBox()
            {
                SetStyle(ControlStyles.Selectable, true);
                TabStop = true;
            }
        }

        // ------------------------------------------------------------------
        // Konstruktor
        // ------------------------------------------------------------------
        public ImageViewerForm(Bitmap image, string sourceUrl)
        {
            _image     = image ?? throw new ArgumentNullException(nameof(image));
            _sourceUrl = sourceUrl ?? string.Empty;

            // Eigenständiges Fenster (kein Owner): dadurch alt-tabbar und hinter das
            // Hauptfenster legbar, sodass währenddessen weiter getippt werden kann.
            Text            = $"Bild-Viewer – Wrok   ({_image.Width}×{_image.Height})";
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition   = FormStartPosition.CenterScreen;
            BackColor       = Color.FromArgb(28, 28, 28);
            KeyPreview      = true;
            ShowInTaskbar   = true;
            MinimumSize     = new Size(240, 180);
            ClientSize      = InitialClientSize();

            TryApplyDarkTitleBar();

            _canvas = new FocusablePictureBox
            {
                Dock      = DockStyle.Fill,
                BackColor = Color.FromArgb(28, 28, 28),
            };
            _canvas.Paint      += OnCanvasPaint;
            _canvas.MouseWheel += OnMouseWheel;
            _canvas.MouseDown  += OnMouseDown;
            _canvas.MouseMove  += OnMouseMove;
            _canvas.MouseUp    += OnMouseUp;
            _canvas.MouseEnter += (_, _) => { if (!_canvas.Focused) _canvas.Focus(); };

            Controls.Add(_canvas);

            KeyDown += OnKeyDown;
            Shown   += (_, _) => { FitToWindow(); _canvas.Invalidate(); };
            // Fenstergröße skaliert das Bild mit.
            Resize  += (_, _) => { FitToWindow(); _canvas.Invalidate(); };
        }

        /// <summary>
        /// Gesamtbreite (inkl. Rahmen), die das Fenster bei der angegebenen
        /// Gesamthöhe braucht, damit das Bild seitenverhältnisgerecht hineinpasst.
        /// Wird für die Nebeneinander-Anordnung benötigt.
        /// </summary>
        public int WidthForHeight(int totalHeight)
        {
            // Rahmen und Titelleiste herausrechnen (Handle existiert bereits,
            // weil der Konstruktor die dunkle Titelleiste gesetzt hat).
            int extraH = Height - ClientSize.Height;
            int extraW = Width  - ClientSize.Width;

            int clientH = Math.Max(1, totalHeight - extraH);
            int clientW = (int)Math.Round(_image.Width * (clientH / (double)_image.Height));

            return clientW + extraW;
        }

        /// <summary>
        /// Fenster so groß wie das Bild – aber höchstens 90 % des Arbeitsbereichs,
        /// damit große Bilder nicht über den Bildschirm hinauslaufen.
        /// </summary>
        private Size InitialClientSize()
        {
            try
            {
                var wa = Screen.FromPoint(Cursor.Position).WorkingArea;
                int maxW = (int)(wa.Width  * 0.9);
                int maxH = (int)(wa.Height * 0.9);

                float scale = Math.Min(1f, Math.Min((float)maxW / _image.Width,
                                                    (float)maxH / _image.Height));

                return new Size(Math.Max(240, (int)(_image.Width  * scale)),
                                Math.Max(180, (int)(_image.Height * scale)));
            }
            catch
            {
                return new Size(960, 720);
            }
        }

        // ------------------------------------------------------------------
        // Anheften (Always-on-top)
        // ------------------------------------------------------------------
        private void TogglePinned()
        {
            TopMost = !TopMost;
            Text = $"Bild-Viewer – Wrok   ({_image.Width}×{_image.Height})" +
                   (TopMost ? "   📌" : string.Empty);
        }

        // ------------------------------------------------------------------
        // Zoom / Pan
        // ------------------------------------------------------------------
        private void FitToWindow()
        {
            int availW = _canvas.Width;
            int availH = _canvas.Height;
            if (availW <= 0 || availH <= 0) return;

            float sx = (float)availW / _image.Width;
            float sy = (float)availH / _image.Height;
            _zoom = Math.Clamp(Math.Min(sx, sy), ZoomMin, ZoomMax);
            _pan  = PointF.Empty;
        }

        private void ApplyZoom(float factor, Point pivot)
        {
            float newZoom = Math.Clamp(_zoom * factor, ZoomMin, ZoomMax);
            // Zoom auf Mauszeiger-Position zentrieren
            float cx = _canvas.Width  / 2f + _pan.X;
            float cy = _canvas.Height / 2f + _pan.Y;
            float dx = pivot.X - cx;
            float dy = pivot.Y - cy;
            _pan.X += dx * (1f - newZoom / _zoom);
            _pan.Y += dy * (1f - newZoom / _zoom);
            _zoom = newZoom;
        }

        // ------------------------------------------------------------------
        // Paint
        // ------------------------------------------------------------------
        private void OnCanvasPaint(object? sender, PaintEventArgs e)
        {
            int w = (int)(_image.Width  * _zoom);
            int h = (int)(_image.Height * _zoom);
            int x = (_canvas.Width  - w) / 2 + (int)_pan.X;
            int y = (_canvas.Height - h) / 2 + (int)_pan.Y;

            e.Graphics.InterpolationMode =
                System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            e.Graphics.PixelOffsetMode =
                System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            e.Graphics.DrawImage(_image, x, y, w, h);

            // Zoom-Anzeige
            string info = $"{(int)(_zoom * 100)} %";
            using var font  = new Font("Segoe UI", 8.5f);
            using var brush = new SolidBrush(Color.FromArgb(110, 255, 255, 255));
            e.Graphics.DrawString(info, font, brush, 8, _canvas.Height - 22);
        }

        // ------------------------------------------------------------------
        // Maus
        // ------------------------------------------------------------------
        private void OnMouseWheel(object? sender, MouseEventArgs e)
        {
            float factor = e.Delta > 0 ? 1f + ZoomStep : 1f - ZoomStep;
            ApplyZoom(factor, e.Location);
            _canvas.Invalidate();
        }

        private void OnMouseDown(object? sender, MouseEventArgs e)
        {
            if (!_canvas.Focused) _canvas.Focus();
            if (e.Button != MouseButtons.Left) return;
            _dragging  = true;
            _dragStart = e.Location;
            _canvas.Cursor = Cursors.SizeAll;
        }

        private void OnMouseMove(object? sender, MouseEventArgs e)
        {
            if (!_dragging) return;
            _pan.X += e.X - _dragStart.X;
            _pan.Y += e.Y - _dragStart.Y;
            _dragStart = e.Location;
            _canvas.Invalidate();
        }

        private void OnMouseUp(object? sender, MouseEventArgs e)
        {
            _dragging      = false;
            _canvas.Cursor = Cursors.Default;
        }

        // ------------------------------------------------------------------
        // Tastatur
        // ------------------------------------------------------------------
        private void OnKeyDown(object? sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Escape:
                    Close();
                    break;

                case Keys.Add:
                case Keys.Oemplus:
                    ApplyZoom(1f + ZoomStep, new Point(_canvas.Width / 2, _canvas.Height / 2));
                    _canvas.Invalidate();
                    break;

                case Keys.Subtract:
                case Keys.OemMinus:
                    ApplyZoom(1f - ZoomStep, new Point(_canvas.Width / 2, _canvas.Height / 2));
                    _canvas.Invalidate();
                    break;

                case Keys.F when e.Control:
                    FitToWindow();
                    _canvas.Invalidate();
                    break;

                case Keys.S when e.Control:
                    SaveImage();
                    break;

                case Keys.P when e.Control:
                    TogglePinned();
                    break;
            }
        }

        // ------------------------------------------------------------------
        // Speichern
        // ------------------------------------------------------------------
        private void SaveImage()
        {
            using var dlg = new SaveFileDialog
            {
                Title           = "Bild speichern",
                Filter          = "PNG-Bild|*.png|JPEG-Bild|*.jpg|Alle Dateien|*.*",
                FileName        = SuggestFilename(_sourceUrl),
                OverwritePrompt = true,
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            try
            {
                var fmt = dlg.FilterIndex == 2 ? ImageFormat.Jpeg : ImageFormat.Png;
                _image.Save(dlg.FileName, fmt);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Speichern fehlgeschlagen:\n{ex.Message}",
                    "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ------------------------------------------------------------------
        // Hilfsmethoden
        // ------------------------------------------------------------------
        private void TryApplyDarkTitleBar()
        {
            try
            {
                int val = 1;
                if (DwmSetWindowAttribute(Handle, 20, ref val, sizeof(int)) != 0)
                    DwmSetWindowAttribute(Handle, 19, ref val, sizeof(int));
            }
            catch { }
        }

        private static string SuggestFilename(string url)
        {
            try
            {
                var name = Path.GetFileName(new Uri(url).AbsolutePath);
                // Grok-CDN liefert Pfade ohne Dateiendung (z. B. .../asset)
                if (!string.IsNullOrEmpty(name))
                    return Path.HasExtension(name) ? name : name + ".png";
            }
            catch { }
            return "bild.png";
        }

        // ------------------------------------------------------------------
        // Dispose
        // ------------------------------------------------------------------
        protected override void Dispose(bool disposing)
        {
            if (disposing) _image.Dispose();
            base.Dispose(disposing);
        }
    }
}
