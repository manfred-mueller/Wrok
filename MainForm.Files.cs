namespace Wrok
{
    public partial class MainForm
    {
        // ------------------------------------------------------------------
        // Dateiverwaltung (grok.com/files)
        // ------------------------------------------------------------------

        private FilesViewerForm? _filesViewer;

        /// <summary>
        /// Öffnet Groks eigene Dateiverwaltung in einem separaten Fenster (siehe
        /// FilesViewerForm). Ist bereits eines offen, wird es nur nach vorne geholt
        /// statt ein zweites zu erzeugen.
        /// </summary>
        private void ShowFilesViewer()
        {
            if (_filesViewer is { IsDisposed: false } existing)
            {
                existing.Activate();
                if (existing.WindowState == FormWindowState.Minimized)
                    existing.WindowState = FormWindowState.Normal;
                return;
            }

            var viewer = new FilesViewerForm(_webViewManager?.CoreEnvironment);
            _filesViewer = viewer;
            viewer.FormClosed += (_, _) =>
            {
                if (ReferenceEquals(_filesViewer, viewer)) _filesViewer = null;
            };
            viewer.Show();
        }
    }
}
