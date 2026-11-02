using System.Diagnostics;
using System.Text.Json;
using System.Windows.Forms;

namespace Wrok
{
    /// <summary>Eine Video-Schnellwahl: Anzeigename und Adresse (URL oder lokaler Pfad).</summary>
    internal record VideoSlot(string Name, string Url);

    /// <summary>
    /// Speichert die Video-Slots in %LOCALAPPDATA%\Wrok\videos.json. Die Slots 1 bis 9
    /// sind per Strg+1 bis Strg+9 direkt erreichbar.
    /// </summary>
    internal static class VideoSlotStore
    {
        public const int SlotCount = 9;

        private static readonly object _lock = new();
        private static List<VideoSlot>? _cache;

        private static string FilePath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Wrok", "videos.json");

        public static List<VideoSlot> Get()
        {
            lock (_lock)
            {
                if (_cache == null)
                {
                    List<VideoSlot>? list = null;
                    try
                    {
                        if (File.Exists(FilePath))
                            list = JsonSerializer.Deserialize<List<VideoSlot>>(File.ReadAllText(FilePath));
                    }
                    catch (Exception ex)
                    {
                        Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [VideoSlotStore] Laden fehlgeschlagen: {ex}");
                    }
                    _cache = Normalize(list);
                }
                return _cache.ToList();
            }
        }

        public static void Set(int index, VideoSlot slot)
        {
            lock (_lock)
            {
                var list = Get();
                list[index] = slot;
                _cache = list;
                Save(list);
            }
        }

        private const int ExportFormatVersion = 1;

        internal record VideoSlotExport(int WrokVideoSlots, List<VideoSlot> Slots);

        /// <summary>Schreibt alle Slots in eine eigene JSON-Datei.</summary>
        public static bool ExportTo(string path)
        {
            try
            {
                var doc = new VideoSlotExport(ExportFormatVersion, Get());
                File.WriteAllText(path, JsonSerializer.Serialize(doc, new JsonSerializerOptions { WriteIndented = true }));
                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [VideoSlotStore] Export fehlgeschlagen: {ex}");
                return false;
            }
        }

        /// <summary>Ersetzt alle Slots durch den Inhalt einer Exportdatei.</summary>
        public static bool ImportFrom(string path)
        {
            try
            {
                var doc = JsonSerializer.Deserialize<VideoSlotExport>(File.ReadAllText(path));
                if (doc == null || doc.WrokVideoSlots != ExportFormatVersion || doc.Slots == null) return false;

                lock (_lock)
                {
                    _cache = Normalize(doc.Slots);
                    Save(_cache);
                }
                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [VideoSlotStore] Import fehlgeschlagen: {ex}");
                return false;
            }
        }

        private static List<VideoSlot> Normalize(List<VideoSlot>? list)
        {
            list ??= new List<VideoSlot>();
            while (list.Count < SlotCount)
                list.Add(new VideoSlot(string.Empty, string.Empty));
            return list.Take(SlotCount).ToList();
        }

        private static void Save(List<VideoSlot> list)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
                File.WriteAllText(FilePath, JsonSerializer.Serialize(list));
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [VideoSlotStore] Speichern fehlgeschlagen: {ex}");
            }
        }
    }

    public partial class MainForm
    {
        // ------------------------------------------------------------------
        // Videos-Menü: neun Schnellwahl-Slots für den Video-Viewer
        // ------------------------------------------------------------------

        private ToolStripMenuItem? videosMenu;

        private void InitializeVideosMenu()
        {
            videosMenu ??= new ToolStripMenuItem(Properties.Resources.VideosMenu);
            RefreshVideosMenu();
        }

        private void RefreshVideosMenu()
        {
            if (videosMenu == null) return;
            videosMenu.DropDownItems.Clear();
            videosMenu.DropDown.ShowItemToolTips = true;

            var slots = VideoSlotStore.Get();
            for (int i = 0; i < slots.Count; i++)
            {
                int idx = i;
                var slot = slots[i];

                string title = !string.IsNullOrWhiteSpace(slot.Name) ? slot.Name
                             : !string.IsNullOrWhiteSpace(slot.Url) ? slot.Url
                             : Properties.Resources.VideoSlotEmpty;

                var item = new ToolStripMenuItem(string.Format(Properties.Resources.VideoSlotMenuItem, i + 1, title)) { Tag = idx };
                item.ToolTipText = string.Format(Properties.Resources.VideoSlotTooltip, i + 1) + "\n" + slot.Url;

                item.MouseDown += (sender, me) =>
                {
                    try
                    {
                        if (me.Button == MouseButtons.Left) PlayVideoSlot(idx);
                        else if (me.Button == MouseButtons.Right) EditVideoSlot(idx);
                    }
                    catch (Exception ex) { Log(ex, "Video-Slot-Klick fehlgeschlagen"); }
                };
                videosMenu.DropDownItems.Add(item);
            }

            videosMenu.DropDownItems.Add(new ToolStripSeparator());
            var videoIoMenu = new ToolStripMenuItem(Properties.Resources.VideosImportExport);
            videoIoMenu.DropDownItems.Add(Properties.Resources.VideosExport, null, (s, e) => ExportVideoSlots());
            videoIoMenu.DropDownItems.Add(Properties.Resources.VideosImport, null, (s, e) => ImportVideoSlots());
            videosMenu.DropDownItems.Add(videoIoMenu);
        }

        private void ExportVideoSlots()
        {
            using var dlg = new SaveFileDialog
            {
                Title      = Properties.Resources.VideosExport,
                Filter     = Properties.Resources.VideosFileFilter,
                FileName   = "wrok-videos.json",
                DefaultExt = "json"
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            if (!VideoSlotStore.ExportTo(dlg.FileName))
                MessageBox.Show(this, Properties.Resources.VideosExportFailed, Properties.Resources.VideosMenu,
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void ImportVideoSlots()
        {
            using var dlg = new OpenFileDialog
            {
                Title      = Properties.Resources.VideosImport,
                Filter     = Properties.Resources.VideosFileFilter,
                DefaultExt = "json"
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            if (MessageBox.Show(this, Properties.Resources.VideosImportConfirm, Properties.Resources.VideosMenu,
                                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;

            if (!VideoSlotStore.ImportFrom(dlg.FileName))
            {
                MessageBox.Show(this, Properties.Resources.VideosImportInvalid, Properties.Resources.VideosMenu,
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MessageBox.Show(this, Properties.Resources.VideosImportDone, Properties.Resources.VideosMenu,
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
            RefreshVideosMenu();
        }

        /// <summary>Öffnet den Video-Viewer für einen Slot. Leere Slots tun nichts.</summary>
        internal void PlayVideoSlot(int index)
        {
            var slots = VideoSlotStore.Get();
            if (index < 0 || index >= slots.Count) return;

            string url = slots[index].Url;
            if (string.IsNullOrWhiteSpace(url)) return;

            ShowVideoViewer(url);
        }

        private void EditVideoSlot(int index)
        {
            try
            {
                var current = VideoSlotStore.Get()[index];
                string name = current.Name;
                string url  = current.Url;

                if (!Dialogs.ShowEditMacro(this, string.Format(Properties.Resources.VideoSlotEditTitle, index + 1), ref name, ref url)) return;

                url = url.Trim().Trim('"');
                VideoSlotStore.Set(index, new VideoSlot(name.Trim(), url));
                RefreshVideosMenu();
            }
            catch (Exception ex)
            {
                Log(ex, "EditVideoSlot fehlgeschlagen");
            }
        }
    }
}
