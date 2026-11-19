using System.Diagnostics;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Security.Principal;

namespace Wrok
{
    /// <summary>
    /// Geschützter Medienordner %LOCALAPPDATA%\Wrok\media für Bilder und Videos.
    ///
    /// Zwei unabhängige Schutzstufen:
    /// - Sperre (gegen Neugierige): Eine Verweigerungsregel „Ordner auflisten" für
    ///   den aktuellen Benutzer, nur auf dem Ordner selbst (keine Vererbung). Das
    ///   geht unabhängig von der Ordnergröße sofort. Dateien bleiben über ihren
    ///   vollständigen Pfad erreichbar – Wrok kann also auch im gesperrten Zustand
    ///   hineinschreiben und daraus abspielen. Der Besitzer darf die Zugriffsliste
    ///   immer ändern, deshalb braucht das Entsperren keine Adminrechte.
    /// - EFS (fester Schutz): Windows-Dateiverschlüsselung, an das Windows-Konto
    ///   gebunden. Ein markierter Ordner verschlüsselt neu angelegte Dateien
    ///   automatisch.
    ///
    /// Der Unterordner „view" enthält Zwischendateien des Medienfensters (angezeigte
    /// Bilder, letztes Bild); er wird von Wrok selbst aufgeräumt.
    /// </summary>
    internal static class MediaProtection
    {
        public static string WrokDir => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Wrok");

        public static string MediaDir       => Path.Combine(WrokDir, "media");
        public static string ViewDir        => Path.Combine(MediaDir, "view");
        public static string WebViewDataDir => Path.Combine(WrokDir, "WebView2Data");

        // Frühere Ablageorte (vor dem gemeinsamen Medienordner).
        private static string LegacyImagesDir => Path.Combine(WrokDir, "images");
        private static string LegacyViewDir   => Path.Combine(WrokDir, "view");

        /// <summary>Legt den Medienordner an (falls nötig) und liefert seinen Pfad.</summary>
        public static string EnsureMediaDir()
        {
            Directory.CreateDirectory(MediaDir);
            return MediaDir;
        }

        /// <summary>Legt den Zwischenordner des Medienfensters an und liefert seinen Pfad.</summary>
        public static string EnsureViewDir()
        {
            Directory.CreateDirectory(ViewDir);
            return ViewDir;
        }

        // ------------------------------------------------------------------
        // Sperre
        // ------------------------------------------------------------------

        /// <summary>
        /// Setzt oder entfernt die Verweigerungsregel auf dem Medienordner.
        /// Liefert false, wenn das nicht geklappt hat (Details im Log).
        /// </summary>
        public static bool SetLocked(bool locked)
        {
            try
            {
                var dir  = new DirectoryInfo(EnsureMediaDir());
                var user = WindowsIdentity.GetCurrent().User;
                if (user == null) return false;

                var rule = new FileSystemAccessRule(user, FileSystemRights.ListDirectory,
                    InheritanceFlags.None, PropagationFlags.None, AccessControlType.Deny);

                var acl = dir.GetAccessControl(AccessControlSections.Access);
                acl.RemoveAccessRuleSpecific(rule);
                if (locked) acl.AddAccessRule(rule);
                dir.SetAccessControl(acl);
                return true;
            }
            catch (Exception ex)
            {
                Log(ex, locked ? "Sperren fehlgeschlagen" : "Entsperren fehlgeschlagen");
                return false;
            }
        }

        // ------------------------------------------------------------------
        // EFS
        // ------------------------------------------------------------------

        /// <summary>Ergebnis einer EFS-Umstellung.</summary>
        internal record EfsResult(int Done, int Failed, string? RootError);

        /// <summary>Ist der Ordner selbst als verschlüsselt markiert?</summary>
        public static bool IsEncrypted(string dir)
        {
            try { return Directory.Exists(dir) && new DirectoryInfo(dir).Attributes.HasFlag(FileAttributes.Encrypted); }
            catch { return false; }
        }

        /// <summary>
        /// Verschlüsselt bzw. entschlüsselt einen Ordner samt Inhalt. Der Ordner und
        /// alle Unterordner werden markiert, damit neue Dateien automatisch folgen.
        /// Bereits passende Dateien werden übersprungen; geöffnete Dateien schlagen
        /// fehl und werden nur gezählt. Schlägt schon der Ordner selbst fehl (z. B.
        /// Windows Home oder kein NTFS), wird abgebrochen und der Grund geliefert.
        /// </summary>
        public static EfsResult ApplyEfs(string root, bool encrypt)
        {
            Directory.CreateDirectory(root);

            try { Apply(root, encrypt); }
            catch (Exception ex)
            {
                Log(ex, $"EFS für '{root}' nicht möglich");
                return new EfsResult(0, 0, ex.Message);
            }

            int done = 1, failed = 0;
            var rootInfo = new DirectoryInfo(root);

            // Nicht lesbare Unterordner überspringen statt die ganze Aktion abzubrechen.
            var options = new EnumerationOptions
            {
                RecurseSubdirectories = true,
                IgnoreInaccessible    = true,
                AttributesToSkip      = 0
            };

            foreach (var d in Safe(() => rootInfo.EnumerateDirectories("*", options)))
                Count(d, d.Attributes);

            foreach (var f in Safe(() => rootInfo.EnumerateFiles("*", options)))
                Count(f, f.Attributes);

            return new EfsResult(done, failed, null);

            void Count(FileSystemInfo item, FileAttributes attrs)
            {
                if (attrs.HasFlag(FileAttributes.Encrypted) == encrypt) return;
                try { Apply(item.FullName, encrypt); done++; }
                catch (Exception ex) { failed++; Log(ex, $"EFS für '{item.FullName}' fehlgeschlagen"); }
            }
        }

        private static void Apply(string path, bool encrypt)
        {
            if (encrypt) File.Encrypt(path);
            else         File.Decrypt(path);
        }

        private static IEnumerable<T> Safe<T>(Func<IEnumerable<T>> source)
        {
            try { return source().ToList(); }
            catch (Exception ex) { Log(ex, "Aufzählen fehlgeschlagen"); return Enumerable.Empty<T>(); }
        }

        /// <summary>Startet den Windows-Assistenten zum Sichern des EFS-Schlüssels.</summary>
        public static bool StartKeyBackup()
        {
            try
            {
                Process.Start(new ProcessStartInfo("rekeywiz.exe") { UseShellExecute = true });
                return true;
            }
            catch (Exception ex)
            {
                Log(ex, "rekeywiz konnte nicht gestartet werden");
                return false;
            }
        }

        // ------------------------------------------------------------------
        // Dateien im Medienordner
        // ------------------------------------------------------------------

        /// <summary>
        /// Liegt die Datei im Medienordner? Der Zwischenordner „view" zählt nicht,
        /// weil Wrok ihn selbst aufräumt.
        /// </summary>
        public static bool IsInMediaDir(string path)
        {
            try
            {
                var full = Path.GetFullPath(path);
                return full.StartsWith(MediaDir + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                    && !full.StartsWith(ViewDir + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        /// <summary>
        /// Kopiert eine Datei in den Medienordner und liefert den neuen Pfad. Bei
        /// Namensgleichheit wird „Name (2).ext" usw. verwendet; nichts wird überschrieben.
        /// </summary>
        public static string CopyIntoMediaDir(string source)
        {
            var target = UniquePath(EnsureMediaDir(), Path.GetFileName(source));
            File.Copy(source, target, overwrite: false);
            return target;
        }

        /// <summary>
        /// Sucht im Medienordner (ohne Unterordner) eine Datei mit identischem Inhalt.
        /// Erst wird die Größe verglichen, nur bei Gleichheit der SHA-256-Hash.
        /// Muss bei entsperrtem Ordner laufen, weil der Ordner aufgelistet wird.
        /// </summary>
        public static string? FindIdenticalFile(string source)
        {
            var src = new FileInfo(source);
            byte[]? srcHash = null;

            foreach (var f in new DirectoryInfo(EnsureMediaDir()).EnumerateFiles())
            {
                if (f.Length != src.Length) continue;
                if (string.Equals(f.FullName, src.FullName, StringComparison.OrdinalIgnoreCase)) return f.FullName;

                try
                {
                    srcHash ??= HashFile(src.FullName);
                    if (HashFile(f.FullName).AsSpan().SequenceEqual(srcHash)) return f.FullName;
                }
                catch (Exception ex) { Log(ex, $"Vergleich mit '{f.FullName}' fehlgeschlagen"); }
            }
            return null;
        }

        private static byte[] HashFile(string path)
        {
            using var stream = File.OpenRead(path);
            return SHA256.HashData(stream);
        }

        /// <summary>
        /// Kopiert alle Medien (ohne den Zwischenordner) unverschlüsselt in einen
        /// Zielordner. Die Kopie läuft über Streams, damit EFS nicht mitwandert.
        /// Liefert (kopiert, fehlgeschlagen).
        /// </summary>
        public static (int Copied, int Failed) ExportTo(string targetDir)
        {
            int copied = 0, failed = 0;
            foreach (var src in Directory.EnumerateFiles(EnsureMediaDir()))
            {
                try
                {
                    var dest = UniquePath(targetDir, Path.GetFileName(src));
                    using (var input  = File.OpenRead(src))
                    using (var output = new FileStream(dest, FileMode.CreateNew, FileAccess.Write))
                        input.CopyTo(output);
                    File.SetLastWriteTime(dest, File.GetLastWriteTime(src));
                    copied++;
                }
                catch (Exception ex)
                {
                    failed++;
                    Log(ex, $"Export von '{src}' fehlgeschlagen");
                }
            }
            return (copied, failed);
        }

        private static string UniquePath(string dir, string fileName)
        {
            var path = Path.Combine(dir, fileName);
            var name = Path.GetFileNameWithoutExtension(fileName);
            var ext  = Path.GetExtension(fileName);
            for (int i = 2; File.Exists(path); i++)
                path = Path.Combine(dir, $"{name} ({i}){ext}");
            return path;
        }

        // ------------------------------------------------------------------
        // Umzug aus den früheren Ordnern
        // ------------------------------------------------------------------

        /// <summary>
        /// Zieht das letzte Bild aus %LOCALAPPDATA%\Wrok\images in den Zwischenordner
        /// des Medienordners um und räumt die alten Ordner images und view auf.
        /// </summary>
        public static void MigrateLegacyFolders()
        {
            try
            {
                var s = Properties.Settings.Default;
                var lastPath = s.LastImagePath;
                if (!string.IsNullOrWhiteSpace(lastPath) && File.Exists(lastPath) &&
                    string.Equals(Path.GetDirectoryName(Path.GetFullPath(lastPath)), LegacyImagesDir, StringComparison.OrdinalIgnoreCase))
                {
                    var target = Path.Combine(EnsureViewDir(), Path.GetFileName(lastPath));
                    File.Copy(lastPath, target, overwrite: true);
                    s.LastImagePath = target;
                    s.Save();
                }
            }
            catch (Exception ex) { Log(ex, "Letztes Bild umziehen fehlgeschlagen"); }

            foreach (var dir in new[] { LegacyImagesDir, LegacyViewDir })
            {
                // Zeigt das letzte Bild noch in einen alten Ordner (Umzug fehlgeschlagen),
                // diesen Ordner lieber stehen lassen.
                if (StillReferenced(dir)) continue;

                try { if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true); }
                catch (Exception ex) { Log(ex, $"Alten Ordner '{dir}' entfernen fehlgeschlagen"); }
            }
        }

        private static bool StillReferenced(string dir)
        {
            try
            {
                var p = Properties.Settings.Default.LastImagePath;
                return !string.IsNullOrWhiteSpace(p) &&
                       string.Equals(Path.GetDirectoryName(Path.GetFullPath(p)), dir, StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        private static void Log(Exception ex, string message) =>
            Trace.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [MediaProtection] {message}: {ex}");
    }
}
