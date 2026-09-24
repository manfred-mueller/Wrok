# Changelog

## [Unreleased]

### Changes
- **Makro-Hotkeys neu: Strg+F1/F2/F3 statt Strg+Ziffer, Nummernblock-Modus entfernt.** Der Nummernblock-Modus war die unzuverlässigste Stelle im Makro-System (erzwang NumLock, funktionierte auf Notebooks ohne Nummernblock gar nicht). Ersetzt durch: Strg+F1, Strg+F2 bzw. Strg+F3 öffnen ein Auswahl-Popup mit den zehn Makrotiteln des jeweiligen Profils (fest zugeordnet, unabhängig vom im Tray-Menü angezeigten Profil); eine folgende Ziffer (1-9, 0 für das zehnte Makro) sendet es, Escape oder Fokusverlust bricht ab. Funktioniert identisch auf jeder Tastatur. Die Einstellung „Nummernblock-Modus" sowie die zugehörigen Menüpunkte entfallen ersatzlos. Tray-Menü (Bearbeiten per Rechtsklick, Senden per Linksklick, Profilwechsel für die Bearbeitungsansicht) bleibt unverändert.

### Fixes
- **`NativeInput.TypeText` blockierte kurz den UI-Thread** beim Senden längerer Makros über den Fallback-Tippweg (`Thread.Sleep(2)` pro Zeichen, synchron aufgerufen). Läuft jetzt in `Task.Run`.

### Refactoring
- Vier identische `try { this.CenterToScreen(); } catch { }`-Stellen in `MainForm.cs` zu einer Hilfsmethode `TryCenterToScreen()` zusammengefasst.
- `MainForm.cs` (zuvor 1798 Zeilen) nach Zuständigkeit in `partial class`-Dateien aufgeteilt: `MainForm.Window.cs`, `MainForm.Tray.cs`, `MainForm.Macros.cs`, `MainForm.Settings.cs`, `MainForm.Media.cs`.
- Grok-Kontowechsel-Feature vollständig entfernt (Menü, Sign-in-Automation, `GrokAccountManager`, zugehörige Einstellungen/Ressourcen) – kein wahrgenommener Zusatznutzen gegenüber manueller Anmeldung. Die Passwort-Autosave-Funktion bleibt unverändert bestehen.

## [1.5.1] - 2026-09-24

### Fixes
- **Offline-Meldung zeigte Umlaute als Steuerzeichen statt lesbarer Zeichen.** In `WebViewManager.cs` war Text an mehreren Stellen doppelt falsch kodiert (UTF-8, versehentlich als Latin-1 gelesen und erneut als UTF-8 gespeichert) – u. a. in der Meldung „Bitte überprüfen Sie Ihre Internetverbindung…". Betraf ausschließlich diese eine Datei; korrigiert.
- **Zweitinstanz-Aktivierung broadcastete die echte Windows-Systemnachricht `WM_SHOWWINDOW`** an alle laufenden Programme (`HWND_BROADCAST`) statt nur an Wrok selbst – andere Anwendungen mit eigener `WM_SHOWWINDOW`-Behandlung konnten das fälschlich als echten Sichtbarkeits-Wechsel interpretieren. Läuft jetzt über eine eigene, per `RegisterWindowMessage` registrierte private Nachricht.
- **Mögliches Wettrennen beim Start zwischen Proxy-Initialisierung und erster Seitennavigation.** `LoadUrlAsync` konnte parallel zur (ggf. proxy-konfigurierten) `CoreWebView2`-Umgebungserstellung einen zweiten, konkurrierenden `EnsureCoreWebView2Async`-Aufruf mit abweichenden Parametern auslösen. Wartet jetzt auf denselben Initialisierungs-Task.
- **Internetverbindungs-Prüfung nutzte eine reine HTTP-URL**, die in einem unsicheren Netz gefälscht werden könnte. Auf HTTPS umgestellt.
- **Fehlschlag der Chef-Taste-Registrierung (Strg+Leertaste) blieb bisher nur im Log sichtbar.** Zeigt jetzt einmal pro Programmlauf einen Warnhinweis im Infobereich.
- **Makro-Text konnte nach einem Minimieren/Wiederherstellen unbemerkt verloren gehen**, wenn der NativeInput-Fallback in der kurzen Lücke zwischen Fensterhandle-Zerstörung und -Neuerzeugung ausgelöst wurde. Wartet jetzt kurz auf die Wiederherstellung, statt sofort stillschweigend aufzugeben.
- **Hilfetext im Info-Dialog beschrieb noch das alte Strg+F1/F2/F3-Auswahl-Popup** statt des aktuellen F1–F10-Direktschemas. Korrigiert (DE + EN).

### Intern
- Bild-Klick-Erkennung in `WrokHelper.js` läuft nur noch im Top-Frame (nicht in eingebetteten iframes) und schließt interaktive Elemente (Buttons, Links, Formularfelder) von der Hintergrundbild-Heuristik aus.
- `SendTextAsync`/`TryClickSendButtonAsync` nutzen jetzt einen gemeinsamen Skript-Ausführung-Helfer statt dupliziertem Try/Catch/Log-Code.
- Log-Rotation begrenzt jetzt die Anzahl aufbewahrter Archiv-Logs (zuvor unbegrenztes Wachstum).
- Tote Resource-Einträge (`MacroPickerRow`, `MacroPickerEmptyRow`) aus dem abgelösten Makro-Popup-Ansatz entfernt.


## [1.3.2] - 2026-09-16

### Features
- **Neue Makro-Variable `{day}`:** voller Wochentagsname, fest auf Deutsch (z. B. „Montag"), unabhängig von der Systemsprache.

### Fixes
- **Strg+Ö öffnete nichts.** `Keys.Oem1` ist auf der deutschen Tastatur tatsächlich die Ü-Taste, nicht Ö (verifiziert über die vom Tastaturtreiber generierte Scancode-Tabelle). Die Registrierung lief deshalb fehlerfrei durch, reagierte aber nie auf das tatsächlich gedrückte Strg+Ö. Jetzt korrekt auf `Keys.Oem3` registriert.
- **Makrotext wurde manchmal nicht ins Grok-Chatfeld eingetragen**, u. a. bei Strg+Ziffer-Makros auf manchen Geräten. Ursache: `ConfigureAwait(false)` an mehreren Stellen im WebView-Zugriff ließ den Code nach einem `await` auf einem beliebigen Hintergrund-Thread weiterlaufen statt auf dem UI-Thread – `CoreWebView2` darf aber nur vom UI-Thread aus angesprochen werden. Entfernt; zusätzliches Logging im Fallback-Pfad ergänzt.
- **Cookie-Einwilligung musste bei jedem Kontowechsel neu bestätigt werden.** Der Kontowechsel löscht die Sitzungscookies für grok.com und accounts.x.ai, was auch das Cookie-Einwilligungs-Cookie mitlöschte. Wird jetzt automatisch weggeklickt („Alle ablehnen") – auf der Anmeldeseite und nach der Rückkehr zu grok.com nach erfolgreichem Login.
- **Automatischer Login klickte „Anmelden" mitunter zu früh.** Der Button ist während der laufenden Cloudflare-Prüfung nicht deaktiviert; es wird jetzt zusätzlich auf das von Cloudflare selbst befüllte Ergebnisfeld gewartet. Der Sign-in-Flow läuft zudem über von C# aus wiederholte Einzelschritte statt verschachtelter JavaScript-Timer, robuster gegenüber Seitennavigationen zwischen E-Mail- und Passwort-Schritt.
- **Fehler beim Öffnen eines Bildes aus der Zwischenablage (Strg+Ö) konnten spurlos verschwinden.** Ein fehlender `catch`-Block ließ Ausnahmen in der „fire-and-forget" aufgerufenen Methode ohne jede Rückmeldung verpuffen. Wird jetzt geloggt und gemeldet.


## [1.3.1] - 2026-07-31

### Features
- **Update-Hinweis.** Wrok sieht beim Start (höchstens einmal täglich) auf GitHub nach, ob eine neuere Version vorliegt, und zeigt dann einen Tray-Ballon sowie einen Menüeintrag, der die Release-Seite öffnet. Es wird nichts heruntergeladen oder ausgeführt – die Aktualisierung selbst läuft weiter über WinGet bzw. den signierten Installer. Abschaltbar unter *Einstellungen → App → Nach Updates suchen* (standardmäßig an); der Menüeintrag weist darauf hin, dass dabei GitHub kontaktiert wird.

### Fixes
- **Sauberes Beenden bei Abmeldung, Herunterfahren und Update.** Das Fenster-X minimiert wie gewohnt in den Infobereich, ein echtes Sitzungsende beendet Wrok dagegen jetzt wirklich. Zuvor konnte die laufende Instanz die Programmdatei festhalten und so ein stilles Upgrade über WinGet behindern. Erkannt wird das über `WM_QUERYENDSESSION`, das jedem Schließen vorausgeht – unabhängig davon, wie der Installer die Anwendung schließt.


## [1.3.0] - 2026-07-31

### Features
- **Optionaler Nummernblock-Modus – alle Profile gleichzeitig, einhändig.** Zusätzlich zu den gewohnten `Strg+1`…`Strg+0` des aktiven Profils lassen sich die ersten fünf Makros **jedes** Profils ohne Zusatztaste über den Nummernblock auslösen: Profil 1 auf `Num0`–`Num4`, Profil 2 auf `Num5`–`Num9`, Profil 3 auf `/ * - + ,`. So sind erstmals alle drei Profile gleichzeitig per Hotkey erreichbar, ohne umzuschalten. Der Modus ist unter *Einstellungen → App* umschaltbar und standardmäßig aus, da die Numpad-Tasten dann systemweit belegt sind. Die klassische Strg-Belegung bleibt unverändert und funktioniert auf jeder Tastatur – auch ohne Nummernblock.
- **Grok-Konten-Verwaltung.** Mehrere Grok-Konten lassen sich mit Bezeichnung und E-Mail hinterlegen (*Einstellungen → Grok-Konten*). Ein Klick wechselt: Wrok löscht die Sitzungs-Cookies von grok.com und accounts.x.ai, ruft die Anmeldeseite auf und trägt die E-Mail automatisch ein. **Kein Passwort wird gespeichert** – das bleibt beim Chromium-eigenen Passwort-Manager oder der manuellen Eingabe.
- **Proxy-Unterstützung.** Adresse, Ausnahmeliste und optionale Anmeldung mit Benutzername und Passwort (*Einstellungen → App*). Das Passwort wird mit der Windows-DPAPI verschlüsselt abgelegt – nie im Klartext – und nur an genau den konfigurierten Proxy-Host herausgegeben, nicht an andere Anmeldeabfragen. Eine Änderung greift nach einem Neustart, den Wrok auf Wunsch selbst auslöst.
- **Passwort-Speichern umschaltbar.** Der Chromium-Passwort-Manager lässt sich über einen Menüpunkt ein- und ausschalten.

### Fixes
- **Bild-aus-Zwischenablage-Kürzel geändert** von `Strg+Shift+P` auf `Strg+Ö`, um Konflikte mit gängigen Web-Editor-Kürzeln zu vermeiden.
- **Einstellungen überstehen ein Versionsupdate.** Beim ersten Start einer neuen Fassung werden die gespeicherten Einstellungen übernommen, statt auf Standardwerte zurückzufallen.

### Intern
- **Sauberer Neustart** (`Program.RestartApplication`): Der Einzelinstanz-Mutex wird zuerst freigegeben, dann die neue Instanz gestartet – sonst blockierte die alte den Start der neuen.
- Grok-Konten und Numpad-Zuordnung als eigene, schlanke Manager (`GrokAccountManager`), Persistenz als JSON in den Settings – dasselbe Muster wie bei den Makros.


## [1.2.0] - 2026-07-19

### Features
- **„Daten löschen" statt „Cache löschen".** Der Eintrag unter *Werkzeuge* räumt jetzt auch die Spuren weg, die Wrok außerhalb des Browserprofils hinterlässt: `Alles löschen` entfernt zusätzlich das zuletzt geöffnete Bild (`%LOCALAPPDATA%\Wrok\images\last-image.png`) samt gespeicherter Quell-URL. Optional — standardmäßig abgewählt — lassen sich auch alle Makroprofile leeren. Die Erfolgsmeldung benennt jetzt ausdrücklich, was **erhalten** bleibt.
- **Makro-Menü zeigt das aktive Profil.** Statt eines Untermenüs „Profile" steht der Name des aktiven Profils als erster Eintrag und klappt zu den übrigen auf. Umbenennen und Zurücksetzen liegen darin und heißen jetzt *Aktives Profil umbenennen* bzw. *… zurücksetzen*.

### Fixes
- **Boss-Taste ließ das Medienfenster stehen.** `MinimizeToTray()` betraf nur das Hauptfenster — der Bild- und Videobetrachter ist ein eigenständiges Top-Level-Fenster und blieb sichtbar, angeheftet sogar über allem anderen. Er wird jetzt mit ausgeblendet und beim Reaktivieren samt Anheftung wiederhergestellt. Betrifft ebenso die Inaktivitäts-Automatik und das Schließen über das Fenster-X, die denselben Weg nehmen.
- **Inaktivitäts-Zeitgeber löste bei geöffnetem Bild nie aus.** Es genügte, dass der Mauszeiger *irgendwo* im Fenster lag. Da die Nebeneinander-Anordnung Bild und Wrok über die gesamte Arbeitsfläche legt, galt fast jede Zeigerposition als Aktivität. Gewertet wird jetzt nur noch ein *bewegter* Zeiger; der Zweig „Wrok hat den Fokus" ist entfallen, da ein Programm im Vordergrund gerade kein Beleg dafür ist, dass jemand davorsitzt. Ein Zeiger über dem Medienfenster zählt weiterhin als Aufmerksamkeit.
- **Offline-Grafik auf dunklem Grund.** `nonet.png` war schwarzes Motiv mit gegen Weiß gerechneten Kanten und zerfiel auf dem `#202020`-Hintergrund der Offline-Seite zu einem hellen Saum. Neu als weißes Motiv mit aus der Helligkeit abgeleitetem Alphakanal — dadurch auf jedem Hintergrund saubere Kanten.

### Intern
- **MainForm entflochten:** `Dialogs`, `ThemeManager` und `InactivityWatcher` in eigene Dateien ausgelagert, MainForm von 1861 auf rund 1330 Zeilen verkleinert.
- Verbleibende hartkodierte Zeichenketten in die Ressourcen überführt; beide Sprachdateien decken sich (143 Einträge).


## [1.1.0] - 2026-07-19

### Features
- **Makro-Profile:** Drei benennbare Sätze zu je zehn Makros. Im Tray unter *Makros → Profile*: Linksklick wechselt, Rechtsklick benennt um — dasselbe Muster wie bei den Makros selbst. Der Wechsel tauscht die komplette Hotkey-Belegung aus, praktisch für verschiedene Figuren oder Szenarien. Bestehende Makros wandern automatisch nach Profil 1.
- **Profilbewusster Import/Export:** Exportiert wird wahlweise nur das aktive Profil oder alle Profile samt Namen; der Dateiname nennt das jeweilige Profil. Beim Import wird die Datei zuerst untersucht: Enthält sie alle Profile, folgt eine Rückfrage vor dem Ersetzen; enthält sie eine einzelne Makroliste, wird das Zielprofil abgefragt (vorausgewählt ist das aktive). Ungültige Dateien werden erkannt, bevor irgendetwas verändert wird. Exportdateien früherer Versionen bleiben lesbar.
- **Mehrzeilige `{input}`-Eingabe:** Der Abfragedialog ist jetzt mehrzeilig und in der Größe veränderbar. Enter erzeugt einen Zeilenumbruch, **Strg+Enter** sendet, Esc bricht ab — für längere Erzählpassagen im Rollenspiel.

### Fixes
- **Header-Manipulation auf Dokument-Anfragen beschränkt.** Bisher wurden `Accept`, `Sec-Fetch-Dest: document` und `Sec-Fetch-Mode: navigate` auf *jede* Anfrage gesetzt — auch auf Bilder, API-Aufrufe und Datei-Uploads, für die diese Werte falsch sind. Zusätzlich entfallen `Accept-Encoding` und `Sec-Fetch-Site`, die Chromium selbst korrekt bestimmt. Nebeneffekt: kein Sprung mehr in verwalteten Code bei jeder einzelnen Anfrage.
- **Desktopsymbol wurde nicht angelegt.** Der `[Tasks]`-Eintrag war nicht mit dem `[Icons]`-Eintrag verknüpft, und `Flags: unchecked` hätte bei stiller Installation über WinGet ohnehin verhindert, dass das Symbol entsteht. Der Task ist jetzt verdrahtet und standardmäßig aktiv.
- Setup-Skript: Publish-Pfad korrigiert (`bin\x64\Release\publish`), dazu Prüfungen gegen fehlende oder versionsfremde Publish-Ausgaben — zuvor konnte stillschweigend ein veraltetes Binary eingepackt werden.
- Installer wird jetzt nach dem *Veröffentlichen* gebaut statt nach dem *Erstellen*.


## [1.0.0] - 2026-07-18

### Features
- Cache löschen mit Auswahl: **Nur Cache** (Cookies/Login bleiben erhalten) oder **Alles** (inkl. Cookies/Login). Dialog nutzt die vorhandenen Resource-Strings.
- **Autostart mit Windows**: umschaltbarer Tray-Eintrag (HKCU\…\Run, kein Admin-Recht nötig); Status wird gegen den aktuellen Programmpfad geprüft.

### Features
- **Video-Wiedergabe im Endlos-Loop.** Klick auf ein Video in Grok (oder `Strg+Shift+P` mit kopierter Video-Adresse) öffnet es in einem eigenen Fenster, links auf voller Bildschirmhöhe wie Bilder. Umgesetzt mit einem zweiten WebView2, das dieselbe `CoreWebView2Environment` — und damit dieselbe Login-Session — nutzt; das Video wird direkt gestreamt statt heruntergeladen. Start stumm (Autoplay-Sperre) mit Bedienelementen zum Aufdrehen; `Esc` schließt, `Strg+P` heftet an.
  - Die Player-Seite wird über `WebResourceRequested` unter einer `grok.com`-URL ausgeliefert, damit das Dokument den richtigen Ursprung hat und SameSite-Cookies greifen.
  - Typerkennung: beim Klick über das DOM (`<video>` vs. `<img>`), beim Zwischenablage-Weg über einen Content-Type-Probe-Request — Grok-URLs haben keine Dateiendung.
  - Bild und Video teilen sich einen Medien-Slot: es ist immer höchstens ein Medienfenster offen, und „Letztes Medium" holt zurück, was zuletzt lief.
- **Nebeneinander-Anordnung:** Ein geöffnetes Bild wird links auf volle Bildschirmhöhe skaliert, Wrok füllt automatisch den Rest rechts daneben. Es gibt immer höchstens ein Viewer-Fenster — ein neues Bild ersetzt das bisherige. Wrok bleibt nach dem Schließen in der Anordnung stehen, sodass das nächste Bild ohne Sprung aufgeht; die in den Settings gespeicherte Fenstergröße bleibt davon unberührt. Die Bildbreite ist auf 60 % der Bildschirmbreite gedeckelt, damit für Wrok genug Platz bleibt.
- **Tray-Menü neu strukturiert** — von 21 auf 8 Einträge auf oberster Ebene:
  - *Einstellungen → App*: Inaktivität, Makros Import/Export, Mit Windows starten
  - *Einstellungen → Grok*: öffnet Groks eigene Einstellungsseite
  - *Werkzeuge*: Rate Limits, Bild aus Zwischenablage öffnen, Letztes Bild öffnen, Cache löschen
  - Makros bleiben auf oberster Ebene (meistgenutzte Funktion)
- Hartkodierte deutsche Menütexte bei Makro-Import/Export durch die bereits vorhandenen Ressourcen ersetzt.
- **Tray-Eintrag „Bild aus Zwischenablage öffnen…"**: nimmt eine per Rechtsklick → „Bildadresse kopieren" kopierte URL, zeigt sie zum Bestätigen/Bearbeiten an und öffnet sie im Viewer. Erlaubt das Öffnen von Bildern unabhängig von der Klick-Erkennung.
- **Bild-Viewer ist jetzt ein eigenständiges Fenster** (kein Owner-Fenster mehr, `ShowInTaskbar`): alt-tabbar und hinter das Hauptfenster legbar, sodass während der Bildanzeige weiter Text eingegeben werden kann.
- **Globaler Hotkey `Strg+Shift+P`** öffnet ein Bild aus der Zwischenablage — ohne Rückfrage, direkt aus der kopierten URL.
- **„Letztes Bild öffnen"**: Das zuletzt angezeigte Bild wird beim Öffnen lokal unter `%LOCALAPPDATA%\Wrok\images\last-image.png` abgelegt (immer dieselbe Datei) und steht nach jedem Neustart sofort wieder zur Verfügung — ohne Netzwerk, Session oder gültige URL.
- **Viewer ohne Toolbar**: geschlossen wird über das X der Titelleiste. Funktionen liegen auf Tastenkürzeln (`Esc` schließen, `+`/`-` und Mausrad zoomen, `Strg+F` einpassen, `Strg+S` speichern, `Strg+P` anheften).
- **Fenster öffnet in Bildgröße** (auf 90 % des Arbeitsbereichs begrenzt) und ist frei skalierbar; die Fenstergröße skaliert das Bild mit. Titelleiste zeigt die Auflösung, bei aktivem Anheften zusätzlich 📌.

### Fixes
- **Bild-Viewer öffnete sich nie.** `CoreWebView2.ExecuteScriptAsync` löst keine Promises auf – die `async`-IIFE lieferte wörtlich `{}` statt Base64, was in `Convert.FromBase64String` als `FormatException` endete. Asynchrone JS-Ergebnisse laufen jetzt über eine Token-basierte `postMessage`-Brücke (`wrokResult:<token>:<payload>`).
- **Rate-Limit-Anzeige lieferte nie Daten** – gleiche Promise-Ursache, ebenfalls auf die Brücke umgestellt. Doppeltes JSON-Unquoting entfällt; Modellnamen sind jetzt Konstanten.
- Base64-Konvertierung großer Bilder erfolgt blockweise statt Zeichen für Zeichen.
- **Bild-Viewer: sporadische GDI+-Fehler behoben.** Das Bitmap wurde aus einem `MemoryStream` erzeugt, der bereits geschlossen war, wenn das Bild später auf dem UI-Thread gezeichnet wurde. Die Pixeldaten werden jetzt in ein eigenständiges Bitmap kopiert.
- **Mausrad-Zoom im Bild-Viewer funktioniert.** `WM_MOUSEWHEEL` geht nur an das fokussierte Control; die `PictureBox` war nicht fokussierbar. Sie ist jetzt selektierbar und holt sich den Fokus, sobald der Zeiger über dem Bild ist.
- **{input}-Dialog erscheint zuverlässig im Vordergrund**, auch wenn das Makro per globalem Hotkey bei minimiertem/hintergründigem Wrok ausgelöst wird (TopMost + Aktivierung, ohne Owner bei unsichtbarem Hauptfenster).

### Docs
- Boss-Taste in README, Changelog und Hilfe-Ressourcen einheitlich auf **Ctrl+Space** korrigiert (entspricht dem tatsächlich registrierten Hotkey; vorher fälschlich als Ctrl+Shift+Space bzw. Ctrl+Tab dokumentiert).

## [0.9.0] - 2025-12-28

### Highlights
- Simplified macros UX: 6 fixed macros (Macro 1…Macro 6).
  - Left-click: send text
  - Left-click + modifier (Alt/Ctrl): send text + Enter
  - Right-click: edit macro (persisted to settings)
- Improved WebView2 text injection robustness and fallbacks.
- Re-introduced global toggle hotkey (Ctrl+Space) for minimize/reactivate.

### Features
- Ensure `Properties.Settings.Default.Macros` contains six entries on startup.
- Inject `__wrokSend` helper into new and already-loaded documents.
- Targeted submit-button click fallback when Enter is requested.

### Fixes
- First-send failure fixed by injecting helper into current document.
- Restored visibility of the 6th macro item in the tray menu.
- Added JS retry + InputSimulator fallback for reliable text + Enter delivery.
- Separated hotkey handling: toggle hotkey vs macro hotkeys.

### Breaking changes / UX
- Removed per-macro subitems (Send, Send+Enter, New Macro).
- Double-click behavior removed (use modifier+click instead). Update docs/user notes.

### QA / Manual test steps
1. Open tray → Macros shows 6 items.
2. Left-click a macro → text inserted into active WebView2 input.
3. Left-click + Alt/Ctrl → text + submit.
4. Right-click macro → edit dialog; changes persist.
5. Press Ctrl+Space → app toggles minimize/reactivate.
6. Press Ctrl+1..Ctrl+5 and Ctrl+^ → corresponding macro sent.

### Rollback
- Revert changes to `RefreshMacrosMenu`, `InitializeMacrosMenu`, `SendTextToWebViewAsync`, `OnHandleCreated`, and `WndProc` to restore previous behavior.

### Suggested commit messages
- `feat(macros): simplify to 6 fixed macros and edit on right-click`
- `fix(webview): inject helper into existing doc and harden SendTextToWebViewAsync`
- `fix(hotkeys): register Ctrl+Space toggle and separate WndProc handling`