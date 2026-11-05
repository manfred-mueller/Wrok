![dunkles wrok icon](/Properties/wrok_black.ico?raw=true "Dunkles Wrok-Symbol")

**Wrok** ist ein moderner, portabler WebView2-basierter Grok™-Client.

## Installation

Einfach per WinGet:

```bash
winget install --id NASS.Wrok
```

### Voraussetzungen

* **Windows 10 (64 Bit) oder neuer**
* **Microsoft Edge WebView2 Runtime** — unter Windows 11 vorinstalliert. Fehlt sie unter Windows 10, bleibt das Fenster leer; sie lässt sich per `winget install Microsoft.EdgeWebView2Runtime` oder [von Microsoft](https://developer.microsoft.com/microsoft-edge/webview2/) nachinstallieren.
* **Grok-Konto** — die Anmeldung erfolgt regulär im eingebetteten Browser.
* Eine .NET-Installation ist **nicht** nötig; die Laufzeit ist im Programm enthalten.

Er arbeitet komplett ohne API-Zugriff und bietet folgende Funktionen:

* Bis zu zehn frei belegbare Makros je Profil, erreichbar über **F1** bis **F10** (im Tray-Menü einschaltbar)
* Makro-Variablen: `{input}`, `{date}`, `{time}`, `{day}`, `{username}`, `{clipboard}`
* Boss-Taste **Strg + Leertaste** (sofortiges Minimieren/Wiederherstellen)
* Panikschalter **Strg + Umschalt + Leertaste** beendet Wrok sofort, ohne Nachfrage
* Medien-Menü mit neun Medien-Slots, abrufbar per **Strg+1** bis **Strg+9**
* Medienansicht für Bilder und Videos, links auf voller Bildschirmhöhe – Videos im Endlos-Loop
* Automatische Fensterminimierung nach einstellbarem Inaktivitäts-Zeitraum
* Speichern von Fenstergröße, -position und Maximierungsstatus
* Vollautomatische Anpassung an Windows Dark/Light-Mode (inkl. Tray-Icon und Titelleiste)
* Intelligentes Löschen der gespeicherten Daten (mit Auswahl: nur Cache oder alles)
* Optionaler Start mit Windows
* Update-Hinweis beim Start (abschaltbar)
* Anzeige der Grok-Rate-Limits
* Tray-Menü mit Schnellzugriff auf alle wichtigen Funktionen

**NASS e.K. und Wrok sind in keiner Weise mit Grok™, xAI oder X® (ehemals Twitter) verbunden.**  
Alle in dieser Software verwendeten Markennamen und Bezeichnungen sind eingetragene Warenzeichen und Marken der jeweiligen Eigentümer und dienen nur der Beschreibung.

## Kurzanleitung

### Tray-Menü

Das Menü ist in zwei Untermenüs gegliedert:

* **Einstellungen → App** — Inaktivität, Mit Windows starten, Passwort speichern, Proxy, Nach Updates suchen
* **Makros → Makros Import/Export** — Makro-Profile exportieren und importieren
* **Medien** — Medien-Slots bearbeiten, abspielen und exportieren oder importieren
* **Einstellungen → Grok** — öffnet Groks eigene Einstellungsseite
* **Werkzeuge** — Rate Limits, Gelöschte Chats wiederherstellen, Daten löschen

### Makro-Profile

Drei benennbare Sätze zu je zehn Makros, etwa für verschiedene Figuren oder Szenarien. `Strg+Umschalt+F1` bis `Strg+Umschalt+F3` wechseln direkt zu Profil 1 bis 3. Das „aktive“ Profil ist das, das im Tray-Menü gerade ausgewählt ist.

Der erste Eintrag im Makro-Menü trägt den Namen des aktiven Profils und klappt zu den übrigen auf.

* **Wechseln:** *Makros → \[aktives Profil]* → Linksklick auf ein anderes Profil
* **Umbenennen:** *Makros → \[aktives Profil] → Aktives Profil umbenennen*
* **Zurücksetzen:** *Aktives Profil zurücksetzen* leert Name und alle zehn Makros des gerade aktiven Profils. Profile lassen sich nicht entfernen — es sind immer genau drei.
* Import und Export beziehen sich auf das aktive Profil, sofern du nicht *Alle Profile exportieren* wählst.

### Makros

* **Bearbeiten:** Rechtsklick auf ein Makro im Tray-Menü.
* **Über das Menü senden:** Linksklick fügt den Text ein und sendet ihn (Text + Enter).
* **Ohne Enter senden:** Linksklick + Umschalt (oder Alt) fügt den Text ein, sendet aber **kein** Enter.
* **Hotkeys:** `F1` bis `F10` senden die zehn Makros des aktiven Profils direkt. Die Funktion ist standardmäßig aus und lässt sich im Tray-Menü unter „F-Tasten für Makros (F1-F10)“ einschalten. Dann sind die Browser-Tastenkürzel im Hauptfenster deaktiviert. `Strg+F5` lädt die Seite neu.

#### Makro-Variablen

Makrotext darf folgende Platzhalter enthalten, die unmittelbar vor dem Senden aufgelöst werden:

|Variable|Wird ersetzt durch|
|-|-|
|`{input}`|Wert, nach dem ein Dialog fragt. Abbrechen verwirft den Sendevorgang.|
|`{input:Frage}`|Dasselbe, aber der Dialog zeigt deine eigene Beschriftung.|
|`{date}`|Aktuelles Datum (`TT.MM.JJJJ`)|
|`{time}`|Aktuelle Uhrzeit (`HH:MM`)|
|`{day}`|Voller Wochentagsname, deutsch (z. B. „Montag")|
|`{username}`|Windows-Benutzername|
|`{clipboard}`|Aktueller Text aus der Zwischenablage|

Der Eingabedialog ist mehrzeilig und in der Größe veränderbar: `Enter` erzeugt einen Zeilenumbruch, **`Strg+Enter` sendet**, `Esc` bricht ab.

Beispiel: `Hallo Schlingeline, es ist jetzt {time} und {input: Text}`

### Medienansicht

Bilder und Videos aus Grok lassen sich in einem eigenen Fenster öffnen, das links auf voller Bildschirmhöhe erscheint, während Wrok den Rest einnimmt. Es ist immer höchstens ein Medienfenster offen — ein neues ersetzt das bisherige.

* **Per Klick öffnen:** Bild oder Video in Grok anklicken.
* **Per Tastenkombination:** Rechtsklick auf das Medium → *Bildadresse kopieren*, dann `Strg+Ö`.
* **Lokale Dateien:** Bild oder Video im Explorer kopieren (`Strg+C` oder *Als Pfad kopieren*), dann `Strg+Ö`. Unterstützt werden JPG, PNG, BMP, GIF, TIFF und ICO sowie MP4, M4V, WebM, MOV und OGV. Auch lokal geöffnete Medien gelten danach als „zuletzt geöffnet“.
* **Zuletzt geöffnetes erneut anzeigen:** `Strg+Ö` drücken, wenn nichts Brauchbares in der Zwischenablage steht. Bilder werden lokal unter `%LOCALAPPDATA%\\Wrok\\images` zwischengespeichert und öffnen daher ohne Netzwerk und ohne gültige Sitzung.
* **Videos** laufen im Endlos-Loop mit Ton. `F11` wechselt zwischen Pause und Wiedergabe, auch wenn das Hauptfenster den Fokus hat. Haltepunkte in eckigen Klammern hinter der Adresse (siehe Medien-Menü) halten das Video an den angegebenen Zeiten an. Solange ein Video offen ist, sind im Hauptfenster die Browser-Tastenkürzel deaktiviert (z. B. Zoom mit Strg+Plus).

Tastenkürzel im Medienfenster:

|Taste|Wirkung|
|-|-|
|`Esc`|Schließen|
|`+` / `-`|Zoomen (Bilder)|
|`0`|100 % (Bilder)|
|`Strg+F`|Einpassen (Bilder)|
|`F11`|Pause / Wiedergabe (Videos)|
|`Strg+P`|Anheften (immer im Vordergrund)|

Ziehen mit der linken Maustaste verschiebt ein Bild.

### Medien-Menü

Neun Slots nehmen jeweils einen Namen und eine Adresse auf. Die Adresse kann ein Link oder eine lokale Datei sein, Video oder Bild.

* **Abspielen:** Linksklick auf einen Slot im Menü *Medien* oder `Strg+1` bis `Strg+9`. Das gilt auch aus dem Video-Fenster heraus.
* **Bearbeiten:** Rechtsklick auf einen Slot. Eine leere Adresse leert den Slot.
* **Import/Export:** *Medien → Medien Import/Export*. Die Datei ist von den Makros getrennt. Beim Import werden alle neun Slots ersetzt.

### Boss-Taste

`Strg+Leertaste` minimiert das Fenster sofort in den Infobereich; erneutes Drücken holt es zurück.

Ein offenes Medienfenster verschwindet mit — auch wenn es angeheftet ist — und kehrt beim Zurückholen samt Anheftung wieder. Die automatische Minimierung bei Inaktivität und das Schließen über das Fenster-X verhalten sich genauso. Ein echtes Sitzungsende — Abmelden, Herunterfahren, Update — beendet Wrok dagegen wirklich.

### Daten löschen

*Werkzeuge → Daten löschen* bietet zwei Möglichkeiten:

* **Nur Cache löschen** — Bilder, Skripte und andere zwischengespeicherte Daten; du bleibst angemeldet.
* **Alles löschen** — zusätzlich Cookies und Anmeldedaten (du wirst abgemeldet) sowie das zuletzt geöffnete Bild samt gespeicherter Adresse.

Zusätzlich ankreuzbar: **alle Makroprofile leeren**, standardmäßig abgewählt. Nicht betroffen sind Fenstergröße, Inaktivitätszeit und die übrigen Programmeinstellungen.

Sinnvoll bei Darstellungsfehlern, Anmeldeproblemen oder wenn eine saubere Sitzung nötig ist.

### Updates

Wrok sieht beim Start (höchstens einmal täglich) nach, ob auf GitHub eine neuere Version vorliegt, und zeigt dann einen Hinweis im Infobereich sowie einen Menüeintrag zur Release-Seite. Es wird nichts heruntergeladen — die Aktualisierung läuft weiter über WinGet bzw. den Installer.

* **Manuell prüfen:** im Info-Dialog (*Über Wrok*) auf *Nach Updates suchen*.
* **Abschalten:** *Einstellungen → App → Nach Updates suchen*. Die Prüfung kontaktiert dabei GitHub.

### Fehlerbehebung und Tipps

* Wird Text nicht abgeschickt: sicherstellen, dass die Seite vollständig geladen und das Eingabefeld sichtbar ist.
* **DevTools öffnen:** `Strg+F12` öffnet die WebView2-Entwicklerwerkzeuge zur Prüfung von DOM und Konsole.
* **Vollbild:** `Strg+F11` schaltet das Vollbild um.
* **Medien-Menü:** Neun Medien-Slots (Menü *Medien*) lassen sich per `Strg+1` bis `Strg+9` abspielen. Linksklick spielt, Rechtsklick bearbeitet.
* **Haltepunkte:** Hinter einer Video-Adresse steht eine Liste in eckigen Klammern, z. B. `https://… [5, 35, 1:41]`. Das Video hält bei jeder Zeit an, `F11` setzt bis zum nächsten Punkt fort. Zeiten in Sekunden oder als `m:ss`.
* **Seite zoomen:** `Strg` + `+`/`-`/`0` oder `Strg` + Mausrad. Die Zoomstufe wird gemerkt.
* **Protokoll:** Meldungen und Fehler stehen in `%LOCALAPPDATA%\\Wrok\\logs\\app.log`.

\---

**Wrok** is a modern, portable WebView2-based Grok™ client.

### Requirements

* **Windows 10 (64-bit) or newer**
* **Microsoft Edge WebView2 Runtime** — preinstalled on Windows 11. If it is missing on Windows 10 the window stays blank; install it via `winget install Microsoft.EdgeWebView2Runtime` or [from Microsoft](https://developer.microsoft.com/microsoft-edge/webview2/).
* **Grok account** — you sign in normally inside the embedded browser.
* No .NET installation required; the runtime ships with the application.

It works entirely without API access and offers the following features:

* Up to ten freely assignable macros per profile, reachable via **F1** to **F10** (can be switched on in the tray menu)
* Media menu with nine slots, playable via **Ctrl+1** to **Ctrl+9**
* Video stop points: add a list such as `[5, 35, 1:41]` after the video address; playback pauses at each time, **F11** continues to the next one
* Macro variables: `{input}`, `{date}`, `{time}`, `{day}`, `{username}`, `{clipboard}`
* Boss key **Ctrl + Spacebar** (instant minimize/restore)
* Panic switch **Ctrl + Shift + Spacebar** quits Wrok immediately, without a prompt
* Media viewer for images and videos, left-aligned at full screen height – videos loop endlessly
* Automatic window minimization after configurable inactivity period
* Saving of window size, position and maximized state
* Full automatic adaptation to Windows Dark/Light mode (including tray icon and title bar)
* Smart clearing of stored data (with choice: cache only or everything)
* Optional start with Windows
* Update notice at startup (can be turned off)
* Grok rate limit display
* Tray menu with quick access to all important functions

**NASS e.K. and Wrok are in no way affiliated with Grok™, xAI or X® (formerly Twitter).**  
All brand names and designations used in this software are registered trademarks and brands of their respective owners and are used for descriptive purposes only.

!\[Wrok Screenshot](/Screenshot.png?raw=true "Wrok Screenshot")

## Short Manual

### Tray menu

The tray menu is grouped into two submenus:

* **Einstellungen → App** — inactivity timeout, macro import/export, start with Windows, password saving, proxy, check for updates
* **Einstellungen → Grok** — opens Grok's own settings page
* **Werkzeuge** — rate limits, restore deleted chats, delete data

### Macro profiles

Three named sets of ten macros each, for different characters or scenarios. `Ctrl+Shift+F1` to `Ctrl+Shift+F3` switch directly to profile 1 to 3. The "active" profile is the one currently selected in the tray menu.

The first entry of the macro menu carries the active profile's name and expands to the others.

* **Switch:** *Makros → \[active profile]* → left-click another profile.
* **Rename:** *Makros → \[active profile] → Aktives Profil umbenennen*.
* **Reset:** *Aktives Profil zurücksetzen* clears the name and all ten macros of the currently active profile. Profiles cannot be removed — there are always exactly three.
* Import/export always refer to the active profile unless you choose *Alle Profile exportieren*.

### Macros

* **Edit:** Right-click a macro in the tray menu.
* **Send via menu:** Left-click a macro to insert its text and send it (Send + Enter).
* **Send without Enter:** Left-click + Shift (or Alt) inserts the text but does NOT send Enter.
* **Hotkeys:** `F1` to `F10` send the ten macros of the active profile directly. This is off by default and can be switched on in the tray menu under "F-Tasten für Makros (F1-F10)". While it is on, the browser shortcuts in the main window are disabled. `Ctrl+F5` reloads the page.

#### Macro variables

Macro text may contain the following placeholders, which are resolved right before sending:

|Variable|Replaced with|
|-|-|
|`{input}`|Value you are asked for in a dialog. Cancelling aborts sending.|
|`{input:Question}`|Same, but the dialog shows your own label.|

The input dialog is multi-line and resizable: `Enter` inserts a line break, **`Ctrl+Enter` sends**, `Esc` aborts.

| `{date}` | Current date (`dd.MM.yyyy`) |
| `{time}` | Current time (`HH:mm`) |
| `{day}` | Full weekday name, in German (e.g. "Montag"), regardless of UI language |
| `{username}` | Windows user name |
| `{clipboard}` | Current clipboard text |

Example: `Hallo Schlingeline, es ist jetzt {time} und {input: Text}`

### Media viewer

Images and videos from Grok can be opened in a separate window that is placed on the left at full screen height, while Wrok fills the remaining space. There is always at most one media window — a new image or video replaces the content of the existing window.

* **Open by click:** Click an image or video in Grok.
* **Open by hotkey:** Right-click the media in Grok → *Copy image/video address*, then press `Ctrl+Ö`.
* **Reopen the last item:** Press `Ctrl+Ö` with nothing usable in the clipboard. Images are cached locally under `%LOCALAPPDATA%\\Wrok\\images`, so they reopen without network or session.
* **Videos** play with sound in an endless loop. `F11` toggles pause and play, even while the main window has focus. Stop points in square brackets after the address (see the media menu) pause the video at the given times. While a video is open, the browser shortcuts in the main window (e.g. Ctrl+Plus zoom) are disabled.

Keyboard shortcuts inside the viewer:

|Key|Action|
|-|-|
|`Esc`|Close|
|`+` / `-`|Zoom (images)|
|`0`|100 % (images)|
|`Ctrl+F`|Fit to window (images)|
|`F11`|Pause / play (videos)|
|`Ctrl+P`|Pin (always on top)|

Drag with the left mouse button to pan an image.

### Media menu

Nine slots each hold a name and an address. The address can be a link or a local file, video or image.

* **Play:** Left-click a slot in the *Media* menu or press `Ctrl+1` to `Ctrl+9`. This also works from the video window.
* **Edit:** Right-click a slot. An empty address clears the slot.
* **Import/Export:** *Medien → Medien Import/Export*. The file is separate from the macros. Importing replaces all nine slots.

### Boss key

* Press `Ctrl+Space` to instantly minimize the window to the tray (and press it again to restore it).
* An open media window is hidden along with it — even when pinned — and comes back pinned as before. The inactivity timer and closing via the window's X behave the same way. A real session end — logging off, shutting down, an update — closes Wrok for good instead of parking it.

### Delete data

* *Werkzeuge → Daten löschen* offers a choice:

  * **Nur Cache löschen** — images, scripts and other cached data; you stay signed in.
  * **Alles löschen** — additionally cookies and login data (you will be signed out) plus the last opened image together with its stored address.
* An extra checkbox clears **all macro profiles**; unchecked by default. Not affected: window size, inactivity timeout and the remaining program settings.
* Use this when you see rendering issues, login problems, or need a clean session.

### Updates

At startup (at most once a day) Wrok checks GitHub for a newer version and then shows a tray notice plus a menu entry linking to the release page. Nothing is downloaded — the update itself goes through WinGet or the installer.

* **Check manually:** in the info dialog (*About Wrok*), click *Check for updates*.
* **Turn it off:** *Einstellungen → App → Check for updates*. The check contacts GitHub.

### Troubleshooting \& Tips

* If text is not sent: ensure the page is fully loaded and the editor is visible.
* **Open DevTools:** `Ctrl+F12` opens the WebView2 developer tools to inspect the DOM and console.
* **Fullscreen:** `Ctrl+F11` toggles fullscreen.
* **Media menu:** Nine media slots (menu *Media*) play with `Ctrl+1` to `Ctrl+9`. Left-click plays, right-click edits.
* **Stop points:** Add a list in square brackets after a video address, e.g. `https://… [5, 35, 1:41]`. Playback pauses at each time; `F11` continues to the next one. Times in seconds or as `m:ss`.
* **Zoom the page:** `Ctrl` + `+`/`-`/`0` or `Ctrl` + mouse wheel. The zoom level is remembered.
* **Logs:** Runtime messages and logs are written to `%LOCALAPPDATA%\\Wrok\\logs\\app.log`.

### Quick operation overview

Left-clicking a macro and using the hotkeys behave the same: the app inserts the text and then attempts to submit it by clicking the send button or sending a real OS Enter keystroke. Some web editors ignore synthetic JS keyboard events; therefore the host application falls back to sending an actual Enter via the input simulator when necessary.

