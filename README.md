![dunkles wrok icon](/wrok_black.ico?raw=true "Dunkles Wrok-Symbol")

**Wrok** ist ein moderner, portabler WebView2-basierter Grok™-Client.

## Installation
Einfach per WinGet:
```bash
winget install --id NASS.Wrok
```

Er arbeitet komplett ohne API-Zugriff und bietet folgende Funktionen:

- Bis zu zehn frei belegbare Makros mit Hotkeys **Strg+1** bis **Strg+0**
- Makro-Variablen: `{input}`, `{date}`, `{time}`, `{username}`, `{clipboard}`
- Boss-Taste **Strg + Leertaste** (sofortiges Minimieren/Wiederherstellen)
- Medienansicht für Bilder und Videos, links auf voller Bildschirmhöhe – Videos im Endlos-Loop
- Automatische Fensterminimierung nach einstellbarem Inaktivitäts-Zeitraum
- Speichern von Fenstergröße, -position und Maximierungsstatus
- Vollautomatische Anpassung an Windows Dark/Light-Mode (inkl. Tray-Icon und Titelleiste)
- Intelligentes Löschen der Browsing-Daten (mit Auswahl: nur Cache oder alles)
- Optionaler Start mit Windows
- Anzeige der Grok-Rate-Limits
- Tray-Menü mit Schnellzugriff auf alle wichtigen Funktionen

**NASS e.K. und Wrok sind in keiner Weise mit Grok™, xAI oder X® (ehemals Twitter) verbunden.**  
Alle in dieser Software verwendeten Markennamen und Bezeichnungen sind eingetragene Warenzeichen und Marken der jeweiligen Eigentümer und dienen nur der Beschreibung.

---

**Wrok** is a modern, portable WebView2-based Grok™ client.

It works entirely without API access and offers the following features:

- Up to ten freely assignable macros with hotkeys **Ctrl+1** to **Ctrl+0**
- Macro variables: `{input}`, `{date}`, `{time}`, `{username}`, `{clipboard}`
- Boss key **Ctrl + Spacebar** (instant minimize/restore)
- Media viewer for images and videos, left-aligned at full screen height – videos loop endlessly
- Automatic window minimization after configurable inactivity period
- Saving of window size, position and maximized state
- Full automatic adaptation to Windows Dark/Light mode (including tray icon and title bar)
- Smart clearing of browsing data (with choice: cache only or everything)
- Optional start with Windows
- Grok rate limit display
- Tray menu with quick access to all important functions

**NASS e.K. and Wrok are in no way affiliated with Grok™, xAI or X® (formerly Twitter).**  
All brand names and designations used in this software are registered trademarks and brands of their respective owners and are used for descriptive purposes only.

![Wrok Screenshot](/Screenshot.png?raw=true "Wrok Screenshot")

## Short Manual

### Tray menu
The tray menu is grouped into two submenus:

- **Einstellungen → App** — inactivity timeout, macro import/export, start with Windows
- **Einstellungen → Grok** — opens Grok's own settings page
- **Werkzeuge** — rate limits, open image/video, clear cache

### Macros
- **Edit:** Right-click a macro in the tray menu.
- **Send via menu:** Left-click a macro to insert its text and send it (Send + Enter).
- **Send without Enter:** Left-click + Shift (or Alt) inserts the text but does NOT send Enter.
- **Hotkeys:** `Ctrl+1` … `Ctrl+0` send the corresponding macro. By default the macro is inserted and followed by Enter (same as left-click). Hold Shift or Alt while pressing the hotkey to suppress the trailing Enter.

#### Macro variables
Macro text may contain the following placeholders, which are resolved right before sending:

| Variable | Replaced with |
|---|---|
| `{input}` | Value you are asked for in a dialog. Cancelling aborts sending. |
| `{input:Question}` | Same, but the dialog shows your own label. |
| `{date}` | Current date (`dd.MM.yyyy`) |
| `{time}` | Current time (`HH:mm`) |
| `{username}` | Windows user name |
| `{clipboard}` | Current clipboard text |

Example: `Hallo Schlingeline, es ist jetzt {time} und {input: Text}`

### Media viewer
Images and videos from Grok can be opened in a separate window that is placed on the left at full screen height, while Wrok fills the remaining space. There is always at most one media window — a new item replaces the previous one.

- **Open by click:** Click an image or video in Grok.
- **Open by hotkey:** Right-click the media in Grok → *Copy image/video address*, then press `Ctrl+Shift+P`.
- **Reopen the last item:** Press `Ctrl+Shift+P` with nothing usable in the clipboard, or use *Werkzeuge → Letztes Bild öffnen*. Images are cached locally under `%LOCALAPPDATA%\Wrok\images`, so they reopen without network or session.
- **Videos** play muted in an endless loop; use the player controls to unmute.

Keyboard shortcuts inside the viewer:

| Key | Action |
|---|---|
| `Esc` | Close |
| `+` / `-` / mouse wheel | Zoom (images) |
| `Ctrl+F` | Fit to window (images) |
| `Ctrl+S` | Save image |
| `Ctrl+P` | Pin (always on top) |

Drag with the left mouse button to pan an image.

### Boss key
- Press `Ctrl+Space` to instantly minimize the window to the tray (and press it again to restore it).

### Clear cache
- *Werkzeuge → Cache löschen* offers a choice:
  - **Nur Cache löschen** — images, scripts and other cached data; you stay signed in.
  - **Alles löschen** — including cookies, login data and settings; you will be signed out.
- Use this when you see rendering issues, login problems, or need a clean session.

### Troubleshooting & Tips
- If text is not sent: ensure the page is fully loaded and the editor is visible.
- **Open DevTools:** `F12` opens the WebView2 developer tools to inspect the DOM and console.
- **Zoom the page:** `Ctrl` + `+`/`-`/`0` or `Ctrl` + mouse wheel. The zoom level is remembered.
- **Logs:** Runtime messages and logs are written to `%LOCALAPPDATA%\Wrok\logs\app.log`.

### Quick operation overview
Left-clicking a macro and using the hotkeys behave the same: the app inserts the text and then attempts to submit it by clicking the send button or sending a real OS Enter keystroke. Some web editors ignore synthetic JS keyboard events; therefore the host application falls back to sending an actual Enter via the input simulator when necessary.
