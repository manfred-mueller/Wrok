![dunkles wrok icon](/wrok_black.ico?raw=true "Dunkles Wrok-Symbol")

**Wrok** ist ein moderner, portabler WebView2-basierter Grok™-Client.

## Installation
Einfach per WinGet:
```bash
winget install --id NASS.Wrok
```

Er arbeitet komplett ohne API-Zugriff und bietet folgende Funktionen:

- Zugriff auf Grok®-Einstellungen  
- Boss-Taste **Strg + Shift + Leertaste** (sofortiges Minimieren)  
- Automatische Fensterminimierung nach einstellbarem Inaktivitäts-Zeitraum  
- Speichern von Fenstergröße, -position und Maximierungsstatus  
- Vollautomatische Anpassung an Windows Dark/Light-Mode (inkl. Tray-Icon und Titelleiste)  
- Intelligentes Löschen der Browsing-Daten (mit Auswahl: nur Cache oder alles)  
- Tray-Menü mit Schnellzugriff auf alle wichtigen Funktionen  

**NASS e.K. und Wrok sind in keiner Weise mit Grok™, xAI oder X® (ehemals Twitter) verbunden.**  
Alle in dieser Software verwendeten Markennamen und Bezeichnungen sind eingetragene Warenzeichen und Marken der jeweiligen Eigentümer und dienen nur der Beschreibung.

---

**Wrok** is a modern, portable WebView2-based Grok™ client.

It works entirely without API access and offers the following features:

- Access to Grok™ settings  
- Boss key **Ctrl + Shift + Spacebar** (instant minimize)  
- Automatic window minimization after configurable inactivity period  
- Saving of window size, position and maximized state  
- Full automatic adaptation to Windows Dark/Light mode (including tray icon and title bar)  
- Smart clearing of browsing data (with choice: cache only or everything)  
- Tray menu with quick access to all important functions  

**NASS e.K. and Wrok are in no way affiliated with Grok™, xAI or X® (formerly Twitter).**  
All brand names and designations used in this software are registered trademarks and brands of their respective owners and are used for descriptive purposes only.

![Wrok Screenshot](/Screenshot.png?raw=true "Wrok Screenshot")

## Short Manual

### Macros
- **Edit:** Right-click a macro in the tray menu → Edit.
- **Send via menu:** Left-click a macro to insert its text and send it (Send + Enter).
- **Send without Enter:** Left-click + Shift (or Alt) will insert the text but NOT send Enter.
- **Hotkeys:** Ctrl+1 .. Ctrl+5 send the corresponding macro. By default, the macro is inserted and followed by Enter (same as left-click). Hold Shift or Alt while pressing the hotkey to suppress the trailing Enter.

### Boss key
- Press `Ctrl+Tab` to quickly minimize the window to the tray.

### Clear cache
- Tray menu → "Clear cache" deletes WebView2 browsing data (cache, cookies, local storage).
- Use this when you see rendering issues, login problems, or need a clean session.
- **Note:** After clearing the cache, you may need to reload the page or sign in again.

### Troubleshooting & Tips
- If text is not sent: ensure the page is fully loaded and the editor is visible.
- **Open DevTools:** For debugging, you can open WebView2 DevTools to inspect the DOM and console.
- **Logs:** Runtime messages and logs are written 	to `%LOCALAPPDATA%\\Wrok\\logs\\app.log`.

### Quick operation overview
- Left-clicking a macro and using the hotkeys behave the same: the app inserts the text and then attempts to submit it by clicking the send button or sending a real OS Enter keystroke. Some web editors ignore synthetic JS keyboard events; therefore, the host application falls back to sending an actual Enter via the InputSimulator when necessary.
