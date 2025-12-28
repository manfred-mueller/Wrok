# CONTRIBUTING.md

## Guidelines

- Folge dem vorhandenen Namensschema: PascalCase für Methoden und Eigenschaften, camelCase für lokale Variablen.
- Asynchrone Methoden müssen das Suffix `Async` tragen.
- Vermeide doppelte Methodensignaturen in derselben Klasse.
- Nutzen von `Properties.Settings.Default` für persistente, benutzerspezifische Einstellungen.

## Coding Standards

- Verwende `async`/`await` für asynchrone APIs.
- Halte Ausnahmebehandlung lokal und logge Fehler mit `Trace.WriteLine`.
- Verwende `using`-Deklarationen für Disposable-Objekte.
- Verwende Nullable-Reference-Typen (sofern aktiviert) korrekt.

## Settings / Makros

- User-scoped Setting: `Macros` (Typ: `System.Collections.Specialized.StringCollection`, Standard: leer)
- Makros werden als Liste von Strings gespeichert; die UI zeigt pro Makro ein Untermenü mit Aktionen (Senden, Senden+Enter, Bearbeiten, Löschen).
- Beim Start werden Makros aus `Properties.Settings.Default.Macros` geladen; Änderungen werden in `Properties.Settings.Default` gespeichert und `Save()` aufgerufen.

## Features implementieren

- Änderungen am Tray-Kontextmenü sollten in `InitializeTrayIcon` eingebunden werden. Erstelle Hilfsmethoden zum Laden/Speichern und Anzeigen der Makros (z.B. `InitializeMacrosMenu`, `RefreshMacrosMenu`, `EditMacro`, `SaveMacros`).
- Asynchrone Interaktionen mit `WebView2` müssen `SendTextToWebViewAsync` verwenden (einzige Implementierung, keine Duplikate).

## Tests

- Schreibe Unit-Tests für Serialisierung / Deserialisierung der Makros (falls Logik vorhanden).

## Pull Requests

- Einer klaren Beschreibung des Features beifügen.
- Kleine, fokussierte Commits. Verwende aussagekräftige Commit-Messages.