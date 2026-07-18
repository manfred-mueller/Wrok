# Refactoring-Anleitung

## Neue Dateien

| Datei | Inhalt |
|---|---|
| `MacroManager.cs` | Makro-Datenmodell, Laden/Speichern mit In-Memory-Cache |
| `WebViewManager.cs` | WebView2-Init, JS-Helper laden, Text senden, Offline-Seite |
| `WrokHelper.js` | JS-Helper (eingebettete Ressource, ersetzt den C#-String) |
| `MainForm.cs` | Schlanke Koordinationsklasse (~400 statt ~2000 Zeilen) |

## Schritte

### 1. Alte MainForm.cs ersetzen
Die neue `MainForm.cs` ersetzt die alte vollständig. Die `NativeInput`-Klasse bleibt unverändert in ihrer eigenen Datei (oder kann am Ende von `WebViewManager.cs` bleiben – sie war vorher am Ende von `MainForm.cs`).

### 2. Neue CS-Dateien hinzufügen
`MacroManager.cs` und `WebViewManager.cs` einfach zum Projekt hinzufügen.

### 3. WrokHelper.js als EmbeddedResource einbinden
In der `.csproj`-Datei:

```xml
<ItemGroup>
  <EmbeddedResource Include="WrokHelper.js" />
</ItemGroup>
```

Danach ist die Datei über `Assembly.GetManifestResourceStream("Wrok.WrokHelper.js")` erreichbar
– genau so, wie `WebViewManager.BuildHelperScript()` es erwartet.

### 4. NativeInput-Klasse prüfen
Die `NativeInput`-Klasse war am Ende der alten `MainForm.cs`. Sie kann:
- In eine eigene Datei `NativeInput.cs` verschoben werden (empfohlen), oder
- Am Ende von `WebViewManager.cs` bleiben (sie ist `internal sealed`, kein Namespace-Konflikt)

## Was sich geändert hat

### MacroManager – Caching
**Vorher:** `LoadMacros()` wurde bei jedem Hotkey-Drücken und jedem Menü-Klick neu
deserialisiert (JSON-Roundtrip).

**Nachher:** `GetMacros()` lädt einmalig beim ersten Aufruf und hält die Liste im Speicher.
`UpdateMacro()` und `SaveAll()` schreiben gleichzeitig in Cache und Settings.

### WrokHelper.js – Eingebettete Ressource
**Vorher:** ~200 Zeilen JS als C#-String-Literal mit `@"..."` und `+`-Konkatenation.

**Nachher:** Eigenständige `.js`-Datei, die als `EmbeddedResource` kompiliert wird.
`WebViewManager.BuildHelperScript()` liest sie zur Laufzeit und setzt nur die
Versionsnummer (`__WROK_VERSION__`) per `Replace()` ein.

Vorteil: Die Datei hat Syntax-Highlighting, kann getestet werden, und Änderungen
am JS erzeugen keinen Diff-Rauschen in C#-Dateien.

### MainForm.cs – Nur noch Koordination
**Vorher:** ~2000 Zeilen, alle Verantwortlichkeiten in einer Datei.

**Nachher:** ~400 Zeilen. Die Klasse kennt `MacroManager` und `WebViewManager`
und verdrahtet sie mit dem UI – mehr nicht.
