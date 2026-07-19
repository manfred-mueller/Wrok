<#
.SYNOPSIS
    Veroeffentlicht Wrok: publizieren, Setup bauen, Release auf GitHub anlegen.

.DESCRIPTION
    Ersetzt den Weg ueber das eingecheckte Setup. Das Installationspaket bleibt
    lokal und wird als Release-Asset hochgeladen - das Repository bleibt schlank.

    Da das Release mit dem persoenlichen Token der GitHub CLI angelegt wird,
    loest es den Workflow .github/workflows/winget.yml aus. (Ein Release, das
    innerhalb einer Action mit GITHUB_TOKEN erstellt wird, taete das nicht.)

.NOTES
    Voraussetzung: GitHub CLI installiert und angemeldet (gh auth login).
    Die Version wird aus Properties\AssemblyInfo.cs gelesen - sie ist die
    einzige Stelle, an der sie gepflegt werden muss (neben InstallScript.iss,
    dessen Uebereinstimmung das Setup-Skript selbst prueft).

.EXAMPLE
    .\publish-release.ps1
    .\publish-release.ps1 -DryRun       # alles bauen, aber kein Release anlegen
    .\publish-release.ps1 -PreRelease   # Release als Vorabversion - WinGet bleibt aussen vor
    .\publish-release.ps1 -NoClean      # ohne vollstaendigen Neubau (nur fuer Probelaeufe)
#>

[CmdletBinding()]
param(
    [switch] $DryRun,

    # Legt das Release als Vorabversion an. GitHub meldet dann das Ereignis
    # "prereleased" statt "released" - der WinGet-Workflow horcht auf "released"
    # und laeuft folglich nicht an. So laesst sich eine Fassung erst selbst
    # benutzen und spaeter, nach Entfernen der Markierung oder von Hand, an
    # WinGet weiterreichen.
    [switch] $PreRelease,

    # Ueberspringt das Loeschen der Ausgabeordner. Spart beim wiederholten
    # Probelauf die Zeit fuer den vollstaendigen Neubau samt ReadyToRun - fuer
    # ein echtes Release aber nicht zu empfehlen.
    [switch] $NoClean
)

$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

# Ausgaben von dotnet/MSBuild auf Englisch. Zwei Gruende:
#   1. Die deutsche Uebersetzung des Terminal-Loggers ist fehlerhaft - im
#      Abschlussbericht fehlt der Projektname ("Erstellen von Erfolgreich in ...").
#   2. Englische Fehlercodes und Meldungstexte lassen sich nachschlagen;
#      Dokumentation und Suchergebnisse sind durchweg englisch.
# Gilt nur fuer diesen Skriptlauf, nicht fuer die uebrige Umgebung.
$env:DOTNET_CLI_UI_LANGUAGE = 'en'

# --- 0. Werkzeuge pruefen ---------------------------------------------------
# Frueh statt spaet: Ohne diese Pruefung faellt ein fehlendes gh erst nach
# mehreren Minuten Bauzeit auf, unmittelbar vor dem Anlegen des Releases.
if (-not $DryRun) {
    if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
        throw "GitHub CLI nicht gefunden. Installieren mit:`n" +
              "  winget install --id GitHub.cli`n" +
              "Danach PowerShell neu starten (PATH) und 'gh auth login' ausfuehren."
    }

    gh auth status *> $null
    if ($LASTEXITCODE -ne 0) {
        throw "GitHub CLI ist nicht angemeldet. Bitte 'gh auth login' ausfuehren."
    }
}

# --- 1. Version ermitteln ---------------------------------------------------
$assemblyInfo = Get-Content 'Properties\AssemblyInfo.cs' -Raw
$version = [regex]::Match($assemblyInfo, 'AssemblyFileVersion\("(\d+\.\d+\.\d+)').Groups[1].Value

if ([string]::IsNullOrWhiteSpace($version)) {
    throw 'Version konnte nicht aus Properties\AssemblyInfo.cs gelesen werden.'
}

$tag   = "v$version"
$setup = "bin\x64\Release\Wrok-Setup-$version.exe"

Write-Host "Version: $version" -ForegroundColor Cyan
Write-Host "Tag:     $tag"     -ForegroundColor Cyan

# --- 1b. Arbeitsverzeichnis pruefen -----------------------------------------
# "gh release create" setzt den Tag auf den Stand, der SERVERSEITIG im
# Standard-Branch liegt. Nicht committete oder nicht gepushte Aenderungen
# wuerden dazu fuehren, dass das Release auf veralteten Quellcode zeigt,
# waehrend das Setup den neuen Stand enthaelt.
if (-not $DryRun) {
    $dirty = git status --porcelain
    if ($dirty) {
        Write-Host "`nNicht committete Aenderungen:" -ForegroundColor Yellow
        $dirty | ForEach-Object { Write-Host "  $_" }
        throw 'Bitte erst committen und pushen - sonst zeigt das Release auf veralteten Quellcode.'
    }

    $branch = (git rev-parse --abbrev-ref HEAD).Trim()
    git fetch --quiet origin $branch 2>$null
    $ahead = (git rev-list --count "origin/$branch..HEAD" 2>$null)

    if ($ahead -and [int]$ahead -gt 0) {
        throw "$ahead Commit(s) noch nicht gepusht. Bitte erst 'git push' ausfuehren."
    }
}

# --- 2. Alte Ausgaben entfernen ---------------------------------------------
# Absichtlich mehr als nur der publish-Ordner: Auch die Zwischenergebnisse unter
# obj verschwinden, damit kein inkrementeller Bau ein veraltetes Teilstueck
# weiterreicht. Die Versionspruefung weiter unten faengt zwar eine falsche
# Versionsnummer ab - nicht aber alten Code, der zufaellig dieselbe Nummer traegt.
# Genau das ist beim Wechsel zwischen Bauen in Visual Studio und Veroeffentlichen
# per Skript der wahrscheinlichere Fall.
#
# Bewusst NUR die x64-Release-Zweige: Debug-Staende und die AnyCPU-Ordner bleiben
# stehen, damit Visual Studio nach einem Release nicht alles neu uebersetzen muss.
if (-not $NoClean) {
    foreach ($dir in @('bin\x64\Release', 'obj\x64\Release')) {
        if (Test-Path $dir) {
            Write-Host "Entferne $dir" -ForegroundColor DarkGray
            Remove-Item $dir -Recurse -Force
        }
    }
}
else {
    Write-Host "-NoClean: alte Ausgaben bleiben stehen." -ForegroundColor Yellow

    # Zumindest das, was sonst unbemerkt weiterverwendet wuerde.
    if (Test-Path 'bin\x64\Release\publish') { Remove-Item 'bin\x64\Release\publish' -Recurse -Force }
    if (Test-Path $setup)                    { Remove-Item $setup -Force }
}

Write-Host "`nVeroeffentliche..." -ForegroundColor Cyan
# "publish" umfasst Wiederherstellen, Kompilieren und Veroeffentlichen.
# Das csproj-Target "BuildInstaller" ruft danach automatisch iscc auf.
# /p:Platform=x64 ist noetig: Ohne diese Angabe nimmt MSBuild den Standard
# "AnyCPU" und legt das Kompilat nach bin\Release\ statt bin\x64\Release\.
# Das <Platform>-Element im Publish-Profil greift dafuer zu spaet, weil aus
# Platform bereits der Ausgabepfad gebildet wird, bevor das Profil geladen ist.
dotnet publish -c Release -r win-x64 /p:Platform=x64 /p:PublishProfile=FolderProfile

# WICHTIG: PowerShell wertet Rueckgabewerte externer Programme nicht als Fehler
# aus - $ErrorActionPreference = 'Stop' greift bei dotnet und iscc nicht. Ohne
# diese Pruefung liefe das Skript nach einem Compilerfehler einfach weiter und
# meldete erst spaeter "Setup wurde nicht erzeugt", was an der falschen Stelle
# suchen laesst.
if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish ist mit Code $LASTEXITCODE fehlgeschlagen. Fehlermeldungen stehen oben."
}

if (-not (Test-Path $setup)) {
    throw "Kompilierung lief durch, aber das Setup fehlt: $setup`n" +
          "Vermutlich ist iscc nicht gelaufen oder hat abgebrochen - siehe Target BuildInstaller in Wrok.csproj " +
          "und die Versionspruefung in InstallScript.iss."
}

$size = [math]::Round((Get-Item $setup).Length / 1MB, 1)
Write-Host "`nSetup erstellt: $setup ($size MB)" -ForegroundColor Green

# --- 3. Gegenprobe: steckt wirklich die richtige Version drin? --------------
$exe = 'bin\x64\Release\publish\Wrok.exe'
$exeVersion = (Get-Item $exe).VersionInfo.FileVersion
Write-Host "Exe-Dateiversion: $exeVersion"

if (-not $exeVersion.StartsWith($version)) {
    throw "Die veroeffentlichte Exe meldet $exeVersion, erwartet wurde $version."
}

if ($DryRun) {
    Write-Host "`n-DryRun: Release wird nicht angelegt." -ForegroundColor Yellow
    exit 0
}

# --- 4. Release anlegen und Asset hochladen ---------------------------------
# Existiert der Tag schon, wird das Release ersetzt - so ist ein zweiter
# Anlauf nach einem Fehlschlag gefahrlos moeglich.
$exists = $false
try { gh release view $tag *> $null; $exists = $true } catch { }

if ($exists) {
    Write-Host "`nRelease $tag existiert bereits - wird geloescht." -ForegroundColor Yellow
    gh release delete $tag --yes --cleanup-tag
}

# --- 5. Release-Notizen aus dem CHANGELOG ziehen ----------------------------
# Der gepflegte CHANGELOG-Abschnitt ist fuer Leser deutlich nuetzlicher als
# eine automatisch erzeugte Liste von Commit-Titeln.
$notesFile = $null

if (Test-Path 'CHANGELOG.md') {
    $changelog = Get-Content 'CHANGELOG.md' -Raw

    # Abschnitt der aktuellen Version: von "## [1.1.0]" bis zur naechsten "## "-Ueberschrift.
    $pattern = '(?ms)^##\s*\[' + [regex]::Escape($version) + '\].*?$(.*?)(?=^##\s|\z)'
    $match   = [regex]::Match($changelog, $pattern)

    if ($match.Success) {
        $notes = $match.Groups[1].Value.Trim()
        if (-not [string]::IsNullOrWhiteSpace($notes)) {
            $notesFile = [System.IO.Path]::GetTempFileName()
            Set-Content -Path $notesFile -Value $notes -Encoding utf8
            Write-Host "Release-Notizen aus CHANGELOG.md uebernommen." -ForegroundColor Cyan
        }
    }
    else {
        Write-Host "Kein CHANGELOG-Abschnitt fuer $version gefunden - GitHub erzeugt die Notizen." -ForegroundColor Yellow
    }
}

$kind = if ($PreRelease) { 'Vorabversion' } else { 'Release' }
Write-Host "`nLege $kind $tag an..." -ForegroundColor Cyan

# Argumente sammeln statt die Aufrufe zu verdoppeln - so kann keine Variante
# beim Aendern vergessen werden.
$ghArgs = @('release', 'create', $tag, $setup, '--title', "Wrok $tag")

if ($notesFile) { $ghArgs += @('--notes-file', $notesFile) }
else            { $ghArgs += '--generate-notes' }

if ($PreRelease) { $ghArgs += '--prerelease' }

gh @ghArgs
$ghExit = $LASTEXITCODE

if ($notesFile) { Remove-Item $notesFile -Force -ErrorAction SilentlyContinue }

# Auch hier gilt: externe Programme lassen PowerShell nicht von selbst abbrechen.
if ($ghExit -ne 0) {
    throw "gh release create ist mit Code $ghExit fehlgeschlagen. " +
          "Angemeldet? Pruefe mit 'gh auth status'."
}

if ($PreRelease) {
    Write-Host "`nFertig - als Vorabversion angelegt." -ForegroundColor Green
    Write-Host "Der Workflow 'WinGet veroeffentlichen' laeuft NICHT an." -ForegroundColor Yellow
    Write-Host "Zum Nachreichen spaeter eines von beidem:" -ForegroundColor Yellow
    Write-Host "  gh release edit $tag --prerelease=false      # loest den Workflow aus"
    Write-Host "  gh workflow run winget.yml -f tag=$tag          # von Hand anstossen"
}
else {
    Write-Host "`nFertig. Der Workflow 'WinGet veroeffentlichen' laeuft jetzt an." -ForegroundColor Green
}
