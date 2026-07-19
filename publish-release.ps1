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
    .\publish-release.ps1 -DryRun     # alles bauen, aber kein Release anlegen
#>

[CmdletBinding()]
param(
    [switch] $DryRun
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

# --- 2. Sauber veroeffentlichen ---------------------------------------------
# Alte Ausgaben entfernen, damit garantiert nichts Veraltetes eingepackt wird.
if (Test-Path 'bin\x64\Release\publish') {
    Remove-Item 'bin\x64\Release\publish' -Recurse -Force
}
if (Test-Path $setup) {
    Remove-Item $setup -Force
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

Write-Host "`nLege Release $tag an..." -ForegroundColor Cyan

if ($notesFile) {
    gh release create $tag $setup --title "Wrok $tag" --notes-file $notesFile
    $ghExit = $LASTEXITCODE
    Remove-Item $notesFile -Force -ErrorAction SilentlyContinue
}
else {
    gh release create $tag $setup --title "Wrok $tag" --generate-notes
    $ghExit = $LASTEXITCODE
}

# Auch hier gilt: externe Programme lassen PowerShell nicht von selbst abbrechen.
if ($ghExit -ne 0) {
    throw "gh release create ist mit Code $ghExit fehlgeschlagen. " +
          "Angemeldet? Pruefe mit 'gh auth status'."
}

Write-Host "`nFertig. Der Workflow 'WinGet veroeffentlichen' laeuft jetzt an." -ForegroundColor Green
