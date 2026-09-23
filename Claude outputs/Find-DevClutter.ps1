#Requires -Version 5.1
<#
.SYNOPSIS
    Findet typische Entwickler-Build-/Cache-Ordner (bin, obj, node_modules, ...) und
    entfernt sie interaktiv über eine Konsolen-Auswahltabelle.

.DESCRIPTION
    Reiner Heuristik-Scanner OHNE KI: durchsucht die angegebenen Wurzelpfade nach
    bekannten Build-/Cache-Ordnermustern, ermittelt Größe und letzte Änderung, und
    zeigt das Ergebnis in einer interaktiven Mehrfachauswahl an (Out-ConsoleGridView,
    sofern PowerShell 7+ und das Modul verfügbar sind - sonst ein einfacher
    Text-Fallback).

    Es wird NICHTS automatisch gelöscht. Erst nach expliziter Auswahl und einer
    zusätzlichen Ja/Nein-Bestätigung werden die markierten Ordner entfernt -
    standardmäßig in den Papierkorb (wiederherstellbar), nicht endgültig.

.PARAMETER Path
    Ein oder mehrere Wurzelverzeichnisse, die durchsucht werden sollen.
    Standard: das eigene Benutzerprofil.

.PARAMETER MinSizeMB
    Nur Ordner ab dieser Größe (MB) anzeigen. Standard: 50.

.PARAMETER MinAgeDays
    Nur Ordner anzeigen, die seit mindestens so vielen Tagen nicht mehr verändert
    wurden (jüngste enthaltene Datei als Referenz). Standard: 30.

.PARAMETER PermanentDelete
    Wenn gesetzt: ausgewählte Ordner werden endgültig gelöscht statt in den
    Papierkorb verschoben.

.PARAMETER LogPath
    Pfad für das Protokoll der tatsächlich gelöschten Ordner.
    Standard: Find-DevClutter-Log.csv neben diesem Skript.

.EXAMPLE
    .\Find-DevClutter.ps1 -Path C:\Users\Manfred\source, D:\Projekte

.EXAMPLE
    .\Find-DevClutter.ps1 -MinSizeMB 200 -MinAgeDays 90 -WhatIf

.NOTES
    Ausführung ggf. mit:  powershell -ExecutionPolicy Bypass -File .\Find-DevClutter.ps1
    (falls die lokale Execution Policy das Ausführen von .ps1-Dateien verbietet).
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string[]]$Path = @([Environment]::GetFolderPath('UserProfile')),

    [int]$MinSizeMB = 50,

    [int]$MinAgeDays = 30,

    [switch]$PermanentDelete,

    [string]$LogPath = (Join-Path $PSScriptRoot 'Find-DevClutter-Log.csv')
)

$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------------
# Bekannte Build-/Cache-Ordnermuster (bewusst konservativ gehalten -
# lieber einen Treffer zu wenig als versehentlich Quellcode-Ordner
# gleichen Namens zu erwischen).
# ------------------------------------------------------------------
$Patterns = @(
    'bin', 'obj', '.vs',
    'node_modules', 'dist', 'build',
    '.gradle', '.idea',
    'packages',            # klassisches NuGet packages-Verzeichnis pro Projekt
    '__pycache__', '.pytest_cache',
    'target'               # Rust/Maven
)

# Ordner, in die grundsätzlich nicht hineingestiegen wird (teuer, heikel, oder
# schlicht kein Entwickler-Ballast).
$ExcludedTopLevel = @(
    'Windows', 'Program Files', 'Program Files (x86)', 'ProgramData',
    '$Recycle.Bin', 'System Volume Information', 'WindowsApps'
)

# ------------------------------------------------------------------
# Papierkorb-Löschung (statt endgültig) über .NET, ohne COM-Dialoge.
# ------------------------------------------------------------------
function Move-ToRecycleBin {
    param([Parameter(Mandatory)][string]$FolderPath)
    Add-Type -AssemblyName Microsoft.VisualBasic
    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory(
        $FolderPath,
        [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
        [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
    )
}

# ------------------------------------------------------------------
# Suche: iterativ (Stack statt Rekursion), bricht die Suche an einem
# Treffer ab (steigt nicht weiter in ein bereits erkanntes bin/obj/
# node_modules hinein - spart Zeit und verhindert doppelte Treffer).
# ------------------------------------------------------------------
function Get-DevArtifactFolders {
    param(
        [string[]]$RootPaths,
        [string[]]$PatternNames
    )

    $patternSet = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase)
    foreach ($p in $PatternNames) { [void]$patternSet.Add($p) }

    $results = [System.Collections.Generic.List[string]]::new()
    $stack = [System.Collections.Generic.Stack[string]]::new()

    foreach ($root in $RootPaths) {
        if (Test-Path -LiteralPath $root) { $stack.Push($root) }
        else { Write-Warning "Pfad nicht gefunden, wird übersprungen: $root" }
    }

    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        $dirName = Split-Path -Leaf $current

        if ($patternSet.Contains($dirName)) {
            $results.Add($current)
            continue   # nicht weiter hineinsteigen
        }

        $children = $null
        try {
            $children = Get-ChildItem -LiteralPath $current -Directory -Force -ErrorAction SilentlyContinue
        } catch { continue }

        foreach ($child in $children) {
            if ($ExcludedTopLevel -contains $child.Name) { continue }
            if ($child.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }  # Symlinks/Junctions meiden
            $stack.Push($child.FullName)
        }
    }

    return $results
}

function Get-FolderInfo {
    param([string]$FolderPath)

    $files = Get-ChildItem -LiteralPath $FolderPath -Recurse -File -Force -ErrorAction SilentlyContinue
    $sizeBytes = ($files | Measure-Object -Property Length -Sum).Sum
    if (-not $sizeBytes) { $sizeBytes = 0 }

    $lastWrite = if ($files) {
        ($files | Sort-Object LastWriteTime -Descending | Select-Object -First 1).LastWriteTime
    } else {
        (Get-Item -LiteralPath $FolderPath).LastWriteTime
    }

    [PSCustomObject]@{
        Path          = $FolderPath
        Kind          = Split-Path -Leaf $FolderPath
        SizeMB        = [math]::Round($sizeBytes / 1MB, 1)
        LastWriteTime = $lastWrite
        AgeDays       = [int][math]::Round(((Get-Date) - $lastWrite).TotalDays)
    }
}

# ------------------------------------------------------------------
# Interaktive Auswahl: Out-ConsoleGridView wenn möglich, sonst ein
# simpler Text-Fallback (z. B. unter Windows PowerShell 5.1, wo das
# Modul nicht läuft).
# ------------------------------------------------------------------
function Select-ClutterItems {
    param([Parameter(Mandatory)][object[]]$Items)

    $hasGridView = $false
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        if (-not (Get-Module -ListAvailable -Name Microsoft.PowerShell.ConsoleGuiTools)) {
            Write-Host "Modul 'Microsoft.PowerShell.ConsoleGuiTools' (für die Auswahltabelle) ist nicht installiert." -ForegroundColor Yellow
            $answer = Read-Host "Jetzt aus der PowerShell Gallery installieren? (j/n)"
            if ($answer -eq 'j') {
                try {
                    Install-Module Microsoft.PowerShell.ConsoleGuiTools -Scope CurrentUser -Force -ErrorAction Stop
                } catch {
                    Write-Warning "Installation fehlgeschlagen: $($_.Exception.Message)"
                }
            }
        }
        if (Get-Module -ListAvailable -Name Microsoft.PowerShell.ConsoleGuiTools) {
            Import-Module Microsoft.PowerShell.ConsoleGuiTools -ErrorAction SilentlyContinue
            $hasGridView = $null -ne (Get-Command Out-ConsoleGridView -ErrorAction SilentlyContinue)
        }
    } else {
        Write-Host "PowerShell $($PSVersionTable.PSVersion) erkannt - Out-ConsoleGridView benötigt PowerShell 7+. Nutze Text-Auswahl." -ForegroundColor Yellow
    }

    if ($hasGridView) {
        return $Items | Out-ConsoleGridView `
            -Title "Build-/Cache-Ordner auswählen  (Leertaste = markieren, Enter = übernehmen, Esc = abbrechen)" `
            -OutputMode Multiple
    }

    # ---- Text-Fallback ----
    Write-Host ""
    Write-Host "Gefundene Ordner:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $it = $Items[$i]
        "{0,3}) {1,8} MB  {2,4} Tage inaktiv  {3}" -f ($i + 1), $it.SizeMB, $it.AgeDays, $it.Path | Write-Host
    }
    Write-Host ""
    $raw = Read-Host "Nummern zum Löschen auswählen, kommagetrennt (z. B. 1,3,5), 'alle' oder Enter für keine"
    if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
    if ($raw.Trim().ToLower() -eq 'alle') { return $Items }

    $indices = $raw -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ - 1 }
    return $indices | Where-Object { $_ -ge 0 -and $_ -lt $Items.Count } | ForEach-Object { $Items[$_] }
}

# ==================================================================
# Ablauf
# ==================================================================

Write-Host "Durchsuche $($Path -join ', ') nach Build-/Cache-Ordnern ..." -ForegroundColor Cyan
$candidates = Get-DevArtifactFolders -RootPaths $Path -PatternNames $Patterns
Write-Host "$($candidates.Count) Treffer nach Namensmuster gefunden. Ermittle Größe und Alter ..." -ForegroundColor Cyan

$items = [System.Collections.Generic.List[object]]::new()
$i = 0
foreach ($c in $candidates) {
    $i++
    Write-Progress -Activity "Größe ermitteln" -Status $c -PercentComplete (($i / [math]::Max(1,$candidates.Count)) * 100)
    $items.Add((Get-FolderInfo -FolderPath $c))
}
Write-Progress -Activity "Größe ermitteln" -Completed

$filtered = $items | Where-Object { $_.SizeMB -ge $MinSizeMB -and $_.AgeDays -ge $MinAgeDays } | Sort-Object SizeMB -Descending

if (-not $filtered -or @($filtered).Count -eq 0) {
    Write-Host ""
    Write-Host "Keine Ordner gefunden, die die Schwellwerte erreichen (>= $MinSizeMB MB, >= $MinAgeDays Tage inaktiv)." -ForegroundColor Green
    Write-Host "Tipp: -MinSizeMB oder -MinAgeDays kleiner setzen, um mehr Kandidaten zu sehen."
    return
}

$totalFoundMB = [math]::Round((($filtered | Measure-Object -Property SizeMB -Sum).Sum), 1)
Write-Host ""
Write-Host "$(@($filtered).Count) Ordner oberhalb der Schwellwerte, zusammen $totalFoundMB MB." -ForegroundColor Cyan

$selected = Select-ClutterItems -Items $filtered

if (-not $selected -or @($selected).Count -eq 0) {
    Write-Host ""
    Write-Host "Keine Auswahl getroffen - es wird nichts gelöscht." -ForegroundColor Green
    return
}

$totalSelectedMB = [math]::Round((($selected | Measure-Object -Property SizeMB -Sum).Sum), 1)
Write-Host ""
Write-Host "Ausgewählt: $(@($selected).Count) Ordner, zusammen $totalSelectedMB MB" -ForegroundColor Yellow
$selected | Format-Table Path, SizeMB, AgeDays -AutoSize | Out-Host

$mode = if ($PermanentDelete) { 'ENDGUELTIG geloescht (kein Papierkorb)' } else { 'in den Papierkorb verschoben' }
$confirm = Read-Host "Wirklich fortfahren? Ordner werden $mode. Tippe 'ja' zum Bestaetigen"

if ($confirm -ne 'ja') {
    Write-Host "Abgebrochen - nichts wurde geloescht." -ForegroundColor Green
    return
}

$logEntries = [System.Collections.Generic.List[object]]::new()

foreach ($item in $selected) {
    if (-not (Test-Path -LiteralPath $item.Path)) {
        Write-Host "UEBERSPRUNGEN (existiert nicht mehr): $($item.Path)" -ForegroundColor DarkYellow
        continue
    }

    if ($PSCmdlet.ShouldProcess($item.Path, 'Loeschen')) {
        $status = 'OK'
        $errorMessage = ''
        try {
            if ($PermanentDelete) {
                Remove-Item -LiteralPath $item.Path -Recurse -Force
            } else {
                Move-ToRecycleBin -FolderPath $item.Path
            }
            Write-Host "OK: $($item.Path)" -ForegroundColor Green
        } catch {
            $status = 'FEHLER'
            $errorMessage = $_.Exception.Message
            Write-Host "FEHLER bei $($item.Path): $errorMessage" -ForegroundColor Red
        }

        $logEntries.Add([PSCustomObject]@{
            Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            Path      = $item.Path
            SizeMB    = $item.SizeMB
            Mode      = $(if ($PermanentDelete) { 'Permanent' } else { 'Papierkorb' })
            Status    = $status
            Error     = $errorMessage
        })
    }
}

if ($logEntries.Count -gt 0) {
    try {
        $writeHeader = -not (Test-Path -LiteralPath $LogPath)
        $logEntries | Export-Csv -LiteralPath $LogPath -Append -NoTypeInformation -Encoding UTF8
        Write-Host ""
        Write-Host "Protokoll geschrieben nach: $LogPath" -ForegroundColor DarkCyan
    } catch {
        Write-Warning "Protokoll konnte nicht geschrieben werden: $($_.Exception.Message)"
    }
}

Write-Host ""
Write-Host "Fertig." -ForegroundColor Cyan
