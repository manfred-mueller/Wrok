<#
.SYNOPSIS
    Raeumt Reste abgebrochener oder behinderter Git-Laeufe unter .git auf.

.DESCRIPTION
    Entfernt:
      - .git\stale-*            (von Hand oder von Werkzeugen umbenannte Lock-Dateien)
      - .git\objects\**\tmp_obj_*, .git\objects\pack\tmp_pack_*, tmp_idx_*
                                (temporaere Objektdateien, die Git nicht loeschen konnte)
      - *.lock unter .git       (nur wenn kein git.exe laeuft UND die Datei aelter
                                 als -MinAgeMinutes ist - sonst koennte es die Sperre
                                 eines gerade laufenden Git-Befehls sein)

    Alle Dateien muessen aelter als -MinAgeMinutes sein, damit ein parallel
    laufender Git-Befehl (z. B. aus Visual Studio) nicht gestoert wird.
    Fehler beim Loeschen werden nur gemeldet, nie als Abbruch behandelt.

.EXAMPLE
    & "$PSScriptRoot\Clear-GitLeftovers.ps1"
    .\Clear-GitLeftovers.ps1 -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    # Repository-Wurzel; Standard ist der Ordner dieses Skripts.
    [string] $RepoRoot = $PSScriptRoot,

    [int] $MinAgeMinutes = 5
)

$gitDir = Join-Path $RepoRoot '.git'
if (-not (Test-Path $gitDir -PathType Container)) { return }

$cutoff = (Get-Date).AddMinutes(-$MinAgeMinutes)

function Get-OldFiles([string] $Path, [string[]] $Patterns, [switch] $Recurse) {
    if (-not (Test-Path $Path)) { return @() }
    Get-ChildItem -Path $Path -File -Force -Recurse:$Recurse -ErrorAction SilentlyContinue |
        Where-Object {
            $file = $_
            $file.LastWriteTime -lt $cutoff -and
                [bool]($Patterns | Where-Object { $file.Name -like $_ })
        }
}

$candidates = @()
$candidates += Get-OldFiles $gitDir @('stale-*')
$candidates += Get-OldFiles (Join-Path $gitDir 'objects') @('tmp_obj_*', 'tmp_pack_*', 'tmp_idx_*') -Recurse

# Echte Lock-Dateien nur, wenn sicher kein Git-Prozess laeuft.
$gitRunning = [bool](Get-Process -Name 'git' -ErrorAction SilentlyContinue)
$locks = Get-OldFiles $gitDir @('*.lock') -Recurse
if ($locks) {
    if ($gitRunning) {
        Write-Host "Git-Lock-Dateien gefunden, aber git.exe laeuft - werden nicht angefasst:" -ForegroundColor Yellow
        $locks | ForEach-Object { Write-Host "  $($_.FullName)" }
    }
    else {
        $candidates += $locks
    }
}

if (-not $candidates) { return }

$removed = 0
foreach ($f in $candidates) {
    if ($PSCmdlet.ShouldProcess($f.FullName, 'Entfernen')) {
        try {
            Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
            $removed++
        }
        catch {
            Write-Host "Konnte nicht entfernen: $($f.FullName) ($($_.Exception.Message))" -ForegroundColor Yellow
        }
    }
}

if ($removed) {
    Write-Host "Git-Reste entfernt: $removed Datei(en) unter .git" -ForegroundColor DarkGray
}
