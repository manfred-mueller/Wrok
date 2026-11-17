# Git-Reste unter .git aufraeumen (siehe Clear-GitLeftovers.ps1).
& (Join-Path $PSScriptRoot 'Clear-GitLeftovers.ps1')

dotnet publish -c Release -r win-x64 /p:Platform=x64 /p:PublishProfile=FolderProfile
