<#
.SYNOPSIS
    Setzt Get-UserMailboxAccess.ps1 aus den Quelldateien in src\ zusammen.

.DESCRIPTION
    Der Quellcode liegt aufgeteilt in src\*.ps1 (sortiert nach Dateiname).
    Ausgeführt wird weiterhin EINE Datei (Get-UserMailboxAccess.ps1), weil
    - die GUI in einem eigenen Runspace läuft und
    - die EXE genau ein Skript einbettet.

    Ausgabe: UTF-8 mit genau einer BOM, Windows-Zeilenenden (CRLF).

.PARAMETER Check
    Nur prüfen, ob Get-UserMailboxAccess.ps1 zu src\ passt (Exitcode 1 bei Abweichung).

.EXAMPLE
    .\build\Merge-Source.ps1
    .\build\Merge-Source.ps1 -Check
#>
param(
    [string]$SourceDir = (Join-Path $PSScriptRoot '..\src'),
    [string]$OutFile   = (Join-Path $PSScriptRoot '..\Get-UserMailboxAccess.ps1'),
    [switch]$Check
)

$ErrorActionPreference = 'Stop'

$Files = @(Get-ChildItem -Path $SourceDir -Filter '*.ps1' -File | Sort-Object Name)
if ($Files.Count -eq 0) { throw "Keine Quelldateien in $SourceDir gefunden." }

$Sb = New-Object System.Text.StringBuilder
foreach ($F in $Files) {
    $T = [System.IO.File]::ReadAllText($F.FullName, [System.Text.Encoding]::UTF8)
    $T = $T.TrimStart([char]0xFEFF) -replace "`r`n", "`n"
    if (-not $T.EndsWith("`n")) { $T += "`n" }
    [void]$Sb.Append($T)
}
$Text = $Sb.ToString() -replace "`n", "`r`n"

if ($Check) {
    if (-not (Test-Path $OutFile)) { Write-Host "FEHLT: $OutFile" -ForegroundColor Red; exit 1 }
    $Cur = [System.IO.File]::ReadAllText($OutFile, [System.Text.Encoding]::UTF8).TrimStart([char]0xFEFF) -replace "`r`n", "`n"
    if ($Cur -ceq ($Text -replace "`r`n", "`n")) {
        Write-Host "OK: Get-UserMailboxAccess.ps1 passt zu src\ ($($Files.Count) Dateien)" -ForegroundColor Green
        exit 0
    }
    Write-Host "ABWEICHUNG: Get-UserMailboxAccess.ps1 passt nicht zu src\ - bitte .\build\Merge-Source.ps1 ausführen" -ForegroundColor Red
    exit 1
}

$OutFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutFile)
$Utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($OutFull, $Text, $Utf8Bom)
Write-Host "Zusammengesetzt: $OutFile ($($Files.Count) Dateien, $([Math]::Round($Text.Length / 1KB, 1)) KB)" -ForegroundColor Green
