<#
.SYNOPSIS
    Baut MailboxAccess.exe als Starter für Get-UserMailboxAccess.ps1.

.DESCRIPTION
    Das Skript Get-UserMailboxAccess.ps1 wird in die EXE eingebettet.
    Beim Start schreibt die EXE das Skript nach %LOCALAPPDATA%\EXOMailboxAccess\
    und führt es im echten powershell.exe mit verstecktem Fenster aus.

    Vorteil gegenüber direktem PS2EXE:
    - Exchange-Online-Modul läuft in der normalen PowerShell-Umgebung (keine Verbindungsabbrüche)
    - WAM-Anmeldung mit Kontoauswahl und YubiKey funktioniert (verstecktes Konsolenfenster als Parent)
    - Weiterhin nur eine EXE zum Verteilen

    Seit v4.0 liegt der Quellcode in src\. Vor dem Bauen wird Get-UserMailboxAccess.ps1
    automatisch mit build\Merge-Source.ps1 zusammengesetzt (abschaltbar mit -SkipMerge).
    Ohne -Version wird die Version aus $script:AppVersion im Skript übernommen.

.EXAMPLE
    .\Build-MailboxAccessExe.ps1
    .\Build-MailboxAccessExe.ps1 -Version 4.0.0.0
    .\Build-MailboxAccessExe.ps1 -IconFile .\mailbox.ico -SkipMerge
#>

param(
    [string]$ScriptPath = '.\Get-UserMailboxAccess.ps1',
    [string]$OutputFile = '.\MailboxAccess.exe',
    [string]$IconFile,
    [string]$Version,
    [switch]$SkipMerge
)

$ErrorActionPreference = 'Stop'

# --- Quellen zusammensetzen (src\*.ps1 -> Get-UserMailboxAccess.ps1) ---
$MergeScript = Join-Path $PSScriptRoot 'build\Merge-Source.ps1'
if (-not $SkipMerge -and (Test-Path $MergeScript) -and (Test-Path (Join-Path $PSScriptRoot 'src'))) {
    & $MergeScript
}

# --- Version aus dem Skript übernehmen (4.0.0 -> 4.0.0.0) ---
if (-not $Version) {
    $Src = [System.IO.File]::ReadAllText((Resolve-Path $ScriptPath).Path, [System.Text.Encoding]::UTF8)
    $M   = [regex]::Match($Src, "\`$script:AppVersion\s*=\s*'([0-9.]+)'")
    $Version = if ($M.Success) { $M.Groups[1].Value } else { '1.0.0' }
}
$VParts = @($Version.TrimStart('v', 'V').Split('.') | Where-Object { $_ -ne '' })
while ($VParts.Count -lt 4) { $VParts += '0' }
$Version = ($VParts | Select-Object -First 4) -join '.'
Write-Host "Version: $Version" -ForegroundColor DarkGray

# --- Voraussetzungen ---
if (-not (Get-Module ps2exe -ListAvailable)) {
    Write-Host "Installiere ps2exe..." -ForegroundColor Yellow
    Install-Module ps2exe -Scope CurrentUser -Force
}
Import-Module ps2exe

$ScriptFull = (Resolve-Path $ScriptPath).Path
$OutputFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFile)

# --- Skript einlesen und als UTF-8 MIT BOM einbetten ---
# Windows PowerShell 5.1 liest Dateien ohne BOM als ANSI -> Umlaute/Sonderzeichen kaputt.
# Daher: Kodierung erkennen und immer mit genau einer BOM einbetten.
$Raw = [System.IO.File]::ReadAllBytes($ScriptFull)
if ($Raw.Length -ge 3 -and $Raw[0] -eq 0xEF -and $Raw[1] -eq 0xBB -and $Raw[2] -eq 0xBF) {
    $Text = [System.Text.Encoding]::UTF8.GetString($Raw, 3, $Raw.Length - 3)
    Write-Host "Kodierung: UTF-8 mit BOM" -ForegroundColor DarkGray
} else {
    try {
        $Strict = New-Object System.Text.UTF8Encoding($false, $true)   # wirft bei ungültigem UTF-8
        $Text   = $Strict.GetString($Raw)
        Write-Host "Kodierung: UTF-8 ohne BOM -> BOM wird ergänzt" -ForegroundColor Yellow
    } catch {
        $Text = [System.Text.Encoding]::GetEncoding(1252).GetString($Raw)
        Write-Host "Kodierung: ANSI (Windows-1252) -> wird nach UTF-8 mit BOM konvertiert" -ForegroundColor Yellow
    }
}
$Text = $Text.TrimStart([char]0xFEFF)   # evtl. doppelte BOM entfernen
$Utf8Bom = New-Object System.Text.UTF8Encoding($true)
$Bytes   = $Utf8Bom.GetPreamble() + $Utf8Bom.GetBytes($Text)
$Payload = [Convert]::ToBase64String([byte[]]$Bytes)

# --- Icon einbetten (für Fenster- und Taskleisten-Icon der App) ---
$IconPayload = ''
if (-not $IconFile) {
    $Default = Join-Path (Split-Path $ScriptFull -Parent) 'mailbox.ico'
    if (Test-Path $Default) { $IconFile = $Default }
}
if ($IconFile) {
    $IconFull    = (Resolve-Path $IconFile).Path
    $IconPayload = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($IconFull))
    Write-Host "Icon: $IconFull" -ForegroundColor DarkGray
}

$Template = @'
# MailboxAccess Starter - generiert von Build-MailboxAccessExe.ps1
$Payload     = '__PAYLOAD__'
$IconPayload = '__ICON__'

try {
    $Dir = Join-Path $env:LOCALAPPDATA 'EXOMailboxAccess'
    if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir -Force | Out-Null }

    $Target = Join-Path $Dir 'Get-UserMailboxAccess.ps1'
    [System.IO.File]::WriteAllBytes($Target, [Convert]::FromBase64String($Payload))

    if ($IconPayload) {
        $IconTarget = Join-Path $Dir 'mailbox.ico'
        [System.IO.File]::WriteAllBytes($IconTarget, [Convert]::FromBase64String($IconPayload))
    }

    $PsExe   = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $ArgList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', "`"$Target`"")
    foreach ($A in $args) {
        if ("$A" -match '\s') { $ArgList += "`"$A`"" } else { $ArgList += "$A" }
    }

    Start-Process -FilePath $PsExe -ArgumentList $ArgList -WindowStyle Hidden
} catch {
    Add-Type -AssemblyName System.Windows.Forms
    [void][System.Windows.Forms.MessageBox]::Show("Start fehlgeschlagen:`n`n$($_.Exception.Message)", 'MailboxAccess', 'OK', 'Error')
}
'@

$LauncherCode = $Template.Replace('__PAYLOAD__', $Payload).Replace('__ICON__', $IconPayload)
$LauncherPath = Join-Path $env:TEMP 'MailboxAccess-Launcher.ps1'
[System.IO.File]::WriteAllText($LauncherPath, $LauncherCode, [System.Text.Encoding]::ASCII)

# --- EXE bauen ---
if (Test-Path $OutputFull) { Remove-Item $OutputFull -Force }

$Params = @{
    inputFile   = $LauncherPath
    outputFile  = $OutputFull
    noConsole   = $true
    noOutput    = $true
    x64         = $true
    title       = 'Postfachberechtigungen'
    description = 'Exchange Online Postfachberechtigungen anzeigen'
    company     = 'MailboxAccess'
    product     = 'MailboxAccess'
    version     = $Version
}
if ($IconFile) { $Params['iconFile'] = $IconFull }

Invoke-ps2exe @Params

Remove-Item $LauncherPath -Force -ErrorAction SilentlyContinue

if (Test-Path $OutputFull) {
    Write-Host ""
    Write-Host "Fertig: $OutputFull" -ForegroundColor Green
    Write-Host "Eingebettet: $ScriptFull ($([Math]::Round((Get-Item $ScriptFull).Length / 1KB, 1)) KB)" -ForegroundColor DarkGray
} else {
    Write-Error "EXE wurde nicht erstellt."
}
