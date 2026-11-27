<#
.SYNOPSIS
    Zeigt an, auf welche Exchange-Online-Postfächer ein Benutzer berechtigt ist.

.DESCRIPTION
    Version 4.0 - Berechtigungs-Cache + GUI in eigenem Thread, zusammengefasste Ansicht pro Postfach.

    FullAccess- und Kalenderrechte aller Postfächer werden EINMAL komplett eingelesen
    und lokal gespeichert (%LOCALAPPDATA%\EXOMailboxAccess\cache.xml).
    Danach dauert jede Benutzerabfrage nur noch wenige Sekunden.

    SendAs und SendOnBehalf werden immer live abgefragt.
    Rechte über mail-aktivierte Gruppen werden mit aufgelöst (Spalte "Quelle").

    Architektur GUI:
    - UI läuft in eigenem STA-Runspace (eigener Thread) -> Fenster bleibt immer bedienbar
    - Exchange-Abfragen laufen im Haupt-Runspace (dort liegt die EXO-Verbindung)
    - Kommunikation über synchronisierte Hashtable + Befehls-Queue

.EXAMPLE
    .\Get-UserMailboxAccess.ps1
    .\Get-UserMailboxAccess.ps1 -User max.mustermann@example.com
    .\Get-UserMailboxAccess.ps1 -User max.mustermann@example.com -NoGui
    .\Get-UserMailboxAccess.ps1 -User max.mustermann@example.com -NoGui -RebuildCache -IncludeCalendar
    .\Get-UserMailboxAccess.ps1 -User max.mustermann@example.com -NoGui -DeltaUpdate

.NOTES
    Benötigt: Modul ExchangeOnlineManagement (v3+), Windows (für die GUI)
    Rolle: Exchange-Administrator oder View-Only Recipients + Mail Recipients
    Delta-Update: zusätzlich Rolle 'View-Only Audit Logs' (Search-UnifiedAuditLog)
    Quellcode: src\*.ps1 - diese Datei wird mit build\Merge-Source.ps1 erzeugt
    Verschachtelte Gruppen werden nicht aufgelöst (nur direkte Mitgliedschaft).
#>

param(
    [string]$User,
    [switch]$NoGui,
    [switch]$IncludeCalendar,
    [switch]$RebuildCache,
    [int]$MaxCacheAgeHours = 1,
    [string]$ExportPath,
    [string]$CachePath = (Join-Path $env:LOCALAPPDATA 'EXOMailboxAccess\cache.xml'),
    [switch]$DeltaUpdate,
    [switch]$NoUpdateCheck
)

$ErrorActionPreference = 'Stop'

$script:AppVersion = '4.0.0'
$script:RepoSlug   = 'RazeAttack/MailboxExchangeGUI'

# Unbehandelte Fehler sichtbar machen (die EXE startet PowerShell versteckt)
$script:DiagDir = Join-Path $env:LOCALAPPDATA 'EXOMailboxAccess'
function Write-DiagLog([string]$File, [string]$Text) {
    try {
        if (-not (Test-Path $script:DiagDir)) { New-Item -ItemType Directory -Path $script:DiagDir -Force | Out-Null }
        Add-Content -Path (Join-Path $script:DiagDir $File) -Value ("{0:yyyy-MM-dd HH:mm:ss}  {1}`r`n" -f (Get-Date), $Text) -Encoding UTF8
    } catch { }
}
function Show-DiagError([string]$Text) {
    try {
        Add-Type -AssemblyName System.Windows.Forms
        [void][System.Windows.Forms.MessageBox]::Show($Text, 'MailboxAccess - Fehler', 'OK', 'Error')
    } catch { }
}
trap {
    $Info = "$($_.Exception.Message)`r`n$($_.InvocationInfo.PositionMessage)"
    Write-DiagLog 'start_fehler.log' $Info
    Show-DiagError "Unerwarteter Fehler:`n`n$Info`n`nLog: $script:DiagDir\start_fehler.log"
    break
}

# TLS 1.2 erzwingen - .NET Framework (Windows PowerShell 5.1) nutzt sonst ggf. veraltete Protokolle
[System.Net.ServicePointManager]::SecurityProtocol =
    [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Modulprüfung (statt #Requires - kompatibel mit PS2EXE)
if (-not (Get-Module ExchangeOnlineManagement -ListAvailable)) {
    $Msg = "Das PowerShell-Modul 'ExchangeOnlineManagement' ist nicht installiert.`n`n" +
           "Installation (Windows PowerShell 5.1):`n" +
           "Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    if ($NoGui) {
        Write-Error $Msg
    } else {
        Add-Type -AssemblyName System.Windows.Forms
        [void][System.Windows.Forms.MessageBox]::Show($Msg, 'Modul fehlt', 'OK', 'Error')
    }
    exit 1
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop

