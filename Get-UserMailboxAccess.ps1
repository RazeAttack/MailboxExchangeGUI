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

#region ===== Allgemein =====

$script:Cache    = $null
$script:TenantId = $null

$script:DefaultProgress = {
    param($Activity, $Status, $Percent)
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $Percent
}

function Connect-EXOIfNeeded {
    $Conn = Get-ConnectionInformation -ErrorAction SilentlyContinue |
        Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
    if (-not $Conn) {
        Write-Host "Verbinde mit Exchange Online..." -ForegroundColor Cyan

        # WAM (normale M365-Kontoauswahl) braucht ein Konsolenfenster als Parent.
        # Ohne Konsole (z. B. PS2EXE -noConsole) -> unsichtbare Konsole erzeugen.
        if (-not ('Win32.ConsoleHelper' -as [type])) {
            Add-Type -Namespace Win32 -Name ConsoleHelper -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
[DllImport("kernel32.dll")] public static extern bool AllocConsole();
[DllImport("user32.dll")]   public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@
        }
        if ([Win32.ConsoleHelper]::GetConsoleWindow() -eq [IntPtr]::Zero) {
            [void][Win32.ConsoleHelper]::AllocConsole()
            $Hwnd = [Win32.ConsoleHelper]::GetConsoleWindow()
            if ($Hwnd -ne [IntPtr]::Zero) { [void][Win32.ConsoleHelper]::ShowWindow($Hwnd, 0) }  # 0 = SW_HIDE
        }

        try {
            # Standard: WAM mit Kontoauswahl
            Connect-ExchangeOnline -ShowBanner:$false
        } catch {
            # Fallback: Browser-Login, falls WAM kein Fenster bekommt
            $HasDisableWam = (Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')
            if ($_.Exception.Message -match 'window handle' -and $HasDisableWam) {
                Connect-ExchangeOnline -ShowBanner:$false -DisableWAM
            } else {
                throw
            }
        }
        $Conn = Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
    }
    $script:TenantId = [string]$Conn.TenantId
    return $Conn
}

function New-KeySet {
    param([object[]]$Values)
    $Set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($V in $Values) {
        if ($V -and "$V".Trim()) { [void]$Set.Add("$V".Trim()) }
    }
    return ,$Set
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$Action,
        [string]$What = 'Aufruf',
        [int]$MaxAttempts = 3,
        [int]$DelaySeconds = 3,
        [scriptblock]$OnProgress
    )
    for ($Attempt = 1; $Attempt -le $MaxAttempts; $Attempt++) {
        try {
            return (& $Action)
        } catch {
            if ($Attempt -ge $MaxAttempts) { throw }
            $Wait = $DelaySeconds * $Attempt
            if ($OnProgress) {
                & $OnProgress 'Wiederholung' "$What fehlgeschlagen ($($_.Exception.Message)) - Versuch $($Attempt + 1)/$MaxAttempts in $Wait s" -1
            }
            Start-Sleep -Seconds $Wait
        }
    }
}

#endregion

#region ===== Cache =====

function Get-CalendarFolderId {
    # Liefert "adresse:\Kalender" bzw. "adresse:\Calendar" (lokalisierter Ordnername)
    param([string]$Address)
    foreach ($Name in 'Kalender', 'Calendar') {
        try {
            [void](Get-EXOMailboxFolderPermission -Identity "$($Address):\$Name" -ErrorAction Stop | Select-Object -First 1)
            return "$($Address):\$Name"
        } catch { }
    }
    $F = Get-EXOMailboxFolderStatistics -Identity $Address -FolderScope Calendar -ErrorAction Stop |
        Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object -First 1
    if ($F) { return "$($Address):\$($F.Name)" }
    throw "Kalenderordner von $Address nicht gefunden"
}

function Get-CalendarPermissions {
    param([string]$Address)

    # Schnellweg: lokalisierter Ordnername direkt (1 Aufruf statt 2)
    foreach ($Name in 'Kalender', 'Calendar') {
        try {
            return @(Get-EXOMailboxFolderPermission -Identity "$($Address):\$Name" -ErrorAction Stop)
        } catch { }
    }

    # Fallback: echten Kalendernamen ermitteln
    try {
        $F = Get-EXOMailboxFolderStatistics -Identity $Address -FolderScope Calendar -ErrorAction Stop |
            Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object -First 1
        if ($F) {
            return @(Get-EXOMailboxFolderPermission -Identity "$($Address):\$($F.Name)" -ErrorAction Stop)
        }
    } catch { }

    return @()
}

function Get-FullAccessEntries {
    # Liefert die FullAccess-Einträge eines Postfachs (wirft bei Fehler)
    param([Parameter(Mandatory)][string]$Address, [scriptblock]$OnProgress)
    $Perms = @(Invoke-WithRetry -What "FullAccess $Address" -OnProgress $OnProgress -Action {
        Get-EXOMailboxPermission -Identity $Address -ErrorAction Stop
    })
    foreach ($P in $Perms) {
        $T = [string]$P.User
        if ($P.Deny -or $T -like 'NT AUTHORITY\*' -or $T -like 'S-1-5-*') { continue }
        [PSCustomObject]@{
            Mailbox   = $Address
            Trustee   = $T
            Rights    = ($P.AccessRights -join ', ')
            Inherited = [bool]$P.IsInherited
        }
    }
}

function Get-CalendarEntries {
    # Liefert die Kalender-Einträge eines Postfachs
    param([Parameter(Mandatory)][string]$Address)
    $Skip = @('Default', 'Anonymous', 'Standard', 'Anonym')
    foreach ($C in (Get-CalendarPermissions -Address $Address)) {
        $TName = [string]$C.User
        if ($C.User -and $C.User.PSObject.Properties['DisplayName'] -and $C.User.DisplayName) {
            $TName = [string]$C.User.DisplayName
        }
        $TSmtp = ''
        try {
            if ($C.User.RecipientPrincipal) { $TSmtp = [string]$C.User.RecipientPrincipal.PrimarySmtpAddress }
        } catch { }

        $Rights = ($C.AccessRights -join ', ')
        if ($TName -in $Skip -or $Rights -eq 'None') { continue }

        [PSCustomObject]@{
            Mailbox     = $Address
            TrusteeName = $TName
            TrusteeSmtp = $TSmtp
            Rights      = $Rights
        }
    }
}

function Get-RoomDelegateEntries {
    # Buchungsstellvertreter (ResourceDelegates) eines Raum-/Gerätepostfachs
    param([Parameter(Mandatory)][string]$Address, [scriptblock]$OnProgress)
    $CP = Invoke-WithRetry -What "Buchung $Address" -OnProgress $OnProgress -Action {
        Get-CalendarProcessing -Identity $Address -ErrorAction Stop
    }
    foreach ($D in @($CP.ResourceDelegates)) {
        if (-not $D) { continue }
        $Name = [string]$D
        $Dn   = ''
        if ($D.PSObject.Properties['Name'] -and $D.Name) { $Name = [string]$D.Name }
        if ($D.PSObject.Properties['DistinguishedName'] -and $D.DistinguishedName) { $Dn = [string]$D.DistinguishedName }
        [PSCustomObject]@{ Mailbox = $Address; Delegate = $Name; DelegateDn = $Dn }
    }
}

function Test-RoomType([string]$Type) {
    return $Type -in 'RoomMailbox', 'EquipmentMailbox'
}

function Build-AccessCache {
    param(
        [switch]$IncludeCalendar,
        [scriptblock]$OnProgress = $script:DefaultProgress
    )

    & $OnProgress 'Cache aufbauen' 'Lade Postfachliste (kann 10-60 s dauern)...' -1
    $Mailboxes = @(Invoke-WithRetry -What 'Postfachliste' -MaxAttempts 4 -OnProgress $OnProgress -Action {
        Get-EXOMailbox -ResultSize Unlimited -Properties RecipientTypeDetails, DisplayName, Alias -ErrorAction Stop |
            Select-Object Name, Alias, DisplayName,
                          @{N = 'PrimarySmtpAddress';   E = { [string]$_.PrimarySmtpAddress } },
                          @{N = 'RecipientTypeDetails'; E = { [string]$_.RecipientTypeDetails } }
    })

    $Total  = [Math]::Max($Mailboxes.Count, 1)
    $Full   = [System.Collections.Generic.List[object]]::new()
    $Cal    = [System.Collections.Generic.List[object]]::new()
    $Failed = [System.Collections.Generic.List[string]]::new()
    $Dele   = [System.Collections.Generic.List[object]]::new()
    $Sw     = [System.Diagnostics.Stopwatch]::StartNew()

    $i = 0
    foreach ($MB in $Mailboxes) {
        $i++
        $Addr = $MB.PrimarySmtpAddress

        $Eta = '...'
        if ($i -gt 5) {
            $Rest = ($Sw.Elapsed.TotalSeconds / $i) * ($Total - $i)
            $Eta  = [TimeSpan]::FromSeconds($Rest).ToString('mm\:ss')
        }
        & $OnProgress 'Cache aufbauen' "$i / $Total - $Addr - Rest ca. $Eta min" ([int]($i / $Total * 100))

        # FullAccess
        try {
            foreach ($E in @(Get-FullAccessEntries -Address $Addr -OnProgress $OnProgress)) { $Full.Add($E) }
        } catch {
            if ($_.Exception -is [System.OperationCanceledException]) { throw }
            $Failed.Add("$Addr - $($_.Exception.Message)")
        }

        # Kalender: optional für alle, für Raum-/Gerätepostfächer immer
        $IsRoom = Test-RoomType $MB.RecipientTypeDetails
        if ($IncludeCalendar -or $IsRoom) {
            foreach ($E in @(Get-CalendarEntries -Address $Addr)) { $Cal.Add($E) }
        }

        # Buchungsstellvertreter (nur Raum/Gerät)
        if ($IsRoom) {
            try {
                foreach ($E in @(Get-RoomDelegateEntries -Address $Addr -OnProgress $OnProgress)) { $Dele.Add($E) }
            } catch {
                if ($_.Exception -is [System.OperationCanceledException]) { throw }
                $Failed.Add("$Addr (Buchung) - $($_.Exception.Message)")
            }
        }
    }

    $script:Cache = [PSCustomObject]@{
        Created     = Get-Date
        TenantId    = $script:TenantId
        HasCalendar = [bool]$IncludeCalendar
        HasRooms    = $true
        Mailboxes   = $Mailboxes
        FullAccess  = $Full.ToArray()
        Calendar    = $Cal.ToArray()
        RoomDelegates = $Dele.ToArray()
        Failed      = $Failed.ToArray()
        Duration    = $Sw.Elapsed.ToString('mm\:ss')
    }

    Save-AccessCache

    if ($Failed.Count -gt 0) {
        $LogFile = Join-Path (Split-Path $CachePath -Parent) 'cache_fehler.log'
        $Failed | Set-Content -Path $LogFile -Encoding UTF8
    }

    $FailText = if ($Failed.Count -gt 0) { " - $($Failed.Count) Postfächer fehlgeschlagen" } else { '' }
    & $OnProgress 'Cache aufbauen' "Fertig in $($script:Cache.Duration) min$FailText" 100
}

function Update-CacheLive {
    <#
        Liest FullAccess (und ggf. Kalender) für ausgewählte Postfächer live nach und
        ersetzt deren Einträge im Cache. Findet neu vergebene Rechte ohne kompletten Neuaufbau.
    #>
    param(
        [string[]]$Addresses,
        [switch]$IncludeSharedRooms,
        [scriptblock]$OnProgress = $script:DefaultProgress
    )
    if (-not $script:Cache) { throw 'Kein Cache vorhanden. Bitte zuerst den Cache aufbauen.' }

    $Targets = [System.Collections.Generic.List[string]]::new()
    $Seen    = New-KeySet @()
    foreach ($A in @($Addresses)) { if ($A -and $Seen.Add([string]$A)) { $Targets.Add([string]$A) } }
    if ($IncludeSharedRooms) {
        foreach ($MB in $script:Cache.Mailboxes) {
            if ([string]$MB.RecipientTypeDetails -in 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox') {
                $A = [string]$MB.PrimarySmtpAddress
                if ($A -and $Seen.Add($A)) { $Targets.Add($A) }
            }
        }
    }

    $RoomSet = New-KeySet @($script:Cache.Mailboxes | Where-Object { Test-RoomType $_.RecipientTypeDetails } | ForEach-Object { [string]$_.PrimarySmtpAddress })

    $Total  = [Math]::Max($Targets.Count, 1)
    $Failed = 0
    $Sw     = [System.Diagnostics.Stopwatch]::StartNew()
    $i      = 0
    foreach ($A in $Targets) {
        $i++
        & $OnProgress 'Live prüfen' "$i / $Total - $A" ([int]($i / $Total * 100))
        try {
            $New = @(Get-FullAccessEntries -Address $A -OnProgress $OnProgress)
            $script:Cache.FullAccess = @(@($script:Cache.FullAccess | Where-Object { $_.Mailbox -ne $A }) + $New)
            $IsRoom = $RoomSet.Contains($A)
            if ($script:Cache.HasCalendar -or $IsRoom) {
                $NewCal = @(Get-CalendarEntries -Address $A)
                $script:Cache.Calendar = @(@($script:Cache.Calendar | Where-Object { $_.Mailbox -ne $A }) + $NewCal)
            }
            if ($IsRoom) {
                $NewDel = @(Get-RoomDelegateEntries -Address $A -OnProgress $OnProgress)
                $script:Cache.RoomDelegates = @(@($script:Cache.RoomDelegates | Where-Object { $_ -and $_.Mailbox -ne $A }) + $NewDel)
            }
        } catch {
            if ($_.Exception -is [System.OperationCanceledException]) { throw }
            $Failed++
        }
    }
    Save-AccessCache

    return [PSCustomObject]@{ Checked = $Targets.Count; Failed = $Failed; Duration = $Sw.Elapsed.ToString('mm\:ss') }
}

function Save-AccessCache {
    $Dir = Split-Path $CachePath -Parent
    if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir -Force | Out-Null }
    $script:Cache | Export-Clixml -Path $CachePath -Depth 5
}

function Import-AccessCache {
    $script:Cache = $null
    if (-not (Test-Path $CachePath)) { return }
    try {
        $C = Import-Clixml -Path $CachePath
        if ($C.TenantId -eq $script:TenantId) { $script:Cache = $C }
    } catch {
        $script:Cache = $null
    }
}

function Get-CacheTimestamp {
    # Stand des Caches = letzter Komplettaufbau oder letztes Delta-Update (das spätere)
    $T = [datetime]$script:Cache.Created
    if ($script:Cache.PSObject.Properties['LastDelta'] -and $script:Cache.LastDelta) {
        $D = [datetime]$script:Cache.LastDelta
        if ($D -gt $T) { $T = $D }
    }
    return $T
}

function Get-CacheInfo {
    if (-not $script:Cache) { return 'Kein Cache vorhanden' }
    $Age = (Get-Date) - (Get-CacheTimestamp)
    $AgeText = if ($Age.TotalHours -ge 1) { '{0:N0} h' -f $Age.TotalHours } else { '{0:N0} min' -f $Age.TotalMinutes }
    $CalText = if ($script:Cache.HasCalendar) { 'mit Kalender' } else { 'ohne Kalender' }
    $FailCount = @($script:Cache.Failed | Where-Object { $_ }).Count
    $FailText  = if ($FailCount -gt 0) { " - ACHTUNG: $FailCount Postfächer fehlgeschlagen (siehe cache_fehler.log)" } else { '' }
    $DeltaText = ''
    if ($script:Cache.PSObject.Properties['LastDelta'] -and $script:Cache.LastDelta) {
        $DeltaText = ", Delta $(([datetime]$script:Cache.LastDelta).ToString('dd.MM. HH:mm'))"
    }
    return "Cache: $(@($script:Cache.Mailboxes).Count) Postfächer, $CalText, Aufbau $(([datetime]$script:Cache.Created).ToString('dd.MM. HH:mm'))$DeltaText (vor $AgeText)$FailText"
}

function Get-MailboxPickList {
    # Kompakte Postfachliste für die Auswahl im Erteilen-Dialog
    if (-not $script:Cache) { return @() }
    return @($script:Cache.Mailboxes | Where-Object { $_.PrimarySmtpAddress } | ForEach-Object {
        [PSCustomObject]@{
            Name    = [string]$(if ($_.DisplayName) { $_.DisplayName } else { $_.Name })
            Address = [string]$_.PrimarySmtpAddress
            Type    = ([string]$_.RecipientTypeDetails) -replace 'Mailbox$', ''
        }
    } | Sort-Object Name)
}

function Test-CacheHasRooms {
    return [bool]($script:Cache -and $script:Cache.PSObject.Properties['HasRooms'] -and $script:Cache.HasRooms)
}

function Test-CacheStale {
    if (-not $script:Cache) { return $true }
    if (-not (Test-CacheHasRooms)) { return $true }
    return (((Get-Date) - (Get-CacheTimestamp)).TotalHours -gt $MaxCacheAgeHours)
}

#endregion

#region ===== v4.0: Delta-Update über das Audit-Log =====

# Exchange-Admin-Operationen, die gecachte Rechte ändern (Vollzugriff, Kalender, Buchungsstellvertreter)
$script:AuditOperations = @(
    'Add-MailboxPermission', 'Remove-MailboxPermission',
    'Add-MailboxFolderPermission', 'Set-MailboxFolderPermission', 'Remove-MailboxFolderPermission',
    'Set-CalendarProcessing'
)
# Das Audit-Log liefert Einträge mit 30 min bis 24 h Verzögerung -> Zeitfenster überlappen
$script:AuditOverlapHours = 24

function Test-AuditAvailable {
    return [bool](Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue)
}

function Get-AuditIdentity($Record) {
    # Zielpostfach eines Audit-Eintrags (Parameter -Identity, sonst ObjectId)
    $Data = $null
    try { $Data = ([string]$Record.AuditData) | ConvertFrom-Json } catch { return $null }
    if (-not $Data) { return $null }
    $Id = ''
    foreach ($P in @($Data.Parameters)) {
        if ($P -and [string]$P.Name -eq 'Identity') { $Id = [string]$P.Value; break }
    }
    if (-not $Id) { $Id = [string]$Data.ObjectId }
    if (-not $Id) { return $null }
    # Ordnerrechte: "postfach:\Kalender" -> "postfach"
    $Pos = $Id.IndexOf(':\')
    if ($Pos -gt 0) { $Id = $Id.Substring(0, $Pos) }
    return $Id.Trim()
}

function Get-AuditChangedMailboxes {
    param(
        [datetime]$Since,
        [scriptblock]$OnProgress = $script:DefaultProgress
    )
    $StartUtc  = $Since.ToUniversalTime()
    $EndUtc    = (Get-Date).ToUniversalTime().AddMinutes(5)
    $SessionId = [guid]::NewGuid().ToString()
    $Ids       = New-KeySet @()
    $Count     = 0
    $Truncated = $false

    while ($true) {
        $Page = @(Invoke-WithRetry -What 'Audit-Log' -OnProgress $OnProgress -Action {
            Search-UnifiedAuditLog -StartDate $StartUtc -EndDate $EndUtc -RecordType ExchangeAdmin `
                -Operations $script:AuditOperations -SessionId $SessionId -SessionCommand ReturnLargeSet `
                -ResultSize 5000 -ErrorAction Stop
        })
        $Page = @($Page | Where-Object { $_ })
        if ($Page.Count -eq 0) { break }
        foreach ($R in $Page) {
            $Count++
            $Id = Get-AuditIdentity $R
            if ($Id) { [void]$Ids.Add($Id) }
        }
        & $OnProgress 'Delta-Update' "Audit-Log: $Count Einträge gelesen" -1
        $Last = $Page[-1]
        if ($Last.PSObject.Properties['ResultCount'] -and [int]$Last.ResultCount -gt 0 -and [int]$Last.ResultIndex -ge [int]$Last.ResultCount) { break }
        if ($Count -ge 50000) { $Truncated = $true; break }
    }
    return [PSCustomObject]@{ Identities = @($Ids); Records = $Count; Truncated = $Truncated }
}

function Update-CacheDelta {
    <#
        Schnelle Aktualisierung statt Komplettaufbau:
        1. Postfachliste neu laden (neue, gelöschte und umgewandelte Postfächer)
        2. Audit-Log seit dem letzten Stand auswerten (minus Überlappung)
        3. Nur betroffene Postfächer live nachlesen (Update-CacheLive)
    #>
    param([scriptblock]$OnProgress = $script:DefaultProgress)

    if (-not $script:Cache) { throw 'Kein Cache vorhanden. Bitte zuerst den Cache aufbauen.' }
    if (-not (Test-AuditAvailable)) {
        throw ("Das Cmdlet Search-UnifiedAuditLog ist nicht verfügbar.`n`n" +
               "Benötigt wird die Exchange-Rolle 'View-Only Audit Logs' oder 'Audit Logs' " +
               "(z. B. über die Rollengruppe 'Compliance Management') und eine aktivierte Überwachung im Tenant.`n`n" +
               "Alternative: 'Komplett neu aufbauen' oder 'Live prüfen'.")
    }

    $Sw    = [System.Diagnostics.Stopwatch]::StartNew()
    $Since = (Get-CacheTimestamp).AddHours(-$script:AuditOverlapHours)

    # 1. Postfachliste
    & $OnProgress 'Delta-Update' 'Lade Postfachliste...' -1
    $NewList = @(Invoke-WithRetry -What 'Postfachliste' -MaxAttempts 4 -OnProgress $OnProgress -Action {
        Get-EXOMailbox -ResultSize Unlimited -Properties RecipientTypeDetails, DisplayName, Alias -ErrorAction Stop |
            Select-Object Name, Alias, DisplayName,
                          @{N = 'PrimarySmtpAddress';   E = { [string]$_.PrimarySmtpAddress } },
                          @{N = 'RecipientTypeDetails'; E = { [string]$_.RecipientTypeDetails } }
    })
    if ($NewList.Count -eq 0) { throw 'Die Postfachliste ist leer - Delta-Update abgebrochen, Cache bleibt unverändert.' }

    $OldType = @{}
    foreach ($MB in @($script:Cache.Mailboxes)) { if ($MB.PrimarySmtpAddress) { $OldType[[string]$MB.PrimarySmtpAddress] = [string]$MB.RecipientTypeDetails } }
    $NewSet  = New-KeySet @($NewList | ForEach-Object { [string]$_.PrimarySmtpAddress })

    $Added       = [System.Collections.Generic.List[string]]::new()
    $TypeChanged = [System.Collections.Generic.List[string]]::new()
    foreach ($MB in $NewList) {
        $A = [string]$MB.PrimarySmtpAddress
        if (-not $A) { continue }
        if (-not $OldType.ContainsKey($A)) { $Added.Add($A) }
        elseif ($OldType[$A] -ne [string]$MB.RecipientTypeDetails) { $TypeChanged.Add($A) }
    }
    $Removed = @($OldType.Keys | Where-Object { -not $NewSet.Contains([string]$_) })

    $script:Cache.Mailboxes = $NewList
    if ($Removed.Count -gt 0) {
        $RemSet = New-KeySet $Removed
        $script:Cache.FullAccess    = @($script:Cache.FullAccess    | Where-Object { $_ -and -not $RemSet.Contains([string]$_.Mailbox) })
        $script:Cache.Calendar      = @($script:Cache.Calendar      | Where-Object { $_ -and -not $RemSet.Contains([string]$_.Mailbox) })
        $script:Cache.RoomDelegates = @($script:Cache.RoomDelegates | Where-Object { $_ -and -not $RemSet.Contains([string]$_.Mailbox) })
    }

    # 2. Audit-Log
    & $OnProgress 'Delta-Update' "Durchsuche Audit-Log ab $($Since.ToString('dd.MM. HH:mm'))..." -1
    $Audit = Get-AuditChangedMailboxes -Since $Since -OnProgress $OnProgress

    # Identity -> Adresse (Name, Alias, Anzeigename, Adresse)
    $Index = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($MB in $NewList) {
        foreach ($K in @($MB.PrimarySmtpAddress, $MB.Alias, $MB.Name, $MB.DisplayName)) {
            $Key = [string]$K
            if ($Key -and -not $Index.ContainsKey($Key)) { $Index[$Key] = [string]$MB.PrimarySmtpAddress }
        }
    }

    $Targets    = New-KeySet @()
    $Unresolved = [System.Collections.Generic.List[string]]::new()
    $Lookups    = 0
    foreach ($Id in @($Audit.Identities)) {
        $A = $null
        if ($Index.ContainsKey($Id)) { $A = $Index[$Id] }
        if (-not $A) {
            $LastPart = @($Id -split '[/\\]')[-1]
            if ($LastPart -and $Index.ContainsKey($LastPart)) { $A = $Index[$LastPart] }
        }
        if (-not $A -and $Lookups -lt 100) {
            $Lookups++
            try {
                $M = Get-EXOMailbox -Identity $Id -ErrorAction Stop
                $A = [string]$M.PrimarySmtpAddress
            } catch { $A = $null }
        }
        if ($A -and $NewSet.Contains($A)) { [void]$Targets.Add($A) } else { $Unresolved.Add($Id) }
    }
    foreach ($A in $Added)       { [void]$Targets.Add($A) }
    foreach ($A in $TypeChanged) { [void]$Targets.Add($A) }

    # 3. Betroffene Postfächer nachlesen (speichert den Cache)
    $Live = Update-CacheLive -Addresses @($Targets) -OnProgress $OnProgress

    $script:Cache | Add-Member -NotePropertyName LastDelta -NotePropertyValue (Get-Date) -Force
    Save-AccessCache

    $Log = "Delta-Update: Audit ab $($Since.ToString('dd.MM.yyyy HH:mm')), $($Audit.Records) Einträge, " +
           "$($Targets.Count) Postfächer nachgelesen (neu $($Added.Count), Typ geändert $($TypeChanged.Count), entfernt $($Removed.Count)), " +
           "Fehler $($Live.Failed), nicht zuordenbar $($Unresolved.Count)"
    if ($Unresolved.Count -gt 0) { $Log += "`r`n  Nicht zuordenbar: " + (($Unresolved | Select-Object -First 50) -join '; ') }
    Write-DiagLog 'delta.log' $Log

    return [PSCustomObject]@{
        Since       = $Since
        Records     = $Audit.Records
        Truncated   = $Audit.Truncated
        Changed     = $Targets.Count
        Added       = $Added.Count
        TypeChanged = $TypeChanged.Count
        Removed     = $Removed.Count
        Unresolved  = $Unresolved.Count
        Failed      = [int]$Live.Failed
        Duration    = $Sw.Elapsed.ToString('mm\:ss')
    }
}

function Format-DeltaReport($D) {
    $Lines = @(
        "Audit-Log ab $($D.Since.ToString('dd.MM. HH:mm')) ausgewertet: $($D.Records) Einträge",
        '',
        "Neu eingelesene Postfächer: $($D.Changed)",
        "   davon neue Postfächer: $($D.Added)",
        "   davon Typ geändert: $($D.TypeChanged)",
        "Entfernte Postfächer: $($D.Removed)",
        "Nicht zuordenbar: $($D.Unresolved)  (Details: delta.log)",
        "Fehler: $($D.Failed)",
        "Dauer: $($D.Duration) min"
    )
    if ($D.Truncated) { $Lines += ''; $Lines += 'ACHTUNG: mehr als 50.000 Audit-Einträge - bitte Cache komplett neu aufbauen.' }
    $Lines += ''
    $Lines += 'Hinweis: Das Audit-Log ist 30 min bis 24 h verzögert.'
    $Lines += "Ganz frische Änderungen findet 'Live prüfen'."
    return ($Lines -join "`n")
}

#endregion

#region ===== v4.0: Update-Prüfung (GitHub-Release) =====

# Läuft in einem eigenen Runspace, damit Start und GUI nicht warten müssen
$script:UpdateCheckScript = {
    param($Sync, [bool]$Manual)
    $Result = $null
    try {
        [System.Net.ServicePointManager]::SecurityProtocol =
            [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
        try { [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials } catch { }

        $Uri     = "https://api.github.com/repos/$($Sync.RepoSlug)/releases/latest"
        $Headers = @{ 'User-Agent' = 'MailboxAccess-UpdateCheck'; 'Accept' = 'application/vnd.github+json' }
        $Rel     = Invoke-RestMethod -Uri $Uri -Headers $Headers -TimeoutSec 8 -UseBasicParsing -ErrorAction Stop

        $Tag     = ([string]$Rel.tag_name) -replace '^[vV]', ''
        $Latest  = $null
        $Current = $null
        [void][version]::TryParse($Tag, [ref]$Latest)
        [void][version]::TryParse([string]$Sync.AppVersion, [ref]$Current)
        $Newer   = [bool]($Latest -and $Current -and ($Latest -gt $Current))
        $ExeUrl  = ''
        $Exe     = @($Rel.assets | Where-Object { [string]$_.name -like '*.exe' } | Select-Object -First 1)
        if ($Exe.Count -gt 0) { $ExeUrl = [string]$Exe[0].browser_download_url }

        $Result = [PSCustomObject]@{
            Ok = $true; Manual = $Manual; Latest = $Tag; Newer = $Newer
            Url = [string]$Rel.html_url; ExeUrl = $ExeUrl; Notes = [string]$Rel.body; Error = ''
        }
    } catch {
        $Msg = $_.Exception.Message
        if ($Msg -match '404') { $Msg = 'Auf GitHub ist noch kein Release veröffentlicht.' }
        $Result = [PSCustomObject]@{
            Ok = $false; Manual = $Manual; Latest = ''; Newer = $false
            Url = "https://github.com/$($Sync.RepoSlug)/releases"; ExeUrl = ''; Notes = ''; Error = $Msg
        }
    }
    $Sync.Update = $Result
    $Sync.UpdateVersion++
}

function Start-UpdateCheck {
    param([bool]$Manual = $false)
    $P = [powershell]::Create()
    [void]$P.AddScript($Sync.UpdateScript).AddArgument($Sync).AddArgument($Manual)
    $script:UpdatePs = $P
    [void]$P.BeginInvoke()
}

#endregion

#region ===== Abfrage =====

function New-AccessEntry {
    param($Mailbox, [string]$Berechtigung, [string]$Rechte, [string]$Quelle, [string]$Fallback)
    [PSCustomObject]@{
        Berechtigung = $Berechtigung
        Postfach     = if ($Mailbox) { [string]$Mailbox.DisplayName } else { $Fallback }
        Adresse      = if ($Mailbox) { [string]$Mailbox.PrimarySmtpAddress } else { '' }
        Typ          = if ($Mailbox) { ([string]$Mailbox.RecipientTypeDetails) -replace 'Mailbox$', '' } else { '' }
        Rechte       = $Rechte
        Quelle       = $Quelle
    }
}

function Get-MailboxAccess {
    param(
        [Parameter(Mandatory)][string]$User,
        [scriptblock]$OnProgress = $script:DefaultProgress
    )

    if (-not $script:Cache) { throw 'Kein Cache vorhanden. Bitte zuerst den Cache aufbauen.' }

    # --- Benutzer auflösen ---
    & $OnProgress 'Abfrage' 'Benutzer auflösen...' 10
    $R = Get-EXORecipient -Identity $User -Properties Alias, DisplayName, Name, DistinguishedName, ExternalDirectoryObjectId
    $OwnAddr  = [string]$R.PrimarySmtpAddress
    $UserKeys = New-KeySet @($R.PrimarySmtpAddress, $R.Name, $R.Alias, $R.DisplayName, $R.Identity, $R.ExternalDirectoryObjectId, $User)
    try {
        $Upn = (Get-EXOMailbox -Identity $User -Properties UserPrincipalName -ErrorAction Stop).UserPrincipalName
        if ($Upn) { [void]$UserKeys.Add($Upn) }
    } catch { }

    # --- Gruppen (direkte Mitgliedschaft) ---
    & $OnProgress 'Abfrage' 'Gruppenmitgliedschaften...' 25
    $DnEsc  = $R.DistinguishedName -replace "'", "''"
    $Groups = @(Get-EXORecipient -ResultSize Unlimited -Filter "Members -eq '$DnEsc'" `
                -Properties Alias, DisplayName, Name, DistinguishedName, ExternalDirectoryObjectId -ErrorAction SilentlyContinue)

    $GroupKeyMap = @{}
    foreach ($G in $Groups) {
        foreach ($K in (New-KeySet @($G.PrimarySmtpAddress, $G.Name, $G.Alias, $G.DisplayName, $G.Identity, $G.ExternalDirectoryObjectId))) {
            if (-not $GroupKeyMap.ContainsKey($K)) { $GroupKeyMap[$K] = [string]$G.DisplayName }
        }
    }

    function Get-Via([string]$Trustee) {
        if (-not $Trustee) { return $null }
        if ($UserKeys.Contains($Trustee)) { return 'direkt' }
        if ($GroupKeyMap.ContainsKey($Trustee)) { return "Gruppe: $($GroupKeyMap[$Trustee])" }
        return $null
    }

    # --- Postfach-Lookup ---
    $MbLookup = @{}
    foreach ($MB in $script:Cache.Mailboxes) {
        foreach ($K in @($MB.PrimarySmtpAddress, $MB.Name, $MB.Alias, $MB.DisplayName)) {
            if ($K -and -not $MbLookup.ContainsKey([string]$K)) { $MbLookup[[string]$K] = $MB }
        }
    }

    $Results = [System.Collections.Generic.List[object]]::new()

    # --- 1. FullAccess (aus Cache) ---
    & $OnProgress 'Abfrage' 'FullAccess (Cache)...' 40
    foreach ($E in $script:Cache.FullAccess) {
        if ($E.Mailbox -eq $OwnAddr) { continue }
        $Via = Get-Via $E.Trustee
        if (-not $Via) { continue }
        if ($E.Inherited) { $Via += ' (vererbt)' }
        $Results.Add((New-AccessEntry -Mailbox $MbLookup[$E.Mailbox] -Berechtigung 'FullAccess' -Rechte $E.Rights -Quelle $Via -Fallback $E.Mailbox))
    }

    # Trustee-Liste für Live-Abfragen (User + Gruppen)
    # -Trustee akzeptiert nur Identity / ExternalDirectoryObjectId (GUID) / UPN -> GUID bevorzugen
    $UserTrusteeId = if ($R.ExternalDirectoryObjectId) { [string]$R.ExternalDirectoryObjectId } elseif ($Upn) { $Upn } else { $OwnAddr }
    $Trustees = @([PSCustomObject]@{ Id = $UserTrusteeId; Dn = $R.DistinguishedName; Via = 'direkt'; IsGroup = $false })
    foreach ($G in $Groups) {
        $GId = if ($G.ExternalDirectoryObjectId) { [string]$G.ExternalDirectoryObjectId } else { [string]$G.Identity }
        $Trustees += [PSCustomObject]@{ Id = $GId; Dn = $G.DistinguishedName; Via = "Gruppe: $($G.DisplayName)"; IsGroup = $true }
    }

    # --- 2. SendAs (live) ---
    & $OnProgress 'Abfrage' 'SendAs (live)...' 60
    foreach ($T in $Trustees) {
        if (-not $T.Id) { continue }
        $SA = @()
        try {
            $SA = @(Get-EXORecipientPermission -Trustee $T.Id -ResultSize Unlimited -ErrorAction Stop)
        } catch {
            try {
                # Fallback: klassisches Cmdlet (REST-basiert in EXO v3)
                $SA = @(Get-RecipientPermission -Trustee $T.Id -ResultSize Unlimited -ErrorAction Stop)
            } catch {
                if (-not $T.IsGroup) { throw "SendAs-Abfrage fehlgeschlagen: $($_.Exception.Message)" }
                continue   # einzelne Gruppe nicht abfragbar -> überspringen
            }
        }
        $SA = $SA | Where-Object { $_.AccessControlType -eq 'Allow' }
        foreach ($S in $SA) {
            $MB = $MbLookup[[string]$S.Identity]
            if ($MB -and $MB.PrimarySmtpAddress -eq $OwnAddr) { continue }
            $Results.Add((New-AccessEntry -Mailbox $MB -Berechtigung 'SendAs' -Rechte ($S.AccessRights -join ', ') -Quelle $T.Via -Fallback ([string]$S.Identity)))
        }
    }

    # --- 3. SendOnBehalf (live) ---
    & $OnProgress 'Abfrage' 'SendOnBehalf (live)...' 80
    foreach ($T in $Trustees) {
        if (-not $T.Dn) { continue }
        $Esc = $T.Dn -replace "'", "''"
        $OB  = Get-EXOMailbox -ResultSize Unlimited -Filter "GrantSendOnBehalfTo -eq '$Esc'" -ErrorAction SilentlyContinue
        foreach ($O in $OB) {
            $MB = $MbLookup[[string]$O.PrimarySmtpAddress]
            $Results.Add((New-AccessEntry -Mailbox $MB -Berechtigung 'SendOnBehalf' -Rechte 'GrantSendOnBehalfTo' -Quelle $T.Via -Fallback ([string]$O.DisplayName)))
        }
    }

    # --- 4. Kalender (aus Cache; ohne Kalender-Option nur Raum-/Gerätepostfächer) ---
    & $OnProgress 'Abfrage' 'Kalender (Cache)...' 90
    foreach ($E in @($script:Cache.Calendar | Where-Object { $_ })) {
        if ($E.Mailbox -eq $OwnAddr) { continue }
        $Via = $null
        foreach ($K in @($E.TrusteeSmtp, $E.TrusteeName)) {
            $Via = Get-Via $K
            if ($Via) { break }
        }
        if (-not $Via) { continue }
        $Results.Add((New-AccessEntry -Mailbox $MbLookup[$E.Mailbox] -Berechtigung 'Kalender' -Rechte $E.Rights -Quelle $Via -Fallback $E.Mailbox))
    }

    # --- 5. Buchungsstellvertreter Raum/Gerät (aus Cache) ---
    $DnSet = New-KeySet @($R.DistinguishedName)
    foreach ($G in $Groups) { if ($G.DistinguishedName) { [void]$DnSet.Add([string]$G.DistinguishedName) } }
    foreach ($E in @($script:Cache.RoomDelegates | Where-Object { $_ })) {
        $Via = Get-Via $E.Delegate
        if (-not $Via -and $E.DelegateDn -and $DnSet.Contains([string]$E.DelegateDn)) {
            $Via = if ($E.DelegateDn -eq $R.DistinguishedName) { 'direkt' } else { 'Gruppe' }
        }
        if (-not $Via) { continue }
        $Results.Add((New-AccessEntry -Mailbox $MbLookup[$E.Mailbox] -Berechtigung 'Buchungsstellvertreter' -Rechte 'ResourceDelegates' -Quelle $Via -Fallback $E.Mailbox))
    }

    & $OnProgress 'Abfrage' "$($Results.Count) Berechtigungen gefunden" 100
    return $Results | Sort-Object Berechtigung, Postfach
}

#endregion

#region ===== v3.9: Weiterleitungen, Ordnerrechte, Gruppen, Tenant-Report =====

function Get-AcceptedDomainSet {
    if (-not $script:AcceptedDomains) {
        $script:AcceptedDomains = New-KeySet @(Get-AcceptedDomain -ErrorAction SilentlyContinue | ForEach-Object { [string]$_.DomainName })
    }
    return ,$script:AcceptedDomains
}

function Test-ExternalAddress([string]$Address) {
    if (-not $Address -or $Address -notlike '*@*') { return $false }
    $Dom = ($Address -split '@')[-1].Trim().TrimEnd(']')
    $Set = Get-AcceptedDomainSet
    if ($Set.Count -eq 0) { return $false }
    return -not $Set.Contains($Dom)
}

function Get-ForwardingInfo {
    # Postfach-Weiterleitungen + Posteingangsregeln eines Benutzers
    param(
        [Parameter(Mandatory)][string]$User,
        [scriptblock]$OnProgress = $script:DefaultProgress
    )
    & $OnProgress 'Weiterleitungen' 'Postfach lesen...' 20
    $MB   = Get-EXOMailbox -Identity $User -Properties DisplayName, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward -ErrorAction Stop
    $Addr = [string]$MB.PrimarySmtpAddress
    $Rows = [System.Collections.Generic.List[object]]::new()
    $Keep = if ($MB.DeliverToMailboxAndForward) { 'Kopie bleibt im Postfach' } else { 'KEINE Kopie im Postfach' }

    if ($MB.ForwardingSmtpAddress) {
        $T   = ([string]$MB.ForwardingSmtpAddress) -replace '^smtp:'
        $Ext = if (Test-ExternalAddress $T) { 'JA' } else { 'nein' }
        $Row = [PSCustomObject]@{ Art = 'Postfach-Weiterleitung (SMTP)'; Name = 'ForwardingSmtpAddress'; Aktiv = 'ja'; Ziel = $T; Extern = $Ext; Details = $Keep }
        $Rows.Add($Row)
    }
    if ($MB.ForwardingAddress) {
        $Row = [PSCustomObject]@{ Art = 'Postfach-Weiterleitung (Empfänger)'; Name = 'ForwardingAddress'; Aktiv = 'ja'; Ziel = [string]$MB.ForwardingAddress; Extern = 'prüfen'; Details = $Keep }
        $Rows.Add($Row)
    }

    & $OnProgress 'Weiterleitungen' 'Posteingangsregeln lesen...' 60
    $RuleErr = ''
    $Rules   = @()
    try { $Rules = @(Get-InboxRule -Mailbox $Addr -ErrorAction Stop) } catch { $RuleErr = $_.Exception.Message }

    foreach ($R in $Rules) {
        $Targets = @()
        $Smtps   = @()
        foreach ($Prop in 'ForwardTo', 'ForwardAsAttachmentTo', 'RedirectTo') {
            foreach ($V in @($R.$Prop | Where-Object { $_ })) {
                $S  = [string]$V
                $Sm = if ($S -match 'SMTP:([^\]\s]+)') { $Matches[1] } else { $S }
                $Smtps   += $Sm
                $Targets += "$Prop -> $Sm"
            }
        }
        $Act = @()
        if ($R.DeleteMessage)       { $Act += 'löschen' }
        if ($R.SoftDeleteMessage)   { $Act += 'weich löschen' }
        if ($R.MoveToFolder)        { $Act += "verschieben nach $($R.MoveToFolder)" }
        if ($R.CopyToFolder)        { $Act += "kopieren nach $($R.CopyToFolder)" }
        if ($R.MarkAsRead)          { $Act += 'als gelesen markieren' }
        if ($R.StopProcessingRules) { $Act += 'weitere Regeln stoppen' }

        $Ext = ''
        if ($Smtps.Count) {
            $Ext = if (@($Smtps | Where-Object { Test-ExternalAddress $_ }).Count) { 'JA' } else { 'nein' }
        }
        $Aktiv = if ($R.Enabled) { 'ja' } else { 'nein' }
        $Row = [PSCustomObject]@{
            Art     = 'Posteingangsregel'
            Name    = [string]$R.Name
            Aktiv   = $Aktiv
            Ziel    = ($Targets -join '; ')
            Extern  = $Ext
            Details = ($Act -join ', ')
        }
        $Rows.Add($Row)
    }

    if ($Rows.Count -eq 0) {
        $Row = [PSCustomObject]@{ Art = '-'; Name = 'Keine Weiterleitungen und keine Posteingangsregeln'; Aktiv = ''; Ziel = ''; Extern = ''; Details = '' }
        $Rows.Add($Row)
    }
    & $OnProgress 'Weiterleitungen' "$($Rows.Count) Einträge" 100

    $Info = "$($MB.DisplayName) <$Addr>"
    if ($RuleErr) { $Info += "  |  Regeln nicht lesbar: $RuleErr" }
    return [PSCustomObject]@{ Info = $Info; Rows = $Rows.ToArray() }
}

function Get-MailboxFolderRights {
    # Ordnerberechtigungen aller sichtbaren Ordner eines Postfachs
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [scriptblock]$OnProgress = $script:DefaultProgress
    )
    & $OnProgress 'Ordnerrechte' 'Ordnerliste laden...' -1
    $MB   = Get-EXOMailbox -Identity $Mailbox -Properties DisplayName -ErrorAction Stop
    $Addr = [string]$MB.PrimarySmtpAddress
    $Sep  = [string][char]0xF8FF

    $Folders = @(Get-EXOMailboxFolderStatistics -Identity $Addr -FolderScope All -ErrorAction Stop | Where-Object {
        $P = [string]$_.FolderPath
        $P -and ($P -notmatch '^/(Recoverable Items|Wiederherstellbare Elemente)') -and
        ([string]$_.FolderType -eq 'Root' -or [string]$_.ContainerClass -like 'IPF*' -or [string]$_.FolderType -eq 'User Created')
    })

    $Rows  = [System.Collections.Generic.List[object]]::new()
    $Total = [Math]::Max($Folders.Count, 1)
    $i     = 0
    foreach ($F in $Folders) {
        $i++
        $P = [string]$F.FolderPath
        if ([string]$F.FolderType -eq 'Root') {
            $Id    = "$($Addr):\"
            $Label = '(Postfach-Stamm)'
        } else {
            $Rel   = ($P -replace '/', '\').Replace($Sep, '/')
            $Id    = "$($Addr):$Rel"
            $Label = $P.Replace($Sep, '/')
        }
        & $OnProgress 'Ordnerrechte' "$i / $Total - $Label" ([int]($i / $Total * 100))

        $Perms = @()
        try {
            $Perms = @(Get-EXOMailboxFolderPermission -Identity $Id -ErrorAction Stop)
        } catch {
            if ($_.Exception -is [System.OperationCanceledException]) { throw }
            continue
        }
        foreach ($Pm in $Perms) {
            $U = [string]$Pm.User
            if ($Pm.User -and $Pm.User.PSObject.Properties['DisplayName'] -and $Pm.User.DisplayName) { $U = [string]$Pm.User.DisplayName }
            $Rg = ($Pm.AccessRights -join ', ')
            if ($Rg -eq 'None' -or $U -in 'Anonymous', 'Anonym') { continue }
            $IsDefault = $U -in 'Default', 'Standard'
            if ($IsDefault -and $Rg -eq 'AvailabilityOnly') { continue }
            $Hint = if ($IsDefault) { 'gilt für ALLE Benutzer' } else { '' }
            $Row  = [PSCustomObject]@{ Ordner = $Label; Benutzer = $U; Rechte = $Rg; Hinweis = $Hint }
            $Rows.Add($Row)
        }
    }

    if ($Rows.Count -eq 0) {
        $Row = [PSCustomObject]@{ Ordner = '-'; Benutzer = 'Keine besonderen Ordnerrechte gefunden'; Rechte = ''; Hinweis = '' }
        $Rows.Add($Row)
    }
    return [PSCustomObject]@{ Info = "$($MB.DisplayName) <$Addr> - $($Folders.Count) Ordner geprüft"; Rows = $Rows.ToArray() }
}

function ConvertTo-GroupItem($G) {
    $T = [string]$G.RecipientTypeDetails
    $Label = switch ($T) {
        'GroupMailbox'                   { 'Microsoft 365' }
        'MailUniversalSecurityGroup'     { 'Sicherheit (E-Mail)' }
        'MailUniversalDistributionGroup' { 'Verteiler' }
        'MailNonUniversalGroup'          { 'Verteiler (alt)' }
        'DynamicDistributionGroup'       { 'Dynamisch' }
        default                          { $T }
    }
    [PSCustomObject]@{ Name = [string]$G.DisplayName; Address = [string]$G.PrimarySmtpAddress; Type = $T; TypeLabel = $Label }
}

function Get-GroupInfo {
    param(
        [Parameter(Mandatory)][string]$User,
        [scriptblock]$OnProgress = $script:DefaultProgress
    )
    & $OnProgress 'Gruppen' 'Benutzer auflösen...' 10
    $R  = Get-EXORecipient -Identity $User -Properties DisplayName, DistinguishedName -ErrorAction Stop
    $Dn = ([string]$R.DistinguishedName) -replace "'", "''"

    & $OnProgress 'Gruppen' 'Mitgliedschaften lesen...' 40
    $Member = @(Get-EXORecipient -ResultSize Unlimited -Filter "Members -eq '$Dn'" -Properties DisplayName, RecipientTypeDetails -ErrorAction SilentlyContinue |
                ForEach-Object { ConvertTo-GroupItem $_ } | Sort-Object Name)

    if (-not $script:AllGroups) {
        & $OnProgress 'Gruppen' 'Alle Gruppen laden (einmalig pro Sitzung)...' -1
        $Types = @('MailUniversalDistributionGroup', 'MailUniversalSecurityGroup', 'MailNonUniversalGroup', 'GroupMailbox')
        $script:AllGroups = @(Get-EXORecipient -ResultSize Unlimited -RecipientTypeDetails $Types -Properties DisplayName, RecipientTypeDetails -ErrorAction Stop |
                              ForEach-Object { ConvertTo-GroupItem $_ } | Sort-Object Name)
    }
    & $OnProgress 'Gruppen' "$($Member.Count) Mitgliedschaften" 100
    return [PSCustomObject]@{ User = $User; UserName = [string]$R.DisplayName; Member = $Member; All = $script:AllGroups }
}

function Set-GroupMembership {
    param(
        [Parameter(Mandatory)][string]$User,
        [object[]]$Add = @(),
        [object[]]$Remove = @(),
        [string]$Admin,
        [scriptblock]$OnProgress = $script:DefaultProgress
    )
    $LogDir  = Split-Path $CachePath -Parent
    $LogFile = Join-Path $LogDir 'aenderungen.log'
    if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }

    $R   = Get-EXORecipient -Identity $User -ErrorAction Stop
    $Upn = $null
    try { $Upn = (Get-EXOMailbox -Identity $User -Properties UserPrincipalName -ErrorAction Stop).UserPrincipalName } catch { }
    $Member = if ($Upn) { [string]$Upn } else { [string]$R.PrimarySmtpAddress }

    $Ops = @()
    foreach ($G in @($Add))    { if ($G) { $Ops += [PSCustomObject]@{ Group = $G; Mode = 'Add' } } }
    foreach ($G in @($Remove)) { if ($G) { $Ops += [PSCustomObject]@{ Group = $G; Mode = 'Remove' } } }

    $Report = [System.Collections.Generic.List[object]]::new()
    $Total  = [Math]::Max($Ops.Count, 1)
    $i      = 0
    foreach ($Op in $Ops) {
        $i++
        $G = $Op.Group
        & $OnProgress 'Gruppen' "$i / $Total - $($G.Name)" ([int]($i / $Total * 100))
        $Ok  = $true
        $Msg = if ($Op.Mode -eq 'Add') { 'hinzugefügt' } else { 'entfernt' }
        try {
            if ($G.Type -eq 'GroupMailbox') {
                if ($Op.Mode -eq 'Add') {
                    Add-UnifiedGroupLinks -Identity $G.Address -LinkType Members -Links $Member -Confirm:$false -ErrorAction Stop | Out-Null
                } else {
                    Remove-UnifiedGroupLinks -Identity $G.Address -LinkType Members -Links $Member -Confirm:$false -ErrorAction Stop | Out-Null
                }
            } else {
                if ($Op.Mode -eq 'Add') {
                    Add-DistributionGroupMember -Identity $G.Address -Member $Member -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop | Out-Null
                } else {
                    Remove-DistributionGroupMember -Identity $G.Address -Member $Member -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop | Out-Null
                }
            }
        } catch {
            $Ok  = $false
            $Msg = $_.Exception.Message
        }
        $Row = [PSCustomObject]@{ Gruppe = $G.Name; Typ = $G.TypeLabel; Aktion = $Op.Mode; Ok = $Ok; Meldung = $Msg }
        $Report.Add($Row)

        $LogState = if (-not $Ok) { 'FEHLER-GRUPPE' } elseif ($Op.Mode -eq 'Add') { 'ERTEILT' } else { 'OK' }
        $Perm     = if ($Op.Mode -eq 'Add') { 'Gruppe hinzufügen' } else { 'Gruppe entfernen' }
        $Line     = '{0:yyyy-MM-dd HH:mm:ss};{1};{2};{3};{4};{5};{6}' -f (Get-Date), $Admin, $User, $G.Address, $Perm, $LogState, ($Msg -replace '[\r\n;]', ' ')
        Add-Content -Path $LogFile -Value $Line -Encoding UTF8
    }
    return $Report
}

function ConvertTo-HtmlText([string]$Text) {
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Export-TenantReport {
    # Gesamtreport aller Postfachrechte + Weiterleitungen als HTML und CSV
    param(
        [Parameter(Mandatory)][string]$Folder,
        [scriptblock]$OnProgress = $script:DefaultProgress
    )
    if (-not $script:Cache) { throw 'Kein Cache vorhanden. Bitte zuerst den Cache aufbauen.' }
    if (-not (Test-Path $Folder)) { New-Item -ItemType Directory -Path $Folder -Force | Out-Null }

    $Lk = @{}
    foreach ($M in $script:Cache.Mailboxes) {
        foreach ($K in @($M.PrimarySmtpAddress, $M.Name, $M.Alias, $M.DisplayName)) {
            if ($K -and -not $Lk.ContainsKey([string]$K)) { $Lk[[string]$K] = $M }
        }
    }

    $Rows = [System.Collections.Generic.List[object]]::new()
    $AddRow = {
        param([string]$MbKey, [string]$Perm, [string]$Trustee, [string]$Rights, [string]$Source)
        $M     = $Lk[$MbKey]
        $PName = $MbKey
        $PAddr = ''
        $PTyp  = ''
        if ($M) {
            $PName = [string]$M.DisplayName
            $PAddr = [string]$M.PrimarySmtpAddress
            $PTyp  = ([string]$M.RecipientTypeDetails) -replace 'Mailbox$'
        }
        $T     = $Lk[$Trustee]
        $TName = $Trustee
        if ($T) { $TName = [string]$T.DisplayName }
        $Obj = [PSCustomObject]@{
            Postfach     = $PName
            Adresse      = $PAddr
            Typ          = $PTyp
            Berechtigung = $Perm
            Benutzer     = $TName
            Rechte       = $Rights
            Quelle       = $Source
        }
        $Rows.Add($Obj)
    }

    & $OnProgress 'Tenant-Report' 'Vollzugriff, Kalender, Buchungsstellvertreter (Cache)...' 10
    foreach ($E in @($script:Cache.FullAccess | Where-Object { $_ })) {
        $Src = if ($E.Inherited) { 'vererbt' } else { 'direkt' }
        & $AddRow $E.Mailbox 'FullAccess' $E.Trustee $E.Rights $Src
    }
    foreach ($E in @($script:Cache.Calendar | Where-Object { $_ })) {
        $Tr = if ($E.TrusteeSmtp) { $E.TrusteeSmtp } else { $E.TrusteeName }
        & $AddRow $E.Mailbox 'Kalender' $Tr $E.Rights 'direkt'
    }
    foreach ($E in @($script:Cache.RoomDelegates | Where-Object { $_ })) {
        & $AddRow $E.Mailbox 'Buchungsstellvertreter' $E.Delegate 'ResourceDelegates' 'direkt'
    }

    & $OnProgress 'Tenant-Report' 'Senden als (alle Postfächer, live)...' -1
    $SA = @()
    try {
        $SA = @(Get-EXORecipientPermission -ResultSize Unlimited -ErrorAction Stop)
    } catch {
        $SA = @(Get-RecipientPermission -ResultSize Unlimited -ErrorAction Stop)
    }
    foreach ($S in $SA) {
        $T = [string]$S.Trustee
        if ($S.AccessControlType -ne 'Allow' -or $T -like 'NT AUTHORITY\*' -or $T -like 'S-1-5-*') { continue }
        & $AddRow ([string]$S.Identity) 'SendAs' $T ($S.AccessRights -join ', ') 'direkt'
    }

    & $OnProgress 'Tenant-Report' 'Senden im Auftrag von + Weiterleitungen (live)...' -1
    $AllMb = @(Get-EXOMailbox -ResultSize Unlimited -Properties DisplayName, RecipientTypeDetails, GrantSendOnBehalfTo, ForwardingSmtpAddress, ForwardingAddress, DeliverToMailboxAndForward -ErrorAction Stop)
    $Fwd = [System.Collections.Generic.List[object]]::new()
    foreach ($M in $AllMb) {
        $A = [string]$M.PrimarySmtpAddress
        foreach ($D in @($M.GrantSendOnBehalfTo | Where-Object { $_ })) {
            & $AddRow $A 'SendOnBehalf' ([string]$D) 'GrantSendOnBehalfTo' 'direkt'
        }
        if ($M.ForwardingSmtpAddress -or $M.ForwardingAddress) {
            $Ziel = if ($M.ForwardingSmtpAddress) { ([string]$M.ForwardingSmtpAddress) -replace '^smtp:' } else { [string]$M.ForwardingAddress }
            $Ext  = if ($M.ForwardingSmtpAddress -and (Test-ExternalAddress $Ziel)) { 'JA' } elseif ($M.ForwardingAddress) { 'prüfen' } else { 'nein' }
            $Kopie = if ($M.DeliverToMailboxAndForward) { 'ja' } else { 'nein' }
            $FRow = [PSCustomObject]@{
                Postfach      = [string]$M.DisplayName
                Adresse       = $A
                Typ           = ([string]$M.RecipientTypeDetails) -replace 'Mailbox$'
                Ziel          = $Ziel
                Extern        = $Ext
                KopieBehalten = $Kopie
            }
            $Fwd.Add($FRow)
        }
    }

    & $OnProgress 'Tenant-Report' 'Dateien schreiben...' 90
    $Stamp   = Get-Date -Format 'yyyyMMdd_HHmm'
    $CsvPath = Join-Path $Folder "Postfachrechte_Gesamt_$Stamp.csv"
    $FwdPath = Join-Path $Folder "Weiterleitungen_$Stamp.csv"
    $HtmPath = Join-Path $Folder "Postfachrechte_Report_$Stamp.html"

    $Sorted = @($Rows | Sort-Object Postfach, Berechtigung, Benutzer)
    $Sorted | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    @($Fwd) | Export-Csv -Path $FwdPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'

    # --- HTML ---
    $Counts = @($Sorted | Group-Object Berechtigung | Sort-Object Name | ForEach-Object {
        '<div class="kpi"><b>{0}</b><span>{1}</span></div>' -f $_.Count, (ConvertTo-HtmlText $_.Name)
    })
    $ExtCount = @($Fwd | Where-Object { $_.Extern -eq 'JA' }).Count
    $Counts += '<div class="kpi"><b>{0}</b><span>Weiterleitungen</span></div>' -f $Fwd.Count
    $Counts += '<div class="kpi{0}"><b>{1}</b><span>davon extern</span></div>' -f $(if ($ExtCount) { ' warn' } else { '' }), $ExtCount

    $Sb = [System.Text.StringBuilder]::new()
    foreach ($R in $Sorted) {
        $Cls = if ($R.Quelle -eq 'vererbt') { ' class="inh"' } else { '' }
        $Line = '<tr{0}><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td></tr>' -f $Cls,
            (ConvertTo-HtmlText $R.Postfach), (ConvertTo-HtmlText $R.Adresse), (ConvertTo-HtmlText $R.Typ),
            (ConvertTo-HtmlText $R.Berechtigung), (ConvertTo-HtmlText $R.Benutzer), (ConvertTo-HtmlText $R.Rechte), (ConvertTo-HtmlText $R.Quelle)
        [void]$Sb.AppendLine($Line)
    }
    $SbF = [System.Text.StringBuilder]::new()
    foreach ($F in $Fwd) {
        $Cls = if ($F.Extern -eq 'JA') { ' class="ext"' } else { '' }
        $Line = '<tr{0}><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>' -f $Cls,
            (ConvertTo-HtmlText $F.Postfach), (ConvertTo-HtmlText $F.Adresse), (ConvertTo-HtmlText $F.Typ),
            (ConvertTo-HtmlText $F.Ziel), (ConvertTo-HtmlText $F.Extern), (ConvertTo-HtmlText $F.KopieBehalten)
        [void]$SbF.AppendLine($Line)
    }

    $Template = @'
<!DOCTYPE html>
<html lang="de"><head><meta charset="utf-8"><title>Postfachrechte - Gesamtreport</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#1b1b1b;background:#f5f7fa}
h1{font-size:22px;margin:0 0 4px}h2{font-size:17px;margin:28px 0 8px}.sub{color:#666;margin-bottom:16px}
.kpis{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:16px}
.kpi{background:#fff;border:1px solid #dde3ea;border-radius:8px;padding:10px 16px;min-width:120px}
.kpi b{display:block;font-size:22px}.kpi span{color:#666;font-size:12px}.kpi.warn{border-color:#d13438;background:#fde7e9}
input{padding:7px 10px;width:360px;border:1px solid #c8ced6;border-radius:6px;margin-bottom:10px}
table{border-collapse:collapse;width:100%;background:#fff;font-size:13px}
th,td{border:1px solid #e1e5ea;padding:5px 8px;text-align:left}
th{background:#0078d4;color:#fff;cursor:pointer;position:sticky;top:0}
tr:nth-child(even){background:#fafbfc}tr.inh td{color:#888}tr.ext td{background:#fde7e9}
</style></head><body>
<h1>Postfachrechte - Gesamtreport</h1>
<div class="sub">Erstellt: {{DATE}} &middot; Cache-Stand: {{CACHE}} &middot; Grau = vererbt</div>
<div class="kpis">{{KPIS}}</div>
<h2>Berechtigungen</h2>
<input id="f1" placeholder="Filter (Postfach, Benutzer, Recht ...)" oninput="flt('t1',this.value)">
<table id="t1"><thead><tr><th>Postfach</th><th>Adresse</th><th>Typ</th><th>Berechtigung</th><th>Benutzer</th><th>Rechte</th><th>Quelle</th></tr></thead><tbody>
{{ROWS}}
</tbody></table>
<h2>Weiterleitungen</h2>
<input id="f2" placeholder="Filter" oninput="flt('t2',this.value)">
<table id="t2"><thead><tr><th>Postfach</th><th>Adresse</th><th>Typ</th><th>Ziel</th><th>Extern</th><th>Kopie behalten</th></tr></thead><tbody>
{{FWD}}
</tbody></table>
<script>
function flt(id,q){q=q.toLowerCase();document.querySelectorAll('#'+id+' tbody tr').forEach(function(r){r.style.display=r.textContent.toLowerCase().indexOf(q)>-1?'':'none';});}
document.querySelectorAll('th').forEach(function(th){th.addEventListener('click',function(){
var t=th.closest('table'),b=t.tBodies[0],i=Array.prototype.indexOf.call(th.parentNode.children,th),asc=th.dataset.asc!=='1';
var rows=Array.prototype.slice.call(b.rows);rows.sort(function(a,c){var x=a.cells[i].textContent,y=c.cells[i].textContent;return asc?x.localeCompare(y,'de'):y.localeCompare(x,'de');});
rows.forEach(function(r){b.appendChild(r);});th.dataset.asc=asc?'1':'0';});});
</script></body></html>
'@
    $CacheStamp = ([datetime]$script:Cache.Created).ToString('dd.MM.yyyy HH:mm')
    $Html = $Template.Replace('{{DATE}}', (Get-Date).ToString('dd.MM.yyyy HH:mm')).Replace('{{CACHE}}', $CacheStamp).Replace('{{KPIS}}', ($Counts -join '')).Replace('{{ROWS}}', $Sb.ToString()).Replace('{{FWD}}', $SbF.ToString())
    [System.IO.File]::WriteAllText($HtmPath, $Html, [System.Text.UTF8Encoding]::new($true))

    & $OnProgress 'Tenant-Report' 'Fertig' 100
    return [PSCustomObject]@{ Html = $HtmPath; Csv = $CsvPath; FwdCsv = $FwdPath; Count = $Sorted.Count; FwdCount = $Fwd.Count; ExtCount = $ExtCount }
}

#endregion

function Merge-ChangeIntoResults {
    <#
        Übernimmt erfolgreiche Änderungen sofort in die Ergebnisliste.
        Exchange liefert "Senden als"/"Senden im Auftrag von" teils verzögert - so ist die Liste trotzdem aktuell.
    #>
    param(
        [object[]]$Results,
        [object[]]$Report,
        [string]$Mode,
        [object[]]$Items,
        [hashtable]$Options = @{}
    )
    $List = [System.Collections.Generic.List[object]]::new()
    foreach ($X in @($Results)) { if ($X) { $List.Add($X) } }

    $MbLk = @{}
    foreach ($M in @($script:Cache.Mailboxes)) { if ($M.PrimarySmtpAddress) { $MbLk[[string]$M.PrimarySmtpAddress] = $M } }
    $RightsMap = @{ FullAccess = 'FullAccess'; SendAs = 'SendAs'; SendOnBehalf = 'GrantSendOnBehalfTo'; Buchungsstellvertreter = 'ResourceDelegates' }

    foreach ($E in @($Report | Where-Object { $_ -and $_.Ok })) {
        $A = [string]$E.Adresse
        $P = [string]$E.Berechtigung
        if ($Mode -eq 'Remove') {
            $Keep = @($List | Where-Object { -not (([string]$_.Adresse) -eq $A -and ([string]$_.Berechtigung) -eq $P -and ([string]$_.Quelle) -eq 'direkt') })
            $List = [System.Collections.Generic.List[object]]::new()
            foreach ($K in $Keep) { $List.Add($K) }
        } else {
            $Exists = @($List | Where-Object { ([string]$_.Adresse) -eq $A -and ([string]$_.Berechtigung) -eq $P -and ([string]$_.Quelle) -eq 'direkt' }).Count
            if ($Exists) { continue }
            $Rights = $RightsMap[$P]
            if ($P -eq 'Kalender') {
                $It = @($Items | Where-Object { ([string]$_.Adresse) -eq $A } | Select-Object -First 1)
                $Rights = 'Reviewer'
                if ($Options.CalendarRole) { $Rights = [string]$Options.CalendarRole }
                if ($It -and $It[0].PSObject.Properties['CalendarRole'] -and $It[0].CalendarRole) { $Rights = [string]$It[0].CalendarRole }
            }
            $New = New-AccessEntry -Mailbox $MbLk[$A] -Berechtigung $P -Rechte $Rights -Quelle 'direkt' -Fallback $A
            $List.Add($New)
        }
    }
    return @($List | Sort-Object Berechtigung, Postfach)
}

#region ===== Postfach-Sicht: wer hat Zugriff? =====

function Get-MailboxAccessors {
    <#
        Liefert alle Benutzer/Gruppen mit Rechten auf ein Postfach.
        FullAccess, Kalender, Buchungsstellvertreter aus dem Cache; SendAs und SendOnBehalf live.
    #>
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [scriptblock]$OnProgress = $script:DefaultProgress
    )
    if (-not $script:Cache) { throw 'Kein Cache vorhanden. Bitte zuerst den Cache aufbauen.' }

    & $OnProgress 'Postfach-Zugriff' 'Postfach auflösen...' 10
    $MB   = Get-EXOMailbox -Identity $Mailbox -Properties DisplayName, RecipientTypeDetails, GrantSendOnBehalfTo -ErrorAction Stop
    $Addr = [string]$MB.PrimarySmtpAddress
    $Type = ([string]$MB.RecipientTypeDetails) -replace 'Mailbox$', ''

    $Lk = @{}
    foreach ($M in $script:Cache.Mailboxes) {
        foreach ($K in @($M.PrimarySmtpAddress, $M.Name, $M.Alias, $M.DisplayName)) {
            if ($K -and -not $Lk.ContainsKey([string]$K)) { $Lk[[string]$K] = $M }
        }
    }

    $Entries = [System.Collections.Generic.List[object]]::new()
    $AddEntry = {
        param([string]$Trustee, [string]$Perm, [string]$Rights, [string]$Source, [string]$Smtp)
        if (-not $Trustee -and -not $Smtp) { return }
        $M = $null
        foreach ($K in @($Smtp, $Trustee)) { if ($K -and $Lk.ContainsKey($K)) { $M = $Lk[$K]; break } }
        # Werte vorab berechnen (Komma in -replace innerhalb von .Add(...) würde als Argumenttrenner gelesen)
        $VName = $Trustee
        $VAddr = ''
        $VTyp  = ''
        if ($M) {
            $VName = [string]$M.DisplayName
            $VAddr = [string]$M.PrimarySmtpAddress
            $VTyp  = ([string]$M.RecipientTypeDetails) -replace 'Mailbox$'
        } elseif ($Smtp) {
            $VAddr = $Smtp
        } elseif ($Trustee -like '*@*') {
            $VAddr = $Trustee
        }
        $Obj = [PSCustomObject]@{
            Benutzer     = $VName
            Adresse      = $VAddr
            Typ          = $VTyp
            Berechtigung = $Perm
            Rechte       = $Rights
            Quelle       = $Source
        }
        $Entries.Add($Obj)
    }

    # FullAccess (Cache)
    & $OnProgress 'Postfach-Zugriff' 'Vollzugriff (Cache)...' 30
    foreach ($E in @($script:Cache.FullAccess | Where-Object { $_ -and $_.Mailbox -eq $Addr })) {
        & $AddEntry $E.Trustee 'FullAccess' $E.Rights $(if ($E.Inherited) { 'vererbt' } else { 'direkt' }) ''
    }

    # SendAs (live)
    & $OnProgress 'Postfach-Zugriff' 'Senden als (live)...' 55
    $SA = @()
    try {
        $SA = @(Get-EXORecipientPermission -Identity $Addr -ResultSize Unlimited -ErrorAction Stop)
    } catch {
        $SA = @(Get-RecipientPermission -Identity $Addr -ResultSize Unlimited -ErrorAction Stop)
    }
    foreach ($S in $SA) {
        $T = [string]$S.Trustee
        if ($S.AccessControlType -ne 'Allow' -or $T -like 'NT AUTHORITY\*' -or $T -like 'S-1-5-*') { continue }
        & $AddEntry $T 'SendAs' ($S.AccessRights -join ', ') 'direkt' ''
    }

    # SendOnBehalf (live, aus dem Postfach)
    & $OnProgress 'Postfach-Zugriff' 'Senden im Auftrag von...' 75
    foreach ($D in @($MB.GrantSendOnBehalfTo | Where-Object { $_ })) {
        & $AddEntry ([string]$D) 'SendOnBehalf' 'GrantSendOnBehalfTo' 'direkt' ''
    }

    # Kalender + Buchungsstellvertreter (Cache)
    & $OnProgress 'Postfach-Zugriff' 'Kalender (Cache)...' 90
    foreach ($E in @($script:Cache.Calendar | Where-Object { $_ -and $_.Mailbox -eq $Addr })) {
        & $AddEntry $E.TrusteeName 'Kalender' $E.Rights 'direkt' $E.TrusteeSmtp
    }
    foreach ($E in @($script:Cache.RoomDelegates | Where-Object { $_ -and $_.Mailbox -eq $Addr })) {
        & $AddEntry $E.Delegate 'Buchungsstellvertreter' 'ResourceDelegates' 'direkt' ''
    }

    & $OnProgress 'Postfach-Zugriff' "$($Entries.Count) Einträge" 100
    return [PSCustomObject]@{
        Name            = [string]$MB.DisplayName
        Address         = $Addr
        Type            = $Type
        Entries         = $Entries.ToArray()
        CalendarInCache = [bool]($script:Cache.HasCalendar -or $Type -in 'Room', 'Equipment')
    }
}

#endregion

#region ===== Berechtigungen erteilen =====

function Grant-UserMailboxAccess {
    <#
        Erteilt einem Benutzer Berechtigungen auf Postfächer.
        Items: @{ Adresse; Postfach; IsRoom; Perms = @('FullAccess','SendAs','SendOnBehalf','Kalender','Buchungsstellvertreter') }
        Options: @{ AutoMapping = $true; CalendarRole = 'Reviewer' }
        Kalender und Buchungsstellvertreter werden nur auf Raum-/Gerätepostfächer angewendet.
    #>
    param(
        [Parameter(Mandatory)][string]$User,
        [Parameter(Mandatory)][object[]]$Items,
        [hashtable]$Options = @{},
        [string]$Admin,
        [scriptblock]$OnProgress = $script:DefaultProgress
    )

    $LogDir  = Split-Path $CachePath -Parent
    $LogFile = Join-Path $LogDir 'aenderungen.log'
    if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }

    $AutoMap = if ($Options.ContainsKey('AutoMapping')) { [bool]$Options.AutoMapping } else { $true }
    $CalRole = if ($Options.CalendarRole) { [string]$Options.CalendarRole } else { 'Reviewer' }

    & $OnProgress 'Erteilen' 'Benutzer auflösen...' -1
    $R   = Get-EXORecipient -Identity $User -Properties Alias, DisplayName, Name, DistinguishedName, ExternalDirectoryObjectId -ErrorAction Stop
    $Upn = $null
    try { $Upn = (Get-EXOMailbox -Identity $User -Properties UserPrincipalName -ErrorAction Stop).UserPrincipalName } catch { }
    $Trustee = if ($Upn) { [string]$Upn } else { [string]$R.PrimarySmtpAddress }
    $Keys    = New-KeySet @($R.PrimarySmtpAddress, $R.Name, $R.Alias, $R.DisplayName, $R.DistinguishedName, $Upn, $User)

    $Total  = [Math]::Max([int](@($Items | ForEach-Object { @($_.Perms).Count }) | Measure-Object -Sum).Sum, 1)
    $Done   = 0
    $Report = [System.Collections.Generic.List[object]]::new()

    foreach ($It in $Items) {
        foreach ($P in @($It.Perms)) {
            $Done++
            & $OnProgress 'Erteilen' "$P - $($It.Adresse)" ([int]($Done / $Total * 100))

            $Ok  = $true
            $Msg = 'erteilt'
            try {
                if ($P -eq 'Buchungsstellvertreter' -and -not $It.IsRoom) {
                    throw 'nur für Raum-/Gerätepostfächer möglich'
                }
                $Role = if ($It.PSObject.Properties['CalendarRole'] -and $It.CalendarRole) { [string]$It.CalendarRole } else { $CalRole }
                $RoleArr = @($Role -split ',\s*' | Where-Object { $_ })
                switch ($P) {
                    'FullAccess' {
                        Add-MailboxPermission -Identity $It.Adresse -User $Trustee -AccessRights FullAccess `
                            -InheritanceType All -AutoMapping $AutoMap -Confirm:$false -ErrorAction Stop | Out-Null
                        if (-not $AutoMap) { $Msg = 'erteilt (ohne AutoMapping)' }
                    }
                    'SendAs' {
                        Add-RecipientPermission -Identity $It.Adresse -Trustee $Trustee -AccessRights SendAs `
                            -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    'SendOnBehalf' {
                        Set-Mailbox -Identity $It.Adresse -GrantSendOnBehalfTo @{ Add = $Trustee } `
                            -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    'Kalender' {
                        $Folder = Get-CalendarFolderId -Address $It.Adresse
                        try {
                            Add-MailboxFolderPermission -Identity $Folder -User $Trustee -AccessRights $RoleArr `
                                -Confirm:$false -ErrorAction Stop | Out-Null
                        } catch {
                            # Eintrag existiert bereits -> Rolle ändern
                            Set-MailboxFolderPermission -Identity $Folder -User $Trustee -AccessRights $RoleArr `
                                -Confirm:$false -ErrorAction Stop | Out-Null
                            $Msg = 'geändert (Eintrag war vorhanden)'
                        }
                        $Msg += " - $Role"
                    }
                    'Buchungsstellvertreter' {
                        $CP  = Get-CalendarProcessing -Identity $It.Adresse -ErrorAction Stop
                        $All = @($CP.ResourceDelegates | Where-Object { $_ })
                        $Has = @($All | Where-Object {
                            $N  = [string]$_
                            $Nm = if ($_.PSObject.Properties['Name']) { [string]$_.Name } else { $N }
                            $Dn = if ($_.PSObject.Properties['DistinguishedName']) { [string]$_.DistinguishedName } else { '' }
                            $Keys.Contains($N) -or $Keys.Contains($Nm) -or ($Dn -and $Keys.Contains($Dn))
                        })
                        if ($Has.Count) {
                            $Msg = 'war bereits Buchungsstellvertreter'
                        } else {
                            $New = @($All | ForEach-Object {
                                if ($_.PSObject.Properties['DistinguishedName'] -and $_.DistinguishedName) { [string]$_.DistinguishedName } else { [string]$_ }
                            }) + @([string]$R.DistinguishedName)
                            Set-CalendarProcessing -Identity $It.Adresse -ResourceDelegates $New -Confirm:$false -ErrorAction Stop | Out-Null
                        }
                    }
                    default { throw "Unbekannte Berechtigung '$P'" }
                }
            } catch {
                $Ok  = $false
                $Msg = $_.Exception.Message
            }

            $Report.Add([PSCustomObject]@{
                Postfach     = $It.Postfach
                Adresse      = $It.Adresse
                Berechtigung = $P
                Ok           = $Ok
                Meldung      = $Msg
            })

            $LogState = if ($Ok) { 'ERTEILT' } else { 'FEHLER-ERTEILEN' }
            $Line     = '{0:yyyy-MM-dd HH:mm:ss};{1};{2};{3};{4};{5};{6}' -f (Get-Date), $Admin, $User, $It.Adresse, $P, $LogState, ($Msg -replace '[\r\n;]', ' ')
            Add-Content -Path $LogFile -Value $Line -Encoding UTF8

            # Cache ergänzen, damit das neue Recht sofort angezeigt wird
            if ($Ok -and $script:Cache) {
                switch ($P) {
                    'FullAccess' {
                        $script:Cache.FullAccess = @(@($script:Cache.FullAccess | Where-Object {
                            -not ($_.Mailbox -eq $It.Adresse -and $Keys.Contains([string]$_.Trustee) -and -not $_.Inherited)
                        }) + [PSCustomObject]@{ Mailbox = $It.Adresse; Trustee = $Trustee; Rights = 'FullAccess'; Inherited = $false })
                    }
                    'Kalender' {
                        $script:Cache.Calendar = @(@($script:Cache.Calendar | Where-Object {
                            -not ($_.Mailbox -eq $It.Adresse -and ($Keys.Contains([string]$_.TrusteeSmtp) -or $Keys.Contains([string]$_.TrusteeName)))
                        }) + [PSCustomObject]@{ Mailbox = $It.Adresse; TrusteeName = [string]$R.DisplayName; TrusteeSmtp = [string]$R.PrimarySmtpAddress; Rights = $Role })
                    }
                    'Buchungsstellvertreter' {
                        $script:Cache.RoomDelegates = @(@($script:Cache.RoomDelegates | Where-Object {
                            $_ -and -not ($_.Mailbox -eq $It.Adresse -and ($Keys.Contains([string]$_.Delegate) -or ($_.DelegateDn -and $Keys.Contains([string]$_.DelegateDn))))
                        }) + [PSCustomObject]@{ Mailbox = $It.Adresse; Delegate = [string]$R.Name; DelegateDn = [string]$R.DistinguishedName })
                    }
                }
            }
        }
    }

    if ($script:Cache) { Save-AccessCache }
    return $Report
}

#endregion

#region ===== Berechtigungen entfernen =====

function Remove-UserMailboxAccess {
    <#
        Entfernt direkte Berechtigungen eines Benutzers.
        Items: @{ Adresse; Postfach; Perms = @('FullAccess','SendAs','SendOnBehalf','Kalender','Buchungsstellvertreter') }
        Jede Aktion wird in %LOCALAPPDATA%\EXOMailboxAccess\aenderungen.log protokolliert.
    #>
    param(
        [Parameter(Mandatory)][string]$User,
        [Parameter(Mandatory)][object[]]$Items,
        [string]$Admin,
        [scriptblock]$OnProgress = $script:DefaultProgress
    )

    $LogDir  = Split-Path $CachePath -Parent
    $LogFile = Join-Path $LogDir 'aenderungen.log'
    if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }

    # Benutzer auflösen (UPN als Trustee, Schlüssel für die Cache-Bereinigung)
    & $OnProgress 'Entfernen' 'Benutzer auflösen...' -1
    $R    = Get-EXORecipient -Identity $User -Properties Alias, DisplayName, Name, DistinguishedName, ExternalDirectoryObjectId
    $Keys = New-KeySet @($R.PrimarySmtpAddress, $R.Name, $R.Alias, $R.DisplayName, $R.Identity, $R.ExternalDirectoryObjectId, $R.DistinguishedName, $User)
    $Upn  = $null
    try {
        $Upn = (Get-EXOMailbox -Identity $User -Properties UserPrincipalName -ErrorAction Stop).UserPrincipalName
        if ($Upn) { [void]$Keys.Add($Upn) }
    } catch { }
    $Trustee = if ($Upn) { [string]$Upn } else { [string]$R.PrimarySmtpAddress }

    $Total  = [Math]::Max([int](@($Items | ForEach-Object { @($_.Perms).Count }) | Measure-Object -Sum).Sum, 1)
    $Done   = 0
    $Report = [System.Collections.Generic.List[object]]::new()

    foreach ($It in $Items) {
        foreach ($P in @($It.Perms)) {
            $Done++
            & $OnProgress 'Entfernen' "$P - $($It.Adresse)" ([int]($Done / $Total * 100))

            $Ok  = $true
            $Msg = 'entfernt'
            try {
                switch ($P) {
                    'FullAccess' {
                        Remove-MailboxPermission -Identity $It.Adresse -User $Trustee -AccessRights FullAccess `
                            -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    'SendAs' {
                        Remove-RecipientPermission -Identity $It.Adresse -Trustee $Trustee -AccessRights SendAs `
                            -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    'SendOnBehalf' {
                        Set-Mailbox -Identity $It.Adresse -GrantSendOnBehalfTo @{ Remove = $Trustee } `
                            -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    'Kalender' {
                        $Folder = Get-CalendarFolderId -Address $It.Adresse
                        Remove-MailboxFolderPermission -Identity $Folder -User $Trustee -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    'Buchungsstellvertreter' {
                        $CP   = Get-CalendarProcessing -Identity $It.Adresse -ErrorAction Stop
                        $All  = @($CP.ResourceDelegates | Where-Object { $_ })
                        $Keep = @($All | Where-Object {
                            $N  = [string]$_
                            $Nm = if ($_.PSObject.Properties['Name']) { [string]$_.Name } else { $N }
                            $Dn = if ($_.PSObject.Properties['DistinguishedName']) { [string]$_.DistinguishedName } else { '' }
                            -not ($Keys.Contains($N) -or $Keys.Contains($Nm) -or ($Dn -and $Keys.Contains($Dn)))
                        })
                        if ($Keep.Count -eq $All.Count) { throw 'Benutzer ist kein Buchungsstellvertreter (mehr)' }
                        $New = @($Keep | ForEach-Object {
                            if ($_.PSObject.Properties['DistinguishedName'] -and $_.DistinguishedName) { [string]$_.DistinguishedName } else { [string]$_ }
                        })
                        if ($New.Count -eq 0) { $New = $null }
                        Set-CalendarProcessing -Identity $It.Adresse -ResourceDelegates $New -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    default { throw "Unbekannte Berechtigung '$P'" }
                }
            } catch {
                $Ok  = $false
                $Msg = $_.Exception.Message
            }

            $Report.Add([PSCustomObject]@{
                Postfach     = $It.Postfach
                Adresse      = $It.Adresse
                Berechtigung = $P
                Ok           = $Ok
                Meldung      = $Msg
            })

            # Protokoll
            $LogState = if ($Ok) { 'OK' } else { 'FEHLER' }
            $Line     = '{0:yyyy-MM-dd HH:mm:ss};{1};{2};{3};{4};{5};{6}' -f (Get-Date), $Admin, $User, $It.Adresse, $P, $LogState, ($Msg -replace '[\r\n;]', ' ')
            Add-Content -Path $LogFile -Value $Line -Encoding UTF8

            # Cache bereinigen - sonst taucht das Recht bei der nächsten Abfrage wieder auf
            if ($Ok -and $script:Cache) {
                switch ($P) {
                    'FullAccess' {
                        $script:Cache.FullAccess = @($script:Cache.FullAccess | Where-Object {
                            -not ($_.Mailbox -eq $It.Adresse -and $Keys.Contains([string]$_.Trustee) -and -not $_.Inherited)
                        })
                    }
                    'Kalender' {
                        $script:Cache.Calendar = @($script:Cache.Calendar | Where-Object {
                            -not ($_.Mailbox -eq $It.Adresse -and ($Keys.Contains([string]$_.TrusteeSmtp) -or $Keys.Contains([string]$_.TrusteeName)))
                        })
                    }
                    'Buchungsstellvertreter' {
                        $script:Cache.RoomDelegates = @($script:Cache.RoomDelegates | Where-Object {
                            $_ -and -not ($_.Mailbox -eq $It.Adresse -and ($Keys.Contains([string]$_.Delegate) -or ($_.DelegateDn -and $Keys.Contains([string]$_.DelegateDn))))
                        })
                    }
                }
            }
        }
    }

    if ($script:Cache) { Save-AccessCache }
    return $Report
}

#endregion

#region ===== Konsolenausgabe =====

function Show-AccessConsole {
    param($Results, [string]$User)

    $Colors = @{ FullAccess = 'Cyan'; SendAs = 'Yellow'; SendOnBehalf = 'Magenta'; Kalender = 'Green'; Buchungsstellvertreter = 'DarkCyan' }

    Write-Host ""
    Write-Host ("=" * 90) -ForegroundColor DarkGray
    Write-Host " Postfachberechtigungen für: $User" -ForegroundColor White
    Write-Host " $(Get-CacheInfo)" -ForegroundColor DarkGray
    Write-Host ("=" * 90) -ForegroundColor DarkGray

    if (-not $Results -or @($Results).Count -eq 0) {
        Write-Host "`n Keine Berechtigungen gefunden.`n" -ForegroundColor Green
        return
    }

    foreach ($Group in ($Results | Group-Object Berechtigung)) {
        $Color = if ($Colors[$Group.Name]) { $Colors[$Group.Name] } else { 'White' }
        Write-Host ""
        Write-Host " >> $($Group.Name) ($($Group.Count))" -ForegroundColor $Color
        Write-Host (" " + "-" * 88) -ForegroundColor DarkGray
        $Group.Group | Format-Table Postfach, Adresse, Typ, Rechte, Quelle -AutoSize | Out-String | Write-Host
    }

    Write-Host ("=" * 90) -ForegroundColor DarkGray
    $Summary = ($Results | Group-Object Berechtigung | ForEach-Object { "$($_.Name): $($_.Count)" }) -join '  |  '
    Write-Host " Gesamt: $(@($Results).Count)   ->   $Summary" -ForegroundColor White
    Write-Host ("=" * 90) -ForegroundColor DarkGray
    Write-Host ""
}

#endregion

#region ===== Start =====

$ConnInfo = Connect-EXOIfNeeded
Import-AccessCache

if ($NoGui) {
    # Konsole: optional Delta-Update über das Audit-Log
    if ($DeltaUpdate -and $script:Cache -and (Test-CacheHasRooms) -and -not $RebuildCache) {
        Write-Host "Delta-Update über das Audit-Log..." -ForegroundColor Yellow
        $D = Update-CacheDelta
        Write-Progress -Activity 'Delta-Update' -Completed
        Write-Host (Format-DeltaReport $D) -ForegroundColor DarkGray
        Write-Host ''
    }
    # Konsole: automatischer Neuaufbau erst ab 24 h (GUI fragt bereits ab MaxCacheAgeHours)
    $TooOld    = (-not $script:Cache) -or (((Get-Date) - (Get-CacheTimestamp)).TotalHours -gt [Math]::Max($MaxCacheAgeHours, 24))
    $NeedBuild = $RebuildCache -or $TooOld -or (-not (Test-CacheHasRooms)) -or ($IncludeCalendar -and -not $script:Cache.HasCalendar)
    if ($NeedBuild) {
        Write-Host "Baue Cache auf (einmalig, dauert einige Minuten)..." -ForegroundColor Yellow
        Build-AccessCache -IncludeCalendar:$IncludeCalendar
        Write-Progress -Activity 'Cache aufbauen' -Completed
    }
    if (-not $User) { $User = Read-Host "Benutzer (UPN/E-Mail)" }
    $Res = Get-MailboxAccess -User $User
    Write-Progress -Activity 'Abfrage' -Completed
    Show-AccessConsole -Results $Res -User $User
    if ($ExportPath -and $Res) {
        $Res | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        Write-Host "Export: $ExportPath" -ForegroundColor Green
    }
    return
}

#endregion

#region ===== GUI: gemeinsamer Zustand =====

# Rechteprüfung: EXO lädt nur Cmdlets, die die RBAC-Rollen des Admins erlauben
$CanWrite = [bool]((Get-Command Add-MailboxPermission    -ErrorAction SilentlyContinue) -and
                   (Get-Command Remove-MailboxPermission -ErrorAction SilentlyContinue) -and
                   (Get-Command Add-RecipientPermission  -ErrorAction SilentlyContinue))

$Sync = [hashtable]::Synchronized(@{
    Queue            = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    Busy             = $false
    Closed           = $false
    Cancel           = $false
    StatusText       = 'Bereit'
    Percent          = 0
    Results          = $null
    ResultUser       = ''
    ResultVersion    = 0
    ErrorText        = ''
    ErrorVersion     = 0
    ReportText       = ''
    ReportFailed     = $false
    ReportVersion    = 0
    ReportTitle      = ''
    CacheInfo        = (Get-CacheInfo)
    CacheStale       = (Test-CacheStale)
    CacheHasRooms    = (Test-CacheHasRooms)
    HasCache         = [bool]$script:Cache
    HasCalendar      = [bool]($script:Cache -and $script:Cache.HasCalendar)
    MailboxList      = (Get-MailboxPickList)
    MbResult         = $null
    MbVersion        = 0
    Extra            = $null
    ExtraVersion     = 0
    GroupData        = $null
    GroupVersion     = 0
    RecheckUser      = ''
    RecheckAddrs     = @()
    RecheckVersion   = 0
    DataDir          = (Split-Path $CachePath -Parent)
    CanWrite         = $CanWrite
    CacheVersion     = 0
    ConnUpn          = [string]$ConnInfo.UserPrincipalName
    IconPath         = $null
    InitialUser      = $User
    InitialCalendar  = [bool]$IncludeCalendar
    RebuildCache     = [bool]$RebuildCache
    MaxCacheAgeHours = $MaxCacheAgeHours
    CanAudit         = (Test-AuditAvailable)
    AppVersion       = $script:AppVersion
    RepoSlug         = $script:RepoSlug
    UpdateScript     = $script:UpdateCheckScript.ToString()
    Update           = $null
    UpdateVersion    = 0
})

# Fortschritt aus dem Arbeits-Thread an die GUI melden (+ Abbruch prüfen)
$SyncProgress = {
    param($Activity, $Text, $Percent)
    $Sync.StatusText = "$Activity - $Text"
    $Sync.Percent    = [int]$Percent
    if ($Sync.Cancel) { throw [System.OperationCanceledException]::new('Vom Benutzer abgebrochen') }
}

#endregion

#region ===== GUI: Oberfläche (läuft in eigenem Thread) =====

$UiScript = {
    $UiLog = Join-Path $env:LOCALAPPDATA 'EXOMailboxAccess\ui_fehler.log'
    trap {
        try {
            Add-Content -Path $UiLog -Encoding UTF8 -Value ("{0:yyyy-MM-dd HH:mm:ss}  {1}`r`n{2}`r`n" -f (Get-Date), $_.Exception.Message, $_.InvocationInfo.PositionMessage)
        } catch { }
        continue
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $Font     = New-Object System.Drawing.Font('Segoe UI', 9)
    $FontBold = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

    $RowColors = @{
        FullAccess   = [System.Drawing.Color]::FromArgb(221, 235, 255)
        SendAs       = [System.Drawing.Color]::FromArgb(255, 243, 205)
        SendOnBehalf = [System.Drawing.Color]::FromArgb(243, 225, 255)
        Kalender     = [System.Drawing.Color]::FromArgb(221, 245, 225)
        Buchungsstellvertreter = [System.Drawing.Color]::FromArgb(214, 240, 240)
    }

    $State = @{ ResultVersion = 0; ErrorVersion = 0; ReportVersion = 0; CacheVersion = -1; Raw = @(); Shown = @()
                Mb = $null; AutoSrc = $null; AutoListRef = $null; Recent = @()
                ExtraVersion = 0; GroupVersion = 0; RecheckVersion = 0; RecheckAt = $null; RecheckUser = ''
                UpdateVersion = 0; UpdateUrl = '' }
    $RecentFile = Join-Path $Sync.DataDir 'recent.txt'
    try { if (Test-Path $RecentFile) { $State.Recent = @(Get-Content $RecentFile -Encoding UTF8 | Where-Object { $_ } | Select-Object -First 15) } } catch { }

    # --- Form ---
    $Form               = New-Object System.Windows.Forms.Form
    $Form.Text          = "Exchange Online - Postfachberechtigungen v$($Sync.AppVersion)   [$($Sync.ConnUpn)]"
    $Form.Size          = New-Object System.Drawing.Size(1400, 790)
    $Form.MinimumSize   = New-Object System.Drawing.Size(1380, 500)
    $Form.StartPosition = 'CenterScreen'
    $Form.Font          = $Font
    if ($Sync.IconPath) {
        try { $Form.Icon = New-Object System.Drawing.Icon($Sync.IconPath) } catch { }
    }

    # --- DataTable ---
    $Table = New-Object System.Data.DataTable
    foreach ($Col in 'Postfach', 'Berechtigung', 'Adresse', 'Typ', 'Rechte', 'Quelle') { [void]$Table.Columns.Add($Col) }

    # --- Grid ---
    $Grid                           = New-Object System.Windows.Forms.DataGridView
    $Grid.Dock                      = 'Fill'
    $Grid.DataSource                = $Table
    $Grid.ReadOnly                  = $true
    $Grid.AllowUserToAddRows        = $false
    $Grid.AllowUserToDeleteRows     = $false
    $Grid.AllowUserToResizeRows     = $false
    $Grid.SelectionMode             = 'FullRowSelect'
    $Grid.MultiSelect               = $true
    $Grid.AutoSizeColumnsMode       = 'Fill'
    $Grid.RowHeadersVisible         = $false
    $Grid.BackgroundColor           = [System.Drawing.Color]::White
    $Grid.BorderStyle               = 'None'
    $Grid.EnableHeadersVisualStyles = $false
    $Grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(45, 62, 80)
    $Grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $Grid.ColumnHeadersDefaultCellStyle.Font      = $FontBold
    $Grid.ColumnHeadersHeight       = 32
    $Grid.RowTemplate.Height        = 26
    $Grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $Grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White

    $Grid.add_CellFormatting({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        $Row = $s.Rows[$e.RowIndex]
        # Farbe nach erster (höchster) Berechtigung - funktioniert auch bei "FullAccess, SendAs"
        $Val = (([string]$Row.Cells['Berechtigung'].Value) -split '[,(]')[0].Trim()
        if ($RowColors.ContainsKey($Val)) { $e.CellStyle.BackColor = $RowColors[$Val] }
        if (([string]$Row.Cells['Quelle'].Value) -like 'Gruppe*') {
            $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(70, 70, 140)
        }
    })

    $Grid.add_DataBindingComplete({
        if ($Grid.Columns.Count -ge 6) {
            $Grid.Columns['Berechtigung'].FillWeight = if ($ChkMerge -and $ChkMerge.Checked) { 22 } else { 11 }
            $Grid.Columns['Postfach'].FillWeight     = 22
            $Grid.Columns['Adresse'].FillWeight      = 26
            $Grid.Columns['Typ'].FillWeight          = 10
            $Grid.Columns['Rechte'].FillWeight       = 14
            $Grid.Columns['Quelle'].FillWeight       = 17
        }
    })

    $Grid.add_CellDoubleClick({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        $Addr = [string]$s.Rows[$e.RowIndex].Cells['Adresse'].Value
        if ($Addr) { Show-MailboxWindow -Address $Addr }
    })

    # --- Kopfbereich ---
    $Top           = New-Object System.Windows.Forms.Panel
    $Top.Dock      = 'Top'
    $Top.Height    = 150
    $Top.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    $LblUser          = New-Object System.Windows.Forms.Label
    $LblUser.Text     = 'Benutzer (UPN):'
    $LblUser.Location = New-Object System.Drawing.Point(12, 16)
    $LblUser.AutoSize = $true

    $TxtUser          = New-Object System.Windows.Forms.TextBox
    $TxtUser.Location = New-Object System.Drawing.Point(120, 12)
    $TxtUser.Width    = 296
    $TxtUser.Text     = $Sync.InitialUser

    $BtnRecent          = New-Object System.Windows.Forms.Button
    $BtnRecent.Text     = [string][char]0x25BE
    $BtnRecent.Location = New-Object System.Drawing.Point(418, 11)
    $BtnRecent.Size     = New-Object System.Drawing.Size(26, 24)
    $CtxRecent          = New-Object System.Windows.Forms.ContextMenuStrip

    $BtnQuery           = New-Object System.Windows.Forms.Button
    $BtnQuery.Text      = 'Abfragen'
    $BtnQuery.Location  = New-Object System.Drawing.Point(455, 10)
    $BtnQuery.Size      = New-Object System.Drawing.Size(110, 28)
    $BtnQuery.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $BtnQuery.ForeColor = [System.Drawing.Color]::White
    $BtnQuery.FlatStyle = 'Flat'
    $BtnQuery.Font      = $FontBold

    $BtnCancel          = New-Object System.Windows.Forms.Button
    $BtnCancel.Text     = 'Abbrechen'
    $BtnCancel.Location = New-Object System.Drawing.Point(575, 10)
    $BtnCancel.Size     = New-Object System.Drawing.Size(100, 28)
    $BtnCancel.Enabled  = $false

    $BtnCache          = New-Object System.Windows.Forms.Button
    $BtnCache.Text     = 'Cache aktualisieren'
    $BtnCache.Location = New-Object System.Drawing.Point(690, 10)
    $BtnCache.Size     = New-Object System.Drawing.Size(150, 28)

    $BtnLive          = New-Object System.Windows.Forms.Button
    $BtnLive.Text     = 'Live prüfen'
    $BtnLive.Location = New-Object System.Drawing.Point(850, 10)
    $BtnLive.Size     = New-Object System.Drawing.Size(150, 28)
    $BtnLive.Enabled  = $false

    $ChkCal           = New-Object System.Windows.Forms.CheckBox
    $ChkCal.Text      = 'Kalenderrechte im Cache'
    $ChkCal.Location  = New-Object System.Drawing.Point(1015, 14)
    $ChkCal.AutoSize  = $true
    $ChkCal.Checked   = $Sync.InitialCalendar -or $Sync.HasCalendar

    $LblFilter          = New-Object System.Windows.Forms.Label
    $LblFilter.Text     = 'Filter:'
    $LblFilter.Location = New-Object System.Drawing.Point(12, 56)
    $LblFilter.AutoSize = $true

    $TxtFilter          = New-Object System.Windows.Forms.TextBox
    $TxtFilter.Location = New-Object System.Drawing.Point(120, 52)
    $TxtFilter.Width    = 320

    $CmbType               = New-Object System.Windows.Forms.ComboBox
    $CmbType.DropDownStyle = 'DropDownList'
    $CmbType.Location      = New-Object System.Drawing.Point(455, 52)
    $CmbType.Width         = 220
    [void]$CmbType.Items.AddRange(@('Alle', 'FullAccess', 'SendAs', 'SendOnBehalf', 'Kalender', 'Buchungsstellvertreter'))
    $CmbType.SelectedIndex = 0

    $BtnExport          = New-Object System.Windows.Forms.Button
    $BtnExport.Text     = 'CSV exportieren'
    $BtnExport.Location = New-Object System.Drawing.Point(690, 50)
    $BtnExport.Size     = New-Object System.Drawing.Size(150, 28)

    $BtnCopy          = New-Object System.Windows.Forms.Button
    $BtnCopy.Text     = 'Auswahl kopieren'
    $BtnCopy.Location = New-Object System.Drawing.Point(850, 50)
    $BtnCopy.Size     = New-Object System.Drawing.Size(150, 28)

    $LblView          = New-Object System.Windows.Forms.Label
    $LblView.Text     = 'Ansicht:'
    $LblView.Location = New-Object System.Drawing.Point(12, 94)
    $LblView.AutoSize = $true

    $ChkMerge          = New-Object System.Windows.Forms.CheckBox
    $ChkMerge.Text     = 'Jedes Postfach nur einmal (Rechte zusammenfassen)'
    $ChkMerge.Location = New-Object System.Drawing.Point(120, 92)
    $ChkMerge.AutoSize = $true
    $ChkMerge.Checked  = $true

    $ChkGroups          = New-Object System.Windows.Forms.CheckBox
    $ChkGroups.Text     = 'Rechte über Gruppen anzeigen'
    $ChkGroups.Location = New-Object System.Drawing.Point(455, 92)
    $ChkGroups.AutoSize = $true
    $ChkGroups.Checked  = $false   # Standard: nur direkte Rechte

    $ChkRooms          = New-Object System.Windows.Forms.CheckBox
    $ChkRooms.Text     = 'Raumpostfächer anzeigen'
    $ChkRooms.Location = New-Object System.Drawing.Point(1015, 56)
    $ChkRooms.AutoSize = $true
    $ChkRooms.Checked  = $true

    $LblCache           = New-Object System.Windows.Forms.Label
    $LblCache.Location  = New-Object System.Drawing.Point(12, 124)
    $LblCache.AutoSize  = $true
    $LblCache.ForeColor = [System.Drawing.Color]::DimGray

    $BtnRemove           = New-Object System.Windows.Forms.Button
    $BtnRemove.Text      = 'Rechte entfernen...'
    $BtnRemove.Location  = New-Object System.Drawing.Point(690, 88)
    $BtnRemove.Size      = New-Object System.Drawing.Size(150, 28)
    $BtnRemove.BackColor = [System.Drawing.Color]::FromArgb(196, 43, 28)
    $BtnRemove.ForeColor = [System.Drawing.Color]::White
    $BtnRemove.FlatStyle = 'Flat'
    $BtnRemove.Font      = $FontBold
    $BtnRemove.Enabled   = $false

    $BtnGrant           = New-Object System.Windows.Forms.Button
    $BtnGrant.Text      = 'Rechte erteilen...'
    $BtnGrant.Location  = New-Object System.Drawing.Point(850, 88)
    $BtnGrant.Size      = New-Object System.Drawing.Size(150, 28)
    $BtnGrant.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
    $BtnGrant.ForeColor = [System.Drawing.Color]::White
    $BtnGrant.FlatStyle = 'Flat'
    $BtnGrant.Font      = $FontBold
    $BtnGrant.Enabled   = $false

    $BtnMbView           = New-Object System.Windows.Forms.Button
    $BtnMbView.Text      = 'Postfach-Zugriff...'
    $BtnMbView.Location  = New-Object System.Drawing.Point(1200, 10)
    $BtnMbView.Size      = New-Object System.Drawing.Size(160, 28)

    $BtnOffboard           = New-Object System.Windows.Forms.Button
    $BtnOffboard.Text      = 'Offboarding...'
    $BtnOffboard.Location  = New-Object System.Drawing.Point(1200, 50)
    $BtnOffboard.Size      = New-Object System.Drawing.Size(160, 28)
    $BtnOffboard.ForeColor = [System.Drawing.Color]::FromArgb(160, 30, 20)
    $BtnOffboard.Font      = $FontBold
    $BtnOffboard.Enabled   = $false

    $BtnCopyRights           = New-Object System.Windows.Forms.Button
    $BtnCopyRights.Text      = 'Rechte kopieren...'
    $BtnCopyRights.Location  = New-Object System.Drawing.Point(1200, 88)
    $BtnCopyRights.Size      = New-Object System.Drawing.Size(160, 28)
    $BtnCopyRights.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $BtnCopyRights.ForeColor = [System.Drawing.Color]::White
    $BtnCopyRights.FlatStyle = 'Flat'
    $BtnCopyRights.Font      = $FontBold
    $BtnCopyRights.Enabled   = $false

    $LnkLog          = New-Object System.Windows.Forms.LinkLabel
    $LnkLog.Text     = 'Protokoll anzeigen'
    $LnkLog.Location = New-Object System.Drawing.Point(1200, 124)
    $LnkLog.AutoSize = $true

    $LblRemoveHint           = New-Object System.Windows.Forms.Label
    $LblRemoveHint.Text      = 'Mehrfachauswahl: Strg / Shift'
    $LblRemoveHint.Location  = New-Object System.Drawing.Point(1015, 94)
    $LblRemoveHint.AutoSize  = $true
    $LblRemoveHint.ForeColor = [System.Drawing.Color]::DimGray

    $Top.Controls.AddRange(@($BtnRecent, $BtnMbView, $BtnOffboard, $BtnCopyRights, $LnkLog))
    $Top.Controls.AddRange(@($LblUser, $TxtUser, $BtnQuery, $BtnCancel, $BtnCache, $BtnLive, $ChkCal,
                             $LblFilter, $TxtFilter, $CmbType, $BtnExport, $BtnCopy,
                             $LblView, $ChkMerge, $ChkGroups, $ChkRooms, $BtnRemove, $BtnGrant, $LblRemoveHint, $LblCache))

    # --- Statusleiste ---
    $Status              = New-Object System.Windows.Forms.StatusStrip
    $LblStatus           = New-Object System.Windows.Forms.ToolStripStatusLabel
    $LblStatus.Spring    = $true
    $LblStatus.TextAlign = 'MiddleLeft'
    $LblStatus.Text      = 'Bereit'
    $LblCount            = New-Object System.Windows.Forms.ToolStripStatusLabel
    $LnkUpdate           = New-Object System.Windows.Forms.ToolStripStatusLabel
    $LnkUpdate.IsLink    = $true
    $LnkUpdate.Visible   = $false
    $LnkUpdate.Font      = $FontBold
    $LnkUpdate.LinkColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $Progress            = New-Object System.Windows.Forms.ToolStripProgressBar
    $Progress.Size       = New-Object System.Drawing.Size(200, 16)
    $Progress.MarqueeAnimationSpeed = 30
    [void]$Status.Items.AddRange(@($LblStatus, $LnkUpdate, $LblCount, $Progress))

    $Form.Controls.Add($Grid)
    $Form.Controls.Add($Top)
    $Form.Controls.Add($Status)

    # --- Menü Extras (v3.9) ---
    $Menu      = New-Object System.Windows.Forms.MenuStrip
    $MnuExtras = New-Object System.Windows.Forms.ToolStripMenuItem('Extras')
    $MnuFwd    = New-Object System.Windows.Forms.ToolStripMenuItem('Weiterleitungen && Posteingangsregeln des Benutzers')
    $MnuGrp    = New-Object System.Windows.Forms.ToolStripMenuItem('Gruppenmitgliedschaften des Benutzers...')
    $MnuFold   = New-Object System.Windows.Forms.ToolStripMenuItem('Ordnerrechte eines Postfachs...')
    $MnuMbView = New-Object System.Windows.Forms.ToolStripMenuItem('Wer hat Zugriff auf ein Postfach?')
    $MnuRep    = New-Object System.Windows.Forms.ToolStripMenuItem('Tenant-Gesamtreport (HTML + CSV)...')
    $MnuLog    = New-Object System.Windows.Forms.ToolStripMenuItem('Änderungsprotokoll anzeigen')
    $MnuDir    = New-Object System.Windows.Forms.ToolStripMenuItem('Datenordner öffnen')
    [void]$MnuExtras.DropDownItems.Add($MnuFwd)
    [void]$MnuExtras.DropDownItems.Add($MnuGrp)
    [void]$MnuExtras.DropDownItems.Add($MnuFold)
    [void]$MnuExtras.DropDownItems.Add($MnuMbView)
    [void]$MnuExtras.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    [void]$MnuExtras.DropDownItems.Add($MnuRep)
    [void]$MnuExtras.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    [void]$MnuExtras.DropDownItems.Add($MnuLog)
    [void]$MnuExtras.DropDownItems.Add($MnuDir)
    [void]$Menu.Items.Add($MnuExtras)

    # --- Menü Cache + Hilfe (v4.0) ---
    $MnuCache      = New-Object System.Windows.Forms.ToolStripMenuItem('Cache')
    $MnuDelta      = New-Object System.Windows.Forms.ToolStripMenuItem('Delta-Update (Audit-Log)')
    $MnuRebuild    = New-Object System.Windows.Forms.ToolStripMenuItem('Komplett neu aufbauen...')
    $MnuLiveM      = New-Object System.Windows.Forms.ToolStripMenuItem('Live prüfen (angezeigte + Shared/Raum)')
    [void]$MnuCache.DropDownItems.Add($MnuDelta)
    [void]$MnuCache.DropDownItems.Add($MnuRebuild)
    [void]$MnuCache.DropDownItems.Add($MnuLiveM)
    [void]$Menu.Items.Add($MnuCache)

    $MnuHelp       = New-Object System.Windows.Forms.ToolStripMenuItem('Hilfe')
    $MnuUpd        = New-Object System.Windows.Forms.ToolStripMenuItem('Nach Updates suchen')
    $MnuGitHub     = New-Object System.Windows.Forms.ToolStripMenuItem('Projekt auf GitHub öffnen')
    $MnuAbout      = New-Object System.Windows.Forms.ToolStripMenuItem('Über MailboxAccess')
    [void]$MnuHelp.DropDownItems.Add($MnuUpd)
    [void]$MnuHelp.DropDownItems.Add($MnuGitHub)
    [void]$MnuHelp.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    [void]$MnuHelp.DropDownItems.Add($MnuAbout)
    [void]$Menu.Items.Add($MnuHelp)

    $Form.MainMenuStrip = $Menu
    $Form.Controls.Add($Menu)

    $MnuFwd.add_Click({ Invoke-ForwardCheck })
    $MnuGrp.add_Click({ Invoke-GroupCheck })
    $MnuFold.add_Click({ Invoke-FolderCheck })
    $MnuMbView.add_Click({ Show-MailboxWindow })
    $MnuRep.add_Click({ Invoke-TenantReport })
    $MnuLog.add_Click({ Show-LogWindow })
    $MnuDir.add_Click({ if (Test-Path $Sync.DataDir) { Start-Process explorer.exe $Sync.DataDir } })
    $MnuCache.add_DropDownOpening({
        $B = Test-Busy
        $MnuDelta.Enabled   = (-not $B) -and $Sync.HasCache -and $Sync.CanAudit -and $Sync.CacheHasRooms
        $MnuDelta.Text      = if ($Sync.CanAudit) { 'Delta-Update (Audit-Log)' } else { 'Delta-Update (keine Audit-Log-Rolle)' }
        $MnuRebuild.Enabled = -not $B
        $MnuLiveM.Enabled   = $BtnLive.Enabled
    })
    $MnuDelta.add_Click({ Invoke-DeltaUpdate })
    $MnuRebuild.add_Click({ [void](Request-CacheBuild) })
    $MnuLiveM.add_Click({ Invoke-LiveCheck })
    $MnuUpd.add_Click({
        $LblStatus.Text = 'Suche nach Updates...'
        Start-UiUpdateCheck -Manual $true
    })
    $MnuGitHub.add_Click({ Start-Process "https://github.com/$($Sync.RepoSlug)" })
    $MnuAbout.add_Click({
        $Txt = "MailboxAccess v$($Sync.AppVersion)`n" +
               "Exchange Online Postfachberechtigungen`n`n" +
               "Angemeldet: $($Sync.ConnUpn)`n" +
               "Schreibrechte: $(if ($Sync.CanWrite) { 'ja' } else { 'nein (Nur-Lesen)' })`n" +
               "Audit-Log (Delta-Update): $(if ($Sync.CanAudit) { 'ja' } else { 'nein' })`n" +
               "Datenordner: $($Sync.DataDir)`n`n" +
               "https://github.com/$($Sync.RepoSlug)`n" +
               "Autor: RazeAttack (GitHub) - Lizenz siehe Repository"
        [void][System.Windows.Forms.MessageBox]::Show($Form, $Txt, 'Über MailboxAccess', 'OK', 'Information')
    })
    $LnkUpdate.add_Click({
        $Url = if ($State.UpdateUrl) { [string]$State.UpdateUrl } else { "https://github.com/$($Sync.RepoSlug)/releases" }
        Start-Process $Url
    })

    $MnuExtras.add_DropDownOpening({
        $B = Test-Busy
        foreach ($It in @($MnuFwd, $MnuGrp, $MnuFold, $MnuRep)) { $It.Enabled = -not $B }
        $MnuGrp.Enabled = $MnuGrp.Enabled -and $Sync.CanWrite
        $MnuRep.Enabled = $MnuRep.Enabled -and $Sync.HasCache
    })
    $Grid.BringToFront()
    $Form.AcceptButton = $BtnQuery

    # --- Tooltips ---
    $Tip = New-Object System.Windows.Forms.ToolTip
    $Tip.SetToolTip($BtnCache, 'Alle Postfachberechtigungen komplett neu einlesen (einige Minuten)')
    $Tip.SetToolTip($BtnLive,  'Vollzugriff live nachlesen: angezeigte Postfächer + alle Shared-, Raum- und Geräte-Postfächer')
    $Tip.SetToolTip($BtnRecent,     'Zuletzt abgefragte Benutzer')
    $Tip.SetToolTip($BtnMbView,     'Wer hat Zugriff auf ein bestimmtes Postfach? (auch per Doppelklick in der Liste)')
    $Tip.SetToolTip($BtnOffboard,   'Alle direkten Rechte des abgefragten Benutzers entfernen - vorher CSV-Nachweis')
    $Tip.SetToolTip($BtnCopyRights, 'Direkte Rechte des abgefragten Benutzers auf einen anderen Benutzer übertragen (Onboarding)')
    $Tip.SetToolTip($BtnGrant, 'Dem Benutzer Vollzugriff / Senden als / Senden im Auftrag von erteilen - bei Raumpostfächern auch Kalender und Buchungsstellvertreter')
    $Tip.SetToolTip($BtnRemove, 'Markierte Postfächer: direkte Rechte des Benutzers entfernen')
    $Tip.SetToolTip($ChkRooms, 'Raum- und Geräte-Postfächer (Kalenderrechte, Buchungsstellvertreter, Vollzugriff) ein-/ausblenden')
    $Tip.SetToolTip($ChkCal,   'Kalenderrechte beim nächsten Cache-Aufbau mit einlesen (ca. doppelte Dauer)')

    # --- Kontextmenü (Rechtsklick in der Liste) ---
    $Ctx      = New-Object System.Windows.Forms.ContextMenuStrip
    $MiMbView = $Ctx.Items.Add('Wer hat Zugriff auf dieses Postfach?')
    $MiFolder = $Ctx.Items.Add('Ordnerrechte dieses Postfachs...')
    [void]$Ctx.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $MiGrant  = $Ctx.Items.Add('Rechte erteilen...')
    $MiRemove = $Ctx.Items.Add('Rechte entfernen...')
    $MiCopy   = $Ctx.Items.Add('Adresse(n) kopieren')
    $Grid.ContextMenuStrip = $Ctx

    $Grid.add_CellMouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right -and $e.RowIndex -ge 0 -and -not $s.Rows[$e.RowIndex].Selected) {
            $s.ClearSelection()
            $s.Rows[$e.RowIndex].Selected = $true
        }
    })

    # --- Hilfsfunktionen ---
    function Test-Busy { return ($Sync.Busy -or $Sync.Queue.Count -gt 0) }

    # Dialog: ausgewählte direkte Berechtigungen entfernen
    function Invoke-RemoveDialog {
        if (Test-Busy) { return }
        $User = [string]$Sync.ResultUser
        if (-not $User -or $Grid.SelectedRows.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show($Form, 'Bitte zuerst ein oder mehrere Postfächer in der Liste markieren.', 'Rechte entfernen', 'OK', 'Information')
            return
        }

        $Removable = @('FullAccess', 'SendAs', 'SendOnBehalf', 'Kalender', 'Buchungsstellvertreter')
        $Sel = @($Grid.SelectedRows | ForEach-Object {
            [PSCustomObject]@{ Adresse = [string]$_.Cells['Adresse'].Value; Berechtigung = [string]$_.Cells['Berechtigung'].Value }
        } | Where-Object { $_.Adresse })

        $Items = @()
        $Notes = @()
        foreach ($A in @($Sel | ForEach-Object { $_.Adresse } | Select-Object -Unique)) {
            $Entries = @($State.Raw | Where-Object { $_ -and ([string]$_.Adresse) -eq $A })
            if (-not $ChkMerge.Checked) {
                # Detailansicht: nur die markierten Zeilen
                $SelPerms = @($Sel | Where-Object { $_.Adresse -eq $A } | ForEach-Object { $_.Berechtigung })
                $Entries  = @($Entries | Where-Object { $SelPerms -contains [string]$_.Berechtigung })
            }
            if (-not $ChkGroups.Checked) {
                $Entries = @($Entries | Where-Object { -not (([string]$_.Quelle) -like 'Gruppe*') })
            }
            if (-not $ChkRooms.Checked) {
                $Entries = @($Entries | Where-Object { ([string]$_.Typ) -notin 'Room', 'Equipment' })
            }
            if ($Entries.Count -eq 0) { continue }

            $Name   = [string]$Entries[0].Postfach
            $Direct = @($Entries | Where-Object { $_.Quelle -eq 'direkt' -and $Removable -contains [string]$_.Berechtigung } |
                        ForEach-Object { [string]$_.Berechtigung } | Select-Object -Unique)

            foreach ($E in ($Entries | Where-Object { $_.Quelle -ne 'direkt' -or $Removable -notcontains [string]$_.Berechtigung })) {
                $Why = if (([string]$E.Quelle) -like 'Gruppe*')      { "$($E.Quelle) -> Gruppenmitgliedschaft ändern" }
                       elseif (([string]$E.Quelle) -like '*vererbt*') { 'vererbt -> an der Quelle entfernen' }
                       else                                          { 'nicht über dieses Tool entfernbar' }
                $Notes += "  $Name - $($E.Berechtigung): $Why"
            }
            if ($Direct.Count -gt 0) {
                $Items += [PSCustomObject]@{ Adresse = $A; Postfach = $Name; Perms = $Direct }
            }
        }
        $Notes = @($Notes | Select-Object -Unique)

        if ($Items.Count -eq 0) {
            $Msg = "In der Auswahl sind keine direkt entfernbaren Berechtigungen."
            if ($Notes.Count) { $Msg += "`n`nNicht entfernbar:`n" + ($Notes -join "`n") }
            [void][System.Windows.Forms.MessageBox]::Show($Form, $Msg, 'Rechte entfernen', 'OK', 'Information')
            return
        }

        # --- Dialog aufbauen ---
        $Dlg                 = New-Object System.Windows.Forms.Form
        $Dlg.Text            = 'Berechtigungen entfernen'
        $Dlg.Size            = New-Object System.Drawing.Size(680, 760)
        $Dlg.StartPosition   = 'CenterParent'
        $Dlg.FormBorderStyle = 'FixedDialog'
        $Dlg.MaximizeBox     = $false
        $Dlg.MinimizeBox     = $false
        $Dlg.Font            = $Font
        if ($Form.Icon) { $Dlg.Icon = $Form.Icon }

        $LblHead          = New-Object System.Windows.Forms.Label
        $LblHead.Text     = "Berechtigungen von  $User  entfernen:"
        $LblHead.Font     = $FontBold
        $LblHead.Location = New-Object System.Drawing.Point(12, 12)
        $LblHead.AutoSize = $true

        # Bezeichnungen wie im Exchange Admin Center
        $PermInfo = [ordered]@{
            FullAccess   = @{ Name = 'Lesen und Verwalten (Vollzugriff)'
                              Desc = 'Anmelden am Postfach und wie der Besitzer handeln.' }
            SendAs       = @{ Name = 'Senden als'
                              Desc = 'E-Mails aus dem Postfach senden - erscheinen als vom Postfachbesitzer gesendet.' }
            SendOnBehalf = @{ Name = 'Senden im Auftrag von'
                              Desc = 'E-Mails im Auftrag senden - Empfänger sehen "<Benutzer> im Auftrag von <Postfach>".' }
            Kalender     = @{ Name = 'Kalenderberechtigung'
                              Desc = 'Ordnerrecht auf den Kalender (z. B. Reviewer, Editor) - auch Raum-/Gerätepostfächer.' }
            Buchungsstellvertreter = @{ Name = 'Buchungsstellvertreter'
                              Desc = 'Darf Buchungsanfragen für das Raum-/Gerätepostfach annehmen oder ablehnen.' }
        }

        $Chk = @{}
        $Y   = 44
        foreach ($P in $Removable) {
            $N = @($Items | Where-Object { $_.Perms -contains $P }).Count

            $C          = New-Object System.Windows.Forms.CheckBox
            $C.Text     = "$($PermInfo[$P].Name) ($N)"
            $C.Font     = $FontBold
            $C.Location = New-Object System.Drawing.Point(12, $Y)
            $C.AutoSize = $true
            $C.Enabled  = $N -gt 0
            $C.Checked  = $N -gt 0
            $Chk[$P]    = $C

            $D           = New-Object System.Windows.Forms.Label
            $D.Text      = $PermInfo[$P].Desc
            $D.Location  = New-Object System.Drawing.Point(30, ($Y + 24))
            $D.AutoSize  = $true
            $D.ForeColor = [System.Drawing.Color]::DimGray

            $Dlg.Controls.AddRange(@($C, $D))
            $Y += 56
        }

        $TxtPrev            = New-Object System.Windows.Forms.TextBox
        $TxtPrev.Multiline  = $true
        $TxtPrev.ReadOnly   = $true
        $TxtPrev.ScrollBars = 'Vertical'
        $TxtPrev.WordWrap   = $false
        $TxtPrev.Font       = New-Object System.Drawing.Font('Consolas', 9)
        $TxtPrev.Location   = New-Object System.Drawing.Point(12, 330)
        $TxtPrev.Size       = New-Object System.Drawing.Size(640, 334)
        $TxtPrev.BackColor  = [System.Drawing.Color]::White

        $BtnOk           = New-Object System.Windows.Forms.Button
        $BtnOk.Text      = 'Entfernen'
        $BtnOk.Location  = New-Object System.Drawing.Point(432, 676)
        $BtnOk.Size      = New-Object System.Drawing.Size(110, 30)
        $BtnOk.BackColor = [System.Drawing.Color]::FromArgb(196, 43, 28)
        $BtnOk.ForeColor = [System.Drawing.Color]::White
        $BtnOk.FlatStyle = 'Flat'
        $BtnOk.Font      = $FontBold
        $BtnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK

        $BtnNo          = New-Object System.Windows.Forms.Button
        $BtnNo.Text     = 'Abbrechen'
        $BtnNo.Location = New-Object System.Drawing.Point(552, 676)
        $BtnNo.Size     = New-Object System.Drawing.Size(100, 30)
        $BtnNo.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

        $Dlg.Controls.AddRange(@($LblHead, $TxtPrev, $BtnOk, $BtnNo))
        $Dlg.CancelButton = $BtnNo

        $UpdatePreview = {
            $Chosen = @($Removable | Where-Object { $Chk[$_].Checked })
            $Lines  = @()
            $Count  = 0
            foreach ($It in $Items) {
                $P = @($It.Perms | Where-Object { $Chosen -contains $_ })
                if ($P.Count -eq 0) { continue }
                $Count += $P.Count
                $Lines += ("  {0,-30} {1}" -f $It.Postfach, (($P | ForEach-Object { $PermInfo[$_].Name }) -join ', '))
            }
            $Text = "Wird entfernt ($Count):`r`n" + ($Lines -join "`r`n")
            if ($Notes.Count) { $Text += "`r`n`r`nNicht entfernbar (bleibt bestehen):`r`n" + ($Notes -join "`r`n") }
            $TxtPrev.Text   = $Text
            $BtnOk.Enabled  = $Count -gt 0
        }
        foreach ($C in $Chk.Values) { $C.add_CheckedChanged($UpdatePreview) }
        & $UpdatePreview

        $Result = $Dlg.ShowDialog($Form)
        $Chosen = @($Removable | Where-Object { $Chk[$_].Checked })
        $Dlg.Dispose()
        if ($Result -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $Final = @(foreach ($It in $Items) {
            $P = @($It.Perms | Where-Object { $Chosen -contains $_ })
            if ($P.Count) { [PSCustomObject]@{ Adresse = $It.Adresse; Postfach = $It.Postfach; Perms = $P } }
        })
        if ($Final.Count -eq 0) { return }
        $Count = [int](@($Final | ForEach-Object { $_.Perms.Count }) | Measure-Object -Sum).Sum

        $Ans = [System.Windows.Forms.MessageBox]::Show($Form,
            "$Count Berechtigung(en) auf $($Final.Count) Postfach/Postfächern für`n$User`nwirklich entfernen?",
            'Bestätigen', 'YesNo', 'Warning', [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($Ans -ne 'Yes') { return }

        $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'Remove'; User = $User; Items = $Final })
        $LblStatus.Text = 'Entferne Berechtigungen...'
    }

    # Dialog: Rechte erteilen (3 Standardrechte, bei Räumen zusätzlich Kalender + Buchungsstellvertreter)
    function Invoke-GrantDialog {
        if (Test-Busy) { return }
        $User = (Get-InputAddress $TxtUser.Text)
        if (-not $User) { $User = [string]$Sync.ResultUser }
        if (-not $User) {
            [void][System.Windows.Forms.MessageBox]::Show($Form, 'Bitte oben zuerst den Benutzer eintragen.', 'Rechte erteilen', 'OK', 'Information')
            return
        }
        $List = @($Sync.MailboxList | Where-Object { $_ })
        if ($List.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show($Form, 'Keine Postfachliste im Cache. Bitte zuerst den Cache aufbauen.', 'Rechte erteilen', 'OK', 'Information')
            return
        }

        $RoomTypes = @('Room', 'Equipment')
        $Picked    = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        $ByAddr    = @{}
        foreach ($M in $List) { $ByAddr[$M.Address] = $M }

        # Vorauswahl: in der Liste markierte Postfächer
        foreach ($Row in $Grid.SelectedRows) {
            $A = [string]$Row.Cells['Adresse'].Value
            if ($A -and $ByAddr.ContainsKey($A)) { [void]$Picked.Add($A) }
        }

        $Dlg                 = New-Object System.Windows.Forms.Form
        $Dlg.Text            = 'Berechtigungen erteilen'
        $Dlg.Size            = New-Object System.Drawing.Size(740, 780)
        $Dlg.StartPosition   = 'CenterParent'
        $Dlg.FormBorderStyle = 'FixedDialog'
        $Dlg.MaximizeBox     = $false
        $Dlg.MinimizeBox     = $false
        $Dlg.Font            = $Font
        if ($Form.Icon) { $Dlg.Icon = $Form.Icon }

        $LblHead          = New-Object System.Windows.Forms.Label
        $LblHead.Text     = "Berechtigungen erteilen für:  $User"
        $LblHead.Font     = $FontBold
        $LblHead.Location = New-Object System.Drawing.Point(12, 12)
        $LblHead.AutoSize = $true

        $LblSearch          = New-Object System.Windows.Forms.Label
        $LblSearch.Text     = 'Postfach suchen:'
        $LblSearch.Location = New-Object System.Drawing.Point(12, 46)
        $LblSearch.AutoSize = $true

        $TxtSearch          = New-Object System.Windows.Forms.TextBox
        $TxtSearch.Location = New-Object System.Drawing.Point(120, 43)
        $TxtSearch.Size     = New-Object System.Drawing.Size(330, 24)

        $ChkOnlyRooms          = New-Object System.Windows.Forms.CheckBox
        $ChkOnlyRooms.Text     = 'nur Raum-/Geräte-Postfächer'
        $ChkOnlyRooms.Location = New-Object System.Drawing.Point(470, 45)
        $ChkOnlyRooms.AutoSize = $true

        $Lv               = New-Object System.Windows.Forms.ListView
        $Lv.Location      = New-Object System.Drawing.Point(12, 76)
        $Lv.Size          = New-Object System.Drawing.Size(700, 280)
        $Lv.View          = 'Details'
        $Lv.CheckBoxes    = $true
        $Lv.FullRowSelect = $true
        $Lv.HideSelection = $false
        [void]$Lv.Columns.Add('Postfach', 280)
        [void]$Lv.Columns.Add('Adresse', 290)
        [void]$Lv.Columns.Add('Typ', 100)

        $LblPicked          = New-Object System.Windows.Forms.Label
        $LblPicked.Location = New-Object System.Drawing.Point(12, 362)
        $LblPicked.AutoSize = $true
        $LblPicked.ForeColor = [System.Drawing.Color]::DimGray

        # --- Standardrechte ---
        $GrpStd          = New-Object System.Windows.Forms.GroupBox
        $GrpStd.Text     = 'Postfachrechte (alle Postfachtypen)'
        $GrpStd.Location = New-Object System.Drawing.Point(12, 388)
        $GrpStd.Size     = New-Object System.Drawing.Size(700, 140)

        $CkFull          = New-Object System.Windows.Forms.CheckBox
        $CkFull.Text     = 'Lesen und Verwalten (Vollzugriff)'
        $CkFull.Font     = $FontBold
        $CkFull.Location = New-Object System.Drawing.Point(14, 24)
        $CkFull.AutoSize = $true

        $CkAuto          = New-Object System.Windows.Forms.CheckBox
        $CkAuto.Text     = 'Automatisch in Outlook einbinden (AutoMapping)'
        $CkAuto.Location = New-Object System.Drawing.Point(34, 48)
        $CkAuto.AutoSize = $true
        $CkAuto.Checked  = $true
        $CkAuto.Enabled  = $false

        $CkSendAs          = New-Object System.Windows.Forms.CheckBox
        $CkSendAs.Text     = 'Senden als'
        $CkSendAs.Font     = $FontBold
        $CkSendAs.Location = New-Object System.Drawing.Point(14, 78)
        $CkSendAs.AutoSize = $true

        $CkBehalf          = New-Object System.Windows.Forms.CheckBox
        $CkBehalf.Text     = 'Senden im Auftrag von'
        $CkBehalf.Font     = $FontBold
        $CkBehalf.Location = New-Object System.Drawing.Point(14, 106)
        $CkBehalf.AutoSize = $true

        $GrpStd.Controls.AddRange(@($CkFull, $CkAuto, $CkSendAs, $CkBehalf))

        # --- Raumrechte ---
        $GrpRoom          = New-Object System.Windows.Forms.GroupBox
        $GrpRoom.Text     = 'Nur Raum-/Geräte-Postfächer'
        $GrpRoom.Location = New-Object System.Drawing.Point(12, 536)
        $GrpRoom.Size     = New-Object System.Drawing.Size(700, 96)

        $CkCal          = New-Object System.Windows.Forms.CheckBox
        $CkCal.Text     = 'Kalenderberechtigung'
        $CkCal.Font     = $FontBold
        $CkCal.Location = New-Object System.Drawing.Point(14, 26)
        $CkCal.AutoSize = $true

        $CmbRole               = New-Object System.Windows.Forms.ComboBox
        $CmbRole.DropDownStyle = 'DropDownList'
        $CmbRole.Location      = New-Object System.Drawing.Point(220, 23)
        $CmbRole.Size          = New-Object System.Drawing.Size(200, 24)
        [void]$CmbRole.Items.AddRange(@('AvailabilityOnly', 'LimitedDetails', 'Reviewer', 'Author', 'Editor', 'PublishingEditor', 'Owner'))
        $CmbRole.SelectedItem  = 'Reviewer'
        $CmbRole.Enabled       = $false

        $LblRole          = New-Object System.Windows.Forms.Label
        $LblRole.Text     = 'Reviewer = lesen, Editor = bearbeiten'
        $LblRole.Location = New-Object System.Drawing.Point(430, 27)
        $LblRole.AutoSize = $true
        $LblRole.ForeColor = [System.Drawing.Color]::DimGray

        $CkDel          = New-Object System.Windows.Forms.CheckBox
        $CkDel.Text     = 'Buchungsstellvertreter (darf Buchungsanfragen annehmen/ablehnen)'
        $CkDel.Font     = $FontBold
        $CkDel.Location = New-Object System.Drawing.Point(14, 60)
        $CkDel.AutoSize = $true

        $GrpRoom.Controls.AddRange(@($CkCal, $CmbRole, $LblRole, $CkDel))

        $BtnOk           = New-Object System.Windows.Forms.Button
        $BtnOk.Text      = 'Erteilen'
        $BtnOk.Location  = New-Object System.Drawing.Point(492, 690)
        $BtnOk.Size      = New-Object System.Drawing.Size(110, 30)
        $BtnOk.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
        $BtnOk.ForeColor = [System.Drawing.Color]::White
        $BtnOk.FlatStyle = 'Flat'
        $BtnOk.Font      = $FontBold
        $BtnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK

        $BtnNo          = New-Object System.Windows.Forms.Button
        $BtnNo.Text     = 'Abbrechen'
        $BtnNo.Location = New-Object System.Drawing.Point(612, 690)
        $BtnNo.Size     = New-Object System.Drawing.Size(100, 30)
        $BtnNo.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

        $LblSum          = New-Object System.Windows.Forms.Label
        $LblSum.Location = New-Object System.Drawing.Point(12, 645)
        $LblSum.Size     = New-Object System.Drawing.Size(700, 40)
        $LblSum.ForeColor = [System.Drawing.Color]::DimGray

        $Dlg.Controls.AddRange(@($LblHead, $LblSearch, $TxtSearch, $ChkOnlyRooms, $Lv, $LblPicked, $GrpStd, $GrpRoom, $LblSum, $BtnOk, $BtnNo))
        $Dlg.CancelButton = $BtnNo

        $DlgState = @{ Loading = $false }

        $UpdateState = {
            $Sel    = @($Picked | ForEach-Object { $ByAddr[$_] } | Where-Object { $_ })
            $Rooms  = @($Sel | Where-Object { $_.Type -in $RoomTypes }).Count
            $Others = $Sel.Count - $Rooms

            $LblPicked.Text   = "Ausgewählt: $($Sel.Count) Postfach/Postfächer ($Rooms Raum/Gerät)"
            $CkAuto.Enabled   = $CkFull.Checked
            $GrpRoom.Enabled  = $Rooms -gt 0
            $CmbRole.Enabled  = $CkCal.Checked -and $Rooms -gt 0

            $N = 0
            foreach ($C in @($CkFull, $CkSendAs, $CkBehalf)) { if ($C.Checked) { $N += $Sel.Count } }
            if ($Rooms -gt 0) { foreach ($C in @($CkCal, $CkDel)) { if ($C.Checked) { $N += $Rooms } } }
            $Hint = if ($Rooms -gt 0 -and $Others -gt 0 -and ($CkCal.Checked -or $CkDel.Checked)) { "  Kalender/Buchungsstellvertreter nur für die $Rooms Raum-/Geräte-Postfächer." } else { '' }
            $LblSum.Text    = "Es werden $N Berechtigung(en) erteilt.$Hint"
            $BtnOk.Enabled  = $N -gt 0
        }

        $FillList = {
            $DlgState.Loading = $true
            $Lv.BeginUpdate()
            $Lv.Items.Clear()
            $Q = [System.Management.Automation.WildcardPattern]::Escape($TxtSearch.Text.Trim())
            foreach ($M in $List) {
                if ($ChkOnlyRooms.Checked -and $M.Type -notin $RoomTypes) { continue }
                if ($Q -and -not ($M.Name -like "*$Q*" -or $M.Address -like "*$Q*") -and -not $Picked.Contains($M.Address)) { continue }
                $Li = New-Object System.Windows.Forms.ListViewItem($M.Name)
                [void]$Li.SubItems.Add($M.Address)
                [void]$Li.SubItems.Add($M.Type)
                $Li.Tag     = $M.Address
                $Li.Checked = $Picked.Contains($M.Address)
                if ($M.Type -in $RoomTypes) { $Li.BackColor = [System.Drawing.Color]::FromArgb(214, 240, 240) }
                [void]$Lv.Items.Add($Li)
            }
            $Lv.EndUpdate()
            $DlgState.Loading = $false
            & $UpdateState
        }

        $Lv.add_ItemChecked({
            param($s, $e)
            if ($DlgState.Loading) { return }
            if ($e.Item.Checked) { [void]$Picked.Add([string]$e.Item.Tag) } else { [void]$Picked.Remove([string]$e.Item.Tag) }
            & $UpdateState
        })
        $TxtSearch.add_TextChanged($FillList)
        $ChkOnlyRooms.add_CheckedChanged($FillList)
        foreach ($C in @($CkFull, $CkSendAs, $CkBehalf, $CkCal, $CkDel)) { $C.add_CheckedChanged($UpdateState) }

        & $FillList
        $Dlg.add_Shown({ [void]$TxtSearch.Focus() })

        $Result = $Dlg.ShowDialog($Form)

        $Std = @()
        if ($CkFull.Checked)   { $Std += 'FullAccess' }
        if ($CkSendAs.Checked) { $Std += 'SendAs' }
        if ($CkBehalf.Checked) { $Std += 'SendOnBehalf' }
        $RoomPerms = @()
        if ($CkCal.Checked) { $RoomPerms += 'Kalender' }
        if ($CkDel.Checked) { $RoomPerms += 'Buchungsstellvertreter' }
        $Options = @{ AutoMapping = [bool]$CkAuto.Checked; CalendarRole = [string]$CmbRole.SelectedItem }
        $Sel = @($Picked | ForEach-Object { $ByAddr[$_] } | Where-Object { $_ })
        $Dlg.Dispose()
        if ($Result -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $Items = @(foreach ($M in $Sel) {
            $IsRoom = $M.Type -in $RoomTypes
            $P = @($Std) + $(if ($IsRoom) { @($RoomPerms) } else { @() })
            if ($P.Count) { [PSCustomObject]@{ Adresse = $M.Address; Postfach = $M.Name; IsRoom = $IsRoom; Perms = $P } }
        })
        if ($Items.Count -eq 0) { return }

        $Names = @{ FullAccess = 'Vollzugriff'; SendAs = 'Senden als'; SendOnBehalf = 'Senden im Auftrag von'; Kalender = "Kalender ($($Options.CalendarRole))"; Buchungsstellvertreter = 'Buchungsstellvertreter' }
        $Preview = @($Items | Select-Object -First 15 | ForEach-Object { "  $($_.Postfach): $((@($_.Perms) | ForEach-Object { $Names[$_] }) -join ', ')" })
        if ($Items.Count -gt 15) { $Preview += "  ... und $($Items.Count - 15) weitere" }
        $Count = [int](@($Items | ForEach-Object { @($_.Perms).Count }) | Measure-Object -Sum).Sum

        $Ans = [System.Windows.Forms.MessageBox]::Show($Form,
            "$Count Berechtigung(en) für`n$User`nerteilen?`n`n$($Preview -join "`n")",
            'Bestätigen', 'YesNo', 'Question', [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($Ans -ne 'Yes') { return }

        $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'Grant'; User = $User; Items = $Items; Options = $Options })
        $LblStatus.Text = 'Erteile Berechtigungen...'
    }

    # ---------- Hilfsfunktionen v3.8 ----------

    # "Max Muster <max@domain>" -> "max@domain"
    function Get-InputAddress([string]$Text) {
        if (-not $Text) { return '' }
        $T = $Text.Trim()
        if ($T -match '<([^>]+)>') { return $Matches[1].Trim() }
        return $T
    }

    function Add-RecentUser([string]$U) {
        if (-not $U) { return }
        $State.Recent = @(@($U) + @($State.Recent | Where-Object { $_ -and $_ -ne $U }) | Select-Object -First 15)
        try {
            if (-not (Test-Path $Sync.DataDir)) { New-Item -ItemType Directory -Path $Sync.DataDir -Force | Out-Null }
            $State.Recent | Set-Content -Path $RecentFile -Encoding UTF8
        } catch { }
    }

    function Update-AutoComplete {
        if (-not [object]::ReferenceEquals($State.AutoListRef, $Sync.MailboxList) -or -not $State.AutoSrc) {
            $Src  = New-Object System.Windows.Forms.AutoCompleteStringCollection
            $Vals = [System.Collections.Generic.List[string]]::new()
            foreach ($M in @($Sync.MailboxList | Where-Object { $_ })) {
                $Vals.Add("$($M.Name) <$($M.Address)>")
                $Vals.Add([string]$M.Address)
            }
            foreach ($U in @($State.Recent)) { if ($U) { $Vals.Add([string]$U) } }
            $Src.AddRange($Vals.ToArray())
            $State.AutoSrc     = $Src
            $State.AutoListRef = $Sync.MailboxList
        }
        $Boxes = @($TxtUser)
        if ($State.Mb -and -not $State.Mb.Form.IsDisposed) { $Boxes += $State.Mb.Txt }
        foreach ($B in $Boxes) { Set-AutoCompleteBox $B }
    }

    function Set-AutoCompleteBox($Box) {
        if (-not $Box -or -not $State.AutoSrc) { return }
        if ([object]::ReferenceEquals($Box.AutoCompleteCustomSource, $State.AutoSrc)) { return }
        $Box.AutoCompleteCustomSource = $State.AutoSrc
        $Box.AutoCompleteSource       = 'CustomSource'
        $Box.AutoCompleteMode         = 'Suggest'
    }

    # ---------- Fenster: Wer hat Zugriff auf Postfach X? (nicht modal) ----------
    function Show-MailboxWindow {
        param([string]$Address)

        if (-not $Sync.HasCache) {
            [void][System.Windows.Forms.MessageBox]::Show($Form, 'Bitte zuerst den Cache aufbauen.', 'Postfach-Zugriff', 'OK', 'Information')
            return
        }

        if (-not $State.Mb -or $State.Mb.Form.IsDisposed) {
            $F               = New-Object System.Windows.Forms.Form
            $F.Text          = 'Wer hat Zugriff auf ein Postfach?'
            $F.Size          = New-Object System.Drawing.Size(1000, 620)
            $F.MinimumSize   = New-Object System.Drawing.Size(860, 400)
            $F.StartPosition = 'Manual'
            $F.Location      = New-Object System.Drawing.Point -ArgumentList ([int]$Form.Left + 80), ([int]$Form.Top + 80)
            $F.Font          = $Font
            if ($Form.Icon) { $F.Icon = $Form.Icon }

            $P           = New-Object System.Windows.Forms.Panel
            $P.Dock      = 'Top'
            $P.Height    = 78
            $P.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

            $L          = New-Object System.Windows.Forms.Label
            $L.Text     = 'Postfach:'
            $L.Location = New-Object System.Drawing.Point(12, 16)
            $L.AutoSize = $true

            $T          = New-Object System.Windows.Forms.TextBox
            $T.Location = New-Object System.Drawing.Point(80, 12)
            $T.Width    = 380

            $BGo          = New-Object System.Windows.Forms.Button
            $BGo.Text     = 'Abfragen'
            $BGo.Location = New-Object System.Drawing.Point(470, 10)
            $BGo.Size     = New-Object System.Drawing.Size(100, 28)
            $BGo.Font     = $FontBold

            $BCp          = New-Object System.Windows.Forms.Button
            $BCp.Text     = 'Kopieren'
            $BCp.Location = New-Object System.Drawing.Point(580, 10)
            $BCp.Size     = New-Object System.Drawing.Size(100, 28)

            $BCsv          = New-Object System.Windows.Forms.Button
            $BCsv.Text     = 'CSV-Export'
            $BCsv.Location = New-Object System.Drawing.Point(690, 10)
            $BCsv.Size     = New-Object System.Drawing.Size(100, 28)

            $BFold          = New-Object System.Windows.Forms.Button
            $BFold.Text     = 'Ordnerrechte'
            $BFold.Location = New-Object System.Drawing.Point(800, 10)
            $BFold.Size     = New-Object System.Drawing.Size(110, 28)
            $BFold.add_Click({
                $A = Get-InputAddress $State.Mb.Txt.Text
                if ($A -and -not (Test-Busy)) {
                    $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'FolderQuery'; Mailbox = $A })
                    $State.Mb.Info.Text = "Lade Ordnerrechte von $A ..."
                }
            })

            $I           = New-Object System.Windows.Forms.Label
            $I.Location  = New-Object System.Drawing.Point(12, 50)
            $I.AutoSize  = $true
            $I.ForeColor = [System.Drawing.Color]::DimGray
            $I.Text      = 'Postfach eingeben und Abfragen klicken. Doppelklick auf einen Benutzer = dessen Rechte im Hauptfenster anzeigen.'

            $P.Controls.AddRange(@($L, $T, $BGo, $BCp, $BCsv, $BFold, $I))

            $Tb = New-Object System.Data.DataTable
            foreach ($C in 'Benutzer', 'Berechtigungen', 'Adresse', 'Typ', 'Quelle') { [void]$Tb.Columns.Add($C) }

            $G                       = New-Object System.Windows.Forms.DataGridView
            $G.Dock                  = 'Fill'
            $G.ReadOnly              = $true
            $G.AllowUserToAddRows    = $false
            $G.AllowUserToDeleteRows = $false
            $G.SelectionMode         = 'FullRowSelect'
            $G.AutoSizeColumnsMode   = 'Fill'
            $G.RowHeadersVisible     = $false
            $G.BackgroundColor       = [System.Drawing.Color]::White
            $G.DataSource            = $Tb

            $F.Controls.Add($G)
            $F.Controls.Add($P)
            $G.BringToFront()
            $F.AcceptButton = $BGo

            $Tm          = New-Object System.Windows.Forms.Timer
            $Tm.Interval = 200

            $State.Mb = @{ Form = $F; Txt = $T; Info = $I; Grid = $G; Table = $Tb; BtnGo = $BGo; Timer = $Tm; Version = $Sync.MbVersion }

            $BGo.add_Click({ Start-MbQuery })
            $Tm.add_Tick({
                if (-not $State.Mb) { return }
                $State.Mb.BtnGo.Enabled = -not ($Sync.Busy -or $Sync.Queue.Count -gt 0)
                if ($Sync.MbVersion -ne $State.Mb.Version) {
                    $State.Mb.Version = $Sync.MbVersion
                    Update-MbTable
                }
            })
            $BCp.add_Click({
                $Tbl = $State.Mb.Table
                if ($Tbl.Rows.Count -eq 0) { return }
                $Lines = @(($Tbl.Columns | ForEach-Object { $_.ColumnName }) -join "`t")
                foreach ($Rw in $Tbl.Rows) { $Lines += (@($Rw.ItemArray) | ForEach-Object { [string]$_ }) -join "`t" }
                [System.Windows.Forms.Clipboard]::SetText(($Lines -join "`r`n"))
                $State.Mb.Info.Text = "$($Tbl.Rows.Count) Zeile(n) kopiert (Excel-tauglich)"
            })
            $BCsv.add_Click({
                $Tbl = $State.Mb.Table
                if ($Tbl.Rows.Count -eq 0) { return }
                $Sd          = New-Object System.Windows.Forms.SaveFileDialog
                $Sd.Filter   = 'CSV-Datei (*.csv)|*.csv'
                $Nm          = ((Get-InputAddress $State.Mb.Txt.Text) -split '@')[0]
                $Sd.FileName = "Postfachzugriff_${Nm}_$(Get-Date -Format yyyyMMdd_HHmm).csv"
                if ($Sd.ShowDialog($State.Mb.Form) -eq 'OK') {
                    $Tbl | Select-Object Benutzer, Berechtigungen, Adresse, Typ, Quelle |
                        Export-Csv -Path $Sd.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ';'
                    $State.Mb.Info.Text = "Exportiert: $($Sd.FileName)"
                }
            })
            $G.add_CellDoubleClick({
                param($s, $e)
                if ($e.RowIndex -lt 0) { return }
                $A = [string]$s.Rows[$e.RowIndex].Cells['Adresse'].Value
                if (-not $A) { return }
                $TxtUser.Text = $A
                $Form.Activate()
                if (-not (Test-Busy)) {
                    $State.Raw = @(); Update-Table
                    $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'Query'; User = $A })
                }
            })
            $F.add_FormClosed({
                if ($State.Mb) { $State.Mb.Timer.Stop(); $State.Mb.Timer.Dispose() }
                $State.Mb = $null
            })

            Set-AutoCompleteBox $T
            $Tm.Start()
            $F.Show($Form)
        } else {
            if ($State.Mb.Form.WindowState -eq 'Minimized') { $State.Mb.Form.WindowState = 'Normal' }
            $State.Mb.Form.Activate()
        }

        if ($Address) {
            $State.Mb.Txt.Text = $Address
            Start-MbQuery
        } else {
            [void]$State.Mb.Txt.Focus()
        }
    }

    function Start-MbQuery {
        if (-not $State.Mb) { return }
        $A = Get-InputAddress $State.Mb.Txt.Text
        if (-not $A) { return }
        if ($Sync.Busy -or $Sync.Queue.Count -gt 0) {
            $State.Mb.Info.Text = 'Bitte warten - im Hauptfenster läuft noch eine Aktion...'
            return
        }
        $State.Mb.Table.Rows.Clear()
        $State.Mb.Info.Text = "Lade Zugriffe auf $A ..."
        $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'MailboxQuery'; Mailbox = $A })
    }

    function Update-MbTable {
        $W   = $State.Mb
        $Res = $Sync.MbResult
        $W.Table.BeginLoadData()
        $W.Table.Rows.Clear()
        if (-not $Res) {
            $W.Table.EndLoadData()
            $W.Info.Text = 'Abfrage fehlgeschlagen - siehe Fehlermeldung im Hauptfenster.'
            return
        }
        $Order = @{ FullAccess = 1; SendAs = 2; SendOnBehalf = 3; Kalender = 4; Buchungsstellvertreter = 5 }
        $Groups = @($Res.Entries | Where-Object { $_ }) |
            Group-Object { if ($_.Adresse) { ([string]$_.Adresse).ToLower() } else { [string]$_.Benutzer } }
        foreach ($Gr in ($Groups | Sort-Object { [string]$_.Group[0].Benutzer })) {
            $First = $Gr.Group[0]
            $Perms = @($Gr.Group | Sort-Object { $Order[[string]$_.Berechtigung] } | ForEach-Object {
                if ($_.Berechtigung -eq 'Kalender') { "Kalender ($($_.Rechte))" } else { [string]$_.Berechtigung }
            } | Select-Object -Unique) -join ', '
            $Src = @($Gr.Group | ForEach-Object { [string]$_.Quelle } | Select-Object -Unique) -join '; '
            [void]$W.Table.Rows.Add($First.Benutzer, $Perms, $First.Adresse, $First.Typ, $Src)
        }
        $W.Table.EndLoadData()
        $CalNote = if (-not $Res.CalendarInCache) { '  |  Kalenderrechte nicht im Cache (Haken "Kalenderrechte im Cache")' } else { '' }
        $W.Info.Text = "$($Res.Name) <$($Res.Address)> [$($Res.Type)]: $(@($Groups).Count) Benutzer/Gruppen mit Zugriff$CalNote"
        $W.Form.Text = "Wer hat Zugriff auf: $($Res.Name)"
    }

    # ---------- Dialog: Rechte kopieren (Onboarding "wie Kollege") ----------
    function Invoke-CopyDialog {
        if (Test-Busy) { return }
        $Src    = [string]$Sync.ResultUser
        $Direct = @($State.Raw | Where-Object { $_ -and $_.Quelle -eq 'direkt' -and $_.Adresse })
        if (-not $Src -or $Direct.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show($Form, 'Bitte zuerst den Vorlage-Benutzer abfragen (der Kollege, dessen Rechte übernommen werden sollen).', 'Rechte kopieren', 'OK', 'Information')
            return
        }
        $GrpCount = @($State.Raw | Where-Object { $_ -and ([string]$_.Quelle) -like 'Gruppe*' }).Count
        $InhCount = @($State.Raw | Where-Object { $_ -and ([string]$_.Quelle) -like '*vererbt*' }).Count
        $Names = @{ FullAccess = 'Lesen und Verwalten (Vollzugriff)'; SendAs = 'Senden als'; SendOnBehalf = 'Senden im Auftrag von'; Kalender = 'Kalenderberechtigung'; Buchungsstellvertreter = 'Buchungsstellvertreter' }
        $Order = @{ FullAccess = 1; SendAs = 2; SendOnBehalf = 3; Kalender = 4; Buchungsstellvertreter = 5 }

        $Dlg                 = New-Object System.Windows.Forms.Form
        $Dlg.Text            = 'Rechte kopieren (Onboarding)'
        $Dlg.Size            = New-Object System.Drawing.Size(780, 700)
        $Dlg.StartPosition   = 'CenterParent'
        $Dlg.FormBorderStyle = 'FixedDialog'
        $Dlg.MaximizeBox     = $false
        $Dlg.MinimizeBox     = $false
        $Dlg.Font            = $Font
        if ($Form.Icon) { $Dlg.Icon = $Form.Icon }

        $LblHead          = New-Object System.Windows.Forms.Label
        $LblHead.Text     = "Vorlage:  $Src"
        $LblHead.Font     = $FontBold
        $LblHead.Location = New-Object System.Drawing.Point(12, 12)
        $LblHead.AutoSize = $true

        $LblT          = New-Object System.Windows.Forms.Label
        $LblT.Text     = 'Kopieren auf:'
        $LblT.Location = New-Object System.Drawing.Point(12, 46)
        $LblT.AutoSize = $true

        $TxtT          = New-Object System.Windows.Forms.TextBox
        $TxtT.Location = New-Object System.Drawing.Point(110, 43)
        $TxtT.Width    = 420
        Set-AutoCompleteBox $TxtT

        $Lv               = New-Object System.Windows.Forms.ListView
        $Lv.Location      = New-Object System.Drawing.Point(12, 80)
        $Lv.Size          = New-Object System.Drawing.Size(740, 390)
        $Lv.View          = 'Details'
        $Lv.CheckBoxes    = $true
        $Lv.FullRowSelect = $true
        [void]$Lv.Columns.Add('Postfach', 250)
        [void]$Lv.Columns.Add('Berechtigung', 230)
        [void]$Lv.Columns.Add('Rechte', 150)
        [void]$Lv.Columns.Add('Typ', 90)

        $BAll          = New-Object System.Windows.Forms.Button
        $BAll.Text     = 'Alle'
        $BAll.Location = New-Object System.Drawing.Point(12, 478)
        $BAll.Size     = New-Object System.Drawing.Size(80, 26)

        $BNone          = New-Object System.Windows.Forms.Button
        $BNone.Text     = 'Keine'
        $BNone.Location = New-Object System.Drawing.Point(100, 478)
        $BNone.Size     = New-Object System.Drawing.Size(80, 26)

        $CkAuto          = New-Object System.Windows.Forms.CheckBox
        $CkAuto.Text     = 'Vollzugriff mit AutoMapping (Postfach erscheint automatisch in Outlook)'
        $CkAuto.Location = New-Object System.Drawing.Point(200, 482)
        $CkAuto.AutoSize = $true
        $CkAuto.Checked  = $true

        $LblInfo           = New-Object System.Windows.Forms.Label
        $LblInfo.Location  = New-Object System.Drawing.Point(12, 516)
        $LblInfo.Size      = New-Object System.Drawing.Size(740, 54)
        $LblInfo.ForeColor = [System.Drawing.Color]::DimGray
        $LblInfo.Text      = "Kopiert werden nur direkte Rechte.`n" +
                             "Nicht kopiert: $GrpCount Recht(e) über Gruppen (Gruppenmitgliedschaft separat setzen), $InhCount vererbte Recht(e)."

        $LblSum           = New-Object System.Windows.Forms.Label
        $LblSum.Location  = New-Object System.Drawing.Point(12, 580)
        $LblSum.AutoSize  = $true
        $LblSum.Font      = $FontBold

        $BtnOk              = New-Object System.Windows.Forms.Button
        $BtnOk.Text         = 'Kopieren'
        $BtnOk.Location     = New-Object System.Drawing.Point(532, 616)
        $BtnOk.Size         = New-Object System.Drawing.Size(110, 30)
        $BtnOk.BackColor    = [System.Drawing.Color]::FromArgb(0, 120, 215)
        $BtnOk.ForeColor    = [System.Drawing.Color]::White
        $BtnOk.FlatStyle    = 'Flat'
        $BtnOk.Font         = $FontBold
        $BtnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK

        $BtnNo              = New-Object System.Windows.Forms.Button
        $BtnNo.Text         = 'Abbrechen'
        $BtnNo.Location     = New-Object System.Drawing.Point(652, 616)
        $BtnNo.Size         = New-Object System.Drawing.Size(100, 30)
        $BtnNo.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

        $UpdateSum = {
            $N  = $Lv.CheckedItems.Count
            $Tg = Get-InputAddress $TxtT.Text
            $LblSum.Text   = if ($Tg) { "$N Berechtigung(en) werden auf $Tg übertragen." } else { "$N Berechtigung(en) ausgewählt - Zielbenutzer eintragen." }
            $BtnOk.Enabled = ($N -gt 0) -and [bool]$Tg
        }

        foreach ($E in ($Direct | Sort-Object { [string]$_.Postfach }, { $Order[[string]$_.Berechtigung] })) {
            $Li = New-Object System.Windows.Forms.ListViewItem([string]$E.Postfach)
            [void]$Li.SubItems.Add($(if ($Names[[string]$E.Berechtigung]) { $Names[[string]$E.Berechtigung] } else { [string]$E.Berechtigung }))
            [void]$Li.SubItems.Add([string]$E.Rechte)
            [void]$Li.SubItems.Add([string]$E.Typ)
            $Li.Tag     = $E
            $Li.Checked = $true
            [void]$Lv.Items.Add($Li)
        }

        $Dlg.Controls.AddRange(@($LblHead, $LblT, $TxtT, $Lv, $BAll, $BNone, $CkAuto, $LblInfo, $LblSum, $BtnOk, $BtnNo))
        $Dlg.CancelButton = $BtnNo

        $Lv.add_ItemChecked($UpdateSum)
        $TxtT.add_TextChanged($UpdateSum)
        $BAll.add_Click({ foreach ($Li in $Lv.Items) { $Li.Checked = $true } })
        $BNone.add_Click({ foreach ($Li in $Lv.Items) { $Li.Checked = $false } })
        $Dlg.add_Shown({ [void]$TxtT.Focus() })
        & $UpdateSum

        $Result = $Dlg.ShowDialog($Form)
        $Target = Get-InputAddress $TxtT.Text
        $Chosen = @($Lv.CheckedItems | ForEach-Object { $_.Tag })
        $Auto   = [bool]$CkAuto.Checked
        $Dlg.Dispose()
        if ($Result -ne [System.Windows.Forms.DialogResult]::OK -or -not $Target -or $Chosen.Count -eq 0) { return }
        if ($Target -eq $Src) {
            [void][System.Windows.Forms.MessageBox]::Show($Form, 'Ziel und Vorlage sind identisch.', 'Rechte kopieren', 'OK', 'Warning')
            return
        }

        $Items = @($Chosen | Group-Object { [string]$_.Adresse } | ForEach-Object {
            $F   = $_.Group[0]
            $Cal = @($_.Group | Where-Object { $_.Berechtigung -eq 'Kalender' } | Select-Object -First 1)
            [PSCustomObject]@{
                Adresse      = [string]$F.Adresse
                Postfach     = [string]$F.Postfach
                IsRoom       = ([string]$F.Typ) -in 'Room', 'Equipment'
                Perms        = @($_.Group | ForEach-Object { [string]$_.Berechtigung } | Select-Object -Unique)
                CalendarRole = if ($Cal) { [string]$Cal[0].Rechte } else { '' }
            }
        })

        $Ans = [System.Windows.Forms.MessageBox]::Show($Form,
            "$($Chosen.Count) Berechtigung(en) auf $($Items.Count) Postfächern`nvon  $Src`nauf  $Target`nübertragen?",
            'Bestätigen', 'YesNo', 'Question', [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($Ans -ne 'Yes') { return }

        $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'Grant'; User = $Target; Items = $Items; Options = @{ AutoMapping = $Auto }; Title = "Ergebnis: Rechte kopiert von $Src" })
        $LblStatus.Text = "Kopiere Rechte von $Src auf $Target..."
    }

    # ---------- Dialog: Offboarding ----------
    function Invoke-OffboardDialog {
        if (Test-Busy) { return }
        $User   = [string]$Sync.ResultUser
        $Removable = @('FullAccess', 'SendAs', 'SendOnBehalf', 'Kalender', 'Buchungsstellvertreter')
        $Direct = @($State.Raw | Where-Object { $_ -and $_.Quelle -eq 'direkt' -and $_.Adresse -and $Removable -contains [string]$_.Berechtigung })
        if (-not $User -or $Direct.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show($Form, 'Bitte zuerst den ausscheidenden Benutzer abfragen.', 'Offboarding', 'OK', 'Information')
            return
        }
        $GrpCount = @($State.Raw | Where-Object { $_ -and ([string]$_.Quelle) -like 'Gruppe*' }).Count

        # Nachweis-CSV (alle Rechte inkl. Gruppen) VOR dem Entfernen
        $Dir = Join-Path $Sync.DataDir 'Offboarding'
        if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir -Force | Out-Null }
        $Csv = Join-Path $Dir ("Offboarding_{0}_{1:yyyyMMdd_HHmm}.csv" -f (($User -split '@')[0]), (Get-Date))
        try {
            @($State.Raw | Where-Object { $_ }) | Select-Object Postfach, Adresse, Typ, Berechtigung, Rechte, Quelle |
                Export-Csv -Path $Csv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show($Form, "Nachweis-CSV konnte nicht geschrieben werden:`n$($_.Exception.Message)`n`nOffboarding abgebrochen.", 'Offboarding', 'OK', 'Error')
            return
        }

        $Items = @($Direct | Group-Object { [string]$_.Adresse } | ForEach-Object {
            [PSCustomObject]@{
                Adresse  = [string]$_.Group[0].Adresse
                Postfach = [string]$_.Group[0].Postfach
                Perms    = @($_.Group | ForEach-Object { [string]$_.Berechtigung } | Select-Object -Unique)
            }
        })
        $Lines = @($Items | Select-Object -First 20 | ForEach-Object { "  $($_.Postfach): $($_.Perms -join ', ')" })
        if ($Items.Count -gt 20) { $Lines += "  ... und $($Items.Count - 20) weitere" }
        $GrpNote = if ($GrpCount) { "`n`nHinweis: $GrpCount Recht(e) über Gruppen bleiben bestehen - Gruppenmitgliedschaften separat entfernen." } else { '' }

        $Ans = [System.Windows.Forms.MessageBox]::Show($Form,
            "OFFBOARDING:  $User`n`nAlle $($Direct.Count) direkten Berechtigungen auf $($Items.Count) Postfächern werden entfernt:`n`n" +
            ($Lines -join "`n") + $GrpNote +
            "`n`nNachweis gespeichert:`n$Csv`n`nJetzt entfernen?",
            'Offboarding bestätigen', 'YesNo', 'Warning', [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($Ans -ne 'Yes') { return }

        $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'Remove'; User = $User; Items = $Items; Title = "Ergebnis: Offboarding $User" })
        $LblStatus.Text = "Offboarding $User - entferne Berechtigungen..."
    }

    # ---------- Fenster: Änderungsprotokoll ----------
    function Show-LogWindow {
        $LogFile = Join-Path $Sync.DataDir 'aenderungen.log'

        $Dlg               = New-Object System.Windows.Forms.Form
        $Dlg.Text          = 'Änderungsprotokoll'
        $Dlg.Size          = New-Object System.Drawing.Size(1100, 600)
        $Dlg.StartPosition = 'CenterParent'
        $Dlg.Font          = $Font
        if ($Form.Icon) { $Dlg.Icon = $Form.Icon }

        $P        = New-Object System.Windows.Forms.Panel
        $P.Dock   = 'Top'
        $P.Height = 44

        $L          = New-Object System.Windows.Forms.Label
        $L.Text     = 'Filter:'
        $L.Location = New-Object System.Drawing.Point(12, 14)
        $L.AutoSize = $true

        $TF          = New-Object System.Windows.Forms.TextBox
        $TF.Location = New-Object System.Drawing.Point(60, 10)
        $TF.Width    = 320

        $BO          = New-Object System.Windows.Forms.Button
        $BO.Text     = 'Ordner öffnen'
        $BO.Location = New-Object System.Drawing.Point(395, 8)
        $BO.Size     = New-Object System.Drawing.Size(120, 28)

        $LC           = New-Object System.Windows.Forms.Label
        $LC.Location  = New-Object System.Drawing.Point(530, 14)
        $LC.AutoSize  = $true
        $LC.ForeColor = [System.Drawing.Color]::DimGray

        $P.Controls.AddRange(@($L, $TF, $BO, $LC))

        $Tb = New-Object System.Data.DataTable
        foreach ($C in 'Zeit', 'Admin', 'Benutzer', 'Postfach', 'Recht', 'Status', 'Meldung') { [void]$Tb.Columns.Add($C) }
        if (Test-Path $LogFile) {
            $Rows = @(Get-Content $LogFile -Encoding UTF8 | Where-Object { $_ })
            [array]::Reverse($Rows)
            foreach ($Ln in $Rows) {
                $Pt = $Ln -split ';', 7
                while ($Pt.Count -lt 7) { $Pt += '' }
                [void]$Tb.Rows.Add($Pt[0], $Pt[1], $Pt[2], $Pt[3], $Pt[4], $Pt[5], $Pt[6])
            }
        }

        $G                     = New-Object System.Windows.Forms.DataGridView
        $G.Dock                = 'Fill'
        $G.ReadOnly            = $true
        $G.AllowUserToAddRows  = $false
        $G.SelectionMode       = 'FullRowSelect'
        $G.AutoSizeColumnsMode = 'Fill'
        $G.RowHeadersVisible   = $false
        $G.BackgroundColor     = [System.Drawing.Color]::White
        $G.DataSource          = $Tb
        $G.add_CellFormatting({
            param($s, $e)
            if ($s.Columns[$e.ColumnIndex].Name -eq 'Status') {
                $V = [string]$e.Value
                if ($V -like 'FEHLER*')   { $e.CellStyle.ForeColor = [System.Drawing.Color]::Firebrick }
                elseif ($V -eq 'ERTEILT') { $e.CellStyle.ForeColor = [System.Drawing.Color]::ForestGreen }
                elseif ($V -eq 'OK')      { $e.CellStyle.ForeColor = [System.Drawing.Color]::DarkOrange }
            }
        })

        $Dlg.Controls.Add($G)
        $Dlg.Controls.Add($P)
        $G.BringToFront()

        $ApplyFilter = {
            $Q = ($TF.Text.Trim() -replace '([\[\]\*%])', '[$1]') -replace "'", "''"
            $Tb.DefaultView.RowFilter = if ($Q) {
                "Zeit LIKE '%$Q%' OR Admin LIKE '%$Q%' OR Benutzer LIKE '%$Q%' OR Postfach LIKE '%$Q%' OR Recht LIKE '%$Q%' OR Status LIKE '%$Q%' OR Meldung LIKE '%$Q%'"
            } else { '' }
            $LC.Text = "$($Tb.DefaultView.Count) von $($Tb.Rows.Count) Einträgen  (OK = entfernt, ERTEILT = vergeben)"
        }
        $TF.add_TextChanged($ApplyFilter)
        $BO.add_Click({ if (Test-Path $Sync.DataDir) { Start-Process explorer.exe $Sync.DataDir } })
        & $ApplyFilter

        [void]$Dlg.ShowDialog($Form)
        $Dlg.Dispose()
    }

    # ---------- v3.9: Tabellenfenster (nicht modal) für Weiterleitungen / Ordnerrechte ----------
    function Show-TableWindow {
        param([string]$Title, [string]$Info, [string[]]$Columns, [object[]]$Rows)

        $F               = New-Object System.Windows.Forms.Form
        $F.Text          = $Title
        $F.Size          = New-Object System.Drawing.Size(1100, 560)
        $F.MinimumSize   = New-Object System.Drawing.Size(700, 300)
        $F.StartPosition = 'Manual'
        $F.Location      = New-Object System.Drawing.Point -ArgumentList ([int]$Form.Left + 120), ([int]$Form.Top + 120)
        $F.Font          = $Font
        if ($Form.Icon) { $F.Icon = $Form.Icon }

        $P           = New-Object System.Windows.Forms.Panel
        $P.Dock      = 'Top'
        $P.Height    = 46
        $P.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

        $I           = New-Object System.Windows.Forms.Label
        $I.Location  = New-Object System.Drawing.Point(12, 15)
        $I.AutoSize  = $true
        $I.Font      = $FontBold
        $I.Text      = $Info

        $BCp          = New-Object System.Windows.Forms.Button
        $BCp.Text     = 'Kopieren'
        $BCp.Size     = New-Object System.Drawing.Size(100, 28)
        $BCp.Anchor   = 'Top, Right'
        $BCp.Location = New-Object System.Drawing.Point(870, 9)

        $BCsv          = New-Object System.Windows.Forms.Button
        $BCsv.Text     = 'CSV-Export'
        $BCsv.Size     = New-Object System.Drawing.Size(100, 28)
        $BCsv.Anchor   = 'Top, Right'
        $BCsv.Location = New-Object System.Drawing.Point(976, 9)

        $P.Controls.AddRange(@($I, $BCp, $BCsv))

        $Tb = New-Object System.Data.DataTable
        foreach ($C in $Columns) { [void]$Tb.Columns.Add($C) }
        foreach ($R in @($Rows | Where-Object { $_ })) {
            $Vals = [object[]]@(foreach ($C in $Columns) { [string]$R.$C })
            [void]$Tb.Rows.Add($Vals)
        }

        $G                       = New-Object System.Windows.Forms.DataGridView
        $G.Dock                  = 'Fill'
        $G.ReadOnly              = $true
        $G.AllowUserToAddRows    = $false
        $G.AllowUserToDeleteRows = $false
        $G.SelectionMode         = 'FullRowSelect'
        $G.AutoSizeColumnsMode   = 'Fill'
        $G.RowHeadersVisible     = $false
        $G.BackgroundColor       = [System.Drawing.Color]::White
        $G.DataSource            = $Tb

        # Auffällige Einträge markieren: externe Weiterleitung (rot), Standard-Ordnerrecht für alle (gelb)
        $G.add_DataBindingComplete({
            param($s, $e)
            foreach ($Rw in $s.Rows) {
                if ($s.Columns.Contains('Extern') -and ([string]$Rw.Cells['Extern'].Value) -eq 'JA') {
                    $Rw.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(253, 231, 233)
                } elseif ($s.Columns.Contains('Hinweis') -and [string]$Rw.Cells['Hinweis'].Value) {
                    $Rw.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 243, 205)
                }
            }
        })

        $F.Controls.Add($G)
        $F.Controls.Add($P)
        $G.BringToFront()
        $F.Tag = @{ Table = $Tb; Info = $I; Title = $Title }

        $BCp.add_Click({
            param($s, $e)
            $Tbl = $s.FindForm().Tag.Table
            if ($Tbl.Rows.Count -eq 0) { return }
            $Lines = @(($Tbl.Columns | ForEach-Object { $_.ColumnName }) -join "`t")
            foreach ($Rw in $Tbl.Rows) { $Lines += (@($Rw.ItemArray) | ForEach-Object { [string]$_ }) -join "`t" }
            [System.Windows.Forms.Clipboard]::SetText(($Lines -join "`r`n"))
            $s.FindForm().Tag.Info.Text = "$($Tbl.Rows.Count) Zeile(n) kopiert (Excel-tauglich)"
        })
        $BCsv.add_Click({
            param($s, $e)
            $Fm  = $s.FindForm()
            $Tbl = $Fm.Tag.Table
            if ($Tbl.Rows.Count -eq 0) { return }
            $Sd          = New-Object System.Windows.Forms.SaveFileDialog
            $Sd.Filter   = 'CSV-Datei (*.csv)|*.csv'
            $Nm          = (([string]$Fm.Tag.Title) -replace '[^\w@.-]+', '_').Trim('_')
            $Sd.FileName = "${Nm}_$(Get-Date -Format yyyyMMdd_HHmm).csv"
            if ($Sd.ShowDialog($Fm) -eq 'OK') {
                $Cols = @($Tbl.Columns | ForEach-Object { $_.ColumnName })
                $Tbl | Select-Object $Cols | Export-Csv -Path $Sd.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ';'
                $Fm.Tag.Info.Text = "Exportiert: $($Sd.FileName)"
            }
        })
        $F.Show($Form)
    }

    # ---------- v3.9: Eingabedialog (Postfach wählen) ----------
    function Read-AddressDialog {
        param([string]$Title, [string]$Prompt, [string]$Default)
        $D                 = New-Object System.Windows.Forms.Form
        $D.Text            = $Title
        $D.Size            = New-Object System.Drawing.Size(520, 170)
        $D.StartPosition   = 'CenterParent'
        $D.FormBorderStyle = 'FixedDialog'
        $D.MaximizeBox     = $false
        $D.MinimizeBox     = $false
        $D.Font            = $Font
        if ($Form.Icon) { $D.Icon = $Form.Icon }

        $L          = New-Object System.Windows.Forms.Label
        $L.Text     = $Prompt
        $L.Location = New-Object System.Drawing.Point(14, 14)
        $L.AutoSize = $true

        $T          = New-Object System.Windows.Forms.TextBox
        $T.Location = New-Object System.Drawing.Point(14, 40)
        $T.Width    = 476
        $T.Text     = $Default
        Set-AutoCompleteBox $T

        $Ok              = New-Object System.Windows.Forms.Button
        $Ok.Text         = 'OK'
        $Ok.Location     = New-Object System.Drawing.Point(300, 82)
        $Ok.Size         = New-Object System.Drawing.Size(90, 28)
        $Ok.DialogResult = [System.Windows.Forms.DialogResult]::OK

        $No              = New-Object System.Windows.Forms.Button
        $No.Text         = 'Abbrechen'
        $No.Location     = New-Object System.Drawing.Point(400, 82)
        $No.Size         = New-Object System.Drawing.Size(90, 28)
        $No.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

        $D.Controls.AddRange(@($L, $T, $Ok, $No))
        $D.AcceptButton = $Ok
        $D.CancelButton = $No
        $Res = $D.ShowDialog($Form)
        $Val = Get-InputAddress $T.Text
        $D.Dispose()
        if ($Res -ne [System.Windows.Forms.DialogResult]::OK) { return '' }
        return $Val
    }

    function Get-CurrentUserInput {
        $U = Get-InputAddress $TxtUser.Text
        if (-not $U) { $U = [string]$Sync.ResultUser }
        return $U
    }

    function Invoke-ForwardCheck {
        if (Test-Busy) { return }
        $U = Get-CurrentUserInput
        if (-not $U) { $U = Read-AddressDialog -Title 'Weiterleitungen und Regeln' -Prompt 'Benutzer / Postfach:' }
        if (-not $U) { return }
        $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'ForwardQuery'; User = $U })
        $LblStatus.Text = "Lade Weiterleitungen und Regeln von $U ..."
    }

    function Invoke-FolderCheck {
        param([string]$Address)
        if (Test-Busy) { return }
        $A = $Address
        if (-not $A -and $Grid.SelectedRows.Count -eq 1) { $A = [string]$Grid.SelectedRows[0].Cells['Adresse'].Value }
        $A = Read-AddressDialog -Title 'Ordnerrechte' -Prompt 'Postfach, dessen Ordnerrechte angezeigt werden sollen:' -Default $A
        if (-not $A) { return }
        $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'FolderQuery'; Mailbox = $A })
        $LblStatus.Text = "Lade Ordnerrechte von $A (kann bei vielen Ordnern 1-2 min dauern)..."
    }

    function Invoke-GroupCheck {
        if (Test-Busy) { return }
        $U = Get-CurrentUserInput
        if (-not $U) { $U = Read-AddressDialog -Title 'Gruppenmitgliedschaften' -Prompt 'Benutzer:' }
        if (-not $U) { return }
        $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'GroupQuery'; User = $U })
        $LblStatus.Text = "Lade Gruppen von $U ..."
    }

    function Invoke-TenantReport {
        if (Test-Busy) { return }
        if (-not $Sync.HasCache) { [void](Request-CacheBuild); return }
        $Fb = New-Object System.Windows.Forms.FolderBrowserDialog
        $Fb.Description  = 'Zielordner für den Gesamtreport (HTML + CSV)'
        $Fb.SelectedPath = [Environment]::GetFolderPath('MyDocuments')
        if ($Fb.ShowDialog($Form) -ne 'OK') { return }
        $Ans = [System.Windows.Forms.MessageBox]::Show($Form,
            "Der Report liest zusätzlich 'Senden als', 'Senden im Auftrag von' und Weiterleitungen für ALLE Postfächer live.`nDas kann einige Minuten dauern.`n`nStarten?",
            'Tenant-Gesamtreport', 'YesNo', 'Question')
        if ($Ans -ne 'Yes') { return }
        $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'TenantReport'; Folder = $Fb.SelectedPath })
        $LblStatus.Text = 'Erstelle Tenant-Gesamtreport...'
    }

    # ---------- v3.9: Dialog Gruppenmitgliedschaften ----------
    function Invoke-GroupDialog {
        param($Data)
        if (-not $Data) { return }
        $Member = @($Data.Member | Where-Object { $_ })
        $All    = @($Data.All | Where-Object { $_ })
        $MemberSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($M in $Member) { [void]$MemberSet.Add([string]$M.Address) }
        $AddSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        $ByAddr = @{}
        foreach ($G in $All) { if ($G.Address) { $ByAddr[[string]$G.Address] = $G } }
        foreach ($G in $Member) { if ($G.Address -and -not $ByAddr.ContainsKey([string]$G.Address)) { $ByAddr[[string]$G.Address] = $G } }

        $Dlg                 = New-Object System.Windows.Forms.Form
        $Dlg.Text            = "Gruppenmitgliedschaften: $($Data.UserName)"
        $Dlg.Size            = New-Object System.Drawing.Size(1040, 640)
        $Dlg.StartPosition   = 'CenterParent'
        $Dlg.FormBorderStyle = 'FixedDialog'
        $Dlg.MaximizeBox     = $false
        $Dlg.MinimizeBox     = $false
        $Dlg.Font            = $Font
        if ($Form.Icon) { $Dlg.Icon = $Form.Icon }

        $L1          = New-Object System.Windows.Forms.Label
        $L1.Text     = "Aktuelle Mitgliedschaften ($($Member.Count)) - Haken = ENTFERNEN"
        $L1.Font     = $FontBold
        $L1.Location = New-Object System.Drawing.Point(12, 12)
        $L1.AutoSize = $true

        $Lv1               = New-Object System.Windows.Forms.ListView
        $Lv1.Location      = New-Object System.Drawing.Point(12, 40)
        $Lv1.Size          = New-Object System.Drawing.Size(490, 470)
        $Lv1.View          = 'Details'
        $Lv1.CheckBoxes    = $true
        $Lv1.FullRowSelect = $true
        [void]$Lv1.Columns.Add('Gruppe', 300)
        [void]$Lv1.Columns.Add('Typ', 160)
        foreach ($M in $Member) {
            $Li = New-Object System.Windows.Forms.ListViewItem([string]$M.Name)
            [void]$Li.SubItems.Add([string]$M.TypeLabel)
            $Li.Tag = [string]$M.Address
            [void]$Lv1.Items.Add($Li)
        }

        $L2          = New-Object System.Windows.Forms.Label
        $L2.Text     = 'Gruppe hinzufügen - suchen und Haken setzen'
        $L2.Font     = $FontBold
        $L2.Location = New-Object System.Drawing.Point(520, 12)
        $L2.AutoSize = $true

        $Ts          = New-Object System.Windows.Forms.TextBox
        $Ts.Location = New-Object System.Drawing.Point(520, 40)
        $Ts.Width    = 490

        $Lv2               = New-Object System.Windows.Forms.ListView
        $Lv2.Location      = New-Object System.Drawing.Point(520, 70)
        $Lv2.Size          = New-Object System.Drawing.Size(490, 440)
        $Lv2.View          = 'Details'
        $Lv2.CheckBoxes    = $true
        $Lv2.FullRowSelect = $true
        [void]$Lv2.Columns.Add('Gruppe', 300)
        [void]$Lv2.Columns.Add('Typ', 160)

        $LSum           = New-Object System.Windows.Forms.Label
        $LSum.Location  = New-Object System.Drawing.Point(12, 524)
        $LSum.AutoSize  = $true
        $LSum.ForeColor = [System.Drawing.Color]::DimGray

        $BtnOk              = New-Object System.Windows.Forms.Button
        $BtnOk.Text         = 'Übernehmen'
        $BtnOk.Location     = New-Object System.Drawing.Point(790, 556)
        $BtnOk.Size         = New-Object System.Drawing.Size(110, 30)
        $BtnOk.BackColor    = [System.Drawing.Color]::FromArgb(0, 120, 215)
        $BtnOk.ForeColor    = [System.Drawing.Color]::White
        $BtnOk.FlatStyle    = 'Flat'
        $BtnOk.Font         = $FontBold
        $BtnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK

        $BtnNo              = New-Object System.Windows.Forms.Button
        $BtnNo.Text         = 'Schließen'
        $BtnNo.Location     = New-Object System.Drawing.Point(910, 556)
        $BtnNo.Size         = New-Object System.Drawing.Size(100, 30)
        $BtnNo.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

        $Dlg.Controls.AddRange(@($L1, $Lv1, $L2, $Ts, $Lv2, $LSum, $BtnOk, $BtnNo))
        $Dlg.CancelButton = $BtnNo

        $GState = @{ Loading = $false }
        $UpdSum = {
            $NR = $Lv1.CheckedItems.Count
            $NA = $AddSet.Count
            $LSum.Text     = "Hinzufügen: $NA   Entfernen: $NR   (dynamische Gruppen und Gruppen ohne E-Mail sind nicht änderbar)"
            $BtnOk.Enabled = ($NA + $NR) -gt 0
        }
        $Fill = {
            $GState.Loading = $true
            $Lv2.BeginUpdate()
            $Lv2.Items.Clear()
            $Q = [System.Management.Automation.WildcardPattern]::Escape($Ts.Text.Trim())
            $Shown = 0
            foreach ($G in $All) {
                if ($MemberSet.Contains([string]$G.Address)) { continue }
                if ($Q -and -not ($G.Name -like "*$Q*" -or $G.Address -like "*$Q*") -and -not $AddSet.Contains([string]$G.Address)) { continue }
                if ($Shown -ge 500 -and -not $AddSet.Contains([string]$G.Address)) { continue }
                $Li = New-Object System.Windows.Forms.ListViewItem([string]$G.Name)
                [void]$Li.SubItems.Add([string]$G.TypeLabel)
                $Li.Tag     = [string]$G.Address
                $Li.Checked = $AddSet.Contains([string]$G.Address)
                [void]$Lv2.Items.Add($Li)
                $Shown++
            }
            $Lv2.EndUpdate()
            $GState.Loading = $false
            & $UpdSum
        }
        $Lv1.add_ItemChecked($UpdSum)
        $Lv2.add_ItemChecked({
            param($s, $e)
            if ($GState.Loading) { return }
            if ($e.Item.Checked) { [void]$AddSet.Add([string]$e.Item.Tag) } else { [void]$AddSet.Remove([string]$e.Item.Tag) }
            & $UpdSum
        })
        $Ts.add_TextChanged($Fill)
        $Dlg.add_Shown({ [void]$Ts.Focus() })
        & $Fill

        $Result = $Dlg.ShowDialog($Form)
        $RemoveList = @($Lv1.CheckedItems | ForEach-Object { $ByAddr[[string]$_.Tag] } | Where-Object { $_ })
        $AddList    = @($AddSet | ForEach-Object { $ByAddr[$_] } | Where-Object { $_ })
        $Dlg.Dispose()
        if ($Result -ne [System.Windows.Forms.DialogResult]::OK -or ($RemoveList.Count + $AddList.Count) -eq 0) { return }

        $Txt = @()
        foreach ($G in $AddList)    { $Txt += "  + $($G.Name) ($($G.TypeLabel))" }
        foreach ($G in $RemoveList) { $Txt += "  - $($G.Name) ($($G.TypeLabel))" }
        $Ans = [System.Windows.Forms.MessageBox]::Show($Form,
            "Gruppen für $($Data.UserName) ändern?`n`n$($Txt -join "`n")",
            'Bestätigen', 'YesNo', 'Question', [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($Ans -ne 'Yes') { return }

        $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'GroupChange'; User = $Data.User; Add = $AddList; Remove = $RemoveList })
        $LblStatus.Text = 'Ändere Gruppenmitgliedschaften...'
    }

    function Update-CacheLabel {
        $LblCache.Text = $Sync.CacheInfo
        if (-not $Sync.CanWrite) { $LblCache.Text = "NUR LESEN (keine Exchange-Rolle zum Ändern)  |  " + $LblCache.Text }
        if (-not $Sync.HasCache -or $Sync.CacheStale) {
            $LblCache.ForeColor = [System.Drawing.Color]::DarkOrange
            if ($Sync.HasCache -and -not $Sync.CacheHasRooms) { $LblCache.Text += "  - ohne Raumpostfach-Rechte, bitte 'Cache aktualisieren'" }
            elseif ($Sync.HasCache) { $LblCache.Text += "  - älter als $($Sync.MaxCacheAgeHours) h, Aktualisierung empfohlen" }
        } else {
            $LblCache.ForeColor = [System.Drawing.Color]::DimGray
        }
    }

    function Request-CacheBuild {
        param([string]$Reason)

        $CanDelta          = [bool]($Sync.HasCache -and $Sync.CanAudit -and $Sync.CacheHasRooms)
        $D                 = New-Object System.Windows.Forms.Form
        $D.Text            = 'Cache aktualisieren'
        $D.Size            = New-Object System.Drawing.Size(560, 330)
        $D.StartPosition   = 'CenterParent'
        $D.FormBorderStyle = 'FixedDialog'
        $D.MaximizeBox     = $false
        $D.MinimizeBox     = $false
        $D.Font            = $Font
        if ($Form.Icon) { $D.Icon = $Form.Icon }

        $L          = New-Object System.Windows.Forms.Label
        $L.Location = New-Object System.Drawing.Point(16, 16)
        $L.Size     = New-Object System.Drawing.Size(515, 150)
        $DeltaInfo  = if ($CanDelta) {
                          "Delta-Update: liest nur Postfächer neu ein, deren Rechte laut Audit-Log geändert wurden (meist unter 1 min)."
                      } elseif ($Sync.HasCache -and -not $Sync.CanAudit) {
                          "Delta-Update nicht möglich: Rolle 'View-Only Audit Logs' fehlt."
                      } else { '' }
        $L.Text     = $(if ($Reason) { "$Reason`n`n" } else { '' }) +
                      "Komplett neu aufbauen: liest alle Postfachberechtigungen neu ein.`n" +
                      "Dauer: einige Minuten (mit Kalenderrechten etwa doppelt so lange).`n" +
                      "Das Fenster bleibt währenddessen bedienbar.`n`n" + $DeltaInfo

        $C          = New-Object System.Windows.Forms.CheckBox
        $C.Text     = 'Kalenderrechte mit einlesen'
        $C.Font     = $FontBold
        $C.Location = New-Object System.Drawing.Point(16, 178)
        $C.AutoSize = $true
        $C.Checked  = $ChkCal.Checked

        $Ok              = New-Object System.Windows.Forms.Button
        $Ok.Text         = 'Komplett neu aufbauen'
        $Ok.Location     = New-Object System.Drawing.Point(226, 235)
        $Ok.Size         = New-Object System.Drawing.Size(200, 30)
        $Ok.DialogResult = [System.Windows.Forms.DialogResult]::OK

        $Dl              = New-Object System.Windows.Forms.Button
        $Dl.Text         = 'Delta-Update'
        $Dl.Location     = New-Object System.Drawing.Point(16, 235)
        $Dl.Size         = New-Object System.Drawing.Size(200, 30)
        $Dl.Enabled      = $CanDelta
        $Dl.DialogResult = [System.Windows.Forms.DialogResult]::Yes

        # Hervorgehoben: Delta, wenn möglich - sonst Komplettaufbau
        $Main            = if ($CanDelta) { $Dl } else { $Ok }
        $Main.BackColor  = [System.Drawing.Color]::FromArgb(0, 120, 215)
        $Main.ForeColor  = [System.Drawing.Color]::White
        $Main.FlatStyle  = 'Flat'
        $Main.Font       = $FontBold

        $No              = New-Object System.Windows.Forms.Button
        $No.Text         = 'Später'
        $No.Location     = New-Object System.Drawing.Point(436, 235)
        $No.Size         = New-Object System.Drawing.Size(90, 30)
        $No.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

        $D.Controls.AddRange(@($L, $C, $Dl, $Ok, $No))
        $D.AcceptButton = $Main
        $D.CancelButton = $No

        $Result = $D.ShowDialog($Form)
        $Cal    = [bool]$C.Checked
        $D.Dispose()
        if ($Result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'DeltaUpdate'; User = [string]$Sync.ResultUser })
            return $true
        }
        if ($Result -ne [System.Windows.Forms.DialogResult]::OK) { return $false }

        $ChkCal.Checked = $Cal
        $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'BuildCache'; Calendar = $Cal })
        return $true
    }

    function Invoke-DeltaUpdate {
        if (Test-Busy) { return }
        if (-not $Sync.HasCache) { [void](Request-CacheBuild); return }
        if (-not $Sync.CanAudit) {
            [void][System.Windows.Forms.MessageBox]::Show($Form, "Delta-Update nicht möglich: Search-UnifiedAuditLog ist nicht verfügbar.`n`nBenötigt: Rolle 'View-Only Audit Logs' bzw. 'Audit Logs'.", 'Delta-Update', 'OK', 'Information')
            return
        }
        $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'DeltaUpdate'; User = [string]$Sync.ResultUser })
        $LblStatus.Text = 'Delta-Update gestartet...'
    }

    function Start-UiUpdateCheck {
        param([bool]$Manual = $false)
        $P = [powershell]::Create()
        [void]$P.AddScript($Sync.UpdateScript).AddArgument($Sync).AddArgument($Manual)
        [void]$P.BeginInvoke()
    }

    function Invoke-LiveCheck {
        if (Test-Busy) { return }
        $User = if ($Sync.ResultUser) { [string]$Sync.ResultUser } else { (Get-InputAddress $TxtUser.Text) }
        if (-not $User) {
            [void][System.Windows.Forms.MessageBox]::Show($Form, 'Bitte zuerst einen Benutzer eingeben bzw. abfragen.', 'Live prüfen', 'OK', 'Information')
            return
        }
        if (-not $Sync.HasCache) { [void](Request-CacheBuild); return }

        $Addrs = @($State.Raw | Where-Object { $_ -and $_.Adresse } | ForEach-Object { [string]$_.Adresse } | Select-Object -Unique)
        $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'LiveCheck'; User = $User; Mailboxes = $Addrs })
        $LblStatus.Text = 'Live-Prüfung gestartet...'
    }

    # Baut die Tabelle aus den Rohdaten neu auf (Gruppen-Filter + Zusammenfassung)
    function Update-Table {
        $Rows = @($State.Raw | Where-Object { $_ })
        if (-not $ChkGroups.Checked) {
            $Rows = @($Rows | Where-Object { -not (([string]$_.Quelle) -like 'Gruppe*') })
        }
        if (-not $ChkRooms.Checked) {
            $Rows = @($Rows | Where-Object { ([string]$_.Typ) -notin 'Room', 'Equipment' })
        }

        $Order = @{ FullAccess = 1; SendAs = 2; SendOnBehalf = 3; Kalender = 4; Buchungsstellvertreter = 5 }

        $Table.BeginLoadData()
        $Table.Rows.Clear()

        if ($ChkMerge.Checked) {
            # Ein Eintrag pro Postfach
            $Groups = $Rows | Group-Object { if ($_.Adresse) { ([string]$_.Adresse).ToLower() } else { [string]$_.Postfach } }
            foreach ($G in ($Groups | Sort-Object { [string]$_.Group[0].Postfach })) {
                $First = $G.Group[0]

                $Perms = @($G.Group |
                    Sort-Object { $Order[[string]$_.Berechtigung] } |
                    ForEach-Object {
                        if ($_.Berechtigung -eq 'Kalender') { "Kalender ($($_.Rechte))" } else { [string]$_.Berechtigung }
                    } | Select-Object -Unique) -join ', '

                $Sources = @($G.Group |
                    ForEach-Object { [string]$_.Quelle } |
                    Select-Object -Unique |
                    Sort-Object { -not ($_ -like 'direkt*') }, { $_ }) -join '; '

                [void]$Table.Rows.Add($First.Postfach, $Perms, $First.Adresse, $First.Typ, '', $Sources)
            }
        } else {
            # Detailansicht: eine Zeile pro Berechtigung
            foreach ($R in $Rows) {
                [void]$Table.Rows.Add($R.Postfach, $R.Berechtigung, $R.Adresse, $R.Typ, $R.Rechte, $R.Quelle)
            }
        }

        $Table.EndLoadData()

        if ($Grid.Columns.Count -ge 6) {
            $Grid.Columns['Rechte'].Visible          = -not $ChkMerge.Checked
            $Grid.Columns['Berechtigung'].HeaderText = if ($ChkMerge.Checked) { 'Berechtigungen' } else { 'Berechtigung' }
            $Grid.Columns['Berechtigung'].FillWeight = if ($ChkMerge.Checked) { 22 } else { 11 }
        }

        $State.Shown = $Rows
        Update-Filter
    }

    function Update-Filter {
        $Parts = @()
        $Txt = $TxtFilter.Text.Trim()
        if ($Txt) {
            $Esc = ($Txt -replace '([\[\]\*%])', '[$1]') -replace "'", "''"
            $Parts += "(Berechtigung LIKE '%$Esc%' OR Postfach LIKE '%$Esc%' OR Adresse LIKE '%$Esc%' OR Typ LIKE '%$Esc%' OR Rechte LIKE '%$Esc%' OR Quelle LIKE '%$Esc%')"
        }
        if ($CmbType.SelectedItem -and $CmbType.SelectedItem -ne 'Alle') {
            # LIKE statt = -> funktioniert auch in der zusammengefassten Ansicht
            $Parts += "Berechtigung LIKE '%$($CmbType.SelectedItem)%'"
        }
        $Table.DefaultView.RowFilter = $Parts -join ' AND '

        # Zähler aus den Rohdaten (nach Gruppen-Filter)
        $Raw         = @($State.Shown | Where-Object { $_ })
        $Counts      = $Raw | Group-Object Berechtigung | ForEach-Object { "$($_.Name): $($_.Count)" }
        $GroupCount  = @($Raw | Where-Object { ([string]$_.Quelle) -like 'Gruppe*' }).Count
        $DirectCount = $Raw.Count - $GroupCount
        $Unit        = if ($ChkMerge.Checked) { 'Postfächer' } else { 'Einträge' }

        $LblCount.Text = "Angezeigt: $($Table.DefaultView.Count) / $($Table.Rows.Count) $Unit   " +
                         "(Rechte direkt: $DirectCount, über Gruppen: $GroupCount)   " + ($Counts -join '  |  ')
    }

    # --- Timer: Status vom Arbeits-Thread übernehmen ---
    $Timer          = New-Object System.Windows.Forms.Timer
    $Timer.Interval = 150
    $Timer.add_Tick({
        $Busy = Test-Busy

        if ($LblStatus.Text -ne $Sync.StatusText) { $LblStatus.Text = $Sync.StatusText }

        if ($Busy -and $Sync.Percent -lt 0) {
            if ($Progress.Style -ne 'Marquee') { $Progress.Style = 'Marquee' }
        } else {
            if ($Progress.Style -ne 'Continuous') { $Progress.Style = 'Continuous' }
            $Progress.Value = if ($Busy) { [Math]::Min(100, [Math]::Max(0, [int]$Sync.Percent)) } else { 0 }
        }

        $BtnQuery.Enabled  = -not $Busy
        $BtnCache.Enabled  = -not $Busy
        $BtnLive.Enabled   = (-not $Busy) -and $Sync.HasCache -and ([bool]$Sync.ResultUser -or [bool](Get-InputAddress $TxtUser.Text))
        $ChkCal.Enabled    = -not $Busy
        $BtnRemove.Enabled = (-not $Busy) -and $Table.Rows.Count -gt 0
        $BtnGrant.Enabled  = (-not $Busy) -and $Sync.HasCache -and ([bool](Get-InputAddress $TxtUser.Text) -or [bool]$Sync.ResultUser)
        $BtnCancel.Enabled = $Busy
        $HasRows = @($State.Raw | Where-Object { $_ -and $_.Quelle -eq 'direkt' }).Count -gt 0
        $BtnMbView.Enabled     = (-not $Busy) -and $Sync.HasCache
        $BtnCopyRights.Enabled = (-not $Busy) -and $Sync.CanWrite -and $HasRows
        $BtnOffboard.Enabled   = (-not $Busy) -and $Sync.CanWrite -and $HasRows
        if (-not $Sync.CanWrite) { $BtnGrant.Enabled = $false; $BtnRemove.Enabled = $false }

        if ($Sync.CacheVersion -ne $State.CacheVersion) {
            $State.CacheVersion = $Sync.CacheVersion
            Update-CacheLabel
            Update-AutoComplete
        }

        if ($Sync.ResultVersion -ne $State.ResultVersion) {
            $State.ResultVersion = $Sync.ResultVersion
            $State.Raw = @($Sync.Results)
            if ($Sync.ResultUser -and (Get-InputAddress $TxtUser.Text) -ne [string]$Sync.ResultUser) { $TxtUser.Text = [string]$Sync.ResultUser }
            Add-RecentUser ([string]$Sync.ResultUser)
            $Form.Text = "Exchange Online - Postfachberechtigungen v$($Sync.AppVersion) - $($Sync.ResultUser)   [$($Sync.ConnUpn)]"
            Update-Table
        }

        if ($Sync.ErrorVersion -ne $State.ErrorVersion) {
            $State.ErrorVersion = $Sync.ErrorVersion
            [void][System.Windows.Forms.MessageBox]::Show($Form, $Sync.ErrorText, 'Fehler', 'OK', 'Error')
        }

        if ($Sync.ExtraVersion -ne $State.ExtraVersion) {
            $State.ExtraVersion = $Sync.ExtraVersion
            $X = $Sync.Extra
            if ($X) { Show-TableWindow -Title $X.Title -Info $X.Info -Columns $X.Columns -Rows $X.Rows }
        }

        if ($Sync.GroupVersion -ne $State.GroupVersion) {
            $State.GroupVersion = $Sync.GroupVersion
            $GD = $Sync.GroupData
            if ($GD) { Invoke-GroupDialog -Data $GD }
        }

        # Nach Erteilen/Entfernen: Postfach-Fenster sofort, Hauptliste nach 45 s erneut live prüfen
        if ($Sync.RecheckVersion -ne $State.RecheckVersion) {
            $State.RecheckVersion = $Sync.RecheckVersion
            $State.RecheckAt      = (Get-Date).AddSeconds(45)
            $State.RecheckUser    = [string]$Sync.RecheckUser
            if ($State.Mb -and $Sync.MbResult -and (@($Sync.RecheckAddrs) -contains [string]$Sync.MbResult.Address)) {
                $State.Mb.Txt.Text = [string]$Sync.MbResult.Address
                $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'MailboxQuery'; Mailbox = [string]$Sync.MbResult.Address })
            }
        }
        if ($State.RecheckAt -and (Get-Date) -ge $State.RecheckAt -and -not $Busy) {
            $State.RecheckAt = $null
            if ($State.RecheckUser -and [string]$Sync.ResultUser -eq $State.RecheckUser) {
                $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'Query'; User = $State.RecheckUser })
            }
        }

        if ($Sync.UpdateVersion -ne $State.UpdateVersion) {
            $State.UpdateVersion = $Sync.UpdateVersion
            $U = $Sync.Update
            if ($U -and $U.Ok -and $U.Newer) {
                $State.UpdateUrl   = [string]$U.Url
                $LnkUpdate.Text    = "Update verfügbar: v$($U.Latest)"
                $LnkUpdate.Visible = $true
                if ($U.Manual) {
                    $Ans = [System.Windows.Forms.MessageBox]::Show($Form, "Neue Version verfügbar: v$($U.Latest)`nInstalliert: v$($Sync.AppVersion)`n`nRelease-Seite jetzt öffnen?", 'Update', 'YesNo', 'Information')
                    if ($Ans -eq 'Yes') { Start-Process ([string]$U.Url) }
                }
            } elseif ($U -and $U.Manual) {
                if ($U.Ok) {
                    [void][System.Windows.Forms.MessageBox]::Show($Form, "Du verwendest die aktuelle Version (v$($Sync.AppVersion)).", 'Update', 'OK', 'Information')
                } else {
                    [void][System.Windows.Forms.MessageBox]::Show($Form, "Update-Prüfung fehlgeschlagen:`n`n$($U.Error)", 'Update', 'OK', 'Warning')
                }
            }
        }

        if ($Sync.ReportVersion -ne $State.ReportVersion) {
            $State.ReportVersion = $Sync.ReportVersion
            $Icon = if ($Sync.ReportFailed) { 'Warning' } else { 'Information' }
            [void][System.Windows.Forms.MessageBox]::Show($Form, $Sync.ReportText, $(if ($Sync.ReportTitle) { [string]$Sync.ReportTitle } else { 'Ergebnis' }), 'OK', $Icon)
        }
    })

    # --- Events ---
    $BtnQuery.add_Click({
        $U = (Get-InputAddress $TxtUser.Text)
        if (-not $U) {
            [void][System.Windows.Forms.MessageBox]::Show($Form, 'Bitte einen Benutzer (UPN/E-Mail) eingeben.', 'Hinweis', 'OK', 'Information')
            return
        }
        if (-not $Sync.HasCache) {
            if (-not (Request-CacheBuild)) { return }
        }
        $State.Raw = @()
        Update-Table
        $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'Query'; User = $U })
    })

    $BtnCache.add_Click({ [void](Request-CacheBuild) })
    $BtnLive.add_Click({ Invoke-LiveCheck })

    $BtnCancel.add_Click({
        $Sync.Cancel    = $true
        $LblStatus.Text = 'Breche ab...'
    })

    $TxtFilter.add_TextChanged({ Update-Filter })
    $CmbType.add_SelectedIndexChanged({ Update-Filter })
    $ChkGroups.add_CheckedChanged({ Update-Table })
    $ChkRooms.add_CheckedChanged({ Update-Table })
    $ChkMerge.add_CheckedChanged({ Update-Table })

    $BtnRemove.add_Click({ Invoke-RemoveDialog })
    $BtnGrant.add_Click({ Invoke-GrantDialog })
    $BtnMbView.add_Click({
        $A = ''
        if ($Grid.SelectedRows.Count -eq 1) { $A = [string]$Grid.SelectedRows[0].Cells['Adresse'].Value }
        Show-MailboxWindow -Address $A
    })
    $MiMbView.add_Click({
        if ($Grid.SelectedRows.Count -ge 1) { Show-MailboxWindow -Address ([string]$Grid.SelectedRows[0].Cells['Adresse'].Value) }
    })
    $BtnCopyRights.add_Click({ Invoke-CopyDialog })
    $MiFolder.add_Click({
        if ($Grid.SelectedRows.Count -ge 1) { Invoke-FolderCheck -Address ([string]$Grid.SelectedRows[0].Cells['Adresse'].Value) }
    })
    $BtnOffboard.add_Click({ Invoke-OffboardDialog })
    $LnkLog.add_LinkClicked({ Show-LogWindow })
    $BtnRecent.add_Click({
        $CtxRecent.Items.Clear()
        if (@($State.Recent).Count -eq 0) {
            $It = $CtxRecent.Items.Add('(noch keine)'); $It.Enabled = $false
        } else {
            foreach ($U in $State.Recent) {
                $It = $CtxRecent.Items.Add([string]$U)
                $It.add_Click({
                    param($s, $e)
                    $TxtUser.Text = [string]$s.Text
                    if (-not (Test-Busy) -and $Sync.HasCache) {
                        $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'Query'; User = [string]$s.Text })
                    }
                })
            }
        }
        $PtMenu = New-Object System.Drawing.Point -ArgumentList 0, ([int]$BtnRecent.Height)
        $CtxRecent.Show($BtnRecent, $PtMenu)
    })
    $MiGrant.add_Click({ Invoke-GrantDialog })
    $MiRemove.add_Click({ Invoke-RemoveDialog })
    $MiCopy.add_Click({
        $Addrs = @($Grid.SelectedRows | Sort-Object Index | ForEach-Object { [string]$_.Cells['Adresse'].Value } | Where-Object { $_ })
        if ($Addrs.Count) {
            [System.Windows.Forms.Clipboard]::SetText(($Addrs -join "`r`n"))
            $LblStatus.Text = "$($Addrs.Count) Adresse(n) kopiert"
        }
    })
    $Ctx.add_Opening({
        $MiRemove.Enabled = (-not (Test-Busy)) -and $Grid.SelectedRows.Count -gt 0
        $MiGrant.Enabled  = (-not (Test-Busy)) -and $Sync.HasCache -and $Sync.CanWrite
        $MiRemove.Enabled = $MiRemove.Enabled -and $Sync.CanWrite
        $MiMbView.Enabled = $Grid.SelectedRows.Count -ge 1 -and $Sync.HasCache
    })

    $BtnExport.add_Click({
        if ($Table.DefaultView.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show($Form, 'Keine Daten zum Exportieren.', 'Hinweis', 'OK', 'Information')
            return
        }
        $Dlg          = New-Object System.Windows.Forms.SaveFileDialog
        $Dlg.Filter   = 'CSV-Datei (*.csv)|*.csv'
        $Name         = ((Get-InputAddress $TxtUser.Text) -split '@')[0]
        $Dlg.FileName = "Postfachrechte_${Name}_$(Get-Date -Format yyyyMMdd_HHmm).csv"
        if ($Dlg.ShowDialog($Form) -eq 'OK') {
            $Export = $Table.DefaultView.ToTable()
            if ($ChkMerge.Checked) {
                $Export | Select-Object Postfach, @{N = 'Berechtigungen'; E = { $_.Berechtigung } }, Adresse, Typ, Quelle |
                    Export-Csv -Path $Dlg.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ';'
            } else {
                $Export | Select-Object Postfach, Berechtigung, Adresse, Typ, Rechte, Quelle |
                    Export-Csv -Path $Dlg.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ';'
            }
            $LblStatus.Text = "Exportiert: $($Dlg.FileName)"
        }
    })

    $BtnCopy.add_Click({
        if ($Grid.SelectedRows.Count -eq 0) { return }
        $Lines = foreach ($Row in ($Grid.SelectedRows | Sort-Object Index)) {
            ($Row.Cells | Where-Object { $_.Visible } | ForEach-Object { [string]$_.Value }) -join "`t"
        }
        $Header = ($Grid.Columns | Where-Object { $_.Visible } | ForEach-Object { $_.HeaderText }) -join "`t"
        [System.Windows.Forms.Clipboard]::SetText(($Header + "`r`n" + ($Lines -join "`r`n")))
        $LblStatus.Text = "$($Grid.SelectedRows.Count) Zeile(n) kopiert (Excel-tauglich)"
    })

    $Form.add_Shown({
        $Form.Activate()
        Update-CacheLabel
        $Timer.Start()

        if ($Sync.RebuildCache -or -not $Sync.HasCache) {
            [void](Request-CacheBuild)
        } elseif ($Sync.CacheStale -and -not $Sync.CacheHasRooms) {
            [void](Request-CacheBuild -Reason "Der Cache enthält noch keine Raumpostfach-Rechte (Kalender, Buchungsstellvertreter).`nEinmal neu aufbauen, damit Raumpostfächer angezeigt werden.")
        } elseif ($Sync.CacheStale) {
            [void](Request-CacheBuild -Reason "Der Cache ist älter als $($Sync.MaxCacheAgeHours) h - neu vergebene Rechte fehlen evtl.")
        }
        if ($TxtUser.Text -and ($Sync.HasCache -or $Sync.Queue.Count -gt 0)) {
            $Sync.Queue.Enqueue([PSCustomObject]@{ Type = 'Query'; User = (Get-InputAddress $TxtUser.Text) })
        } else {
            [void]$TxtUser.Focus()
        }
    })

    $Form.add_FormClosing({
        param($s, $e)
        if (Test-Busy) {
            $Ans = [System.Windows.Forms.MessageBox]::Show($Form, "Es läuft noch ein Vorgang.`nTrotzdem beenden?", 'Beenden', 'YesNo', 'Warning')
            if ($Ans -ne 'Yes') { $e.Cancel = $true; return }
            $Sync.Cancel = $true
        }
        $Timer.Stop()
        $Sync.Closed = $true
    })

    [void]$Form.ShowDialog()
    $Form.Dispose()
    $Sync.Closed = $true
}

#endregion

#region ===== GUI: Start + Arbeitsschleife =====

# App-Icon suchen (neben dem Skript - im Repo bzw. von der EXE nach %LOCALAPPDATA% geschrieben)
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
foreach ($Candidate in @((Join-Path $ScriptDir 'mailbox.ico'), (Join-Path (Split-Path $CachePath -Parent) 'mailbox.ico'))) {
    if ($Candidate -and (Test-Path $Candidate)) { $Sync.IconPath = $Candidate; break }
}

# Eigene AppUserModelID -> Windows gruppiert das Fenster nicht unter powershell.exe
# und zeigt in der Taskleiste das Fenster-Icon statt des PowerShell-Logos
try {
    if (-not ('Win32.AppId' -as [type])) {
        Add-Type -Namespace Win32 -Name AppId -MemberDefinition @'
[DllImport("shell32.dll", SetLastError = true)]
public static extern int SetCurrentProcessExplicitAppUserModelID([MarshalAs(UnmanagedType.LPWStr)] string AppID);
'@
    }
    [void][Win32.AppId]::SetCurrentProcessExplicitAppUserModelID('Example.MailboxAccess')
} catch { }

$UiRunspace                = [runspacefactory]::CreateRunspace()
$UiRunspace.ApartmentState = 'STA'
$UiRunspace.ThreadOptions  = 'ReuseThread'
$UiRunspace.Open()
$UiRunspace.SessionStateProxy.SetVariable('Sync', $Sync)

$UiPs          = [powershell]::Create()
$UiPs.Runspace = $UiRunspace
[void]$UiPs.AddScript($UiScript.ToString())
$UiHandle = $UiPs.BeginInvoke()

# Update-Prüfung im Hintergrund (abschaltbar mit -NoUpdateCheck)
if (-not $NoUpdateCheck) {
    try { Start-UpdateCheck -Manual $false } catch { Write-DiagLog 'ui_fehler.log' "Update-Prüfung: $($_.Exception.Message)" }
}

# Arbeitsschleife: führt Befehle der GUI aus (hier liegt die EXO-Verbindung)
while (-not $Sync.Closed) {
    $Cmd = $null
    if (-not $Sync.Queue.TryDequeue([ref]$Cmd)) {
        # GUI-Thread unerwartet beendet?
        if ($UiHandle.IsCompleted) { break }
        Start-Sleep -Milliseconds 100
        continue
    }

    $Sync.Busy   = $true
    $Sync.Cancel = $false
    try {
        switch ($Cmd.Type) {
            'BuildCache' {
                Build-AccessCache -IncludeCalendar:$Cmd.Calendar -OnProgress $SyncProgress
                $FailCount = @($script:Cache.Failed | Where-Object { $_ }).Count
                $FailText  = if ($FailCount -gt 0) { " - $FailCount Postfächer fehlgeschlagen" } else { '' }
                $Sync.StatusText = "Cache aufgebaut in $($script:Cache.Duration) min$FailText"
            }
            'Query' {
                $Sw  = [System.Diagnostics.Stopwatch]::StartNew()
                $Res = @(Get-MailboxAccess -User $Cmd.User -OnProgress $SyncProgress)
                $Sync.Results    = $Res
                $Sync.ResultUser = $Cmd.User
                $Sync.ResultVersion++
                $Sync.StatusText = "Fertig in $([Math]::Round($Sw.Elapsed.TotalSeconds, 1)) s: $($Res.Count) Berechtigungen für $($Cmd.User)  (Doppelklick = wer hat Zugriff?)"
            }
            'DeltaUpdate' {
                $D = Update-CacheDelta -OnProgress $SyncProgress
                $Sync.MailboxList = Get-MailboxPickList
                if ($Cmd.User) {
                    $Res = @(Get-MailboxAccess -User $Cmd.User -OnProgress $SyncProgress)
                    $Sync.Results    = $Res
                    $Sync.ResultUser = $Cmd.User
                    $Sync.ResultVersion++
                }
                $Sync.ReportTitle  = 'Delta-Update'
                $Sync.ReportText   = Format-DeltaReport $D
                $Sync.ReportFailed = [bool]($D.Failed -gt 0 -or $D.Truncated)
                $Sync.ReportVersion++
                $Sync.StatusText = "Delta-Update: $($D.Changed) Postfächer nachgelesen in $($D.Duration) min ($($D.Records) Audit-Einträge)"
            }
            'LiveCheck' {
                $Live = Update-CacheLive -Addresses $Cmd.Mailboxes -IncludeSharedRooms -OnProgress $SyncProgress
                $Res  = @(Get-MailboxAccess -User $Cmd.User -OnProgress $SyncProgress)
                $Sync.Results    = $Res
                $Sync.ResultUser = $Cmd.User
                $Sync.ResultVersion++
                $FailText = if ($Live.Failed) { ", $($Live.Failed) Fehler" } else { '' }
                $Sync.StatusText = "Live geprüft: $($Live.Checked) Postfächer in $($Live.Duration) min$FailText - $($Res.Count) Berechtigungen für $($Cmd.User)"
            }
            'ForwardQuery' {
                $FI = Get-ForwardingInfo -User $Cmd.User -OnProgress $SyncProgress
                $Sync.Extra = [PSCustomObject]@{
                    Title     = "Weiterleitungen und Regeln: $($Cmd.User)"
                    Info      = $FI.Info
                    Columns   = @('Art', 'Name', 'Aktiv', 'Ziel', 'Extern', 'Details')
                    Rows      = $FI.Rows
                }
                $Sync.ExtraVersion++
                $Sync.StatusText = "Weiterleitungen/Regeln geladen: $($Cmd.User)"
            }
            'FolderQuery' {
                $FR = Get-MailboxFolderRights -Mailbox $Cmd.Mailbox -OnProgress $SyncProgress
                $Sync.Extra = [PSCustomObject]@{
                    Title     = "Ordnerrechte: $($Cmd.Mailbox)"
                    Info      = $FR.Info
                    Columns   = @('Ordner', 'Benutzer', 'Rechte', 'Hinweis')
                    Rows      = $FR.Rows
                }
                $Sync.ExtraVersion++
                $Sync.StatusText = "Ordnerrechte geladen: $($Cmd.Mailbox)"
            }
            'GroupQuery' {
                $Sync.GroupData = Get-GroupInfo -User $Cmd.User -OnProgress $SyncProgress
                $Sync.GroupVersion++
                $Sync.StatusText = "Gruppen geladen: $($Cmd.User)"
            }
            'GroupChange' {
                $Rep       = @(Set-GroupMembership -User $Cmd.User -Add $Cmd.Add -Remove $Cmd.Remove -Admin $Sync.ConnUpn -OnProgress $SyncProgress)
                $OkCount   = @($Rep | Where-Object { $_.Ok }).Count
                $FailCount = $Rep.Count - $OkCount
                $Lines = foreach ($E in $Rep) {
                    $Act = if ($E.Aktion -eq 'Add') { 'hinzugefügt' } else { 'entfernt' }
                    if ($E.Ok) { "OK       $($E.Gruppe) ($($E.Typ)) - $Act" }
                    else       { "FEHLER   $($E.Gruppe) ($($E.Typ))`n         $($E.Meldung)" }
                }
                $Sync.ReportText   = "Gruppenmitgliedschaften: $($Cmd.User)`nErfolgreich: $OkCount   Fehler: $FailCount`n`n" + ($Lines -join "`n") +
                                     "`n`nHinweis: Rechte über Gruppen werden erst nach der Replikation sichtbar (einige Minuten)."
                $Sync.ReportFailed = $FailCount -gt 0
                $Sync.ReportTitle  = 'Ergebnis: Gruppen'
                $Sync.ReportVersion++
                $Res = @(Get-MailboxAccess -User $Cmd.User -OnProgress $SyncProgress)
                $Sync.Results    = $Res
                $Sync.ResultUser = $Cmd.User
                $Sync.ResultVersion++
                $Sync.StatusText = "Gruppen geändert: $OkCount OK, $FailCount Fehler"
            }
            'TenantReport' {
                $TR = Export-TenantReport -Folder $Cmd.Folder -OnProgress $SyncProgress
                $Sync.ReportText   = "Tenant-Gesamtreport erstellt:`n`n$($TR.Count) Berechtigungen`n$($TR.FwdCount) Weiterleitungen (davon extern: $($TR.ExtCount))`n`n" +
                                     "HTML: $($TR.Html)`nCSV:  $($TR.Csv)`nCSV:  $($TR.FwdCsv)"
                $Sync.ReportFailed = $false
                $Sync.ReportTitle  = 'Tenant-Gesamtreport'
                $Sync.ReportVersion++
                try { Start-Process $TR.Html } catch { }
                $Sync.StatusText = "Report erstellt: $($TR.Html)"
            }
            'MailboxQuery' {
                $Sw  = [System.Diagnostics.Stopwatch]::StartNew()
                $Res = Get-MailboxAccessors -Mailbox $Cmd.Mailbox -OnProgress $SyncProgress
                $Sync.MbResult = $Res
                $Sync.MbVersion++
                $Sync.StatusText = "Postfach $($Res.Name): $(@($Res.Entries).Count) Berechtigungen in $([Math]::Round($Sw.Elapsed.TotalSeconds, 1)) s"
            }
            'Grant' {
                $Rep       = @(Grant-UserMailboxAccess -User $Cmd.User -Items $Cmd.Items -Options $Cmd.Options -Admin $Sync.ConnUpn -OnProgress $SyncProgress)
                $OkCount   = @($Rep | Where-Object { $_.Ok }).Count
                $FailCount = $Rep.Count - $OkCount

                $EacName = @{ FullAccess = 'Lesen und Verwalten (Vollzugriff)'; SendAs = 'Senden als'; SendOnBehalf = 'Senden im Auftrag von'; Kalender = 'Kalenderberechtigung'; Buchungsstellvertreter = 'Buchungsstellvertreter' }
                $Lines = foreach ($E in $Rep) {
                    $PName = if ($EacName[$E.Berechtigung]) { $EacName[$E.Berechtigung] } else { $E.Berechtigung }
                    if ($E.Ok) { "OK       $($E.Postfach) - $PName  ($($E.Meldung))" }
                    else       { "FEHLER   $($E.Postfach) - $PName`n         $($E.Meldung)" }
                }
                $Sync.ReportText = "Rechte erteilt für: $($Cmd.User)`nErfolgreich: $OkCount   Fehler: $FailCount`n`n" +
                                   ($Lines -join "`n") +
                                   "`n`nHinweis: Exchange Online braucht teils bis zu 60 Minuten, bis Rechte (v. a. AutoMapping in Outlook) wirksam sind." +
                                   "`nProtokoll: %LOCALAPPDATA%\EXOMailboxAccess\aenderungen.log"
                $Sync.ReportFailed = $FailCount -gt 0
                $Sync.ReportTitle  = if ($Cmd.PSObject.Properties['Title'] -and $Cmd.Title) { [string]$Cmd.Title } else { 'Ergebnis: Rechte erteilen' }
                $Sync.ReportVersion++

                $Res = @(Get-MailboxAccess -User $Cmd.User -OnProgress $SyncProgress)
                $Opt = if ($Cmd.PSObject.Properties['Options'] -and $Cmd.Options) { $Cmd.Options } else { @{} }
                $Res = @(Merge-ChangeIntoResults -Results $Res -Report $Rep -Mode 'Grant' -Items $Cmd.Items -Options $Opt)
                $Sync.Results    = $Res
                $Sync.ResultUser = $Cmd.User
                $Sync.ResultVersion++
                $Sync.RecheckUser  = $Cmd.User
                $Sync.RecheckAddrs = @($Cmd.Items | ForEach-Object { [string]$_.Adresse })
                $Sync.RecheckVersion++
                $Sync.StatusText = "Erteilt: $OkCount, Fehler: $FailCount - Liste aktualisiert (Nachprüfung in 45 s)"
            }
            'Remove' {
                $Rep       = @(Remove-UserMailboxAccess -User $Cmd.User -Items $Cmd.Items -Admin $Sync.ConnUpn -OnProgress $SyncProgress)
                $OkCount   = @($Rep | Where-Object { $_.Ok }).Count
                $FailCount = $Rep.Count - $OkCount

                $EacName = @{ FullAccess = 'Lesen und Verwalten (Vollzugriff)'; SendAs = 'Senden als'; SendOnBehalf = 'Senden im Auftrag von'; Kalender = 'Kalenderberechtigung'; Buchungsstellvertreter = 'Buchungsstellvertreter' }
                $Lines = foreach ($E in $Rep) {
                    $PName = if ($EacName[$E.Berechtigung]) { $EacName[$E.Berechtigung] } else { $E.Berechtigung }
                    if ($E.Ok) { "OK       $($E.Postfach) - $PName" }
                    else       { "FEHLER   $($E.Postfach) - $PName`n         $($E.Meldung)" }
                }
                $Sync.ReportText = "Benutzer: $($Cmd.User)`nErfolgreich: $OkCount   Fehler: $FailCount`n`n" +
                                   ($Lines -join "`n") +
                                   "`n`nHinweis: Exchange Online braucht teils einige Minuten, bis Änderungen überall wirksam sind." +
                                   "`nProtokoll: %LOCALAPPDATA%\EXOMailboxAccess\aenderungen.log"
                $Sync.ReportFailed = $FailCount -gt 0
                $Sync.ReportTitle  = if ($Cmd.PSObject.Properties['Title'] -and $Cmd.Title) { [string]$Cmd.Title } else { 'Ergebnis: Rechte entfernen' }
                $Sync.ReportVersion++

                # Liste neu laden
                $Res = @(Get-MailboxAccess -User $Cmd.User -OnProgress $SyncProgress)
                $Res = @(Merge-ChangeIntoResults -Results $Res -Report $Rep -Mode 'Remove' -Items $Cmd.Items)
                $Sync.Results    = $Res
                $Sync.ResultUser = $Cmd.User
                $Sync.ResultVersion++
                $Sync.RecheckUser  = $Cmd.User
                $Sync.RecheckAddrs = @($Cmd.Items | ForEach-Object { [string]$_.Adresse })
                $Sync.RecheckVersion++
                $Sync.StatusText = "Entfernt: $OkCount, Fehler: $FailCount - Liste aktualisiert (Nachprüfung in 45 s)"
            }
        }
    } catch {
        if ($Cmd -and $Cmd.Type -eq 'MailboxQuery') { $Sync.MbResult = $null; $Sync.MbVersion++ }
        $Ex = $_.Exception
        while ($Ex.InnerException -and -not ($Ex -is [System.OperationCanceledException])) { $Ex = $Ex.InnerException }
        if ($Ex -is [System.OperationCanceledException] -or $_.Exception.Message -match 'Vom Benutzer abgebrochen') {
            $Sync.StatusText = 'Abgebrochen - bisheriger Cache bleibt erhalten'
        } else {
            $Sync.StatusText = 'Fehler - siehe Meldung'
            $Sync.ErrorText  = "Fehler bei '$($Cmd.Type)':`n`n$($_.Exception.Message)"
            $Sync.ErrorVersion++
        }
    } finally {
        $Sync.Percent     = 0
        $Sync.CacheInfo   = Get-CacheInfo
        $Sync.CacheStale  = Test-CacheStale
        $Sync.CacheHasRooms = Test-CacheHasRooms
        $Sync.HasCache    = [bool]$script:Cache
        $Sync.HasCalendar = [bool]($script:Cache -and $script:Cache.HasCalendar)
        if ($Cmd.Type -eq 'BuildCache') { $Sync.MailboxList = Get-MailboxPickList }
        $Sync.CacheVersion++
        $Sync.Cancel      = $false
        $Sync.Busy        = $false
    }
}

try {
    [void]$UiPs.EndInvoke($UiHandle)
} catch {
    Write-DiagLog 'ui_fehler.log' "EndInvoke: $($_.Exception.Message)"
}
$UiErrors = @($UiPs.Streams.Error | ForEach-Object { "$($_.Exception.Message)`r`n$($_.InvocationInfo.PositionMessage)" })
if ($UiErrors.Count) { Write-DiagLog 'ui_fehler.log' ("GUI-Fehler:`r`n" + ($UiErrors -join "`r`n---`r`n")) }
if (-not $Sync.Closed) {
    $First = if ($UiErrors.Count) { $UiErrors[0] } else { 'keine Details' }
    Show-DiagError "Die Oberfläche wurde unerwartet beendet.`n`n$First`n`nLog: $script:DiagDir\ui_fehler.log"
}
$UiPs.Dispose()
$UiRunspace.Close()
$UiRunspace.Dispose()

#endregion
