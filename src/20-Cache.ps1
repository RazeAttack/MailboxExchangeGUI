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

