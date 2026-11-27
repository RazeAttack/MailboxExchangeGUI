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

