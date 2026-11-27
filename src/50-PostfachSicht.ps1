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

