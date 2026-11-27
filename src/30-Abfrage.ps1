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

