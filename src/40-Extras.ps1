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

