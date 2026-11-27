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
