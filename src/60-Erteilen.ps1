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

