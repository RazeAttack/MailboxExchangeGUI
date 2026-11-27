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

