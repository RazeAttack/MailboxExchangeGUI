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

