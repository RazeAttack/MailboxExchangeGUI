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

