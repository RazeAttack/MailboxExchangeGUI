#region ===== v4.0: Update-Prüfung (GitHub-Release) =====

# Läuft in einem eigenen Runspace, damit Start und GUI nicht warten müssen
$script:UpdateCheckScript = {
    param($Sync, [bool]$Manual)
    $Result = $null
    try {
        [System.Net.ServicePointManager]::SecurityProtocol =
            [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
        try { [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials } catch { }

        $Uri     = "https://api.github.com/repos/$($Sync.RepoSlug)/releases/latest"
        $Headers = @{ 'User-Agent' = 'MailboxAccess-UpdateCheck'; 'Accept' = 'application/vnd.github+json' }
        $Rel     = Invoke-RestMethod -Uri $Uri -Headers $Headers -TimeoutSec 8 -UseBasicParsing -ErrorAction Stop

        $Tag     = ([string]$Rel.tag_name) -replace '^[vV]', ''
        $Latest  = $null
        $Current = $null
        [void][version]::TryParse($Tag, [ref]$Latest)
        [void][version]::TryParse([string]$Sync.AppVersion, [ref]$Current)
        $Newer   = [bool]($Latest -and $Current -and ($Latest -gt $Current))
        $ExeUrl  = ''
        $Exe     = @($Rel.assets | Where-Object { [string]$_.name -like '*.exe' } | Select-Object -First 1)
        if ($Exe.Count -gt 0) { $ExeUrl = [string]$Exe[0].browser_download_url }

        $Result = [PSCustomObject]@{
            Ok = $true; Manual = $Manual; Latest = $Tag; Newer = $Newer
            Url = [string]$Rel.html_url; ExeUrl = $ExeUrl; Notes = [string]$Rel.body; Error = ''
        }
    } catch {
        $Msg = $_.Exception.Message
        if ($Msg -match '404') { $Msg = 'Auf GitHub ist noch kein Release veröffentlicht.' }
        $Result = [PSCustomObject]@{
            Ok = $false; Manual = $Manual; Latest = ''; Newer = $false
            Url = "https://github.com/$($Sync.RepoSlug)/releases"; ExeUrl = ''; Notes = ''; Error = $Msg
        }
    }
    $Sync.Update = $Result
    $Sync.UpdateVersion++
}

function Start-UpdateCheck {
    param([bool]$Manual = $false)
    $P = [powershell]::Create()
    [void]$P.AddScript($Sync.UpdateScript).AddArgument($Sync).AddArgument($Manual)
    $script:UpdatePs = $P
    [void]$P.BeginInvoke()
}

#endregion

