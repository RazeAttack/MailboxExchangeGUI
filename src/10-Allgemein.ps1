#region ===== Allgemein =====

$script:Cache    = $null
$script:TenantId = $null

$script:DefaultProgress = {
    param($Activity, $Status, $Percent)
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $Percent
}

function Connect-EXOIfNeeded {
    $Conn = Get-ConnectionInformation -ErrorAction SilentlyContinue |
        Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
    if (-not $Conn) {
        Write-Host "Verbinde mit Exchange Online..." -ForegroundColor Cyan

        # WAM (normale M365-Kontoauswahl) braucht ein Konsolenfenster als Parent.
        # Ohne Konsole (z. B. PS2EXE -noConsole) -> unsichtbare Konsole erzeugen.
        if (-not ('Win32.ConsoleHelper' -as [type])) {
            Add-Type -Namespace Win32 -Name ConsoleHelper -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
[DllImport("kernel32.dll")] public static extern bool AllocConsole();
[DllImport("user32.dll")]   public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@
        }
        if ([Win32.ConsoleHelper]::GetConsoleWindow() -eq [IntPtr]::Zero) {
            [void][Win32.ConsoleHelper]::AllocConsole()
            $Hwnd = [Win32.ConsoleHelper]::GetConsoleWindow()
            if ($Hwnd -ne [IntPtr]::Zero) { [void][Win32.ConsoleHelper]::ShowWindow($Hwnd, 0) }  # 0 = SW_HIDE
        }

        try {
            # Standard: WAM mit Kontoauswahl
            Connect-ExchangeOnline -ShowBanner:$false
        } catch {
            # Fallback: Browser-Login, falls WAM kein Fenster bekommt
            $HasDisableWam = (Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')
            if ($_.Exception.Message -match 'window handle' -and $HasDisableWam) {
                Connect-ExchangeOnline -ShowBanner:$false -DisableWAM
            } else {
                throw
            }
        }
        $Conn = Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
    }
    $script:TenantId = [string]$Conn.TenantId
    return $Conn
}

function New-KeySet {
    param([object[]]$Values)
    $Set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($V in $Values) {
        if ($V -and "$V".Trim()) { [void]$Set.Add("$V".Trim()) }
    }
    return ,$Set
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$Action,
        [string]$What = 'Aufruf',
        [int]$MaxAttempts = 3,
        [int]$DelaySeconds = 3,
        [scriptblock]$OnProgress
    )
    for ($Attempt = 1; $Attempt -le $MaxAttempts; $Attempt++) {
        try {
            return (& $Action)
        } catch {
            if ($Attempt -ge $MaxAttempts) { throw }
            $Wait = $DelaySeconds * $Attempt
            if ($OnProgress) {
                & $OnProgress 'Wiederholung' "$What fehlgeschlagen ($($_.Exception.Message)) - Versuch $($Attempt + 1)/$MaxAttempts in $Wait s" -1
            }
            Start-Sleep -Seconds $Wait
        }
    }
}

#endregion

