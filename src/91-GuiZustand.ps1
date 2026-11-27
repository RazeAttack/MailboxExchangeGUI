#region ===== GUI: gemeinsamer Zustand =====

# Rechteprüfung: EXO lädt nur Cmdlets, die die RBAC-Rollen des Admins erlauben
$CanWrite = [bool]((Get-Command Add-MailboxPermission    -ErrorAction SilentlyContinue) -and
                   (Get-Command Remove-MailboxPermission -ErrorAction SilentlyContinue) -and
                   (Get-Command Add-RecipientPermission  -ErrorAction SilentlyContinue))

$Sync = [hashtable]::Synchronized(@{
    Queue            = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    Busy             = $false
    Closed           = $false
    Cancel           = $false
    StatusText       = 'Bereit'
    Percent          = 0
    Results          = $null
    ResultUser       = ''
    ResultVersion    = 0
    ErrorText        = ''
    ErrorVersion     = 0
    ReportText       = ''
    ReportFailed     = $false
    ReportVersion    = 0
    ReportTitle      = ''
    CacheInfo        = (Get-CacheInfo)
    CacheStale       = (Test-CacheStale)
    CacheHasRooms    = (Test-CacheHasRooms)
    HasCache         = [bool]$script:Cache
    HasCalendar      = [bool]($script:Cache -and $script:Cache.HasCalendar)
    MailboxList      = (Get-MailboxPickList)
    MbResult         = $null
    MbVersion        = 0
    Extra            = $null
    ExtraVersion     = 0
    GroupData        = $null
    GroupVersion     = 0
    RecheckUser      = ''
    RecheckAddrs     = @()
    RecheckVersion   = 0
    DataDir          = (Split-Path $CachePath -Parent)
    CanWrite         = $CanWrite
    CacheVersion     = 0
    ConnUpn          = [string]$ConnInfo.UserPrincipalName
    IconPath         = $null
    InitialUser      = $User
    InitialCalendar  = [bool]$IncludeCalendar
    RebuildCache     = [bool]$RebuildCache
    MaxCacheAgeHours = $MaxCacheAgeHours
    CanAudit         = (Test-AuditAvailable)
    AppVersion       = $script:AppVersion
    RepoSlug         = $script:RepoSlug
    UpdateScript     = $script:UpdateCheckScript.ToString()
    Update           = $null
    UpdateVersion    = 0
})

# Fortschritt aus dem Arbeits-Thread an die GUI melden (+ Abbruch prüfen)
$SyncProgress = {
    param($Activity, $Text, $Percent)
    $Sync.StatusText = "$Activity - $Text"
    $Sync.Percent    = [int]$Percent
    if ($Sync.Cancel) { throw [System.OperationCanceledException]::new('Vom Benutzer abgebrochen') }
}

#endregion

