<#
.SYNOPSIS
    Bereitet eine neue Version vor: AppVersion hochsetzen, CHANGELOG-Eintrag, prüfen, committen, pushen, taggen.

.DESCRIPTION
    Schritt 1: $script:AppVersion in src\00-Header.ps1 setzen
    Schritt 2: Abschnitt "## [x.y.z] – Datum" oben in CHANGELOG.md einfügen
    Schritt 3: build\Merge-Source.ps1 + build\Test-Project.ps1
    Optional:
      -Push  -> git add, commit, push
      -Tag   -> wartet auf grünes CI (mit GitHub CLI 'gh', sonst Rückfrage), dann Tag setzen und pushen
                -> der Release-Workflow baut die EXE und veröffentlicht das Release

.PARAMETER Bump
    Patch (4.0.0 -> 4.0.1), Minor (4.0.1 -> 4.1.0), Major (4.1.0 -> 5.0.0). Standard: Patch

.PARAMETER Version
    Feste Version statt -Bump, z. B. 4.2.0

.PARAMETER Notes
    Einträge für den CHANGELOG. Ohne Angabe wird interaktiv gefragt (leere Zeile = fertig).

.PARAMETER Section
    Überschrift im CHANGELOG: Behoben, Geändert, Hinzugefügt. Standard: Behoben

.EXAMPLE
    .\build\New-Release.ps1
    .\build\New-Release.ps1 -Notes 'Delta-Update: Fehler bei leerem Audit-Log behoben'
    .\build\New-Release.ps1 -Bump Minor -Section Hinzugefügt -Notes 'Neue Funktion X', 'Neue Funktion Y' -Push
    .\build\New-Release.ps1 -Version 4.0.1 -Notes 'Kleinigkeiten' -Push -Tag
#>
param(
    [ValidateSet('Patch', 'Minor', 'Major')]
    [string]$Bump = 'Patch',
    [string]$Version,
    [string[]]$Notes,
    [ValidateSet('Behoben', 'Geändert', 'Hinzugefügt')]
    [string]$Section = 'Behoben',
    [switch]$Push,
    [switch]$Tag
)

$ErrorActionPreference = 'Stop'
$Root      = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$Header    = Join-Path $Root 'src\00-Header.ps1'
$Changelog = Join-Path $Root 'CHANGELOG.md'
$Utf8Bom   = New-Object System.Text.UTF8Encoding($true)
$Utf8      = New-Object System.Text.UTF8Encoding($false)

function Invoke-Git {
    # einfache Funktion ohne param(): $args reicht git-Schalter wie -m unverändert durch
    & git -C $Root @args
    if ($LASTEXITCODE -ne 0) { throw "git $($args -join ' ') fehlgeschlagen (Exitcode $LASTEXITCODE)" }
}

# --- 1. Aktuelle Version lesen ---
$HeaderText = [System.IO.File]::ReadAllText($Header, [System.Text.Encoding]::UTF8).TrimStart([char]0xFEFF)
$Pattern    = "(\`$script:AppVersion\s*=\s*')([0-9]+\.[0-9]+\.[0-9]+)(')"
$M = [regex]::Match($HeaderText, $Pattern)
if (-not $M.Success) { throw "AppVersion nicht gefunden in $Header" }
$Old = [version]$M.Groups[2].Value

if ($Version) {
    $New = [version]$Version.TrimStart('v', 'V')
} else {
    switch ($Bump) {
        'Major' { $New = [version]::new($Old.Major + 1, 0, 0) }
        'Minor' { $New = [version]::new($Old.Major, $Old.Minor + 1, 0) }
        default { $New = [version]::new($Old.Major, $Old.Minor, $Old.Build + 1) }
    }
}
$NewText = '{0}.{1}.{2}' -f $New.Major, $New.Minor, [Math]::Max($New.Build, 0)
if ([version]$NewText -le $Old) { throw "Neue Version $NewText muss größer sein als $Old." }

Write-Host "Version: $Old -> $NewText" -ForegroundColor Cyan

# Tag darf noch nicht existieren
$Existing = & git -C $Root tag --list "v$NewText"
if ($Existing) { throw "Tag v$NewText existiert bereits." }

# --- 2. CHANGELOG-Einträge ---
if (-not $Notes -or $Notes.Count -eq 0) {
    Write-Host "CHANGELOG-Einträge für '$Section' eingeben (leere Zeile = fertig):" -ForegroundColor Yellow
    $List = New-Object System.Collections.Generic.List[string]
    while ($true) {
        $Line = Read-Host '  -'
        if (-not $Line) { break }
        $List.Add($Line)
    }
    $Notes = $List.ToArray()
}
if (-not $Notes -or $Notes.Count -eq 0) { throw 'Keine CHANGELOG-Einträge angegeben - abgebrochen, nichts geändert.' }

# --- 3. Dateien ändern ---
$HeaderNew = [regex]::Replace($HeaderText, $Pattern, ('${1}' + $NewText + '${3}'), 'None', [TimeSpan]::FromSeconds(5))
$HeaderNew = $HeaderNew -replace "Version $($Old.Major)\.$($Old.Minor) - ", "Version $($New.Major).$($New.Minor) - "
[System.IO.File]::WriteAllText($Header, $HeaderNew, $Utf8Bom)
Write-Host "[OK] src\00-Header.ps1: AppVersion = '$NewText'" -ForegroundColor Green

$Cl  = [System.IO.File]::ReadAllText($Changelog, [System.Text.Encoding]::UTF8).TrimStart([char]0xFEFF)
$Nl  = if ($Cl.Contains("`r`n")) { "`r`n" } else { "`n" }
$Entry = "## [$NewText] – $(Get-Date -Format 'yyyy-MM-dd')$Nl$Nl### $Section$Nl" +
         (($Notes | ForEach-Object { "- $_" }) -join $Nl) + "$Nl$Nl"
$Head = "# Changelog$Nl$Nl"
if ($Cl.StartsWith($Head)) { $Cl = $Head + $Entry + $Cl.Substring($Head.Length) }
else                       { $Cl = $Head + $Entry + $Cl }
[System.IO.File]::WriteAllText($Changelog, $Cl, $Utf8)
Write-Host "[OK] CHANGELOG.md: Abschnitt [$NewText] eingefügt" -ForegroundColor Green

# --- 4. Zusammensetzen + prüfen ---
& (Join-Path $PSScriptRoot 'Merge-Source.ps1')
& (Join-Path $PSScriptRoot 'Test-Project.ps1')
if ($LASTEXITCODE -ne 0) { throw 'Prüfung fehlgeschlagen - bitte Fehler beheben (Dateien sind bereits geändert).' }

Write-Host ''
& git -C $Root status --short

if (-not $Push) {
    Write-Host ''
    Write-Host "Fertig vorbereitet. Weiter mit:" -ForegroundColor Cyan
    Write-Host "  git add -A; git commit -m `"chore: Version $NewText`"; git push"
    Write-Host "  (nach grünem CI)  git tag v$NewText; git push origin v$NewText"
    Write-Host "  oder alles auf einmal: .\build\New-Release.ps1 -Version $NewText ... -Push -Tag"
    return
}

# --- 5. Commit + Push ---
$Msg = "chore: Version $NewText`n`n" + (($Notes | ForEach-Object { "- $_" }) -join "`n")
Invoke-Git add -A
Invoke-Git commit -m $Msg
Invoke-Git push
$Sha = (& git -C $Root rev-parse HEAD).Trim()
Write-Host "[OK] Gepusht: $Sha" -ForegroundColor Green

if (-not $Tag) {
    Write-Host ''
    Write-Host "Nach grünem CI: git tag v$NewText; git push origin v$NewText" -ForegroundColor Cyan
    return
}

# --- 6. Auf CI warten, dann taggen ---
$Gh = Get-Command gh -ErrorAction SilentlyContinue
if ($Gh) {
    Write-Host 'Warte auf CI (GitHub CLI)...' -ForegroundColor Yellow
    $RunId = $null
    $ErrorActionPreference = 'Continue'   # gh schreibt Hinweise nach stderr
    for ($i = 0; $i -lt 30 -and -not $RunId; $i++) {
        Start-Sleep -Seconds 5
        $Json = & gh run list --workflow ci.yml --commit $Sha --json databaseId --limit 1 2>$null
        if ($LASTEXITCODE -eq 0 -and $Json) {
            $Runs = @($Json | ConvertFrom-Json)
            if ($Runs.Count -gt 0) { $RunId = $Runs[0].databaseId }
        }
    }
    if (-not $RunId) { throw 'CI-Lauf nicht gefunden - Tag manuell setzen, sobald CI grün ist.' }
    & gh run watch $RunId --exit-status
    $WatchExit = $LASTEXITCODE
    $ErrorActionPreference = 'Stop'
    if ($WatchExit -ne 0) { throw "CI ist fehlgeschlagen - kein Tag gesetzt. Details: gh run view $RunId --log-failed" }
    Write-Host '[OK] CI grün' -ForegroundColor Green
} else {
    Write-Host ''
    Write-Host 'GitHub CLI (gh) nicht installiert - bitte CI im Browser prüfen:' -ForegroundColor Yellow
    Write-Host '  https://github.com/RazeAttack/MailboxExchangeGUI/actions'
    $Ans = Read-Host 'Ist der CI-Lauf grün? (j/n)'
    if ($Ans -notmatch '^[jJyY]') { Write-Host "Kein Tag gesetzt. Später: git tag v$NewText; git push origin v$NewText"; return }
}

Invoke-Git tag "v$NewText"
Invoke-Git push origin "v$NewText"
Write-Host "[OK] Tag v$NewText gepusht - Release-Workflow läuft:" -ForegroundColor Green
Write-Host '  https://github.com/RazeAttack/MailboxExchangeGUI/actions/workflows/release.yml'
