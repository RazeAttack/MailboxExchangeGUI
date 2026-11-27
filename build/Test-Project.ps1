<#
.SYNOPSIS
    Prüft das Projekt: Quellen zusammengesetzt, Kodierung, Syntax, PSScriptAnalyzer, Version.

.DESCRIPTION
    Wird lokal und in GitHub Actions (CI + Release) verwendet.
    Exitcode 0 = alles OK, 1 = Fehler.

.EXAMPLE
    .\build\Test-Project.ps1
    .\build\Test-Project.ps1 -Analyzer
    .\build\Test-Project.ps1 -Analyzer -ExpectedVersion 4.0.0
#>
param(
    [switch]$Analyzer,
    [string]$ExpectedVersion
)

$ErrorActionPreference = 'Stop'
$Root   = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$Main   = Join-Path $Root 'Get-UserMailboxAccess.ps1'
$Failed = 0

function Write-Result([bool]$Ok, [string]$Text) {
    if ($Ok) { Write-Host "[OK]     $Text" -ForegroundColor Green }
    else     { Write-Host "[FEHLER] $Text" -ForegroundColor Red; $script:Failed++ }
}

# 1. Hauptskript passt zu src\
& (Join-Path $PSScriptRoot 'Merge-Source.ps1') -Check | Out-Host
Write-Result ($LASTEXITCODE -eq 0) 'Get-UserMailboxAccess.ps1 ist aus src\ zusammengesetzt'

# 2. Kodierung: alle .ps1 UTF-8 mit genau einer BOM
$AllPs1 = @(Get-ChildItem -Path $Root -Filter '*.ps1' -File -Recurse | Where-Object { $_.FullName -notmatch '\\\.git\\' })
foreach ($F in $AllPs1) {
    $B      = [System.IO.File]::ReadAllBytes($F.FullName)
    $HasBom = $B.Length -ge 3 -and $B[0] -eq 0xEF -and $B[1] -eq 0xBB -and $B[2] -eq 0xBF
    $Double = $B.Length -ge 6 -and $HasBom -and $B[3] -eq 0xEF -and $B[4] -eq 0xBB -and $B[5] -eq 0xBF
    $Rel    = $F.FullName.Substring($Root.Length + 1)
    Write-Result ($HasBom -and -not $Double) "Kodierung UTF-8 mit BOM: $Rel"
}

# 3. Syntax (Parser von Windows PowerShell)
foreach ($F in $AllPs1) {
    $Err = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($F.FullName, [ref]$null, [ref]$Err)
    $Rel = $F.FullName.Substring($Root.Length + 1)
    Write-Result (@($Err).Count -eq 0) "Syntax: $Rel"
    foreach ($E in @($Err)) { Write-Host ("         Zeile {0}: {1}" -f $E.Extent.StartLineNumber, $E.Message) -ForegroundColor Red }
}

# 4. Version im Skript
$Content = [System.IO.File]::ReadAllText($Main, [System.Text.Encoding]::UTF8)
$M = [regex]::Match($Content, "\`$script:AppVersion\s*=\s*'([^']+)'")
Write-Result $M.Success 'AppVersion im Skript gefunden'
if ($M.Success) {
    Write-Host "         AppVersion = $($M.Groups[1].Value)" -ForegroundColor DarkGray
    if ($ExpectedVersion) {
        $Exp = $ExpectedVersion.TrimStart('v', 'V')
        Write-Result ($M.Groups[1].Value -eq $Exp) "AppVersion passt zum Tag ($Exp)"
    }
}

# 5. PSScriptAnalyzer (nur Fehler brechen ab, Warnungen werden gezählt)
if ($Analyzer) {
    if (-not (Get-Module PSScriptAnalyzer -ListAvailable)) {
        Write-Result $false 'PSScriptAnalyzer ist nicht installiert (Install-Module PSScriptAnalyzer -Scope CurrentUser)'
    } else {
        Import-Module PSScriptAnalyzer
        $Targets = @($Main, (Join-Path $Root 'Build-MailboxAccessExe.ps1')) + @(Get-ChildItem (Join-Path $Root 'build') -Filter '*.ps1' | ForEach-Object FullName)
        foreach ($T in $Targets) {
            $Res  = @(Invoke-ScriptAnalyzer -Path $T -Severity Error, Warning, ParseError)
            $Errs = @($Res | Where-Object { [string]$_.Severity -in 'Error', 'ParseError' })
            $Warn = @($Res | Where-Object { [string]$_.Severity -eq 'Warning' })
            $Rel  = $T.Substring($Root.Length + 1)
            Write-Result ($Errs.Count -eq 0) "PSScriptAnalyzer: $Rel ($($Errs.Count) Fehler, $($Warn.Count) Warnungen)"
            foreach ($E in $Errs) { Write-Host ("         Zeile {0}: [{1}] {2}" -f $E.Line, $E.RuleName, $E.Message) -ForegroundColor Red }
            $Warn | Group-Object RuleName | Sort-Object Count -Descending |
                ForEach-Object { Write-Host ("         Warnung {0,4} x {1}" -f $_.Count, $_.Name) -ForegroundColor DarkYellow }
        }
    }
}

Write-Host ''
if ($Failed -gt 0) {
    Write-Host "$Failed Prüfung(en) fehlgeschlagen." -ForegroundColor Red
    exit 1
}
Write-Host 'Alle Prüfungen bestanden.' -ForegroundColor Green
exit 0
