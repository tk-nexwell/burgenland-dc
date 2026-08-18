[CmdletBinding()]
param(
    [switch]$Register,
    [switch]$Unregister,
    [int]$IntervalMinutes = 15,
    [switch]$SkipMirror
)

# Keeps the canonical Windows checkout, and the generated OneDrive mirror, consistent with the
# published GitHub main branch without anyone running a release script by hand.
#
#   powershell -File scripts\auto_sync.ps1              run one synchronisation now
#   powershell -File scripts\auto_sync.ps1 -Register    install the scheduled task
#   powershell -File scripts\auto_sync.ps1 -Unregister  remove the scheduled task
#
# GitHub Pages already publishes every commit pushed to main, so this script only moves in one
# direction: GitHub is the record, the local checkout follows it. It never pushes and never
# discards real local work. If the checkout carries genuine edits, it stops and says so, because a
# silent reset would destroy them. Line-ending-only differences are not genuine edits: this
# checkout stores LF and Windows editors write CRLF, so a plain 'git status' marks every file as
# modified even when the content matches the commit exactly.

$ErrorActionPreference = 'Stop'
$repo = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot)).TrimEnd([char[]]'\/')
$taskName = 'GDC Nickelsdorf mirror sync'

function Write-Step([string]$Message) { Write-Host "  $Message" }

function Invoke-RepoGit {
    param([Parameter(Mandatory = $true)][string[]]$Arguments, [switch]$AllowFailure)

    # git reports ordinary progress on stderr, and 'fetch' always does. Under
    # $ErrorActionPreference = 'Stop', merging that into the success stream makes PowerShell raise
    # NativeCommandError and abort a run that in fact succeeded. Judge the result by the exit code.
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $output = & git -C $repo @Arguments 2>&1
        $code = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previous
    }
    if ($code -ne 0 -and -not $AllowFailure) {
        throw "git $($Arguments -join ' ') failed: $(($output | Out-String).Trim())"
    }
    return ($output | Out-String).Trim()
}

if ($Register -and $Unregister) { throw 'Choose either -Register or -Unregister, not both.' }

if ($Unregister) {
    schtasks.exe /Delete /TN $taskName /F | Out-Null
    Write-Host "Removed the scheduled task '$taskName'." -ForegroundColor Green
    exit 0
}

if ($Register) {
    if ($IntervalMinutes -lt 5 -or $IntervalMinutes -gt 1440) {
        throw 'IntervalMinutes must be between 5 and 1440.'
    }
    $script = Join-Path $repo 'scripts\auto_sync.ps1'
    $command = "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$script`""
    schtasks.exe /Create /TN $taskName /TR $command /SC MINUTE /MO $IntervalMinutes /F | Out-Null
    Write-Host "Registered '$taskName' to run every $IntervalMinutes minutes." -ForegroundColor Green
    Write-Host "Run it now with: schtasks /Run /TN `"$taskName`""
    exit 0
}

Write-Host "Synchronising $repo with origin/main"

Invoke-RepoGit -Arguments @('fetch', 'origin', 'main', '--prune') | Out-Null
$local = Invoke-RepoGit -Arguments @('rev-parse', 'HEAD')
$remote = Invoke-RepoGit -Arguments @('rev-parse', 'origin/main')

# Real content differences, with line-ending-only noise excluded.
$dirty = Invoke-RepoGit -Arguments @('diff', '--ignore-cr-at-eol', '--name-only', 'HEAD')
$staged = Invoke-RepoGit -Arguments @('diff', '--ignore-cr-at-eol', '--cached', '--name-only')
$untracked = Invoke-RepoGit -Arguments @('ls-files', '--others', '--exclude-standard')

if ($dirty -or $staged) {
    Write-Warning 'The checkout carries genuine local edits, so nothing was reset.'
    Write-Warning 'Commit and push them, or discard them, then run this script again.'
    ($dirty + "`n" + $staged).Split("`n") | Where-Object { $_ } | Sort-Object -Unique |
        ForEach-Object { Write-Warning "  modified: $_" }
    exit 1
}
if ($untracked) {
    Write-Step 'Untracked files are present and were left alone:'
    $untracked.Split("`n") | Where-Object { $_ } | ForEach-Object { Write-Step "  $_" }
}

if ($local -eq $remote) {
    Write-Step "Already at the published commit $($local.Substring(0, 7))."
} else {
    Write-Step "Moving from $($local.Substring(0, 7)) to the published $($remote.Substring(0, 7))."
    Invoke-RepoGit -Arguments @('reset', '--hard', 'origin/main') | Out-Null
}

# Normalise the working tree back to the committed bytes. Windows editors reintroduce CRLF, which
# leaves every file looking modified and makes a later real change hard to see in git status.
$eolNoise = Invoke-RepoGit -Arguments @('diff', '--name-only')
if ($eolNoise) {
    Write-Step 'Rewriting line-ending-only differences back to the committed bytes.'
    Invoke-RepoGit -Arguments @('checkout', '--', '.') | Out-Null
}

if ($SkipMirror) {
    Write-Host 'Checkout is in step with origin/main. Mirror rebuild skipped.' -ForegroundColor Green
    exit 0
}

$sync = Join-Path $repo 'scripts\sync_onedrive.ps1'
if (-not (Test-Path -LiteralPath $sync -PathType Leaf)) {
    throw 'scripts/sync_onedrive.ps1 is missing, so the OneDrive mirror cannot be rebuilt.'
}

$mirrorDestination = $null
$manifestPath = Join-Path $repo 'MIRROR_MANIFEST.json'
if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
    $mirrorDestination = (Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json).mirror.destination
}

$headFile = if ($mirrorDestination) { Join-Path $mirrorDestination 'MIRROR_HEAD.txt' } else { $null }
$mirrorHead = if ($headFile -and (Test-Path -LiteralPath $headFile -PathType Leaf)) {
    (Get-Content -LiteralPath $headFile -Raw).Trim()
} else { '' }

if ($mirrorHead -eq $remote) {
    Write-Host "Checkout and OneDrive mirror are both at $($remote.Substring(0, 7))." -ForegroundColor Green
    exit 0
}

Write-Step 'Rebuilding the OneDrive mirror from the published commit.'
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $sync -Apply
if ($LASTEXITCODE -ne 0) { throw 'scripts/sync_onedrive.ps1 -Apply reported a failure.' }

Write-Host "Checkout and OneDrive mirror are both at $($remote.Substring(0, 7))." -ForegroundColor Green
