[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$failures = [System.Collections.Generic.List[string]]::new()
$warnings = [System.Collections.Generic.List[string]]::new()

function Add-Failure([string]$Message) { $failures.Add($Message) }
function Add-Warning([string]$Message) { $warnings.Add($Message) }

$required = @(
    'index.html', 'gdc.css', 'gdc_app.js', 'gdc_data.js', 'model_export.js',
    'assets/nickelsdorf-masterplan-concept.png', 'CNAME'
)
foreach ($item in $required) {
    if (-not (Test-Path -LiteralPath (Join-Path $repo $item))) {
        Add-Failure "Missing release asset: $item"
    }
}

$index = Get-Content -LiteralPath (Join-Path $repo 'index.html') -Raw
$tokens = [regex]::Matches($index, '[?&]v=([A-Za-z0-9._-]+)') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique
if ($tokens.Count -ne 1) {
    Add-Failure 'index.html must use one matching cache token for CSS and JavaScript assets.'
}

$data = Get-Content -LiteralPath (Join-Path $repo 'gdc_data.js') -Raw
if ($data -match 'const\s+BEDATA\s*=') {
    Add-Failure 'gdc_data.js contains clear-text Burgenland Energie production aggregates (BEDATA). Use a sanitized public data module.'
}
if ($data -match '(?i)confidential under the NDA|confidential under the NDAs|NDA-derived') {
    Add-Failure 'A public asset contains an explicit confidential or NDA-derived source marker.'
}

$app = Get-Content -LiteralPath (Join-Path $repo 'gdc_app.js') -Raw
$unsafeClaims = @(
    'A confidential tenant from the U.S.',
    'PPA volumes carry no network charges.',
    'built &amp; owned by the Power SPV · no grid fees',
    'MW firm, 8,760 hours a year'
)
foreach ($claim in $unsafeClaims) {
    if ($app.Contains($claim)) { Add-Failure "Unqualified public claim remains: $claim" }
}
if ($app -match 'gateLocked|ACCESS_HASH|PASS_HASH') {
    Add-Warning 'Client-side access controls are present. They are navigation aids, not authentication.'
}

$git = Get-Command git -ErrorAction SilentlyContinue
if ($git) {
    & $git.Source -C $repo diff --check
    if ($LASTEXITCODE -ne 0) { Add-Failure 'git diff --check reported whitespace errors.' }
} else {
    Add-Warning 'Git was not available, so repository consistency checks were skipped.'
}

$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) {
    $bundledNode = Join-Path $env:USERPROFILE '.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node.exe'
    if (Test-Path -LiteralPath $bundledNode) { $node = Get-Item -LiteralPath $bundledNode }
}
if ($node) {
    foreach ($script in @('gdc_app.js', 'model_export.js')) {
        & $node.FullName --check (Join-Path $repo $script)
        if ($LASTEXITCODE -ne 0) { Add-Failure "JavaScript syntax check failed: $script" }
    }
} else {
    Add-Warning 'Node.js was not available, so JavaScript syntax checks were skipped.'
}

foreach ($warning in $warnings) { Write-Warning $warning }
if ($failures.Count -gt 0) {
    Write-Host "`nPUBLIC RELEASE BLOCKED" -ForegroundColor Red
    foreach ($failure in $failures) { Write-Host " - $failure" -ForegroundColor Red }
    exit 1
}

Write-Host 'Release checks passed.' -ForegroundColor Green
