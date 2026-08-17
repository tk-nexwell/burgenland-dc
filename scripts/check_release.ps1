[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$failures = [System.Collections.Generic.List[string]]::new()
$warnings = [System.Collections.Generic.List[string]]::new()

function Add-Failure([string]$Message) { $failures.Add($Message) }
function Add-Warning([string]$Message) { $warnings.Add($Message) }

$manifestPath = Join-Path $repo 'MIRROR_MANIFEST.json'
$manifest = $null
if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
    try {
        $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    } catch {
        Add-Failure "MIRROR_MANIFEST.json is not valid JSON: $($_.Exception.Message)"
    }
} else {
    Add-Failure 'Missing release control: MIRROR_MANIFEST.json'
}

$required = @(
    'index.html',
    'gdc.css',
    'gdc_app.js',
    'gdc_data.js',
    'model_export.js',
    'prices_hourly.csv',
    'assets/nickelsdorf-masterplan-future-state.png',
    'scripts/sync_onedrive.ps1',
    'scripts/test_model_integrity.mjs',
    'scripts/test_xirr.mjs',
    '.github/workflows/release-check.yml',
    '.gitignore',
    'CNAME'
)
foreach ($item in $required) {
    if (-not (Test-Path -LiteralPath (Join-Path $repo $item) -PathType Leaf)) {
        Add-Failure "Missing release asset or control: $item"
    }
}

$manifestFiles = @()
if ($manifest) {
    if ([int]$manifest.version -ne 4) {
        Add-Failure 'Manifest version must be 4.'
    }
    if ($manifest.kind -ne 'canonical-source-to-onedrive-mirror') {
        Add-Failure 'Manifest kind must be canonical-source-to-onedrive-mirror.'
    }
    if ([string]$manifest.canonical.repository -ne 'tk-nexwell/burgenland-dc' -or
        [string]$manifest.canonical.remote -ne 'origin' -or
        [string]$manifest.canonical.branch -ne 'main') {
        Add-Failure 'Manifest canonical repository must be tk-nexwell/burgenland-dc origin/main.'
    }

    $canonicalCheckout = ([string]$manifest.canonical.localCheckout).Replace('\', '/').TrimEnd('/')
    if ($canonicalCheckout -ne 'C:/Users/ThomasKoenig/GitHub/burgenland-dc') {
        Add-Failure 'Manifest canonical checkout must be C:/Users/ThomasKoenig/GitHub/burgenland-dc.'
    }
    $mirrorDestination = ([string]$manifest.mirror.destination).Replace('\', '/').TrimEnd('/')
    if ($mirrorDestination -ne
        'C:/Users/ThomasKoenig/OneDrive - nexwell.com/Output/Austria DC/github-pages-deploy') {
        Add-Failure 'Manifest OneDrive destination is not the approved generated-mirror path.'
    }
    if ([string]$manifest.mirror.proofFile -ne 'MIRROR_PROOF.json' -or
        [string]$manifest.mirror.commitFile -ne 'MIRROR_HEAD.txt' -or
        [string]$manifest.mirror.backupPrefix -ne 'github-pages-deploy.backup.') {
        Add-Failure 'Manifest proof, commit or backup naming is not the approved release-control set.'
    }
    if (@($manifest.releaseFiles) -notcontains 'assets/nickelsdorf-masterplan-future-state.png') {
        Add-Failure 'Manifest releaseFiles must include the future-state masterplan.'
    }

    $seenManifestFiles = [System.Collections.Generic.HashSet[string]]::new(
        [StringComparer]::OrdinalIgnoreCase
    )
    foreach ($itemValue in @($manifest.releaseFiles) + @($manifest.controlFiles)) {
        $item = ([string]$itemValue).Replace('\', '/').Trim()
        if ([string]::IsNullOrWhiteSpace($item) -or
            [IO.Path]::IsPathRooted($item) -or
            $item.StartsWith('/') -or
            $item.EndsWith('/') -or
            $item -match '(^|/)\.\.?(/|$)' -or
            $item -match '//' -or
            $item -match '[:*?"<>|]') {
            Add-Failure "Manifest contains an unsafe tracked path: $item"
            continue
        }
        if (-not $seenManifestFiles.Add($item)) {
            Add-Failure "Manifest contains a duplicate or case-colliding tracked path: $item"
            continue
        }
        $manifestFiles += $item
    }
    if ($manifestFiles.Count -eq 0) {
        Add-Failure 'Manifest declares no tracked mirror files.'
    }
    foreach ($generatedName in @('MIRROR_PROOF.json', 'MIRROR_HEAD.txt')) {
        if ($seenManifestFiles.Contains($generatedName)) {
            Add-Failure "Generated mirror proof must not be tracked or declared: $generatedName"
        }
    }

    foreach ($requiredLocalPattern in @(
        '*.bundle',
        'MIRROR_HEAD.txt',
        'MIRROR_HASHES.json',
        'MIRROR_PROOF.json',
        'bedata_enc.bin',
        'assets/nickelsdorf-masterplan-concept.png',
        'assets/nickelsdorf-masterplan-burgenland.png'
    )) {
        if (@($manifest.localOnlyPatterns) -notcontains $requiredLocalPattern) {
            Add-Failure "Manifest localOnlyPatterns omits: $requiredLocalPattern"
        }
    }
}

$indexPath = Join-Path $repo 'index.html'
$appPath = Join-Path $repo 'gdc_app.js'
$dataPath = Join-Path $repo 'gdc_data.js'
$modelPath = Join-Path $repo 'model_export.js'
$index = if (Test-Path -LiteralPath $indexPath) { Get-Content -LiteralPath $indexPath -Raw } else { '' }
$app = if (Test-Path -LiteralPath $appPath) { Get-Content -LiteralPath $appPath -Raw } else { '' }
$data = if (Test-Path -LiteralPath $dataPath) { Get-Content -LiteralPath $dataPath -Raw } else { '' }
$model = if (Test-Path -LiteralPath $modelPath) { Get-Content -LiteralPath $modelPath -Raw } else { '' }
$publicText = $index + "`n" + $app + "`n" + $data + "`n" + $model

$assetMatches = [regex]::Matches($index, '(?:href|src)="(gdc\.(?:css|js)|gdc_(?:app|data)\.js)\?v=([A-Za-z0-9._-]+)"')
$tokens = @($assetMatches | ForEach-Object { $_.Groups[2].Value } | Select-Object -Unique)
if ($assetMatches.Count -ne 3 -or $tokens.Count -ne 1) {
    Add-Failure 'index.html must load gdc.css, gdc_data.js and gdc_app.js with one matching cache token.'
}

$cnamePath = Join-Path $repo 'CNAME'
if (Test-Path -LiteralPath $cnamePath) {
    $cname = (Get-Content -LiteralPath $cnamePath -Raw).Trim()
    if ($cname -ne 'gdc-nickelsdorf.com') {
        Add-Failure "CNAME must be exactly gdc-nickelsdorf.com, found '$cname'."
    }
}

$futureImage = Join-Path $repo 'assets\nickelsdorf-masterplan-future-state.png'
if (Test-Path -LiteralPath $futureImage -PathType Leaf) {
    $imageInfo = Get-Item -LiteralPath $futureImage
    if ($imageInfo.Length -lt 250000) {
        Add-Failure 'The future-state masterplan is unexpectedly small; verify that the final render is present.'
    } else {
        $imageBytes = [IO.File]::ReadAllBytes($futureImage)
        $signature = [BitConverter]::ToString($imageBytes[0..7])
        if ($signature -ne '89-50-4E-47-0D-0A-1A-0A') {
            Add-Failure 'The future-state masterplan does not have a valid PNG signature.'
        }
    }
}
if ($app -notmatch 'assets/nickelsdorf-masterplan-future-state\.png') {
    Add-Failure 'gdc_app.js does not reference the required future-state masterplan.'
}
if ($app -notmatch '(?i)not an approved site plan') {
    Add-Failure 'The future-state visualization must state that it is not an approved site plan.'
}

foreach ($privateConstant in @('BEDATA', 'BENCH', 'NEWD', 'MEAS', 'CLIP')) {
    if ($data -match "(?m)^\s*const\s+$privateConstant\s*=") {
        Add-Failure "gdc_data.js contains a private or meter-derived dataset: $privateConstant."
    }
}
if ($publicText -match '(?i)confidential under the NDA|confidential under the NDAs|NDA-derived') {
    Add-Failure 'A public asset contains an explicit confidential or NDA-derived source marker.'
}
if ($publicText -match '(?i)bedata_enc\.bin|raw 15-minute data explorer') {
    Add-Failure 'Browser-delivered project-meter archive or explorer code remains in a public asset.'
}
if (Test-Path -LiteralPath (Join-Path $repo 'bedata_enc.bin')) {
    Add-Failure 'The encrypted project-meter archive must not exist inside the public checkout.'
}

$internalTermsPatterns = @(
    @{ Pattern = '(?i)\bTERMS_BLOB\b'; Message = 'Encrypted internal-terms payload remains in public code.' },
    @{ Pattern = '(?i)\btermsDerive\s*\('; Message = 'Internal-terms key derivation remains in public code.' },
    @{ Pattern = '(?i)\btermsTry\s*\('; Message = 'Internal-terms unlock code remains in public code.' },
    @{ Pattern = '(?i)\bapplyTerms\s*\('; Message = 'Internal negotiated-term application code remains in public code.' },
    @{ Pattern = '(?i)sessionStorage[^\r\n]*gdcTerms'; Message = 'Internal-terms browser storage remains in public code.' },
    @{ Pattern = '(?i)Commercial terms open'; Message = 'Public code contains the internal-terms unlocked state.' },
    @{ Pattern = '(?i)the ones in the executed drafts'; Message = 'Public code contains an executed-drafts claim.' }
)
foreach ($rule in $internalTermsPatterns) {
    if ($publicText -match $rule.Pattern) { Add-Failure $rule.Message }
}
if ($app -match '(?m)^\s*const\s+TERMS\s*=\s*(?!null\s*;)') {
    Add-Failure 'The public TERMS variable must be absent or explicitly null.'
}

if ($publicText -match '(?i)\bACCESS_HASH\b|\bPASS_HASH\b|function\s+gateTry\s*\(' -or
    $app -match '(?is)const\s+GATE\s*=\s*\{[^}]*\b(?:h|hash|salt)\s*:') {
    Add-Failure 'Public assets contain client-side credential or password-verifier material.'
}

$unsafeClaims = @(
    'A confidential tenant from the U.S.',
    'PPA volumes carry no network charges.',
    'built &amp; owned by the Power SPV · no grid fees',
    'MW firm, 8,760 hours a year',
    'Burgenland Energie guaranteed',
    'That is N+1 on electrons',
    'PHASE 1 UNDER CONSTRUCTION'
)
foreach ($claim in $unsafeClaims) {
    if ($publicText.Contains($claim)) { Add-Failure "Unqualified public claim remains: $claim" }
}

$git = Get-Command git -ErrorAction SilentlyContinue
if ($git) {
    $tracked = @(& $git.Source -C $repo ls-files)
    if ($LASTEXITCODE -ne 0) {
        Add-Failure 'Unable to enumerate tracked files.'
        $tracked = @()
    }

    foreach ($path in $tracked) {
        if ($path -ieq 'bedata_enc.bin' -or
            $path -ieq 'MIRROR_HEAD.txt' -or
            $path -ieq 'MIRROR_HASHES.json' -or
            $path -ieq 'MIRROR_PROOF.json' -or
            $path -like '*.bundle') {
            Add-Failure "Local-only artifact is tracked and would be public: $path"
        }
    }

    if ($manifest) {
        foreach ($patternValue in @($manifest.localOnlyPatterns)) {
            $pattern = ([string]$patternValue).Replace('\', '/')
            foreach ($path in $tracked) {
                if ($path -like $pattern) {
                    Add-Failure "Manifest local-only path is tracked and would be public: $path"
                }
            }
        }
    }

    foreach ($itemValue in $manifestFiles) {
        $item = ([string]$itemValue).Replace('\', '/')
        if ($tracked -cnotcontains $item) {
            Add-Failure "Manifest file is not tracked: $item"
        }
    }
    foreach ($path in $tracked) {
        if ($manifestFiles -cnotcontains $path) {
            Add-Failure "Tracked file is not declared for the public release or mirror: $path"
        }
    }

    & $git.Source -C $repo diff --check
    if ($LASTEXITCODE -ne 0) { Add-Failure 'git diff --check reported whitespace errors.' }
    & $git.Source -C $repo diff --cached --check
    if ($LASTEXITCODE -ne 0) { Add-Failure 'git diff --cached --check reported staged whitespace errors.' }
} else {
    Add-Warning 'Git was not available, so tracked-file and whitespace checks were skipped.'
}

$syncScript = Join-Path $repo 'scripts\sync_onedrive.ps1'
if (Test-Path -LiteralPath $syncScript) {
    $parseTokens = $null
    $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile(
        $syncScript,
        [ref]$parseTokens,
        [ref]$parseErrors
    )
    foreach ($parseError in @($parseErrors)) {
        Add-Failure "PowerShell syntax error in scripts/sync_onedrive.ps1: $($parseError.Message)"
    }
}

$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) {
    $bundledNode = Join-Path $env:USERPROFILE '.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node.exe'
    if (Test-Path -LiteralPath $bundledNode) { $node = Get-Item -LiteralPath $bundledNode }
}
if ($node) {
    foreach ($script in @(
        'gdc_data.js',
        'gdc_app.js',
        'model_export.js',
        'scripts/sanitize_public_release.mjs',
        'scripts/test_model_integrity.mjs',
        'scripts/test_xirr.mjs'
    )) {
        & $node.FullName --check (Join-Path $repo $script)
        if ($LASTEXITCODE -ne 0) { Add-Failure "JavaScript syntax check failed: $script" }
    }
    & $node.FullName (Join-Path $repo 'scripts/test_xirr.mjs')
    if ($LASTEXITCODE -ne 0) { Add-Failure 'XIRR regression checks failed.' }
    & $node.FullName (Join-Path $repo 'scripts/test_model_integrity.mjs')
    if ($LASTEXITCODE -ne 0) { Add-Failure 'Model integrity regression checks failed.' }
} else {
    Add-Warning 'Node.js was not available, so JavaScript syntax and XIRR checks were skipped.'
}

foreach ($warning in $warnings) { Write-Warning $warning }
if ($failures.Count -gt 0) {
    Write-Host "`nPUBLIC RELEASE BLOCKED" -ForegroundColor Red
    foreach ($failure in $failures) { Write-Host " - $failure" -ForegroundColor Red }
    exit 1
}

Write-Host 'Release checks passed.' -ForegroundColor Green
