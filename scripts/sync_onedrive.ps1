[CmdletBinding()]
param(
    [string]$Destination = "$env:USERPROFILE\OneDrive - nexwell.com\Output\Austria DC\github-pages-deploy"
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$manifestPath = Join-Path $repo 'MIRROR_MANIFEST.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json

$dirty = & git -C $repo status --porcelain
if ($LASTEXITCODE -ne 0) { throw 'Unable to read repository status.' }
if ($dirty) { throw 'Mirror refused: commit or discard local changes first.' }

$head = (& git -C $repo rev-parse HEAD).Trim()
$main = (& git -C $repo rev-parse origin/main).Trim()
if ($head -ne $main) { throw 'Mirror refused: local HEAD does not match origin/main.' }

if (-not (Test-Path -LiteralPath $Destination)) {
    New-Item -ItemType Directory -Path $Destination | Out-Null
}

foreach ($entry in $manifest.files) {
    $source = Join-Path $repo $entry
    $target = Join-Path $Destination $entry
    if (-not (Test-Path -LiteralPath $source)) { throw "Manifest source is missing: $entry" }
    if ((Get-Item -LiteralPath $source).PSIsContainer) {
        New-Item -ItemType Directory -Path $target -Force | Out-Null
        Get-ChildItem -LiteralPath $source -Force | Copy-Item -Destination $target -Recurse -Force
        foreach ($sourceFile in Get-ChildItem -LiteralPath $source -File -Recurse) {
            $relative = $sourceFile.FullName.Substring($source.Length).TrimStart('\')
            $targetFile = Join-Path $target $relative
            if (-not (Test-Path -LiteralPath $targetFile)) { throw "Mirror file is missing: $entry/$relative" }
            if ((Get-FileHash -LiteralPath $sourceFile.FullName -Algorithm SHA256).Hash -ne
                (Get-FileHash -LiteralPath $targetFile -Algorithm SHA256).Hash) {
                throw "Hash mismatch after mirroring: $entry/$relative"
            }
        }
        continue
    }
    $parent = Split-Path -Parent $target
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    Copy-Item -LiteralPath $source -Destination $target -Force
    $sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    $targetHash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
    if ($sourceHash -ne $targetHash) { throw "Hash mismatch after mirroring: $entry" }
}

Set-Content -LiteralPath (Join-Path $Destination 'MIRROR_HEAD.txt') -Value $head -Encoding ascii
Write-Host "OneDrive mirror verified at commit $head" -ForegroundColor Green
