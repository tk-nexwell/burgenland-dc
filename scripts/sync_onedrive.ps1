[CmdletBinding()]
param(
    [string]$Destination = "$env:USERPROFILE\OneDrive - nexwell.com\Output\Austria DC\github-pages-deploy"
)

$ErrorActionPreference = 'Stop'
$repo = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$destPath = [IO.Path]::GetFullPath($Destination)
$destParent = [IO.Path]::GetDirectoryName($destPath)
$manifestPath = Join-Path $repo 'MIRROR_MANIFEST.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json

if ($destPath -eq $repo) { throw 'Mirror refused: source and destination are the same directory.' }
if ([IO.Path]::GetPathRoot($destPath) -eq $destPath) { throw 'Mirror refused: destination cannot be a drive root.' }
if (-not $destParent) { throw 'Mirror refused: destination parent could not be resolved.' }

& git -C $repo fetch --quiet origin main
if ($LASTEXITCODE -ne 0) { throw 'Unable to fetch origin/main.' }
$dirty = & git -C $repo status --porcelain
if ($LASTEXITCODE -ne 0) { throw 'Unable to read repository status.' }
if ($dirty) { throw 'Mirror refused: commit or discard local changes first.' }
$head = (& git -C $repo rev-parse HEAD).Trim()
$remote = (& git -C $repo rev-parse origin/main).Trim()
if ($head -ne $remote) { throw 'Mirror refused: local HEAD does not match the freshly fetched origin/main.' }

if (-not (Test-Path -LiteralPath $destParent)) {
    New-Item -ItemType Directory -Path $destParent -Force | Out-Null
}
$stage = "$destPath.staging.$([guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $stage | Out-Null

try {
    foreach ($entry in $manifest.files) {
        $source = Join-Path $repo $entry
        $target = Join-Path $stage $entry
        if (-not (Test-Path -LiteralPath $source)) { throw "Manifest source is missing: $entry" }
        if ((Get-Item -LiteralPath $source).PSIsContainer) {
            Copy-Item -LiteralPath $source -Destination $target -Recurse -Force
        } else {
            $parent = Split-Path -Parent $target
            if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
            Copy-Item -LiteralPath $source -Destination $target -Force
        }
    }

    $hashes = foreach ($file in Get-ChildItem -LiteralPath $stage -File -Recurse | Sort-Object FullName) {
        [pscustomobject]@{
            path = $file.FullName.Substring($stage.Length).TrimStart('\').Replace('\','/')
            sha256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }
    [pscustomobject]@{ commit = $head; generated = (Get-Date).ToUniversalTime().ToString('o'); files = @($hashes) } |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $stage 'MIRROR_HASHES.json') -Encoding utf8
    Set-Content -LiteralPath (Join-Path $stage 'MIRROR_HEAD.txt') -Value $head -Encoding ascii

    foreach ($item in $hashes) {
        $sourceFile = Join-Path $repo $item.path
        $stageFile = Join-Path $stage $item.path
        if ((Get-FileHash -LiteralPath $sourceFile -Algorithm SHA256).Hash.ToLowerInvariant() -ne $item.sha256) {
            throw "Hash mismatch in staged mirror: $($item.path)"
        }
    }

    $archive = $null
    if (Test-Path -LiteralPath $destPath) {
        $archive = "$destPath.archive.$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        if (Test-Path -LiteralPath $archive) { $archive = "$archive.$([guid]::NewGuid().ToString('N').Substring(0,8))" }
        if ([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($archive)) -ne $destParent) {
            throw 'Mirror refused: archive target escaped the intended OneDrive parent.'
        }
        Move-Item -LiteralPath $destPath -Destination $archive
    }
    try {
        Move-Item -LiteralPath $stage -Destination $destPath
    } catch {
        if ($archive -and -not (Test-Path -LiteralPath $destPath) -and (Test-Path -LiteralPath $archive)) {
            Move-Item -LiteralPath $archive -Destination $destPath
        }
        throw
    }
    Write-Host "OneDrive exact mirror verified at commit $head" -ForegroundColor Green
    if ($archive) { Write-Host "Previous mirror preserved at $archive" -ForegroundColor Yellow }
} finally {
    if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force }
}
