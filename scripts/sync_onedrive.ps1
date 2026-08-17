[CmdletBinding()]
param(
    [string]$Destination,
    [switch]$Apply
)

# One-way publication mirror. The Git checkout is authoritative; OneDrive is generated from a
# verified commit and never copied back. Without -Apply this script performs preflight checks only.
$ErrorActionPreference = 'Stop'
$repo = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot)).TrimEnd([char[]]'\/')
$manifestPath = Join-Path $repo 'MIRROR_MANIFEST.json'

function Get-NormalizedFullPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'A required path is empty.' }
    $expanded = [Environment]::ExpandEnvironmentVariables($Path)
    return [IO.Path]::GetFullPath($expanded).TrimEnd([char[]]'\/')
}

function Invoke-RepoGit {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)

    $output = & git -C $repo @Arguments 2>&1
    if ($LASTEXITCODE -ne 0) {
        $detail = ($output | Out-String).Trim()
        throw "git $($Arguments -join ' ') failed: $detail"
    }
    return $output
}

function ConvertTo-SafeRelativePath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $relative = $Path.Replace('\', '/').Trim()
    if ([string]::IsNullOrWhiteSpace($relative) -or
        [IO.Path]::IsPathRooted($relative) -or
        $relative.StartsWith('/') -or
        $relative.EndsWith('/') -or
        $relative -match '(^|/)\.\.?(/|$)' -or
        $relative -match '//' -or
        $relative -match '[:*?"<>|]') {
        throw "Unsafe path in MIRROR_MANIFEST.json: '$Path'"
    }
    return $relative
}

function Get-RelativeFileList {
    param([Parameter(Mandatory = $true)][string]$Root)

    $normalizedRoot = Get-NormalizedFullPath -Path $Root
    $prefix = $normalizedRoot + [IO.Path]::DirectorySeparatorChar
    return @(
        Get-ChildItem -LiteralPath $normalizedRoot -Recurse -Force -File |
            ForEach-Object {
                $_.FullName.Substring($prefix.Length).Replace('\', '/')
            }
    )
}

function Assert-ExactFileSet {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string[]]$Expected
    )

    if (-not (Test-Path -LiteralPath $Root -PathType Container)) {
        throw "Mirror directory is missing: $Root"
    }
    if (Test-Path -LiteralPath (Join-Path $Root '.git')) {
        throw "A .git entry was found in the generated mirror: $Root"
    }

    $actual = @(Get-RelativeFileList -Root $Root | Sort-Object)
    $wanted = @($Expected | Sort-Object)
    $difference = @(Compare-Object -ReferenceObject $wanted -DifferenceObject $actual -CaseSensitive)
    if ($difference.Count -gt 0) {
        $detail = ($difference | ForEach-Object { "$($_.SideIndicator) $($_.InputObject)" }) -join ', '
        throw "Generated mirror has missing or extra files: $detail"
    }
}

function Get-GitBlobForFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    return ((Invoke-RepoGit -Arguments @('hash-object', '--no-filters', '--', $Path)) |
        Out-String).Trim()
}

function Write-Utf8NoBom {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Text
    )

    $encoding = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, $Text, $encoding)
}

function Assert-MirrorProof {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string[]]$DeclaredFiles,
        [Parameter(Mandatory = $true)][string]$ProofFile,
        [Parameter(Mandatory = $true)][string]$CommitFile,
        [Parameter(Mandatory = $true)][string]$Commit
    )

    $expectedFiles = @($DeclaredFiles) + @($ProofFile, $CommitFile)
    Assert-ExactFileSet -Root $Root -Expected $expectedFiles

    $headValue = (Get-Content -LiteralPath (Join-Path $Root $CommitFile) -Raw).Trim()
    if ($headValue -ne $Commit) {
        throw "Mirror commit marker '$headValue' does not match '$Commit'."
    }

    try {
        $proof = Get-Content -LiteralPath (Join-Path $Root $ProofFile) -Raw | ConvertFrom-Json
    } catch {
        throw "Mirror proof is not valid JSON: $($_.Exception.Message)"
    }
    if ([int]$proof.schemaVersion -ne 2 -or [string]$proof.status -ne 'verified-mirror') {
        throw 'Mirror proof schema or status is invalid.'
    }
    if ([string]$proof.commit -ne $Commit -or [string]$proof.originCommit -ne $Commit) {
        throw 'Mirror proof does not identify the verified origin commit.'
    }

    $proofEntries = @($proof.files)
    if ($proofEntries.Count -ne $DeclaredFiles.Count) {
        throw 'Mirror proof contains the wrong number of file records.'
    }
    $proofByPath = @{}
    foreach ($record in $proofEntries) {
        $recordPath = [string]$record.path
        if ($proofByPath.ContainsKey($recordPath)) {
            throw "Mirror proof contains a duplicate file record: $recordPath"
        }
        $proofByPath[$recordPath] = $record
    }

    foreach ($relative in $DeclaredFiles) {
        if (-not $proofByPath.ContainsKey($relative)) {
            throw "Mirror proof omits a declared file: $relative"
        }
        $record = $proofByPath[$relative]
        $fullPath = Join-Path $Root $relative
        $sha256 = (Get-FileHash -LiteralPath $fullPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $gitBlob = Get-GitBlobForFile -Path $fullPath
        $bytes = (Get-Item -LiteralPath $fullPath).Length
        if ([string]$record.sha256 -ne $sha256 -or
            [string]$record.gitBlob -ne $gitBlob -or
            [long]$record.bytes -ne $bytes) {
            throw "Mirror proof mismatch for: $relative"
        }
    }
}

function Assert-GeneratedCleanupPath {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Parent,
        [Parameter(Mandatory = $true)][string]$LeafPrefix
    )

    $fullPath = Get-NormalizedFullPath -Path $Path
    $actualParent = Get-NormalizedFullPath -Path (Split-Path -Parent $fullPath)
    $leaf = Split-Path -Leaf $fullPath
    if ($actualParent -ne $Parent -or
        -not $leaf.StartsWith($LeafPrefix, [StringComparison]::Ordinal)) {
        throw "Refusing cleanup outside the generated staging scope: $fullPath"
    }
    return $fullPath
}

function Remove-GeneratedDirectory {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Parent,
        [Parameter(Mandatory = $true)][string]$LeafPrefix
    )

    $safePath = Assert-GeneratedCleanupPath -Path $Path -Parent $Parent -LeafPrefix $LeafPrefix
    if (Test-Path -LiteralPath $safePath) {
        if (-not (Test-Path -LiteralPath $safePath -PathType Container)) {
            throw "Refusing recursive cleanup of a non-directory path: $safePath"
        }
        Remove-Item -LiteralPath $safePath -Recurse -Force
    }
}

function Remove-GeneratedFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Parent,
        [Parameter(Mandatory = $true)][string]$LeafPrefix
    )

    $safePath = Assert-GeneratedCleanupPath -Path $Path -Parent $Parent -LeafPrefix $LeafPrefix
    if (Test-Path -LiteralPath $safePath) {
        if (-not (Test-Path -LiteralPath $safePath -PathType Leaf)) {
            throw "Refusing file cleanup of a non-file path: $safePath"
        }
        Remove-Item -LiteralPath $safePath -Force
    }
}

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    throw 'Git is required to build the OneDrive mirror.'
}
if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
    throw 'MIRROR_MANIFEST.json is missing.'
}

try {
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
} catch {
    throw "MIRROR_MANIFEST.json is not valid JSON: $($_.Exception.Message)"
}
if ([int]$manifest.version -ne 4 -or
    [string]$manifest.kind -ne 'canonical-source-to-onedrive-mirror') {
    throw 'MIRROR_MANIFEST.json must use canonical-source-to-onedrive-mirror schema version 4.'
}

$remote = [string]$manifest.canonical.remote
$branch = [string]$manifest.canonical.branch
$repository = [string]$manifest.canonical.repository
$expectedCheckout = Get-NormalizedFullPath -Path ([string]$manifest.canonical.localCheckout)
$expectedDestination = Get-NormalizedFullPath -Path ([string]$manifest.mirror.destination)
$proofFile = ConvertTo-SafeRelativePath -Path ([string]$manifest.mirror.proofFile)
$commitFile = ConvertTo-SafeRelativePath -Path ([string]$manifest.mirror.commitFile)
$backupPrefix = [string]$manifest.mirror.backupPrefix

if (-not $remote -or -not $branch -or -not $repository) {
    throw 'Manifest canonical repository, remote and branch values are required.'
}
if ($proofFile.Contains('/') -or $commitFile.Contains('/')) {
    throw 'Mirror proofFile and commitFile must be root-level filenames.'
}
if ([string]::IsNullOrWhiteSpace($backupPrefix) -or
    $backupPrefix -match '[\\/:*?"<>|]') {
    throw 'Manifest backupPrefix must be a safe filename prefix.'
}

if ($repo -ne $expectedCheckout) {
    throw "Mirror refused: run this script from the canonical checkout '$expectedCheckout', not '$repo'."
}
if (-not $Destination) { $Destination = [string]$manifest.mirror.destination }
$destinationFull = Get-NormalizedFullPath -Path $Destination
if ($destinationFull -ne $expectedDestination) {
    throw "Mirror refused: destination must be '$expectedDestination', not '$destinationFull'."
}
$destinationRoot = Get-NormalizedFullPath -Path ([IO.Path]::GetPathRoot($destinationFull))
if ($destinationFull -eq $destinationRoot -or $destinationFull -eq $repo) {
    throw "Unsafe mirror destination: $destinationFull"
}
$destinationParent = Get-NormalizedFullPath -Path (Split-Path -Parent $destinationFull)
$destinationLeaf = Split-Path -Leaf $destinationFull
if (-not (Test-Path -LiteralPath $destinationParent -PathType Container)) {
    throw "Mirror parent directory does not exist: $destinationParent"
}
if ($backupPrefix -ne "$destinationLeaf.backup.") {
    throw "Manifest backupPrefix must be '$destinationLeaf.backup.'."
}

$declaredFiles = @()
$seenFiles = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($value in @($manifest.releaseFiles) + @($manifest.controlFiles)) {
    $relative = ConvertTo-SafeRelativePath -Path ([string]$value)
    if (-not $seenFiles.Add($relative)) {
        throw "Duplicate or case-colliding manifest file: $relative"
    }
    $declaredFiles += $relative
}
if ($declaredFiles.Count -eq 0) { throw 'Manifest declares no mirror files.' }
if ($seenFiles.Contains($proofFile) -or $seenFiles.Contains($commitFile)) {
    throw 'Generated proof filenames must not also be declared tracked files.'
}

$topLevel = ((Invoke-RepoGit -Arguments @('rev-parse', '--show-toplevel')) | Out-String).Trim()
if ((Get-NormalizedFullPath -Path $topLevel) -ne $repo) {
    throw "Canonical path is not the root of this Git checkout: $repo"
}

[void](Invoke-RepoGit -Arguments @('fetch', '--prune', $remote))

$currentBranch = ((Invoke-RepoGit -Arguments @('branch', '--show-current')) | Out-String).Trim()
if ($currentBranch -ne $branch) {
    throw "Mirror refused: current branch is '$currentBranch', expected '$branch'."
}
$dirty = @(
    Invoke-RepoGit -Arguments @('status', '--porcelain=v1', '--untracked-files=all') |
        Where-Object { $_ -and ([string]$_).Trim() }
)
if ($dirty.Count -gt 0) {
    throw "Mirror refused: the canonical checkout is not clean.`n$($dirty -join [Environment]::NewLine)"
}

$head = ((Invoke-RepoGit -Arguments @('rev-parse', 'HEAD')) | Out-String).Trim()
$originHead = ((Invoke-RepoGit -Arguments @('rev-parse', "$remote/$branch")) | Out-String).Trim()
if ($head -ne $originHead) {
    throw "Mirror refused: HEAD $head does not match fetched $remote/$branch $originHead."
}

$remoteUrl = ((Invoke-RepoGit -Arguments @('remote', 'get-url', $remote)) | Out-String).Trim()
$allowedRemoteUrls = @(
    "https://github.com/$repository",
    "https://github.com/$repository.git",
    "git@github.com:$repository",
    "git@github.com:$repository.git",
    "ssh://git@github.com/$repository",
    "ssh://git@github.com/$repository.git"
)
if ($allowedRemoteUrls -notcontains $remoteUrl) {
    throw "Mirror refused: remote URL '$remoteUrl' is not an approved URL for '$repository'."
}

$treeFiles = @(
    Invoke-RepoGit -Arguments @('ls-tree', '-r', '--name-only', $head) |
        ForEach-Object { ([string]$_).Replace('\', '/') }
)
$treeDifference = @(
    Compare-Object -ReferenceObject @($declaredFiles | Sort-Object) `
        -DifferenceObject @($treeFiles | Sort-Object) -CaseSensitive
)
if ($treeDifference.Count -gt 0) {
    $detail = ($treeDifference | ForEach-Object { "$($_.SideIndicator) $($_.InputObject)" }) -join ', '
    throw "Tracked Git tree does not exactly match the manifest: $detail"
}
foreach ($patternValue in @($manifest.localOnlyPatterns)) {
    $pattern = ([string]$patternValue).Replace('\', '/')
    foreach ($trackedPath in $treeFiles) {
        if ($trackedPath -like $pattern) {
            throw "Local-only path is tracked and cannot be mirrored: $trackedPath"
        }
    }
}

$powerShell = Get-Command pwsh -ErrorAction SilentlyContinue
if (-not $powerShell) { $powerShell = Get-Command powershell -ErrorAction SilentlyContinue }
if (-not $powerShell) { throw 'PowerShell is required to run the release check.' }
& $powerShell.Source -NoProfile -NonInteractive -ExecutionPolicy Bypass `
    -File (Join-Path $repo 'scripts\check_release.ps1')
if ($LASTEXITCODE -ne 0) {
    throw 'Public release checks failed; the OneDrive mirror was not changed.'
}

Write-Host "Preflight passed for $head." -ForegroundColor Green
if (-not $Apply) {
    Write-Host 'No files were copied or renamed. Re-run with -Apply to replace the OneDrive mirror.' `
        -ForegroundColor Yellow
    return
}

$operationId = [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '.' +
    [Guid]::NewGuid().ToString('N')
$stagePrefix = ".$destinationLeaf.staging."
$archivePrefix = ".$destinationLeaf.archive."
$stagePath = Join-Path $destinationParent ($stagePrefix + $operationId)
$archivePath = Join-Path $destinationParent ($archivePrefix + $operationId + '.zip')
$shortCommit = $head.Substring(0, 12)
$backupPath = Join-Path $destinationParent ($backupPrefix +
    [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '.' + $shortCommit)
if (Test-Path -LiteralPath $backupPath) {
    $backupPath += '.' + [Guid]::NewGuid().ToString('N').Substring(0, 8)
}

[void](Assert-GeneratedCleanupPath -Path $stagePath -Parent $destinationParent `
    -LeafPrefix $stagePrefix)
[void](Assert-GeneratedCleanupPath -Path $archivePath -Parent $destinationParent `
    -LeafPrefix $archivePrefix)
[void](Assert-GeneratedCleanupPath -Path $backupPath -Parent $destinationParent `
    -LeafPrefix $backupPrefix)
if ((Test-Path -LiteralPath $stagePath) -or (Test-Path -LiteralPath $archivePath)) {
    throw 'Generated staging paths unexpectedly already exist.'
}
if (Test-Path -LiteralPath $backupPath) {
    throw "Recoverable backup path unexpectedly already exists: $backupPath"
}

$oldMirrorMoved = $false
$newMirrorInstalled = $false
$preserveStageOnFailure = $false
try {
    [void](New-Item -ItemType Directory -Path $stagePath)
    $archiveArguments = @('archive', '--format=zip', "--output=$archivePath", $head, '--') +
        @($declaredFiles)
    [void](Invoke-RepoGit -Arguments $archiveArguments)
    Expand-Archive -LiteralPath $archivePath -DestinationPath $stagePath
    Assert-ExactFileSet -Root $stagePath -Expected $declaredFiles

    $fileProof = @()
    foreach ($relative in $declaredFiles) {
        $fullPath = Join-Path $stagePath $relative
        $treeBlob = ((Invoke-RepoGit -Arguments @('rev-parse', "${head}:$relative")) |
            Out-String).Trim()
        $exportedBlob = Get-GitBlobForFile -Path $fullPath
        if ($treeBlob -ne $exportedBlob) {
            throw "Exported Git content mismatch for: $relative"
        }
        $fileProof += [ordered]@{
            path = $relative
            bytes = (Get-Item -LiteralPath $fullPath).Length
            sha256 = (Get-FileHash -LiteralPath $fullPath -Algorithm SHA256).Hash.ToLowerInvariant()
            gitBlob = $treeBlob
        }
    }

    $proof = [ordered]@{
        schemaVersion = 2
        status = 'verified-mirror'
        generatedAtUtc = [DateTime]::UtcNow.ToString('o')
        repository = $repository
        canonicalCheckout = $repo
        mirrorDestination = $destinationFull
        remote = $remote
        remoteUrl = $remoteUrl
        branch = $branch
        commit = $head
        originCommit = $originHead
        declaredFileCount = $declaredFiles.Count
        files = @($fileProof)
    }
    Write-Utf8NoBom -Path (Join-Path $stagePath $proofFile) `
        -Text (($proof | ConvertTo-Json -Depth 7) + "`n")
    Write-Utf8NoBom -Path (Join-Path $stagePath $commitFile) -Text ($head + "`n")
    Assert-MirrorProof -Root $stagePath -DeclaredFiles $declaredFiles -ProofFile $proofFile `
        -CommitFile $commitFile -Commit $head

    if (Test-Path -LiteralPath $destinationFull) {
        if (-not (Test-Path -LiteralPath $destinationFull -PathType Container)) {
            throw "Mirror destination exists but is not a directory: $destinationFull"
        }
        $destinationItem = Get-Item -LiteralPath $destinationFull -Force
        if (($destinationItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Mirror destination is a reparse point and will not be renamed: $destinationFull"
        }
        [IO.Directory]::Move($destinationFull, $backupPath)
        $oldMirrorMoved = $true
    }

    [IO.Directory]::Move($stagePath, $destinationFull)
    $newMirrorInstalled = $true
    Assert-MirrorProof -Root $destinationFull -DeclaredFiles $declaredFiles `
        -ProofFile $proofFile -CommitFile $commitFile -Commit $head

    Write-Host "OneDrive mirror verified at commit $head." -ForegroundColor Green
    Write-Host "Mirror: $destinationFull" -ForegroundColor Green
    if ($oldMirrorMoved) {
        Write-Host "Recoverable previous mirror: $backupPath" -ForegroundColor Green
    } else {
        Write-Host 'No previous mirror existed; no backup was needed.' -ForegroundColor Green
    }
} catch {
    $installError = $_.Exception
    $rollbackError = $null
    try {
        if ($newMirrorInstalled -and (Test-Path -LiteralPath $destinationFull)) {
            [IO.Directory]::Move($destinationFull, $stagePath)
            $newMirrorInstalled = $false
        }
        if ($oldMirrorMoved -and (Test-Path -LiteralPath $backupPath) -and
            -not (Test-Path -LiteralPath $destinationFull)) {
            [IO.Directory]::Move($backupPath, $destinationFull)
            $oldMirrorMoved = $false
        }
    } catch {
        $rollbackError = $_.Exception
    }
    if ($rollbackError) {
        $preserveStageOnFailure = $true
        throw "Mirror install failed: $($installError.Message) Rollback also failed: $($rollbackError.Message) Generated recovery material, if present, was preserved at '$stagePath'."
    }
    throw $installError
} finally {
    if (Test-Path -LiteralPath $archivePath) {
        Remove-GeneratedFile -Path $archivePath -Parent $destinationParent `
            -LeafPrefix $archivePrefix
    }
    if (-not $preserveStageOnFailure -and (Test-Path -LiteralPath $stagePath)) {
        Remove-GeneratedDirectory -Path $stagePath -Parent $destinationParent `
            -LeafPrefix $stagePrefix
    }
}
