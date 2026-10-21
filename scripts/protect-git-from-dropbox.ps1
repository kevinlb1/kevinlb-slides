<#
.SYNOPSIS
  Stop Dropbox interfering with .git, without breaking sync between machines.

.DESCRIPTION
  This repository lives inside Dropbox, so every file under .git is a Dropbox
  placeholder and each read and write goes through Dropbox's filter driver.
  Marking the folder "available offline" does not help: the attributes on
  .git/index are 0x80420, PINNED *and* REPARSE_POINT, so pinning fixes where the
  data lives while leaving every file a placeholder. The symptoms are
  intermittent and confusing:

    fatal: Unable to write index
    error: unable to write file .git/objects/...
    Unable to create '.git/index.lock': File exists

  -IgnoreRegenerable tells Dropbox to skip only the volatile, machine-local
  files that git rebuilds by itself. Everything a clone or a pull needs --
  objects, refs, packed-refs, HEAD, config -- keeps syncing, so the repository
  stays usable from any machine.

  What is excluded, and why it is safe:

    index          A stat cache over the working tree. Machine-specific by
                   nature, rewritten on nearly every git command, and the file
                   that fails most. Rebuilt with `git reset`, which this script
                   does automatically when it finds it missing.
    logs/          The reflog. Local history of where refs pointed here; means
                   nothing on another machine.
    FETCH_HEAD     Rewritten by every fetch.
    ORIG_HEAD      Written before a merge or reset.
    COMMIT_EDITMSG A scratch buffer for the last commit message.
    MERGE_*, AUTO_MERGE, CHERRY_PICK_HEAD, REVERT_HEAD
                   State for an operation in progress. Syncing a half-finished
                   merge to another machine is actively harmful.
    objects/info/commit-graphs
                   A cache. Rebuilt by `git gc` or `git commit-graph write`.
    gc.log         Output from a failed background gc.

  Run with no arguments to diagnose and repair. That is safe and idempotent, and
  is also what to run on a machine that has just picked the repository up.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File scripts\protect-git-from-dropbox.ps1
  powershell -ExecutionPolicy Bypass -File scripts\protect-git-from-dropbox.ps1 -IgnoreRegenerable
#>
[CmdletBinding()]
param(
    [switch]$IgnoreRegenerable,
    [int]$StaleLockMinutes = 30
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$gitDir = Join-Path $repo '.git'
if (-not (Test-Path $gitDir)) { throw "No .git directory at $repo" }

# Paths under .git that git recreates on demand. Nothing here is needed to
# clone, fetch, pull or push.
$regenerable = @(
    'index', 'logs', 'FETCH_HEAD', 'ORIG_HEAD', 'COMMIT_EDITMSG', 'gc.log',
    'AUTO_MERGE', 'MERGE_HEAD', 'MERGE_MSG', 'MERGE_MODE',
    'CHERRY_PICK_HEAD', 'REVERT_HEAD', 'BISECT_LOG',
    'objects\info\commit-graph', 'objects\info\commit-graphs'
)

Write-Host "Repository: $repo`n"

# Never touch a repository another git process is using.
$running = Get-Process git, git-remote-https -ErrorAction SilentlyContinue
if ($running) {
    Write-Warning "git is running (pid $($running.Id -join ', ')). Re-run when it has finished."
    exit 1
}

# --- Stale locks --------------------------------------------------------------
# git writes the new index into index.lock, so a zero-byte lock never got
# started. Age-gate anyway: a fresh lock may belong to a live operation.
$locks = @(Get-ChildItem $gitDir -Filter '*.lock' -Recurse -Force -ErrorAction SilentlyContinue)
foreach ($lock in $locks) {
    $age = (Get-Date) - $lock.CreationTime
    if ($age.TotalMinutes -ge $StaleLockMinutes) {
        Remove-Item $lock.FullName -Force
        Write-Host ("  removed stale lock: {0} ({1:N0} min old, {2} bytes)" -f $lock.Name, $age.TotalMinutes, $lock.Length)
    } else {
        Write-Warning ("  lock {0} is only {1:N0} min old -- left in place" -f $lock.Name, $age.TotalMinutes)
    }
}
if ($locks.Count -eq 0) { Write-Host "  no lock files" }

# --- Rehydrate placeholders ---------------------------------------------------
$files = Get-ChildItem $gitDir -Recurse -File -Force -ErrorAction SilentlyContinue
$placeholders = @($files | Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint })
$rehydrated = 0
foreach ($file in $placeholders) {
    try {
        $bytes = [IO.File]::ReadAllBytes($file.FullName)
        $tmp = "$($file.FullName).__hydrate"
        [IO.File]::WriteAllBytes($tmp, $bytes)
        Move-Item -LiteralPath $tmp -Destination $file.FullName -Force
        $rehydrated++
    } catch {
        Write-Warning "  could not rehydrate $($file.Name): $($_.Exception.Message)"
    }
}
Write-Host "  placeholders: $($placeholders.Count); rehydrated: $rehydrated"

# --- Corrupt commit-graph cache ----------------------------------------------
$graph = Join-Path $gitDir 'objects\info\commit-graph'
$fsck = & git -C $repo fsck --no-progress 2>&1 | Where-Object { $_ -match 'commit-graph' }
if ($fsck -and (Test-Path $graph)) {
    Remove-Item $graph -Force
    Write-Host "  removed corrupt commit-graph cache"
}

# --- Rebuild a missing index --------------------------------------------------
# Expected on a machine that has just synced the repository, because the index
# is excluded below. Without it git reports every tracked file as deleted.
if (-not (Test-Path (Join-Path $gitDir 'index'))) {
    Write-Host "  no index (normal on a freshly synced machine); rebuilding from HEAD"
    & git -C $repo reset --mixed --quiet
    if ($LASTEXITCODE -ne 0) { throw "git reset failed while rebuilding the index" }
}

# --- Exclude the regenerable paths from Dropbox -------------------------------
# Rehydrating a file rewrites it, and rewriting drops its alternate data
# streams -- including com.dropbox.ignored. So the repair pass above silently
# undoes the exclusions. The intent is therefore recorded in git config, which
# survives because rehydration preserves file contents, and the exclusions are
# re-applied on every run rather than only when the switch is passed.
if ($IgnoreRegenerable) {
    & git -C $repo config --local ai100.dropboxIgnoreRegenerable true
}
$wanted = (& git -C $repo config --local --get ai100.dropboxIgnoreRegenerable) -eq 'true'

if ($wanted) {
    Write-Host ""
    $applied = 0
    $already = 0
    foreach ($relative in $regenerable) {
        $path = Join-Path $gitDir $relative
        if (-not (Test-Path $path)) { continue }
        if (Get-Item -Path $path -Stream com.dropbox.ignored -ErrorAction SilentlyContinue) {
            $already++
            continue
        }
        Set-Content -Path $path -Stream com.dropbox.ignored -Value 1
        Write-Host "  Dropbox will ignore: .git\$relative"
        $applied++
    }
    Write-Host "  exclusions applied: $applied; already in place: $already"
    Write-Host ""
    Write-Host "Objects, refs, packed-refs, HEAD and config still sync, so this"
    Write-Host "repository remains clonable and pullable from any machine."
} else {
    Write-Host "  Dropbox exclusions not enabled -- re-run with -IgnoreRegenerable"
}

# --- Prove git can write ------------------------------------------------------
Write-Host ""
& git -C $repo status --porcelain | Out-Null
if ($LASTEXITCODE -ne 0) { throw "git status failed" }
& git -C $repo update-index --refresh | Out-Null
Write-Host "git can read and write the index. OK."

if ((& git -C $repo rev-parse --is-shallow-repository) -eq 'true') {
    Write-Warning "This clone is shallow. `git fetch --unshallow` restores full history, but note that gh-pages history is large."
}
