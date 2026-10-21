<#
.SYNOPSIS
  Install the prebuilt AlertText add-in on this machine. Nothing to compile.

.DESCRIPTION
  For every machine other than the one where the macro is edited, and only the
  first time. Copies AlertText.ppam from beside this script into your AddIns
  folder and registers it so PowerPoint loads it at every start. After that,
  every deck shows an "Alert" button at the end of the Home tab (Alt, H, A, L).

  There is no need to run it again: the installed add-in looks in this folder
  whenever a deck opens and replaces itself with anything newer it finds, so a
  machine that syncs the course folder stays current by itself. Run it again
  only if something has gone wrong.

  Needs no PowerPoint settings: the add-in is already compiled, so the VBA
  trust setting that update-macro-on-change.ps1 requires does not apply here.
  PowerPoint must be closed, because a loaded add-in's file is locked.

  The window stays open at the end so the result is readable when the script
  is run from Explorer (right-click > Run with PowerShell).

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File macros\deploy-to-new-machine.ps1
#>
[CmdletBinding()]
param(
    [string]$InstallFolder = (Join-Path $env:APPDATA 'Microsoft\AddIns')
)

$ErrorActionPreference = 'Stop'
# $PSScriptRoot is not always set while parameter defaults are bound, so resolve them here.
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$names = @('AlertText')

function Finish([int]$Code) {
    Write-Host ''
    if ($Code -eq 0) { Write-Host 'Done.' -ForegroundColor Green } else { Write-Host 'Nothing was installed.' -ForegroundColor Yellow }
    # Read-Host throws when there is no console to read from (-NonInteractive).
    try { if ($Host.Name -eq 'ConsoleHost') { Read-Host 'Press Enter to close' | Out-Null } } catch { }
    exit $Code
}

function Fail([string]$Message) {
    Write-Host ''
    Write-Host $Message -ForegroundColor Red
    Finish 1
}

# PowerPoint rejects automation calls while it is busy with its own startup,
# with "Call was rejected by callee". Asking again a moment later works.
function Use-PowerPoint([scriptblock]$what, [int]$tries = 25) {
    for ($i = 1; $i -le $tries; $i++) {
        try { return & $what }
        catch { if ($i -eq $tries) { throw }; Start-Sleep -Milliseconds 700 }
    }
}

# A PowerPoint to work through. New-Object is the usual way and it sometimes
# fails on this machine with "Server execution failed" -- Office is a packaged
# (Store) install, and after PowerPoint has been killed rather than closed its
# COM activation stays broken until PowerPoint has been started normally once.
# Starting it the ordinary way and attaching to that always works.
function Get-PowerPoint {
    try { return New-Object -ComObject PowerPoint.Application } catch { }
    Write-Host '  PowerPoint refused to start through COM; starting it the ordinary way'
    Start-Process 'powerpnt.exe'
    for ($i = 0; $i -lt 15; $i++) {
        Start-Sleep -Seconds 2
        try {
            $running = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
            Use-PowerPoint { $running.Presentations.Count } | Out-Null     # wait until it answers
            return $running
        } catch { }
    }
    throw 'PowerPoint would not start'
}

try {
    foreach ($name in $names) {
        $source = Join-Path $here "$name.ppam"
        if (-not (Test-Path $source)) {
            Fail "No prebuilt add-in at $source.`nIt is committed in macros of the kevinlb-slides repository; pull the latest version, or build it on the development machine with update-macro-on-change.ps1."
        }
    }
    if (Get-Process -Name POWERPNT -ErrorAction SilentlyContinue) {
        Fail "PowerPoint is running. Close every PowerPoint window (check the taskbar and the system tray), then run this script again.`nA loaded add-in's file is locked and cannot be replaced while PowerPoint is open."
    }

    New-Item -ItemType Directory -Force $InstallFolder | Out-Null
    foreach ($name in $names) {
        $target = Join-Path $InstallFolder "$name.ppam"
        Copy-Item (Join-Path $here "$name.ppam") $target -Force
        Write-Host "Installed  $target"
    }

    # Registering means asking PowerPoint. Writing the registry entry by hand
    # produces something that looks identical and that PowerPoint ignores: an
    # add-in put there never appears in its add-in list at all.
    $app = Get-PowerPoint
    try {
        foreach ($name in $names) {
            foreach ($existing in @($app.AddIns)) {
                if ($existing.Name -ieq $name) { $app.AddIns.Remove($existing.Name) }
            }
            $added = Use-PowerPoint { $app.AddIns.Add((Join-Path $InstallFolder "$name.ppam")) }
            $added.Loaded = -1
            $added.AutoLoad = -1
            Write-Host "Registered $name to load whenever PowerPoint starts."
        }
    }
    finally {
        try { Use-PowerPoint { $app.Quit() } } catch { }
    }
    Write-Host ''
    Write-Host 'Start PowerPoint: the Home tab ends with an "AI 100" group holding the Alert button (Alt, H, A, L).'
    Write-Host 'Right-click that button > Add to Quick Access Toolbar for an Alt+digit shortcut.'
    Write-Host ''
    Write-Host 'If PowerPoint shows a security notice about an add-in the first time, choose Enable.'
    Finish 0
}
catch {
    Fail ("Unexpected error: " + $_.Exception.Message + "`nAt: " + $_.InvocationInfo.PositionMessage)
}
