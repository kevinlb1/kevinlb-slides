<#
.SYNOPSIS
  Rebuild and reinstall the AlertText add-in after editing the VBA beside this
  script.

.DESCRIPTION
  For the machine where the macros are developed. The sources of truth are the
  .bas and .cls files beside this script; customUI14.xml beside it is the
  ribbon group the add-in shows on the Home tab. This script:

    1. compiles the module into a PowerPoint add-in (through PowerPoint),
    2. embeds the ribbon XML in the package,
    3. installs it to your AddIns folder and registers it to load at every
       start, replacing the previous copy,
    4. leaves the finished AlertText.ppam beside this script, which is both
       the copy to commit and the copy other machines update themselves from:
       every installation looks in this folder when a deck opens and replaces
       itself, as PowerPoint closes, with anything newer it finds here.

  Requirements, checked up front with instructions if they fail:
    - PowerPoint closed (a loaded add-in's file is locked);
    - "Trust access to the VBA project object model" ticked, once, in
      PowerPoint's Trust Center. Without it PowerPoint refuses scripted access
      to VBA and nothing can be compiled.

  The window stays open at the end so the result is readable when the script
  is run from Explorer (right-click > Run with PowerShell).

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File macros\update-macro-on-change.ps1
#>
[CmdletBinding()]
param(
    [string]$InstallFolder = (Join-Path $env:APPDATA 'Microsoft\AddIns')
)

$ErrorActionPreference = 'Stop'
# $PSScriptRoot is not always set while parameter defaults are bound, so resolve them here.
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$ppSaveAsOpenXMLAddin = 30
$ppSaveAsOpenXMLPresentationMacroEnabled = 25

# The add-in is one module, the class modules it needs, and a ribbon. It stays
# a list because everything below is a loop over it, and because a second
# add-in has been in here once already.
$addIns = @(
    [pscustomobject]@{ Name = 'AlertText'; Module = 'AlertText.bas'; Classes = @('AlertEvents.cls'); Ribbon = 'customUI14.xml' }
)

$TrustHelp = @'
PowerPoint is not allowing scripts to work with VBA projects, so the macros
cannot be compiled into add-ins. Turn the setting on once (it is per user,
and it can be turned off again afterwards):

  1. Open PowerPoint.
  2. File > Options > Trust Center > Trust Center Settings...
  3. Macro Settings > tick "Trust access to the VBA project object model".
  4. OK, OK, then close PowerPoint completely.
  5. Run this script again.

Nothing was changed on this machine.
'@

function Finish([int]$Code) {
    Write-Host ''
    if ($Code -eq 0) { Write-Host 'Done.' -ForegroundColor Green } else { Write-Host 'Nothing was installed.' -ForegroundColor Yellow }
    # Keep the window open when launched from Explorer; harmless in a terminal.
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
    # ---- Preconditions ------------------------------------------------------
    foreach ($a in $addIns) {
        foreach ($f in @($a.Module) + $a.Classes + @($a.Ribbon | Where-Object { $_ })) {
            if (-not (Test-Path (Join-Path $here $f))) { Fail "Missing file: $f`nThis script expects the macro sources beside it, in macros of the kevinlb-slides repository." }
        }
    }
    $sec = Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security' -ErrorAction SilentlyContinue
    if (-not $sec -or $sec.AccessVBOM -ne 1) { Fail $TrustHelp }
    if (Get-Process -Name POWERPNT -ErrorAction SilentlyContinue) {
        Fail "PowerPoint is running. Close every PowerPoint window (check the taskbar and the system tray), then run this script again.`nA loaded add-in's file is locked and cannot be replaced while PowerPoint is open."
    }

    # ---- 1. Compile the module into an add-in ------------------------------
    $builds = @{}
    $app = Get-PowerPoint
    try {
        foreach ($a in $addIns) {
            foreach ($addin in @($app.AddIns)) {
                if ($addin.Name -ieq $a.Name) { $addin.Loaded = 0; $app.AddIns.Remove($addin.Name) }
            }
        }
        foreach ($a in $addIns) {
            Write-Host "Compiling $($a.Module) ..."
            $build = Join-Path $env:TEMP ("{0}-build.ppam" -f $a.Name)
            $viaPptm = [IO.Path]::ChangeExtension($build, '.pptm')
            foreach ($stale in $build, $viaPptm) { if (Test-Path $stale) { Remove-Item $stale -Force } }
            $pres = Use-PowerPoint { $app.Presentations.Add(0) }
            try {
                $project = $pres.VBProject
                if ($null -eq $project) { throw 'VBProject unavailable' }
                $module = $project.VBComponents.Import((Join-Path $here $a.Module))
                # The event sink that watches the selection lives in a class
                # module, because WithEvents needs one.
                foreach ($cls in $a.Classes) {
                    $extra = $project.VBComponents.Import((Join-Path $here $cls))
                    Write-Host ("  imported class {0}" -f $extra.Name)
                }
            }
            catch {
                # The registry said yes but PowerPoint still refused: same fix.
                $pres.Close()
                Fail ($TrustHelp + "`n(PowerPoint reported: $($_.Exception.Message))")
            }
            Write-Host ("  imported module {0}, {1} lines" -f $module.Name, $module.CodeModule.CountOfLines)
            # Saving straight to .ppam is the short way, and on 18 September 2026
            # this machine's PowerPoint began refusing it -- E_FAIL for any
            # presentation, VBA or not, in any folder, while .pptm saved fine. So
            # take the presentation form when the add-in form is refused and change
            # the package's type afterwards; the two differ in the content type of
            # the presentation part and nothing else.
            try {
                $pres.SaveAs($build, $ppSaveAsOpenXMLAddin)
            }
            catch {
                Write-Host '  PowerPoint refused to save an add-in; going through a .pptm instead'
                $pres.SaveAs($viaPptm, $ppSaveAsOpenXMLPresentationMacroEnabled)
            }
            $pres.Close()
            $builds[$a.Name] = $build
        }
    }
    finally {
        try { Use-PowerPoint { $app.Quit() } } catch { }
    }

    Add-Type -AssemblyName WindowsBase
    foreach ($a in $addIns) {
        $build = $builds[$a.Name]

        # ---- 1b. Turn the .pptm into a .ppam, if that is the road we took ---
        $viaPptm = [IO.Path]::ChangeExtension($build, '.pptm')
        if (-not (Test-Path $build) -and (Test-Path $viaPptm)) {
            $presentationType = 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml'
            $addinType = 'application/vnd.ms-powerpoint.addin.macroEnabled.main+xml'
            $src = [System.IO.Packaging.Package]::Open($viaPptm, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
            $dst = [System.IO.Packaging.Package]::Open($build, [System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite)
            try {
                foreach ($part in $src.GetParts()) {
                    $type = $part.ContentType
                    if ($type -eq $presentationType) { $type = $addinType }
                    $copy = $dst.CreatePart($part.Uri, $type, [System.IO.Packaging.CompressionOption]::Normal)
                    $from = $part.GetStream()
                    $to = $copy.GetStream()
                    $from.CopyTo($to)
                    $from.Dispose(); $to.Dispose()
                }
                # The relationship parts come along with everything else, so the
                # relationships are already in place; creating them again is an
                # id conflict.
            }
            finally {
                $src.Close(); $dst.Close()
                Remove-Item $viaPptm -Force -ErrorAction SilentlyContinue
            }
            Write-Host ("  converted the .pptm for {0} into an add-in" -f $a.Name)
        }

        # ---- 2. Embed the ribbon XML in the package -------------------------
        # A .ppam is an Open Packaging Conventions zip. The ribbon is a part at
        # customUI/customUI14.xml, reached from the package root through a
        # relationship of the customui type. Write it with System.IO.Packaging, the
        # same OPC library Office uses: a zip rewritten with System.IO.Compression
        # in update mode is structurally different and PowerPoint silently refuses
        # to load the add-in ("couldn't load the ... add-in").
        if ($a.Ribbon) {
            $pkg = [System.IO.Packaging.Package]::Open($build, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
            try {
                $uri = New-Object System.Uri('/customUI/customUI14.xml', [System.UriKind]::Relative)
                if ($pkg.PartExists($uri)) { $pkg.DeletePart($uri) }
                $part = $pkg.CreatePart($uri, 'application/xml', [System.IO.Packaging.CompressionOption]::Normal)
                $bytes = [System.IO.File]::ReadAllBytes((Join-Path $here $a.Ribbon))
                $stream = $part.GetStream()
                $stream.Write($bytes, 0, $bytes.Length)
                $stream.Dispose()
                $ribbonType = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
                $already = @($pkg.GetRelationships() | Where-Object { $_.RelationshipType -eq $ribbonType })
                if ($already.Count -eq 0) {
                    $pkg.CreateRelationship($uri, [System.IO.Packaging.TargetMode]::Internal, $ribbonType, 'rIdRibbon') | Out-Null
                }
            }
            finally {
                $pkg.Close()
            }
            Write-Host ("  embedded the ribbon in {0}" -f $a.Name)
        }

        # ---- 3. Install ------------------------------------------------------
        $target = Join-Path $InstallFolder ("{0}.ppam" -f $a.Name)
        $beside = Join-Path $here ("{0}.ppam" -f $a.Name)
        New-Item -ItemType Directory -Force $InstallFolder | Out-Null
        Copy-Item $build $target -Force
        Copy-Item $build $beside -Force
        Remove-Item $build -Force
        Write-Host "Installed  $target"
        Write-Host "Beside me  $beside"
        Write-Host ''
    }

    # ---- 4. Register, through PowerPoint ------------------------------------
    # Writing HKCU\...\PowerPoint\AddIns\<name> by hand looks identical to what
    # PowerPoint writes, and PowerPoint ignores it: an add-in registered that
    # way never appeared in Application.AddIns, let alone loaded. Asking
    # PowerPoint to add it produces the same registry entry and does work.
    $app = Get-PowerPoint
    try {
        foreach ($a in $addIns) {
            foreach ($existing in @($app.AddIns)) {
                if ($existing.Name -ieq $a.Name) { $app.AddIns.Remove($existing.Name) }
            }
            $added = Use-PowerPoint { $app.AddIns.Add((Join-Path $InstallFolder ("{0}.ppam" -f $a.Name))) }
            $added.Loaded = -1
            $added.AutoLoad = -1
            Write-Host ("Registered {0}; loaded = {1}" -f $a.Name, $added.Loaded)
        }
    }
    finally {
        try { Use-PowerPoint { $app.Quit() } } catch { }
    }
    Write-Host ''

    Write-Host 'Commit AlertText.ppam together with the sources: every other machine'
    Write-Host 'picks it up from this folder by itself, and deploy-to-new-machine.ps1'
    Write-Host 'installs it the first time.'
    Write-Host ''
    Write-Host 'Start PowerPoint: the Home tab ends with an "AI 100" group holding the Alert button (Alt, H, A, L).'
    Write-Host 'Right-click that button > Add to Quick Access Toolbar for an Alt+digit shortcut.'
    Finish 0
}
catch {
    Fail ("Unexpected error: " + $_.Exception.Message + "`nAt: " + $_.InvocationInfo.PositionMessage)
}
