<#
.SYNOPSIS
  Install the course fonts, for this session as well as every later one.

.DESCRIPTION
  The lecture decks and templates use Inter, Hanken Grotesk and Libre Caslon
  Display, several weights each as separate static files. Windows and Office
  ship none of them, and without them PowerPoint quietly draws a stand-in.

  Installs machine-wide, into the Windows font directory and HKLM, asking for
  Administrator rights itself when it does not have them. -PerUser installs
  for the current user instead and needs no rights.

  Installing a font takes two steps, and until 25 September 2026 this script
  did only the first. Copying the file into the font directory and listing it
  in the registry is what makes Windows load it at the next sign-in; handing
  it to AddFontResource and telling every window (WM_FONTCHANGE) is what makes
  it usable now. Without the second step the fonts stayed invisible until the
  next sign-out -- "Restart PowerPoint" was not enough -- while right-clicking
  the files and choosing Install, which does both, worked at once. The same
  missing step is the likeliest reason a per-user install was once recorded as
  invisible to Office: fonts installed per user from Explorer render fine.

  Close PowerPoint first; it reads the font list at startup.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File fonts\install-fonts.ps1
#>
[CmdletBinding()]
param(
    [switch]$Force,
    [switch]$PerUser
)

$ErrorActionPreference = 'Stop'
$source = $PSScriptRoot

function Finish([int]$Code) {
    # Keep the window open when launched by "Run with PowerShell" or by the
    # elevation below, so the result can be read; harmless in a terminal.
    try { if ($Host.Name -eq 'ConsoleHost') { Read-Host 'Press Enter to close' | Out-Null } } catch { }
    exit $Code
}

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$elevated = (New-Object Security.Principal.WindowsPrincipal($identity)).IsInRole(
    [Security.Principal.WindowsBuiltinRole]::Administrator)

if (-not $PerUser -and -not $elevated) {
    # Refusing and printing a command used to be all this did, in a window that
    # "Run with PowerShell" closes at once -- so nothing was installed and
    # nothing said so. Ask for the rights instead.
    Write-Host 'Asking Windows for Administrator rights to install the fonts for every user...'
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
    if ($Force) { $arguments += '-Force' }
    try {
        Start-Process powershell -Verb RunAs -ArgumentList $arguments
    } catch {
        Write-Warning "Administrator rights were not granted, so nothing was installed.`nRun this again and accept the prompt, or pass -PerUser to install for your account only."
        Finish 1
    }
    exit 0
}

if ($PerUser) {
    $target = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts'
    $regKey = 'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'
    Write-Host 'Installing for the current user.'
} else {
    $target = Join-Path $env:WINDIR 'Fonts'
    $regKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'
    Write-Host 'Installing machine-wide.'
}

New-Item -ItemType Directory -Force -Path $target | Out-Null
if (-not (Test-Path $regKey)) { New-Item -Path $regKey -Force | Out-Null }

Add-Type -AssemblyName PresentationCore
Add-Type -Namespace AI100 -Name FontTable -MemberDefinition @'
[DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
public static extern int AddFontResourceW(string fileName);
[DllImport("user32.dll", CharSet = CharSet.Unicode)]
public static extern IntPtr SendMessageTimeoutW(IntPtr hWnd, uint message, UIntPtr wParam, IntPtr lParam,
    uint flags, uint timeout, out UIntPtr result);
'@

$installed = 0
$skipped = 0
$loaded = 0

try {
    foreach ($file in Get-ChildItem -Path $source -Filter *.ttf) {
        $dest = Join-Path $target $file.Name

        $glyphs = [Windows.Media.GlyphTypeface]::new([Uri]$file.FullName)
        $family = $glyphs.Win32FamilyNames[[System.Globalization.CultureInfo]::InvariantCulture]
        if (-not $family) { $family = ($glyphs.FamilyNames.Values | Select-Object -First 1) }

        # Compose the face from the font's own weight and style rather than trusting
        # Win32FaceNames, which reports "Regular" for every file in this set -- so
        # all four faces of a family would register under one name and overwrite
        # each other, leaving bold and italic unavailable.
        $parts = @()
        if ($glyphs.Weight.ToString() -notin @('Normal', 'Regular')) { $parts += $glyphs.Weight.ToString() }
        if ($glyphs.Style.ToString() -ne 'Normal') { $parts += 'Italic' }
        $face = ($parts -join ' ')
        $name = if ($face) { "$family $face (TrueType)" } else { "$family (TrueType)" }

        $existing = (Get-ItemProperty -Path $regKey -Name $name -ErrorAction SilentlyContinue).$name
        if ($existing -and (Test-Path $dest) -and -not $Force) {
            Write-Host "  already installed: $name"
            $skipped++
        } else {
            Copy-Item $file.FullName $dest -Force
            # Machine-wide entries take a bare filename; per-user entries need the path.
            $value = if ($PerUser) { $dest } else { $file.Name }
            New-ItemProperty -Path $regKey -Name $name -Value $value -PropertyType String -Force | Out-Null
            Write-Host "  installed: $name"
            $installed++
        }

        # Load it into this session too, installed now or earlier: a font a
        # previous run of this script put in place may never have been loaded.
        # Loading one that is already loaded only adds a reference.
        if ([AI100.FontTable]::AddFontResourceW($dest) -gt 0) { $loaded++ }
        else { Write-Warning "  Windows would not load $dest" }
    }
} catch {
    Write-Warning ("Stopped: " + $_.Exception.Message)
    Finish 1
}

# Tell every open window the font list has changed (HWND_BROADCAST,
# WM_FONTCHANGE, SMTO_ABORTIFHUNG so one hung window cannot stall this).
$ignored = [UIntPtr]::Zero
[void][AI100.FontTable]::SendMessageTimeoutW([IntPtr]0xffff, 0x001D, [UIntPtr]::Zero, [IntPtr]::Zero, 0x0002, 1000, [ref]$ignored)

Write-Host ""
Write-Host "Installed $installed font file(s), $skipped already there; $loaded loaded for this session."
Write-Host "Start PowerPoint again (and any other open Office app) to pick them up. No sign-out needed."
Finish 0
