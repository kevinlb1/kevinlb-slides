# Installing the AlertText add-in on a Mac

The two PowerShell scripts beside this file are for Windows. On a Mac the same
prebuilt add-in, `AlertText.ppam`, installs by hand in a few minutes. Nothing is
compiled; the `.bas` source and `update-macro-on-change.ps1` stay on the
Windows machine where the macro is edited.

Requirements: PowerPoint for Mac 2016 or later (Microsoft 365 is fine). The
macro uses only the ordinary text and ribbon object model, with no Windows
system calls, so it runs unchanged. These steps have not yet been tried on a
Mac by the course team; if something below does not match what you see, note
it here.

## Install

1. **Get `AlertText.ppam`.** Clone the repository, or copy the file out of
   `macros` on any machine that has it. Pull first so the add-in is
   the latest committed build.

2. **Copy it into PowerPoint's add-in folder.** In Finder, choose Go > Go to
   Folder and paste:

       ~/Library/Group Containers/UBF8T346G9.Office/User Content/Add-Ins

   (Finder may show it as "Add-Ins.localized"; that is the same folder.) Drop
   `AlertText.ppam` there. Any folder works, but PowerPoint lists this one
   first in the next step.

3. **Register it.** Open PowerPoint, then Tools > PowerPoint Add-ins. Click
   the + button, choose `AlertText.ppam`, and make sure its checkbox is
   ticked. Registered add-ins load every time PowerPoint starts.

4. **Allow the macros.** If PowerPoint asks whether to enable macros in the
   add-in, choose Enable Macros. This is asked once for the file.

5. **Check the Home tab.** It should end with an "AI 100" group holding the
   Alert, Bolder and Smaller buttons, as on Windows. Ribbon buttons carried by
   an add-in have been supported in PowerPoint for Mac since 2016. If the
   group is missing, see the fallback below.

## Updating

On Windows the add-in replaces itself from the course folder; that machinery
is a Windows command file, so on a Mac the update is still by hand. The same
goes for the missing-font warning, which asks Windows' font mapper: on a Mac
it stays silent, so check Font Book for Inter's weights yourself.

When a pull brings a newer `AlertText.ppam`, quit PowerPoint completely
(PowerPoint > Quit, not just closing the window: a loaded add-in's file is
locked), replace the file in the folder from step 2, and start PowerPoint
again. The registration from step 3 survives; only the file changes. If
PowerPoint reports that it could not load the add-in after an update, quit and
start it once more: a failed load is remembered for the rest of the session.

## Keyboard

The Mac ribbon has no Alt keytips, so Alt, H, A, L does not exist here. The
Quick Access Toolbar route used on Windows is not available either: the Mac
version has no "Add to Quick Access Toolbar" for add-in buttons. The buttons
are mouse-driven on a Mac.

## Fallback if the buttons do not appear

The macros can be run without the ribbon, per deck and per session:

1. Tools > Macro > Visual Basic Editor.
2. In the editor, File > Import File, and choose `AlertText.bas` from
   `macros`. It appears under the open deck's Modules.
3. Back in PowerPoint, select text, then Tools > Macro > Macros, pick
   `ToggleAlert`, `ToggleBolder` or `StepSmaller`, and Run.

The decks are `.pptx` files, which cannot hold macros, so the module is
silently dropped when the deck is saved. That is intended: the macro must
never be embedded in a deck. Re-import it when you next need it. Do not save
a deck as `.pptm` to keep it.

## Security settings

The default Mac setting, "Disable all macros with notification", is enough;
it produces the one-time prompt in step 4. If PowerPoint > Preferences >
Security has "Disable all macros without notification" selected, the add-in
loads without its macros and the buttons do nothing: choose the "with
notification" option instead.
