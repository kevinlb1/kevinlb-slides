Attribute VB_Name = "AlertText"
' Alert text: a heavier weight of the current font family, in a theme colour.
' Never a synthesised bold.
'
' The course templates carry weight in the family name ("Inter Light",
' "Inter SemiBold"), because PowerPoint only knows Regular/Bold/Italic inside
' a family. Each outline level has its own base weight (Regular, Light,
' ExtraLight). Two alert styles sit above it:
'
'   Alert 1: in Accent 1 (the blue). Bold on levels 1 and 2 -- the family's
'            genuine Bold face, reached through the plain family name plus the
'            bold flag -- SemiBold below.
'   Alert 2: in Accent 2 (the teal). SemiBold on levels 1 and 2, Medium below.
'
' An alert below level 1 used to be one weight lighter again (SemiBold, then
' Medium, for alert 1), which with the lighter tint left sub-bullet alerts too
' faint to carry the eye. Levels 2 and 3 each went up a step on 16 September
' 2026, so level 2 now wears level 1's weight and is told apart by its tint and
' size. Keep in step with python/deck_text.py's LOOKS and lectures/STYLE.md.
'
' Both colours come from the deck's theme, so they follow the template. Below
' level 1 the accent is blended toward white by the level, so an alert sits at
' about the lightness of the text around it. The result is written as a plain
' colour: "Accent N, Lighter X%" is a lightness change in HSL, which holds HSL
' saturation and so makes a dark accent *more* vivid as it lightens -- Accent 2
' at level 3 came out a brighter turquoise than the teal it was a faded version
' of -- and transparency would change with whatever sits behind the text. The
' theme is still where the colour comes from: it is read at the moment the
' button is pressed. Keep the ladder in step with python/deck_text.py's BLEND.
'
' The Alert button cycles: default -> alert 1 -> alert 2 -> default. Text
' that is exactly alert 1 for its level (family, bold state, colour) becomes
' alert 2; text that is exactly alert 2 goes back to the level's default;
' anything else, including text that is merely different from the default,
' becomes alert 1.
'
' Two more buttons in the same group:
'
'   Bolder:  alert 2's weight for the level, colour untouched -- emphasis
'            without an alert (SemiBold on levels 1 and 2, Medium below).
'            Pressed on text already at that weight, it goes back.
'   Smaller: one step down on each of three ladders, judged from the text as
'            it is now, not from its indent: the next smaller size in the
'            master's body ladder, the next lighter weight, and the next
'            lighter colour step. In place, so the bullet and indent stay.
'            When no smaller size is left, in an outline placeholder it
'            returns to the paragraph's own look; elsewhere it stops.
'
' "Regular" means the paragraph's level look inside an outline placeholder,
' and the plain family at full strength anywhere else (tables, text boxes).
' With nothing selected, every button acts on the word under the cursor,
' as Ctrl+B does.
'
' Install once per authoring machine, then it works in every deck:
'   update-macro-on-change.ps1 beside this file compiles it into AlertText.ppam and
'   loads it (needs "Trust access to the VBA project object model" ticked
'   once). Re-run it after editing this file. By hand instead:
'   1. Alt+F11, File > Import File..., pick this .bas file.
'   2. File > Save As > PowerPoint Add-in (*.ppam), e.g. AlertText.ppam.
'   3. File > Options > Add-ins > Manage: PowerPoint Add-ins > Go > Add New.
'   Other machines: deploy-to-new-machine.ps1 copies the prebuilt AlertText.ppam
'   from this folder and registers it; nothing to compile there.
' The add-in puts the three buttons in an "AI 100" group at the end of the Home
' tab (customUI14.xml): Alert (Alt, H, A, L), Bolder (Alt, H, B, L) and
' Smaller (Alt, H, S, M). Add-in macros never appear in the Quick Access
' Toolbar's Macros list, so for Alt+digit shortcuts right-click a button and
' choose "Add to Quick Access Toolbar"; AutoHotkey can remap keys onto those.
Option Explicit

' ---- Watching the selection ------------------------------------------------
'
' Promoting or demoting a bullet leaves its alert at the old level's colour,
' and PowerPoint raises no event for either. What it does raise is a change of
' selection, so the add-in keeps the last shape the cursor was in and checks it
' over when the cursor moves somewhere else. Only the colour of alerted text is
' touched -- its colour and its weight -- and only where they belong to
' another level: nothing else is written, so a shape that is already right is
' left alone entirely.
'
' "Alerted text" is text in one of the alert colours. PowerPoint has no way to
' tag a run -- Tags belong to shapes, slides and presentations, not to runs --
' so the colour is the mark, which is the same thing the Alert button reads.
' Two kinds of colour count: the add-in's own opaque ones, and Accent 1 or 2 as
' a theme colour at any tint. The second is how every alert was written before
' the colours became opaque, and how an alert still arrives when a slide is
' pasted from an older deck; tinted with lumMod/lumOff it matches none of the
' opaque colours, so recognising alerts by exact colour alone passed them over
' and Fix Alerts reported "every alert already matches" with two different
' teals side by side on the same level (Lecture 6, slide 18).
'
' The Watch button on the Home tab turns this off and on, and the setting is
' remembered per machine. Note that VBA's edits do not join PowerPoint's undo
' stack: a correction cannot be taken back with Ctrl+Z, only pressed away.

Private Const STATE_DEFAULT As Long = 0
Private Const STATE_ALERT1 As Long = 1
Private Const STATE_ALERT2 As Long = 2
Private Const DEEPEST_LEVEL As Long = 5       ' the master styles levels 1 to 5
' Maths is set in this face (lectures/STYLE.md rule 12). It is part of the
' text, so a selection can hold it, and Font.Name applies to a whole range at
' once: an alert over "k-nearest neighbours" would read Cambria Math off the k
' and give it to the words, or write Inter over the k. Both are wrong, so the
' buttons take the level's family from the ordinary text and put the maths
' back afterwards. The colour is applied to all of it, k included.
Private Const MATH_FAMILY As String = "Cambria Math"

' ---- Fonts that are not there ------------------------------------------------
'
' A deck set in a font this machine lacks does not look broken: PowerPoint
' draws a stand-in and still shows the missing name in the font box, so the
' slides look nearly right while every line break and every alert weight is
' judged against the wrong letters. The only way to tell is to ask Windows
' which face it actually hands out for a name, which is what FontInstalled
' does: GDI picks a substitute when the face is missing, and GetTextFace
' reports the substitute. Declared for Windows only; on a Mac the check says
' yes to everything rather than failing.
#If Mac Then
#Else
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function CreateFontW Lib "gdi32" (ByVal nHeight As Long, ByVal nWidth As Long, _
    ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, _
    ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, _
    ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, _
    ByVal fdwPitchAndFamily As Long, ByVal lpszFace As LongPtr) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function GetTextFaceW Lib "gdi32" (ByVal hDC As LongPtr, ByVal nCount As Long, _
    ByVal lpFaceName As LongPtr) As Long
#End If

Private events As AlertEvents
Private fontsWarned As String       ' "|name|name|": each missing font is mentioned once a session
Private lastFontCheck As String
Private pendingUpdate As Boolean
Private lastUpdate As String
Private lastSlide As Long
Private lastShape As String
Private working As Boolean
Private colours(1 To 2, 1 To 5) As Long
Private coloursFor As String

Public Sub Auto_Open()
    Set events = New AlertEvents
    Set events.App = Application
    pendingUpdate = True
End Sub

' Look for a newer copy once, when the first deck opens: the folder to look
' in is the one that deck came from.
Public Sub MaybeUpdate()
    If Not pendingUpdate Then Exit Sub
    pendingUpdate = False
    On Error Resume Next
    lastUpdate = ArrangeUpdate()
End Sub

' ---- Staying current -------------------------------------------------------
'
' A newer AlertText.ppam turns up in the macros folder beside the decks
' whenever the build script has been run on another machine and Dropbox has
' caught up. Copying it over the installed one from in here does not work:
' PowerPoint holds a loaded add-in's file open, and VBA cannot unload the
' project while any VBA is running, so both "unload then copy" and "take it
' out of the add-in list then copy" fail with "Permission denied". Nor can a
' second add-in do it on this one's behalf, for the same reason -- which is
' what the AI100Update companion add-in was for, and why it is gone.
'
' So the copy is left to something that is not PowerPoint: a hidden PowerShell
' that waits for PowerPoint to close, replaces the file, and stops. The newer
' version is in place at the next start. Nothing here needs to know where
' Dropbox is, and no machine needs to be told to update anything.

Private Function AddInsFolder() As String
    AddInsFolder = Environ$("APPDATA") & "\Microsoft\AddIns\"
End Function

Private Function SourceNote() As String
    SourceNote = AddInsFolder() & "AI100-source.txt"
End Function

' Where the add-in is published: the macros folder beside the deck that is
' open, or the one remembered from the last time a deck was open.
Private Function SourceFolder() As String
    Dim guess As String, handle As Integer, note As String
    On Error Resume Next
    If Application.Presentations.Count > 0 Then
        guess = Application.ActivePresentation.Path & "\macros"
        If Dir$(guess & "\AlertText.ppam") <> "" Then
            handle = FreeFile
            Open SourceNote() For Output As #handle
            Print #handle, guess
            Close #handle
            SourceFolder = guess
            Exit Function
        End If
    End If
    If Dir$(SourceNote()) = "" Then Exit Function
    handle = FreeFile
    Open SourceNote() For Input As #handle
    Line Input #handle, note
    Close #handle
    SourceFolder = Trim$(note)
End Function

' Leave a command file running that swaps the add-in in once PowerPoint is out
' of the way. It waits rather than copying once, so a PowerPoint started again
' straight away only delays it, and it deletes itself when it is done.
'
' A batch file, not PowerShell: this is a packaged (Store) Office, and a
' powershell.exe started from in here dies immediately, though cmd.exe runs
' perfectly well. Every tool is named by its full path, since the environment
' PowerPoint hands on is not necessarily one where they are on the PATH.
Public Function ArrangeUpdate() As String
    Dim folder As String, src As String, dst As String
    Dim script As String, handle As Integer, tools As String, nap As String
    On Error GoTo failed
    folder = SourceFolder()
    If folder = "" Then
        ArrangeUpdate = "no macros folder to update from yet"
        Exit Function
    End If
    src = folder & "\AlertText.ppam"
    dst = AddInsFolder() & "AlertText.ppam"
    If Dir$(src) = "" Then
        ArrangeUpdate = "no copy of AlertText in " & folder
        Exit Function
    End If
    If Dir$(dst) <> "" Then
        If FileDateTime(src) <= FileDateTime(dst) Then
            ArrangeUpdate = "this is the newest AlertText in " & folder
            Exit Function
        End If
    End If
    tools = Environ$("SystemRoot") & "\System32\"
    nap = tools & "ping.exe -n 3 127.0.0.1 >nul"
    script = Environ$("TEMP") & "\AlertText-update.cmd"
    handle = FreeFile
    Open script For Output As #handle
    Print #handle, "@echo off"
    Print #handle, ":wait"
    Print #handle, tools & "tasklist.exe /fi ""imagename eq POWERPNT.EXE"" /nh | " & tools & "find.exe /i ""POWERPNT.EXE"" >nul"
    Print #handle, "if not errorlevel 1 (" & nap & " & goto wait)"
    Print #handle, "set /a tries=0"
    Print #handle, ":again"
    Print #handle, "copy /y """ & src & """ """ & dst & """ >nul"
    Print #handle, "if not errorlevel 1 goto done"
    Print #handle, "set /a tries+=1"
    Print #handle, "if %tries% lss 30 (" & nap & " & goto again)"
    Print #handle, ":done"
    Print #handle, "del ""%~f0"""
    Close #handle
    Shell "cmd.exe /c """ & script & """", vbHide
    ArrangeUpdate = "a newer AlertText is in " & folder & "; it goes in when PowerPoint closes"
    Exit Function
failed:
    ArrangeUpdate = "could not arrange the update: " & Err.Description
End Function

' What the last look for a newer copy found, for WatchStatus to report.
Public Function UpdateReport() As String
    UpdateReport = lastUpdate
End Function

' The date this add-in was built, for the ribbon's label: the file's own
' timestamp, so nothing has to be bumped by hand.
Public Function BuildDate() As String
    Dim path As String
    On Error Resume Next
    BuildDate = ""
    path = AddInsFolder() & "AlertText.ppam"
    If Dir$(path) <> "" Then BuildDate = Format$(FileDateTime(path), "d mmm")
End Function

Public Sub GroupLabel(control As IRibbonControl, ByRef returnedVal)
    Dim built As String
    built = BuildDate()
    returnedVal = IIf(built = "", "AI 100", "AI 100 " & ChrW$(183) & " " & built)
End Sub

Public Sub Auto_Close()
    Set events = Nothing
End Sub

' Off until the Watch button turns it on: an event handler that rewrites text
' behind you is something to opt into, and VBA's edits cannot be undone with
' Ctrl+Z.
'
' The setting is a file beside the add-in rather than SaveSetting/GetSetting.
' This Office is a packaged install, so its registry is virtualised: the add-in
' writes to a private copy of HKCU and reads nothing anyone else wrote there,
' which made the setting look empty from outside and impossible to preset.
Private Function SettingFile() As String
    SettingFile = Environ$("APPDATA") & "\Microsoft\AddIns\AlertText-watch.txt"
End Function

Private Function WatchOn() As Boolean
    Dim handle As Integer, flag As String
    On Error GoTo none
    If Dir$(SettingFile()) = "" Then GoTo none
    handle = FreeFile
    Open SettingFile() For Input As #handle
    Line Input #handle, flag
    Close #handle
    WatchOn = (Trim$(flag) = "1")
    Exit Function
none:
    WatchOn = False
End Function

Private Sub SetWatch(ByVal enabled As Boolean)
    Dim handle As Integer
    On Error Resume Next
    handle = FreeFile
    Open SettingFile() For Output As #handle
    Print #handle, IIf(enabled, "1", "0")
    Close #handle
End Sub

Public Sub ToggleWatchRibbon(control As IRibbonControl, pressed As Boolean)
    SetWatch pressed
End Sub

Public Sub WatchPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = WatchOn()
End Sub

' For checking from outside that the add-in is awake and what it thinks.
Public Function WatchStatus() As String
    WatchStatus = "hooked=" & IIf(events Is Nothing, "no", "yes") & _
                  " watching=" & IIf(WatchOn(), "yes", "no") & _
                  " file=[" & SettingFile() & "]" & _
                  " last=" & lastSlide & "/" & lastShape & _
                  " update=[" & lastUpdate & "]" & _
                  " fonts=[" & lastFontCheck & "]"
End Function

' ---- Checking the fonts ----------------------------------------------------

' Whether Windows has the face itself, not a stand-in for it.
Public Function FontInstalled(ByVal faceName As String) As Boolean
#If Mac Then
    FontInstalled = True
#Else
    Dim dc As LongPtr, hFont As LongPtr, hOld As LongPtr, buffer As String, cut As Long
    dc = GetDC(0)
    ' height -20, any weight, DEFAULT_CHARSET: let the font mapper choose freely
    hFont = CreateFontW(-20, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, StrPtr(faceName))
    hOld = SelectObject(dc, hFont)
    buffer = String$(64, vbNullChar)
    GetTextFaceW dc, 64, StrPtr(buffer)
    SelectObject dc, hOld
    DeleteObject hFont
    ReleaseDC 0, dc
    cut = InStr(buffer, vbNullChar)
    If cut > 0 Then buffer = Left$(buffer, cut - 1)
    FontInstalled = (StrComp(buffer, faceName, vbTextCompare) = 0)
    ' A name in Windows' substitution table is a stand-in by definition, and
    ' GDI reports it back as asked: "Helvetica" comes back as "Helvetica" and
    ' is drawn in Arial.
    If FontInstalled Then FontInstalled = Not IsFontAlias(faceName)
#End If
End Function

Private Function IsFontAlias(ByVal faceName As String) As Boolean
    Dim target As String
    On Error Resume Next
    target = CreateObject("WScript.Shell").RegRead( _
        "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\FontSubstitutes\" & faceName)
    IsFontAlias = (target <> "")
End Function

' A font name as the master gives it, with a theme reference such as +mn-lt
' turned into the font the theme names.
Private Function RealFontName(ByVal pres As Presentation, ByVal name As String) As String
    On Error Resume Next
    RealFontName = name
    If Left$(name, 4) = "+mn-" Then
        RealFontName = pres.SlideMaster.Theme.ThemeFontScheme.MinorFont(msoThemeLatin).name
    ElseIf Left$(name, 4) = "+mj-" Then
        RealFontName = pres.SlideMaster.Theme.ThemeFontScheme.MajorFont(msoThemeLatin).name
    End If
End Function

' The fonts a deck cannot do without: every weight of the template's families
' that the deck uses, the weights the buttons write whether or not it uses
' them yet, and the maths face. A stray font pasted in from elsewhere is not
' on the list -- Courier New in one code sample should not raise a warning
' every time the deck opens.
Private Function NeededFonts(ByVal pres As Presentation) As Collection
    Dim needed As New Collection, families As String, body As String, title As String
    Dim f As Object, name As String
    On Error Resume Next                       ' Add raises on a duplicate: that is the dedupe
    body = BaseFamily(RealFontName(pres, pres.SlideMaster.TextStyles(ppBodyStyle).Levels(1).Font.name))
    title = BaseFamily(RealFontName(pres, pres.SlideMaster.TextStyles(ppTitleStyle).Levels(1).Font.name))
    families = "|" & body & "|" & title & "|"
    For Each f In pres.Fonts
        name = f.name
        If InStr(families, "|" & BaseFamily(name) & "|") > 0 Then needed.Add name, name
    Next f
    needed.Add body, body
    needed.Add body & " SemiBold", body & " SemiBold"
    needed.Add body & " Medium", body & " Medium"
    needed.Add MATH_FAMILY, MATH_FAMILY
    Set NeededFonts = needed
End Function

' What is wrong with the fonts for the deck in front of you, as the warning
' would say it, or "" when nothing is.
Public Function FontProblems() As String
    On Error Resume Next
    FontProblems = FontProblemsFor(ActivePresentation, False)
End Function

' The same for a given deck; with `fresh`, only what has not already been
' mentioned this session.
Private Function FontProblemsFor(ByVal pres As Presentation, ByVal fresh As Boolean) As String
    Dim needed As Collection, name As Variant
    Dim missing As String
    On Error Resume Next
    If pres Is Nothing Then Exit Function
    Set needed = NeededFonts(pres)
    For Each name In needed
        If Not FontInstalled(CStr(name)) Then
            If Not fresh Or InStr(fontsWarned, "|" & name & "|") = 0 Then missing = missing & vbCr & "    " & name
        End If
    Next name
    If missing <> "" Then
        FontProblemsFor = "This computer does not have fonts these slides are set in:" & vbCr & missing & vbCr & vbCr & _
            "PowerPoint quietly draws a stand-in instead and still shows the missing name, " & _
            "so the slides look nearly right while line breaks and alert weights are judged " & _
            "against the wrong letters."
    End If
    If FontProblemsFor <> "" Then
        FontProblemsFor = FontProblemsFor & vbCr & vbCr & _
            "To fix it: close PowerPoint, then right-click fonts\install-fonts.ps1 in the folder " & _
            "beside the lecture decks and choose Run with PowerShell (it asks for Administrator " & _
            "rights itself), or select the font files there, right-click and choose Install."
    End If
End Function

' Once a deck is open: warn about any font it needs that is not here. Each
' problem is raised once a session, however many decks are opened.
Public Sub CheckFonts(ByVal pres As Presentation)
    Dim message As String, name As Variant
    On Error Resume Next
    message = FontProblemsFor(pres, True)
    lastFontCheck = IIf(FontProblemsFor(pres, False) = "", "all present", "problems")
    If message = "" Then Exit Sub
    For Each name In NeededFonts(pres)
        If Not FontInstalled(CStr(name)) Then fontsWarned = fontsWarned & "|" & name & "|"
    Next name
    MsgBox message, vbExclamation, "AI 100: fonts missing"
End Sub

Public Sub ForgetShape()
    lastSlide = 0
    lastShape = ""
End Sub

' The selection has moved: look over the shape it came from.
'
' Writing to a shape can itself move the selection, so the handler remembers
' where the cursor is now before it touches anything, and refuses to run
' inside itself. Without that, the first live test left PowerPoint wedged.
Public Sub OnSelectionChange(ByVal Sel As Selection)
    Dim slideIndex As Long, shapeName As String
    Dim fromSlide As Long, fromShape As String
    On Error Resume Next
    If working Then Exit Sub
    If Sel.Type = ppSelectionText Or Sel.Type = ppSelectionShapes Then
        slideIndex = Sel.SlideRange(1).SlideIndex
        shapeName = Sel.ShapeRange(1).Name
    End If
    fromSlide = lastSlide
    fromShape = lastShape
    lastSlide = slideIndex
    lastShape = shapeName
    If Not WatchOn() Then Exit Sub
    If fromSlide = 0 Then Exit Sub
    If fromSlide = slideIndex And fromShape = shapeName Then Exit Sub
    working = True
    FixShape ActivePresentation.Slides(fromSlide), fromShape
    working = False
End Sub

Private Sub FixShape(ByVal sld As Slide, ByVal shapeName As String)
    Dim shp As Shape
    On Error Resume Next
    Set shp = sld.Shapes(shapeName)
    If shp Is Nothing Then Exit Sub
    If Not shp.HasTextFrame Then Exit Sub
    If Not shp.TextFrame2.HasText Then Exit Sub
    FixFrame shp.TextFrame2
End Sub

' Every alerted run takes the look of the level its paragraph is at now: the
' level's colour, and the level's weight for whichever alert it is wearing. A
' symbol set as maths keeps its face and takes only the colour. Nothing else
' in the text is touched, and a run that is already right is not written to.
' Returns how many runs were changed.
Public Function FixFrame(ByVal frame As TextFrame2) As Long
    Dim p As Long, r As Long, para As TextRange2, run As TextRange2
    Dim level As Long, which As Long
    Dim wanted As Long, family As String, base As String, bold As Boolean
    Dim touched As Boolean
    On Error Resume Next
    For p = 1 To frame.TextRange.Paragraphs.Count
        Set para = frame.TextRange.Paragraphs(p)
        level = para.ParagraphFormat.IndentLevel
        For r = 1 To para.Runs.Count
            Set run = para.Runs(r)
            which = AlertOfRun(run)
            If which = 0 Then GoTo nextRun
            touched = False
            wanted = AlertRGB(AccentIndex(which), level)
            ' A theme colour is replaced even where it resolves to the right
            ' shade: it is the old kind of alert, and only the plain colour
            ' holds still when the text is promoted or the theme changes.
            If run.Font.Fill.ForeColor.Type <> msoColorTypeRGB _
               Or run.Font.Fill.ForeColor.RGB <> wanted Then
                Paint run, wanted
                touched = True
            End If
            If run.Font.name <> MATH_FAMILY Then
                base = BaseFamily(run.Font.name)
                If base = "" Or base = MATH_FAMILY Then base = MasterFamily(level)
                If which = 1 Then
                    family = WithSuffix(base, Alert1Suffix(level))
                    bold = Alert1IsBold(level)
                Else
                    family = WithSuffix(base, Alert2Suffix(level))
                    bold = False
                End If
                If run.Font.name <> family Then
                    run.Font.name = family
                    touched = True
                End If
                If (run.Font.Bold = msoTrue) <> bold Then
                    If bold Then run.Font.Bold = msoTrue Else run.Font.Bold = msoFalse
                    touched = True
                End If
            End If
            If touched Then FixFrame = FixFrame + 1
nextRun:
        Next r
    Next p
End Function

' Every shape on every slide, for when the watching has been off.
Public Sub FixAlertsRibbon(control As IRibbonControl)
    Dim fixed As Long
    fixed = FixDeck()
    MsgBox IIf(fixed = 0, "Every alert already matches its level.", _
               fixed & " run(s) put back to their level's look."), _
           vbInformation, "AI 100"
End Sub

' Every shape on every slide. Returns how many runs were put right.
Public Function FixDeck() As Long
    Dim sld As Slide, shp As Shape, fixed As Long
    working = True
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame2.HasText Then fixed = fixed + FixFrame(shp.TextFrame2)
            End If
        Next shp
    Next sld
    working = False
    FixDeck = fixed
End Function


' The template's weight ladder for the first three levels. Keep in step with
' the master and with python/adopt_template.py.
Private Function LevelSuffix(ByVal level As Long) As String
    Select Case level
        Case 1: LevelSuffix = ""
        Case 2: LevelSuffix = "Light"
        Case Else: LevelSuffix = "ExtraLight"
    End Select
End Function

' Alert 1: Bold on the first two levels, SemiBold below.
Private Function Alert1Suffix(ByVal level As Long) As String
    Select Case level
        Case 1, 2: Alert1Suffix = ""         ' plain family, with the Bold face
        Case Else: Alert1Suffix = "SemiBold"
    End Select
End Function

' Weight 700 has no family of its own: Regular and Bold share the plain name,
' so level 1 reaches Bold through the flag. That is the real Bold file, not a
' synthesised bold, which only happens when a family has no Bold face.
Private Function Alert1IsBold(ByVal level As Long) As Boolean
    Alert1IsBold = (level <= 2)
End Function

' Alert 2: SemiBold on the first two levels, Medium below, never bold.
Private Function Alert2Suffix(ByVal level As Long) As String
    Select Case level
        Case 1, 2: Alert2Suffix = "SemiBold"
        Case Else: Alert2Suffix = "Medium"
    End Select
End Function

' The family the master gives this level's plain text, with no weight on it.
' The last resort when the selection cannot say what family it is in.
Private Function MasterFamily(ByVal level As Long) As String
    Dim name As String
    On Error GoTo fallback
    If level < 1 Then level = 1
    If level > DEEPEST_LEVEL Then level = DEEPEST_LEVEL
    name = BaseFamily(ActivePresentation.SlideMaster.TextStyles(ppBodyStyle).Levels(level).Font.name)
    If name <> "" Then
        MasterFamily = name
        Exit Function
    End If
fallback:
    MasterFamily = "Inter"
End Function

' The first character that is not maths; 0 when the selection is all maths.
Private Function FirstPlain(ByVal rng As TextRange2) As Long
    Dim i As Long
    For i = 1 To rng.Length
        If rng.Characters(i, 1).Font.Name <> MATH_FAMILY Then
            FirstPlain = i
            Exit Function
        End If
    Next i
    FirstPlain = 0
End Function

' The character the buttons judge the selection by: the first ordinary one, so
' a state and a family are read from the words rather than from an equation.
Private Function PlainChar(ByVal rng As TextRange2) As TextRange2
    Dim i As Long
    i = FirstPlain(rng)
    If i = 0 Then i = 1
    Set PlainChar = rng.Characters(i, 1)
End Function

' Where the maths sits and how it is slanted, so it can be put back after a
' range-wide font change.
Private Function MathSpots(ByVal rng As TextRange2) As Collection
    Dim i As Long, spots As New Collection
    For i = 1 To rng.Length
        If rng.Characters(i, 1).Font.Name = MATH_FAMILY Then
            spots.Add Array(i, rng.Characters(i, 1).Font.Italic)
        End If
    Next i
    Set MathSpots = spots
End Function

' Set the family and weight of everything except the maths, which keeps its
' own face and italics.
Private Sub SetFamily(ByVal rng As TextRange2, ByVal family As String, ByVal bold As Long)
    Dim spots As Collection, spot As Variant, ch As TextRange2
    Set spots = MathSpots(rng)
    rng.Font.Name = family
    rng.Font.Bold = bold
    For Each spot In spots
        Set ch = rng.Characters(CLng(spot(0)), 1)
        ch.Font.Name = MATH_FAMILY
        ch.Font.Italic = spot(1)
    Next spot
End Sub

Private Function WithSuffix(ByVal base As String, ByVal suffix As String) As String
    If suffix = "" Then WithSuffix = base Else WithSuffix = base & " " & suffix
End Function

' The family a level's plain text uses.
Private Function LevelFamily(ByVal base As String, ByVal level As Long) As String
    LevelFamily = WithSuffix(base, LevelSuffix(level))
End Function

' Weight suffixes the course font families use, longest first so that
' "ExtraLight" is matched before "Light".
Private Function WeightSuffixes() As Variant
    WeightSuffixes = Array("ExtraLight", "ExtraBold", "SemiBold", "Thin", "Light", _
                           "Medium", "Bold", "Black")
End Function

Private Function BaseFamily(ByVal fontName As String) As String
    Dim suffix As Variant
    For Each suffix In WeightSuffixes()
        If Len(fontName) > Len(suffix) + 1 Then
            If Right$(fontName, Len(suffix) + 1) = " " & suffix Then
                BaseFamily = Left$(fontName, Len(fontName) - Len(suffix) - 1)
                Exit Function
            End If
        End If
    Next suffix
    BaseFamily = fontName
End Function

' The master's body colour is 002040 at full strength on level 1 and
' progressively more transparent below it.
Private Function LevelTransparency(ByVal level As Long) As Single
    Select Case level
        Case 1: LevelTransparency = 0
        Case 2: LevelTransparency = 0.3
        Case 3: LevelTransparency = 0.45
        Case Else: LevelTransparency = 0.7
    End Select
End Function

' How much of the accent is left at each level, as whole percents. Mixing is
' integer arithmetic so that python/deck_text.py, which writes the same
' colours from outside, lands on the same byte every time.
Private Function BlendFactor(ByVal level As Long) As Long
    Select Case level
        Case 1: BlendFactor = 100
        Case 2: BlendFactor = 85
        Case 3: BlendFactor = 70
        Case Else: BlendFactor = 50
    End Select
End Function

' The colour an alert wears at this level: the theme's accent blended toward
' white. PowerPoint stores an RGB long as &HBBGGRR, which is what RGB() builds.
' The ten colours are cached per presentation: the watcher asks for them once
' per run, and reading the theme every time made it crawl.
Private Function AlertRGB(ByVal accent As Long, ByVal level As Long) As Long
    Dim which As Long
    If level > DEEPEST_LEVEL Then level = DEEPEST_LEVEL
    If level < 1 Then level = 1
    If accent = msoThemeAccent1 Then which = 1 Else which = 2
    If coloursFor <> ActivePresentation.FullName Then FillColours
    AlertRGB = colours(which, level)
End Function

Private Sub FillColours()
    Dim accent As Long, level As Long, base As Long, keep As Long
    For accent = 1 To 2
        base = ActivePresentation.SlideMaster.Theme.ThemeColorScheme( _
                   AccentIndex(accent)).RGB
        For level = 1 To DEEPEST_LEVEL
            keep = BlendFactor(level)
            colours(accent, level) = RGB(Mix(base And &HFF&, keep), _
                                         Mix((base \ &H100&) And &HFF&, keep), _
                                         Mix((base \ &H10000) And &HFF&, keep))
        Next level
    Next accent
    coloursFor = ActivePresentation.FullName
End Sub

Private Function Mix(ByVal channel As Long, ByVal keep As Long) As Long
    Mix = (channel * keep + 255 * (100 - keep) + 50) \ 100
End Function

' Accent 1 and Accent 2's places in the theme colour scheme.
Private Function AccentIndex(ByVal which As Long) As Long
    If which = 1 Then AccentIndex = msoThemeAccent1 Else AccentIndex = msoThemeAccent2
End Function

' The level an alert colour belongs to, 0 when the colour is not an alert.
Private Function AlertLevelOf(ByVal colour As Long, ByRef which As Long) As Long
    Dim accent As Long, level As Long
    For accent = 1 To 2
        For level = 1 To DEEPEST_LEVEL
            If colour = AlertRGB(AccentIndex(accent), level) Then
                which = accent
                AlertLevelOf = level
                Exit Function
            End If
        Next level
    Next accent
    AlertLevelOf = 0
End Function

' Which alert a run is wearing: 1, 2, or 0 for none. See "Alerted text" at the
' top for why a theme colour counts as well as the add-in's own colours.
Private Function AlertOfRun(ByVal run As TextRange2) As Long
    Dim which As Long
    On Error Resume Next
    If AlertLevelOf(run.Font.Fill.ForeColor.RGB, which) > 0 Then
        AlertOfRun = which
        Exit Function
    End If
    If run.Font.Fill.ForeColor.Type = msoColorTypeScheme Then
        Select Case run.Font.Fill.ForeColor.ObjectThemeColor
            Case msoThemeColorAccent1: AlertOfRun = 1
            Case msoThemeColorAccent2: AlertOfRun = 2
        End Select
    End If
End Function

' Which of the three states the selection is exactly in. Anything that is not
' exactly alert 1 or alert 2 for its level counts as default, so the next
' press makes it alert 1.
Private Function CurrentState(ByVal rng As TextRange2, ByVal base As String, ByVal level As Long) As Long
    ' Judged from the first ordinary character: a selection that also holds an
    ' equation reports no font name of its own, which would read as "not an
    ' alert" and stop the button cycling.
    Dim probe As TextRange2, isBold As Boolean
    Set probe = PlainChar(rng)
    isBold = (probe.Font.Bold = msoTrue)
    CurrentState = STATE_DEFAULT
    If probe.Font.Name = WithSuffix(base, Alert1Suffix(level)) _
       And isBold = Alert1IsBold(level) _
       And probe.Font.Fill.ForeColor.RGB = AlertRGB(msoThemeAccent1, level) Then
        CurrentState = STATE_ALERT1
    ElseIf probe.Font.Name = WithSuffix(base, Alert2Suffix(level)) _
       And Not isBold _
       And probe.Font.Fill.ForeColor.RGB = AlertRGB(msoThemeAccent2, level) Then
        CurrentState = STATE_ALERT2
    End If
End Function

Private Sub ApplyDefault(ByVal rng As TextRange2, ByVal base As String, ByVal level As Long)
    SetFamily rng, LevelFamily(base, level), msoFalse
    With rng.Font.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = RGB(0, 32, 64)
        .Transparency = LevelTransparency(level)
    End With
End Sub

Private Sub ApplyAlert1(ByVal rng As TextRange2, ByVal base As String, ByVal level As Long)
    Dim bold As Long
    If Alert1IsBold(level) Then bold = msoTrue Else bold = msoFalse
    SetFamily rng, WithSuffix(base, Alert1Suffix(level)), bold
    Paint rng, AlertRGB(msoThemeAccent1, level)
End Sub

' One colour, opaque, over whatever is behind the text.
Private Sub Paint(ByVal rng As TextRange2, ByVal colour As Long)
    With rng.Font.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = colour
        .Transparency = 0
    End With
End Sub

Private Sub ApplyAlert2(ByVal rng As TextRange2, ByVal base As String, ByVal level As Long)
    SetFamily rng, WithSuffix(base, Alert2Suffix(level)), msoFalse
    Paint rng, AlertRGB(msoThemeAccent2, level)
End Sub

' The text the buttons act on: the selection, or the word under the cursor.
' Returns False when nothing textual is selected.
Private Function Selected(ByRef rng As TextRange2, ByRef base As String, ByRef level As Long) As Boolean
    Selected = False
    If ActiveWindow.Selection.Type <> ppSelectionText Then Exit Function
    Set rng = ActiveWindow.Selection.TextRange2
    If rng.Length = 0 Then Set rng = WordAt(rng)
    If rng Is Nothing Then Exit Function
    If rng.Length = 0 Then Exit Function
    base = BaseFamily(rng.Font.Name)
    If base = MATH_FAMILY Then base = ""
    ' A selection spanning two weights reports no name of its own; go by its
    ' first ordinary character, never by an equation's Cambria Math.
    If base = "" Then base = BaseFamily(PlainChar(rng).Font.Name)
    If base = MATH_FAMILY Or base = "" Then base = MasterFamily(rng.ParagraphFormat.IndentLevel)
    level = rng.ParagraphFormat.IndentLevel
    Selected = True
End Function

' The word around an insertion point, without its trailing space, so a click
' with nothing selected behaves like Ctrl+B.
Private Function WordAt(ByVal point As TextRange2) As TextRange2
    Dim whole As TextRange2
    On Error Resume Next
    Set whole = point.Parent.TextRange          ' the frame's full text
    On Error GoTo 0
    If whole Is Nothing Then Exit Function
    Dim pos As Long, w As TextRange2, i As Long, n As Long
    pos = point.Start
    For i = 1 To whole.Words.Count
        Set w = whole.Words(i)
        If pos >= w.Start And pos <= w.Start + w.Length Then
            n = Len(RTrim$(w.Text))
            If n = 0 Then Exit Function
            Set WordAt = whole.Characters(w.Start, n)
            Exit Function
        End If
    Next i
End Function

' True when the text sits in an outline placeholder, where the paragraph's
' level defines its regular look. Tables and text boxes have no levels, with
' one exception: a placeholder PowerPoint has detached from the layout (the
' layout's own is taken by a table, as on Lecture 4 slide 6) keeps the name
' "Content Placeholder N" and a copy of the master's level styles, so its text
' is an outline in everything but the object model's answer. Without this the
' buttons wrote the level-1 look on every line of such a box.
' python/deck_text.py and python/tidy_slide_text.py judge it the same way.
Private Function InOutline() As Boolean
    InOutline = False
    On Error GoTo done
    Dim shp As Shape
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If shp.Type <> msoPlaceholder Then
        InOutline = (shp.Name Like "Content Placeholder*")
        Exit Function
    End If
    Select Case shp.PlaceholderFormat.Type
        Case ppPlaceholderBody, ppPlaceholderObject, ppPlaceholderVerticalBody, ppPlaceholderSubtitle
            InOutline = True
    End Select
done:
End Function

' The level whose look counts as "regular" for the selection.
Private Function ReferenceLevel(ByVal level As Long) As Long
    If InOutline() Then ReferenceLevel = level Else ReferenceLevel = 1
End Function

' Ribbon entry points. customUI14.xml names these as onAction; Office passes
' the clicked control.
Public Sub ToggleAlertRibbon(control As IRibbonControl)
    ToggleAlert
End Sub

Public Sub ToggleBolderRibbon(control As IRibbonControl)
    ToggleBolder
End Sub

Public Sub StepSmallerRibbon(control As IRibbonControl)
    StepSmaller
End Sub

Public Sub ToggleAlert()
    Dim rng As TextRange2, base As String, level As Long
    If Not Selected(rng, base, level) Then Exit Sub
    level = ReferenceLevel(level)

    Select Case CurrentState(rng, base, level)
        Case STATE_ALERT1: ApplyAlert2 rng, base, level
        Case STATE_ALERT2: ApplyDefault rng, base, level
        Case Else: ApplyAlert1 rng, base, level
    End Select
End Sub

' ---- Bolder: two weights up from regular, colour untouched ------------------
Public Sub ToggleBolder()
    Dim rng As TextRange2, base As String, level As Long
    If Not Selected(rng, base, level) Then Exit Sub
    level = ReferenceLevel(level)
    Dim bolder As String, probe As TextRange2
    bolder = WithSuffix(base, Alert2Suffix(level))     ' alert 2's weight for the level
    Set probe = PlainChar(rng)
    If probe.Font.Name = bolder And probe.Font.Bold <> msoTrue Then
        SetFamily rng, LevelFamily(base, level), msoFalse
    Else
        SetFamily rng, bolder, msoFalse
    End If
End Sub

' ---- Smaller: one step down each ladder, judged from the text itself --------

' Body text sizes come from the deck's master, so the ladder follows the
' template rather than this file. Levels 1 to 5 are styled.
Private Function MasterLevelSize(ByVal level As Long) As Single
    MasterLevelSize = ActivePresentation.SlideMaster.TextStyles(ppBodyStyle).Levels(level).Font.Size
End Function

' The largest master size clearly smaller than the current one; 0 when none.
Private Function NextSmallerSize(ByVal current As Single) As Single
    Dim n As Long, candidate As Single
    NextSmallerSize = 0
    For n = 1 To 5
        candidate = MasterLevelSize(n)
        If candidate < current - 0.5 And candidate > NextSmallerSize Then NextSmallerSize = candidate
    Next n
End Function

' The weight suffix the text wears now, with the bold flag read as "Bold".
Private Function CurrentSuffix(ByVal rng As TextRange2, ByVal base As String) As String
    Dim name As String
    name = PlainChar(rng).Font.Name
    If PlainChar(rng).Font.Bold = msoTrue Then
        CurrentSuffix = "Bold"
    ElseIf name = base Then
        CurrentSuffix = ""
    ElseIf Left$(name, Len(base) + 1) = base & " " Then
        CurrentSuffix = Mid$(name, Len(base) + 2)
    Else
        CurrentSuffix = ""
    End If
End Function

' One step lighter on the weight ladder; ExtraLight is the floor (Thin is not
' installed).
Private Function LighterSuffix(ByVal suffix As String) As String
    Select Case suffix
        Case "Black": LighterSuffix = "ExtraBold"
        Case "ExtraBold": LighterSuffix = "Bold"
        Case "Bold": LighterSuffix = "SemiBold"
        Case "SemiBold": LighterSuffix = "Medium"
        Case "Medium": LighterSuffix = ""
        Case "": LighterSuffix = "Light"
        Case "Light": LighterSuffix = "ExtraLight"
        Case Else: LighterSuffix = "ExtraLight"
    End Select
End Function

Private Sub ApplyWeight(ByVal rng As TextRange2, ByVal base As String, ByVal suffix As String)
    If suffix = "Bold" Then
        SetFamily rng, base, msoTrue
    Else
        SetFamily rng, WithSuffix(base, suffix), msoFalse
    End If
End Sub

' The next value on a ladder above the current one; the last value is a floor.
Private Function NextStep(ByVal current As Single, ByVal ladder As Variant) As Single
    Dim i As Long
    For i = LBound(ladder) To UBound(ladder)
        If ladder(i) > current + 0.01 Then
            NextStep = ladder(i)
            Exit Function
        End If
    Next i
    NextStep = ladder(UBound(ladder))
End Function

' One step lighter in colour: an alert takes the next level's alert colour,
' anything else the next transparency on the ink ladder.
Private Sub StepColourLighter(ByVal rng As TextRange2)
    Dim which As Long, level As Long
    level = AlertLevelOf(rng.Font.Fill.ForeColor.RGB, which)
    If level > 0 Then
        If level < DEEPEST_LEVEL Then level = level + 1
        Paint rng, AlertRGB(AccentIndex(which), level)
    Else
        rng.Font.Fill.Transparency = NextStep(rng.Font.Fill.Transparency, Array(0, 0.3, 0.45, 0.7))
    End If
End Sub

' Size, weight and lightness of a level, applied to the text in place. An
' accent colour stays and takes the level's tint, as indenting would leave it.
Private Sub ApplyLevelLook(ByVal rng As TextRange2, ByVal base As String, ByVal level As Long)
    Dim which As Long, wasAlert As Long
    wasAlert = AlertLevelOf(rng.Font.Fill.ForeColor.RGB, which)
    rng.Font.Size = MasterLevelSize(level)
    SetFamily rng, LevelFamily(base, level), msoFalse
    If wasAlert > 0 Then
        Paint rng, AlertRGB(AccentIndex(which), level)
    Else
        With rng.Font.Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.RGB = RGB(0, 32, 64)
            .Transparency = LevelTransparency(level)
        End With
    End If
End Sub

Public Sub StepSmaller()
    Dim rng As TextRange2, base As String, level As Long
    If Not Selected(rng, base, level) Then Exit Sub
    Dim smaller As Single
    smaller = NextSmallerSize(rng.Font.Size)
    If smaller = 0 Then
        ' Bottom of the ladder. In an outline, come back to the paragraph's own
        ' look; elsewhere there is no known "regular" size, so stop here.
        If InOutline() Then ApplyLevelLook rng, base, level
        Exit Sub
    End If
    rng.Font.Size = smaller
    ApplyWeight rng, base, LighterSuffix(CurrentSuffix(rng, base))
    StepColourLighter rng
End Sub
