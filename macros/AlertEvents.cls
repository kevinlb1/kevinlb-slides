VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AlertEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The add-in's ear on PowerPoint. PowerPoint has no event for "the text
' changed" or "this bullet was promoted", so the nearest thing is to watch the
' selection: when it leaves a shape, the shape is checked over. AlertText.bas
' does the work; this only passes the event on.
Option Explicit

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
    On Error Resume Next
    AlertText.OnSelectionChange Sel
End Sub

' A deck has opened: the first one says where to look for a newer add-in,
' and every one has its fonts checked.
Private Sub App_AfterPresentationOpen(ByVal Pres As Presentation)
    On Error Resume Next
    AlertText.MaybeUpdate
    AlertText.CheckFonts Pres
End Sub

Private Sub App_AfterNewPresentation(ByVal Pres As Presentation)
    On Error Resume Next
    AlertText.MaybeUpdate
End Sub

Private Sub App_PresentationClose(ByVal Pres As Presentation)
    On Error Resume Next
    AlertText.ForgetShape
End Sub
