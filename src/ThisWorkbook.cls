Option Explicit

' Excel Chess - draw the board as soon as the workbook opens with macros enabled.

Private Sub Workbook_Open()
    On Error GoTo Fail
    Sheet1.Activate
    Game.OnOpen
    Exit Sub
Fail:
    XL.XFail Err.Description
End Sub
