Option Explicit

' Excel Chess - the "Chess" sheet. Every click and every edit goes to Game.bas.

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' not Target.CountLarge: it crashes LibreOffice; Count alone overflows in Excel on a whole-sheet select
    If Target.Rows.Count <> 1 Or Target.Columns.Count <> 1 Then Exit Sub
    On Error GoTo Fail
    Game.OnClick Target.Row, Target.Column
    Exit Sub
Fail:
    XL.XFail Err.Description
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Fail
    Game.OnEdit
    Exit Sub
Fail:
    XL.XFail Err.Description
End Sub
