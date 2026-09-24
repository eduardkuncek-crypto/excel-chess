Attribute VB_Name = "XL"
' Excel Chess - the only module that touches Excel objects.
' Sheet1 = the "Chess" sheet, Sheet2 = the hidden "Data" sheet that remembers the game.
Option Explicit

Public Sub XBegin()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Sub

Public Sub XEnd()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' a cell's text. LibreOffice reports a number-looking text cell as '1 (with the apostrophe), Excel as 1
Private Function CellText(ByVal cell As Range) As String
    Dim f As String
    f = CStr(cell.Formula)
    If Left$(f, 1) = "'" Then f = Mid$(f, 2)
    CellText = f
End Function

Public Sub XSetText(ByVal r As Long, ByVal c As Long, ByVal s As String)
    If CellText(Sheet1.Cells(r, c)) <> s Then Sheet1.Cells(r, c).Value = s
End Sub

Public Function XGetText(ByVal r As Long, ByVal c As Long) As String
    XGetText = CellText(Sheet1.Cells(r, c))
End Function

Public Sub XSetFill(ByVal r As Long, ByVal c As Long, ByVal col As Long)
    With Sheet1.Cells(r, c).Interior
        If .Color <> col Then .Color = col
    End With
End Sub

' puts the cursor in the move box, so the next click anywhere counts as a click
Public Sub XPark()
    On Error Resume Next
    Application.EnableEvents = False
    ' compare names, not objects: "ActiveSheet Is Sheet1" is never True in LibreOffice
    If ActiveSheet.Name = Sheet1.Name Then Sheet1.Cells(6, 13).Select
    Application.EnableEvents = True
End Sub

' repaint now (before the engine starts thinking)
Public Sub XShowNow()
    Application.ScreenUpdating = True
    DoEvents
End Sub

Public Sub XWait(ByVal waiting As Boolean)
    If waiting Then Application.Cursor = xlWait Else Application.Cursor = xlDefault
End Sub

Public Function XLoad(ByVal key As Long) As String
    XLoad = CellText(Sheet2.Cells(key, 2))
End Function

Public Sub XSave(ByVal key As Long, ByVal v As String)
    Application.EnableEvents = False
    Sheet2.Cells(key, 2).Value = v
    Application.EnableEvents = True
End Sub

Public Function XLevel() As String
    XLevel = CellText(Sheet1.Cells(9, 13))
End Function

' One font for all 12 pieces. Segoe UI Symbol only exists on Windows; LibreOffice on Linux would
' borrow each piece from a different font (and draw the black pawn as a colour emoji).
Private Function PieceFont() As String
    If InStr(1, Application.OperatingSystem, "Unix") > 0 Or InStr(1, Application.OperatingSystem, "Linux") > 0 Then
        PieceFont = "Noto Sans Symbols 2"
    Else
        PieceFont = "Segoe UI Symbol"
    End If
End Function

' puts the board's look back (a drag moves the cell's formatting along with the piece) and sets
' the piece font for this computer; also called once when the file opens
Public Sub XRepairBoard()
    Application.EnableEvents = False
    Sheet1.Range("M10:P10").Font.Name = PieceFont()
    Sheet1.Range("L11:L12").Font.Name = PieceFont()
    With Sheet1.Range("C3:J10")
        .Font.Name = PieceFont()
        .Font.Size = 30
        .Font.Bold = False
        .Font.Italic = False
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "@"
        .Borders.LineStyle = xlNone
        .BorderAround xlContinuous, xlThick
    End With
    Application.EnableEvents = True
End Sub

' last line of defence: never leave Excel with events switched off
Public Sub XFail(ByVal desc As String)
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Game.ResetAfterError
    Sheet1.Cells(5, 12).Value = "Something went wrong (" & desc & "). Click any square to carry on."
End Sub
