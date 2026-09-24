Attribute VB_Name = "Game"
' Excel Chess - the game on the sheet: clicks, drags, typed moves, drawing the board, saving.
' Every worksheet read/write goes through XL.bas, so this module can be tested outside Excel.
Option Explicit

Private Const BR As Long = 3        ' board rows 3..10 (row 3 is rank 8 unless the board is flipped)
Private Const BC As Long = 3        ' board columns C..J
Private Const PC As Long = 12       ' side panel starts in column L
Private Const IN_R As Long = 6      ' the move box is M6
Private Const IN_C As Long = 13
Private Const BTN_R As Long = 8     ' New (White) / New (Black) / Take back / Resign
Private Const OPT_R As Long = 9     ' Level / Flip / Hint
Private Const PRO_R As Long = 10    ' promotion picker
Private Const LIST_R As Long = 15   ' first row of the move list

Private Loaded As Boolean
Private Busy As Boolean
Private StartFen As String
Private GMv(0 To MAXHIST) As Long
Private GSan(0 To MAXHIST) As String
Private GCap(0 To MAXHIST) As Long
Private NMoves As Long
Private Player As Long
Private Flipped As Boolean
Private Result As String
Private SelSq As Long
Private PromoFr As Long
Private PromoTo As Long
Private HintFr As Long
Private HintTo As Long
Private Shown(0 To 63) As String
Private Info As String
Private Msg As String
Private ListRows As Long
Private IgnR As Long
Private IgnC As Long
Private IgnT As Double
Private LvDepth As Long
Private LvMs As Long

' ====================================================================== entry points (called from the sheet)

Public Sub OnOpen()
    Loaded = False
    Busy = False
    XRepairBoard
    EnsureLoaded
    Msg = ""
    Render
    If Result = "" And Side <> Player Then EngineMove
End Sub

Public Sub ResetAfterError()
    Busy = False
    Loaded = False
End Sub

Public Sub OnClick(ByVal r As Long, ByVal c As Long)
    If Busy Then Exit Sub
    ' Excel selects the drop cell right after a drag; that click was the drag, not a new click
    If r = IgnR And c = IgnC Then
        IgnR = 0
        If Abs(Timer - IgnT) < 0.7 Then Exit Sub
    End If
    IgnR = 0
    EnsureLoaded
    If r >= BR And r <= BR + 7 And c >= BC And c <= BC + 7 Then
        If Result = "" And Side <> Player Then
            EngineMove
            XPark
        Else
            ClickSquare DispSq(r, c)
        End If
        Exit Sub
    End If
    If r = BTN_R And c = PC Then
        NewGame 1
    ElseIf r = BTN_R And c = PC + 1 Then
        NewGame -1
    ElseIf r = BTN_R And c = PC + 2 Then
        TakeBack
    ElseIf r = BTN_R And c = PC + 3 Then
        Resign
    ElseIf r = OPT_R And c = PC + 2 Then
        Flip
    ElseIf r = OPT_R And c = PC + 3 Then
        Hint
    ElseIf r = PRO_R And c >= PC + 1 And c <= PC + 4 Then
        If PromoFr = 0 Then Exit Sub
        ChoosePromo 5 - (c - PC - 1)
    Else
        Exit Sub
    End If
    XPark
End Sub

' Something on the sheet was edited: a move typed into the box, or a piece dragged to another square.
Public Sub OnEdit()
    Dim txt As String, i As Long, r As Long, c As Long, cur As String
    Dim nFill As Long, nVac As Long, fillSq As Long, vacSq As Long, fillR As Long, fillC As Long
    If Busy Then Exit Sub
    EnsureLoaded
    txt = Trim$(XGetText(IN_R, IN_C))
    If txt <> "" Then
        XBegin
        XSetText IN_R, IN_C, ""
        XEnd
        TypedMove txt
        Exit Sub
    End If
    For i = 0 To 63
        r = BR + i \ 8
        c = BC + i Mod 8
        cur = XGetText(r, c)
        If cur <> Shown(i) Then
            If cur = "" Then
                nVac = nVac + 1
                vacSq = DispSq(r, c)
            Else
                nFill = nFill + 1
                fillSq = DispSq(r, c)
                fillR = r
                fillC = c
            End If
        End If
    Next i
    If nVac + nFill = 0 Then
        Render
        Exit Sub
    End If
    XRepairBoard
    If nFill = 1 And nVac <= 1 Then
        If nVac = 0 Then vacSq = SelSq
        If vacSq <> 0 Then
            If XGetText(fillR, fillC) = Shown(SqIdx(vacSq)) Then
                DragMove vacSq, fillSq
                IgnR = fillR
                IgnC = fillC
                IgnT = Timer
                Exit Sub
            End If
        End If
    End If
    Msg = "Don't type on the board - click a piece, drag it, or type a move in the box."
    Render
End Sub

' ====================================================================== loading and saving

Private Sub EnsureLoaded()
    Dim mv() As String, i As Long, m As Long
    If Loaded Then Exit Sub
    InitEngine
    StartFen = XLoad(1)
    If StartFen = "" Then StartFen = START_FEN
    If Not SetFEN(StartFen) Then
        StartFen = START_FEN
        SetFEN StartFen
    End If
    If XLoad(3) = "b" Then Player = -1 Else Player = 1
    Flipped = (XLoad(4) = "1")
    Result = XLoad(5)
    Info = XLoad(6)
    NMoves = 0
    mv = Split(Trim$(XLoad(2)), " ")
    For i = 0 To UBound(mv)
        If mv(i) <> "" Then
            m = ParseMove(mv(i))
            If m = 0 Then Exit For
            RecordMove m
        End If
    Next i
    SelSq = 0
    PromoFr = 0
    PromoTo = 0
    HintFr = 0
    HintTo = 0
    ListRows = 400
    Loaded = True
End Sub

Private Function MovesUci() As String
    Dim i As Long, s As String
    For i = 0 To NMoves - 1
        If i > 0 Then s = s & " "
        s = s & MoveUci(GMv(i))
    Next i
    MovesUci = s
End Function

' plays a legal move on the engine board and writes it into the game record
Private Sub RecordMove(ByVal m As Long)
    GSan(NMoves) = MoveSan(m)
    If MoveFlag(m) = 1 Then GCap(NMoves) = -Side Else GCap(NMoves) = Bd(MoveTo(m))
    MakeMove m
    GMv(NMoves) = m
    NMoves = NMoves + 1
End Sub

Private Sub PlayMove(ByVal m As Long)
    RecordMove m
    SelSq = 0
    PromoFr = 0
    PromoTo = 0
    HintFr = 0
    HintTo = 0
    XSave 2, MovesUci()
    CheckEnd
End Sub

Private Function ScoreText(ByVal winner As Long) As String
    If winner = 1 Then ScoreText = "(1-0)" Else ScoreText = "(0-1)"
End Function

Private Sub CheckEnd()
    Dim st As Long
    st = GameState()
    If st = 1 Then
        If Side = Player Then
            Result = "Checkmate - the engine wins " & ScoreText(-Side)
        Else
            Result = "Checkmate - you win! " & ScoreText(-Side)
        End If
    ElseIf st = 2 Then
        Result = "Stalemate - it's a draw (1/2-1/2)"
    ElseIf st = 3 Then
        Result = "Draw by the 50-move rule (1/2-1/2)"
    ElseIf st = 4 Then
        Result = "Draw by threefold repetition (1/2-1/2)"
    ElseIf st = 5 Then
        Result = "Draw - nobody has enough pieces left to checkmate (1/2-1/2)"
    ElseIf NMoves >= MAXHIST - 100 Then
        Result = "Draw - the game got too long for this sheet (1/2-1/2)"
    End If
    If Result <> "" Then XSave 5, Result
End Sub

' ====================================================================== player actions

Private Function CanPlay() As Boolean
    If Result <> "" Then
        Msg = "The game is over - click a New game button to play again."
    ElseIf Side <> Player Then
        Msg = "It's the engine's turn."
    Else
        CanPlay = True
    End If
End Function

Private Function ColorName(ByVal s As Long) As String
    If s = 1 Then ColorName = "White" Else ColorName = "Black"
End Function

Private Sub ClickSquare(ByVal sq As Long)
    Dim p As Long, t As Long
    If Not CanPlay() Then
        SelSq = 0
        Render
        XPark
        Exit Sub
    End If
    If PromoFr <> 0 Then
        PromoFr = 0
        PromoTo = 0
        Msg = "Promotion cancelled."
    End If
    p = Bd(sq)
    If SelSq = 0 Then
        If p = 0 Then
            Msg = "Click one of your pieces first."
        ElseIf p * Side < 0 Then
            Msg = "That's not your piece - you are " & ColorName(Player) & "."
        Else
            SelSq = sq
            Msg = ""
        End If
        Render
        Exit Sub
    End If
    If sq = SelSq Then
        SelSq = 0
        Render
        XPark
        Exit Sub
    End If
    If p * Side > 0 Then
        ' king first, then its own rook = castle that way
        t = 0
        If Abs(Bd(SelSq)) = 6 And Abs(p) = 4 And sq \ 10 = SelSq \ 10 Then
            If sq > SelSq Then t = SelSq + 2 Else t = SelSq - 2
            If FindMove(SelSq, t, 0) = 0 Then t = 0
        End If
        If t <> 0 Then
            TryMove SelSq, t
        Else
            SelSq = sq
            Msg = ""
            Render
        End If
        Exit Sub
    End If
    TryMove SelSq, sq
End Sub

Private Sub DragMove(ByVal fr As Long, ByVal t As Long)
    SelSq = 0
    If Not CanPlay() Then
        Render
        XPark
        Exit Sub
    End If
    TryMove fr, t
End Sub

' the legal move fr -> t (with promotion piece pr, 0 = any); 0 if there is none
Private Function FindMove(ByVal fr As Long, ByVal t As Long, ByVal pr As Long) As Long
    Dim n As Long, i As Long, m As Long
    n = GenLegal(SLOT_UI)
    For i = 0 To n - 1
        m = LegalMove(SLOT_UI, i)
        If MoveFrom(m) = fr And MoveTo(m) = t Then
            If pr = 0 Or MovePromo(m) = pr Then
                FindMove = m
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub TryMove(ByVal fr As Long, ByVal t As Long)
    Dim m As Long
    SelSq = 0
    m = FindMove(fr, t, 0)
    If m = 0 Then
        Msg = "Illegal move: " & ExplainIllegal(fr, t)
        Render
        XPark
    ElseIf MovePromo(m) <> 0 Then
        PromoFr = fr
        PromoTo = t
        Msg = "Your pawn promotes! Pick the new piece on the right (M10 to P10)."
        Render
        XPark
    Else
        UserMove m
    End If
End Sub

Private Sub ChoosePromo(ByVal pt As Long)
    Dim m As Long
    m = FindMove(PromoFr, PromoTo, pt)
    PromoFr = 0
    PromoTo = 0
    If m = 0 Then
        Render
    Else
        UserMove m
    End If
End Sub

Private Sub UserMove(ByVal m As Long)
    Msg = ""
    PlayMove m
    If Result = "" Then
        EngineMove
    Else
        Render
    End If
    XPark
End Sub

Private Sub TypedMove(ByVal txt As String)
    Dim m As Long, fr As Long, t As Long, s As String
    If Not CanPlay() Then
        Render
        XPark
        Exit Sub
    End If
    PromoFr = 0
    PromoTo = 0
    SelSq = 0
    m = ParseMove(txt)
    If m <> 0 Then
        UserMove m
        Exit Sub
    End If
    s = LCase$(Replace(Replace(Replace(txt, "-", ""), " ", ""), "x", ""))
    If Len(s) >= 4 Then
        fr = SqFromName(Left$(s, 2))
        t = SqFromName(Mid$(s, 3, 2))
    End If
    If fr <> 0 And t <> 0 Then
        Msg = "Illegal move " & txt & ": " & ExplainIllegal(fr, t)
    Else
        Msg = "Not a legal move here: " & txt & ". Type moves like e2e4, Nf3, exd5 or O-O."
    End If
    Render
    XPark
End Sub

Private Sub NewGame(ByVal asSide As Long)
    If asSide = 1 Then
        XSave 3, "w"
    Else
        XSave 3, "b"
    End If
    XSave 1, START_FEN
    XSave 2, ""
    XSave 4, ""
    XSave 5, ""
    XSave 6, ""
    If asSide = -1 Then XSave 4, "1"
    Loaded = False
    EnsureLoaded
    If asSide = 1 Then Msg = "New game - you are White. Your move!" Else Msg = "New game - you are Black."
    Render
    If Side <> Player Then EngineMove
End Sub

Private Sub TakeBack()
    If NMoves = 0 Then
        Msg = "Nothing to take back."
        Render
        Exit Sub
    End If
    UnmakeMove
    NMoves = NMoves - 1
    If Side <> Player And NMoves > 0 Then
        UnmakeMove
        NMoves = NMoves - 1
    End If
    Result = ""
    Info = ""
    SelSq = 0
    PromoFr = 0
    PromoTo = 0
    HintFr = 0
    HintTo = 0
    Msg = "Move taken back."
    XSave 2, MovesUci()
    XSave 5, ""
    XSave 6, ""
    Render
    If Side <> Player Then EngineMove
End Sub

Private Sub Resign()
    If Result <> "" Then
        Msg = "The game is already over."
    Else
        Result = "You resigned - the engine wins " & ScoreText(-Player)
        XSave 5, Result
        Msg = ""
    End If
    Render
End Sub

Private Sub Flip()
    Flipped = Not Flipped
    If Flipped Then
        XSave 4, "1"
    Else
        XSave 4, ""
    End If
    Render
End Sub

Private Sub Hint()
    Dim m As Long
    If Not CanPlay() Then
        Render
        Exit Sub
    End If
    Busy = True
    Msg = "Looking for a good move for you..."
    Render
    XShowNow
    XWait True
    EvalNoise = 0
    m = Think(64, 1500)
    XWait False
    Busy = False
    If m <> 0 Then
        HintFr = MoveFrom(m)
        HintTo = MoveTo(m)
        Msg = "Hint: " & MoveSan(m) & " (highlighted in green)"
    End If
    Render
End Sub

' ====================================================================== the engine's turn

Private Sub SetLevel()
    Select Case LCase$(Trim$(XLevel()))
    Case "easy"
        LvDepth = 2
        LvMs = 1500
        EvalNoise = 60
    Case "medium"
        LvDepth = 3
        LvMs = 3000
        EvalNoise = 15
    Case "insane"
        LvDepth = MAXPLY - 4
        LvMs = 10000
        EvalNoise = 0
    Case Else
        LvDepth = MAXPLY - 4
        LvMs = 3000
        EvalNoise = 0
    End Select
End Sub

Private Function EvalText(ByVal sc As Long) As String
    Dim a As Long, s As String
    a = Abs(sc)
    If a > MATE - 1000 Then
        If sc > 0 Then s = "White" Else s = "Black"
        EvalText = s & " mates in " & CStr((MATE - a + 1) \ 2)
        Exit Function
    End If
    If sc < 0 Then s = "-" Else s = "+"
    EvalText = s & CStr(a \ 100) & "." & Right$("0" & CStr(a Mod 100), 2)
End Function

Private Sub EngineMove()
    Dim m As Long, fromBook As Boolean
    If Result <> "" Then Exit Sub
    Busy = True
    Msg = ""
    Render
    XShowNow
    XWait True
    SetLevel
    m = 0
    If StartFen = START_FEN And LvDepth > 2 Then m = BookMove(MovesUci())
    If m <> 0 Then
        fromBook = True
    Else
        m = Think(LvDepth, LvMs)
    End If
    EvalNoise = 0
    If m <> 0 Then
        If fromBook Then
            Info = "Engine played " & MoveSan(m) & "  |  from its opening book"
        Else
            Info = "Engine played " & MoveSan(m) & "  |  depth " & CStr(LastDepth) & "  |  eval " & EvalText(LastScore * Side) & "  |  " & CStr(LastNodes) & " positions in " & CStr(LastMs \ 1000) & "." & CStr((LastMs Mod 1000) \ 100) & "s"
        End If
        PlayMove m
        XSave 6, Info
    End If
    XWait False
    Busy = False
    Render
End Sub

' ====================================================================== drawing

Private Function DispSq(ByVal r As Long, ByVal c As Long) As Long
    If Flipped Then
        DispSq = 21 + (7 - (c - BC)) + (r - BR) * 10
    Else
        DispSq = 21 + (c - BC) + (7 - (r - BR)) * 10
    End If
End Function

Private Function SqIdx(ByVal sq As Long) As Long
    Dim f As Long, rk As Long
    f = (sq Mod 10) - 1
    rk = (sq \ 10) - 2
    If Flipped Then SqIdx = rk * 8 + (7 - f) Else SqIdx = (7 - rk) * 8 + f
End Function

Private Function Glyph(ByVal p As Long) As String
    If p = 0 Or p = OFFB Then
        Glyph = ""
    ElseIf p > 0 Then
        Glyph = ChrW(9818 - p)
    Else
        Glyph = ChrW(9824 + p)
    End If
End Function

Private Function PieceValue(ByVal p As Long) As Long
    Select Case Abs(p)
    Case 1
        PieceValue = 1
    Case 2, 3
        PieceValue = 3
    Case 4
        PieceValue = 5
    Case 5
        PieceValue = 9
    End Select
End Function

Private Function SquareColor(ByVal sq As Long, ByVal lastFr As Long, ByVal lastTo As Long, ByVal chkSq As Long, ByVal isTgt As Boolean) As Long
    Dim dark As Boolean
    dark = (((sq Mod 10) + (sq \ 10)) Mod 2 = 1)
    If sq = chkSq Then
        SquareColor = RGB(232, 90, 75)
    ElseIf sq = SelSq Or sq = PromoFr Then
        SquareColor = RGB(247, 236, 93)
    ElseIf isTgt Or sq = PromoTo Then
        If dark Then SquareColor = RGB(96, 146, 196) Else SquareColor = RGB(160, 199, 235)
    ElseIf sq = HintFr Or sq = HintTo Then
        If dark Then SquareColor = RGB(86, 160, 86) Else SquareColor = RGB(152, 212, 152)
    ElseIf sq = lastFr Or sq = lastTo Then
        If dark Then SquareColor = RGB(170, 162, 58) Else SquareColor = RGB(205, 210, 106)
    Else
        If dark Then SquareColor = RGB(181, 136, 99) Else SquareColor = RGB(240, 217, 181)
    End If
End Function

Private Sub Render()
    Dim i As Long, r As Long, c As Long, sq As Long, n As Long, k As Long, m As Long
    Dim tgt(0 To 119) As Boolean, lastFr As Long, lastTo As Long, chkSq As Long
    XBegin
    If SelSq <> 0 Then
        n = GenLegal(SLOT_UI)
        For k = 0 To n - 1
            m = LegalMove(SLOT_UI, k)
            If MoveFrom(m) = SelSq Then tgt(MoveTo(m)) = True
        Next k
    End If
    If NMoves > 0 Then
        lastFr = MoveFrom(GMv(NMoves - 1))
        lastTo = MoveTo(GMv(NMoves - 1))
    End If
    If InCheck(Side) Then chkSq = KingSq(Side)
    For i = 0 To 63
        r = BR + i \ 8
        c = BC + i Mod 8
        sq = DispSq(r, c)
        Shown(i) = Glyph(Bd(sq))
        XSetText r, c, Shown(i)
        XSetFill r, c, SquareColor(sq, lastFr, lastTo, chkSq, tgt(sq))
    Next i
    For i = 0 To 7
        If Flipped Then
            XSetText BR + i, BC - 1, CStr(i + 1)
            XSetText BR + 8, BC + i, Chr$(104 - i)
        Else
            XSetText BR + i, BC - 1, CStr(8 - i)
            XSetText BR + 8, BC + i, Chr$(97 + i)
        End If
    Next i
    RenderPanel
    XEnd
End Sub

Private Sub RenderPanel()
    Dim st As String, i As Long, k As Long, wTook As String, bTook As String, diff As Long
    If Result <> "" Then
        st = Result
    ElseIf Busy Then
        st = "Engine is thinking..."
    ElseIf Side = Player Then
        st = "Your move (" & ColorName(Player) & ")"
        If InCheck(Side) Then st = st & " - you're in CHECK!"
    Else
        st = "Engine to move - click the board"
    End If
    XSetText 3, PC, st
    XSetText 4, PC, Info
    XSetText 5, PC, Msg
    XSetText IN_R, PC, "Your move:"
    XSetText IN_R + 1, PC, "Type e2e4, Nf3, exd5 or O-O in the yellow box and press Enter"
    XSetText BTN_R, PC, "New game as White"
    XSetText BTN_R, PC + 1, "New game as Black"
    XSetText BTN_R, PC + 2, "Take back"
    XSetText BTN_R, PC + 3, "Resign"
    XSetText OPT_R, PC, "Level:"
    XSetText OPT_R, PC + 2, "Flip board"
    XSetText OPT_R, PC + 3, "Hint"
    If PromoFr <> 0 Then
        XSetText PRO_R, PC, "Promote to:"
        For i = 0 To 3
            XSetText PRO_R, PC + 1 + i, Glyph((5 - i) * Side)
        Next i
    Else
        XSetText PRO_R, PC, ""
        For i = 0 To 3
            XSetText PRO_R, PC + 1 + i, ""
        Next i
    End If
    For k = 0 To NMoves - 1
        If GCap(k) < 0 Then
            wTook = wTook & Glyph(GCap(k))
            diff = diff + PieceValue(GCap(k))
        ElseIf GCap(k) > 0 Then
            bTook = bTook & Glyph(GCap(k))
            diff = diff - PieceValue(GCap(k))
        End If
    Next k
    If diff > 0 Then wTook = wTook & "  +" & CStr(diff)
    If diff < 0 Then bTook = bTook & "  +" & CStr(-diff)
    XSetText 11, PC, "White took: " & wTook
    XSetText 12, PC, "Black took: " & bTook
    XSetText LIST_R - 1, PC, "Move"
    XSetText LIST_R - 1, PC + 1, "White"
    XSetText LIST_R - 1, PC + 2, "Black"
    k = (NMoves + 1) \ 2
    For i = 0 To MaxL(k, ListRows) - 1
        If i < k Then
            XSetText LIST_R + i, PC, CStr(i + 1) & "."
            XSetText LIST_R + i, PC + 1, GSan(2 * i)
            If 2 * i + 1 < NMoves Then
                XSetText LIST_R + i, PC + 2, GSan(2 * i + 1)
            Else
                XSetText LIST_R + i, PC + 2, ""
            End If
        Else
            XSetText LIST_R + i, PC, ""
            XSetText LIST_R + i, PC + 1, ""
            XSetText LIST_R + i, PC + 2, ""
        End If
    Next i
    ListRows = k
End Sub
