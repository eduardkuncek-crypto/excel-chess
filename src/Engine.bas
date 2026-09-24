Attribute VB_Name = "Engine"
' Excel Chess - the engine: board, every rule of chess, and the search.
' No worksheet access in here. Game.bas drives it, XL.bas is the only module that talks to Excel.
'
' Board: 10x12 mailbox. a1 = 21, h1 = 28, a8 = 91, h8 = 98, the border is off-board.
' Pieces: 1..6 = white pawn, knight, bishop, rook, queen, king. Negative = black. 0 = empty. 7 = off-board.
' A move is one Long: from + to*128 + promotion*16384 + flag*131072
' (flag 1 = en passant, 2 = castling, 3 = double pawn push).
Option Explicit

Public Const OFFB As Long = 7
Public Const MATE As Long = 100000
Public Const INF As Long = 1000000
Public Const MAXPLY As Long = 60
Public Const MAXHIST As Long = 4000
Public Const SLOT_UI As Long = 64
Public Const SLOT_SAN As Long = 65
Public Const SLOT_CHK As Long = 66
Public Const SLOT_PARSE As Long = 67
Public Const START_FEN As String = "rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR w KQkq - 0 1"

' the position
Public Bd(0 To 119) As Long
Public Side As Long
Public Castle As Long
Public EpSq As Long
Public Half As Long
Public FullMove As Long
Public KingSq(-1 To 1) As Long
Public HashKey As Long
Public HistPly As Long
Public PosHash(0 To MAXHIST) As Long

' what the last Think found
Public LastDepth As Long
Public LastScore As Long
Public LastNodes As Long
Public LastMs As Long
Public EvalNoise As Long

' undo stack, one entry per move made (game moves and search moves share it)
Private PHash As Long
Private UMove(0 To MAXHIST) As Long
Private UCap(0 To MAXHIST) As Long
Private UCastle(0 To MAXHIST) As Long
Private UEp(0 To MAXHIST) As Long
Private UHalf(0 To MAXHIST) As Long
Private UPHash(0 To MAXHIST) As Long
Private UHash(0 To MAXHIST) As Long

' move lists: slot = search ply (0..59), plus fixed slots for the UI, SAN, checks, parsing
Private Mv(0 To 67, 0 To 299) As Long
Private MvScore(0 To 67, 0 To 299) As Long
Private MvCount(0 To 67) As Long

Private Zob(0 To 12, 0 To 119) As Long
Private ZCastle(0 To 15) As Long
Private ZEp(0 To 119) As Long
Private ZSide As Long
Private RndState As Double

Private KnOff(0 To 7) As Long
Private SlideDir(0 To 7) As Long
Private DirLo(0 To 6) As Long
Private DirHi(0 To 6) As Long
Private CMask(0 To 119) As Long
Private Sq64(0 To 119) As Long
Private PVal(0 To 6) As Long
Private PhaseW(0 To 6) As Long
Private Pst(0 To 7, 0 To 63) As Long

Private TTKey(0 To 65535) As Long
Private TTMove(0 To 65535) As Long
Private Killer1(0 To 67) As Long
Private Killer2(0 To 67) As Long
Private HistH(0 To 119, 0 To 119) As Long

Private Ply As Long
Private Nodes As Long
Private StopFlag As Boolean
Private StartT As Double
Private TimeLimitMs As Long
Private RootBest As Long
Private Inited As Boolean
Private BookLines As String

' ====================================================================== setup

Public Sub InitEngine()
    Dim i As Long, j As Long, r As Long, f As Long, v() As String
    If Inited Then Exit Sub
    Inited = True
    Randomize
    RndState = 20260924#
    For i = 0 To 12
        For j = 0 To 119
            Zob(i, j) = NextRand()
        Next j
    Next i
    For i = 0 To 15
        ZCastle(i) = NextRand()
    Next i
    For i = 0 To 119
        ZEp(i) = NextRand()
        Sq64(i) = -1
        CMask(i) = 15
    Next i
    ZSide = NextRand()
    For r = 0 To 7
        For f = 0 To 7
            Sq64(21 + f + r * 10) = r * 8 + f
        Next f
    Next r
    ' castling rights die when anything moves from or to these squares
    CMask(21) = 13
    CMask(25) = 12
    CMask(28) = 14
    CMask(91) = 7
    CMask(95) = 3
    CMask(98) = 11
    v = Split("-21,-19,-12,-8,8,12,19,21", ",")
    For i = 0 To 7
        KnOff(i) = CLng(v(i))
    Next i
    ' 0..3 diagonal, 4..7 straight; the king uses all eight
    v = Split("-11,-9,9,11,-10,-1,1,10", ",")
    For i = 0 To 7
        SlideDir(i) = CLng(v(i))
    Next i
    DirLo(3) = 0
    DirHi(3) = 3
    DirLo(4) = 4
    DirHi(4) = 7
    DirLo(5) = 0
    DirHi(5) = 7
    v = Split("0,100,320,330,500,900,0", ",")
    For i = 0 To 6
        PVal(i) = CLng(v(i))
    Next i
    PhaseW(2) = 1
    PhaseW(3) = 1
    PhaseW(4) = 2
    PhaseW(5) = 4
    ' piece-square tables, a8 first (Tomasz Michniewski's simplified evaluation)
    LoadPst 1, "0,0,0,0,0,0,0,0,50,50,50,50,50,50,50,50,10,10,20,30,30,20,10,10,5,5,10,25,25,10,5,5,0,0,0,20,20,0,0,0,5,-5,-10,0,0,-10,-5,5,5,10,10,-20,-20,10,10,5,0,0,0,0,0,0,0,0"
    LoadPst 2, "-50,-40,-30,-30,-30,-30,-40,-50,-40,-20,0,0,0,0,-20,-40,-30,0,10,15,15,10,0,-30,-30,5,15,20,20,15,5,-30,-30,0,15,20,20,15,0,-30,-30,5,10,15,15,10,5,-30,-40,-20,0,5,5,0,-20,-40,-50,-40,-30,-30,-30,-30,-40,-50"
    LoadPst 3, "-20,-10,-10,-10,-10,-10,-10,-20,-10,0,0,0,0,0,0,-10,-10,0,5,10,10,5,0,-10,-10,5,5,10,10,5,5,-10,-10,0,10,10,10,10,0,-10,-10,10,10,10,10,10,10,-10,-10,5,0,0,0,0,5,-10,-20,-10,-10,-10,-10,-10,-10,-20"
    LoadPst 4, "0,0,0,0,0,0,0,0,5,10,10,10,10,10,10,5,-5,0,0,0,0,0,0,-5,-5,0,0,0,0,0,0,-5,-5,0,0,0,0,0,0,-5,-5,0,0,0,0,0,0,-5,-5,0,0,0,0,0,0,-5,0,0,0,5,5,0,0,0"
    LoadPst 5, "-20,-10,-10,-5,-5,-10,-10,-20,-10,0,0,0,0,0,0,-10,-10,0,5,5,5,5,0,-10,-5,0,5,5,5,5,0,-5,0,0,5,5,5,5,0,-5,-10,5,5,5,5,5,0,-10,-10,0,5,0,0,0,0,-10,-20,-10,-10,-5,-5,-10,-10,-20"
    LoadPst 6, "-30,-40,-40,-50,-50,-40,-40,-30,-30,-40,-40,-50,-50,-40,-40,-30,-30,-40,-40,-50,-50,-40,-40,-30,-30,-40,-40,-50,-50,-40,-40,-30,-20,-30,-30,-40,-40,-30,-30,-20,-10,-20,-20,-20,-20,-20,-20,-10,20,20,0,0,0,0,20,20,20,30,10,0,0,10,30,20"
    ' 7 = king in the endgame
    LoadPst 7, "-50,-40,-30,-20,-20,-30,-40,-50,-30,-20,-10,0,0,-10,-20,-30,-30,-10,20,30,30,20,-10,-30,-30,-10,30,40,40,30,-10,-30,-30,-10,30,40,40,30,-10,-30,-30,-10,20,30,30,20,-10,-30,-30,-30,0,0,0,0,-30,-30,-50,-30,-30,-30,-30,-30,-30,-50"
    InitBook
End Sub

Private Sub LoadPst(ByVal pt As Long, ByVal csv As String)
    Dim v() As String, i As Long
    v = Split(csv, ",")
    For i = 0 To 63
        Pst(pt, i) = CLng(v(i))
    Next i
End Sub

Private Sub InitBook()
    ' a small opening book, UCI moves; the engine picks a random line that matches the game so far
    BookLines = "e2e4 e7e5 g1f3 b8c6 f1b5 a7a6 b5a4 g8f6 e1g1 f8e7"
    BookLines = BookLines & "|e2e4 e7e5 g1f3 b8c6 f1c4 f8c5 c2c3 g8f6 d2d4 e5d4"
    BookLines = BookLines & "|e2e4 e7e5 g1f3 b8c6 d2d4 e5d4 f3d4 g8f6 d4c6 b7c6"
    BookLines = BookLines & "|e2e4 e7e5 g1f3 g8f6 f3e5 d7d6 e5f3 f6e4 d2d4 d6d5"
    BookLines = BookLines & "|e2e4 e7e5 b1c3 g8f6 f2f4 d7d5 f4e5 f6e4 g1f3 f8e7"
    BookLines = BookLines & "|e2e4 e7e5 f2f4 e5f4 g1f3 g7g5 h2h4 g5g4 f3e5 g8f6"
    BookLines = BookLines & "|e2e4 c7c5 g1f3 d7d6 d2d4 c5d4 f3d4 g8f6 b1c3 a7a6"
    BookLines = BookLines & "|e2e4 c7c5 g1f3 b8c6 d2d4 c5d4 f3d4 g8f6 b1c3 e7e5"
    BookLines = BookLines & "|e2e4 c7c5 g1f3 e7e6 d2d4 c5d4 f3d4 b8c6 b1c3 d8c7"
    BookLines = BookLines & "|e2e4 e7e6 d2d4 d7d5 b1c3 g8f6 c1g5 f8e7 e4e5 f6d7"
    BookLines = BookLines & "|e2e4 c7c6 d2d4 d7d5 b1c3 d5e4 c3e4 c8f5 e4g3 f5g6"
    BookLines = BookLines & "|e2e4 d7d5 e4d5 d8d5 b1c3 d5a5 d2d4 g8f6 g1f3 c8f5"
    BookLines = BookLines & "|e2e4 d7d6 d2d4 g8f6 b1c3 g7g6 g1f3 f8g7 f1e2 e8g8"
    BookLines = BookLines & "|d2d4 d7d5 c2c4 e7e6 b1c3 g8f6 c1g5 f8e7 e2e3 e8g8"
    BookLines = BookLines & "|d2d4 d7d5 c2c4 c7c6 g1f3 g8f6 b1c3 d5c4 a2a4 c8f5"
    BookLines = BookLines & "|d2d4 d7d5 g1f3 g8f6 c2c4 e7e6 b1c3 f8e7 c1f4 e8g8"
    BookLines = BookLines & "|d2d4 d7d5 c1f4 g8f6 e2e3 c7c5 c2c3 b8c6 g1f3 e7e6"
    BookLines = BookLines & "|d2d4 g8f6 c2c4 g7g6 b1c3 f8g7 e2e4 d7d6 g1f3 e8g8"
    BookLines = BookLines & "|d2d4 g8f6 c2c4 e7e6 b1c3 f8b4 e2e3 e8g8 f1d3 d7d5"
    BookLines = BookLines & "|d2d4 g8f6 c2c4 e7e6 g1f3 b7b6 g2g3 c8b7 f1g2 f8e7"
    BookLines = BookLines & "|c2c4 e7e5 b1c3 g8f6 g1f3 b8c6 g2g3 d7d5 c4d5 f6d5"
    BookLines = BookLines & "|g1f3 d7d5 g2g3 g8f6 f1g2 e7e6 e1g1 f8e7 d2d3 e8g8"
End Sub

' Park-Miller generator, so the hash keys are the same every time
Private Function NextRand() As Long
    RndState = RndState * 16807#
    RndState = RndState - Int(RndState / 2147483647#) * 2147483647#
    NextRand = CLng(RndState)
End Function

' ====================================================================== position setup

Public Function SetFEN(ByVal fen As String) As Boolean
    Dim parts() As String, i As Long, ch As String, r As Long, f As Long, k As Long, p As Long
    Dim wk As Long, bk As Long, cs As String
    InitEngine
    For i = 0 To 119
        Bd(i) = OFFB
    Next i
    For r = 0 To 7
        For f = 0 To 7
            Bd(21 + f + r * 10) = 0
        Next f
    Next r
    parts = Split(Trim$(fen), " ")
    If UBound(parts) < 1 Then Exit Function
    r = 7
    f = 0
    For i = 1 To Len(parts(0))
        ch = Mid$(parts(0), i, 1)
        If ch = "/" Then
            If f <> 8 Then Exit Function
            r = r - 1
            f = 0
            If r < 0 Then Exit Function
        ElseIf ch >= "1" And ch <= "8" Then
            f = f + CLng(ch)
            If f > 8 Then Exit Function
        Else
            k = InStr("PNBRQKpnbrqk", ch)
            If k = 0 Or f > 7 Then Exit Function
            If k <= 6 Then p = k Else p = 6 - k
            Bd(21 + f + r * 10) = p
            If p = 6 Then
                wk = wk + 1
                KingSq(1) = 21 + f + r * 10
            ElseIf p = -6 Then
                bk = bk + 1
                KingSq(-1) = 21 + f + r * 10
            End If
            f = f + 1
        End If
    Next i
    If r <> 0 Or f <> 8 Or wk <> 1 Or bk <> 1 Then Exit Function
    For f = 21 To 28
        If Abs(Bd(f)) = 1 Or Abs(Bd(f + 70)) = 1 Then Exit Function
    Next f
    If parts(1) = "w" Then
        Side = 1
    ElseIf parts(1) = "b" Then
        Side = -1
    Else
        Exit Function
    End If
    Castle = 0
    If UBound(parts) >= 2 Then
        cs = parts(2)
        If InStr(cs, "K") > 0 And Bd(25) = 6 And Bd(28) = 4 Then Castle = Castle Or 1
        If InStr(cs, "Q") > 0 And Bd(25) = 6 And Bd(21) = 4 Then Castle = Castle Or 2
        If InStr(cs, "k") > 0 And Bd(95) = -6 And Bd(98) = -4 Then Castle = Castle Or 4
        If InStr(cs, "q") > 0 And Bd(95) = -6 And Bd(91) = -4 Then Castle = Castle Or 8
    End If
    EpSq = 0
    If UBound(parts) >= 3 Then
        k = SqFromName(parts(3))
        ' only keep it if a pawn really did just make a double step
        If k <> 0 Then
            If Side = 1 And k \ 10 = 7 Then
                If Bd(k - 10) = -1 And Bd(k) = 0 And Bd(k + 10) = 0 Then EpSq = k
            ElseIf Side = -1 And k \ 10 = 4 Then
                If Bd(k + 10) = 1 And Bd(k) = 0 And Bd(k - 10) = 0 Then EpSq = k
            End If
        End If
    End If
    Half = 0
    FullMove = 1
    If UBound(parts) >= 4 Then Half = NumOr(parts(4), 0)
    If UBound(parts) >= 5 Then FullMove = NumOr(parts(5), 1)
    If FullMove < 1 Then FullMove = 1
    ' the side that just moved can't still be in check
    If Attacked(KingSq(-Side), Side) Then Exit Function
    HistPly = 0
    ComputeHash
    PosHash(0) = HashKey
    SetFEN = True
End Function

Public Function GetFEN() As String
    Dim r As Long, f As Long, e As Long, s As String, p As Long, c As String
    For r = 7 To 0 Step -1
        e = 0
        For f = 0 To 7
            p = Bd(21 + f + r * 10)
            If p = 0 Then
                e = e + 1
            Else
                If e > 0 Then s = s & CStr(e)
                e = 0
                If p > 0 Then s = s & Mid$("PNBRQK", p, 1) Else s = s & Mid$("pnbrqk", -p, 1)
            End If
        Next f
        If e > 0 Then s = s & CStr(e)
        If r > 0 Then s = s & "/"
    Next r
    If Side = 1 Then s = s & " w " Else s = s & " b "
    If (Castle And 1) <> 0 Then c = c & "K"
    If (Castle And 2) <> 0 Then c = c & "Q"
    If (Castle And 4) <> 0 Then c = c & "k"
    If (Castle And 8) <> 0 Then c = c & "q"
    If c = "" Then c = "-"
    s = s & c & " "
    If EpSq <> 0 Then s = s & SqName(EpSq) Else s = s & "-"
    GetFEN = s & " " & CStr(Half) & " " & CStr(FullMove)
End Function

Private Function NumOr(ByVal s As String, ByVal dflt As Long) As Long
    Dim i As Long
    NumOr = dflt
    If Len(s) = 0 Or Len(s) > 6 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    NumOr = CLng(s)
End Function

Private Sub ComputeHash()
    Dim sq As Long
    PHash = 0
    For sq = 21 To 98
        If Bd(sq) <> 0 And Bd(sq) <> OFFB Then PHash = PHash Xor Zob(Bd(sq) + 6, sq)
    Next sq
    HashKey = FullHash()
End Sub

' pieces + castling + side + en passant (only when a capture is actually possible)
Private Function FullHash() As Long
    Dim h As Long, cs As Long
    h = PHash Xor ZCastle(Castle)
    If Side = -1 Then h = h Xor ZSide
    If EpSq <> 0 Then
        cs = EpSq - 10 * Side
        If Bd(cs - 1) = Side Or Bd(cs + 1) = Side Then h = h Xor ZEp(EpSq)
    End If
    FullHash = h
End Function

' ====================================================================== squares and moves as text

Public Function SqName(ByVal sq As Long) As String
    SqName = Chr$(96 + (sq Mod 10)) & CStr(sq \ 10 - 1)
End Function

Public Function SqFromName(ByVal s As String) As Long
    Dim f As Long, r As Long
    If Len(s) <> 2 Then Exit Function
    f = InStr("abcdefgh", LCase$(Left$(s, 1)))
    r = InStr("12345678", Right$(s, 1))
    If f = 0 Or r = 0 Then Exit Function
    SqFromName = 10 + f + r * 10
End Function

Public Function MoveFrom(ByVal m As Long) As Long
    MoveFrom = m And 127
End Function

Public Function MoveTo(ByVal m As Long) As Long
    MoveTo = (m \ 128) And 127
End Function

Public Function MovePromo(ByVal m As Long) As Long
    MovePromo = (m \ 16384) And 7
End Function

Public Function MoveFlag(ByVal m As Long) As Long
    MoveFlag = m \ 131072
End Function

Private Function MkMove(ByVal fr As Long, ByVal t As Long, ByVal pr As Long, ByVal fl As Long) As Long
    MkMove = fr + t * 128 + pr * 16384 + fl * 131072
End Function

Public Function MoveUci(ByVal m As Long) As String
    Dim s As String
    s = SqName(m And 127) & SqName((m \ 128) And 127)
    If MovePromo(m) <> 0 Then s = s & Mid$("pnbrqk", MovePromo(m), 1)
    MoveUci = s
End Function

' Standard notation (Nf3, exd5, O-O, e8=Q+). The move must be legal in the current position.
Public Function MoveSan(ByVal m As Long) As String
    Dim fr As Long, t As Long, pr As Long, pt As Long, s As String
    Dim i As Long, n As Long, m2 As Long, f2 As Long, amb As Boolean, sameF As Boolean, sameR As Boolean
    fr = MoveFrom(m)
    t = MoveTo(m)
    pr = MovePromo(m)
    pt = Abs(Bd(fr))
    If MoveFlag(m) = 2 Then
        If t > fr Then s = "O-O" Else s = "O-O-O"
    ElseIf pt = 1 Then
        If Bd(t) <> 0 Or MoveFlag(m) = 1 Then s = Left$(SqName(fr), 1) & "x"
        s = s & SqName(t)
        If pr <> 0 Then s = s & "=" & Mid$("PNBRQK", pr, 1)
    Else
        s = Mid$("PNBRQK", pt, 1)
        n = GenLegal(SLOT_SAN)
        For i = 0 To n - 1
            m2 = Mv(SLOT_SAN, i)
            f2 = MoveFrom(m2)
            If f2 <> fr And MoveTo(m2) = t And Abs(Bd(f2)) = pt Then
                amb = True
                If (f2 Mod 10) = (fr Mod 10) Then sameF = True
                If (f2 \ 10) = (fr \ 10) Then sameR = True
            End If
        Next i
        If amb Then
            If Not sameF Then
                s = s & Left$(SqName(fr), 1)
            ElseIf Not sameR Then
                s = s & Right$(SqName(fr), 1)
            Else
                s = s & SqName(fr)
            End If
        End If
        If Bd(t) <> 0 Then s = s & "x"
        s = s & SqName(t)
    End If
    If MakeMove(m) Then
        If InCheck(Side) Then
            If HasLegal() Then s = s & "+" Else s = s & "#"
        End If
        UnmakeMove
    End If
    MoveSan = s
End Function

Private Function NormMove(ByVal s As String) As String
    s = Replace(s, "e.p.", "")
    s = Replace(s, " ", "")
    s = Replace(s, "+", "")
    s = Replace(s, "#", "")
    s = Replace(s, "!", "")
    s = Replace(s, "?", "")
    s = Replace(s, "x", "")
    s = Replace(s, "X", "")
    s = Replace(s, ":", "")
    s = Replace(s, "=", "")
    s = Replace(s, "-", "")
    s = Replace(s, "0", "O")
    NormMove = s
End Function

' Turns typed text into a legal move: UCI (e2e4, e7e8q) or SAN (Nf3, exd5, O-O, e8=Q). 0 = no match.
Public Function ParseMove(ByVal txt As String) As Long
    Dim n As Long, i As Long, m As Long, a As String, cnt As Long, found As Long
    a = NormMove(Trim$(txt))
    If a = "" Then Exit Function
    n = GenLegal(SLOT_PARSE)
    For i = 0 To n - 1
        m = Mv(SLOT_PARSE, i)
        If LCase$(MoveUci(m)) = LCase$(a) Then
            ParseMove = m
            Exit Function
        End If
    Next i
    For i = 0 To n - 1
        m = Mv(SLOT_PARSE, i)
        If NormMove(MoveSan(m)) = a Then
            ParseMove = m
            Exit Function
        End If
    Next i
    ' last try: ignore upper/lower case, but only if that leaves exactly one move
    For i = 0 To n - 1
        m = Mv(SLOT_PARSE, i)
        If LCase$(NormMove(MoveSan(m))) = LCase$(a) Then
            cnt = cnt + 1
            found = m
        End If
    Next i
    If cnt = 1 Then ParseMove = found
End Function

Private Function PieceName(ByVal pt As Long) As String
    Dim v() As String
    v = Split("nothing,pawn,knight,bishop,rook,queen,king", ",")
    PieceName = v(pt)
End Function

' Why moving the piece on fr to t is not allowed, in plain words.
Public Function ExplainIllegal(ByVal fr As Long, ByVal t As Long) As String
    Dim p As Long, pt As Long, i As Long, found As Boolean, m As Long
    p = Bd(fr)
    If p = 0 Or p = OFFB Then
        ExplainIllegal = "there is no piece on " & SqName(fr) & "."
        Exit Function
    End If
    If p * Side < 0 Then
        ExplainIllegal = "that's not your piece."
        Exit Function
    End If
    pt = p * Side
    If fr = t Then
        ExplainIllegal = "the piece has to go somewhere."
        Exit Function
    End If
    If Bd(t) * Side > 0 And Bd(t) <> OFFB Then
        ExplainIllegal = "you can't capture your own piece."
        Exit Function
    End If
    If pt = 6 And Abs(t - fr) = 2 And (fr = 25 Or fr = 95) Then
        ExplainIllegal = CastleReason(fr, t)
        Exit Function
    End If
    GenMoves SLOT_SAN, False
    For i = 0 To MvCount(SLOT_SAN) - 1
        m = Mv(SLOT_SAN, i)
        If MoveFrom(m) = fr And MoveTo(m) = t Then found = True
    Next i
    If found Then
        If pt = 6 Then
            ExplainIllegal = "your king can't move into check."
        ElseIf InCheck(Side) Then
            ExplainIllegal = "you're in check, and that move doesn't get your king out of it."
        Else
            ExplainIllegal = "that " & PieceName(pt) & " is pinned - moving it would leave your king in check."
        End If
        Exit Function
    End If
    If pt = 1 Then
        If (t - fr) * Side < 0 Then
            ExplainIllegal = "pawns can't move backwards."
        ElseIf (t - fr) * Side = 9 Or (t - fr) * Side = 11 Then
            ExplainIllegal = "pawns only move diagonally when they capture something."
        ElseIf (t - fr) * Side = 10 Or (t - fr) * Side = 20 Then
            ExplainIllegal = "pawns can't capture straight ahead, and can't jump over pieces."
        Else
            ExplainIllegal = "a pawn can't move like that."
        End If
    Else
        ExplainIllegal = "a " & PieceName(pt) & " can't move like that."
    End If
End Function

Private Function CastleReason(ByVal k As Long, ByVal t As Long) As String
    Dim bit As Long, d As Long, i As Long
    If t > k Then
        bit = 1
        d = 1
    Else
        bit = 2
        d = -1
    End If
    If Side = -1 Then bit = bit * 4
    If (Castle And bit) = 0 Then
        CastleReason = "you can't castle that way any more - the king or that rook has already moved."
        Exit Function
    End If
    If InCheck(Side) Then
        CastleReason = "you can't castle while you're in check."
        Exit Function
    End If
    For i = 1 To 3
        If i < 3 Or d = -1 Then
            If Bd(k + i * d) <> 0 Then
                CastleReason = "there are pieces in the way."
                Exit Function
            End If
        End If
    Next i
    If Attacked(k + d, -Side) Or Attacked(k + 2 * d, -Side) Then
        CastleReason = "the king can't castle through or into check."
        Exit Function
    End If
    CastleReason = "you can't castle right now."
End Function

' ====================================================================== move generation

Private Sub AddMove(ByVal slot As Long, ByVal m As Long)
    Mv(slot, MvCount(slot)) = m
    MvCount(slot) = MvCount(slot) + 1
End Sub

Private Sub AddPromos(ByVal slot As Long, ByVal fr As Long, ByVal t As Long)
    AddMove slot, MkMove(fr, t, 5, 0)
    AddMove slot, MkMove(fr, t, 2, 0)
    AddMove slot, MkMove(fr, t, 4, 0)
    AddMove slot, MkMove(fr, t, 3, 0)
End Sub

' All pseudo-legal moves for the side to move (captures and promotions only if capsOnly).
Public Sub GenMoves(ByVal slot As Long, ByVal capsOnly As Boolean)
    Dim sq As Long, p As Long, t As Long, c As Long, d As Long, i As Long, s As Long, fwd As Long
    s = Side
    fwd = 10 * s
    MvCount(slot) = 0
    For sq = 21 To 98
        p = Bd(sq) * s
        If p > 0 And p < 7 Then
            If p = 1 Then
                t = sq + fwd
                If Bd(t) = 0 Then
                    If t > 90 Or t < 30 Then
                        AddPromos slot, sq, t
                    ElseIf Not capsOnly Then
                        AddMove slot, MkMove(sq, t, 0, 0)
                        If (s = 1 And sq < 40) Or (s = -1 And sq > 80) Then
                            If Bd(t + fwd) = 0 Then AddMove slot, MkMove(sq, t + fwd, 0, 3)
                        End If
                    End If
                End If
                For i = -1 To 1 Step 2
                    t = sq + fwd + i
                    c = Bd(t)
                    If c <> OFFB And c * s < 0 Then
                        If t > 90 Or t < 30 Then
                            AddPromos slot, sq, t
                        Else
                            AddMove slot, MkMove(sq, t, 0, 0)
                        End If
                    ElseIf t = EpSq And EpSq <> 0 Then
                        AddMove slot, MkMove(sq, t, 0, 1)
                    End If
                Next i
            ElseIf p = 2 Or p = 6 Then
                For i = 0 To 7
                    If p = 2 Then t = sq + KnOff(i) Else t = sq + SlideDir(i)
                    c = Bd(t)
                    If c = 0 Then
                        If Not capsOnly Then AddMove slot, MkMove(sq, t, 0, 0)
                    ElseIf c <> OFFB And c * s < 0 Then
                        AddMove slot, MkMove(sq, t, 0, 0)
                    End If
                Next i
                If p = 6 And Not capsOnly Then GenCastles slot
            Else
                For i = DirLo(p) To DirHi(p)
                    d = SlideDir(i)
                    t = sq + d
                    c = Bd(t)
                    Do While c = 0
                        If Not capsOnly Then AddMove slot, MkMove(sq, t, 0, 0)
                        t = t + d
                        c = Bd(t)
                    Loop
                    If c <> OFFB And c * s < 0 Then AddMove slot, MkMove(sq, t, 0, 0)
                Next i
            End If
        End If
    Next sq
End Sub

Private Sub GenCastles(ByVal slot As Long)
    Dim k As Long, e As Long
    If Side = 1 Then k = 25 Else k = 95
    If Bd(k) <> 6 * Side Then Exit Sub
    If Side = 1 Then e = Castle And 3 Else e = (Castle \ 4) And 3
    If e = 0 Then Exit Sub
    If Attacked(k, -Side) Then Exit Sub
    If (e And 1) <> 0 Then
        If Bd(k + 1) = 0 And Bd(k + 2) = 0 And Bd(k + 3) = 4 * Side Then
            If Not Attacked(k + 1, -Side) And Not Attacked(k + 2, -Side) Then AddMove slot, MkMove(k, k + 2, 0, 2)
        End If
    End If
    If (e And 2) <> 0 Then
        If Bd(k - 1) = 0 And Bd(k - 2) = 0 And Bd(k - 3) = 0 And Bd(k - 4) = 4 * Side Then
            If Not Attacked(k - 1, -Side) And Not Attacked(k - 2, -Side) Then AddMove slot, MkMove(k, k - 2, 0, 2)
        End If
    End If
End Sub

' Is square sq attacked by side "by"?
Public Function Attacked(ByVal sq As Long, ByVal by As Long) As Boolean
    Dim i As Long, t As Long, c As Long, d As Long
    If by = 1 Then
        If Bd(sq - 9) = 1 Or Bd(sq - 11) = 1 Then
            Attacked = True
            Exit Function
        End If
    Else
        If Bd(sq + 9) = -1 Or Bd(sq + 11) = -1 Then
            Attacked = True
            Exit Function
        End If
    End If
    For i = 0 To 7
        If Bd(sq + KnOff(i)) = 2 * by Or Bd(sq + SlideDir(i)) = 6 * by Then
            Attacked = True
            Exit Function
        End If
    Next i
    For i = 0 To 7
        d = SlideDir(i)
        t = sq + d
        c = Bd(t)
        Do While c = 0
            t = t + d
            c = Bd(t)
        Loop
        If c = 5 * by Then
            Attacked = True
            Exit Function
        End If
        If i < 4 Then
            If c = 3 * by Then
                Attacked = True
                Exit Function
            End If
        Else
            If c = 4 * by Then
                Attacked = True
                Exit Function
            End If
        End If
    Next i
End Function

Public Function InCheck(ByVal s As Long) As Boolean
    InCheck = Attacked(KingSq(s), -s)
End Function

' ====================================================================== make / unmake

Private Sub MoveRook(ByVal a As Long, ByVal b As Long)
    PHash = PHash Xor Zob(Bd(a) + 6, a)
    Bd(b) = Bd(a)
    Bd(a) = 0
    PHash = PHash Xor Zob(Bd(b) + 6, b)
End Sub

' Plays a pseudo-legal move. Returns False (and takes it back) if it leaves the mover's king in check.
Public Function MakeMove(ByVal m As Long) As Boolean
    Dim fr As Long, t As Long, pr As Long, fl As Long, p As Long, cap As Long, s As Long, cs As Long
    fr = m And 127
    t = (m \ 128) And 127
    pr = (m \ 16384) And 7
    fl = m \ 131072
    s = Side
    p = Bd(fr)
    cap = Bd(t)
    HistPly = HistPly + 1
    UMove(HistPly) = m
    UCap(HistPly) = cap
    UCastle(HistPly) = Castle
    UEp(HistPly) = EpSq
    UHalf(HistPly) = Half
    UPHash(HistPly) = PHash
    UHash(HistPly) = HashKey
    PHash = PHash Xor Zob(p + 6, fr)
    If cap <> 0 Then PHash = PHash Xor Zob(cap + 6, t)
    Bd(fr) = 0
    If pr <> 0 Then Bd(t) = pr * s Else Bd(t) = p
    PHash = PHash Xor Zob(Bd(t) + 6, t)
    If p = s Or cap <> 0 Then Half = 0 Else Half = Half + 1
    If fl = 1 Then
        cs = t - 10 * s
        PHash = PHash Xor Zob(Bd(cs) + 6, cs)
        Bd(cs) = 0
    ElseIf fl = 2 Then
        ' block If on purpose: LibreOffice's Basic crashes on "If .. Then Sub a, b Else Sub c, d"
        If t > fr Then
            MoveRook fr + 3, fr + 1
        Else
            MoveRook fr - 4, fr - 1
        End If
    End If
    If p = 6 * s Then KingSq(s) = t
    Castle = Castle And CMask(fr) And CMask(t)
    If fl = 3 Then EpSq = fr + 10 * s Else EpSq = 0
    If s = -1 Then FullMove = FullMove + 1
    Side = -s
    HashKey = FullHash()
    PosHash(HistPly) = HashKey
    If Attacked(KingSq(s), -s) Then
        UnmakeMove
        MakeMove = False
    Else
        MakeMove = True
    End If
End Function

Public Sub UnmakeMove()
    Dim m As Long, fr As Long, t As Long, pr As Long, fl As Long, s As Long, p As Long
    m = UMove(HistPly)
    fr = m And 127
    t = (m \ 128) And 127
    pr = (m \ 16384) And 7
    fl = m \ 131072
    Side = -Side
    s = Side
    p = Bd(t)
    If pr <> 0 Then p = s
    Bd(fr) = p
    Bd(t) = UCap(HistPly)
    If fl = 1 Then
        Bd(t - 10 * s) = -s
    ElseIf fl = 2 Then
        If t > fr Then
            Bd(fr + 3) = Bd(fr + 1)
            Bd(fr + 1) = 0
        Else
            Bd(fr - 4) = Bd(fr - 1)
            Bd(fr - 1) = 0
        End If
    End If
    If p = 6 * s Then KingSq(s) = fr
    If s = -1 Then FullMove = FullMove - 1
    Castle = UCastle(HistPly)
    EpSq = UEp(HistPly)
    Half = UHalf(HistPly)
    PHash = UPHash(HistPly)
    HashKey = UHash(HistPly)
    HistPly = HistPly - 1
End Sub

' ====================================================================== legal moves and game end

' Fills the slot with only the legal moves and returns how many there are.
Public Function GenLegal(ByVal slot As Long) As Long
    Dim i As Long, n As Long, m As Long
    GenMoves slot, False
    For i = 0 To MvCount(slot) - 1
        m = Mv(slot, i)
        If MakeMove(m) Then
            UnmakeMove
            Mv(slot, n) = m
            n = n + 1
        End If
    Next i
    MvCount(slot) = n
    GenLegal = n
End Function

Public Function LegalMove(ByVal slot As Long, ByVal i As Long) As Long
    LegalMove = Mv(slot, i)
End Function

Public Function HasLegal() As Boolean
    Dim i As Long
    GenMoves SLOT_CHK, False
    For i = 0 To MvCount(SLOT_CHK) - 1
        If MakeMove(Mv(SLOT_CHK, i)) Then
            UnmakeMove
            HasLegal = True
            Exit Function
        End If
    Next i
End Function

' How many times the current position has occurred (same side to move, same rights).
Public Function RepCount() As Long
    Dim i As Long, n As Long, lim As Long
    lim = HistPly - Half
    If lim < 0 Then lim = 0
    i = HistPly
    Do While i >= lim
        If PosHash(i) = HashKey Then n = n + 1
        i = i - 2
    Loop
    RepCount = n
End Function

Private Function IsRep() As Boolean
    Dim i As Long, lim As Long
    lim = HistPly - Half
    If lim < 0 Then lim = 0
    i = HistPly - 2
    Do While i >= lim
        If PosHash(i) = HashKey Then
            IsRep = True
            Exit Function
        End If
        i = i - 2
    Loop
End Function

' True when neither side can possibly checkmate (K v K, K+minor v K, bishops all on one colour).
Public Function Insufficient() As Boolean
    Dim sq As Long, p As Long, knights As Long, bLight As Long, bDark As Long
    For sq = 21 To 98
        p = Abs(Bd(sq))
        If p = 1 Or p = 4 Or p = 5 Then Exit Function
        If p = 2 Then knights = knights + 1
        If p = 3 Then
            If ((sq Mod 10) + (sq \ 10)) Mod 2 = 1 Then bDark = bDark + 1 Else bLight = bLight + 1
        End If
    Next sq
    If knights + bLight + bDark <= 1 Then
        Insufficient = True
    ElseIf knights = 0 And (bLight = 0 Or bDark = 0) Then
        Insufficient = True
    End If
End Function

' 0 = still playing, 1 = checkmate (side to move lost), 2 = stalemate,
' 3 = 50-move rule, 4 = threefold repetition, 5 = insufficient material
Public Function GameState() As Long
    If Not HasLegal() Then
        If InCheck(Side) Then GameState = 1 Else GameState = 2
    ElseIf Half >= 100 Then
        GameState = 3
    ElseIf RepCount() >= 3 Then
        GameState = 4
    ElseIf Insufficient() Then
        GameState = 5
    End If
End Function

' ====================================================================== evaluation

Public Function MaxL(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxL = a Else MaxL = b
End Function

' pushes a lone king to the edge and brings the winning king closer, so won endgames get finished
Private Function MopUp(ByVal lk As Long, ByVal wk As Long) As Long
    Dim lf As Long, lr As Long, wf As Long, wr As Long
    lf = (lk Mod 10) - 1
    lr = (lk \ 10) - 2
    wf = (wk Mod 10) - 1
    wr = (wk \ 10) - 2
    MopUp = 10 * (MaxL(3 - lf, lf - 4) + MaxL(3 - lr, lr - 4)) + 4 * (14 - Abs(lf - wf) - Abs(lr - wr))
End Function

' Score in centipawns from the point of view of the side to move.
Public Function EvalPos() As Long
    Dim sq As Long, p As Long, s64 As Long, score As Long, phase As Long
    Dim wk As Long, bk As Long, wB As Long, bB As Long, wMat As Long, bMat As Long, wP As Long, bP As Long
    For sq = 21 To 98
        p = Bd(sq)
        If p > 0 And p < 6 Then
            s64 = Sq64(sq) Xor 56
            score = score + PVal(p) + Pst(p, s64)
            phase = phase + PhaseW(p)
            If p = 1 Then
                wP = wP + 1
            Else
                wMat = wMat + PVal(p)
                If p = 3 Then wB = wB + 1
            End If
        ElseIf p = 6 Then
            wk = Sq64(sq) Xor 56
        ElseIf p < 0 And p > -6 Then
            s64 = Sq64(sq)
            score = score - PVal(-p) - Pst(-p, s64)
            phase = phase + PhaseW(-p)
            If p = -1 Then
                bP = bP + 1
            Else
                bMat = bMat + PVal(-p)
                If p = -3 Then bB = bB + 1
            End If
        ElseIf p = -6 Then
            bk = Sq64(sq)
        End If
    Next sq
    If phase > 24 Then phase = 24
    ' the king wants shelter in the middlegame and the centre in the endgame
    score = score + (Pst(6, wk) * phase + Pst(7, wk) * (24 - phase)) \ 24
    score = score - (Pst(6, bk) * phase + Pst(7, bk) * (24 - phase)) \ 24
    If wB >= 2 Then score = score + 30
    If bB >= 2 Then score = score - 30
    If bMat = 0 And bP = 0 And wMat >= 500 Then score = score + MopUp(KingSq(-1), KingSq(1))
    If wMat = 0 And wP = 0 And bMat >= 500 Then score = score - MopUp(KingSq(1), KingSq(-1))
    If EvalNoise > 0 Then score = score + CLng(Rnd() * (2 * EvalNoise)) - EvalNoise
    EvalPos = score * Side
End Function

' ====================================================================== search

Private Function ElapsedMs() As Long
    Dim e As Double
    e = Timer - StartT
    If e < 0 Then e = e + 86400
    ElapsedMs = CLng(e * 1000)
End Function

' every 256 positions: often enough that LibreOffice (about 300 positions a second) never overruns
Private Sub CheckTime()
    If ElapsedMs() >= TimeLimitMs Then StopFlag = True
    DoEvents
End Sub

Private Sub ScoreMoves(ByVal slot As Long, ByVal ttm As Long)
    Dim i As Long, m As Long, v As Long, pr As Long, sc As Long
    For i = 0 To MvCount(slot) - 1
        m = Mv(slot, i)
        pr = MovePromo(m)
        If m = ttm Then
            sc = 1000000
        Else
            v = Abs(Bd(MoveTo(m)))
            If MoveFlag(m) = 1 Then v = 1
            If v <> 0 Then
                sc = 100000 + PVal(v) * 10 - Abs(Bd(MoveFrom(m)))
            ElseIf m = Killer1(slot) Then
                sc = 80000
            ElseIf m = Killer2(slot) Then
                sc = 70000
            Else
                sc = HistH(MoveFrom(m), MoveTo(m))
            End If
            If pr <> 0 Then sc = sc + 90000 + PVal(pr)
        End If
        MvScore(slot, i) = sc
    Next i
End Sub

' selection sort, one step: bring the best remaining move to position i
Private Sub PickMove(ByVal slot As Long, ByVal i As Long, ByVal n As Long)
    Dim j As Long, b As Long, tmp As Long
    b = i
    For j = i + 1 To n - 1
        If MvScore(slot, j) > MvScore(slot, b) Then b = j
    Next j
    If b <> i Then
        tmp = Mv(slot, i)
        Mv(slot, i) = Mv(slot, b)
        Mv(slot, b) = tmp
        tmp = MvScore(slot, i)
        MvScore(slot, i) = MvScore(slot, b)
        MvScore(slot, b) = tmp
    End If
End Sub

Private Function QSearch(ByVal alpha As Long, ByVal beta As Long) As Long
    Dim i As Long, n As Long, m As Long, score As Long, stand As Long
    Nodes = Nodes + 1
    If (Nodes And 255) = 0 Then CheckTime
    If StopFlag Then Exit Function
    stand = EvalPos()
    If stand >= beta Or Ply >= MAXPLY Then
        QSearch = stand
        Exit Function
    End If
    If stand > alpha Then alpha = stand
    GenMoves Ply, True
    n = MvCount(Ply)
    ScoreMoves Ply, 0
    For i = 0 To n - 1
        PickMove Ply, i, n
        m = Mv(Ply, i)
        If MakeMove(m) Then
            Ply = Ply + 1
            score = -QSearch(-beta, -alpha)
            Ply = Ply - 1
            UnmakeMove
            If StopFlag Then Exit Function
            If score > alpha Then
                alpha = score
                If alpha >= beta Then Exit For
            End If
        End If
    Next i
    QSearch = alpha
End Function

' Negamax alpha-beta with principal variation search, check extension and a move-ordering hash table.
Private Function Search(ByVal depth As Long, ByVal alpha As Long, ByVal beta As Long) As Long
    Dim i As Long, n As Long, m As Long, score As Long, best As Long, bestM As Long
    Dim legal As Long, chk As Boolean, idx As Long, ttm As Long
    Nodes = Nodes + 1
    If (Nodes And 255) = 0 Then CheckTime
    If StopFlag Then Exit Function
    If Ply > 0 Then
        If Half >= 100 Then Exit Function
        If IsRep() Then Exit Function
    End If
    chk = InCheck(Side)
    If chk Then depth = depth + 1
    If depth <= 0 Then
        Search = QSearch(alpha, beta)
        Exit Function
    End If
    If Ply >= MAXPLY Then
        Search = EvalPos()
        Exit Function
    End If
    idx = HashKey And 65535
    If TTKey(idx) = HashKey Then ttm = TTMove(idx)
    GenMoves Ply, False
    n = MvCount(Ply)
    ScoreMoves Ply, ttm
    best = -INF
    For i = 0 To n - 1
        PickMove Ply, i, n
        m = Mv(Ply, i)
        If MakeMove(m) Then
            legal = legal + 1
            Ply = Ply + 1
            If legal = 1 Then
                score = -Search(depth - 1, -beta, -alpha)
            Else
                score = -Search(depth - 1, -alpha - 1, -alpha)
                If score > alpha And score < beta Then score = -Search(depth - 1, -beta, -alpha)
            End If
            Ply = Ply - 1
            UnmakeMove
            If StopFlag Then Exit Function
            If score > best Then
                best = score
                bestM = m
                If score > alpha Then
                    alpha = score
                    If Ply = 0 Then RootBest = m
                    If alpha >= beta Then
                        If Bd(MoveTo(m)) = 0 And MoveFlag(m) <> 1 And MovePromo(m) = 0 Then
                            If Killer1(Ply) <> m Then
                                Killer2(Ply) = Killer1(Ply)
                                Killer1(Ply) = m
                            End If
                            If HistH(MoveFrom(m), MoveTo(m)) < 60000 Then HistH(MoveFrom(m), MoveTo(m)) = HistH(MoveFrom(m), MoveTo(m)) + depth * depth
                        End If
                        Exit For
                    End If
                End If
            End If
        End If
    Next i
    If legal = 0 Then
        If chk Then Search = -MATE + Ply Else Search = 0
        Exit Function
    End If
    TTKey(idx) = HashKey
    TTMove(idx) = bestM
    Search = best
End Function

' Iterative deepening. Returns the best move found (0 if there is no legal move).
Public Function Think(ByVal maxDepth As Long, ByVal maxMs As Long) As Long
    Dim d As Long, score As Long, best As Long, n As Long, i As Long, j As Long
    InitEngine
    Ply = 0
    Nodes = 0
    StopFlag = False
    StartT = Timer
    TimeLimitMs = maxMs
    LastDepth = 0
    LastScore = 0
    For i = 0 To 67
        Killer1(i) = 0
        Killer2(i) = 0
    Next i
    For i = 21 To 98
        For j = 21 To 98
            HistH(i, j) = 0
        Next j
    Next i
    n = GenLegal(0)
    If n = 0 Then Exit Function
    best = Mv(0, 0)
    If n > 1 Then
        For d = 1 To maxDepth
            RootBest = 0
            score = Search(d, -INF, INF)
            If StopFlag Then Exit For
            If RootBest <> 0 Then best = RootBest
            LastDepth = d
            LastScore = score
            ' the next depth takes a few times longer, so stop if half the time is gone
            If ElapsedMs() * 2 > maxMs Then Exit For
            If score > MATE - 500 Or score < 500 - MATE Then Exit For
        Next d
    End If
    LastNodes = Nodes
    LastMs = ElapsedMs()
    Think = best
End Function

' ====================================================================== opening book

Private Function FirstToken(ByVal s As String) As String
    Dim k As Long
    k = InStr(s, " ")
    If k = 0 Then FirstToken = s Else FirstToken = Left$(s, k - 1)
End Function

' hist = the game so far as UCI moves separated by spaces (from the normal start position)
Public Function BookMove(ByVal hist As String) As Long
    Dim ln() As String, tok() As String, i As Long, cand As String, nxt As String, h As String
    InitEngine
    h = Trim$(hist)
    ln = Split(BookLines, "|")
    For i = 0 To UBound(ln)
        nxt = ""
        If h = "" Then
            nxt = FirstToken(ln(i))
        ElseIf Left$(ln(i), Len(h) + 1) = h & " " Then
            nxt = FirstToken(Mid$(ln(i), Len(h) + 2))
        End If
        If nxt <> "" Then cand = cand & nxt & " "
    Next i
    If cand = "" Then Exit Function
    tok = Split(Trim$(cand), " ")
    BookMove = ParseMove(tok(Int(Rnd() * (UBound(tok) + 1))))
End Function
