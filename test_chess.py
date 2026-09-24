#!/usr/bin/env python3
"""Tests for Excel Chess. Translates the real .bas files with vba2py (which also lints
them), then checks the engine against python-chess and drives the Game module through a
fake worksheet. Needs: pip install chess.   Run: python3 test_chess.py
"""
# aios-run: agent
import importlib.util
import os
import random
import sys
import tempfile
import time

import chess

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import vba2py  # noqa: E402


def load_engine():
    src = {m: open(os.path.join(HERE, 'src', f'{m}.bas')).read() for m in ('Engine', 'Game')}
    code = vba2py.translate(src, externs={'XL': open(os.path.join(HERE, 'src', 'XL.bas')).read()})
    path = os.path.join(tempfile.mkdtemp(), 'chess_gen.py')
    open(path, 'w').write(code)
    spec = importlib.util.spec_from_file_location('chess_gen', path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    mod.V_InitEngine()
    return mod


E = load_engine()


def sq(name):
    return E.V_SqFromName(name)


def legal_uci():
    n = E.V_GenLegal(E.V_SLOT_UI)
    return sorted(E.V_MoveUci(E.V_LegalMove(E.V_SLOT_UI, i)) for i in range(n))


def play_uci(u):
    m = E.V_ParseMove(u)
    assert m, u
    assert E.V_MakeMove(m)


# ---------------------------------------------------------------------- rules vs python-chess

def test_random_games(games=40, plies=150, seed=1):
    rng = random.Random(seed)
    checked = 0
    for g in range(games):
        b = chess.Board()
        assert E.V_SetFEN(b.fen())
        for _ in range(plies):
            mine = legal_uci()
            ref = sorted(m.uci() for m in b.legal_moves)
            assert mine == ref, (b.fen(), set(mine) ^ set(ref))
            assert E.V_GetFEN() == b.fen(en_passant='fen'), (E.V_GetFEN(), b.fen(en_passant='fen'))
            st = E.V_GameState()
            if b.is_checkmate():
                assert st == 1, b.fen()
            elif b.is_stalemate():
                assert st == 2, b.fen()
            elif b.halfmove_clock >= 100:
                assert st == 3, b.fen()
            elif b.is_repetition(3):
                assert st == 4, b.fen()
            elif b.is_insufficient_material():
                assert st == 5, b.fen()
            else:
                assert st == 0, (b.fen(), st)
            if st:
                break
            # bias toward captures/promotions/castling so rare rules come up more
            moves = list(b.legal_moves)
            special = [x for x in moves if b.is_capture(x) or x.promotion or b.is_castling(x)]
            mv = rng.choice(special if special and rng.random() < 0.4 else moves)
            n = E.V_GenLegal(E.V_SLOT_UI)
            mine_moves = [E.V_LegalMove(E.V_SLOT_UI, i) for i in range(n)]
            for m in mine_moves:                       # SAN printing: every move
                u = E.V_MoveUci(m)
                assert E.V_MoveSan(m) == b.san(chess.Move.from_uci(u)), (b.fen(), u, E.V_MoveSan(m))
            for x in [mv] + rng.sample(moves, min(2, len(moves))):   # SAN parsing: a sample
                assert E.V_MoveUci(E.V_ParseMove(b.san(x))) == x.uci(), (b.fen(), b.san(x))
            play_uci(mv.uci())
            b.push(mv)
            checked += 1
    return f'{checked} positions'


def test_rule_positions():
    cases = [
        # en passant, including the pinned pawn that must not capture
        ('8/8/8/KPp4r/8/8/8/7k w - c6 0 1', 'b5c6', False),
        ('8/8/8/1Pp5/8/8/8/K6k w - c6 0 1', 'b5c6', True),
        # castling through check / out of check / with rights gone
        ('r3k2r/8/8/8/8/8/8/R3K1r1 w KQkq - 0 1', 'e1c1', False),
        ('r3k2r/8/8/8/8/8/5r2/R3K2R w KQkq - 0 1', 'e1g1', False),
        ('r3k2r/8/8/8/8/8/8/R3K2R w Qkq - 0 1', 'e1g1', False),
        ('r3k2r/8/8/8/8/8/8/R3K2R w KQkq - 0 1', 'e1g1', True),
        ('r3k2r/8/8/8/8/8/8/R3K2R w KQkq - 0 1', 'e1c1', True),
        ('r3k2r/8/8/8/8/8/8/RN2K2R w KQkq - 0 1', 'e1c1', False),
        # promotion to every piece
        ('8/P6k/8/8/8/8/8/K7 w - - 0 1', 'a7a8n', True),
        ('8/P6k/8/8/8/8/8/K7 w - - 0 1', 'a7a8', False),
    ]
    for fen, uci, ok in cases:
        assert E.V_SetFEN(fen), fen
        assert (uci in legal_uci()) == ok, (fen, uci)
    # explanations for illegal moves
    E.V_SetFEN('r3k2r/8/8/8/8/8/8/R3K2R w Qkq - 0 1')
    assert 'already moved' in E.V_ExplainIllegal(sq('e1'), sq('g1'))
    E.V_SetFEN('4k3/8/8/8/8/8/8/R3K2r w Q - 0 1')
    assert 'in check' in E.V_ExplainIllegal(sq('e1'), sq('c1'))
    E.V_SetFEN('4k3/4r3/8/8/8/8/4B3/4K3 w - - 0 1')
    assert 'pinned' in E.V_ExplainIllegal(sq('e2'), sq('d3'))
    E.V_SetFEN(chess.STARTING_FEN)
    assert "can't move like that" in E.V_ExplainIllegal(sq('b1'), sq('b3'))
    assert 'own piece' in E.V_ExplainIllegal(sq('d1'), sq('d2'))
    assert "not your piece" in E.V_ExplainIllegal(sq('e7'), sq('e5'))
    assert 'diagonally' in E.V_ExplainIllegal(sq('e2'), sq('d3'))
    # threefold repetition through the game history
    E.V_SetFEN(chess.STARTING_FEN)
    for u in 'g1f3 g8f6 f3g1 f6g8 g1f3 g8f6 f3g1 f6g8'.split():
        play_uci(u)
    assert E.V_GameState() == 4
    # a bad FEN is refused
    assert not E.V_SetFEN('8/8/8/8/8/8/8/8 w - - 0 1')
    assert not E.V_SetFEN('rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR x KQkq - 0 1')
    assert not E.V_SetFEN('4k3/8/8/8/8/8/8/4K2r b - - 0 1')   # side not to move is in check
    return 'ok'


def test_book_lines():
    lines = E.V_Engine_BookLines.split('|')
    for ln in lines:
        b = chess.Board()
        for u in ln.split():
            b.push_uci(u)          # raises if a book move is illegal
    for _ in range(30):
        E.V_SetFEN(chess.STARTING_FEN)
        m = E.V_BookMove('')
        assert E.V_MoveUci(m) in {ln.split()[0] for ln in lines}
        play_uci(E.V_MoveUci(m))
        m2 = E.V_BookMove(E.V_MoveUci(m))
        assert m2 != 0
    return f'{len(lines)} lines'


# ---------------------------------------------------------------------- the engine plays

def mate_in(b, n):
    """True if the side to move can force mate in n moves (brute force, python-chess)."""
    for mv in list(b.legal_moves):
        b.push(mv)
        if b.is_checkmate():
            b.pop()
            return True
        ok = n > 1 and not b.is_game_over() and all(
            (b.push(r) or True) and (mate_in(b, n - 1), b.pop())[0] for r in list(b.legal_moves))
        b.pop()
        if ok:
            return True
    return False


def think(fen, depth=64, ms=3000):
    assert E.V_SetFEN(fen), fen
    E.V_EvalNoise = 0
    t = time.time()
    m = E.V_Think(depth, ms)
    return E.V_MoveUci(m), time.time() - t


def test_tactics():
    out = []
    u, t = think('6k1/5ppp/8/8/8/8/5PPP/R5K1 w - - 0 1', 4)
    assert u == 'a1a8', u
    out.append(f'mate in 1 {t:.1f}s')
    fen = 'r2qkb1r/pp2nppp/3p4/2pNN1B1/2BnP3/3P4/PPP2PPP/R2bK2R w KQkq - 1 1'
    assert mate_in(chess.Board(fen), 2) and not mate_in(chess.Board(fen), 1)
    u, t = think(fen, 5, 60000)
    b = chess.Board(fen)
    b.push_uci(u)
    assert all((b.push(r) or True) and (mate_in(b, 1), b.pop())[0] for r in list(b.legal_moves)), u
    out.append(f'mate in 2 ({u}) {t:.1f}s')
    u, t = think('rnb1kbnr/pppp1ppp/8/4q3/8/5N2/PPPPPPPP/RNBQKB1R w KQkq - 0 1', 3)
    assert u == 'f3e5', u
    out.append(f'wins the queen {t:.1f}s')
    u, t = think('4k3/8/8/8/8/8/8/R3K3 w Q - 0 1', 64, 700)
    assert t < 3, t
    out.append(f'stops on time ({t:.1f}s for 0.7s)')
    return ', '.join(out)


def test_converts_kq_vs_k(seed=3):
    rng = random.Random(seed)
    b = chess.Board('8/8/8/4k3/8/8/8/3QK3 w - - 0 1')
    E.V_SetFEN(b.fen())
    for ply in range(120):
        if b.is_game_over():
            break
        if b.turn == chess.WHITE:
            E.V_EvalNoise = 0
            m = E.V_Think(3, 60000)
            u = E.V_MoveUci(m)
        else:
            u = rng.choice(list(b.legal_moves)).uci()
        play_uci(u)
        b.push_uci(u)
    assert b.is_checkmate(), b.fen()
    return f'mated in {b.fullmove_number} moves'


def test_selfplay(seed=7):
    b = chess.Board()
    E.V_SetFEN(b.fen())
    random.seed(seed)
    while not b.is_game_over(claim_draw=False) and b.ply() < 100:
        E.V_EvalNoise = 30
        m = E.V_Think(2, 60000)
        u = E.V_MoveUci(m)
        assert chess.Move.from_uci(u) in b.legal_moves, (b.fen(), u)
        play_uci(u)
        b.push_uci(u)
    return f'{b.ply()} legal plies, {b.result(claim_draw=True)}'


# ---------------------------------------------------------------------- the sheet

class Sheet:
    def __init__(self):
        self.cells, self.fill, self.data, self.parks = {}, {}, {}, 0

    def install(self):
        E.V_XBegin = lambda: None
        E.V_XEnd = lambda: None
        E.V_XSetText = lambda r, c, s: self.cells.__setitem__((r, c), s)
        E.V_XGetText = lambda r, c: self.cells.get((r, c), '')
        E.V_XSetFill = lambda r, c, col: self.fill.__setitem__((r, c), col)
        E.V_XPark = self.park
        E.V_XShowNow = lambda: None
        E.V_XWait = lambda w: None
        E.V_XLoad = lambda k: self.data.get(k, '')
        E.V_XSave = lambda k, v: self.data.__setitem__(k, v)
        E.V_XLevel = lambda: self.cells.get((9, 13), 'Easy')
        E.V_XRepairBoard = lambda: None
        E.V_XFail = lambda d: (_ for _ in ()).throw(AssertionError(d))
        return self

    def park(self):
        self.parks += 1

    def cell(self, name, flipped=False):
        f, r = ord(name[0]) - 97, int(name[1])
        return (3 + r - 1, 3 + 7 - f) if flipped else (3 + 8 - r, 3 + f)

    def click(self, name, flipped=False):
        E.V_OnClick(*self.cell(name, flipped))

    def type(self, text):
        self.cells[(6, 13)] = text
        E.V_OnEdit()

    def status(self):
        return self.cells.get((3, 12), '')

    def msg(self):
        return self.cells.get((5, 12), '')

    def moves(self):
        return self.data.get(2, '').split()

    def board(self):
        b = chess.Board(self.data.get(1) or chess.STARTING_FEN)
        for u in self.moves():
            b.push_uci(u)
        return b

    def open(self, fen=None, player='w', level='Easy'):
        self.cells, self.fill = {}, {}
        self.data = {1: fen or '', 2: '', 3: player, 4: '1' if player == 'b' else '', 5: '', 6: ''}
        self.cells[(9, 13)] = level
        E.V_OnOpen()


GLYPH = {'K': '♔', 'Q': '♕', 'R': '♖', 'B': '♗', 'N': '♘', 'P': '♙',
         'k': '♚', 'q': '♛', 'r': '♜', 'b': '♝', 'n': '♞', 'p': '♟'}


def check_board_cells(sh, flipped=False):
    b = sh.board()
    for s in chess.SQUARES:
        name = chess.square_name(s)
        p = b.piece_at(s)
        assert sh.cells[sh.cell(name, flipped)] == (GLYPH[p.symbol()] if p else ''), name


def test_sheet():
    sh = Sheet().install()
    sh.open()
    check_board_cells(sh)
    assert sh.status() == 'Your move (White)', sh.status()
    # click-click move; the engine answers
    sh.click('e2')
    assert sh.fill[sh.cell('e4')] != sh.fill[sh.cell('e5')], 'e4 should be highlighted'
    sh.click('e4')
    assert sh.moves()[0] == 'e2e4' and len(sh.moves()) == 2, sh.moves()
    assert sh.cells[(15, 13)] == 'e4' and sh.cells[(15, 14)] != '', 'move list'
    assert sh.cells[(4, 12)].startswith('Engine played'), sh.cells[(4, 12)]
    check_board_cells(sh)
    # illegal click-click: nothing is played, the reason is shown
    sh.click('b1')
    sh.click('b3')
    assert len(sh.moves()) == 2 and "knight can't move like that" in sh.msg(), sh.msg()
    sh.click('e7' if sh.board().piece_at(chess.E7) else 'd8')
    assert "not your piece" in sh.msg(), sh.msg()
    sh.click('e3')
    assert 'Click one of your pieces' in sh.msg(), sh.msg()
    # typed move and typed rubbish
    sh.type('hello')
    assert 'Not a legal move' in sh.msg() and sh.cells[(6, 13)] == ''
    sh.type('e2e5')
    assert 'Illegal move' in sh.msg(), sh.msg()
    b = sh.board()
    san = next(b.san(m) for m in b.legal_moves if b.piece_at(m.from_square).piece_type == chess.KNIGHT)
    sh.type(san)
    assert len(sh.moves()) == 4, (san, sh.moves(), sh.msg())
    # drag and drop: source emptied, glyph appears on the target
    b = sh.board()
    mv = next(m for m in b.legal_moves if not m.promotion and not b.is_castling(m))
    src, dst = chess.square_name(mv.from_square), chess.square_name(mv.to_square)
    sh.cells[sh.cell(dst)] = sh.cells[sh.cell(src)]
    sh.cells[sh.cell(src)] = ''
    E.V_OnEdit()
    assert sh.moves()[4] == mv.uci(), (sh.moves(), mv.uci(), sh.msg())
    E.V_OnClick(*sh.cell(dst))        # Excel then selects the drop cell: must be ignored
    assert len(sh.moves()) == 6 and sh.msg() == '', sh.msg()
    check_board_cells(sh)
    # drag where Excel reports only the target (piece was selected by the mouse-down)
    b = sh.board()
    mv = next(m for m in b.legal_moves if not m.promotion and not b.is_castling(m))
    src, dst = chess.square_name(mv.from_square), chess.square_name(mv.to_square)
    E.V_OnClick(*sh.cell(src))
    sh.cells[sh.cell(dst)] = sh.cells[sh.cell(src)]
    E.V_OnEdit()
    assert sh.moves()[6] == mv.uci(), (sh.moves(), mv.uci(), sh.msg())
    E.V_OnClick(*sh.cell(dst))
    # typing on the board is undone
    sh.cells[sh.cell('d5')] = 'x'
    E.V_OnEdit()
    assert "Don't type on the board" in sh.msg() and len(sh.moves()) == 8
    check_board_cells(sh)
    # take back removes your move and the reply
    E.V_OnClick(8, 14)
    assert len(sh.moves()) == 6 and sh.board().turn == chess.WHITE
    check_board_cells(sh)
    # reopening the file rebuilds exactly the same game
    before = sh.board().fen()
    saved = dict(sh.data)
    sh.cells = {(9, 13): 'Easy'}
    E.V_ResetAfterError()
    E.V_OnOpen()
    assert sh.data == saved and sh.board().fen() == before
    check_board_cells(sh)
    # hint, flip, resign
    E.V_OnClick(9, 15)
    assert sh.msg().startswith('Hint: '), sh.msg()
    E.V_OnClick(9, 14)
    check_board_cells(sh, flipped=True)
    E.V_OnClick(9, 14)
    E.V_OnClick(8, 15)
    assert 'You resigned' in sh.status(), sh.status()
    sh.click('e2')
    assert 'game is over' in sh.msg()
    return 'ok'


def test_sheet_endings():
    sh = Sheet().install()
    out = []
    # promotion via the picker
    sh.open('8/P6k/8/8/8/8/8/K7 w - - 0 1')
    sh.click('a7')
    sh.click('a8')
    assert sh.cells[(10, 13)] == GLYPH['Q'] and 'promotes' in sh.msg(), sh.msg()
    E.V_OnClick(10, 15)
    assert sh.moves()[0] == 'a7a8b', sh.moves()
    out.append('promotion picker')
    # castling by clicking the king and then the rook
    sh.open('r3k2r/pppppppp/8/8/8/8/PPPPPPPP/R3K2R w KQkq - 0 1')
    sh.click('e1')
    sh.click('h1')
    assert sh.moves()[0] == 'e1g1', sh.moves()
    out.append('castle by clicking the rook')
    # you checkmate the engine
    sh.open('6k1/5ppp/8/8/8/8/5PPP/R5K1 w - - 0 1')
    sh.type('Ra8#')
    assert sh.status() == 'Checkmate - you win! (1-0)', sh.status()
    # the engine checkmates you (it is black and to move when the file opens)
    sh.open('r5k1/8/8/8/8/8/5PPP/6K1 b - - 0 1', level='Hard')
    assert sh.moves() == ['a8a1'] and 'engine wins' in sh.status(), (sh.moves(), sh.status())
    out.append('checkmate both ways')
    sh.open('k7/8/8/2Q5/8/8/8/K7 w - - 0 1')
    sh.type('Qc7')
    assert sh.status().startswith('Stalemate'), sh.status()
    sh.open('k7/8/8/8/8/8/1n6/K7 w - - 0 1')
    sh.type('Kxb2')
    assert 'enough pieces' in sh.status(), sh.status()
    sh.open('k7/8/8/8/8/8/8/KR6 w - - 99 80')
    sh.type('Rb2')
    assert '50-move' in sh.status(), sh.status()
    out.append('stalemate, insufficient material, 50 moves')
    # new game as black: the engine opens, the board is flipped
    sh.open()
    E.V_OnClick(8, 13)
    assert len(sh.moves()) == 1 and sh.data[3] == 'b'
    check_board_cells(sh, flipped=True)
    assert sh.status() == 'Your move (Black)', sh.status()
    # pinned piece explanation through the UI
    sh.open('4k3/4r3/8/8/8/8/4B3/4K3 w - - 0 1')
    sh.click('e2')
    sh.click('d3')
    assert 'pinned' in sh.msg(), sh.msg()
    out.append('play as black, pin message')
    return ', '.join(out)


if __name__ == '__main__':
    t0 = time.time()
    failed = 0
    for name, fn in list(globals().items()):
        if name.startswith('test_') and callable(fn):
            t = time.time()
            try:
                res = fn()
                print(f'PASS {name}: {res}  ({time.time() - t:.1f}s)', flush=True)
            except Exception as e:  # noqa: BLE001
                failed += 1
                import traceback
                traceback.print_exc()
                print(f'FAIL {name}: {e!r}', flush=True)
    print(f'{"ALL PASSED" if not failed else f"{failed} FAILED"} in {time.time() - t0:.0f}s')
    sys.exit(1 if failed else 0)
