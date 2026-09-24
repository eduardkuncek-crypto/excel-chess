#!/usr/bin/env python3
"""End-to-end test of Excel Chess.xlsm inside a real LibreOffice Calc (headless, private
profile): opens it with macros on, then clicks squares, types moves, drags a piece and
presses the buttons, the way a player would. Every board it shows is checked against
python-chess. This is how the game runs on the Arch laptop.

    python3 lo_test.py                 (builds a fresh copy of the .xlsm in /tmp first)
Needs: LibreOffice (pacman -S libreoffice-fresh), python-chess importable.
"""
# aios-run: agent
import os
import subprocess
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build  # noqa: E402

try:
    import chess
except ImportError:          # python-chess lives in a venv; borrow it
    import glob
    sys.path += glob.glob('/tmp/claude-*/**/venv/lib/python3*/site-packages', recursive=True)
    import chess

GLYPH = dict(zip('KQRBNPkqrbnp', [chr(9812 + i) for i in range(12)]))


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def start_office(profile):
    """A private headless LibreOffice on port 2002. Never touches a LibreOffice he has open."""
    global OFFICE
    OFFICE = subprocess.Popen(['soffice', '--headless', '--invisible', '--norestore', '--nologo',
                      f'-env:UserInstallation=file://{profile}',
                      '--accept=socket,host=127.0.0.1,port=2002;urp;'],
                     env=dict(os.environ, SAL_DISABLE_OPENCL='1'),
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    for _ in range(80):
        try:
            return resolver.resolve('uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext')
        except Exception:  # noqa: BLE001 - still starting
            time.sleep(0.25)
    raise SystemExit('LibreOffice did not start')


class Game:
    def __init__(self, doc, ctx):
        self.doc = doc
        self.dispatcher = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)
        self.ch = doc.Sheets.getByName('Chess')
        self.data = doc.Sheets.getByName('Data')

    def text(self, a):
        return self.ch.getCellRangeByName(a).getString()

    def moves(self):
        return self.data.getCellRangeByName('B2').getString().split()

    def cell(self, sq, flipped=False):
        f, r = ord(sq[0]) - 97, int(sq[1])
        col, row = (7 - f, r - 1) if flipped else (f, 8 - r)
        return 'CDEFGHIJ'[col] + str(3 + row)

    def click(self, a, wait=0.4):
        self.doc.CurrentController.select(self.ch.getCellRangeByName(a))
        time.sleep(wait)

    def square(self, sq, flipped=False, wait=0.4):
        self.click(self.cell(sq, flipped), wait)

    def ui(self, cmd, *args):
        frame = self.doc.CurrentController.Frame
        disp = self.dispatcher
        disp.executeDispatch(frame, cmd, '', 0, tuple(prop(k, v) for k, v in args))

    def type(self, txt, wait=1.0):
        """Type into the move box the way a person does (a UI edit, so the sheet event fires)."""
        self.click('M6', wait=0.2)
        self.ui('.uno:EnterString', ('StringName', txt))
        time.sleep(wait)

    def drag(self, src, dst, wait=1.5):
        """Cut + paste = what a cell drag does: source emptied, piece lands on the target."""
        self.square(src, wait=0.3)
        self.ui('.uno:Cut')
        self.square(dst, wait=0.3)
        self.ui('.uno:Paste')
        time.sleep(wait)

    def board(self):
        b = chess.Board(self.data.getCellRangeByName('B1').getString() or chess.STARTING_FEN)
        for u in self.moves():
            b.push_uci(u)
        return b

    def check_board(self, flipped=False):
        b = self.board()
        for s in chess.SQUARES:
            p = b.piece_at(s)
            want = GLYPH[p.symbol()] if p else ''
            got = self.text(self.cell(chess.square_name(s), flipped))
            assert got == want, (chess.square_name(s), got, want)
        return b


def main():
    tmp = tempfile.mkdtemp()
    path = build.build_xlsm(os.path.join(tmp, 'chess.xlsm'))
    ctx = start_office(os.path.join(tmp, 'profile'))
    desk = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
    doc = desk.loadComponentFromURL(uno.systemPathToFileUrl(path), '_blank', 0,
                                    (prop('MacroExecutionMode', 4),))
    g = Game(doc, ctx)
    ok = []
    assert g.text('L3') == 'Your move (White)', g.text('L3')
    g.check_board()
    ok.append('opens with macros')
    g.click('M9', wait=0.2)
    g.ui('.uno:EnterString', ('StringName', 'Easy'))
    g.square('e2')
    g.square('e4', wait=1.5)
    assert g.moves()[0] == 'e2e4' and len(g.moves()) == 2, g.moves()
    g.check_board()
    ok.append(f'click-click e4, engine replied {g.moves()[1]}')
    g.square('b1')
    g.square('b4')
    assert "knight can't move like that" in g.text('L5'), g.text('L5')
    ok.append('illegal move explained')
    b = g.board()
    san = next(b.san(m) for m in b.legal_moves if b.piece_at(m.from_square).piece_type == chess.KNIGHT)
    g.type(san, wait=1.5)
    assert len(g.moves()) == 4, (san, g.moves(), g.text('L5'))
    assert g.text('M6') == ''
    g.check_board()
    ok.append(f'typed {san}')
    # drag: select the piece, then the sheet shows it moved to the target
    b = g.board()
    mv = next(m for m in b.legal_moves if not m.promotion and not b.is_castling(m) and not b.is_capture(m))
    src, dst = chess.square_name(mv.from_square), chess.square_name(mv.to_square)
    g.drag(src, dst)
    assert g.moves()[4] == mv.uci(), (g.moves(), mv.uci(), g.text('L5'))
    g.check_board()
    ok.append(f'drag {src}-{dst}')
    g.click('N8', wait=1)                           # take back
    assert len(g.moves()) == 4, g.moves()
    g.check_board()
    g.click('N9')                                   # flip
    g.check_board(flipped=True)
    assert doc.CurrentController.Selection.AbsoluteName.endswith('$M$6'), 'cursor not parked'
    g.click('N9')
    g.check_board()
    ok.append('take back + flip')
    g.click('M9', wait=0.2)
    g.ui('.uno:EnterString', ('StringName', 'Hard'))
    t = time.time()
    g.click('O9', wait=0.5)                         # hint
    while not g.text('L5').startswith('Hint') and time.time() - t < 20:
        time.sleep(0.3)
    assert g.text('L5').startswith('Hint:'), g.text('L5')
    ok.append(f'{g.text("L5")} in {time.time() - t:.1f}s')
    b = g.board()
    mv = next(iter(b.legal_moves))
    g.type(mv.uci(), wait=0.5)
    t = time.time()
    while len(g.moves()) < 6 and time.time() - t < 30:
        time.sleep(0.3)
    assert len(g.moves()) == 6, g.moves()
    g.check_board()
    ok.append(f'Hard level: {g.text("L4")}  (wall {time.time() - t:.1f}s)')
    g.click('M8', wait=1.5)                         # new game as black
    assert len(g.moves()) == 1 and g.text('L3') == 'Your move (Black)', (g.moves(), g.text('L3'))
    g.check_board(flipped=True)
    ok.append('new game as black')
    doc.close(True)
    try:
        desk.terminate()           # only our own headless instance
    except Exception:  # noqa: BLE001 - the bridge drops as it exits
        pass
    return ok


if __name__ == '__main__':
    for line in main():
        print('PASS', line)
    print('ALL PASSED')
