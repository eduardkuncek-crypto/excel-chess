#!/usr/bin/env python3
"""Measure Excel Chess's strength: the real VBA engine (translated by vba2py) plays a match
against Stockfish locked to a fixed Elo (UCI_LimitStrength). Each opening is played twice,
once with each colour. Prints the score and the Elo it implies.

    python3 elo_match.py <stockfish binary> --depth 2 --elo 1320 --openings 12
Needs: python-chess.
"""
# aios-run: agent
import argparse
import math
import os
import sys
from multiprocessing import Pool

import chess
import chess.engine

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)


def play(job):
    import test_chess                       # translates the .bas files, once per worker process
    E = test_chess.E
    sf_path, elo, depth, opening, engine_white = job
    board = chess.Board()
    E.V_SetFEN(board.fen())
    for u in opening:
        board.push_uci(u)
        test_chess.play_uci(u)
    sf = chess.engine.SimpleEngine.popen_uci(sf_path)
    sf.configure({'UCI_LimitStrength': True, 'UCI_Elo': elo, 'Threads': 1})
    try:
        while not board.is_game_over(claim_draw=True) and board.ply() < 300:
            if (board.turn == chess.WHITE) == engine_white:
                E.V_EvalNoise = 0
                u = E.V_MoveUci(E.V_Think(depth, 10 ** 8))
            else:
                u = sf.play(board, chess.engine.Limit(time=0.1)).move.uci()
            board.push_uci(u)
            test_chess.play_uci(u)
    finally:
        sf.quit()
    result = board.result(claim_draw=True)
    if result == '1/2-1/2' or result == '*':
        return 0.5, result
    return (1.0 if (result == '1-0') == engine_white else 0.0), result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('stockfish')
    ap.add_argument('--depth', type=int, default=2)
    ap.add_argument('--elo', type=int, default=1320)
    ap.add_argument('--openings', type=int, default=12)
    ap.add_argument('--workers', type=int, default=4)
    a = ap.parse_args()
    import re
    book = re.findall(r'"\|?([a-h1-8 ]{9,})"', open(os.path.join(HERE, 'src', 'Engine.bas')).read())
    openings = []
    for line in book:
        o = tuple(line.split()[:4])
        if len(o) == 4 and o not in openings:
            openings.append(o)
    openings = openings[:a.openings]
    jobs = [(a.stockfish, a.elo, a.depth, o, w) for o in openings for w in (True, False)]
    with Pool(a.workers) as pool:
        res = pool.map(play, jobs)
    score = sum(r[0] for r in res)
    n = len(res)
    wins = sum(r[0] == 1 for r in res)
    draws = sum(r[0] == 0.5 for r in res)
    s = min(max(score / n, 0.5 / n), 1 - 0.5 / n)          # keep log finite at 0% / 100%
    diff = -400 * math.log10(1 / s - 1)
    print(f'depth {a.depth} vs Stockfish {a.elo}: +{wins} ={draws} -{n - wins - draws}  '
          f'score {score}/{n}  ->  about {a.elo + diff:.0f} Elo ({diff:+.0f})')


if __name__ == '__main__':
    main()
