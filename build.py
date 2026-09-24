#!/usr/bin/env python3
"""Build Excel Chess.xlsm: the sheet (XlsxWriter) plus a vbaProject.bin written from
scratch, since there is no Excel on Linux to save one. The .bin follows Microsoft's specs:
MS-OVBA (VBA project: dir stream, compression, PROJECT stream) inside an MS-CFB compound
file. Only source code is stored; Excel compiles it the first time the file is opened.

    python3 build.py            -> "Excel Chess.xlsm" next to this script
"""
# aios-run: agent
import os
import random
import struct
import sys
import tempfile

import xlsxwriter

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, 'src')
OUT = os.path.join(HERE, 'Excel Chess.xlsm')
START_FEN = 'rnbqkbnr/pppppppp/8/8/8/8/PPPPPPPP/RNBQKBNR w KQkq - 0 1'

# ============================================================== MS-OVBA 2.4.1 compression

def compress(data):
    out = bytearray(b'\x01')
    for start in range(0, len(data), 4096):
        chunk = data[start:start + 4096]
        comp = _compress_chunk(chunk)
        if len(comp) <= 4096:
            out += struct.pack('<H', 0xB000 | (len(comp) + 2 - 3)) + comp
        else:                      # incompressible: a raw chunk, only allowed at 4096 bytes
            assert len(chunk) == 4096
            out += struct.pack('<H', 0x3FFF) + chunk
    return bytes(out)


def _compress_chunk(chunk):
    out, pos, n, seen = bytearray(), 0, len(chunk), {}

    def remember(p):
        if p + 3 <= n:
            seen.setdefault(chunk[p:p + 3], []).append(p)

    while pos < n:
        flag_at = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if pos >= n:
                break
            bit_count = max((pos - 1).bit_length(), 4) if pos else 4
            max_len = min((0xFFFF >> bit_count) + 3, n - pos)
            best_len = best_off = 0
            for c in reversed(seen.get(chunk[pos:pos + 3], [])[-80:]):
                length = 0
                while length < max_len and chunk[c + length] == chunk[pos + length]:
                    length += 1
                if length > best_len:
                    best_len, best_off = length, pos - c
                    if length == max_len:
                        break
            if best_len >= 3:
                out += struct.pack('<H', ((best_off - 1) << (16 - bit_count)) | (best_len - 3))
                flags |= 1 << bit
                for k in range(best_len):
                    remember(pos + k)
                pos += best_len
            else:
                out.append(chunk[pos])
                remember(pos)
                pos += 1
        out[flag_at] = flags
    return out


def decompress(data):
    """Straight from the spec; used to prove compress() round-trips."""
    assert data[0] == 1
    out, i = bytearray(), 1
    while i < len(data):
        header = struct.unpack_from('<H', data, i)[0]
        size = (header & 0x0FFF) + 3
        assert header & 0x7000 == 0x3000, 'bad chunk signature'
        end, i = i + size, i + 2
        start = len(out)
        if not header & 0x8000:
            out += data[i:i + 4096]
            i += 4096
            continue
        while i < end:
            flags = data[i]
            i += 1
            for bit in range(8):
                if i >= end:
                    break
                if flags >> bit & 1:
                    token = struct.unpack_from('<H', data, i)[0]
                    i += 2
                    diff = len(out) - start
                    bit_count = max((diff - 1).bit_length(), 4)
                    length = (token & (0xFFFF >> bit_count)) + 3
                    offset = (token >> (16 - bit_count)) + 1
                    for _ in range(length):
                        out.append(out[-offset])
                else:
                    out.append(data[i])
                    i += 1
    return bytes(out)

# ============================================================== MS-OVBA 2.4.3 data encryption (CMG / DPB / GC)

def encrypt(project_id, data, rng):
    seed = rng.randrange(256)
    version = 2
    proj_key = sum(project_id.encode('cp1252')) & 0xFF
    out = [seed, seed ^ version, seed ^ proj_key]
    unenc1, enc1, enc2 = proj_key, seed ^ proj_key, seed ^ version
    for _ in range((seed & 6) // 2):
        temp = rng.randrange(256)
        b = temp ^ ((enc2 + unenc1) & 0xFF)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, temp
    for byte in struct.pack('<I', len(data)) + data:
        b = byte ^ ((enc2 + unenc1) & 0xFF)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, byte
    return ''.join(f'{b:02X}' for b in out)


def decrypt(hexstr):
    d = bytes.fromhex(hexstr)
    seed, version_enc, proj_key_enc = d[0], d[1], d[2]
    assert seed ^ version_enc == 2
    proj_key = seed ^ proj_key_enc
    unenc1, enc1, enc2 = proj_key, proj_key_enc, version_enc
    i = 3
    plain = []
    for _ in range((seed & 6) // 2 + 4):
        b = d[i] ^ ((enc2 + unenc1) & 0xFF)
        enc2, enc1, unenc1 = enc1, d[i], b
        plain.append(b)
        i += 1
    length = struct.unpack('<I', bytes(plain[-4:]))[0]
    data = []
    for _ in range(length):
        b = d[i] ^ ((enc2 + unenc1) & 0xFF)
        enc2, enc1, unenc1 = enc1, d[i], b
        data.append(b)
        i += 1
    assert i == len(d)
    return proj_key, bytes(data)

# ============================================================== MS-CFB compound file

ENDOFCHAIN, FREESECT, FATSECT, NOSTREAM = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFF


def cfb(files):
    """files: {'PROJECT': bytes, 'VBA/dir': bytes, ...} -> compound file bytes."""
    # directory tree: root, storages, streams
    entries = [{'name': 'Root Entry', 'type': 5, 'kids': []}]
    index = {'': 0}
    for path in files:
        parts = path.split('/')
        for depth in range(1, len(parts) + 1):
            key = '/'.join(parts[:depth])
            if key not in index:
                index[key] = len(entries)
                is_stream = depth == len(parts)
                entries.append({'name': parts[depth - 1], 'type': 2 if is_stream else 1, 'kids': [],
                                'data': files[path] if is_stream else None})
                entries[index['/'.join(parts[:depth - 1])]]['kids'].append(index[key])
    for e in entries:   # MS-CFB 2.6.3: a storage's start sector and size MUST be zero
        e.update(left=NOSTREAM, right=NOSTREAM, child=NOSTREAM,
                 start=0 if e['type'] == 1 else ENDOFCHAIN, size=0)

    def order(i):
        name = entries[i]['name']
        return (len(name), name.upper())

    def tree(ids):          # balanced binary search tree, all nodes black
        if not ids:
            return NOSTREAM
        mid = len(ids) // 2
        entries[ids[mid]]['left'] = tree(ids[:mid])
        entries[ids[mid]]['right'] = tree(ids[mid + 1:])
        return ids[mid]

    for e in entries:
        if e['kids']:
            e['child'] = tree(sorted(e['kids'], key=order))

    # small streams live in the mini stream (64-byte sectors), big ones in 512-byte sectors
    mini, minifat, big = bytearray(), [], []
    for e in entries:
        if e['type'] != 2:
            continue
        data = e['data']
        e['size'] = len(data)
        if len(data) < 4096:
            first = len(mini) // 64
            count = max(1, -(-len(data) // 64))
            e['start'] = first if data else ENDOFCHAIN
            minifat += [first + k + 1 for k in range(count - 1)] + [ENDOFCHAIN]
            mini += data + b'\0' * (count * 64 - len(data))
        else:
            big.append(e)

    sectors, fat = [], []

    def put(blob):
        first = len(sectors)
        count = -(-len(blob) // 512)
        for k in range(count):
            sectors.append(blob[k * 512:(k + 1) * 512].ljust(512, b'\0'))
        fat.extend([first + k + 1 for k in range(count - 1)] + [ENDOFCHAIN])
        return first

    for e in big:
        e['start'] = put(e['data'])
    root = entries[0]
    if mini:
        root['start'] = put(bytes(mini))
        root['size'] = len(mini)
    minifat_blob = b''.join(struct.pack('<I', v) for v in minifat)
    minifat_blob += struct.pack('<I', FREESECT) * ((-len(minifat)) % 128)
    first_minifat = put(minifat_blob) if minifat else ENDOFCHAIN
    n_minifat = len(minifat_blob) // 512

    dir_blob = bytearray()
    for e in entries:
        name = e['name'].encode('utf-16-le')
        assert len(name) <= 62
        dir_blob += name.ljust(64, b'\0')
        dir_blob += struct.pack('<HBB', len(name) + 2, e['type'], 1)
        dir_blob += struct.pack('<III', e['left'], e['right'], e['child'])
        dir_blob += b'\0' * 16 + struct.pack('<I', 0) + b'\0' * 16
        dir_blob += struct.pack('<IQ', e['start'], e['size'])
    while len(dir_blob) % 512:
        dir_blob += b'\0' * 64 + struct.pack('<HBB', 0, 0, 0)
        dir_blob += struct.pack('<III', NOSTREAM, NOSTREAM, NOSTREAM) + b'\0' * 36 + struct.pack('<IQ', 0, 0)
    first_dir = put(bytes(dir_blob))

    n_fat = 1
    while n_fat * 128 < len(sectors) + n_fat:
        n_fat += 1
    assert n_fat <= 109
    fat_start = len(sectors)
    fat += [FATSECT] * n_fat
    fat += [FREESECT] * (n_fat * 128 - len(fat))
    fat_blob = b''.join(struct.pack('<I', v) for v in fat)
    for k in range(n_fat):
        sectors.append(fat_blob[k * 512:(k + 1) * 512])

    header = bytearray(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1' + b'\0' * 16)
    header += struct.pack('<HHHHH', 0x3E, 3, 0xFFFE, 9, 6) + b'\0' * 6
    header += struct.pack('<IIIIIIIII', 0, n_fat, first_dir, 0, 4096, first_minifat, n_minifat,
                          ENDOFCHAIN, 0)
    difat = [fat_start + k for k in range(n_fat)] + [FREESECT] * (109 - n_fat)
    header += b''.join(struct.pack('<I', v) for v in difat)
    assert len(header) == 512
    return bytes(header) + b''.join(sectors)

# ============================================================== the VBA project

DOC_BASE = {'ThisWorkbook': '0{00020819-0000-0000-C000-000000000046}',
            'Sheet1': '0{00020820-0000-0000-C000-000000000046}',
            'Sheet2': '0{00020820-0000-0000-C000-000000000046}'}
STD_MODULES = ['Engine', 'Game', 'XL']
PROJECT_ID = '{00000000-0000-0000-0000-000000000000}'


def module_source(name):
    if name in DOC_BASE:
        text = open(os.path.join(SRC, name + '.cls')).read()
        head = [f'Attribute VB_Name = "{name}"', f'Attribute VB_Base = "{DOC_BASE[name]}"',
                'Attribute VB_GlobalNameSpace = False', 'Attribute VB_Creatable = False',
                'Attribute VB_PredeclaredId = True', 'Attribute VB_Exposed = True',
                'Attribute VB_TemplateDerived = False', 'Attribute VB_Customizable = True']
        text = '\n'.join(head) + '\n' + text
    else:
        text = open(os.path.join(SRC, name + '.bas')).read()
        assert text.startswith(f'Attribute VB_Name = "{name}"')
    lines = text.replace('\r\n', '\n').rstrip('\n').split('\n')
    return ('\r\n'.join(lines) + '\r\n').encode('cp1252')     # raises if anything isn't cp1252


def rec(rid, data):
    return struct.pack('<HI', rid, len(data)) + data


def dir_stream(modules):
    u = lambda s: s.encode('utf-16-le')                      # noqa: E731
    a = lambda s: s.encode('cp1252')                         # noqa: E731
    d = bytearray()
    d += rec(0x0001, struct.pack('<I', 1))                    # SysKind: Win32 (Office recompiles anyway)
    d += rec(0x0002, struct.pack('<I', 0x409))                # LCID
    d += rec(0x0014, struct.pack('<I', 0x409))                # LCID for Invoke
    d += rec(0x0003, struct.pack('<H', 1252))                 # code page
    d += rec(0x0004, a('VBAProject'))                         # project name
    d += rec(0x0005, b'') + rec(0x0040, b'')                  # doc string (ANSI + Unicode)
    d += rec(0x0006, b'') + rec(0x003D, b'')                  # help file (x2)
    d += rec(0x0007, struct.pack('<I', 0))                    # help context
    d += rec(0x0008, struct.pack('<I', 0))                    # lib flags
    d += struct.pack('<HIIH', 0x0009, 4, 1, 0)                # version: "size" 4, then 6 bytes
    d += rec(0x000C, b'') + rec(0x003C, b'')                  # constants (x2)
    libid = a('*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#'
              'C:\\Windows\\System32\\stdole2.tlb#OLE Automation')
    d += rec(0x0016, a('stdole')) + rec(0x003E, u('stdole'))  # reference name
    d += rec(0x000D, struct.pack('<I', len(libid)) + libid + struct.pack('<IH', 0, 0))
    d += struct.pack('<HIH', 0x000F, 2, len(modules))        # module count
    d += struct.pack('<HIH', 0x0013, 2, 0xFFFF)              # project cookie
    for name in modules:
        d += rec(0x0019, a(name)) + rec(0x0047, u(name))
        d += rec(0x001A, a(name)) + rec(0x0032, u(name))      # stream name
        d += rec(0x001C, b'') + rec(0x0048, b'')              # doc string
        d += rec(0x0031, struct.pack('<I', 0))                # source starts at offset 0
        d += rec(0x001E, struct.pack('<I', 0))                # help context
        d += rec(0x002C, struct.pack('<H', 0xFFFF))           # cookie
        d += struct.pack('<HI', 0x0022 if name in DOC_BASE else 0x0021, 0)
        d += struct.pack('<HI', 0x002B, 0)                    # end of module
    d += struct.pack('<HI', 0x0010, 0)                        # end of dir
    return bytes(d)


def project_stream(modules, rng):
    not_protected, no_password, visible = struct.pack('<I', 0), b'\x00', b'\xff'
    lines = [f'ID="{PROJECT_ID}"']
    lines += [f'Document={m}/&H00000000' for m in modules if m in DOC_BASE]
    lines += [f'Module={m}' for m in modules if m not in DOC_BASE]
    lines += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
              f'CMG="{encrypt(PROJECT_ID, not_protected, rng)}"',
              f'DPB="{encrypt(PROJECT_ID, no_password, rng)}"',
              f'GC="{encrypt(PROJECT_ID, visible, rng)}"',
              '', '[Host Extender Info]',
              '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000',
              '', '[Workspace]']
    lines += [f'{m}=0, 0, 0, 0, C' for m in modules]
    return ('\r\n'.join(lines) + '\r\n').encode('cp1252')


def vba_project():
    rng = random.Random(2026)
    modules = ['ThisWorkbook', 'Sheet1', 'Sheet2'] + STD_MODULES
    files = {'PROJECT': project_stream(modules, rng)}
    wm = b''
    for m in modules:
        wm += m.encode('cp1252') + b'\0' + m.encode('utf-16-le') + b'\0\0'
    files['PROJECTwm'] = wm + b'\0\0'
    files['VBA/_VBA_PROJECT'] = b'\xCC\x61\xFF\xFF\x00\x00\x00'
    dirs = dir_stream(modules)
    files['VBA/dir'] = compress(dirs)
    assert decompress(files['VBA/dir']) == dirs
    for m in modules:
        src = module_source(m)
        files[f'VBA/{m}'] = compress(src)
        assert decompress(files[f'VBA/{m}']) == src, m
    return cfb(files)

# ============================================================== the workbook

GLYPH = dict(zip('KQRBNPkqrbnp', [chr(9812 + i) for i in range(12)]))
LIGHT, DARK = '#F0D9B5', '#B58863'


def build_xlsm(path):
    fd, binpath = tempfile.mkstemp(suffix='.bin')
    os.write(fd, vba_project())
    os.close(fd)
    wb = xlsxwriter.Workbook(path)
    wb.set_properties({'title': 'Excel Chess', 'comments': 'A chess engine written in VBA'})
    wb.set_vba_name('ThisWorkbook')
    ws = wb.add_worksheet('Chess')
    ws.set_vba_name('Sheet1')
    data = wb.add_worksheet('Data')
    data.set_vba_name('Sheet2')
    wb.add_vba_project(binpath)

    F = {}

    def fmt(**kw):
        key = tuple(sorted(kw.items()))
        if key not in F:
            F[key] = wb.add_format(dict(kw, num_format='@'))
        return F[key]

    base = dict(font_name='Segoe UI', font_size=11, valign='vcenter')
    ws.hide_gridlines(2)
    ws.set_zoom(100)
    ws.set_column('A:A', 2, fmt(**base))
    ws.set_column('B:B', 3.5, fmt(**base))
    ws.set_column('C:J', 7.6, fmt(**base))
    ws.set_column('K:K', 3, fmt(**base))
    ws.set_column('L:L', 20, fmt(**base))
    ws.set_column('M:P', 17, fmt(**base))
    ws.set_row(0, 8)
    ws.set_row(1, 34)
    for r in range(2, 10):
        ws.set_row(r, 44)
    ws.set_row(10, 22)
    ws.set_row(11, 22)

    ws.write('C2', 'Excel Chess', fmt(**dict(base, bold=True, font_size=20, font_color='#3B2A1A')))
    ws.write('L2', 'A real chess engine, written in VBA, living in this sheet.',
             fmt(**dict(base, italic=True, font_color='#7A7A7A', font_size=10)))

    # board with the starting position (drawn again by the macros once enabled)
    rows = START_FEN.split()[0].split('/')
    for i, row in enumerate(rows):
        cells = []
        for ch in row:
            cells += [''] * int(ch) if ch.isdigit() else [GLYPH[ch]]
        for j, g in enumerate(cells):
            dark = (i + j) % 2 == 1
            border = dict(top=5 if i == 0 else 0, bottom=5 if i == 7 else 0,
                          left=5 if j == 0 else 0, right=5 if j == 7 else 0)
            f = fmt(font_name='Segoe UI Symbol', font_size=30, align='center', valign='vcenter',
                    bg_color=DARK if dark else LIGHT, **border)
            ws.write_string(2 + i, 2 + j, g, f) if g else ws.write_blank(2 + i, 2 + j, None, f)
        ws.write_string(2 + i, 1, str(8 - i), fmt(**dict(base, bold=True, align='center', font_color='#8A7A6A')))
    for j in range(8):
        ws.write_string(10, 2 + j, 'abcdefgh'[j], fmt(**dict(base, bold=True, align='center', font_color='#8A7A6A')))

    status = fmt(**dict(base, bold=True, font_size=15, font_color='#1F1F1F'))
    ws.write('L3', 'Enable macros to play: click "Enable Content" in the yellow bar at the top.', status)
    ws.write('L4', '(If Excel says macros are blocked: close the file, right-click it > Properties > tick "Unblock".)',
             fmt(**dict(base, italic=True, font_size=10, font_color='#6B6B6B')))
    ws.write_blank('L5', None, fmt(**dict(base, bold=True, font_color='#B42318')))
    ws.write('L6', 'Your move:', fmt(**dict(base, bold=True, align='right')))
    ws.write_blank('M6', None, fmt(**dict(base, bold=True, font_size=14, align='center', bg_color='#FFF2A8',
                                           border=2, border_color='#C9A400')))
    ws.write('L7', 'Type e2e4, Nf3, exd5 or O-O in the yellow box and press Enter',
             fmt(**dict(base, italic=True, font_size=9, font_color='#7A7A7A')))
    button = dict(base, bold=True, align='center', font_color='#FFFFFF', border=1, border_color='#FFFFFF')
    ws.write('L8', 'New game as White', fmt(**dict(button, bg_color='#3A4A5A')))
    ws.write('M8', 'New game as Black', fmt(**dict(button, bg_color='#3A4A5A')))
    ws.write('N8', 'Take back', fmt(**dict(button, bg_color='#6B5B3E')))
    ws.write('O8', 'Resign', fmt(**dict(button, bg_color='#8B2E2E')))
    ws.write('L9', 'Level:', fmt(**dict(base, bold=True, align='right')))
    ws.write('M9', 'Hard', fmt(**dict(base, bold=True, align='center', bg_color='#DCE8F5', border=1,
                                      border_color='#7A99BB')))
    ws.data_validation('M9', {'validate': 'list', 'source': ['Easy', 'Medium', 'Hard', 'Insane'],
                              'input_title': 'Engine strength',
                              'input_message': 'Easy / Medium / Hard (3 s) / Insane (10 s)'})
    ws.write('N9', 'Flip board', fmt(**dict(button, bg_color='#3A4A5A')))
    ws.write('O9', 'Hint', fmt(**dict(button, bg_color='#2E7D4F')))
    ws.write_blank('L10', None, fmt(**dict(base, bold=True, align='right')))
    for c in 'MNOP':
        ws.write_blank(f'{c}10', None, fmt(font_name='Segoe UI Symbol', font_size=26, align='center',
                                            valign='vcenter', bg_color='#FFF7D6'))
    took = fmt(**dict(base, font_name='Segoe UI Symbol', font_size=12))
    ws.write('L11', 'White took: ', took)
    ws.write('L12', 'Black took: ', took)
    head = fmt(**dict(base, bold=True, bottom=2))
    ws.write('L14', 'Move', head)
    ws.write('M14', 'White', head)
    ws.write('N14', 'Black', head)

    help_title = fmt(**dict(base, bold=True, font_size=12))
    help_text = fmt(**dict(base, font_size=10, font_color='#4A4A4A'))
    ws.write('B14', 'How to play', help_title)
    tips = [
        'Click your piece, then the square it goes to. Blue = legal squares.',
        'Or drag it: grab the edge of the cell and drop it on the new square.',
        'Or type the move in the yellow box: e2e4, Nf3, exd5, O-O, e8=Q.',
        'Castle: click the king, then the rook (or two squares over).',
        'Pawn on the last row: pick the new piece in the row that pops up.',
        'Illegal moves are refused, and it tells you why (pin, check ...).',
        'All rules: check, mate, stalemate, castling, en passant,',
        'promotion, 50-move rule, threefold repetition, dead positions.',
        'Levels: Easy, Medium, Hard (3 s), Insane (10 s). Hint = help.',
        'The game is saved in the file - close it, carry on later.',
    ]
    for k, t in enumerate(tips):
        ws.write(14 + k, 1, t, help_text)
    ws.set_selection('M6')
    ws.activate()

    dfmt = wb.add_format({'num_format': '@'})
    data.set_column('A:A', 14)
    data.set_column('B:B', 80, dfmt)
    labels = ['start position', 'moves', 'you play (w/b)', 'board flipped', 'result', 'engine info']
    values = [START_FEN, '', 'w', '', '', '']
    for k, (lab, val) in enumerate(zip(labels, values)):
        data.write_string(k, 0, lab)
        data.write_string(k, 1, val, dfmt) if val else data.write_blank(k, 1, None, dfmt)
    data.hide()
    wb.close()
    os.remove(binpath)
    return path


if __name__ == '__main__':
    out = build_xlsm(sys.argv[1] if len(sys.argv) > 1 else OUT)
    print('built', out, os.path.getsize(out), 'bytes')
