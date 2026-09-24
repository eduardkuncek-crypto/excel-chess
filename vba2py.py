#!/usr/bin/env python3
"""Translate the small VBA subset Excel Chess is written in into Python, so the real
.bas files can be run and tested on Linux. It doubles as a linter: in Excel a compile
error kills the whole workbook and nothing here can open Excel, so anything Excel would
reject (undeclared names, wrong argument counts, unbalanced blocks, type mismatches,
16-bit Integer overflow traps, writes to ByRef parameters) is an error here too.
"""
# aios-run: agent
import re

KEYWORDS = set('''and or xor not mod is then else elseif end if for to step next do loop
while until wend select case exit sub function dim private public const as byval byref
call true false set with new nothing option explicit attribute goto on error resume
global static redim preserve'''.split())
TYPES = {'long': 'Long', 'double': 'Double', 'single': 'Double', 'boolean': 'Boolean',
         'string': 'String'}
DEFAULT = {'Long': '0', 'Double': '0.0', 'Boolean': 'False', 'String': "''"}

TOK = re.compile(r'\s*(?:(?P<str>"(?:[^"]|"")*")|(?P<hex>&H[0-9A-Fa-f]+&?)'
                 r'|(?P<num>\d+\.\d+#?|\d+#?)|(?P<id>[A-Za-z][A-Za-z0-9_]*\$?)'
                 r'|(?P<op><>|<=|>=|:=|[-+*/\\^&=<>(),.:]))')


class VBAError(Exception):
    pass


def strip_comment(line):
    q = False
    for i, ch in enumerate(line):
        if ch == '"':
            q = not q
        elif ch == "'" and not q:
            return line[:i]
    return line


def tokenize(s, where):
    toks, pos, s = [], 0, s.rstrip()
    while pos < len(s):
        m = TOK.match(s, pos)
        if not m or m.end() == pos:
            raise VBAError(f'{where}: cannot read {s[pos:]!r}')
        pos = m.end()
        toks.append((m.lastgroup, m.group(m.lastgroup)))
    return toks


def low(tok):
    return tok[1].lower().rstrip('$') if tok[0] == 'id' else tok[1]


def split_top(toks, sep):
    """Split a token list on a separator (op text or keyword) at paren depth 0."""
    parts, cur, depth = [], [], 0
    for t in toks:
        if t[1] == '(':
            depth += 1
        elif t[1] == ')':
            depth -= 1
        if depth == 0 and low(t) == sep:
            parts.append(cur)
            cur = []
        else:
            cur.append(t)
    parts.append(cur)
    return parts


def match_paren(toks, i):
    """Index of the ')' matching the '(' at toks[i]."""
    depth = 0
    for j in range(i, len(toks)):
        if toks[j][1] == '(':
            depth += 1
        elif toks[j][1] == ')':
            depth -= 1
            if depth == 0:
                return j
    raise VBAError(f'missing ) in {" ".join(t[1] for t in toks)}')


def find_top(toks, sep):
    depth = 0
    for i, t in enumerate(toks):
        if t[1] == '(':
            depth += 1
        elif t[1] == ')':
            depth -= 1
        elif depth == 0 and low(t) == sep:
            return i
    return -1


class Sym:
    def __init__(self, kind, name, py, typ, module, public, dims=None, params=None):
        self.kind, self.name, self.py, self.typ = kind, name, py, typ
        self.module, self.public, self.dims, self.params = module, public, dims, params
        self.assigned_params = set()


class Node:
    def __init__(self, code, typ, lit=False, arith=False):
        self.code, self.typ, self.lit, self.arith = code, typ, lit, arith


BUILTINS = {
    # name: (min args, max args, python, return type; 'arg0' = same as first arg)
    'abs': (1, 1, 'abs({0})', 'arg0'), 'mid': (2, 3, 'vmid({a})', 'String'),
    'left': (2, 2, 'vleft({a})', 'String'), 'right': (2, 2, 'vright({a})', 'String'),
    'len': (1, 1, 'len({0})', 'Long'), 'instr': (2, 3, 'vinstr({a})', 'Long'),
    'lcase': (1, 1, '({0}).lower()', 'String'), 'ucase': (1, 1, '({0}).upper()', 'String'),
    'trim': (1, 1, 'vtrim({0})', 'String'), 'chr': (1, 1, 'vchr({0})', 'String'),
    'chrw': (1, 1, 'vchr({0})', 'String'), 'asc': (1, 1, 'ord(({0})[0])', 'Long'),
    'cstr': (1, 1, 'vstr({0})', 'String'), 'clng': (1, 1, 'vclng({0})', 'Long'),
    'split': (2, 2, 'vsplit({a})', 'Array'), 'ubound': (1, 1, 'vubound({0})', 'Long'),
    'replace': (3, 3, 'vreplace({a})', 'String'), 'int': (1, 1, 'vint({0})', 'Double'),
    'rnd': (0, 0, 'vrnd()', 'Double'), 'timer': (0, 0, 'vtimer()', 'Double'),
    'rgb': (3, 3, 'vrgb({a})', 'Long'),
}
STRING_ARGS = {'mid': [0], 'left': [0], 'right': [0], 'len': [0], 'lcase': [0], 'ucase': [0],
               'trim': [0], 'asc': [0], 'split': [0, 1], 'replace': [0, 1, 2]}


class Translator:
    def __init__(self):
        self.syms = {}        # lower name -> list of module-level Sym
        self.modules = []     # (name, header lines, procs)
        self.out = []

    # ------------------------------------------------------------------ pass 1: declarations
    def add_module(self, name, source, bodies=True):
        lines, buf = [], ''
        for n, raw in enumerate(source.splitlines(), 1):
            text = strip_comment(raw).rstrip()
            if text.endswith(' _'):
                buf += text[:-2] + ' '
                continue
            text = (buf + text).strip()
            buf = ''
            if text:
                lines.append((n, text))
        decls, procs, cur = [], [], None
        for n, text in lines:
            where = f'{name}.bas:{n}'
            toks = tokenize(text, where)
            if ':' in [t[1] for t in toks if t[0] == 'op']:
                raise VBAError(f'{where}: one statement per line please (":" found)')
            words = [low(t) for t in toks]
            if words[0] in ('attribute', 'option'):
                continue
            if cur is None:
                w = words[1] if words[0] in ('public', 'private') else words[0]
                if w in ('sub', 'function'):
                    cur = {'hdr': (where, toks), 'body': []}
                    continue
                if procs:
                    raise VBAError(f'{where}: declarations must come before all procedures')
                decls.append((where, toks))
            else:
                if words[:2] in (['end', 'sub'], ['end', 'function']):
                    kind = 'sub' if 'sub' in [low(t) for t in cur['hdr'][1][:2]] else 'function'
                    if words[1] != kind:
                        raise VBAError(f'{where}: End {words[1].title()} closes a {kind.title()}')
                    procs.append(cur)
                    cur = None
                else:
                    cur['body'].append((where, toks))
        if cur is not None:
            raise VBAError(f'{name}: procedure {cur["hdr"][0]} never ends')
        for where, toks in decls:
            self.declare_module_level(name, where, toks)
        for p in procs:
            words = [low(t) for t in p['hdr'][1][:1]]
            if not bodies and words == ['private']:
                continue          # extern module: only its Public procedures are visible
            self.declare_proc(name, p)
        if bodies:
            self.modules.append((name, decls, procs))

    def pyname(self, module, name, public):
        return f'V_{name}' if public else f'V_{module}_{name}'

    def add_sym(self, sym, where):
        key = sym.name.lower()
        for other in self.syms.get(key, []):
            if other.public or sym.public or other.module == sym.module:
                raise VBAError(f'{where}: {sym.name} is declared twice')
        self.syms.setdefault(key, []).append(sym)

    def parse_type(self, toks, where):
        if len(toks) != 2 or low(toks[0]) != 'as':
            raise VBAError(f'{where}: every declaration needs "As <type>" (got {toks})')
        t = low(toks[1])
        if t == 'integer':
            raise VBAError(f'{where}: no Integer, it overflows at 32767 - use Long')
        if t not in TYPES:
            raise VBAError(f'{where}: unsupported type {toks[1][1]}')
        return TYPES[t]

    def declare_module_level(self, module, where, toks):
        words = [low(t) for t in toks]
        public = words[0] in ('public', 'global')
        if words[0] not in ('public', 'private', 'dim', 'global', 'const'):
            raise VBAError(f'{where}: unexpected module-level statement')
        rest = toks[1:] if words[0] != 'const' else toks
        if low(rest[0]) == 'const':
            eq = find_top(rest, '=')
            name = rest[1][1]
            typ = self.parse_type(rest[2:eq], where)
            sym = Sym('const', name, self.pyname(module, name, public), typ, module, public)
            sym.init = rest[eq + 1:]
            self.add_sym(sym, where)
            return
        for part in split_top(rest, ','):
            self.add_sym(self.parse_var(part, where, module, public), where)

    def parse_var(self, part, where, module, public, local=False):
        name = part[0][1]
        if part[0][0] != 'id' or low(part[0]) in KEYWORDS:
            raise VBAError(f'{where}: bad variable name {name}')
        py = f'V_{name}' if local else self.pyname(module, name, public)
        if len(part) > 1 and part[1][1] == '(':
            close = match_paren(part, 1)
            inner = part[2:close]
            typ = self.parse_type(part[close + 1:], where)
            if not inner:
                sym = Sym('dynarray', name, py, typ, module, public)
            else:
                dims = []
                for d in split_top(inner, ','):
                    k = find_top(d, 'to')
                    if k < 0:
                        raise VBAError(f'{where}: write array bounds as "lo To hi"')
                    dims.append((d[:k], d[k + 1:]))
                sym = Sym('array', name, py, typ, module, public, dims=dims)
            return sym
        return Sym('var', name, py, self.parse_type(part[1:], where), module, public)

    def declare_proc(self, module, p):
        where, toks = p['hdr']
        words = [low(t) for t in toks]
        public = words[0] != 'private'
        i = 1 if words[0] in ('public', 'private') else 0
        kind = words[i]
        name = toks[i + 1][1]
        op = i + 2
        if toks[op][1] != '(':
            raise VBAError(f'{where}: procedure header needs ()')
        close = match_paren(toks, op)
        params = []
        inner = toks[op + 1:close]
        if inner:
            for part in split_top(inner, ','):
                byval = False
                if low(part[0]) in ('byval', 'byref'):
                    byval = low(part[0]) == 'byval'
                    part = part[1:]
                ps = self.parse_var(part, where, module, True, local=True)
                ps.kind = 'param' if ps.kind == 'var' else ps.kind
                ps.byval = byval
                params.append(ps)
        rtype = None
        if kind == 'function':
            rtype = self.parse_type(toks[close + 1:], where)
        elif close + 1 != len(toks):
            raise VBAError(f'{where}: a Sub has no return type')
        sym = Sym(kind, name, self.pyname(module, name, public), rtype, module, public,
                  params=params)
        self.add_sym(sym, where)
        p['sym'] = sym
        p['module'] = module

    # ------------------------------------------------------------------ name lookup
    def lookup(self, name, module):
        for s in self.syms.get(name.lower(), []):
            if s.public or s.module == module:
                return s
        return None

    # ------------------------------------------------------------------ pass 2: emit
    def translate(self):
        out = ['# generated by vba2py.py - do not edit', 'from vbart import *', '']
        # constants first, in order (they may refer to each other), then variables
        for module, decls, procs in self.modules:
            self.ctx = Ctx(self, module, None)
            for where, toks in decls:
                self.ctx.where = where
                for s in self.decl_syms(module, toks):
                    if s.kind == 'const':
                        e = self.ctx.expr(s.init)
                        self.check_assign(s.typ, e, where)
                        out.append(f'{s.py} = {e.code}')
        for module, decls, procs in self.modules:
            self.ctx = Ctx(self, module, None)
            for where, toks in decls:
                self.ctx.where = where
                for s in self.decl_syms(module, toks):
                    if s.kind != 'const':
                        out.append(f'{s.py} = {self.ctx.init_code(s)}')
        for module, decls, procs in self.modules:
            for p in procs:
                out.extend(ProcEmitter(self, module, p).emit())
                out.append('')
        return '\n'.join(out) + '\n'

    def decl_syms(self, module, toks):
        rest = toks[1:] if low(toks[0]) != 'const' else toks
        if low(rest[0]) == 'const':
            return [self.lookup(rest[1][1], module)]
        return [self.lookup(part[0][1], module) for part in split_top(rest, ',')]

    def check_assign(self, typ, e, where):
        if typ == 'Long':
            if e.typ not in ('Long', 'Double'):
                raise VBAError(f'{where}: assigning {e.typ} to a Long')
            return f'vlong({e.code})' if (e.typ == 'Double' or e.arith) else e.code
        if typ == 'Double':
            if e.typ not in ('Long', 'Double'):
                raise VBAError(f'{where}: assigning {e.typ} to a Double')
            return f'float({e.code})'
        if typ != e.typ:
            raise VBAError(f'{where}: assigning {e.typ} to a {typ}')
        return e.code


class Ctx:
    """Expression parser for one scope."""

    def __init__(self, tr, module, proc):
        self.tr, self.module, self.proc = tr, module, proc
        self.locals = {}
        self.where = module

    def init_code(self, s):
        if s.kind == 'dynarray':
            return 'VList()'
        if s.kind == 'array':
            bounds = []
            for lo, hi in s.dims:
                a, b = self.expr(lo), self.expr(hi)
                bounds.append(f'{a.code}, {b.code}')
            cls = 'VArr' if len(bounds) == 1 else 'VArr2'
            return f'{cls}({DEFAULT[s.typ]}, {", ".join(bounds)})'
        return DEFAULT[s.typ]

    def resolve(self, name):
        s = self.locals.get(name.lower())
        return s if s else self.tr.lookup(name, self.module)

    def err(self, msg):
        raise VBAError(f'{self.where}: {msg}')

    # precedence climbing, VBA order (low -> high):
    # Xor, Or, And, Not, comparisons, &, + -, Mod, \, * /, unary -, ^
    def expr(self, toks):
        self.toks, self.i = toks, 0
        if not toks:
            self.err('missing expression')
        n = self.p_bin(0)
        if self.i != len(toks):
            self.err(f'unexpected {toks[self.i][1]!r}')
        return n

    LEVELS = [['xor'], ['or'], ['and'], None, ['=', '<>', '<', '>', '<=', '>='], ['&'],
              ['+', '-'], ['mod'], ['\\'], ['*', '/']]

    def peek(self):
        return low(self.toks[self.i]) if self.i < len(self.toks) else None

    def p_bin(self, lvl):
        if lvl == 3:
            if self.peek() == 'not':
                self.i += 1
                a = self.p_bin(3)
                if a.typ == 'Boolean':
                    return Node(f'(not {a.code})', 'Boolean')
                if a.typ == 'Long':
                    return Node(f'(~{a.code})', 'Long')
                self.err(f'Not on a {a.typ}')
            return self.p_bin(4)
        if lvl == len(self.LEVELS):
            return self.p_neg()
        a = self.p_bin(lvl + 1)
        while self.peek() in self.LEVELS[lvl]:
            op = self.peek()
            self.i += 1
            b = self.p_bin(lvl + 1)
            a = self.binop(op, a, b)
            if lvl == 4 and self.peek() in self.LEVELS[4]:
                self.err('chained comparison')
        return a

    def p_neg(self):
        if self.peek() in ('-', '+'):
            op = self.peek()
            self.i += 1
            a = self.p_neg()
            if a.typ not in ('Long', 'Double'):
                self.err(f'unary {op} on {a.typ}')
            return Node(f'({op}{a.code})', a.typ, lit=a.lit, arith=a.arith)
        a = self.p_primary()
        while self.peek() == '^':
            self.i += 1
            b = self.p_primary()
            a = Node(f'({a.code} ** {b.code})', 'Double', arith=True)
        return a

    def binop(self, op, a, b):
        num = ('Long', 'Double')
        if op in ('+', '-', '*', '/'):
            if a.typ not in num or b.typ not in num:
                self.err(f'{a.typ} {op} {b.typ} (use & to join strings)')
            if a.lit and b.lit:
                self.err('arithmetic on two number literals can overflow VBA Integer - use a Const As Long')
            typ = 'Double' if (op == '/' or 'Double' in (a.typ, b.typ)) else 'Long'
            return Node(f'({a.code} {op} {b.code})', typ, arith=True)
        if op in ('\\', 'mod'):
            if a.typ not in num or b.typ not in num:
                self.err(f'{op} on {a.typ}/{b.typ}')
            fn = 'vidiv' if op == '\\' else 'vmod'
            return Node(f'{fn}({a.code}, {b.code})', 'Long', arith=True)
        if op == '&':
            for x in (a, b):
                if x.typ not in ('String', 'Long'):
                    self.err(f'& with a {x.typ}')
            return Node(f'vcat({a.code}, {b.code})', 'String')
        if op in ('=', '<>', '<', '>', '<=', '>='):
            ok = (a.typ == b.typ == 'String') or (a.typ in num and b.typ in num) or \
                 (a.typ == b.typ == 'Boolean' and op in ('=', '<>'))
            if not ok:
                self.err(f'comparing {a.typ} with {b.typ} (Type mismatch in Excel)')
            pyop = {'=': '==', '<>': '!='}.get(op, op)
            return Node(f'({a.code} {pyop} {b.code})', 'Boolean')
        # and / or / xor: both Boolean or both Long; Python & | ^ evaluate both sides like VBA
        if a.typ != b.typ or a.typ not in ('Boolean', 'Long'):
            self.err(f'{op} between {a.typ} and {b.typ}')
        pyop = {'and': '&', 'or': '|', 'xor': '^'}[op]
        return Node(f'({a.code} {pyop} {b.code})', a.typ)

    def args(self):
        """Parse '(' a, b ')' starting at self.i (which is at '(')."""
        start = self.i
        depth = 0
        for j in range(self.i, len(self.toks)):
            if self.toks[j][1] == '(':
                depth += 1
            elif self.toks[j][1] == ')':
                depth -= 1
                if depth == 0:
                    break
        else:
            self.err('missing )')
        inner = self.toks[start + 1:j]
        self.i = j + 1
        if not inner:
            return []
        saved = (self.toks, self.i)
        res = [Ctx.expr(self, part) for part in split_top(inner, ',')]
        self.toks, self.i = saved
        return res

    def p_primary(self):
        if self.i >= len(self.toks):
            self.err('expression ends too early')
        kind, text = self.toks[self.i]
        if kind == 'num':
            self.i += 1
            if text.endswith('#') or '.' in text:
                return Node(repr(float(text.rstrip('#'))), 'Double', lit=True)
            v = int(text)
            if v > 2147483647:
                self.err('number too big for a Long')
            return Node(str(v), 'Long', lit=True)
        if kind == 'hex':
            self.i += 1
            v = int(text[2:].rstrip('&'), 16)
            if v >= 0x8000 and not text.endswith('&'):
                self.err(f'{text} is a negative Integer in VBA - add a trailing &')
            return Node(str(v), 'Long', lit=True)
        if kind == 'str':
            self.i += 1
            return Node(repr(text[1:-1].replace('""', '"')), 'String')
        if text == '(':
            self.i += 1
            depth, j = 1, self.i
            while depth:
                if j >= len(self.toks):
                    self.err('missing )')
                if self.toks[j][1] == '(':
                    depth += 1
                elif self.toks[j][1] == ')':
                    depth -= 1
                j += 1
            saved = (self.toks, self.i)
            n = Ctx.expr(self, self.toks[self.i:j - 1])
            self.toks, self.i = saved[0], j
            return Node(f'({n.code})', n.typ, lit=n.lit, arith=n.arith)
        if kind != 'id':
            self.err(f'unexpected {text!r}')
        name = low(self.toks[self.i])
        self.i += 1
        if name in ('true', 'false'):
            return Node('True' if name == 'true' else 'False', 'Boolean')
        if name in KEYWORDS:
            self.err(f'unexpected keyword {text}')
        has_paren = self.peek() == '('
        # a function's own name (no parentheses) is its return value
        if self.proc and name == self.proc.name.lower() and not has_paren:
            if self.proc.kind != 'function':
                self.err('a Sub has no return value')
            return Node('_ret', self.proc.typ)
        sym = self.resolve(text.rstrip('$'))
        if sym is None:
            if name not in BUILTINS:
                self.err(f'{text} is not declared (Option Explicit would stop here)')
            lo_, hi_, tpl, rtyp = BUILTINS[name]
            a = self.args() if has_paren else []
            if not lo_ <= len(a) <= hi_:
                self.err(f'{text} takes {lo_}..{hi_} arguments')
            for k in STRING_ARGS.get(name, []):
                if k < len(a) and a[k].typ != 'String':
                    self.err(f'{text} argument {k + 1} must be a String')
            if name == 'ubound' and a[0].typ != 'Array':
                self.err('UBound needs an array')
            code = tpl.format(*[x.code for x in a], a=', '.join(x.code for x in a))
            return Node(code, a[0].typ if rtyp == 'arg0' else rtyp)
        if sym.kind in ('var', 'param', 'const'):
            if has_paren:
                self.err(f'{text} is not an array or function')
            return Node(sym.py, sym.typ)
        if sym.kind in ('array', 'dynarray'):
            if not has_paren:
                return Node(sym.py, 'Array')
            a = self.args()
            need = len(sym.dims) if sym.kind == 'array' else 1
            if len(a) != need:
                self.err(f'{text} has {need} dimension(s), got {len(a)} subscripts')
            for x in a:
                if x.typ == 'Double':          # VBA rounds a Double subscript
                    x.code = f'vlong({x.code})'
                elif x.typ != 'Long':
                    self.err(f'subscript of {text} is a {x.typ}')
            return Node(f'{sym.py}[{", ".join(x.code for x in a)}]', sym.typ)
        if sym.kind == 'sub':
            self.err(f'{text} is a Sub, it has no value')
        a = self.args() if has_paren else []
        self.check_call(sym, a)
        return Node(f'{sym.py}({", ".join(x.code for x in a)})', sym.typ)

    def check_call(self, sym, a):
        if len(a) != len(sym.params):
            self.err(f'{sym.name} takes {len(sym.params)} arguments, got {len(a)}')
        for p, x in zip(sym.params, a):
            if p.kind == 'param':
                if p.typ in ('Long', 'Double') and x.typ not in ('Long', 'Double'):
                    self.err(f'{sym.name}: {p.name} wants a number, got {x.typ}')
                if p.typ in ('String', 'Boolean') and x.typ != p.typ:
                    self.err(f'{sym.name}: {p.name} wants a {p.typ}, got {x.typ}')
                if p.typ == 'Long' and x.typ == 'Double' and not p.byval:
                    self.err(f'{sym.name}: ByRef Long given a Double')


class ProcEmitter:
    def __init__(self, tr, module, p):
        self.tr, self.module, self.p, self.sym = tr, module, p, p['sym']
        self.ctx = Ctx(tr, module, self.sym)
        self.lines, self.ind, self.stack, self.globals_ = [], 1, [], set()
        self.sel = 0
        for ps in self.sym.params:
            self.ctx.locals[ps.name.lower()] = ps

    def w(self, s):
        self.lines.append('    ' * self.ind + s)

    def err(self, where, msg):
        raise VBAError(f'{where}: {msg}')

    def emit(self):
        sym = self.sym
        hdr_where = self.p['hdr'][0]
        body = self.p['body']
        # hoist Dims: VBA locals live for the whole procedure
        decl = []
        for where, toks in body:
            if low(toks[0]) in ('dim', 'static', 'const'):
                self.ctx.where = where
                if low(toks[0]) == 'const':
                    self.err(where, 'no local Const, put it at module level')
                if low(toks[0]) == 'static':
                    self.err(where, 'no Static')
                for part in split_top(toks[1:], ','):
                    s = self.tr.parse_var(part, where, self.module, False, local=True)
                    if s.name.lower() in self.ctx.locals:
                        self.err(where, f'{s.name} declared twice')
                    self.ctx.locals[s.name.lower()] = s
                    decl.append(s)
        for s in decl:
            self.w(f'{s.py} = {self.ctx.init_code(s)}')
        if sym.kind == 'function':
            self.w(f'_ret = {DEFAULT[sym.typ]}')
        for where, toks in body:
            self.ctx.where = where
            self.stmt(where, toks)
        if self.stack:
            self.err(hdr_where, f'{self.stack[-1][0]} block never closed in {sym.name}')
        self.w('return _ret' if sym.kind == 'function' else 'return')
        params = ', '.join(p.py for p in sym.params)
        head = [f'def {sym.py}({params}):']
        if self.globals_:
            head.append('    global ' + ', '.join(sorted(self.globals_)))
        return head + self.lines

    def cond(self, toks, where):
        e = self.ctx.expr(toks)
        if e.typ not in ('Boolean', 'Long'):
            self.err(where, f'condition is a {e.typ}')
        return e.code

    def stmt(self, where, toks):
        w0 = low(toks[0])
        words = [low(t) for t in toks]
        if w0 in ('dim',):
            return
        if w0 == 'if':
            k = find_top(toks, 'then')
            if k < 0:
                self.err(where, 'If without Then')
            c = self.cond(toks[1:k], where)
            rest = toks[k + 1:]
            if not rest:
                self.w(f'if {c}:')
                self.ind += 1
                self.w('pass')
                self.stack.append(('if', self.ind))
                return
            e = find_top(rest, 'else')
            if low(rest[0]) == 'if' or (e >= 0 and rest[e + 1:] and low(rest[e + 1]) == 'if'):
                self.err(where, 'nested single-line If')
            if e >= 0:
                for branch in (rest[:e], rest[e + 1:]):
                    s = self.ctx.resolve(branch[0][1]) if branch[0][0] == 'id' else None
                    if low(branch[0]) == 'call' or (s is not None and s.kind in ('sub', 'function')
                                                    and find_top(branch, '=') < 0):
                        self.err(where, 'single-line If ... Else with a Sub call crashes '
                                        "LibreOffice's Basic compiler - write a block If")
            self.w(f'if {c}:')
            self.ind += 1
            self.simple(where, rest if e < 0 else rest[:e])
            self.ind -= 1
            if e >= 0:
                self.w('else:')
                self.ind += 1
                self.simple(where, rest[e + 1:])
                self.ind -= 1
            return
        if w0 == 'elseif':
            self.top(where, 'if')
            k = find_top(toks, 'then')
            if k < 0 or k != len(toks) - 1:
                self.err(where, 'ElseIf ... Then must end the line')
            self.ind -= 1
            self.w(f'elif {self.cond(toks[1:k], where)}:')
            self.ind += 1
            self.w('pass')
            return
        if w0 == 'else' and len(toks) == 1:
            self.top(where, 'if')
            self.ind -= 1
            self.w('else:')
            self.ind += 1
            self.w('pass')
            return
        if words[:2] == ['end', 'if']:
            self.top(where, 'if')
            self.stack.pop()
            self.ind -= 1
            return
        if w0 == 'for':
            if len(toks) < 4 or toks[2][1] != '=':
                self.err(where, 'bad For')
            var = self.ctx.resolve(toks[1][1])
            if var is None or var.kind not in ('var', 'param') or var.typ != 'Long':
                self.err(where, 'For needs a declared Long loop variable')
            if var.kind == 'param' and not var.byval:
                self.err(where, 'loop variable is a ByRef parameter')
            self.note_assign(var)
            rest = toks[3:]
            k = find_top(rest, 'to')
            s = find_top(rest, 'step')
            a = self.ctx.expr(rest[:k])
            b = self.ctx.expr(rest[k + 1:s if s >= 0 else None])
            step = 1
            if s >= 0:
                st = rest[s + 1:]
                txt = ''.join(t[1] for t in st)
                if not re.fullmatch(r'-?\d+', txt):
                    self.err(where, 'Step must be a number literal')
                step = int(txt)
            for x in (a, b):
                if x.typ != 'Long':
                    self.err(where, 'For bounds must be Long')
            end = f'({b.code}) + 1' if step > 0 else f'({b.code}) - 1'
            self.w(f'for {var.py} in range({a.code}, {end}, {step}):')
            self.ind += 1
            self.w('pass')
            self.stack.append(('for', var.name.lower()))
            return
        if w0 == 'next':
            self.top(where, 'for')
            if len(toks) > 1 and low(toks[1]) != self.stack[-1][1]:
                self.err(where, f'Next {toks[1][1]} closes For {self.stack[-1][1]}')
            self.stack.pop()
            self.ind -= 1
            return
        if w0 == 'do':
            if len(toks) == 1:
                self.w('while True:')
            elif words[1] == 'while':
                self.w(f'while {self.cond(toks[2:], where)}:')
            elif words[1] == 'until':
                self.w(f'while not {self.cond(toks[2:], where)}:')
            else:
                self.err(where, 'bad Do')
            self.ind += 1
            self.w('pass')
            self.stack.append(('do', None))
            return
        if w0 == 'loop':
            self.top(where, 'do')
            if len(toks) > 1:
                c = self.cond(toks[2:], where)
                self.w(f'if not {c}: break' if words[1] == 'while' else f'if {c}: break')
            self.stack.pop()
            self.ind -= 1
            return
        if words[:2] == ['select', 'case']:
            e = self.ctx.expr(toks[2:])
            self.sel += 1
            v = f'_sel{self.sel}'
            self.w(f'{v} = {e.code}')
            self.stack.append(('select', v, e.typ, self.ind, [False]))
            return
        if w0 == 'case':
            self.top(where, 'select')
            _, v, typ, ind, seen = self.stack[-1]
            self.ind = ind
            if len(toks) == 2 and words[1] == 'else':
                if not seen[0]:
                    self.err(where, 'Case Else first')
                self.w('else:')
            else:
                conds = []
                for part in split_top(toks[1:], ','):
                    k = find_top(part, 'to')
                    if k >= 0:
                        lo, hi = self.ctx.expr(part[:k]), self.ctx.expr(part[k + 1:])
                        conds.append(f'({lo.code} <= {v} <= {hi.code})')
                    else:
                        x = self.ctx.expr(part)
                        if (x.typ == 'String') != (typ == 'String'):
                            self.err(where, 'Case type mismatch')
                        conds.append(f'({v} == {x.code})')
                self.w(('elif ' if seen[0] else 'if ') + ' or '.join(conds) + ':')
                seen[0] = True
            self.ind = ind + 1
            self.w('pass')
            return
        if words[:2] == ['end', 'select']:
            self.top(where, 'select')
            self.ind = self.stack[-1][3]
            self.stack.pop()
            return
        self.simple(where, toks)

    def top(self, where, kind):
        if not self.stack or self.stack[-1][0] != kind:
            got = self.stack[-1][0] if self.stack else 'nothing'
            self.err(where, f'expected to be inside {kind}, but inside {got}')

    def note_assign(self, sym):
        if sym.kind == 'param':
            if not sym.byval:
                self.err(self.ctx.where, f'{sym.name} is ByRef - assigning it changes the caller (make it ByVal)')
        elif sym.name.lower() not in self.ctx.locals:
            self.globals_.add(sym.py)

    def simple(self, where, toks):
        """One statement that is not a block keyword: Exit, Call, assignment, sub call."""
        words = [low(t) for t in toks]
        w0 = words[0]
        if w0 == 'exit':
            what = words[1]
            if what == 'function':
                if self.sym.kind != 'function':
                    self.err(where, 'Exit Function in a Sub')
                self.w('return _ret')
            elif what == 'sub':
                if self.sym.kind != 'sub':
                    self.err(where, 'Exit Sub in a Function')
                self.w('return')
            elif what in ('for', 'do'):
                loops = [s[0] for s in self.stack if s[0] in ('for', 'do')]
                if not loops or loops[-1] != what:
                    self.err(where, f'Exit {what.title()} is not inside a {what.title()} (Python would break the wrong loop)')
                self.w('break')
            else:
                self.err(where, 'bad Exit')
            return
        if w0 in ('doevents', 'randomize'):
            self.w('pass')
            return
        if w0 in ('set', 'with', 'on', 'goto', 'redim', 'end', 'else', 'elseif', 'loop',
                  'next', 'case', 'select'):
            self.err(where, f'{toks[0][1]} not allowed here')
        if w0 == 'call':
            toks = toks[1:]
            words = words[1:]
            if len(toks) < 2 or toks[1][1] != '(' or toks[-1][1] != ')':
                self.err(where, 'Call needs Name(args)')
            self.call(where, toks[0], toks[2:-1])
            return
        eq = find_top(toks, '=')
        if eq > 0:
            self.assign(where, toks[:eq], toks[eq + 1:])
            return
        if toks[0][0] != 'id':
            self.err(where, 'what is this statement?')
        args = toks[1:]
        if args and args[0][1] == '(' and match_paren(args, 0) == len(args) - 1:
            inner = args[1:-1]
            if len(split_top(inner, ',')) > 1:
                self.err(where, 'Sub call with (a, b) needs Call - Excel says "Expected: ="')
            args = inner
        self.call(where, toks[0], args)

    def call(self, where, name_tok, arg_toks):
        sym = self.ctx.resolve(name_tok[1])
        if sym is None or sym.kind not in ('sub', 'function'):
            self.err(where, f'{name_tok[1]} is not a Sub/Function')
        a = [self.ctx.expr(p) for p in split_top(arg_toks, ',')] if arg_toks else []
        self.ctx.check_call(sym, a)
        self.w(f'{sym.py}({", ".join(x.code for x in a)})')

    def assign(self, where, lhs, rhs):
        name = lhs[0][1]
        if self.sym.kind == 'function' and name.lower() == self.sym.name.lower() and len(lhs) == 1:
            e = self.ctx.expr(rhs)
            self.w(f'_ret = {self.tr.check_assign(self.sym.typ, e, where)}')
            return
        sym = self.ctx.resolve(name)
        if sym is None:
            self.err(where, f'{name} is not declared')
        if sym.kind == 'const':
            self.err(where, f'{name} is a Const')
        if sym.kind in ('var', 'param'):
            if len(lhs) != 1:
                self.err(where, f'{name} is not an array')
            self.note_assign(sym)
            e = self.ctx.expr(rhs)
            self.w(f'{sym.py} = {self.tr.check_assign(sym.typ, e, where)}')
            return
        if sym.kind == 'dynarray' and len(lhs) == 1:
            e = self.ctx.expr(rhs)
            if e.typ != 'Array':
                self.err(where, f'{name} needs an array')
            self.note_assign(sym)
            self.w(f'{sym.py} = {e.code}')
            return
        if sym.kind in ('array', 'dynarray'):
            target = self.ctx.expr(lhs)   # checks subscripts
            e = self.ctx.expr(rhs)
            self.w(f'{target.code} = {self.tr.check_assign(sym.typ, e, where)}')
            return
        self.err(where, f'cannot assign to {name}')


def translate(files, externs=()):
    """files: {module: source}. externs: {module: source} whose procedures exist (headers
    are checked) but whose bodies are provided by the test harness."""
    tr = Translator()
    for name, src in dict(externs).items():
        tr.add_module(name, src, bodies=False)
    for name, src in files.items():
        tr.add_module(name, src)
    return tr.translate()
