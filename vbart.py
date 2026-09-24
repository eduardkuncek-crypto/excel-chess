"""Runtime for code translated by vba2py.py: VBA semantics that differ from Python's
(bounds-checked arrays, Long overflow, \\ and Mod truncating toward zero, banker's rounding)."""
# aios-run: agent
import math
import random
import time


class VBARuntimeError(Exception):
    pass


class VArr:
    __slots__ = ('lo', 'hi', 'a')

    def __init__(self, default, lo, hi):
        self.lo, self.hi = lo, hi
        self.a = [default] * (hi - lo + 1)

    def __getitem__(self, i):
        if type(i) is not int or not self.lo <= i <= self.hi:
            raise VBARuntimeError(f'Subscript out of range: {i!r} not in {self.lo}..{self.hi}')
        return self.a[i - self.lo]

    def __setitem__(self, i, v):
        if type(i) is not int or not self.lo <= i <= self.hi:
            raise VBARuntimeError(f'Subscript out of range: {i!r} not in {self.lo}..{self.hi}')
        self.a[i - self.lo] = v


class VArr2:
    __slots__ = ('lo1', 'hi1', 'lo2', 'hi2', 'w', 'a')

    def __init__(self, default, lo1, hi1, lo2, hi2):
        self.lo1, self.hi1, self.lo2, self.hi2 = lo1, hi1, lo2, hi2
        self.w = hi2 - lo2 + 1
        self.a = [default] * ((hi1 - lo1 + 1) * self.w)

    def _k(self, k):
        i, j = k
        if type(i) is not int or type(j) is not int or not (
                self.lo1 <= i <= self.hi1 and self.lo2 <= j <= self.hi2):
            raise VBARuntimeError(f'Subscript out of range: ({i!r}, {j!r})')
        return (i - self.lo1) * self.w + j - self.lo2

    def __getitem__(self, k):
        return self.a[self._k(k)]

    def __setitem__(self, k, v):
        self.a[self._k(k)] = v


class VList(list):
    def __getitem__(self, i):
        if type(i) is not int or not 0 <= i < len(self):
            raise VBARuntimeError(f'Subscript out of range: {i!r} (array has {len(self)})')
        return list.__getitem__(self, i)


def vround(x):
    return x if type(x) is int else round(x)   # round() is banker's rounding, like VBA


def vlong(x):
    if type(x) is bool:
        raise VBARuntimeError('Boolean used as a number')
    x = vround(x)
    if not -2147483648 <= x <= 2147483647:
        raise VBARuntimeError(f'Overflow: {x} does not fit a Long')
    return x


def vclng(x):
    if isinstance(x, str):
        try:
            x = float(x.strip())
        except ValueError:
            raise VBARuntimeError(f'Type mismatch: CLng({x!r})')
    return vlong(x)


def vidiv(a, b):
    a, b = vround(a), vround(b)
    if b == 0:
        raise VBARuntimeError('Division by zero')
    q = abs(a) // abs(b)
    return q if (a >= 0) == (b > 0) else -q


def vmod(a, b):
    a, b = vround(a), vround(b)
    if b == 0:
        raise VBARuntimeError('Division by zero')
    r = abs(a) % abs(b)
    return r if a >= 0 else -r


def vint(x):
    return float(math.floor(x)) if type(x) is float else x


def vstr(x):
    if type(x) is bool:
        return 'True' if x else 'False'
    return str(x)


def vcat(a, b):
    return vstr(a) + vstr(b)


def vmid(s, i, n=None):
    if i < 1 or (n is not None and n < 0):
        raise VBARuntimeError(f'Invalid procedure call: Mid$({s!r}, {i}, {n})')
    return s[i - 1:] if n is None else s[i - 1:i - 1 + n]


def vleft(s, n):
    if n < 0:
        raise VBARuntimeError('Left$ with a negative length')
    return s[:n]


def vright(s, n):
    if n < 0:
        raise VBARuntimeError('Right$ with a negative length')
    return s[len(s) - n:] if n else ''


def vinstr(*a):
    start, s, sub = (1, *a) if len(a) == 2 else a
    return s.find(sub, start - 1) + 1


def vtrim(s):
    return s.strip(' ')


def vchr(n):
    return chr(n)


def vsplit(s, d):
    return VList() if s == '' else VList(s.split(d))


def vubound(a):
    return a.hi if isinstance(a, VArr) else len(a) - 1


def vreplace(s, a, b):
    return s.replace(a, b)


def vrnd():
    return random.random()


def vtimer():
    t = time.time()
    lt = time.localtime(t)
    return lt.tm_hour * 3600 + lt.tm_min * 60 + lt.tm_sec + (t % 1)


def vrgb(r, g, b):
    return r + g * 256 + b * 65536
