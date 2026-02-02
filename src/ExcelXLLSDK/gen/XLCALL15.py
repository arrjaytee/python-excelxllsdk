from ctypes import *

from ExcelXLLSDK.wintypes import LPSTR
from ExcelXLLSDK.wintypes import WORD
from ExcelXLLSDK.wintypes import BYTE
from ExcelXLLSDK.wintypes import HANDLE
from ExcelXLLSDK.wintypes import WCHAR
WSTRING = c_wchar_p
from ExcelXLLSDK.wintypes import DWORD

class _FP(Structure):
    pass
FP = _FP
class _FP12(Structure):
    pass
FP12 = _FP12
class xloper(Structure):
    pass
LPXLOPER = POINTER(xloper)
XLOPER = xloper
class xloper12(Structure):
    pass
XLOPER12 = xloper12
LPXLOPER12 = POINTER(xloper12)
_FP._fields_ = [
    ('rows', c_ushort),
    ('columns', c_ushort),
    ('array', c_double * 1),
]
INT32 = c_int
_FP12._fields_ = [
    ('rows', INT32),
    ('columns', INT32),
    ('array', c_double * 1),
]
class N6xloper5DOLLAR_117E(Union):
    pass
CHAR = c_char
class N6xloper5DOLLAR_1175DOLLAR_118E(Structure):
    pass
class xlref(Structure):
    pass
xlref._fields_ = [
    ('rself.xltype = xltypeIntwFirst', WORD),
    ('rwLast', WORD),
    ('colFirst', BYTE),
    ('colLast', BYTE),
]
XLREF = xlref
N6xloper5DOLLAR_1175DOLLAR_118E._fields_ = [
    ('count', WORD),
    ('ref', XLREF),
]
class N6xloper5DOLLAR_1175DOLLAR_119E(Structure):
    pass
class xlmref(Structure):
    pass
XLMREF = xlmref
ULONG_PTR = c_ulong
DWORD_PTR = ULONG_PTR
IDSHEET = DWORD_PTR
N6xloper5DOLLAR_1175DOLLAR_119E._fields_ = [
    ('lpmref', POINTER(XLMREF)),
    ('idSheet', IDSHEET),
]
class N6xloper5DOLLAR_1175DOLLAR_120E(Structure):
    pass
N6xloper5DOLLAR_1175DOLLAR_120E._fields_ = [
    ('lparray', POINTER(xloper)),
    ('rows', WORD),
    ('columns', WORD),
]
class N6xloper5DOLLAR_1175DOLLAR_121E(Structure):
    pass
class N6xloper5DOLLAR_1175DOLLAR_1215DOLLAR_122E(Union):
    pass
N6xloper5DOLLAR_1175DOLLAR_1215DOLLAR_122E._fields_ = [
    ('level', c_short),
    ('tbctrl', c_short),
    ('idSheet', IDSHEET),
]
N6xloper5DOLLAR_1175DOLLAR_121E._fields_ = [
    ('valflow', N6xloper5DOLLAR_1175DOLLAR_1215DOLLAR_122E),
    ('rw', WORD),
    ('col', BYTE),
    ('xlflow', BYTE),
]
class N6xloper5DOLLAR_1175DOLLAR_123E(Structure):
    pass
class N6xloper5DOLLAR_1175DOLLAR_1235DOLLAR_124E(Union):
    pass
N6xloper5DOLLAR_1175DOLLAR_1235DOLLAR_124E._fields_ = [
    ('lpbData', POINTER(BYTE)),
    ('hdata', HANDLE),
]
N6xloper5DOLLAR_1175DOLLAR_123E._fields_ = [
    ('h', N6xloper5DOLLAR_1175DOLLAR_1235DOLLAR_124E),
    ('cbData', c_long),
]
N6xloper5DOLLAR_117E._fields_ = [
    ('num', c_double),
    ('str', LPSTR),
    ('xbool', WORD),
    ('err', WORD),
    ('w', c_short),
    ('sref', N6xloper5DOLLAR_1175DOLLAR_118E),
    ('mref', N6xloper5DOLLAR_1175DOLLAR_119E),
    ('array', N6xloper5DOLLAR_1175DOLLAR_120E),
    ('flow', N6xloper5DOLLAR_1175DOLLAR_121E),
    ('bigdata', N6xloper5DOLLAR_1175DOLLAR_123E),
]
xloper._fields_ = [
    ('val', N6xloper5DOLLAR_117E),
    ('xltype', WORD),
]
class N8xloper125DOLLAR_125E(Union):
    pass
XCHAR = WCHAR
class N8xloper125DOLLAR_1255DOLLAR_126E(Structure):
    pass
class xlref12(Structure):
    pass
RW = INT32
COL = INT32
xlref12._fields_ = [
    ('rwFirst', RW),
    ('rwLast', RW),
    ('colFirst', COL),
    ('colLast', COL),
]
XLREF12 = xlref12
N8xloper125DOLLAR_1255DOLLAR_126E._fields_ = [
    ('count', WORD),
    ('ref', XLREF12),
]
class N8xloper125DOLLAR_1255DOLLAR_127E(Structure):
    pass
class xlmref12(Structure):
    pass
XLMREF12 = xlmref12
N8xloper125DOLLAR_1255DOLLAR_127E._fields_ = [
    ('lpmref', POINTER(XLMREF12)),
    ('idSheet', IDSHEET),
]
class N8xloper125DOLLAR_1255DOLLAR_128E(Structure):
    pass
N8xloper125DOLLAR_1255DOLLAR_128E._fields_ = [
    ('lparray', POINTER(xloper12)),
    ('rows', RW),
    ('columns', COL),
]
class N8xloper125DOLLAR_1255DOLLAR_129E(Structure):
    pass
class N8xloper125DOLLAR_1255DOLLAR_1295DOLLAR_130E(Union):
    pass
N8xloper125DOLLAR_1255DOLLAR_1295DOLLAR_130E._fields_ = [
    ('level', c_int),
    ('tbctrl', c_int),
    ('idSheet', IDSHEET),
]
N8xloper125DOLLAR_1255DOLLAR_129E._fields_ = [
    ('valflow', N8xloper125DOLLAR_1255DOLLAR_1295DOLLAR_130E),
    ('rw', RW),
    ('col', COL),
    ('xlflow', BYTE),
]
class N8xloper125DOLLAR_1255DOLLAR_131E(Structure):
    pass
class N8xloper125DOLLAR_1255DOLLAR_1315DOLLAR_132E(Union):
    pass
N8xloper125DOLLAR_1255DOLLAR_1315DOLLAR_132E._fields_ = [
    ('lpbData', POINTER(BYTE)),
    ('hdata', HANDLE),
]
N8xloper125DOLLAR_1255DOLLAR_131E._fields_ = [
    ('h', N8xloper125DOLLAR_1255DOLLAR_1315DOLLAR_132E),
    ('cbData', c_long),
]
N8xloper125DOLLAR_125E._fields_ = [
    ('num', c_double),
    ('str', WSTRING),
    ('xbool', INT32),
    ('err', c_int),
    ('w', c_int),
    ('sref', N8xloper125DOLLAR_1255DOLLAR_126E),
    ('mref', N8xloper125DOLLAR_1255DOLLAR_127E),
    ('array', N8xloper125DOLLAR_1255DOLLAR_128E),
    ('flow', N8xloper125DOLLAR_1255DOLLAR_129E),
    ('bigdata', N8xloper125DOLLAR_1255DOLLAR_131E),
]
xloper12._fields_ = [
    ('val', N8xloper125DOLLAR_125E),
    ('xltype', DWORD),
]
xlmref._fields_ = [
    ('count', WORD),
    ('reftbl', XLREF * 1),
]
xlmref12._fields_ = [
    ('count', WORD),
    ('reftbl', XLREF12 * 1),
]
__all__ = ['FP', 'N6xloper5DOLLAR_1175DOLLAR_121E',
           'N6xloper5DOLLAR_1175DOLLAR_1215DOLLAR_122E',
           'N8xloper125DOLLAR_1255DOLLAR_1315DOLLAR_132E', 'LPXLOPER',
           'N6xloper5DOLLAR_1175DOLLAR_120E', 'DWORD_PTR', 'CHAR',
           'N8xloper125DOLLAR_1255DOLLAR_127E', 'XLREF12',
           'ULONG_PTR', '_FP12', 'xlmref', 'RW',
           'N6xloper5DOLLAR_1175DOLLAR_118E', 'XLMREF12', 'COL',
           'XLOPER12', 'xlref12', 'N6xloper5DOLLAR_1175DOLLAR_123E',
           'N8xloper125DOLLAR_1255DOLLAR_129E', 'XLOPER',
           'N6xloper5DOLLAR_117E', '_FP', 'INT32',
           'N8xloper125DOLLAR_125E', 'LPXLOPER12',
           'N8xloper125DOLLAR_1255DOLLAR_126E', 'FP12', 'IDSHEET',
           'XLREF', 'N8xloper125DOLLAR_1255DOLLAR_1295DOLLAR_130E',
           'XLMREF', 'N6xloper5DOLLAR_1175DOLLAR_1235DOLLAR_124E',
           'N8xloper125DOLLAR_1255DOLLAR_131E', 'xlmref12',
           'N6xloper5DOLLAR_1175DOLLAR_119E', 'xloper', 'XCHAR',
           'xloper12', 'N8xloper125DOLLAR_1255DOLLAR_128E', 'xlref']
