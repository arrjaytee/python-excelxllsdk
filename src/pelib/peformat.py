from ctypes import *

from ctypes.wintypes import LPCSTR
_stdcall_libraries = {}
_stdcall_libraries['kernel32'] = WinDLL('kernel32')
from ctypes.wintypes import LPSTR
from ctypes.wintypes import BOOL
from ctypes.wintypes import LPVOID
from ctypes.wintypes import DWORD
from ctypes.wintypes import LPCWSTR
from ctypes.wintypes import LPWSTR
from ctypes.wintypes import WORD
from ctypes.wintypes import LONG
from ctypes.wintypes import BYTE


lstrlenA = _stdcall_libraries['kernel32'].lstrlenA
lstrlenA.restype = c_int
lstrlenA.argtypes = [LPCSTR]
lstrlen = lstrlenA # alias
lstrcpynA = _stdcall_libraries['kernel32'].lstrcpynA
lstrcpynA.restype = LPSTR
lstrcpynA.argtypes = [LPSTR, LPCSTR, c_int]
lstrcpyn = lstrcpynA # alias
lstrcatA = _stdcall_libraries['kernel32'].lstrcatA
lstrcatA.restype = LPSTR
lstrcatA.argtypes = [LPSTR, LPCSTR]
lstrcat = lstrcatA # alias
lstrcmpA = _stdcall_libraries['kernel32'].lstrcmpA
lstrcmpA.restype = c_int
lstrcmpA.argtypes = [LPCSTR, LPCSTR]
lstrcmp = lstrcmpA # alias
lstrcpyA = _stdcall_libraries['kernel32'].lstrcpyA
lstrcpyA.restype = LPSTR
lstrcpyA.argtypes = [LPSTR, LPCSTR]
lstrcpy = lstrcpyA # alias
lstrcmpiA = _stdcall_libraries['kernel32'].lstrcmpiA
lstrcmpiA.restype = c_int
lstrcmpiA.argtypes = [LPCSTR, LPCSTR]
lstrcmpi = lstrcmpiA # alias
ULONG_PTR = c_ulong
SIZE_T = ULONG_PTR
PDWORD = POINTER(DWORD)
VirtualProtect = _stdcall_libraries['kernel32'].VirtualProtect
VirtualProtect.restype = BOOL
VirtualProtect.argtypes = [LPVOID, SIZE_T, DWORD, PDWORD]
lstrcmpW = _stdcall_libraries['kernel32'].lstrcmpW
lstrcmpW.restype = c_int
lstrcmpW.argtypes = [LPCWSTR, LPCWSTR]
lstrcmpiW = _stdcall_libraries['kernel32'].lstrcmpiW
lstrcmpiW.restype = c_int
lstrcmpiW.argtypes = [LPCWSTR, LPCWSTR]
lstrcpynW = _stdcall_libraries['kernel32'].lstrcpynW
lstrcpynW.restype = LPWSTR
lstrcpynW.argtypes = [LPWSTR, LPCWSTR, c_int]
lstrcpyW = _stdcall_libraries['kernel32'].lstrcpyW
lstrcpyW.restype = LPWSTR
lstrcpyW.argtypes = [LPWSTR, LPCWSTR]
lstrcatW = _stdcall_libraries['kernel32'].lstrcatW
lstrcatW.restype = LPWSTR
lstrcatW.argtypes = [LPWSTR, LPCWSTR]
lstrlenW = _stdcall_libraries['kernel32'].lstrlenW
lstrlenW.restype = c_int
lstrlenW.argtypes = [LPCWSTR]
class _IMAGE_DOS_HEADER(Structure):
    pass
IMAGE_DOS_HEADER = _IMAGE_DOS_HEADER
class _IMAGE_FILE_HEADER(Structure):
    pass
IMAGE_FILE_HEADER = _IMAGE_FILE_HEADER
class _IMAGE_OPTIONAL_HEADER(Structure):
    pass
IMAGE_OPTIONAL_HEADER32 = _IMAGE_OPTIONAL_HEADER
class _IMAGE_NT_HEADERS(Structure):
    pass
IMAGE_NT_HEADERS32 = _IMAGE_NT_HEADERS
class _IMAGE_EXPORT_DIRECTORY(Structure):
    pass
IMAGE_EXPORT_DIRECTORY = _IMAGE_EXPORT_DIRECTORY
PAGE_EXECUTE_READ = 32 # Variable c_int '32'
PAGE_EXECUTE = 16 # Variable c_int '16'
PAGE_WRITECOPY = 8 # Variable c_int '8'
PAGE_NOCACHE = 512 # Variable c_int '512'
PAGE_READONLY = 2 # Variable c_int '2'
PAGE_READWRITE = 4 # Variable c_int '4'
PAGE_EXECUTE_READWRITE = 64 # Variable c_int '64'
PAGE_WRITECOMBINE = 1024 # Variable c_int '1024'
PAGE_GUARD = 256 # Variable c_int '256'
PAGE_NOACCESS = 1 # Variable c_int '1'
PAGE_EXECUTE_WRITECOPY = 128 # Variable c_int '128'
_IMAGE_DOS_HEADER._pack_ = 2
_IMAGE_DOS_HEADER._fields_ = [
    ('e_magic', WORD),
    ('e_cblp', WORD),
    ('e_cp', WORD),
    ('e_crlc', WORD),
    ('e_cparhdr', WORD),
    ('e_minalloc', WORD),
    ('e_maxalloc', WORD),
    ('e_ss', WORD),
    ('e_sp', WORD),
    ('e_csum', WORD),
    ('e_ip', WORD),
    ('e_cs', WORD),
    ('e_lfarlc', WORD),
    ('e_ovno', WORD),
    ('e_res', WORD * 4),
    ('e_oemid', WORD),
    ('e_oeminfo', WORD),
    ('e_res2', WORD * 10),
    ('e_lfanew', LONG),
]
_IMAGE_FILE_HEADER._fields_ = [
    ('Machine', WORD),
    ('NumberOfSections', WORD),
    ('TimeDateStamp', DWORD),
    ('PointerToSymbolTable', DWORD),
    ('NumberOfSymbols', DWORD),
    ('SizeOfOptionalHeader', WORD),
    ('Characteristics', WORD),
]
class _IMAGE_DATA_DIRECTORY(Structure):
    pass
_IMAGE_DATA_DIRECTORY._fields_ = [
    ('VirtualAddress', DWORD),
    ('Size', DWORD),
]
IMAGE_DATA_DIRECTORY = _IMAGE_DATA_DIRECTORY
_IMAGE_OPTIONAL_HEADER._fields_ = [
    ('Magic', WORD),
    ('MajorLinkerVersion', BYTE),
    ('MinorLinkerVersion', BYTE),
    ('SizeOfCode', DWORD),
    ('SizeOfInitializedData', DWORD),
    ('SizeOfUninitializedData', DWORD),
    ('AddressOfEntryPoint', DWORD),
    ('BaseOfCode', DWORD),
    ('BaseOfData', DWORD),
    ('ImageBase', DWORD),
    ('SectionAlignment', DWORD),
    ('FileAlignment', DWORD),
    ('MajorOperatingSystemVersion', WORD),
    ('MinorOperatingSystemVersion', WORD),
    ('MajorImageVersion', WORD),
    ('MinorImageVersion', WORD),
    ('MajorSubsystemVersion', WORD),
    ('MinorSubsystemVersion', WORD),
    ('Win32VersionValue', DWORD),
    ('SizeOfImage', DWORD),
    ('SizeOfHeaders', DWORD),
    ('CheckSum', DWORD),
    ('Subsystem', WORD),
    ('DllCharacteristics', WORD),
    ('SizeOfStackReserve', DWORD),
    ('SizeOfStackCommit', DWORD),
    ('SizeOfHeapReserve', DWORD),
    ('SizeOfHeapCommit', DWORD),
    ('LoaderFlags', DWORD),
    ('NumberOfRvaAndSizes', DWORD),
    ('DataDirectory', IMAGE_DATA_DIRECTORY * 16),
]
_IMAGE_NT_HEADERS._fields_ = [
    ('Signature', DWORD),
    ('FileHeader', IMAGE_FILE_HEADER),
    ('OptionalHeader', IMAGE_OPTIONAL_HEADER32),
]
_IMAGE_EXPORT_DIRECTORY._fields_ = [
    ('Characteristics', DWORD),
    ('TimeDateStamp', DWORD),
    ('MajorVersion', WORD),
    ('MinorVersion', WORD),
    ('Name', DWORD),
    ('Base', DWORD),
    ('NumberOfFunctions', DWORD),
    ('NumberOfNames', DWORD),
    ('AddressOfFunctions', DWORD),
    ('AddressOfNames', DWORD),
    ('AddressOfNameOrdinals', DWORD),
]
__all__ = ['_IMAGE_OPTIONAL_HEADER', '_IMAGE_DATA_DIRECTORY',
           'PAGE_EXECUTE', 'PAGE_EXECUTE_READ', 'PDWORD', 'lstrlenA',
           'lstrcmpiA', 'lstrcmp', 'lstrlen', 'VirtualProtect',
           'PAGE_NOCACHE', 'lstrcmpiW', 'PAGE_READONLY', 'lstrlenW',
           'PAGE_EXECUTE_WRITECOPY', 'IMAGE_FILE_HEADER',
           '_IMAGE_FILE_HEADER', 'lstrcatA', 'ULONG_PTR', 'lstrcat',
           'IMAGE_NT_HEADERS32', 'PAGE_EXECUTE_READWRITE', 'lstrcatW',
           'PAGE_READWRITE', '_IMAGE_DOS_HEADER', 'IMAGE_DOS_HEADER',
           'lstrcpynW', 'lstrcpyW', 'IMAGE_EXPORT_DIRECTORY',
           'lstrcmpA', 'lstrcpy', 'PAGE_WRITECOPY', 'lstrcpyA',
           'IMAGE_OPTIONAL_HEADER32', 'PAGE_NOACCESS', 'lstrcpynA',
           'lstrcmpW', '_IMAGE_EXPORT_DIRECTORY', 'lstrcmpi',
           'IMAGE_DATA_DIRECTORY', 'PAGE_GUARD', '_IMAGE_NT_HEADERS',
           'SIZE_T', 'lstrcpyn', 'PAGE_WRITECOMBINE']
