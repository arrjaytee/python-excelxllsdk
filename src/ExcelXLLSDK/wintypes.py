"""define a few windows types necessary for excel

ctypes.wintypes won't import on any other platform
"""
import ctypes

LPSTR = ctypes.c_char_p
WORD = ctypes.c_ushort
BYTE = ctypes.c_byte
HANDLE = ctypes.c_void_p
WCHAR = ctypes.c_wchar
DWORD = ctypes.c_ulong





