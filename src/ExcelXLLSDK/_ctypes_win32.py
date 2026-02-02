# handmade definitions of win32 API functions to support _xltypes
# should really automate this using ctypeslib

from ctypes import *
from ctypes.wintypes import *

user32 = windll.user32

WNDENUMPROC = WINFUNCTYPE(BOOL, HWND, LPARAM)

EnumWindows = WINFUNCTYPE(BOOL, WNDENUMPROC, LPARAM)(('EnumWindows', user32))
EnumChildWindows = WINFUNCTYPE(BOOL, HWND, WNDENUMPROC, LPARAM)(('EnumChildWindows', user32))
GetWindowThreadProcessId = WINFUNCTYPE(DWORD, HWND, POINTER(DWORD))(('GetWindowThreadProcessId', user32))

from comtypes import IID

REFIID = POINTER(IID)

oleacc = oledll.oleacc
AccessibleObjectFromWindow = oleacc.AccessibleObjectFromWindow
AccessibleObjectFromWindow.argtypes = [HWND, DWORD, REFIID, POINTER(c_void_p)]
OBJID_NATIVEOM = 0xFFFFFFF0

DLL_PROCESS_ATTACH = 1
LPCTSTR = c_char_p

DisableThreadLibraryCalls = WINFUNCTYPE(BOOL, HMODULE)(('DisableThreadLibraryCalls', windll.kernel32))

GetModuleHandle = WINFUNCTYPE(HMODULE, LPCTSTR)(('GetModuleHandleA', windll.kernel32))
GetModuleFileName = WINFUNCTYPE(DWORD, HMODULE, LPSTR, DWORD)(('GetModuleFileNameA', windll.kernel32))

# expose the console API
# import stuff
FreeConsole = WINFUNCTYPE(BOOL)(('FreeConsole', windll.kernel32))
AllocConsole = WINFUNCTYPE(BOOL)(('AllocConsole', windll.kernel32))
GetConsoleWindow = WINFUNCTYPE(HWND)(('GetConsoleWindow', windll.kernel32))
