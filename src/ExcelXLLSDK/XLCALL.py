import sys
import ctypes
import logging
from ctypes import c_int, POINTER, cast, pointer, byref

from ExcelXLLSDK.xltypes import (
    XLOPER12, LPXLOPER12, OPER12
)
import ExcelXLLSDK.gen.xlerr
import ExcelXLLSDK.gen.xlcall


from ExcelXLLSDK.gen.xltype import *
from ExcelXLLSDK.gen.xlret import *


OPER = OPER12
XLOPER = XLOPER12
LPXLOPER = LPXLOPER12

_log = logging.getLogger(__name__)

class ExcelError(StandardError):
    pass


class AbortError(ExcelError):
    pass


class InvXlfnError(ExcelError):
    pass


class InvCountError(ExcelError):
    pass


class InvXloperError(ExcelError):
    pass


class StackOvflError(ExcelError):
    pass


class FailedError(ExcelError):
    pass


class UncalcedError(ExcelError):
    pass


class NotThreadSafeError(ExcelError):
    pass


class InvAsynchronousContextError(ExcelError):
    pass

def _xlret_errcheck(result, _, __):
    if result == xlretSuccess:
        return None
    if result == xlretAbort:
        raise AbortError()
    if result == xlretInvXlfn:
        raise InvXlfnError()
    if result == xlretInvCount:
        raise InvCountError()
    if result == xlretInvXloper:
        raise InvXloperError()
    if result == xlretStackOvfl:
        raise StackOvflError()
    if result == xlretFailed:
        raise FailedError()
    if result == xlretUncalced:
        raise UncalcedError()
    if result == xlretNotThreadSafe:
        raise NotThreadSafeError()
    if result == xlretInvAsynchronousContext:
        raise InvAsynchronousContextError()
    raise ExcelError("unknown Excel error code")


class NoExcelError(RuntimeError):
    pass


def _is_excel():
    """figure out if this process is an excel process"""
    filename = ctypes.create_string_buffer(4096)
    ctypes.windll.kernel32.GetModuleFileNameA(0, filename, ctypes.sizeof(filename))
    return filename.value.endswith('\\EXCEL.EXE')

if not _is_excel():
    def Excel(*_):
        raise NoExcelError()
else:
    # Excel12v is really a wrapper  around MdCallBack12, which is (unconventionally)
    # exported from EXCEL.EXE, so we do our own thing to load up the module
    _EXCEL = ctypes.windll.kernel32.GetModuleHandleA(None)
    _MdCallback12 = ctypes.windll.kernel32.GetProcAddress(_EXCEL, "MdCallBack12")
    _MdCallBack12 = ctypes.WINFUNCTYPE(c_int)(_MdCallback12)
    _MdCallBack12.restype = c_int
    _MdCallBack12.argtypes = [c_int, c_int, POINTER(LPXLOPER12), LPXLOPER12]
    _MdCallBack12.errcheck = _xlret_errcheck

    def Excel12(xlfn, *args):
        """convert arguments to XLOPER12s and invoke Excel12 API"""
        # pylint: disable=W0142

        if len(args) > 255:
            raise ExcelError('Too many arguments for Excel12')
        if len(args) == 0:
            rgx = (LPXLOPER12 * 1)()
        else:
            opers = [arg if isinstance(arg, XLOPER12) else XLOPER12(arg) for arg in args]
            rgx = (LPXLOPER12 * len(args))(*[cast(pointer(xloper), LPXLOPER12) for xloper in opers])

        res = XLOPER12()
        _MdCallBack12(xlfn, len(args), rgx, cast(pointer(res), LPXLOPER12))

        if res.xltype in (xltypeStr, xltypeRef, xltypeBigData, xltypeMulti):
            res.xltype |= xlbitXLFree

        return res

    Excel = Excel12


    _log.info('Excel12 callback in use')

def _make_err(xlerr):
    res = XLOPER()
    res.xltype = xltypeErr
    res.val.err = xlerr
    return res

globals().update(
    [(name, _make_err(getattr(ExcelXLLSDK.gen.xlerr, name))) for name in ExcelXLLSDK.gen.xlerr.__all__]
)


class _Wrapper(object):
    """wrapper module to allow us to neatly invoke the excel API directly"""
    def __init__(self, wrapped):
        self.wrapped = wrapped

    def __getattr__(self, name):
        res = getattr(self.wrapped, name, None)
        if hasattr(self.wrapped, name):
            return getattr(self.wrapped, name)
        xlfn = getattr(ExcelXLLSDK.gen.xlcall, name)

        def _wrapper(*args):
            if _log.isEnabledFor(logging.DEBUG):
                _log.debug('%s(%s)' % (name, ', '.join((repr(arg) for arg in args))))
            return Excel(xlfn, *args)
        return _wrapper
sys.modules[__name__] = _Wrapper(sys.modules[__name__])
