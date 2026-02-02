from __future__ import absolute_import
# pylint: disable=C0103,E0602,R0903,W0614,W0401,W0212,W0621
import sys
import ctypes
import logging
from ctypes import c_int, POINTER, cast, pointer, byref

from ExcelXLLSDK.xltypes import (
    XLOPER4, LPXLOPER4, OPER4,
    XLOPER12, LPXLOPER12, OPER12
)
import ExcelXLLSDK.gen.xlerr
import ExcelXLLSDK.gen.xlcall


from ExcelXLLSDK.gen.xltype import *
from ExcelXLLSDK.gen.xlret import *

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
    xlver = 0    
else:
    _XLCallVer = ctypes.cdll.XLCALL32.XLCallVer
    _XLCallVer.restype = ctypes.c_int32
    _XLCallVer.argtypes = []
    try:
        xlver = {0x0500: 11, 0x0C00: 12}[_XLCallVer()]
    except KeyError:
        _log.warning('unrecognised excel version')
        xlver = 0

_log.info('xlver = %d', xlver)

if xlver >= 12:
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
    OPER = OPER12
    XLOPER = XLOPER12
    LPXLOPER = LPXLOPER12

    _log.info('Excel12 callback in use')

elif xlver >= 11:
    _Excel4v = ctypes.windll.XLCALL32.Excel4v
    _Excel4v.restype = c_int
    _Excel4v.argtypes = [c_int, LPXLOPER4, c_int, POINTER(LPXLOPER4)]
    _Excel4v.errcheck = _xlret_errcheck

    # convenient wrapper on Excel4 function
    def Excel4(xlfn, *args):
        """translate arguments to XLOPERS and invoke the Excel4 API"""
        # pylint: disable=W0142

        if len(args) > 30:
            raise ExcelError('Too many arguments for Excel4')

        if len(args) == 0:
            rgx = (LPXLOPER4 * 1)()
        else:
            opers = [arg if isinstance(arg, XLOPER4) else XLOPER4(arg) for arg in args]
            rgx = (LPXLOPER4 * len(opers))(*[cast(pointer(xloper), LPXLOPER4) for xloper in opers])

        res = XLOPER4()
        _Excel4v(xlfn, byref(res), len(args), rgx)

        if res.xltype in (xltypeStr, xltypeRef, xltypeBigData, xltypeMulti):
            res.xltype |= xlbitXLFree

        return res

    Excel = Excel4
    OPER = OPER4
    XLOPER = XLOPER4
    LPXLOPER = LPXLOPER4

    _log.info('Excel4 callback in use')

else:
    _log.info('Excel not found')

    def Excel(*_):
        raise NoExcelError()
    XLOPER = XLOPER12
    LPXLOPER = LPXLOPER12
    OPER = OPER12

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
