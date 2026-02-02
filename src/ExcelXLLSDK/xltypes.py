from __future__ import absolute_import

from multimethod import multimethod

from ctypes import *
from ctypes.util import find_msvcrt
from .wintypes import *
from .gen.xltype import *
from .gen.xlerr import *
from .gen.XLCALL15 import (
    XLREF, XLREF12, XLOPER as _XLOPER4, XLOPER12 as _XLOPER12
)
from .gen.xlcall import (xlFree, xlSheetNm, xlCoerce)
from datetime import datetime, timedelta

# we make use of lots of excel constants in here
from .gen import xltype as _xltypes

_EXCEL_TIME_ZERO = datetime(1899, 12, 30, 0, 0, 0)

_xltypeMask = ~(xlbitDLLFree | xlbitXLFree)
_xltypeAtoms = xltypeNil | xltypeBool | xltypeInt | xltypeNum | xltypeStr

def _xltypeName(xltype):
    _xltype = xltype & _xltypeMask

    suffix = ""
    if xltype & xlbitDLLFree:
        suffix += '|DLL'
    if xltype & xlbitXLFree:
        suffix += '|XL'

    for key, value in vars(_xltypes).iteritems():
        if key.startswith('xltype') and value == _xltype:
            return key + suffix
    return "0x{:04X}".format(_xltype) + suffix

def _malloc_errcheck(result, _, __):
    """capture 0 pointers from malloc and convert to exceptions"""
    if result == None:
        raise MemoryError('malloc: failed to allocate memory')
    return result

_crt = ctypes.CDLL(find_msvcrt())
_malloc = CFUNCTYPE(c_void_p, c_size_t)(('malloc', _crt))
_malloc.errcheck = _malloc_errcheck
_free = CFUNCTYPE(None, c_void_p)(('free', _crt))

def _colref(col):
    """convert column number to column letter"""
    div, mod = divmod(col, 26)
    return chr(ord('A') + mod) if div == 0 else _colref(div) + _colref(mod)

class _XLOPER(object):
    def __init__(self, value=None):  # pylint: disable=W0231        
        self._set(value)

    def is_ref(self):
        return self.xltype&(xltypeRef|xltypeSRef)

    @classmethod
    def from_err(cls, err):
        xlo = cls()
        xlo.xltype = xltypeErr
        xlo.val.err = err
        return xlo

    def __eq__(self, value):
        xltype = self.xltype & _xltypeMask
        if value is None:
            return xltype == xltypeNil or xltype == xltypeMissing

        if isinstance(value, self.__class__):
            if xltype != (value.xltype & _xltypeMask):
                return False
            if xltype == xltypeErr:
                return self.val.err == value.val.err
            if xltype == xltypeBool:
                return self.val.xbool == value.val.xbool
            if xltype == xltypeInt:
                return self.val.w == value.val.w
            if xltype == xltypeNum:
                return self.val.num == value.val.num
            if xltype == xltypeStr:
                return self.val._get_Str() == value.val._get_Str()

        if xltype == xltypeErr or xltype == xltypeNil:
            return False

        # last ditch, extract value and compare
        return self.value == value

    def _get(self):  # pylint: disable=R0911
        xltype = self.xltype & _xltypeMask
        if xltype == xltypeNil:
            return None
        elif xltype == xltypeMissing:
            return None  # python convention - missing parameters are none
        elif xltype == xltypeBool:
            return True if self.val.xbool else False
        elif xltype == xltypeInt:
            return self.val.w
        elif xltype == xltypeNum:
            return self.val.num
        elif xltype == xltypeStr:
            return self._get_Str()
        elif xltype == xltypeMulti:
            return [
                tuple([self._cell_Multi(i, j).value for j in xrange(self.val.array.columns)])
                for i in xrange(self.val.array.rows)
            ]
        elif xltype == xltypeSRef or xltype == xltypeRef:
            import ExcelXLLSDK.XLCALL
            return ExcelXLLSDK.XLCALL.Excel(xlCoerce, self)._get()
        elif xltype == xltypeErr and self.val.err == xlerrNA:
            return None
        else:
            raise TypeError('cannot convert %s to a python type' % _xltypeName(self.xltype))

    def _cell_Multi(self, i, j):
        """ indexed access to xltypeMulti conntents"""
        if i < 0 or i >= self.val.array.rows or j < 0 or j >= self.val.array.columns:
            raise IndexError('xltypeMulti index out of range')
        res = self.__class__.from_address(addressof(self.val.array.lparray[i * self.val.array.columns + j]))
        return res

    def __getitem__(self, index):
        if self.xltype & _xltypeMask != xltypeMulti:
            raise TypeError("'%s:%s' object cannot be indexed" %
                           (self.__class__.__name__, _xltypeName(self.xltype)))
        return self._cell_Multi(*index).value  # pylint: disable=W0142




    def _iter_SRef(self):
        ref = self.val.sref.ref
        for col in xrange(ref.colFirst, ref.colLast + 1):
            for rw in xrange(ref.rwFirst, ref.rwLast + 1):
                xlo = self.__class__()
                xlo.xltype = xltypeSRef
                xlo.val.sref.ref.rwFirst = rw
                xlo.val.sref.ref.rwLast = rw
                xlo.val.sref.ref.colFirst = col
                xlo.val.sref.ref.colLast = col
                yield xlo


    def _iterrefs(self):
        _type = (self.val.mref.lpmref.contents.reftbl._type_ * self.val.mref.lpmref.contents.count)
        refs = _type.from_address(addressof(self.val.mref.lpmref.contents.reftbl))
        for ref in refs:
            yield ref

    def _iter_Ref(self):
        for ref in self._iterrefs():
            for col in xrange(ref.colFirst, ref.colLast + 1):
                for rw in xrange(ref.rwFirst, ref.rwLast + 1):
                    xlo = self.__class__()
                    xlo._set_Ref(1)
                    xlo.val.mref.idSheet = self.val.mref.idSheet
                    reftbl = xlo.val.mref.lpmref.contents.reftbl
                    reftbl[0].rwFirst = rw
                    reftbl[0].rwLast = rw
                    reftbl[0].colFirst = col
                    reftbl[0].colLast = col
                    yield xlo

    # TODO iterkeys, itervalues: keys defined as leftmost column of strings, with empty cells below
    # values are the cells to the right hand side of them. also implement __getitem__, __setitem__ for indexing?
    # what does __setitem__ on a range do? perhaps that is a better way to xlSet?

    def iterrows(self):
        """
        loop over references to all the rows in this multi reference
        """
        xltype = self.xltype & _xltypeMask
        if xltype == xltypeMulti:
            for ref in self._iterrefs():
                for rw in xrange(ref.rwFirst, ref.rwLast + 1):
                    xlo = self.__class__()
                    xlo._set_Ref(1)
                    xlo.val.mref.idSheet = self.val.mref.idSheet
                    reftbl = xlo.val.mref.lpmref.contents.reftbl
                    reftbl[0].rwFirst = rw
                    reftbl[0].rwLast = rw
                    reftbl[0].colFirst = ref.colFirst
                    reftbl[0].colLast = ref.colLast
                    yield xlo


    def itercols(self):
        """
        loop over references to all the columns in this multi reference
        """
        xltype = self.xltype & _xltypeMask
        if xltype == xltypeMulti:
            for ref in self._iterrefs():
                for col in xrange(ref.colFirst, ref.colLast + 1):
                    xlo = self.__class__()
                    xlo._set_Ref(1)
                    xlo.val.mref.idSheet = self.val.mref.idSheet
                    reftbl = xlo.val.mref.lpmref.contents.reftbl
                    reftbl[0].rwFirst = ref.rwFirst
                    reftbl[0].rwLast = ref.rwLast
                    reftbl[0].colFirst = col
                    reftbl[0].colLast = col
                    yield xlo

    def __iter__(self):
        xltype = self.xltype & _xltypeMask
        if xltype == xltypeMulti:
            buftype = ((self.__class__ * self.val.array.columns) * self.val.array.rows)
            buf = cast(self.val.array.lparray, POINTER(buftype)).contents
            return iter(buf)
        elif xltype == xltypeSRef:
            return self._iter_SRef()
        elif xltype == xltypeRef:
            return self._iter_Ref()
        else:
            raise TypeError("'%s:%s' object cannot be iterated" %
                           (self.__class__.__name__, _xltypeName(self.xltype)))

    def _set_Multi(self, rows, columns, _iter):
        rows = max(rows, 0)
        columns = max(columns, 0)

        self.val.array.rows = rows
        self.val.array.columns = columns

        size = sizeof(self.val.array.lparray._type_) * self.val.array.rows * self.val.array.columns
        ptr = _malloc(size)
        ctypes.memset(ptr, 0, size)
        self.val.array.lparray = cast(ptr, type(self.val.array.lparray))
        self.xltype = xltypeMulti
        for i in xrange(0, self.val.array.rows):
            for j in xrange(0, self.val.array.columns):
                cell = self._cell_Multi(i, j)
                cell.xltype = xltypeErr
                cell.val.err = xlerrNA
                
        for i, ii in zip(xrange(0, self.val.array.rows), _iter):
            for j, v in zip(xrange(0, self.val.array.columns), ii):
                cell = self._cell_Multi(i, j)
                cell._set(v)


    def _set(self, value):  # pylint: disable=R0912
        """assign a python value to the xloper, allocate memory if necessary"""
        self._del()
        from_value(self, value)
        if self.xltype == 0:
            raise TypeError("could not assign type for " + str(type(value)))

    def _del(self):
        """clear up manually allocated memory"""
        xltype = self.xltype & _xltypeMask
        if xltype == xltypeStr:
            self._del_Str()
        elif xltype == xltypeRef:
            ptr = cast(self.val.mref.lpmref, c_void_p)
            _free(ptr)
        elif xltype == xltypeMulti:
            for i in xrange(0, self.val.array.rows):
                for j in xrange(0, self.val.array.columns):
                    self._cell_Multi(i, j)._del()
            _free(cast(self.val.array.lparray, c_void_p))

    value = property(fget=_get, fset=_set, fdel=_del)
    
    def _get_datetime(self):
        xltype = self.xltype & _xltypeMask
        # if it's a string, parse it? 
        if xltype == xltypeNum:  
            return _EXCEL_TIME_ZERO + timedelta(days=self.val.num)
        elif xltype == xltypeSRef or xltype == xltypeRef:
            import ExcelXLLSDK.XLCALL
            return ExcelXLLSDK.XLCALL.Excel(xlCoerce, self)._get_datetime()
        else:
            raise TypeError('cannot convert %s to a python datetime' % _xltypeName(self.xltype))

    datetime = property(fget=_get_datetime)


    def __del__(self):
        if self._b_needsfree_:
            if self.xltype & xlbitXLFree:
                # clear the XLFree bit, so we are as if we had been returned from XL
                self.xltype &= ~xlbitXLFree
                self._xlFree()

            elif self.xltype & xlbitDLLFree:
                self._del()
            # blank this, in case we are called twice
            self.xltype = 0

    def __nonzero__(self):
        xltype = self.xltype & _xltypeMask
        if xltype == xltypeErr:
            return False
        elif xltype == xltypeMulti:
            return self.val.array.rows and self.val.array.columns
        elif xltype == xltypeSRef:
            return True
        elif xltype == xltypeRef:
            return self.val.mref.count
        else:
            return True if self.value else False



    @property
    def rows(self):
        xltype = self.xltype & _xltypeMask
        if xltype == xltypeMulti:
            return self.val.array.rows
        elif xltype == xltypeSRef:
            return self.val.sref.ref.rwLast - self.val.sref.ref.rwFirst
        else:
            raise TypeError('cannot get length of %s', _xltypeName(self.xltype))

    @property 
    def size(self):
        xltype = self.xltype & _xltypeMask
        if xltype == xltypeMulti:
            return (self.val.array.rows, self.val.array.columns)
        elif xltype == xltypeSRef:
            ref = self.val.sref.ref        
        else:
            raise TypeError('cannot get length of %s', _xltypeName(self.xltype))

        return (ref.rwLast - ref.rwFirst + 1, ref.colLast - ref.colFirst + 1)

    def __int__(self):
        xltype = self.xltype & _xltypeMask
        if self.xltype & _xltypeAtoms:
            return int(self.value)
        raise TypeError('cannot convert %s to int', _xltypeName(self.xltype))

    def __float__(self):
        if self.xltype & _xltypeAtoms:
            return float(self.value)
        raise TypeError('cannot convert %s to float', _xltypeName(self.xltype))

    def _repr_ref(self, ref):
        res = _colref(ref.colFirst) + str(int(ref.rwFirst) + 1)
        if ref.rwFirst != ref.rwLast or ref.colFirst != ref.colLast:
            res += ':' + _colref(ref.colLast) + str(int(ref.rwLast) + 1)
        return res

    def _repr_mref(self, mref):
        buf = (mref.reftbl._type_ * mref.count).from_address(addressof(mref.reftbl))
        return ','.join([self._repr_ref(ref) for ref in buf])

    def __str__(self):
        return unicode(self).encode("utf-8")

    def __unicode__(self):  # pylint: disable=R0911,R0912
        """return a string reprentation of the XLOPER

        we use an excel convention for formatting, as you would type
        the value in an excel formula.
        """
        xltype = self.xltype & _xltypeMask
        if xltype == xltypeMissing:
            return "Missing"
        elif xltype == xltypeNil:
            return "Nil"
        elif xltype == xltypeBool:
            return "TRUE" if self.val.xbool else "FALSE"
        elif xltype == xltypeInt:
            return str(int(self.val.w))
        elif xltype == xltypeNum:            
            return str(float(self.val.num))
        elif xltype == xltypeStr:
            return self._get_Str()
        elif xltype == xltypeMulti:
            return '{ ' + '; '.join(', '.join((repr(cell) for cell in row)) for row in self) + ' }'            
        elif xltype == xltypeSRef:            
            return '%s!%s' % (self._xlSheetNm(), self._repr_ref(self.val.sref.ref))
        elif xltype == xltypeRef:
            return '%s!%s' % (self._xlSheetNm(), self._repr_mref(self.val.mref.lpmref.contents))
        elif xltype == xltypeErr:
            if self.val.err == xlerrValue:
                return '#VALUE!'
            elif self.val.err == xlerrName:
                return '#NAME?'
            elif self.val.err == xlerrNum:
                return '#NUM!'
            elif self.val.err == xlerrNull:
                return '#NULL!'
            elif self.val.err == xlerrNA:
                return '#N/A'
            elif self.val.err == xlerrRef:
                return '#REF!'
            elif self.val.err == xlerrDiv0:
                return '#DIV/0!'
            else:
                return '#UNKNOWN!'
        else:
            raise TypeError('repr: cannot represent %s' % _xltypeName(self.xltype))

    def __repr__(self):
        """debugging description of including full type information"""       
        xltype = self.xltype & _xltypeMask
        # repr strings with the right quotes        
        if xltype == xltypeStr:
            return '"' + self._get_Str() + '"'
        try:
            return unicode(self)
        except TypeError:
            return '<%s at 0x%08X>' % (_xltypeName(self.xltype), addressof(self))
  


class XLOPER4(Structure, _XLOPER):
    _fields_ = _XLOPER4._fields_

    # TODO move _Set stuff that is XLOPER4 specific to here.

    def __init__(self, value=None):
        _XLOPER.__init__(self, value)

    def _xlFree(self):
        import ExcelXLLSDK.XLCALL
        ExcelXLLSDK.XLCALL.Excel4(xlFree, self)

    def _xlSheetNm(self):
        """callback to figure out sheet name"""
        import ExcelXLLSDK.XLCALL
        return ExcelXLLSDK.XLCALL.Excel4(xlSheetNm, self)

    def _set_Ref(self, count=1):
        # how to allocate this? need to assign to the lpmref buffer
        class _type(ctypes.Structure):
            _fields_ = [('count', WORD), ('reftbl', XLREF * count)]

        size = sizeof(_type)
        ptr = _malloc(size)
        ctypes.memset(ptr, 0, size)
        self.val.mref.lpmref = cast(ptr, type(self.val.mref.lpmref))
        self.val.mref.lpmref.contents.count = count
        self.xltype = xltypeRef

    def _get_Str(self):
        # how to extact the .val.str pointer? using the struct member
        # just gives us the value, which is not what we want.
        _len = POINTER(c_ubyte).from_address(addressof(self)).contents.value

        class _type(Structure):
            _fields_ = [('len', c_ubyte), ('str', c_char * _len)]

        ptr = POINTER(_type).from_address(addressof(self)).contents
        return ptr.str

    def _set_Str(self, value):
        if isinstance(value, unicode):
            return self._set_Str(value.encode('ascii'))

        if len(value) > 0xFF:
            raise ValueError('xltypeStr: string must be <= %d characters long' % 0xFF)

        class _type(Structure):
            _fields_ = [('len', c_ubyte), ('str', c_char * len(value))]

        ptr = _type.from_address(_malloc(sizeof(_type)))
        ptr.len = len(value)
        ptr.str = value

        self.val.str = addressof(ptr)
        self.xltype = xltypeStr

    def _del_Str(self):
        _free(cast(addressof(self), POINTER(c_void_p)).contents)


LPXLOPER4 = POINTER(XLOPER4)

class OPER4(XLOPER4):
    """
    derive an OPER class (values only) so we can discriminate
    between the types when specifying arguments to functions
    """
    pass

class XLOPER12(Structure, _XLOPER):
    _fields_ = _XLOPER12._fields_

    def __init__(self, value=None):
        _XLOPER.__init__(self, value)

    def _xlFree(self):
        import ExcelXLLSDK.XLCALL
        ExcelXLLSDK.XLCALL.Excel12(xlFree, self)

    def _xlSheetNm(self):
        import ExcelXLLSDK.XLCALL
        return ExcelXLLSDK.XLCALL.Excel12(xlSheetNm, self)

    def _set_Ref(self, count=1):
        # how to allocate this? need to assign to the lpmref buffer
        class _type(ctypes.Structure):
            _fields_ = [('count', WORD), ('reftbl', XLREF12 * count)]

        size = sizeof(_type)
        ptr = _malloc(size)
        ctypes.memset(ptr, 0, size)
        self.val.mref.lpmref = cast(ptr, type(self.val.mref.lpmref))
        self.val.mref.lpmref.contents.count = count
        self.xltype = xltypeRef

    def _get_Str(self):
        _len = POINTER(c_ushort).from_address(addressof(self)).contents.value

        class _type(Structure):
            _fields_ = [('len', c_ushort), ('str', c_wchar * _len)]

        ptr = POINTER(_type).from_address(addressof(self)).contents
        return ptr.str

    def _set_Str(self, value):
        if len(value) > 0xFFFF:
            raise ValueError('xltypeStr: string must be <= %d characters long' % 0xFFFF)

        class _type(Structure):
            _fields_ = [('len', c_ushort), ('str', c_wchar * len(value))]

        ptr = _type.from_address(_malloc(sizeof(_type)))
        ptr.len = len(value)
        ptr.str = value

        self.val.str = addressof(ptr)
        self.xltype = xltypeStr

    def _del_Str(self):
        _free(cast(addressof(self), POINTER(c_void_p)).contents)

LPXLOPER12 = POINTER(XLOPER12)

class OPER12(XLOPER12):
    """
    OPER type which permits values only
    """
    pass


@multimethod(_XLOPER, object)
def from_value(self, value):
    raise TypeError('cannot convert object to XLOPER')

# TODO improve exception mapping - build an xlerrException,
import ExcelXLLSDK._from
import ExcelXLLSDK._from_exception



