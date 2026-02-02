import unittest
from unittest import skipIf, skipUnless
from ExcelXLLSDK.XLCALL import xlver, _xlret_errcheck
from ExcelXLLSDK.gen.xltype import xltypeNum

import ExcelXLLSDK.gen.xlret as xlret

from ExcelXLLSDK.XLCALL import (
    Excel,
    ExcelError,
    NoExcelError,
    InvXlfnError,
    InvCountError,
    FailedError,
    AbortError,
    InvXloperError,
    InvAsynchronousContextError,
    StackOvflError,
    NotThreadSafeError,
    xlCoerce,
    xlfGetWorkspace,
    xlAsyncReturn,
    XLOPER,
)


class XLCALLTestCase(unittest.TestCase):
    @skipUnless(xlver, "requires excel")
    def test_xlret(self):
        """verify that the xlret return code translation works via the API"""
        with self.assertRaises(InvXlfnError):
            Excel(0xFFFF)
        with self.assertRaises(InvCountError):
            print(xlfGetWorkspace())
        with self.assertRaises(FailedError):
            xlCoerce("fish & chips", xltypeNum)
        with self.assertRaises(InvXloperError):
            xlo = XLOPER()
            xlo.xltype = 0xDEAD
            xlCoerce(xlo, xltypeNum)

    def test_xlret_weak(self):
        """test a few xlret codes that are hard to mock"""
        with self.assertRaises(AbortError):
            _xlret_errcheck(xlret.xlretAbort, None, None)
        with self.assertRaises(StackOvflError):
            _xlret_errcheck(xlret.xlretStackOvfl, None, None)
        with self.assertRaises(NotThreadSafeError):
            _xlret_errcheck(xlret.xlretNotThreadSafe, None, None)
        with self.assertRaises(ExcelError):
            _xlret_errcheck(0xDEAD, None, None)

    @skipUnless(xlver >= 14, "asynchronous functions are only in excel 2010")
    def test_Async(self):
        with self.assertRaises(InvAsynchronousContextError):
            xlAsyncReturn(XLOPER(), 123)

    @skipUnless(xlver >= 12, "requires excel12 API")
    def test_args12(self):
        with self.assertRaises(ExcelError):
            Excel(0, *([None] * 256))

    @skipUnless(xlver == 11, "requires excel4 API")
    def test_args4(self):
        with self.assertRaises(ExcelError):
            Excel(0, *([None] * 31))

    @skipIf(xlver, "must not have excel")
    def test_NoExcelError(self):
        with self.assertRaises(NoExcelError):
            Excel(123)
