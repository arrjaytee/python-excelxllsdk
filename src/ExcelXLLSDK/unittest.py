"""framework for expressing assertions and test cases on excel sheets"""
from __future__ import absolute_import

import os, sys, unittest

from os.path import splitext
from glob import glob

from ExcelXLLSDK.xll import xlarg, XLLModule
from unittest import TestCase
from ExcelXLLSDK.com import find_Application
from ExcelXLLSDK.com.Excel import xlCalculationManual

from ExcelXLLSDK.gen.xltype import xltypeRef

from ExcelXLLSDK.XLCALL import (
    xlfCaller, xlfRegister,
    xlCoerce, xlcOpen, xlcNew,
    xlfUnregister,
    xlfDocuments,
    xlcActivate,
    xlfGetCell,
    xlcCloseAll,
    xlcOptionsCalculation,
    xlver,
    ExcelError,
    XLOPER
)

def assertCell(cell, caller=None):    
    value = xlCoerce(cell)    
    if not value:
        raise AssertionError("%s : %s" % (str(caller), repr(value)))

class RangeTestCase(TestCase):
    def __init__(self, caller, asserts):
        self.caller = caller
        self.asserts = asserts
        TestCase.__init__(self)

    def runTest(self):        
        for cell in self.asserts:
            assertCell(cell, cell)

    def shortDescription(self):
        return None

    def __str__(self):
        return "%s" % str(self.caller)

    def __repr__(self):
        return "<RangeTestCase %s>" % str(self)

# place to send output if we are already part of a test run
_devnull = open(os.devnull, "w")
_result = None

xll = XLLModule(Category='Python Unit Test')
DllMain = xll.DllMain

# can we figure out nose.main here? or somehow supplant the macros?
@xll
@xlarg('Asserts', type=XLOPER, help='range containing assertion results')
def TESTCASE(Asserts):
    """build a test case from all the input cells, treating each as a
    separate assertion.

    Returns TRUE if all assertions pass, #VALUE! if not
    """    
    # ensure assertions are all calculated first
    xlCoerce(Asserts)

    # build a test case from the input range
    # TODO nose doesn't like this hack - we should simply
    # run the assertion against the current workbook testcase,     
    testcase = RangeTestCase(
        xlCoerce(xlfCaller(), xltypeRef),
        xlCoerce(Asserts, xltypeRef),
    )

    # run the test against the test runner - or would it be better
    # just to rememb er the range and run it later? so testcase == AND
    if not _result is None:    
        testcase.run(result=_result)        
    
    # also run in our own running, and generate text output
    # is it ok to do it twice - only checking cells, so fine.
    runner = unittest.TextTestRunner(stream=_devnull if _result else sys.stderr)
    res = runner.run(testcase)
    return res.wasSuccessful()

@xll
@xlarg('Condition', type=XLOPER, help='Returns True if passed, #VALUE! if failed')
def ASSERT(Condition):
    """Verify that Condition is True, return an error if not"""
    assertCell(Condition, xlfCaller())
    return True

from exceltools.registry import _find_xll

class WorkbookTestCase(TestCase):
    @classmethod
    def setUpClass(cls):
        if not xlver:
            return

        xlcNew(1)
        cls.Application = find_Application(os.getpid())
        xlcCloseAll()

        # can we use packcage_listdir here?  should we use the
        # class or the module name?
        _dir = splitext(sys.modules[cls.__module__].__file__)[0]
        wbks = glob(_dir + '/*.xls') + glob(_dir + '/*.xlsx')

        for wbk in wbks:
            if not xlcOpen(wbk, 0, True):
                raise ExcelError('could not open "%s"' % wbk)

        if cls.Application.Workbooks.Count == 0:
            raise RuntimeError("could not find any workbooks")

    # reset stuff for each test case?
    def setUp(self):
        if not xlver:
            self.skipTest("not running under excel")

        self.Application.Calculation = xlCalculationManual

    def run(self, result=None):
        """invoke excel to calculate all open sheets, and gather results into this test"""
        global _result
        _result = result
        try:
            return super(WorkbookTestCase, self).run(result)
        finally:
            _result = None


    def runTest(self):
        # this is not recalculating the #NAME! cells
        self.Application.CalculateFull()

    @classmethod
    def tearDownClass(cls):
        if not xlver:
            return
        xlcCloseAll()
        del cls.Application

