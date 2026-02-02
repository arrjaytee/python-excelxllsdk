import sys
import random
from mock import patch

import ExcelXLLSDK.unittest 
import ExcelXLLSDK.XLCALL

from ..gen.xlerr import xlerrNum

from unittest import skipIf

from ExcelXLLSDK.xll import *
from ExcelXLLSDK.XLCALL import OPER, XLOPER, xlver

xll = XLLModule(
    Category='ExcelXLLSDK Test',
    AddInManagerInfo="Test XLL Functions"
    )

_magic = random.random()
_volatile = 0
_nonvolatile = 0

@xll(volatile=True)
def _Volatile():
    global _volatile    
    _volatile += 1    
    return _magic

@xll(volatile=False)
def _NonVolatile():
    global _nonvolatile
    _nonvolatile += 1    
    return _magic

@xll(thread_safe=True)
def _ThreadSafe():
    return _magic

@xll(macro_sheet=True)
def _MacroSheet():
    return _magic

@xll
def _RaiseRuntimeError():
    raise RuntimeError('runtime error')

@xll
def _bool(x=False):
    return x, type(x).__name__

@xll
def _int(x=1):
    return x, type(x).__name__

@xll 
def _float(x=123.123):
    return x, type(x).__name__

@xll
def _str(x='fish'):
    return x, type(x).__name__

@xll
def _unicode(x=u'foreignfish'):
    return x, type(x).__name__

@xll 
def _OPER(x=OPER()):
    return x.value, type(x).__name__

@xll
def __XLOPER(x=XLOPER()):
    return repr(x), type(x).__name__

@xll
def _object(x=object()):
    return repr(x), type(x).__name__

#  _async = 0

# @xll
# @asynchronous
# def test_Async(seconds, _callback=None):            
#     global _async
#     def _wait(seconds):   
#         global _async                
#         if _async:
#             time.sleep(seconds)
#             _async  = _async +  1    
#             _callback(_magic)

#     import threading
#     thread = threading.Thread(target=_wait, args=(seconds.value, ))
#     thread.start()

DllMain = xll.DllMain

_eval = ExcelXLLSDK.XLCALL.xlfEvaluate              

class RegisterTestCase(ExcelXLLSDK.unittest.WorkbookTestCase):
    """verify that volatile functions get recalculated on every
    calculation cycle, and nonvolatile ones don't"""

    def test_excepthook(self):        
        """verify that excepthook is called"""
        with patch('sys.excepthook'):                        
            xlo = _eval('_RaiseRuntimeError()')
            self.assertEqual(xlo, XLOPER.from_err(xlerrNum))
            self.assertTrue(sys.excepthook.called)
            
    def test_excepthook_fail(self):
        """verify that if our excepthook breaks, then we still give an error"""
        def _excepthook(*args):
            raise RuntimeError("sys.excepthook failed")

        with patch("sys.excepthook", new=_excepthook):            
            xlo = _eval('_RaiseRuntimeError()')
            self.assertEqual(xlo, XLOPER.from_err(xlerrNum))

    def test_RegisteredFunctions(self):
        from exceltools.registry import _find_xll
        loaded = set((xll for xll, _, _ in self.Application.RegisteredFunctions()))
        for xll in map(_find_xll, ['ExcelXLLSDK_test']):
            self.assertIn(xll, loaded)

    @skipIf(True, "AddIns2 doesn't appear to work?")
    def test_Addins2(self):
        """check excel thinks everything is loaded"""
        if xlver < 12:
            raise SkipTest("need excel 2007 or above for AddIns2")

        from exceltools.registry import _find_xll

        # chekc all entry points are loaded?
        # addins = [ addin.FullName for addin in self.Application.Addins2 ]
        # for xll in map(_find_xll, self.__xlls__):
        #     self.assertIn(xll, addins)

        for addin in self.Application.Addins2:
            if addin.Name == 'ExcelXLLSDK_test.xll':
                self.assertEqual(addin.Title, 'ExcelXLLSDK.test._test-script.py')

    @skipIf(True, "Volatile broken? No sheets are loading...")
    def test_Volatile(self):
        global _volatile, _nonvolatile
        _volatile, _nonvolatile = (0, 0)

        self.Application.CalculateFull()
        self.assertEqual(_volatile, 1)
        self.assertEqual(_nonvolatile, 1)

        self.Application.Calculate()
        self.assertEqual(_volatile, 2)
        self.assertEqual(_nonvolatile, 1)

        self.Application.CalculateFull()
        self.assertEqual(_volatile, 3)
        self.assertEqual(_nonvolatile, 2)

    @skipIf(True, "Async not yet supported?")
    def test_Async(self):
        global _async
        _async = 1
        
        before = time.time()
        self.Application.CalculateFullRebuild()
        after = time.time()
        self.assertEqual(_async, 3)
        print 'recalculation took {0:f}'.format(after - before)
        self.assertTrue(after - before, 5.0)
              
    def test_Register(self):        
        global _magic, _eval
          
        self.assertEqual(_eval('_Volatile()').value, _magic)
        self.assertEqual(_eval('_NonVolatile()').value, _magic)
        self.assertEqual(_eval('_ThreadSafe()').value, _magic)
        self.assertEqual(_eval('_MacroSheet()').value, _magic)
        with self.assertRaises(TypeError):            
            self.assertEqual(_eval('_RaiseRuntimeError()').value, _magic)
        

    