from __future__ import absolute_import
import logging

from ExcelXLLSDK.xll import xlarg, XLLModule
from ExcelXLLSDK.logging import LogContext
from ExcelXLLSDK.XLCALL import XLOPER

xll = XLLModule(Category="Python Builtins", context=LogContext(logging.getLogger(__name__)))
DllMain = xll.DllMain

@xll(hidden=True)
def interact():
    """
    startup a python interpeter and return when it has finished
    """
    import code
    return code.interact()

@xll(name="repr")
@xlarg("Object", type=XLOPER, help='object to convert to python string representation')
def _repr(Object):
    """Return a string representing the argument, and dump it to stdout

    Returns python __repr__ of the argument
    """    
    res = repr(Object)  
    return res


@xll(name='repr.value')
@xlarg("Object", type=XLOPER, help='object to convert to python string representation')
def _repr(Object):
    """Return a string representing the argument, and dump it to stdout

    Returns python __repr__ of the argument
    """
    res = repr(Object.value)
    return res

@xll(name='eval')
@xlarg('source', help='python code to evaluate')
def _eval(source='', _1=XLOPER()):
    """Evaluate source as python expression and return its XLOPER representation

    Returns xloper representation of the result
    """
    return eval(str(source), locals(), globals())



