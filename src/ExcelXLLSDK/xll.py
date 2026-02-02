from __future__ import absolute_import


import gc
import sys
import ctypes
import inspect
import logging

from types import GeneratorType, MethodType

import ExcelXLLSDK.XLCALL

from ctypes import WINFUNCTYPE, c_int, addressof
from pelib import PEExportDict

from ExcelXLLSDK.XLCALL import (
    OPER, XLOPER,
    OPER4, XLOPER4, LPXLOPER4, 
    OPER12, XLOPER12, LPXLOPER12,
    ExcelError, UncalcedError,
    xlfRegister, xlGetName,
    xlfGetWorkspace, xlCoerce,
    xlAsyncReturn, xlfCaller,
    xlEventRegister
    )

xleventCalculationEnded = 1
xleventCalculationCanceled = 2

from ExcelXLLSDK.gen.xltype import xlbitDLLFree, xlbitXLFree, xltypeInt
from ExcelXLLSDK.XLCALL import xlver

from ExcelXLLSDK._ctypes_win32 import (
    DisableThreadLibraryCalls, 
    DLL_PROCESS_ATTACH
)

#pylint: disable=R0903,C0111,C0103,W0232,C0301

_log = logging.getLogger(__name__)

_pxAutoFree = {}
_xlNone = XLOPER(None)
_xlTrue = XLOPER(True)
_xlFalse = XLOPER(False)


def _xlResult(value):        
    """
    convert a return value into an XLOPER, caching in a dict so that it
    is not garbage collected until we get the xlAutoFree callback.
    """
    if value is None: return addressof(_xlNone)
    if value is True: return addressof(_xlTrue)
    if value is False: return addressof(_xlFalse)
    
    if isinstance(value, GeneratorType):
        rows, columns = xlfCaller().size
        gen = value
        value = XLOPER()              
        value._set_Multi(rows, columns, gen)

    if not isinstance(value, XLOPER):
        value = XLOPER(value)              

    if value.xltype & xlbitXLFree:
        _log.warning("returning XL allocated data to excel? %d" % sys.getrefcount(value))
        return addressof(value)

    _pxAutoFree[addressof(value)] = value

    if not value.xltype & xlbitXLFree:
        value.xltype |= xlbitDLLFree

    return addressof(value)


def xlAutoFree(pxFree):
    del _pxAutoFree[pxFree]
    
def xlAutoFree12(pxFree):
    del _pxAutoFree[pxFree]


from ExcelXLLSDK.XLCALL import xlerrValue

# mappings from python buildins to ctypes codes
# can we map X to async result, and have a callback type there
_argtypes = {
    int: ('J', ctypes.c_long, lambda x: x),
    float: ('B', ctypes.c_double, lambda x: x),
    str: ('C', ctypes.c_char_p, lambda x: str(x)),
    unicode: ('C%', ctypes.c_wchar_p, lambda x: unicode(x)),
    bool: ('A', ctypes.c_short, lambda x: True if x else False),
    OPER4: ('P', ctypes.c_long, OPER4.from_address),
    XLOPER4: ('R', ctypes.c_long, XLOPER4.from_address),
    OPER12: ('Q', ctypes.c_long, OPER12.from_address),
    XLOPER12: ('U', ctypes.c_long, XLOPER12.from_address)
}

def _argtype(type):
    #  if it's a direct conversion, just use the converter in the map
    if type in _argtypes:
        return _argtypes[type]

    # otherwise treat as an XLOPER, and defer to the XLOPER.to function
    code, ctype, conv = _argtypes[XLOPER]
    return (code, ctype, lambda x : XLOPER.from_address(x).value)
 
class NullContext(object):
    """
    null context handler as a default
    """
    def __enter__(self):
        pass
    def __exit__(self, *args):
        pass

    def __call__(self, *args) :
        return self



class XLLFunction(object):
    def __init__(self, func, context):
        """
        setup initial xlfRegister parameters from the python function definition
        """
        # function to invoke - we can wrap this with logging, exception 
        # translation later on.
        self.func = func
        self.context = context

        # get nicely organised argument spec
        argspec = inspect.getargspec(func)

        # arguments to xlfRegister with sensible defaults
        # if it's a bound method, construct a better name?

        args = argspec.args

        if isinstance(func, MethodType):
            if not func.im_self:
                # or can we, and allow objects to be passed in as first argument
                raise ExcelError("cannot register an unbound method as an xll function")

            self.Procedure = "{1:08X}.{0.__name__}.".format(func, id(func.im_self))

            # hack off the first argspec
            args = args[1:]
        else:
            self.Procedure = "{0.__module__}.{0.__name__}.".format(func)

        self.FunctionText = func.__name__
        self.ArgumentText = ','.join(args)
        self.MacroType = 1
        self.Category = None
        self.FunctionHelp = inspect.getdoc(func) or '.'.join((func.__module__, func.__name__))
        self.HelpTopic = None
        self.ShortcutText = None

        # result is always an oper pointer
        self.result_TypeText = _argtype(OPER)[0]
       
        # infer types of arguments from python default values, and assume
        # OPER values for the rest
        types = map(type, argspec.defaults) if argspec.defaults else []
        types = [OPER]*(len(args) - len(types)) + types

        # store all argument details as arrays mapped to the func_
        self.TypeText = []
        self.argtypes = []
        self.argconvs = []
    
        # populate the initial array representing argument details with defaults
        for code, argtype, argconv in map(_argtype, types):
            self.TypeText.append(code)
            self.argtypes.append(argtype)
            self.argconvs.append(argconv)

        self.ArgumentHelp = [ t.__name__ for t in types ]

        # apply argument overrides from @arg decorators
        for name, info in getattr(func, 'xl_args', {}).iteritems():
            i = func.func_code.co_varnames.index(name)
            default = info.get('default', None)
            _type = info.get('type', type(default))
            help = info.get('help', None)

            if type or default:
                self.TypeText[i], self.argtypes[i], self.argconvs[i] = _argtype(_type)
            if help:
                self.ArgumentHelp[i] = help


    def __call__(self, *args):
        """
        invoke the XLL function as called by the xloper wrapper, this is the entry
        point for ctypes
        """
        try:
            # should we allow the context to bind to the args + argconvs here
            # so we get better tracing? probably.
            args = [argconv(arg) for arg, argconv in zip(args, self.argconvs)]
            with self.context(self, args):
                return _xlResult(self.func(*args))
        except UncalcedError as e:
            return 0
        except BaseException as e:
            sys.excepthook(*sys.exc_info())
            try:
                _log.exception(e)
                return _xlResult(e)
            except BaseException as e:
                return 0
        finally:
            # IMPORTANT: run the GC before returning to excel so that we deallocate
            # all XL supplied XLOPER data within the same UDF call, else excel will
            # do it instead, and we will crash when he GC is eventually called.
            gc.collect()

    @property
    def entry_point(self):
        """ 
        construct the ctypes callback to expose to excel
        """
        if not hasattr(self, '_procedure'):
            # important to hold a reference to the function type otherwise it is GC'd before use
            self._entry_point_type = WINFUNCTYPE(ctypes.c_void_p, *self.argtypes)
            self._entry_point = self._entry_point_type(self.__call__)
        return self._entry_point

def xlarg(name, default=None, type=None, help=None):
    def _wrapper(func):
        if not name in func.func_code.co_varnames:
            raise ExcelError('function does not have an argument named {0}'.format(name))
        func.__dict__.setdefault("xl_args", {})[name] = { 'default' : default, 'type' : type, 'help' : help }
        return func

    return _wrapper


class XLLModule(object):
    def __init__(self, Category=None, AddInManagerInfo=None, context=NullContext(), thread_safe=False):
        """
        setup xll module, default is to use this instance to implement 
        methods.
        """        
        self.hInstDll = None
        self.registered = []
        self.Category = Category
        self.AddInManagerInfo = AddInManagerInfo
        self.context = context
        self.thread_safe = thread_safe

    def register(self, func, name=None,
                 context=None,
                 category=None,
                 command=False,
                 hidden=False,
                 volatile=True,
                 thread_safe=None,
                 macro_sheet=False):
        if thread_safe == None:
            thread_safe = self.thread_safe

        reg = XLLFunction(func, context or self.context)

        if name: reg.FunctionText = name
        reg.Category = category or self.Category

        if hidden and command:
            raise ExcelError('cannot register a macro command as hidden')

        if hidden:
            reg.MacroType = 2
        elif command:
            reg.MacroType = 0
        else:
            reg.MacroType = 1

        if volatile:
            reg.TypeText.append('!')
        if thread_safe:
            reg.TypeText.append('$')
        if macro_sheet:
            reg.TypeText.append('#')

        if [r for r in self.registered if r.FunctionText == reg.FunctionText]:
            raise ExcelError('registered excel function with duplicate name: ' + reg.FunctionText)

        self.registered.append(reg)
        return reg

    def __call__(self, func=None, **kwargs):
        def _wrapper(func):
            self.register(func, **kwargs)
            return func

        if func:
            return _wrapper(func)

        return _wrapper

    def DllMain(self, hInstDll, fdwReason, lpvReserved):
        """
        DllMain function to import into embedded xll which sets 
        up xlAutoOpen 
        This allows the object to provide an entry point for the addin dll.
        """
        _log.debug( 'DllMain(0x%08X %d, 0x%08X)' %  (hInstDll, fdwReason, lpvReserved))

        if self.hInstDll is None:
            self.hInstDll = hInstDll
        
        if self.hInstDll != hInstDll:
            raise RuntimeError('XLLModule.DllMain called with wrong hInstDll') 

        # install the thunk - should this pass back tot he module? 
        if fdwReason == DLL_PROCESS_ATTACH:
            DisableThreadLibraryCalls(hInstDll)
            thunk = PEExportDict.from_handle(hInstDll, readonly=False)
                
            thunk['xlAutoAdd'] = WINFUNCTYPE(c_int)(self.xlAutoAdd)
            thunk['xlAutoClose'] = WINFUNCTYPE(c_int)(self.xlAutoClose)        
            thunk['xlAutoOpen'] = WINFUNCTYPE(c_int)(self.xlAutoOpen)            
            thunk['xlAutoRemove'] = WINFUNCTYPE(c_int)(self.xlAutoRemove)
            
            if xlver >= 12:
                thunk['xlAddInManagerInfo12'] = WINFUNCTYPE(c_int, LPXLOPER12)(self.xlAddInManagerInfo)
                thunk['xlAutoRegister12'] = WINFUNCTYPE(c_int, LPXLOPER12)(self.xlAutoRegister)
                thunk['xlAutoFree12'] = WINFUNCTYPE(None, c_int)(xlAutoFree12)
            else:
                thunk['xlAddInManagerInfo'] = WINFUNCTYPE(c_int, LPXLOPER4)(self.xlAddInManagerInfo)
                thunk['xlAutoRegister'] = WINFUNCTYPE(c_int, LPXLOPER4)(self.xlAutoRegister)
                thunk['xlAutoFree'] = WINFUNCTYPE(None, c_int)(xlAutoFree)
            
            self.thunk = thunk
        
        return 1

    def xleventCalculationEnded(self):
        _log.info("xleventCalculationEnded")


    def xleventCalculationCanceled(self):
        _log.info("xleventCalculationEnded")

    def xlAutoOpen(self):
        """
        callback from excel which should register the functions supplied
        to the decorators
        """

        ModuleName = xlGetName()

        _log.info("xlAutoOpen: %s", ModuleName)
        
        # setup the excel version number accurately - we can't use 
        # xlfGetWorkspace in the module setup.
        ExcelXLLSDK.XLCALL.version = float(xlfGetWorkspace(2))

        # register event handlers under custom uuids
        regCalculationEnded = self.register(self.xleventCalculationEnded, command=True)
        regCalculationCanceled = self.register(self.xleventCalculationCanceled, command=True)

        for reg in self.registered:
            if hasattr(reg, 'RegisterId'):
                _log.warning('{0} has already been registered'.format(reg.FunctionText))
            
            # install the dll export 
            self.thunk[reg.Procedure] = reg.entry_point
            TypeText = reg.result_TypeText+''.join(reg.TypeText)
            # NOTE if xlfRegister fails we get False, not and exception      
            # NOTE add space suffix to ArgumentHelp as excel removes the last char
            reg.RegisterId = xlfRegister(
                ModuleName,
                reg.Procedure,            
                TypeText,
                reg.FunctionText,
                reg.ArgumentText,
                reg.MacroType,
                reg.Category or self.Category,
                reg.ShortcutText,
                reg.HelpTopic,
                reg.FunctionHelp+' ',
                *[ arg+' ' if arg else arg for arg in reg.ArgumentHelp ]
            )
      
            # failed to register the function? 
            if not reg.RegisterId:
                raise ExcelError('xlfRegister of {0} failed'.format(reg.FunctionText))

            _log.info('={reg.FunctionText}({reg.ArgumentText}) [{TypeText}] -> {reg.func.__module__}.{reg.func.func_name}'.format(reg=reg, TypeText=TypeText))

        # use excel to tell us the right dll to register under - should match
        # with self.hInstDll
        global xleventCalculationCanceled, xleventCalculationEnded
        xlEventRegister(regCalculationEnded.Procedure, xleventCalculationEnded)
        xlEventRegister(regCalculationCanceled.Procedure, xleventCalculationCanceled)

        return 1
    
    def xlAutoRegister(self, pxName):
        """
        unused
        """
        _log.info("xlAutoRegister")
        return 0

    def xlAutoClose(self):
        """called when the XLL is unloaded
        
        however if the shutdown is aborted by a save/cancel then we won't be opened again. 
        so we shouldn't do anything here.
        """
        _log.info("xlAutoClose: %s", str(xlGetName()))
        return 1

    def xlAutoAdd():
        """
        called when addin is selected from the addin list
        """
        _log.info("xlAutoAdd: %s", str(xlGetName()))
        return 1

    def xlAutoRemove():    
        """
        called when the addin is deselected from the addin list
        """        
        _log.info("xlAutoRemove: %s", str(xlGetName()))
        return 1

    def xlAddInManagerInfo(self, pxAction):
        """
        return information for the addin manager to display
        """
        _log.info("xlAddInManagerInfo: %s", str(xlGetName()))

        # convert to an integer - not sure this is working
        action = xlCoerce(pxAction.contents, xltypeInt)
        
        if int(action) == 1:
            res = self.AddInManagerInfo or repr(self)
            _log.info("xlAddInManagerInfo(%s) = %s" % (repr(pxAction.contents), repr(res)))
            return _xlResult(res)

        # should work out the pacakge + versio for the entry point here?
          
        _log.warning("Ignoring xlAddInManagerInfo: %s", repr(pxAction.contents))
        return addressof(xlerrValue)



__all__ = [
    'XLLModule',
    'xlarg',
]
     