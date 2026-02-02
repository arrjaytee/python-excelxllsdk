import os
import sys
import runpy
import atexit

from ExcelXLLSDK.xll import XLLModule

xll = XLLModule()
DllMain = xll.DllMain

@xll(macro_sheet=True)
def _pyxcel_import_user():
    __import__('user')
    return 0

@xll(macro_sheet=True)
def _pyxcel_chdir(dir):
    os.chdir(dir.value)
    return 0

@xll(macro_sheet=True)
def _pyxcel_cmdline(argv=''):
    """run a python script or module if we have leading -m"""
    try:
        argv = eval(argv)

        if argv[0] == '-c':
            sys.argv = [argv[0]] + argv[2:]
            print 'i got ', sys.argv
            exec (argv[1], vars(sys.modules['__main__']))
        elif argv[0] == '-m':
            sys.argv = [__file__] + argv[2:]
            # run module seems to overwrite [0] with '', so put a dummy
            # entry in to avoid that.
            sys.path = [ '<dummy> '] + sys.path
            runpy.run_module(argv[1], run_name='__main__')
        else:
            sys.argv = [__file__] + argv[1:]
            runpy.run_path(argv[0], run_name='__main__')
        return 0
    except SystemExit as e:
        return e.code
    except:
        sys.excepthook(*sys.exc_info())
        return 1
    finally:
        # fake a call to atexit
        atexit._run_exitfuncs()



