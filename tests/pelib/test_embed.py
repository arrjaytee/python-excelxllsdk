import os.path, sys

from pelib import PEExportDict
import ctypes
import unittest

from pkg_resources import resource_filename


# problem here is that tests can be run aside from the entry points
# how do we find the right entry point? 

class EmbedTestCase(unittest.TestCase):
    def test_kernel32_exports(self):
        """check we can interpret the export table from kernel32.dll"""    
        kernel32 = ctypes.windll.LoadLibrary('kernel32.dll')
        exp = PEExportDict.from_dll(kernel32)
     
        # some examples of things that should be there
        self.assertIn("Beep", exp)
        self.assertIn("Sleep", exp)

    def test_embed(self):
        """check we can insert and retrieve new ctypes callbacks into a dll"""    

        path = resource_filename(__name__, '_test_pelib.pyd')
        
        dll = ctypes.windll.LoadLibrary(path)    
        thunk = PEExportDict.from_dll(dll, readonly=False)
           
        # make a callback that binds to a local variable
        bound = 0
        def add(added):    
            """callback function registered into an export table"""
            return added + bound
        func = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_int)(add)

        # install into the export table
        thunk['b'] = func
        thunk['c'] = func
        thunk['a'] = func        
               
        self.assertEqual(list(thunk.keys()), [ 'a', 'b', 'c', 'init_test_pelib' ])
        self.assertEqual(len(list(thunk.keys())), len(list(thunk.values())))
        self.assertEqual(len(list(thunk.ordinals())), len(list(thunk.keys())))
       
        # symbol b is already defined in the dll - ordinals should be allocated
        # in the insertion order
        self.assertEqual(list(thunk.ordinals()), [ 3, 1, 2, 0 ])

        # check we got what we expected
        self.assertIn('a', thunk)
        self.assertIn('b', thunk)
        self.assertIn('c', thunk)

        # check that GetProcAddress gives us what we think we shold have
        ptr = ctypes.cast(func, ctypes.c_void_p).value
        self.assertEqual(ctypes.windll.kernel32.GetProcAddress(dll._handle, 'a'), ptr)
        self.assertEqual(ctypes.windll.kernel32.GetProcAddress(dll._handle, 'b'), ptr)
        self.assertEqual(ctypes.windll.kernel32.GetProcAddress(dll._handle, 'c'), ptr)

        # read back the entry point from the dll
        thunk_add = dll['a']
        thunk_add.restype = ctypes.c_int
        thunk_add.argtypes = [ ctypes.c_int ]
        
        # check it all adds up - prove we're calling the same function
        # by changing b
        for bound in range(0, 10):
            for unbound in range(0, 10):
                self.assertEqual(thunk_add(unbound), bound + unbound)

        # now replace the entry b with some other function
        func = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_int)(
                lambda multiplied : multiplied * bound
                )
        thunk['b'] = func    
        
        self.assertEqual(ctypes.windll.kernel32.GetProcAddress(dll._handle, 'b'), ctypes.cast(func, ctypes.c_void_p).value)
        thunk_mul = dll['b']
        for bound in range(0, 10):
            for unbound in range(0, 10):
                self.assertEqual(thunk_mul(unbound), unbound * bound)
        
if __name__ == '__main__':
    unittest.main()

    


