"""interrogate and manipulate the a PE format image in memory"""
import ctypes

from ctypes import byref, sizeof, cast, c_char_p, c_void_p, create_string_buffer, cdll, c_int, addressof
from ctypes.wintypes import DWORD, WORD, LPCSTR

from pelib.peformat import VirtualProtect, PAGE_READWRITE 
from pelib.peformat import IMAGE_EXPORT_DIRECTORY, IMAGE_DOS_HEADER, IMAGE_NT_HEADERS32
from pelib.peformat import lstrcmpA
import sys

_kernel32 = ctypes.windll.kernel32

class PEExportDict(IMAGE_EXPORT_DIRECTORY): 
    @classmethod
    def from_dll(cls, dll, readonly=True):
        return PEExportDict.from_handle(dll._handle, readonly=readonly)

    def _rva(self, addr):
        return ctypes.c_long(addr).value + self.base
    def _torva(self, addr):
        return addr - self.base

    @classmethod
    def from_handle(cls, handle, readonly=True):
        # Windows DLL handles are their memory addresses

        # traverse the PE header structures to get to the export table
        dos = IMAGE_DOS_HEADER.from_address(handle)
        assert dos.e_magic == 0x5A4D
        nt = IMAGE_NT_HEADERS32.from_address(dos.e_lfanew + handle)
        assert nt.Signature == 0x4550
        opt = nt.OptionalHeader
        assert opt.Magic == 0x010B

        # check that dll actually has an export table for us to tweak
        # if not present, can we add a new one? 
        if opt.DataDirectory[0].VirtualAddress == 0 \
        or opt.DataDirectory[0].Size == 0:
            raise RuntimeError("Could not find export table")

        # initialize an instance around the location
        self = PEExportDict.from_address(opt.DataDirectory[0].VirtualAddress + handle)
        self.base = handle
        
        # build wrapper types around the existing export structure
        self.Names = (DWORD * self.NumberOfNames).from_address(self._rva(self.AddressOfNames))
        self.NameOrdinals = (WORD * self.NumberOfNames).from_address(self._rva(self.AddressOfNameOrdinals))
        self.Functions = (DWORD * self.NumberOfFunctions).from_address(self._rva(self.AddressOfFunctions))
        
        # place to cache buffers allocated for C strings? or should we just use malloc
        self._bufs = dict()
        self._vals = dict()

        # if we are not read only then get access to the pages resize to allocate
        # our own memory
        if not readonly:
            if not VirtualProtect(byref(self), sizeof(self), PAGE_READWRITE, byref(DWORD())):
                raise RuntimeError("Could not make export table pages writable")
                
        return self

    def keys(self):
        Names = (DWORD * self.NumberOfNames).from_address(self._rva(self.AddressOfNames))
        for name in Names:
            yield cast(name + self.base, c_char_p).value

    def values(self):
        # return functions by looking at the ordinals first
        NameOrdinals = (WORD *self.NumberOfNames).from_address(self._rva(self.AddressOfNameOrdinals))
        Functions = (DWORD * self.NumberOfFunctions).from_address(self._rva(self.AddressOfFunctions))
        for ordinal in NameOrdinals:
            assert ordinal < self.NumberOfFunctions
            yield cast(Functions[ordinal] + self.base, c_void_p)

    def ordinals(self):
        NameOrdinals = (WORD  * self.NumberOfNames).from_address(self._rva(self.AddressOfNameOrdinals))
        for ordinal in NameOrdinals:
            yield ordinal

    def exports(self):
        Names = (DWORD * self.NumberOfNames).from_address(self._rva(self.AddressOfNames))
        NameOrdinals = (WORD * self.NumberOfNames).from_address(self._rva(self.AddressOfNameOrdinals))
        Functions = (DWORD * self.NumberOfFunctions).from_address(self._rva(self.AddressOfFunctions))
        for name, ordinal in zip(Names, NameOrdinals):
            yield (cast(name + self.base, c_char_p).value, ordinal, self._rva(Functions[ordinal]))

    def __iter__(self):
        return self.keys()

    def __contains__(self, key):                
        return key in self.keys()
        
    def __setitem__(self, key, value):
        # where to keep these? should be a __dict__?
        buf = create_string_buffer(key)

        # get the c strcmp function to use in the comparisons below so we don't spend
        # a lot of time converting to/from python strings
        strcmp = cdll.msvcrt.strcmp
        strcmp.restype = c_int
        strcmp.argtypes = [ c_int, c_int ]

        # our own definition of bisect which compares the strings directly with strcmp
        def bisect_left(a, x, lo=0, hi=None):
            if lo < 0:
                raise ValueError('lo must be non-negative')
            if hi is None:
                hi = len(a)
            while lo < hi:
                mid = (lo+hi)//2
                cmp = strcmp(a[mid] + self.base, addressof(x))
                if cmp < 0: lo = mid+1
                else: hi = mid
            return lo
        
        i = bisect_left((DWORD * self.NumberOfNames).from_address(self._rva(self.AddressOfNames)), buf)

        ptr = cast(value, c_void_p).value

        # hold a reference to the value
        self._vals[key] = value

        # if we find an exact match, replace the existing engtry
        if i < self.NumberOfNames and lstrcmpA(
                cast(self._rva(self.Names[i]), LPCSTR), cast(addressof(buf), LPCSTR)) == 0:       
            # overwrite the existing entry, preserve the ordinal
            ordinal = self.NameOrdinals[i]
            self.Functions[ordinal] = self._torva(ptr)  # first make space in all the arrays
        else:            
            # else expand the arrays to make space, and add a new entry
            factor = 2

            if len(self.Functions) == self.NumberOfFunctions:
                self.Functions = (DWORD * (self.NumberOfFunctions * factor))(*(self.Functions))
                self.AddressOfFunctions = self._torva(addressof(self.Functions))

            if len(self.Names) == self.NumberOfNames:
                self.Names = (DWORD * (self.NumberOfNames* 2))(*(self.Names))
                self.AddressOfNames = self._torva(addressof(self.Names))

            if len(self.NameOrdinals) == self.NumberOfNames:
                self.NameOrdinals = (WORD * (self.NumberOfNames * factor))(*(self.NameOrdinals))
                self.AddressOfNameOrdinals = self._torva(addressof(self.NameOrdinals))

            # the new function pointer simply goes on the end
            ordinal = self.NumberOfFunctions
            self.Functions[self.NumberOfFunctions] = self._torva(ptr)
            self.NumberOfFunctions += 1

            # TODO use malloc here and forget about the memory else we crash if the thunk is lost
            # TODO or perhaps put some module wrapper around and store in sys.modules?
            
            # insert the new function name and it's ordinal at the correct location
            self.Names[i+1:len(self.Names)] = self.Names[i:len(self.Names)-1]
            self.NameOrdinals[i+1:len(self.NameOrdinals)] = self.NameOrdinals[i:len(self.NameOrdinals)-1]

            # insert the new value into the array
            self.Names[i] = addressof(buf) - self.base
            self.NameOrdinals[i] = ordinal
            self.NumberOfNames += 1 
        
            # keep a reference to the string
            self._bufs[key] = buf

        # since this is quite hairy, double check that GetProcAddress works just to be safe
        assert ptr == _kernel32.GetProcAddress(self.base, key)

    def print_exports(self):
         # do we think it's there? 
        print "%32s % 4s %10s" % ("Name", "Ord", "Address") 
        print "%32s % 4s %10s" % ('-' * 32, '-' * 4, '-'*10) 
        for export in self.exports():
            print "%32s % 4d 0x%08X" % export
        print '\n'
      














