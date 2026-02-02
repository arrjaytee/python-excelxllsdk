from multimethod import multimethod
from .xltypes import from_value, _XLOPER

from .gen.xlerr import *
from .gen.xltype import xltypeErr

@multimethod(_XLOPER, Exception)
def from_value(self, value):
    self.val.err = xlerrNull
    self.xltype = xltypeErr

@multimethod(_XLOPER, ZeroDivisionError)
def from_value(self, value):
    self.val.err = xlerrDiv0
    self.xltype = xltypeErr

@multimethod(_XLOPER, NameError)
def from_value(self, value):
    self.val.err = xlerrName
    self.xltype = xltypeErr

@multimethod(_XLOPER, SyntaxError)
def from_value(self, value):
    self.val.err = xlerrName
    self.xltype = xltypeErr

@multimethod(_XLOPER, AssertionError)
def from_value(self, value):
    self.val.err = xlerrValue
    self.xltype = xltypeErr

@multimethod(_XLOPER, TypeError)
@multimethod(_XLOPER, ValueError)
def from_value(self, value):
    self.val.err = xlerrValue
    self.xltype = xltypeErr

@multimethod(_XLOPER, StandardError)
def from_value(self, value):
    self.val.err = xlerrNum
    self.xltype = xltypeErr

