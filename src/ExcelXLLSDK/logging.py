from __future__ import absolute_import

from logging.handlers import MemoryHandler

from ExcelXLLSDK.XLCALL import xlEventRegister, xlfCaller

# note that ON.DOUBLECLICK and friends may be useful as events on given sheeet etc.

# need some thread local flag that logging should go the event?

def LogContext(log):
    class _BoundLogContext(object):
        def __init__(self, reg, args):
             self.msg = "={}({})".format(reg.FunctionText, ', '.join(map(repr, args)))

        def __enter__(self):
             log.debug(self.msg)

        def __exit__(self, *exc_info):
            if exc_info[0]:
                log.exception(repr(exc_info[0]), exc_info=exc_info)
            else:
                log.debug(log.info(self.msg))

    return _BoundLogContext

# named range handler needs
#
# # xlfCaller needs to wrap,
#
# class DeferContext(object):
#     def __enter__(self):
#         pass
#
#     def __exit__(self, *exc_info):
#         pass
#
# _deferred = []
#
# # need an isUDF? if xlfCaller gives xltypeSRef/Ref, then we are in-cell, and can't log (yet)
#
# class NamedRangeHandler(MemoryHandler):
#     def __init__(self, range):
#         self.range = range
#         self.iter = self.range.iterrows()
#
#     def emit(self, record):
#         data = XLOPER([[ record.created, record.levelname, record.name, record.getMessage() ]])
#         row = self.iter.next()
#
#         if xlfCaller.is_ref():
#             # need to put it in a queue?
#         else:
#             xlSet(row, data)
