from ctypes import *



xlretNotThreadSafe = 128 # Variable c_int '128'
xlretInvXlfn = 2 # Variable c_int '2'
xlretStackOvfl = 16 # Variable c_int '16'
xlretNotClusterSafe = 512 # Variable c_int '512'
xlretInvXloper = 8 # Variable c_int '8'
xlretInvAsynchronousContext = 256 # Variable c_int '256'
xlretUncalced = 64 # Variable c_int '64'
xlretInvCount = 4 # Variable c_int '4'
xlretFailed = 32 # Variable c_int '32'
xlretAbort = 1 # Variable c_int '1'
xlretSuccess = 0 # Variable c_int '0'
__all__ = ['xlretStackOvfl', 'xlretNotClusterSafe',
           'xlretInvAsynchronousContext', 'xlretAbort',
           'xlretNotThreadSafe', 'xlretInvXlfn', 'xlretUncalced',
           'xlretInvCount', 'xlretFailed', 'xlretInvXloper',
           'xlretSuccess']
