#define PY_SSIZE_T_CLEAN
#include <Python.h>

#include <windows.h>
#include <tchar.h>
#include <stdio.h>
#include <io.h>
#include <fcntl.h>

// problems:
//
// currently the addin lives in the scripts/ folder, which allows us to figure out that the entry
// point exists whether in production or development as it is always copied.

// however: wheel files do not let us deploy arbitrary stuff into the scripts folder. it only supports entry_points
//
// options are: relocate the xll addin into the python packages folder.
//				when fully installed, go up the tree to find the path with the lib folder, and initialize th path from there
//     			in development - we may come from a source folder, which is no good, as ther eis not lib.
//              do both...?
//
// hack the wheel code to support better entry point deployment?
//
// how to develop: allow some environment variable
//
// use a single xll which hooks into a particular module imported from it's virtual env.
// have to break the -script.py
//

//
// how to debug...

// TODO move this logic to a python library, now it is more stable:
// TODO we may wish to optionally create a new console even if we can AttachConsole...
// TODO does this work for inherited file handles form pyxcel commands?
// TODO make exceltools separate to the XLL stuff - we cant install entrypoints without it...
// TODO can we speciy dll entry points as ctypes functions in modules using setuptools entry points...

// the dll containing this code lives alongside the python27dll for the
// virtualenv, so out init must set that up.
// all virtaulenv init should go from there...

// if a buffer wasn't initialized because the inherited handle
// was a console, yet the application was not then reopen it
// and make sure it is unbuffered.
static void _reopen_console(const char *name, FILE *fp, const char *mode)
{
	if (_fileno(fp) == -2)
	{
		freopen(name, mode, fp);
		setvbuf(fp, NULL, _IONBF, 0);
	}
}

static void _console()
{
	// standard handles that are not redirected to console will already
	// have been initialized correctly - mscrt obeys pipe/file redirections
	// even if we are not in a console process.

	// however console handles are not inherited, so first we attach a console
	if (GetConsoleWindow() == NULL)
		AttachConsole(ATTACH_PARENT_PROCESS);

	// if we have no console, and no redirection, then allocate a new console
	if (GetConsoleWindow() == NULL)
		if (!AllocConsole())
			return;

	// -2 means we had no console window at startup, so fixup the streams
	// that aren't propertly initialized
	_reopen_console("CONIN$", stdin, "rb");
	_reopen_console("CONOUT$", stdout, "ab");
	_reopen_console("CONOUT$", stderr, "ab");

	// some packacges (pytest) assume that file descriptors 0, 1, 2
	// are always stdin/out/err, so dup them back so it works
	if (_fileno(stdin) != 0)
		_dup2(_fileno(stdin), 0);
	if (_fileno(stdout) != 1)
		_dup2(_fileno(stdout), 1);
	if (_fileno(stderr) != 2)
		_dup2(_fileno(stderr), 2);
}

static PyObject *_initmodule(const char *name, HINSTANCE hInstDll)
{
	PyObject *m, *d, *h;
	m = PyImport_AddModule(name);
	;
	if (m == NULL)
		Py_FatalError("can't create __main__ module");
	d = PyModule_GetDict(m);
	if (PyDict_GetItemString(d, "__builtins__") == NULL)
	{
		PyObject *bimod = PyImport_ImportModule("__builtin__");
		if (bimod == NULL ||
			PyDict_SetItemString(d, "__builtins__", bimod) != 0)

			Py_FatalError("can't add __builtins__ to __main__");
		Py_XDECREF(bimod);
	}

	h = PyLong_FromSize_t((size_t)hInstDll);
	if (h == NULL)
		return NULL;
	if (PyDict_SetItemString(d, "__hInstDll__", h) != 0)
	{
		Py_XDECREF(h);
		return NULL;
	}
	Py_XDECREF(h);

	return m;
}

BOOL Run_DllMain(HINSTANCE hInstDll, DWORD fdwReason, LPVOID lpvReserved)
{
	static PyObject *m = 0;
	PyObject *d, *r, *f;
	FILE *fp;

	TCHAR filename[MAX_PATH];
	TCHAR *pos;

	// argc is just the script name
	char *argc = filename;

	// initialize the module on attach
	if (fdwReason != DLL_PROCESS_ATTACH)
		return TRUE;

	// figure out the name of this module
	if (GetModuleFileName(hInstDll, filename, MAX_PATH) == 0)
		return FALSE;

	// remove the trailing extension to get the module to import
	pos = _tcsrchr(filename, '.');
	if (pos == NULL)
		return FALSE;
	*pos = 0;

	// add on the script suffix
	_tcscat(filename, "-script");

	// find the trailing '\', this gives us the module name
	// dll names must be unique within their processes
	pos = _tcsrchr(filename, '\\');
	if (pos == NULL)
		return FALSE;
	++pos;

	// construct the module with members setup - use the full file name
	// of the shared library
	m = _initmodule(pos, hInstDll);
	if (m == 0)
	{
		PyErr_Print();
		return FALSE;
	}
	// get the dictionary to add more members
	d = PyModule_GetDict(m);

	_tcscat(filename, ".py");

	// setup the argv as being just the script name, no arguments
	PySys_SetArgvEx(1, &argc, 0);

	// setup the __file__ member
	f = PyUnicode_FromString(filename);
	if (f == NULL)
	{
		PyErr_Print();
		return FALSE;
	}
	if (PyDict_SetItemString(d, "__file__", f) < 0)
	{
		Py_DECREF(f);
		PyErr_Print();
		return FALSE;
	}

	fp = fopen(filename, "r");
	if (fp == NULL)
	{
		PyErr_Format(PyExc_RuntimeError, "failed to load %s", filename);
		PyErr_Print();
		return FALSE;
	}

	// run the file in our new module
	r = PyRun_FileExFlags(
		fp,
		filename,
		Py_file_input,
		d, d,
		1, 0);

	if (r == NULL)
	{
		PyErr_Print();
		return FALSE;
	}
	Py_DECREF(r);

	if (PyObject_HasAttrString(m, "DllMain"))
	{
		f = PyObject_GetAttrString(m, "DllMain");
		if (!f)
		{
			PyErr_Print();
			return FALSE;
		}

		// invoke DllMain with the same arguments.
		r = PyObject_CallFunction(f, "iii", hInstDll, fdwReason, lpvReserved);
		Py_DECREF(f);

		if (!r)
		{
			PyErr_Print();
			return FALSE;
		}
		else
		{
			int result = PyObject_IsTrue(r);
			if (result == -1)
			{
				PyErr_Print();
				result = FALSE;
			}

			Py_DECREF(r);
			return result;
		}
	}

	return TRUE;
}

// TODO DllMain initializes the interpreter - should this be in the xlAutoOpen/xlAutoClose
//      and somehow isolated?

BOOL WINAPI DllMain(HINSTANCE hInstDll, DWORD fdwReason, LPVOID lpvReserved)
{
	TCHAR filename[MAX_PATH];
	TCHAR home[MAX_PATH];
	TCHAR *prefix;
	TCHAR *pos;

	PyGILState_STATE state;
	BOOL result;

	DisableThreadLibraryCalls(hInstDll);

	// somehow python is shut down before we get the detach
	// so calls to GIL etc will just hang. ignore detachment
	// in liu of just letting python do it's normal cleanup
	if (fdwReason == DLL_PROCESS_DETACH || fdwReason == DLL_THREAD_DETACH)
		return TRUE;

	// initialize the module on attach
	if (fdwReason == DLL_PROCESS_ATTACH)
	{
		// fix console output first, so we see stdout/err early on.
		_console();

		// figure out the name of this module
		if (GetModuleFileName(hInstDll, filename, MAX_PATH) == 0)
			return FALSE;

		// figure out the name of this module
		if (GetModuleFileName(hInstDll, home, MAX_PATH) == 0)
			return FALSE;

		pos = _tcsrchr(home, '\\');
		*pos = 0;
		pos = _tcsrchr(home, '\\');
		*pos = 0;

		// TODO remember if we are initialized and de-initialize if unloaded?
		if (!Py_IsInitialized())
		{
			Py_SetProgramName(filename);
			Py_SetPythonHome(home);

			PyEval_InitThreads();
			Py_InitializeEx(1);
			PyEval_SaveThread();
		}

		// now we lock GIL again via the state API and load our module
		// state = PyGILState_Ensure();
		// result = Run_DllMain(hInstDll, fdwReason, lpvReserved);
		// PyGILState_Release(state);

		return TRUE;
	}

	return TRUE;
}

static PyMethodDef _addin_methods[] = {
	{NULL, NULL, 0, NULL} /* Sentinel */
};

static struct PyModuleDef _addin_module_def = {
	PyModuleDef_HEAD_INIT,
	"spam", /* name of module */
	NULL,	/* module documentation, may be NULL */
	-1,		/* size of per-interpreter state of the module,
			   or -1 if the module keeps state in global variables. */
	_addin_methods};

PyMODINIT_FUNC
PyInit__addin(void)
{
	// PyErr_SetString(PyExc_RuntimeError, "exceltools._addin should never be imported by python");
	return PyModule_Create(&_addin_module_def);
}

#pragma comment(linker, "/export:xlAutoOpen@0=xlAutoOpen")
int __stdcall xlAutoOpen() { return 0; }
