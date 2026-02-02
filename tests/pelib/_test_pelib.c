#include "Python.h"

PyMODINIT_FUNC
init_test_pelib()
{
	PyErr_SetString(PyExc_RuntimeError, "exceltools._addin should never be imported by python");
	return;
}