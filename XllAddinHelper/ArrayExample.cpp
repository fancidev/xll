////////////////////////////////////////////////////////////////////////////
// ArrayExample.cpp
//
// This file demonstrates how to take array arguments from and return
// array to Excel. 
//
// The following types are supported as arguments and return value:
//
//   SAFEARRAY *
//
// (scalar, vector, matrix)
//
// (column major, row major)
//
// (fortran array)
//
// In Excel 2003 and earlier, array size is limited to 65,536 rows by
// 256 columns. 

#include "XllAddin.h"
#include <cassert>
#include <comdef.h>

double Trace(SAFEARRAY *mat)
{
	if (mat == nullptr)
		return 0.0;

	assert(SafeArrayGetDim(mat) == 2);

	HRESULT hr;
	VARTYPE vt;
	hr = SafeArrayGetVartype(mat, &vt);
	if (FAILED(hr))
		throw std::invalid_argument("unsupported argument");
	if (vt != VT_VARIANT)
		throw std::invalid_argument("unsupported argument");

	VARIANT *data;
	hr = SafeArrayAccessData(mat, (void**)&data);
	if (FAILED(hr))
		throw std::invalid_argument("Cannot access data");

	if (mat->rgsabound[0].cElements != mat->rgsabound[1].cElements)
	{
		SafeArrayUnaccessData(mat);
		throw std::invalid_argument("Only supports square matrix.");
	}

	ULONG n = mat->rgsabound[0].cElements;
	double sum = 0.0;
	for (ULONG i = 0; i < n; i++)
	{
		VARIANT v;
		VariantInit(&v);
		HRESULT hr = VariantChangeType(&v, data, 0, VT_R8);
		if (FAILED(hr))
			throw std::invalid_argument("Cannot convert to double");
		sum += V_R8(&v);
		data += (n + 1);
	}
	SafeArrayUnaccessData(mat);
	return sum;
}

EXPORT_XLL_FUNCTION(Trace)
.Description(L"Returns the sum of the diagonal elements of a square matrix.");
