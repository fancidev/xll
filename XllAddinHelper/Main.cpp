#include "XllAddin.h"

int Plus(int a, int b)
{
	return a + b;
}

int Minus(int a, int b)
{
	return a - b;
}

double GetCircleArea(double r)
{
	return 3.1415926 * r * r;
}

double Square(double x)
{
	return x * x;
}

std::wstring ReverseString(const std::wstring &s)
{
	return std::wstring(s.crbegin(), s.crend());
}

double Trace(const VARIANT &v)
{
	if ((v.vt & VT_ARRAY) == 0)
		throw std::invalid_argument("argument must be a matrix");

	SAFEARRAY *mat = v.parray;
	if (SafeArrayGetDim(mat) != 2)
		throw std::invalid_argument("argument must be a matrix");

	HRESULT hr;
	VARTYPE vt;
	hr = SafeArrayGetVartype(mat, &vt);
	if (FAILED(hr))
		throw std::invalid_argument("unsupported argument");
	if (vt != VT_R8)
		throw std::invalid_argument("unsupported argument");

	double *data;
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
		sum += *data;
		data += (n + 1);
	}
	SafeArrayUnaccessData(mat);
	return sum;
}

EXPORT_XLL_FUNCTION(Plus)
.Description(L"Returns the sum of two numbers.")
.Arg(L"a", L"first number")
.Arg(L"b", L"second number");

EXPORT_XLL_FUNCTION(Minus)
.Category(XS("Test Functions"))
.Pure()
.ThreadSafe();

EXPORT_XLL_FUNCTION(Square)
.Description(L"Returns the square of a number.")
.Arg(L"x", L"The number to square");

EXPORT_XLL_FUNCTION(ReverseString);

//EXPORT_XLL_FUNCTION(Trace)
//.Description(L"Returns the sum of the diagonal elements of a square matrix.");

BOOL WINAPI DllMain(HANDLE hInstance, ULONG fdwReason, LPVOID lpReserved)
{
	return TRUE;
}
