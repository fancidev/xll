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

EXPORT_XLL_FUNCTION(Plus)
.Description(L"Returns the sum of two numbers.")
.Arg(L"a", L"first number")
.Arg(L"b", L"second number");

EXPORT_XLL_FUNCTION(Minus);
EXPORT_XLL_FUNCTION(Square)
.Description(L"Returns the square of a number.")
.Arg(L"x", L"The number to square");

EXPORT_XLL_FUNCTION(ReverseString);

BOOL WINAPI DllMain(HANDLE hInstance, ULONG fdwReason, LPVOID lpReserved)
{
	return TRUE;
}
