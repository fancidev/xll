#include "XllAddin.h"

int Plus(int a, int b)
{
	return a + b;
}

int __stdcall Minus(int a, int b)
{
	return a - b;
}

double GetCircleArea(double r)
{
	return 3.1415926 * r * r;
}

double __fastcall Square(double x)
{
	return x * x;
}

// Test using EXPORT_XLL_FUNCTION in a namespace.
namespace
{
	EXPORT_XLL_FUNCTION(Plus)
		.Description(L"Returns the sum of two numbers.")
		.Arg(L"a", L"first number")
		.Arg(L"b", L"second number");
}

EXPORT_XLL_FUNCTION(Minus)
.Category(L"Test Functions")
.Pure()
.ThreadSafe();

EXPORT_XLL_FUNCTION(Square)
.Description(L"Returns the square of a number.")
.Arg(L"x", L"The number to square");
