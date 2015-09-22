#include "XllAddin.h"

int Plus(int a, int b)
{
	return a + b;
}

int __stdcall Minus(int a, int b)
{
	return a - b;
}

namespace custom_ns
{
	double GetCircleArea(double r)
	{
		return 3.1415926 * r * r;
	}
}

EXPORT_XLL_FUNCTION_AS(custom_ns::GetCircleArea, "GetCircleArea");

namespace
{
	double Divide(double a, double b)
	{
		return a / b;
	}
}

// This example demonstrates how to check input arguments to
// avoid throwing a division-by-zero runtime exception, which
// will crash Excel. A C++ exception, on the other hand, is 
// handled by XLL Connector and returned to Excel as #VALUE!.
int DivInt(int a, int b)
{
	if (b == 0)
		throw std::invalid_argument("b cannot be zero");
	return a / b;
}

EXPORT_XLL_FUNCTION(DivInt);

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

EXPORT_XLL_FUNCTION(Minus, XLL_NOT_VOLATILE | XLL_THREADSAFE)
.Category(L"Test Functions");

EXPORT_XLL_FUNCTION(Square)
.Description(L"Returns the square of a number.")
.Arg(L"x", L"The number to square");

EXPORT_XLL_FUNCTION(Divide);