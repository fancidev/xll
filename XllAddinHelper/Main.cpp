//#include "AutoWrap.h"
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

/// <summary>Returns the square of a number.</summary>
/// <param name="x">The value to square.</param>
/// <returns>The square of x.</returns>
/// <example>
/// <c>=Square(1.2)</c> returns 1.44.
/// </example>
double Square(double x)
{
	return x * x;
}

EXPORT_XLL_FUNCTION(Plus)
.Description(L"Returns the sum of two numbers.")
.Arg(L"a", L"first number")
.Arg(L"b", L"second number");

EXPORT_XLL_FUNCTION(Minus);
EXPORT_XLL_FUNCTION(Square);

//static auto &x = NAMED_FUNCTION(Plus)
//.Name("Plus")
//.Description("Returns the sum of two numbers.")
//.Arg("a", "First number")
//.Arg("b", "Second number");

#if 0
class MyTempClass
{
	int x;
};

/// <summary>Creates an instance of a class.</summary>
/// <param name="p2">Parameter 2</param>
/// <param name="far">Parameter 2</param>
/// <param name="haha">Parameter 2</param>
/// <returns>The object.</returns>
MyTempClass CreateInstance(int p1, int p2, const MyTempClass &from, VARIANTARG v)
{
	return MyTempClass();
}
#endif

BOOL WINAPI DllMain(HANDLE hInstance, ULONG fdwReason, LPVOID lpReserved)
{
	return TRUE;
}
