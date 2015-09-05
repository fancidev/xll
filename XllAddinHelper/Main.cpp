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



DWORD SlowFunc()
{
	Sleep(1000);
	return GetCurrentThreadId();
}

EXPORT_XLL_FUNCTION(GetCurrentThreadId)
.Volatile()
.ThreadSafe()
.Description(L"Returns the id of the thread that is evaluating this function.")
.HelpTopic(L"https://msdn.microsoft.com/en-us/library/windows/desktop/ms683183(v=vs.85).aspx!0");

EXPORT_XLL_FUNCTION(SlowFunc)
.Volatile()
.ThreadSafe();

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

BOOL WINAPI DllMain(HANDLE hInstance, ULONG fdwReason, LPVOID lpReserved)
{
	return TRUE;
}
