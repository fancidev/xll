#include "XllAddin.h"

#if _DEBUG
XLL_ADDIN_NAME(L"XLL Connector Examples (Debug)");
#else
XLL_ADDIN_NAME(L"XLL Connector Examples (Release)");
#endif

BOOL WINAPI DllMain(HANDLE hInstance, ULONG fdwReason, LPVOID lpReserved)
{
	return TRUE;
}
