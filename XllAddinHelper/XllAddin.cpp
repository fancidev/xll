#include <Windows.h>
#include "XllAddin.h"

#define EXPORT_UNDECORATED_NAME comment(linker, "/export:" __FUNCTION__ "=" __FUNCDNAME__)

int WINAPI xlAutoOpen()
{
#pragma EXPORT_UNDECORATED_NAME

#if 0
	static XLOPER12 xDLL;
	Excel12(xlGetName, &xDLL, 0);
	//MessageBoxW(NULL, L"xlAutoOpen", L"MyAddin", MB_OK);
	Excel12(xlfRegister, 0, 4,
		&xDLL,
		&ExcelVariant(L"CalcCircum"),
		&ExcelVariant(L"BB"),
		&ExcelVariant(L"CalcCircum"));

	Excel12(xlfRegister, 0, 4,
		&xDLL,
		&ExcelVariant(L"XLSquare"),
		&ExcelVariant(L"BB"),
		&ExcelVariant(L"Mysquare"));

	Excel12(xlFree, 0, 1, &xDLL);
#else
	AddinRegistrar::RegisterAllFunctions();
#endif

	return 1;
}

int WINAPI xlAutoClose()
{
#pragma EXPORT_UNDECORATED_NAME

	return 1;
}

LPXLOPER12 WINAPI xlAutoRegister12(LPXLOPER12 pxName)
{
	return 0;
}

int WINAPI xlAutoAdd(void)
{
	return 1;
}

int WINAPI xlAutoRemove(void)
{
	return 1;
}

LPXLOPER12 WINAPI xlAddInManagerInfo12(LPXLOPER12 xAction)
{
	return 0;
}
