#include <Windows.h>
#include "XllAddin.h"

#define EXPORT_UNDECORATED_NAME comment(linker, "/export:" __FUNCTION__ "=" __FUNCDNAME__)

const ExcelVariant ExcelVariant::Empty(ExcelVariant::FromType(xltypeNil));
const ExcelVariant ExcelVariant::Missing(ExcelVariant::FromType(xltypeMissing));

double __stdcall MyTestFunc(double x, double y)
{
#pragma EXPORT_UNDECORATED_NAME
	return x * y;
}

const wchar_t * __stdcall MyToString(double x)
{
#pragma EXPORT_UNDECORATED_NAME
	static wchar_t result[100];
	swprintf(result, 100, L"%lf", x);
	return result;
}

static int RegisterFunction(LPXLOPER12 dllName, const FunctionInfo &f)
{
	if (f.arguments.size() > 245)
		throw std::invalid_argument("Too many arguments");

	std::wstring argumentText;
	if (f.arguments.size() > 0)
	{
		argumentText = f.arguments[0].name();
		for (size_t i = 1; i < f.arguments.size(); i++)
		{
			argumentText += L",";
			argumentText += f.arguments[i].name();
		}
	}

	ExcelVariant opers[256];
	// opers[0] = dllName;
	opers[1] = f.entryPoint;
	opers[2] = f.typeText;
	opers[3] = f.name;
	opers[4] = argumentText;
	opers[5] = f.macroType;
	opers[6] = f.category;
	opers[7] = f.shortcut;
	opers[8] = f.helpTopic;
	opers[9] = f.description;
	for (size_t i = 0; i < f.arguments.size(); i++)
	{
		// Excel sometimes truncates the last one or two characters of the
		// last argument description. Therefore we need to append two spaces
		// to the last argument description to counter this behavior. See 
		// https://msdn.microsoft.com/en-us/library/office/bb687841.aspx
		if (i == f.arguments.size() - 1 && f.arguments[i].description() != nullptr)
			opers[10 + i] = std::wstring(f.arguments[i].description()) + L"  ";
		else
			opers[10 + i] = f.arguments[i].description();
	}

	LPXLOPER12 popers[256];
	popers[0] = dllName;
	for (size_t i = 1; i < 10u + f.arguments.size(); i++)
		popers[i] = &opers[i];

	// If opers[9] is supplied, regardless of its value, Excel will not
	// automatically fill in argument text. So we do not supply it unless
	// user has specified something.
	int n;
	if (f.description == nullptr && f.arguments.size() == 0)
		n = 9;
	else
		n = 10 + f.arguments.size();

	XLOPER12 id;
	int ret = Excel12v(xlfRegister, &id, n, popers);
	return ret;
}

static void RegisterAllFunctions()
{
	XLOPER12 xDLL;
	Excel12(xlGetName, &xDLL, 0); // TODO: check return value

#if 0
	Excel12(xlfRegister, 0, 6,
		&xDLL,
		&ExcelVariant(L"MyToString"),
		&ExcelVariant(L"C%B"),
		&ExcelVariant(L"MyToString"),
		&ExcelVariant(L"x"),
		&ExcelVariant(1.0));
#endif

#if 0
	Excel12(xlfRegister, 0, 10,
		&xDLL,
		&ExcelVariant(L"MyTestFunc"),
		&ExcelVariant(L"BBB"),
		&ExcelVariant(L"MyTestFunc"),
		&ExcelVariant(L"a,b,c"), // 4: argumentText; extra arguments are shown but if you fill them, you get error
		&ExcelVariant(1.0),
		&ExcelVariant(L""),
		&ExcelVariant(L""), 
		&ExcelVariant(L""), // 8
		//&ExcelVariant((wchar_t*)nullptr)  // 9
		&ExcelVariant::Missing // 9 -- if this argument is supplied, regardless of
								// what value it is, Excel will not automatically
								// fill in argument text. You must supply it in
								// argumentText.
		);
#endif

	for (FunctionInfo& f : AddinRegistrar::registry())
	{
		RegisterFunction(&xDLL, f);
	}

	Excel12(xlFree, 0, 1, &xDLL);
}

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
	RegisterAllFunctions();
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
