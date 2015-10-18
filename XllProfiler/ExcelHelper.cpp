////////////////////////////////////////////////////////////////////////////
// ExcelHelper.h -- utility functions for XLL

#include <Windows.h>
#include <strsafe.h>
#include "XLCALL.H"
#include "ExcelHelper.h"
#include "XLString.h"

using namespace xll;

static LPCWSTR s_specialEntryPoints[] =
{
	L"xlAddInManagerInfo",
	L"xlAddInManagerInfo12",
	L"xlAutoAdd",
	L"xlAutoClose",
	L"xlAutoFree",
	L"xlAutoFree12",
	L"xlAutoOpen",
	L"xlAutoRegister",
	L"xlAutoRegister12",
	L"xlAutoRemove",
};

bool IsSpecialEntryPoint(const XLStringW &entryPoint)
{
	for (LPCWSTR specialEntryPoint : s_specialEntryPoints)
	{
		if (entryPoint == specialEntryPoint)
			return true;
	}
	return false;
}

void GetDefinedNames()
{
	XLOPER12 result;
	XLOPER12 arg1, arg2;
	arg1.xltype = xltypeMissing;
	arg2.xltype = xltypeInt;
	arg2.val.w = 2;
	if (Excel12(xlfNames, &result, 2, &arg1, &arg2) == xlretSuccess)
	{
		if (result.xltype == xltypeMulti)
		{
			const XLOPER12 *p = result.val.array.lparray;
			for (int i = 0; i < result.val.array.rows*result.val.array.columns; i++)
			{
				if (p[i].xltype == xltypeStr)
				{
					LPCWSTR lpName = p[i].val.str;
				}
			}
		}
	}
}

std::wstring GetRegisteredName(LPXLOPER12 pxRegisteredId)
{
	std::wstring result;
	XLOPER12 xName;
	if (Excel12(xlfGetDef, &xName, 1, pxRegisteredId) == xlretSuccess)
	{
		if (xName.xltype == xltypeStr)
		{
			result = XLStringW::FromBuffer(xName.val.str);
		}
		Excel12(xlFree, nullptr, 1, &xName);
	}
	return result;
}

FARPROC GetEntryPointAddress(LPCWSTR moduleName, const XLOPER12 &xProcedure)
{
	FARPROC address = NULL;
	if (xProcedure.xltype == xltypeStr)
	{
		char proc[1000];
		int n = WideCharToMultiByte(CP_ACP, 0, &xProcedure.val.str[1],
			xProcedure.val.str[0], proc, sizeof(proc) - 1, nullptr, nullptr);
		if (n > 0)
		{
			proc[n] = '\0';
			HMODULE hModule = GetModuleHandleW(moduleName);
			if (hModule != NULL)
			{
				address = GetProcAddress(hModule, proc);
			}
		}
	}
	else if (xProcedure.xltype == xltypeNum)
	{
		unsigned short ordinal = (unsigned short)xProcedure.val.num;
		HMODULE hModule = GetModuleHandleW(moduleName);
		if (hModule != NULL)
		{
			address = GetProcAddress(hModule, MAKEINTRESOURCEA(ordinal));
		}
	}
	return address;
}

// Returns a list of all registered XLL functions.
void GetRegisteredFunctions(std::vector<RegisteredFunctionInfo> &info)
{
	XLOPER12 result;
	XLOPER12 arg;
	arg.xltype = xltypeInt;
	arg.val.w = 44;
	if (Excel12(xlfGetWorkspace, &result, 1, &arg) == xlretSuccess)
	{
		if (result.xltype == xltypeMulti &&
			result.val.array.lparray != nullptr &&
			result.val.array.columns >= 3)
		{
			const XLOPER12 *p = result.val.array.lparray;
			for (int i = 0; i < result.val.array.rows; i++)
			{
				if (p[0].xltype == xltypeStr &&
					p[1].xltype == xltypeStr &&
					p[2].xltype == xltypeStr)
				{
					auto& lpDllName = XLStringW::FromBuffer(p[0].val.str);
					auto& lpFuncName = XLStringW::FromBuffer(p[1].val.str);
					auto& lpTypeText = XLStringW::FromBuffer(p[2].val.str);

					// Exclude special entry points.
					if (!IsSpecialEntryPoint(lpFuncName))
					{
						// Get register id of the function.
						XLOPER12 xProcedure;
						if (lpFuncName.length() > 0 && lpFuncName.cbegin()[0] == '#')
						{
							std::wstring t = lpFuncName;
							xProcedure.xltype = xltypeNum;
							xProcedure.val.num = wcstod(&t[1], nullptr);
						}
						else
						{
							xProcedure.xltype = xltypeStr;
							xProcedure.val.str = p[1].val.str;
						}

						XLOPER12 xId;
						if (Excel12(xlfRegisterId, &xId, 3, &p[0], &xProcedure, &p[2]) == xlretSuccess)
						{
							if (xId.xltype == xltypeNum)
							{
								RegisteredFunctionInfo entry;
								entry.id = xId.val.num;
								entry.functionName = GetRegisteredName(&xId);
								entry.dllName = lpDllName;
								entry.entryPointName = lpFuncName;
								entry.typeText = lpTypeText;

								FARPROC entryPointAddress = GetEntryPointAddress(
									entry.dllName.c_str(), xProcedure);
								if (entryPointAddress != nullptr)
								{
									entry.entryPointAddress = entryPointAddress;
									info.push_back(entry);
								}
							}
							Excel12(xlFree, nullptr, 1, &xId);
						}
					}
				}
				p += result.val.array.columns;
			}
		}
		Excel12(xlFree, nullptr, 1, &result);
	}
}