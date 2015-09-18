////////////////////////////////////////////////////////////////////////////
// Addin.h -- book keeping of exported functions and manages registration

#include "xlldef.h"
#include "FunctionInfo.h"
#include "ExcelVariant.h"
#include "Conversion.h"
#include <vector>
#include <cassert>

namespace XLL_NAMESPACE
{
#if XLL_SUPPORT_THREAD_LOCAL
	__declspec(thread) XLOPER12 xllReturnValue;
#endif
}

using namespace XLL_NAMESPACE;

static HMODULE GetThisModuleHandle()
{
	HMODULE hThisDll;
	if (::GetModuleHandleEx(
		 GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS
		|GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
		(LPCTSTR)(void*)(&GetThisModuleHandle), &hThisDll))
	{
		return hThisDll;
	}
	return NULL;
}

class ExportTableHelper
{
	struct ExportEntry
	{
		FARPROC proc;
		WORD ordinal;
		LPCSTR name;
	};

	std::vector<ExportEntry> m_exportEntries; // todo: handle exceptions

public:
	ExportTableHelper(HMODULE hModule)
	{
		BYTE* pImageBase = (BYTE*)hModule;
		PIMAGE_DOS_HEADER pDosHeader = (PIMAGE_DOS_HEADER)hModule;
		if (pDosHeader->e_magic != IMAGE_DOS_SIGNATURE)
			return;

		PIMAGE_NT_HEADERS pNtHeaders = (PIMAGE_NT_HEADERS)&pImageBase[pDosHeader->e_lfanew];
		if (pNtHeaders->Signature != IMAGE_NT_SIGNATURE)
			return;
		if (pNtHeaders->OptionalHeader.NumberOfRvaAndSizes == 0)
			return;

		PIMAGE_EXPORT_DIRECTORY pExportDirectory = (PIMAGE_EXPORT_DIRECTORY)
			&pImageBase[pNtHeaders->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT].VirtualAddress];

		DWORD *pFunctionRVAs = (DWORD*)&pImageBase[pExportDirectory->AddressOfFunctions];
		for (DWORD i = 0; i < pExportDirectory->NumberOfFunctions; ++i)
		{
			if (pFunctionRVAs[i] != 0)
			{
				ExportEntry ent;
				ent.name = nullptr;
				ent.proc = (FARPROC)&pImageBase[pFunctionRVAs[i]];
				ent.ordinal = pExportDirectory->Base + i;
				m_exportEntries.push_back(ent);
			}
		}
		
		DWORD* pNameRVAs = (DWORD*)&pImageBase[pExportDirectory->AddressOfNames];
		for (DWORD i = 0; i < pExportDirectory->NumberOfNames; i++)
		{
			LPCSTR name = (LPCSTR)&pImageBase[pNameRVAs[i]];
		}

		// Sort the entries by proc address to make it easier to search.
	}

	WORD GetProcOrdinal(FARPROC proc)
	{
		for (size_t i = 0; i < m_exportEntries.size(); ++i)
		{
			if (m_exportEntries[i].proc == proc)
				return m_exportEntries[i].ordinal;
		}
		return 0;
	}
};

static ExportTableHelper exportTable(GetThisModuleHandle());

static int RegisterFunction(LPXLOPER12 dllName, const FunctionInfo &f)
{
	// This is enforced by a static_assert in XLWrapper.
	assert(f.arguments.size() <= XLL_MAX_ARG_COUNT);

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

	// Find ordinal of entry point. We may support export by name
	// in the future.
	WORD ordinal = exportTable.GetProcOrdinal(f.entryPoint);
	if (ordinal == 0)
		return xlretFailed;
	opers[1] = ordinal;

	opers[2] = std::wstring(f.typeText) + (f.isPure ? L"" : L"!") + (f.isThreadSafe ? L"$" : L"");
	opers[3] = f.name;
	// BUG: if the function description is given, then even if the UDF takes
	//      no arguments, Excel still shows a box to let the user input the
	//      argument. Need to find a way to get rid of the box.
	if (!argumentText.empty())
		opers[4] = argumentText;
	else
		opers[4] = (wchar_t*)nullptr;
	opers[5] = f.macroType;
	opers[6] = f.category;
	opers[7] = f.shortcut;
	opers[8] = f.helpTopic;
	//opers[8] = L"e:\\Dev\\Repos\\Xll\\Test\\A.chm!123";
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
		n = 10 + static_cast<int>(f.arguments.size());

	XLOPER12 id;
	int ret = Excel12v(xlfRegister, &id, n, popers);
	return ret;
}

#if 0
static int RegisterFunctionTest(LPXLOPER12 dllName)
{
	const int N = 6;

	XLOPER12 opers[N];
	xll::SetValue(&opers[1], 1.0); // entry point ordinal
	xll::SetValue(&opers[2], L"C!");
	xll::SetValue(&opers[3], L"MagicString"); // name
	xll::SetValue(&opers[4], L""); // arguments
	xll::SetValue(&opers[5], 1.0); // macro type

	LPXLOPER12 popers[N];
	popers[0] = dllName;
	for (size_t i = 1; i < N; i++)
		popers[i] = &opers[i];

	XLOPER12 id;
	int ret = Excel12v(xlfRegister, &id, N, popers);
	return ret;
}
#endif

// http://blogs.msdn.com/b/oldnewthing/archive/2014/03/21/10509670.aspx
// The DLL that links to this LIB must reference some symbol within this
// file to make this OBJ file into the final image. Otherwise the OBJ
// file will not be included and the xlAuto*() functions will not be
// exported, and the XLL will not work.

#define EXPORT_UNDECORATED_NAME comment(linker, "/export:" __FUNCTION__ "=" __FUNCDNAME__)

int WINAPI xlAutoOpen()
{
#pragma EXPORT_UNDECORATED_NAME

	XLOPER12 xDLL;
	if (Excel12(xlGetName, &xDLL, 0) == xlretSuccess)
	{
		for (FunctionInfo &f : XLL_NAMESPACE::FunctionInfo::registry())
		{
			RegisterFunction(&xDLL, f);
		}
		// RegisterFunctionTest(&xDLL);
		Excel12(xlFree, 0, 1, &xDLL);
	}
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

void WINAPI xlAutoFree12(LPXLOPER12 p)
{
#pragma EXPORT_UNDECORATED_NAME
	if (p)
	{
#if XLL_SUPPORT_THREAD_LOCAL
		assert(p == &xllReturnValue);
		XLOPER12_Clear(p);
#else
		XLOPER12_Clear(p);
		free(p);
#endif
	}
}
