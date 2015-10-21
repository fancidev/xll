#include <Windows.h>
#include "XLCALL.H"
#include "ExcelHelper.h"
#include "ThunkManager.h"

BOOL WINAPI DllMain(HANDLE hInstance, ULONG fdwReason, LPVOID lpReserved)
{
	switch (fdwReason)
	{
	case DLL_PROCESS_DETACH:
		// Uninstall thunks
		break;
	}
	return TRUE;
}

static ThunkManager s_thunkManager;

#if 0
static RegisteredFunctionInfo *pFunctionInfo;

DWORD GetPageSize()
{
	SYSTEM_INFO systemInfo;
	GetSystemInfo(&systemInfo);
	return systemInfo.dwPageSize;
}

bool InstallThunk(RegisteredFunctionInfo *pInfo)
{
	unsigned char *instruction = (unsigned char *)(pInfo->entryPointAddress);

	// Must be a JMP [...] instruction of the form 
	// FF 25 xx yy zz ww   jmp ds:[xx yy zz ww]
	// Otherwise we don't know how to insert the thunk.
	if (instruction[0] == 0xFF && instruction[1] == 0x25)
	{
		FARPROC *pAddress;
		memcpy(&pAddress, &instruction[2], 4);
		pInfo->importThunkLocation = pAddress;
		pInfo->realProcAddress = *pAddress;

		// Create a code stub with a single instruction:
		// CALL thunk

		// Allocate a page with read/write access.
		DWORD dwPageSize = GetPageSize();
		LPVOID page = VirtualAlloc(NULL, 100, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE);
		if (page != nullptr)
		{
			StubInstruction inst = { { 0xFF, 0x15 }, &s_thunkAddress }; // CALL thunk
			memcpy(page, &inst, sizeof(StubInstruction));

			bool ok = false;
			DWORD dwOldProtect;
			if (VirtualProtect(page, 100, PAGE_EXECUTE_READ, &dwOldProtect) &&
				FlushInstructionCache(GetCurrentProcess(), page, 100))
			{
				ok = true;
			}
		}

		// *pAddress = (FARPROC)stub;
		*pAddress = (FARPROC)page;

		pFunctionInfo = new RegisteredFunctionInfo;
		*pFunctionInfo = *pInfo;

		return true;
	}
	return false;
}
#endif

void XllBeforeCall(ThunkInfo *pThunkInfo, void *returnAddress, va_list)
{
	RegisteredFunctionInfo *pFunctionInfo = (RegisteredFunctionInfo*)pThunkInfo->cookie;
	wchar_t msg[1000];
	swprintf_s(msg, L"BeforeCall:%s:%s\n",
		pFunctionInfo->functionName.c_str(), 
		pFunctionInfo->typeText.c_str());
	OutputDebugStringW(msg);
}

void XllAfterCall(ThunkInfo *pThunkInfo, size_t intRetVal)
{
	RegisteredFunctionInfo *pFunctionInfo = (RegisteredFunctionInfo*)pThunkInfo->cookie;
	wchar_t msg[1000];
	swprintf_s(msg, L"AfterCall:%s:0x%p\n",
		pFunctionInfo->functionName.c_str(),
		intRetVal);
	OutputDebugStringW(msg);
}

bool __stdcall IsProfilerPresent()
{
	return true;
}

static void InstallThunk(const RegisteredFunctionInfo &functionInfo)
{
	RegisteredFunctionInfo *pFunctionInfo =
		new RegisteredFunctionInfo(functionInfo);
	s_thunkManager.InstallThunk(NULL,
		pFunctionInfo->procAddress, pFunctionInfo,
		XllBeforeCall, XllAfterCall);
}

int WINAPI xlAutoOpen()
{
	std::vector<RegisteredFunctionInfo> info;
	GetRegisteredFunctions(info);

	for (size_t i = 0; i < info.size(); i++)
	{
		if (info[i].functionName == L"GetCircleArea")
		{
			InstallThunk(info[i]);
		}
	}

	// Register some function to prevent this DLL from
	// being unloaded by Excel.
	XLOPER12 xDllName;
	if (Excel12(xlGetName, &xDllName, 0) == xlretSuccess)
	{
		XLOPER12 xProcedure, xSignature, xFunction;
		xProcedure.xltype = xltypeStr;
		xProcedure.val.str = L"\021IsProfilerPresent";
		xSignature.xltype = xltypeStr;
		xSignature.val.str = L"\01A";
		xFunction.xltype = xltypeStr;
		xFunction.val.str = L"\021IsProfilerPresent";
		Excel12(xlfRegister, nullptr, 4, &xDllName, &xProcedure, &xSignature, &xFunction);
		Excel12(xlFree, 0, 1, xDllName);
	}

	return 1;
}

int WINAPI xlAutoClose()
{
	return 1;
}
