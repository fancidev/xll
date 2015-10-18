#include <Windows.h>
#include "XLCALL.H"
#include "ExcelHelper.h"
#include <stack>

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

static RegisteredFunctionInfo *pFunctionInfo;

struct CallStackEntry
{
	const RegisteredFunctionInfo *pInfo;
	void *returnAddress;
};

static std::stack<CallStackEntry> thunkCallStack;

// returns the address of the real function
static FARPROC __cdecl precall(RegisteredFunctionInfo *pInfo, void *returnAddress, ...)
{
	std::wstring msg;
	msg = L"Precall(" + pInfo->functionName + L")\n";
	OutputDebugStringW(msg.c_str());

	CallStackEntry entry;
	entry.pInfo = pInfo;
	entry.returnAddress = returnAddress;
	thunkCallStack.push(entry);
	return pInfo->realProcAddress;
}

// Return the real return address.
static void* __cdecl postcall(size_t returnValue)
{
	CallStackEntry entry = thunkCallStack.top();
	thunkCallStack.pop();
	std::wstring msg = L"Postcall(" + entry.pInfo->functionName + L")\n";
	OutputDebugStringW(msg.c_str());
	return entry.returnAddress;
}

// When the thunk is called, the stack looks like this:
//
// ARG(n)
// ...
// ARG2
// ARG1
// RETURN ADDRESS <- ESP
//
__declspec(naked) void thunk()
{
	__asm
	{
		; // top of stack contains return address
		push pFunctionInfo;
		call precall;
		; // EAX now contains the address of the UDF

		; // cleans the stack (pFunctionInfo & return address)
		add esp, 8;

		; // ESP points to the first real argument to the UDF
		; // Call the UDF now.
		call eax;

		; // Now EAX contains the return value of the UDF.
		; // Because all UDF are __stdcall, the arguments are
		; // cleared from the stack.
		; // TODO: handle double return value

		; // Log the return value.
		push eax;
		call postcall;

		; // Now EAX contains the real return address.
		; // Save it on the stack so as to return to it later.
		push eax;

		; // Restore the return value of the UDF into EAX.
		mov eax, ss:[esp + 4];

		; // Return back to the real return address.
		ret 8
	}
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
		*pAddress = (FARPROC)thunk;

		pFunctionInfo = new RegisteredFunctionInfo;
		*pFunctionInfo = *pInfo;

		return true;
	}
	return false;
}

bool __stdcall IsProfilerPresent()
{
	return true;
}

int WINAPI xlAutoOpen()
{
	std::vector<RegisteredFunctionInfo> info;
	GetRegisteredFunctions(info);

	for (size_t i = 0; i < info.size(); i++)
	{
		if (info[i].functionName == L"GetCircleArea")
		{
			InstallThunk(&info[i]);
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
