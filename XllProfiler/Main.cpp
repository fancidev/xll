#include <Windows.h>
#include "XLCALL.H"
#include "ExcelHelper.h"
#include <stack>
#include <map>

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

static RegisteredFunctionInfo* __cdecl GetStubFunctionInfo(void *stubReturnAddress)
{
	// TODO: do the actual work
	return pFunctionInfo;
}

// returns the address of the real function
static FARPROC __cdecl precall(void *stubReturnAddress, void *returnAddress, ...)
{
	RegisteredFunctionInfo *pInfo = GetStubFunctionInfo(stubReturnAddress);

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

// The import entry looks like this:
// JMP ds:[xx yy zz ww]
//
// We replace ds:[xx yy zz ww] with the address of our own stub.
// This stub contains a single instruction:
//   call thunk
// This pushes the EIP onto the stack, so that "thunk" can know
// which function is called from.
//
// When the thunk is called, the stack looks like this:
//
// ARG(n)
// ...
// ARG2
// ARG1
// REAL RETURN ADDRESS
// STUB RETURN ADDRESS <- ESP
//
__declspec(naked) void thunk()
{
	__asm
	{
		; // Top of stack contains stub return address; it is used
		; // to identify the UDF being called. On top of it is the
		; // real return address.
		call precall;
		; // EAX now contains the address of the UDF.

		; // Pops StubReturnAddress and RealReturnAddress off the
		; // stack.
		add esp, 8;
		; // Now ESP points to the first actual argument to the UDF.

		; // Call the UDF now.
		call eax;

		; // Now EAX contains the return value of the UDF.
		; // Because all UDFs are __stdcall, the arguments are
		; // already popped from the stack.

		; // All return value types except double return value
		; // in EAX; double returns value in ST(0). We push
		; // both onto the stack, and then call postcall to
		; // log the return value.
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

typedef const void * address_t;

static const address_t s_thunkAddress = thunk;

__declspec(naked) void stub()
{
	__asm call [s_thunkAddress];
}

struct ThunkInfo
{
	address_t stubReturnAddress;
};

std::map<address_t, ThunkInfo> s_installedThunks;

#pragma pack(push, 1)
struct StubInstruction // CALL [s_thunkAddress]
{
	unsigned char opcode[2]; // FF 15 
	address_t target; // &s_thunkAddress
};
#pragma pack(pop)

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
