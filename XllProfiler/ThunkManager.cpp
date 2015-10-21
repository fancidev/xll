////////////////////////////////////////////////////////////////////////////
// ThunkManager.cpp -- instrument DLL routines through thunking

#include "ThunkManager.h"
#include <string>
#include <stack>

struct CallStackEntry
{
	ThunkInfo *thunk;
	void *returnAddress;
};

static std::stack<CallStackEntry> thunkCallStack;

static void DefaultBeforeCallHandler(ThunkInfo *pThunkInfo, void *returnAddress, va_list)
{
	std::wstring msg;
	// msg = L"Precall(" + pInfo->functionName + L")\n";
	msg = L"Precall()\n";
	OutputDebugStringW(msg.c_str());
}

static void DefaultAfterCallHandler(ThunkInfo *pThunkInfo, size_t intReturnValue)
{
	//std::wstring msg = L"Postcall(" + entry.pInfo->functionName + L")\n";
	std::wstring msg = L"Postcall()\n";
	OutputDebugStringW(msg.c_str());
}

// Returns the actual entry point address of the dll procedure.
static FARPROC __cdecl precall(ThunkInfo **ppThunkInfo, void *returnAddress, ...)
{
	ThunkInfo *pThunkInfo = *ppThunkInfo;

	if (pThunkInfo->beforeCall)
	{
		va_list ap;
		va_start(ap, returnAddress);
		pThunkInfo->beforeCall(pThunkInfo, returnAddress, ap);
		va_end(ap);
	}

	CallStackEntry entry;
	entry.thunk = pThunkInfo;
	entry.returnAddress = returnAddress;
	thunkCallStack.push(entry);
	return pThunkInfo->procEntryPoint;
}

// Returns the real return address.
static void* __cdecl postcall(size_t returnValue)
{
	CallStackEntry entry = thunkCallStack.top();
	thunkCallStack.pop();
	if (entry.thunk->afterCall)
	{ 
		entry.thunk->afterCall(entry.thunk, returnValue);
	}
	return entry.returnAddress;
}

// GetProcAddress() returns the address of an import entry. The
// import entry looks like the following:
//
//   JMP ds:[...]
//
// where ds:[...] contains the true entry point of the procedure.
//
// We replace the contents in ds:[...] with the address of a stub
// generated at run-time. This stub contains a single instruction
// followed by a LPVOID datum:
//
//   CALL thunk
//   DD pThunkInfo
//
// This instruction pushes the EIP onto the stack, which happens
// to point to the address containing pThunkInfo. This allows 
// thunk() to know which dll procedure is being called.
//
// When thunk() is called, the stack looks like this:
//
//   ARG(n)
//   ...
//   ARG2
//   ARG1
//   REAL RETURN ADDRESS
//   pThunkInfo Address  <- ESP
//
__declspec(naked) void thunk()
{
	__asm
	{
		; // Top of stack contains the address that contains a
		; // pointer to ThunkInfo. It is used to identify the 
		; // dll procedure being called. On top of it is the
		; // real return address.
		call precall;
		; // EAX now contains the real entry point address of 
		; // the dll procedure.

		; // Pop ppThunkInfo and returnAddress off the stack.
		add esp, 8;
		; // Now ESP points to the first actual argument to the 
		; // dll procedure.

		; // Call the DLL procedure now.
		call eax;

		; // Now EAX contains the return value of the DLL procedure.
		; // Because we only support procedures using the __stdcall
		; // calling convention, the arguments are already popped
		; // off the stack by the callee.

		; // All return value types except double are stored in EAX;
		; // double return value is stored in ST(0). We push both
		; // onto the stack and then call postcall to log the return
		; // value.
		push eax;
		call postcall;

		; // Now EAX contains the real return address.
		; // Save it on the stack so as to return to it later.
		push eax;

		; // Restore the original return value of the dll procedure.
		mov eax, ss:[esp + 4];

		; // Return back to the real return address.
		ret 8
	}
}

// CC                 INT 3
// CC                 INT 3
// FF 15 xx xx xx xx  CALL [s_thunkAddress]
// DD yy yy yy yy
struct StubInstruction // CALL [s_thunkAddress]
{
	unsigned char padding[sizeof(size_t) - 2]; // CC CC
	unsigned char opcode[2];  // FF 15
	void **pThunkAddress;     // &s_thunkAddress
	ThunkInfo *pThunkInfo;
};
static_assert(sizeof(StubInstruction) == sizeof(size_t) * 3, "");
static void* s_thunkAddress = thunk;

ThunkManager::ThunkManager()
{
	// Create a heap to allocate memory to store stub code.
	m_hCodeHeap = HeapCreate(HEAP_CREATE_ENABLE_EXECUTE, 0, 0);
}

ThunkManager::~ThunkManager()
{
	if (m_hCodeHeap != NULL)
	{
		HeapDestroy(m_hCodeHeap);
		m_hCodeHeap = NULL;
	}
}

#pragma pack(push, 1)
struct JmpInstruction
{
	unsigned char opcode[2]; // FF 25
	FARPROC* pEntryPoint;
};
#pragma pack(pop)

BOOL ThunkManager::InstallThunk(
	HMODULE hModule, FARPROC procAddress, void* cookie,
	BeforeCallHandler beforeCall, AfterCallHandler afterCall)
{
	if (procAddress == NULL)
		return FALSE;

	if (m_hCodeHeap == NULL)
		return FALSE;

	// The code at procAddress must be a JMP instruction of the form
	//
	//   FF 25 xx xx xx xx   jmp ds:[xx xx xx xx]
	//
	// where [xx xx xx xx] contains the actual entry point address
	// of the procedure. If the instruction is not of this form, we
	// do not support thunking it.
	const JmpInstruction *instruction = (const JmpInstruction *)procAddress;
	if (!(instruction->opcode[0] == 0xFF && instruction->opcode[1] == 0x25))
		return FALSE;

	// Allocate a code stub from the code heap.
	StubInstruction *stub = static_cast<StubInstruction*>(
		HeapAlloc(m_hCodeHeap, 0, sizeof(StubInstruction)));
	if (stub == NULL)
		return FALSE;

	// Create book-keeping entry for the thunk.
	ThunkInfo *pThunkInfo = new ThunkInfo;
	pThunkInfo->hModule = hModule;
	pThunkInfo->procAddress = procAddress;
	pThunkInfo->cookie = cookie;
	pThunkInfo->procEntryPoint = *instruction->pEntryPoint;
	pThunkInfo->stubEntryPoint = (FARPROC)&stub->opcode;
	pThunkInfo->beforeCall = beforeCall;
	pThunkInfo->afterCall = afterCall;

	// Fill the instructions in the stub.
	memset(stub->padding, 0xCC, sizeof(stub->padding)); // INT 3
	stub->opcode[0] = 0xFF;
	stub->opcode[1] = 0x15;
	stub->pThunkAddress = &s_thunkAddress;
	stub->pThunkInfo = pThunkInfo;

	// Flush instruction cache.
	if (!FlushInstructionCache(GetCurrentProcess(), stub, sizeof(StubInstruction)))
	{
		HeapFree(m_hCodeHeap, 0, stub);
		return FALSE;
	}

	// TODO: make it exception free
	m_thunks.push_back(pThunkInfo);

	// TODO: lock the library in memory, or register a hook to
	// uninstall the thunk when the DLL is unloaded.

	// Redirect the entry point to our thunk.
	*instruction->pEntryPoint = pThunkInfo->stubEntryPoint;

	return TRUE;
}
