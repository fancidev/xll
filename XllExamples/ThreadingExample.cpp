#include "XllAddin.h"

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
