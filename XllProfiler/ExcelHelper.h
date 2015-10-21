////////////////////////////////////////////////////////////////////////////
// ExcelHelper.h -- utility functions for XLL

#include <Windows.h>
#include <string>
#include <vector>

struct RegisteredFunctionInfo
{
	double id;
	std::wstring functionName;
	std::wstring dllName;
	std::wstring entryPointName;
	std::wstring typeText;
	FARPROC procAddress; // as returned by GetProcAddress
};

void GetRegisteredFunctions(std::vector<RegisteredFunctionInfo> &info);