////////////////////////////////////////////////////////////////////////////
// Invoke.cpp -- implement helper functions to call into Excel

#include "Invoke.h"
#include <stdio.h>

using namespace XLL_NAMESPACE;

////////////////////////////////////////////////////////////////////////////
// ExcelException implementation

static const char * GetExcelErrorMessage(int errorCode)
{
	switch (errorCode)
	{
	case xlretSuccess:
		return "success";
	case xlretAbort:
		return "macro halted";
	case xlretInvXlfn:
		return "invalid function number";
	case xlretInvCount:
		return "invalid number of arguments";
	case xlretInvXloper:
		return "invalid OPER structure";
	case xlretStackOvfl:
		return "stack overflow";
	case xlretFailed:
		return "command failed";
	case xlretUncalced:
		return "uncalced cell";
	case xlretNotThreadSafe:
		return "not allowed during multi-threaded calc";
	case xlretInvAsynchronousContext:
		return "invalid asynchronous function handle";
	case xlretNotClusterSafe:
		return "not supported on cluster";
	default:
		return nullptr;
	}
}

ExcelException::ExcelException(int errorCode)
	: m_errorCode(errorCode)
{
	const char *s = GetExcelErrorMessage(errorCode);
	if (s == nullptr)
		sprintf_s(m_errorMessage, "xll error %d", m_errorCode);
	else
		sprintf_s(m_errorMessage, "%s", s);
}

////////////////////////////////////////////////////////////////////////////
// Excel invoke implementation

#if 0
//ExcelVariant ExcelCall(int xlfn);
//ExcelVariant ExcelCall(int xlfn, const ExcelVariant &); // cannot pass XLOPER12 &

template <typename... T>
ExcelVariant ExcelCall(int xlfn, const T&... args)
{
	//va_list x;
	//va_start;
	// TODO: this function won't work or would invoke a copy constructor
	// if the supplied argument is of type XLOPER12. This is not what we want.
	const int NumArgs = sizeof...(args);
	ExcelVariant xlArgs[NumArgs] = { args... };
	LPXLOPER12 pArgs[NumArgs];
	for (int i = 0; i < NumArgs; i++)
		pArgs[i] = &xlArgs[i];

	XLOPER12 result;
	int ret = Excel12v(xlfn, &result, numargs, pArgs);
	if (ret != xlretSuccess)
		throw ExcelException(ret);

	ExcelVariant vResult(result);
	Excel12(xlFree, 0, 1, &result);
	return vResult;
}
#endif