////////////////////////////////////////////////////////////////////////////
// Invoke.h
// 
// Declares helper functions to call into Excel.

#pragma once

#include "xlldef.h"
#include <Windows.h>
#include "XLCALL.H"
#include <exception>

namespace XLL_NAMESPACE
{
	// Represents an exception from calling Excel.
	class ExcelException : public std::exception
	{
		int m_errorCode;
		char m_errorMessage[100];

	public:
		explicit ExcelException(int errorCode);
		const char* what() const override { return m_errorMessage; }
	};
}
