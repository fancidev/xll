////////////////////////////////////////////////////////////////////////////
// Conversion.h -- helper functions to convert between data types

#pragma once

#include "xlldef.h"
#include <string>

//
// CreateValue(Destination, Source)
//
// Creates a value at Destination by making a deep copy of Source.
//
// This function behaves like a constructor. The memory at Destination
// is assumed to be uninitialized before the call. If the function
// succeeds, the memory must have been properly initialized. If the
// function fails, the memory remains uninitialized.
//
// This function always makes a deep copy. For example, the following
// line creates a VARIANT from a double, and then makes a deep copy
// of it:
//
//   VARIANT u, v;
//   CreateValue(&u, 1.0);
//   CreateValue(&v, u);
//
// If the function succeeds, it returns S_OK. If it fails, it returns
// an HRESULT error code. Always check the return value to make sure
// the value is actually created.
//
// To destroy a created value, call DeleteValue().
//
// XLL Connector uses (an overload of) this function to perform all
// data type conversions; in particular, it marshals the return value
// of a UDF to XLOPER12 using one of the overloads of CreateValue().
// Therefore, you must overload this function to support marshalling
// your custom type as return value.
//

namespace XLL_NAMESPACE
{
	//
	// Conversions to XLOPER12.
	//

	template <typename T>
	HRESULT CreateValue(LPXLOPER12, T)
	{
		static_assert(false, 
			"Don't know how to convert the given type to XLOPER12. "
			"Overload CreateValue(LPXLOPER12, T) to implement it.");
	}
	HRESULT CreateValue(LPXLOPER12, const XLOPER12 &);
	HRESULT CreateValue(LPXLOPER12, double);
	HRESULT CreateValue(LPXLOPER12, int);
	HRESULT CreateValue(LPXLOPER12, unsigned long);
	HRESULT CreateValue(LPXLOPER12, bool);
	HRESULT CreateValue(LPXLOPER12, const wchar_t *, size_t);
	HRESULT CreateValue(LPXLOPER12, const wchar_t *);
	HRESULT CreateValue(LPXLOPER12, const std::wstring &);
	HRESULT DeleteValue(LPXLOPER12);

	// Conversions from XLOPER12.
	HRESULT CreateValue(double*, const XLOPER12 &);

	//
	// Conversions to VARIANT.
	//

	HRESULT CreateValue(VARIANT*, const XLOPER12 &);
	HRESULT DeleteValue(VARIANT*);

	//
	// Conversions to LPSAFEARRAY.
	//

	HRESULT CreateValue(SAFEARRAY**, const XLOPER12 &);
	HRESULT DeleteValue(SAFEARRAY**);
}