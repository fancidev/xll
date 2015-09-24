////////////////////////////////////////////////////////////////////////////
// Conversion.h -- helper functions to convert between data types

#pragma once

#include "xlldef.h"
#include <string>

//
// SetValue(Destination, Source)
//
// Sets the value at Destination to Source by making a deep copy.
//
// XLL Connector uses (the overloads of) this function to perform all
// conversions between data types. A deep copy is always made; hence
// Source can be released after the call.
// 
// 
//
// This template function is used as the universal conversion mechanism
// in this library. It makes it possible to convert between values
// without the need of dedicated constructors or type cast operators.
// 
// For example, to convert a double to VARIANT, use
//
//   VARIANT dest = make<VARIANT>(1.0);
//
// To convert a VARIANT to a double, use
//
//   double x = make<double>(dest);
//
// Note that the function always creates a new object. For example, the
// following code creates a deep copy of an existing VARIANT:
//
//   VARIANT copy = make<VARIANT>(dest);
//
// If conversion fails because of the values are not compatible, an
// implementation should throw a BadConversion exception. If conversion
// fails for other reasons such as memory allocation failure, the
// implementation should throw the corresponding appropriate exception.
//
// This library provides several routines to convert between the common
// data types used in marshalling.
//

namespace XLL_NAMESPACE
{
	// TODO: should be called CreateValue because we assume
	// the dest to be uninitialized.

	//
	// Conversions to XLOPER12.
	//

	HRESULT SetValue(LPXLOPER12, const XLOPER12 &);
	HRESULT SetValue(LPXLOPER12, double);
	HRESULT SetValue(LPXLOPER12, int);
	HRESULT SetValue(LPXLOPER12, unsigned long);
	HRESULT SetValue(LPXLOPER12, bool);
	HRESULT SetValue(LPXLOPER12, const wchar_t *, size_t);
	HRESULT SetValue(LPXLOPER12, const wchar_t *);
	HRESULT SetValue(LPXLOPER12, const std::wstring &);

	//
	// Conversions to VARIANT.
	//

	HRESULT SetValue(VARIANT*, const XLOPER12 &);
	//void ClearValue(VARIANT*);

	//
	// Conversions to LPSAFEARRAY.
	//

	HRESULT SetValue(SAFEARRAY**, const XLOPER12 &);
	//void ClearValue(SAFEARRAY*);
}