////////////////////////////////////////////////////////////////////////////
// Conversion.h -- helper functions to convert between data types

#pragma once

#include "xlldef.h"
#include <string>

XLL_BEGIN_NAMEPSACE

//
// make -- creates a value of type T from the value of 'src'
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

#define XLL_ALLOW_MAKE_FROM(TFrom) \
	template <typename T> T make(TFrom) { \
		static_assert(false, "Don't know how to make the requested conversion."); \
	}

XLL_ALLOW_MAKE_FROM(bool);
XLL_ALLOW_MAKE_FROM(int);
XLL_ALLOW_MAKE_FROM(unsigned int);
XLL_ALLOW_MAKE_FROM(unsigned long);
XLL_ALLOW_MAKE_FROM(double);
XLL_ALLOW_MAKE_FROM(const wchar_t *);
XLL_ALLOW_MAKE_FROM(const std::wstring &);
XLL_ALLOW_MAKE_FROM(const XLOPER12 &);

// enum class 
//In these implementations, if the
// source data is too large to fit into the destination type, it checks
// the return value of GetTruncationPolicy() to determine the action.
//void SetTruncationPolicy();
//int GetTruncationPolicy();

//
// Conversions to XLOPER12.
//

template <> XLOPER12 make<XLOPER12>(const XLOPER12 &);
template <> XLOPER12 make<XLOPER12>(double);
template <> XLOPER12 make<XLOPER12>(bool);
template <> XLOPER12 make<XLOPER12>(int);
template <> XLOPER12 make<XLOPER12>(unsigned long);
template <> XLOPER12 make<XLOPER12>(unsigned int);
template <> XLOPER12 make<XLOPER12>(const wchar_t *);
template <> XLOPER12 make<XLOPER12>(const std::wstring &);

//inline void XLOPER12_Create(LPXLOPER12 pv, unsigned long value)
//{
//	XLOPER12_Create(pv, static_cast<double>(value));
//}

//
// Conversions to VARIANT.
//

template <> VARIANT make<VARIANT>(const XLOPER12 &);

//
// Conversions to LPSAFEARRAY.
//

template <> LPSAFEARRAY make<LPSAFEARRAY>(const XLOPER12 &);

XLL_END_NAMESPACE