////////////////////////////////////////////////////////////////////////////
// StringExample.cpp
//
// This file demonstrates how to take string arguments from and return
// string to Excel. 
//
// The following types are supported as arguments and return value:
//
//   [const] char * [const]
//   [const] wchar_t * [const]
//   [const] std::wstring [&]
//
// In Excel 2003 and earlier, strings are ansi-encoded and can contain
// up to 255 characters (not including the nul-terminator). A passed-in
// string argument will never contain more bytes. If a string longer than
// 255 bytes is returned, the wrapper throws std::invalid_argument() and
// returns #VALUE! to Excel.
//
// In Excel 2007 and later, strings are unicode-encoded and can contain
// up to 32,767 characters, not including the nul-terminator. A passed-in
// argument will never contain more characters. If a string longer than
// 32,767 characters is returned, the wrapper throws std::invalid_argument()
// and returns #VALUE! to Excel.
//
// The wrapper automatically performs ansi-unicode conversion where
// necessary. If some characters cannot be converted, it is replaced by a
// placeholder.
//
// When passed as argument, wchar_t* and char* are more efficient than
// their std::[w]string counterparts because they are passed directly 
// from Excel and incurs no allocation or copying overhead. When used
// in return value, strings are always copied, so there is not much
// difference in performance.

#include "XllAddin.h"

// ReverseString:
//   Simple example that takes a string as input and returns a string
//   with its characters reversed.
std::wstring ReverseString(const std::wstring &s)
{
	return std::wstring(s.crbegin(), s.crend());
}

EXPORT_XLL_FUNCTION(ReverseString, XLL_NOT_VOLATILE | XLL_THREADSAFE);

// GetTooLongString:
//   Returns a string of 40,000 characters. The wrapper throws an exception
//   and returns #VALUE! to Excel.
std::wstring GetTooLongString()
{
	return std::wstring(40000, L'x');
}

EXPORT_XLL_FUNCTION(GetTooLongString, XLL_NOT_VOLATILE | XLL_THREADSAFE);

// MultiByteStrLen:
//   Returns the length of a byte string. The purpose of this function
//   is to test string marshalling. If marshalled as "C", the maximum
//   input allowed is 255 bytes; if a longer string is input, Excel returns
//   #VALUE! directly without calling the function at all.
int MultiByteStrLen(const char *s)
{
	return static_cast<int>(strlen(s));
}

EXPORT_XLL_FUNCTION(MultiByteStrLen);
