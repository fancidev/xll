////////////////////////////////////////////////////////////////////////////
// Marshal.h -- template classes to marshal arguments from Excel to UDF

#pragma once

#include "xlldef.h"
#include <string>
#include "Conversion.h"

//
// Excel supports calling XLL functions with a limited set of argument
// types. To adapt to the native argument types of a UDF, XLL Connector
// creates a wrapper function for each UDF, which marshals the arguments
// and the return value, as well as handling any exception thrown by the
// UDF.
//
// The following diagram illustrates the marshalling steps performed in
// a call from Excel to UDF:
//
//             +-----------------+                   +-----------------+
//   Excel     |  Wire Arg Type  |                   |  Wire Ret Type  |
//             +--------+--------+                   +--------+--------+
//                      |                                     ^
//                      v                                     |
//   Marshaler       marshal                               marshal
//                      |                                     ^
//                      v                                     |
//             +--------+--------+                   +--------+--------+
//   UDF       |  User Arg Type  | ---> Compute ---> |  User Ret Type  |
//             +-----------------+                   +-----------------+
//
// 
// When Excel calls a UDF, it passes arguments in "wire type", such as
// LPXLOPER12 and primitive numeric types. XLL Connector "marshals" these
// arguments into "user type" and pass them to the UDF. This process is
// implemented by the ArgumentMarshaler<T> template class. You can specialize
// this struct to provide marshaling support for custom types.
//
// When the UDF returns, XLL Connector marshals the return value to
// LPXLOPER12 and return back to Excel. If an exception is thrown, the
// wrapper returns #VALUE!.
//

XLL_BEGIN_NAMEPSACE

// ArgumentMarshaler<T> -- marshal an argument of user type T from
// Excel to UDF. The following members are required:
//
//   UserType      argument type of the UDF
//   WireType      type that Excel passes
//   Marshal       function to marshal the argument
//   GetTypeText   returns the type text used in registration
//
// The default implementation produces a static_assert error. 

template <typename T>
struct ArgumentMarshaler
{
	// typedef T UserType;
	// typedef ? WireType;
	// static inline ? Marshal(WireType arg);
	template <typename U> struct always_false : std::false_type {};
	static_assert(always_false<T>::value,
		"Do not know how to marshal the supplied argument type. "
		"Specialize xll::ArgumentMarshaler<T> to support it.");
};

// ArgumentMarshalerImpl - generic implementation of ArgumentMarshaler
//
// ArgUserType     native argument type of udf
// ArgWireType     wire type, default to the same as user type
// ArgAdapterType  return type of Marshal(). This type must be:
//                 1) implicitly constructible from WireType, and 
//                 2) implicitly convertible to UserType.
//                 The adapter object is used to automatically free any
//                 resources allocated during argument marshalling.
template <
	typename ArgUserType, 
	typename ArgWireType = ArgUserType,
	typename ArgAdapterType = ArgUserType>
struct ArgumentMarshalerImpl
{
	typedef ArgUserType UserType;
	typedef ArgWireType WireType;
	static inline ArgAdapterType Marshal(ArgWireType arg)
	{
		return arg;
	}
};

#define IMPLEMENT_ARGUMENT_MARSHALER(UserType, ...) \
	template <> struct ArgumentMarshaler<UserType> \
		: ArgumentMarshalerImpl<UserType, __VA_ARGS__> {}

#define IMPLEMENT_ARGUMENT_MARSHALER_AS(UserType, AsUserType) \
	template <> struct ArgumentMarshaler<UserType> : ArgumentMarshaler<AsUserType> {}

//
// Primitive type marshalling
//

IMPLEMENT_ARGUMENT_MARSHALER(double);
IMPLEMENT_ARGUMENT_MARSHALER(int);
IMPLEMENT_ARGUMENT_MARSHALER(FP12*);

//
// String marshalling
//

IMPLEMENT_ARGUMENT_MARSHALER(const wchar_t *);
IMPLEMENT_ARGUMENT_MARSHALER(std::wstring, LPCWSTR);
IMPLEMENT_ARGUMENT_MARSHALER_AS(const std::wstring &, std::wstring);

class UnicodeToAnsiAdapter
{
	char *m_str;
public:
	UnicodeToAnsiAdapter(const wchar_t *s)
	{
		int cb = WideCharToMultiByte(CP_ACP, 0, s, -1, nullptr, 0, nullptr, nullptr);
		if (cb <= 0)
			throw std::invalid_argument("Input string is not a valid Unicode string.");
		m_str = (char*)malloc((size_t)cb);
		if (m_str == nullptr)
			throw std::bad_alloc();
		if (WideCharToMultiByte(CP_ACP, 0, s, -1, m_str, cb, nullptr, nullptr) <= 0)
		{
			free(m_str);
			throw std::invalid_argument("Cannot convert input string from Unicode to Ansi.");
		}
	}
	UnicodeToAnsiAdapter(UnicodeToAnsiAdapter&& other)
	{
		if (this != &other)
		{
			m_str = other.m_str;
			other.m_str = nullptr;
		}
	}
	UnicodeToAnsiAdapter(const UnicodeToAnsiAdapter &) = delete;
	UnicodeToAnsiAdapter& operator = (const UnicodeToAnsiAdapter &) = delete;
	~UnicodeToAnsiAdapter()
	{
		if (m_str != nullptr)
		{
			free(m_str);
			m_str = nullptr;
		}
	}
	operator char*() { return m_str; }
};

// In Excel 2007 and later, if a string argument is declared as char*, then
// at most 255 bytes can be passed, or #VALUE! is returned. Therefore we 
// always marshal a string as wchar_t*.
IMPLEMENT_ARGUMENT_MARSHALER(const char *, LPCWSTR, UnicodeToAnsiAdapter);

//
// VARIANT marshalling
//

class VariantAdapter
{
private:
	VARIANT m_value;
public:
	VariantAdapter(const VariantAdapter &) = delete;
	VariantAdapter& operator=(const VariantAdapter &) = delete;
	VariantAdapter(VariantAdapter &&other)
	{
		if (this != &other)
		{
			memcpy(&this->m_value, &other.m_value, sizeof(VARIANT));
			memset(&other.m_value, 0, sizeof(VARIANT));
		}
	}
	VariantAdapter(LPXLOPER12 pv)
	{
		m_value = make<VARIANT>(*pv);
	}
	~VariantAdapter()
	{
		VariantClear(&m_value);
	}
	operator VARIANT*() { return &m_value; }
};

IMPLEMENT_ARGUMENT_MARSHALER(VARIANT*, LPXLOPER12, VariantAdapter);

//
// SAFEARRAY marshalling
//

class SafeArrayAdapter
{
private:
	SAFEARRAY *psa;
public:
	SafeArrayAdapter(const SafeArrayAdapter &) = delete;
	SafeArrayAdapter& operator=(const SafeArrayAdapter &) = delete;
	SafeArrayAdapter(SafeArrayAdapter &&other)
	{
		if (this != &other)
		{
			this->psa = other.psa;
			other.psa = nullptr;
		}
	}
	SafeArrayAdapter(LPXLOPER12 pv) : psa(make<SAFEARRAY*>(*pv)){}
	~SafeArrayAdapter()
	{
		if (psa)
		{
			SafeArrayDestroy(psa);
			psa = nullptr;
		}
	}
	operator SAFEARRAY*() { return psa; }
};

IMPLEMENT_ARGUMENT_MARSHALER(SAFEARRAY*, LPXLOPER12, SafeArrayAdapter);

//template <typename T> struct ArgumentWrapper<T &> : ArgumentWrapper < T > {};
//template <typename T> struct ArgumentWrapper<T &&> : ArgumentWrapper < T > {};
//template <typename T> struct ArgumentWrapper<T const> : ArgumentWrapper < T > {};
//template <typename T> struct ArgumentWrapper<T volatile> : ArgumentWrapper < T > {};
//template <typename T> struct ArgumentWrapper<T const volatile> : ArgumentWrapper < T > {};

XLL_END_NAMESPACE