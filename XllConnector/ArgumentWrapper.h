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
// The following diagram illustrates the steps in a call from Excel to
// a UDF:
//
//             +-----------------+                   +-----------------+
//   Excel     |   XL Arg Type   |                   |   XL Ret Type   |
//             +--------+--------+                   +--------+--------+
//                      |                                     ^
//                      v                                     |
//   Marshaler       unwrap                                 wrap
//                      |                                     ^
//                      v                                     |
//             +--------+--------+                   +--------+--------+
//   UDF       | Native Arg Type | ---> Compute ---> | Native Ret Type |
//             +-----------------+                   +-----------------+
//
// 
// When Excel calls a UDF, it supports passing arguments of type LPXLOPER12
// as well as several other native types. We call these "wrapped arguments".
// We must "unwrap" these incoming arguments before forwarding the call to
// the UDF. Likewise, we must "wrap" the return value of the udf to one of
// the several return types supported by Excel.
//

XLL_BEGIN_NAMEPSACE

// template <typename T> struct fake_dependency : public std::false_type {};

// ArgumentMarshaler<T> -- marshal an argument of native type T from
// Excel to UDF. The default implementation exposes the argument as
// LPXLOPER12 and uses make<T>() to convert it. This requires the
// type cast operator to be provided for that type.
//
// NativeType: the argument type of the UDF
// StorageType: an intermediary type that is convertible from
// XLOPER12 and convertible to NativeType. 
// The purpose of StorageType is to automatically release the
// intermediary object after the call.

//template <typename NativeType>
//struct ArgumentMarshaler
//{
//	typedef LPXLOPER12 MarshaledType;
//	static inline const wchar_t * GetTypeText() { return L"Q"; }
//	static inline NativeType Marshal(MarshaledType arg)
//	{
//		return make<NativeType>(*arg);
//	}
//};

//#define PACK_TYPE_TEXT(c1, c2) ((int)(unsigned char)(c1) | ((int)(unsigned char)(c2) << 8))
//#define UNPACK_TYPE_TEXT(x) (char)((x) & 0xFF), (char)(((x)>>8) & 0xFF)

template <typename T>
struct ArgumentMarshaler
{
	// typedef T NativeType;
	// typedef ? MarshaledType;
	// typedef ? StorageType;
	// static inline StorageType Marshal(MarshaledType arg);
	// enum { PackedTypeText = ? };
	// static inline LPCWSTR GetTypeText() { return L"Q"; }
	template <typename U> struct always_false : std::false_type {};
	static_assert(always_false<T>::value,
		"Do not know how to marshal the supplied argument type. "
		"Specialize xll::ArgumentMarshaler<T> to support it.");
};

template <
	typename TNative, 
	char TypeChar1, 
	char TypeChar2 = 0, 
	typename TMarshaled = TNative, 
	typename TStorage = TNative>
struct ArgumentMarshalerImpl
{
	typedef TNative NativeType; // UserType, UdfType, udf_arg_type
	typedef TMarshaled MarshaledType; // WireType, XllType, xll_arg_type
	typedef TStorage StorageType;
	static inline const wchar_t * GetTypeText()
	{
		static const wchar_t typeText[] = { TypeChar1, TypeChar2, 0 };
		return typeText;
	}
	static inline TStorage Marshal(TMarshaled arg)
	{
		return arg;
	}
};

#define IMPLEMENT_ARGUMENT_MARSHALER(UserType, ...) \
	template <> struct ArgumentMarshaler<UserType> \
		: ArgumentMarshalerImpl<UserType, __VA_ARGS__> {}

IMPLEMENT_ARGUMENT_MARSHALER(double, 'B');
IMPLEMENT_ARGUMENT_MARSHALER(int, 'J');
IMPLEMENT_ARGUMENT_MARSHALER(std::wstring, 'C', '%', LPCWSTR, std::wstring);
IMPLEMENT_ARGUMENT_MARSHALER(const std::wstring &, 'C', '%', LPCWSTR, std::wstring);
IMPLEMENT_ARGUMENT_MARSHALER(const char *, 'C');

// TODO: use a wrapper to free memory
IMPLEMENT_ARGUMENT_MARSHALER(VARIANT, 'Q', 0, LPXLOPER12);

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

IMPLEMENT_ARGUMENT_MARSHALER(SAFEARRAY*, 'Q', 0, LPXLOPER12, SafeArrayAdapter);

//template <typename T> struct ArgumentWrapper<T &> : ArgumentWrapper < T > {};
//template <typename T> struct ArgumentWrapper<T &&> : ArgumentWrapper < T > {};
//template <typename T> struct ArgumentWrapper<T const> : ArgumentWrapper < T > {};
//template <typename T> struct ArgumentWrapper<T volatile> : ArgumentWrapper < T > {};
//template <typename T> struct ArgumentWrapper<T const volatile> : ArgumentWrapper < T > {};

XLL_END_NAMESPACE