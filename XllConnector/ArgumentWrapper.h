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

template <typename T>
struct ArgumentMarshaler
{
	// typedef T NativeType;
	// typedef ? MarshaledType;
	// typedef ? StorageType;
	// static inline StorageType Marshal(MarshaledType arg);
	// static inline LPCWSTR GetTypeText() { return L"Q"; }
	template <typename U> struct always_false : std::false_type {};
	static_assert(always_false<T>::value,
		"Do not know how to marshal the supplied argument type. "
		"Specialize xll::ArgumentMarshaler<T> to support it.");
};

template <typename T, typename TStorage = T>
struct VariantArgumentMarshaler
{
	typedef T NativeType;
	typedef LPXLOPER12 MarshaledType;
	static inline const wchar_t * GetTypeText() { return L"Q"; }
	static inline TStorage Marshal(LPXLOPER12 arg)
	{
		//return make<TStorage>(*arg);
		return TStorage(*arg);
	}
};

template <typename T, char TypeChar1, char TypeChar2 = 0>
struct SimpleArgumentMarshaler
{
	typedef T NativeType;
	typedef T UdfType;
	typedef T MarshaledType;
	typedef T XllType;
	static inline const wchar_t * GetTypeText()
	{
		static const wchar_t typeText[] = { TypeChar1, TypeChar2, 0 };
		return typeText;
	}
	static inline T Marshal(T arg)
	{
		return arg;
	}
};

template <typename TNative, typename TMarshaled, typename TStorage, char TypeChar1, char TypeChar2 = 0>
struct ArgumentMarshalerImpl
{
	typedef TNative NativeType;
	typedef TMarshaled MarshaledType;
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

#define DEFINE_SIMPLE_ARGUMENT_MARSHALER(nativeType, marshaledType, typeText) \
	template <> struct ArgumentMarshaler<nativeType> \
	{ \
		typedef marshaledType MarshaledType; \
		static inline const wchar_t * GetTypeText() { return L#typeText; } \
		static inline nativeType Marshal(marshaledType v) { return v; } \
	}

template <> struct ArgumentMarshaler<double> : SimpleArgumentMarshaler<double, 'B'> {};
template <> struct ArgumentMarshaler<int> : SimpleArgumentMarshaler<int, 'J'>{};
template <> struct ArgumentMarshaler<const std::wstring &> 
	: ArgumentMarshalerImpl < std::wstring, LPCWSTR, std::wstring, 'C', '%' > {};
template <> struct ArgumentMarshaler<const char *>
	: ArgumentMarshalerImpl < const char *, const char *, const char *, 'C' > {};
//DEFINE_SIMPLE_ARGUMENT_WRAPPER(int, int, "J");
//DEFINE_SIMPLE_ARGUMENT_WRAPPER(double, double, "B");
//DEFINE_SIMPLE_ARGUMENT_WRAPPER(const char *, const char *, "C");
//DEFINE_SIMPLE_ARGUMENT_WRAPPER(std::wstring, const wchar_t *, "C%");

// TODO: use a wrapper to free memory
template <> 
struct ArgumentMarshaler<VARIANT> 
	: VariantArgumentMarshaler<VARIANT>
{
};

class SafeArrayWrapper
{
private:
	SAFEARRAY *psa;
public:
	SafeArrayWrapper(const SafeArrayWrapper &) = delete;
	SafeArrayWrapper& operator=(const SafeArrayWrapper &) = delete;
	SafeArrayWrapper(SafeArrayWrapper &&other)
	{
		if (this != &other)
		{
			this->psa = other.psa;
			other.psa = nullptr;
		}
	}
	SafeArrayWrapper(const XLOPER12 &v) : psa(make<SAFEARRAY*>(v)){}
	~SafeArrayWrapper()
	{
		if (psa)
		{
			SafeArrayDestroy(psa);
			psa = nullptr;
		}
	}
	operator SAFEARRAY*() { return psa; }
};

template <> 
struct ArgumentMarshaler<SAFEARRAY*> 
	: VariantArgumentMarshaler<SAFEARRAY*, SafeArrayWrapper>
{
};

//template <typename T> struct ArgumentWrapper<T &> : ArgumentWrapper < T > {};
//template <typename T> struct ArgumentWrapper<T &&> : ArgumentWrapper < T > {};
//template <typename T> struct ArgumentWrapper<T const> : ArgumentWrapper < T > {};
//template <typename T> struct ArgumentWrapper<T volatile> : ArgumentWrapper < T > {};
//template <typename T> struct ArgumentWrapper<T const volatile> : ArgumentWrapper < T > {};

XLL_END_NAMESPACE