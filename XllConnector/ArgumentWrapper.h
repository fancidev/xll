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

template <typename T> struct fake_dependency : public std::false_type {};

template <typename T>
struct ArgumentWrapper
{
	static_assert(fake_dependency<T>::value,
		"Do not know how to wrap the supplied argument type. "
		"Specialize xll::ArgumentWrapper<T> to support it.");
	// typedef ?? wrapped_type;
	// static inline T unwrap(wrapped_type value);
};

#define DEFINE_SIMPLE_ARGUMENT_WRAPPER(NativeType, WrappedType) \
template <> struct ArgumentWrapper<NativeType> \
	{ \
	typedef WrappedType wrapped_type; \
	static inline NativeType unwrap(WrappedType v) { return v; } \
	}

DEFINE_SIMPLE_ARGUMENT_WRAPPER(int, int);
DEFINE_SIMPLE_ARGUMENT_WRAPPER(double, double);
DEFINE_SIMPLE_ARGUMENT_WRAPPER(const char *, const char *);
DEFINE_SIMPLE_ARGUMENT_WRAPPER(std::wstring, const wchar_t *);

template <> struct ArgumentWrapper < VARIANT >
{
	typedef LPXLOPER12 wrapped_type;
	static VARIANT unwrap(LPXLOPER12 v)
	{
		// TODO: use a wrapper to free memory
		return make<VARIANT>(*v);
	}
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
	SafeArrayWrapper(const XLOPER12 *pv) : psa(make<SAFEARRAY*>(*pv)){}
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

template <> struct ArgumentWrapper < SAFEARRAY* >
{
	typedef LPXLOPER12 wrapped_type;
	static SafeArrayWrapper unwrap(LPXLOPER12 v)
	{
		return SafeArrayWrapper(v);
	}
};

template <typename T> struct ArgumentWrapper<T &> : ArgumentWrapper < T > {};
template <typename T> struct ArgumentWrapper<T &&> : ArgumentWrapper < T > {};
template <typename T> struct ArgumentWrapper<T const> : ArgumentWrapper < T > {};
template <typename T> struct ArgumentWrapper<T volatile> : ArgumentWrapper < T > {};
template <typename T> struct ArgumentWrapper<T const volatile> : ArgumentWrapper < T > {};

XLL_END_NAMESPACE