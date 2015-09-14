////////////////////////////////////////////////////////////////////////////
// Addin.h -- defines macros to export functions to Excel

#pragma once

#include "xlldef.h"
#include "FunctionInfo.h"
#include "Conversion.h"
#include "TypeText.h"
#include "Marshal.h"

XLL_BEGIN_NAMEPSACE

FunctionInfoBuilder AddFunction(FunctionInfo &f);

inline LPXLOPER12 getReturnValue()
{
#if XLL_SUPPORT_THREAD_LOCAL
	__declspec(thread) extern XLOPER12 xllReturnValue;
	return &xllReturnValue;
#else
	LPXLOPER12 p = (LPXLOPER12)malloc(sizeof(XLOPER12));
	if (p == nullptr)
		throw std::bad_alloc();
	return p;
#endif
}

template <typename Func> struct StripCallingConvention;
template <typename TRet, typename... TArgs>
struct StripCallingConvention < TRet __cdecl(TArgs...) >
{
	typedef TRet type(TArgs...);
};
#ifndef _WIN64
template <typename TRet, typename... TArgs>
struct StripCallingConvention < TRet __stdcall(TArgs...) >
{
	typedef TRet type(TArgs...);
};
template <typename TRet, typename... TArgs>
struct StripCallingConvention < TRet __fastcall(TArgs...) >
{
	typedef TRet type(TArgs...);
};
#else
template <typename TRet, typename... TArgs>
struct StripCallingConvention < TRet __vectorcall(TArgs...) >
{
	typedef TRet type(TArgs...);
};
#endif

template <typename Func, Func *func, typename TRet, typename... TArgs>
inline LPXLOPER12 XLWrapperImpl(typename ArgumentMarshaler<TArgs>::WireType... args)
{
	try
	{
		LPXLOPER12 pvRetVal = getReturnValue();
		HRESULT hr = SetValue(pvRetVal, func(ArgumentMarshaler<TArgs>::Marshal(args)...));
		if (FAILED(hr))
			throw std::invalid_argument("Cannot convert return value to XLOPER12.");
		return pvRetVal;
	}
	catch (const std::exception &)
	{
		// todo: report exception
	}
	catch (...)
	{
		// todo: report exception
	}
	return const_cast<ExcelVariant*>(&ExcelVariant::ErrValue);
}

XLL_END_NAMESPACE

#define XLL_QUOTE_(x) #x
#define XLL_QUOTE(x) XLL_QUOTE_(x)
#define XLL_CONCAT_(x,y) x##y
#define XLL_CONCAT(x,y) XLL_CONCAT_(x,y)

#define EXPORT_XLL_FUNCTION(f) \
	template <typename Func> struct XLWrapper_##f; \
	template <typename TRet, typename... TArgs> \
	struct XLWrapper_##f < TRet(TArgs...) > \
	{ \
		static LPXLOPER12 __stdcall Call(typename ::XLL_NAMESPACE::ArgumentMarshaler<TArgs>::WireType... args) \
		{ \
			__pragma(comment(linker, "/export:" XLL_WRAPPER_PREFIX #f "=" __FUNCDNAME__)) \
			return ::XLL_NAMESPACE::XLWrapperImpl<decltype(f), f, TRet, TArgs...>(args...); \
		} \
		static const wchar_t * GetTypeText() \
		{ \
			auto &r = Call; \
			return ::XLL_NAMESPACE::TypeText<std::remove_reference<decltype(r)>::type>::SeqTypeW::ToArray(); \
		} \
	}; \
	static ::XLL_NAMESPACE::FunctionInfoBuilder XLFun_##f = ::XLL_NAMESPACE::AddFunction( \
		::XLL_NAMESPACE::FunctionInfo(XLWrapper_##f<::XLL_NAMESPACE::StripCallingConvention<decltype(f)>::type>::GetTypeText(), \
		XLL_CONCAT(L,XLL_WRAPPER_PREFIX) L#f)).Name(L#f)
