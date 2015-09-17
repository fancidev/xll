////////////////////////////////////////////////////////////////////////////
// Wrapper.h -- type-safe wrapper to expose user functions to Excel

#pragma once

#include "xlldef.h"
#include "FunctionInfo.h"
#include "Conversion.h"
#include "Marshal.h"

namespace XLL_NAMESPACE
{
	//
	// StripCallingConvention<Func>
	//
	// Removes the calling convention from a function type.
	//                

#ifndef _WIN64
#define XLL_HAVE_CDECL      1
#define XLL_HAVE_STDCALL    1
#define XLL_HAVE_FASTCALL   1
#define XLL_HAVE_VECTORCALL 0
#else
#define XLL_HAVE_CDECL      1
#define XLL_HAVE_STDCALL    0
#define XLL_HAVE_FASTCALL   0
#define XLL_HAVE_VECTORCALL 1
#endif

	template <typename Func> struct StripCallingConvention;

#if XLL_HAVE_CDECL
	template <typename TRet, typename... TArgs>
	struct StripCallingConvention < TRet __cdecl(TArgs...) >
	{
		typedef TRet type(TArgs...);
	};
#endif

#if XLL_HAVE_STDCALL
	template <typename TRet, typename... TArgs>
	struct StripCallingConvention < TRet __stdcall(TArgs...) >
	{
		typedef TRet type(TArgs...);
	};
#endif

#if XLL_HAVE_FASTCALL
	template <typename TRet, typename... TArgs>
	struct StripCallingConvention < TRet __fastcall(TArgs...) >
	{
		typedef TRet type(TArgs...);
	};
#endif

#if XLL_HAVE_VECTORCALL
	template <typename TRet, typename... TArgs>
	struct StripCallingConvention < TRet __vectorcall(TArgs...) >
	{
		typedef TRet type(TArgs...);
	};
#endif

	//template <typename Func>
	//using strip_cc_t = typename StripCallingConvention<Func>::type;


	//
	// getReturnValue
	//
	// Returns a pointer an XLOPER12 that holds the return value of a
	// wrapper function. Because the code that fills the return value
	// is guaranteed never to be called recursively, we allocate the
	// return value in thread-local storage (TLS) where supported. 
	//
	// TLS is properly supported starting from WIndows Vista. On
	// earlier platforms, we allocate the return value on the heap.
	//

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

	//
	// XLWrapper<...>
	//
	// Actual implementation of UDF wrappers. This is used together with
	// EntryPointHolder structs, which contains a single entry point
	// function that is exported with a pretty name.
	// 

	template <template <void *, typename, typename...> class EntryPointHolder,
		typename Func,
		Func *func,
		typename = typename StripCallingConvention<Func>::type>
	struct XLWrapper;

	template <template <void *, typename, typename...> class EntryPointHolder,
		typename Func,
		Func *func,
		typename TRet,
		typename... TArgs>
	struct XLWrapper < EntryPointHolder, Func, func, TRet(TArgs...) >
	{
		// Type of the entry point of the wrapped UDF.
		typedef LPXLOPER12(__stdcall EntryPointType)
			(typename ArgumentMarshaler<TArgs>::WireType...);

		// Actual implementation of the wrapper.
		static inline LPXLOPER12 __stdcall Call(
			typename ArgumentMarshaler<TArgs>::WireType... args)
		{
			try
			{
				LPXLOPER12 pvRetVal = xll::getReturnValue();
				HRESULT hr = xll::SetValue(pvRetVal,
					func(ArgumentMarshaler<TArgs>::Marshal(args)...));
				if (FAILED(hr))
				{
					throw std::invalid_argument(
						"Cannot convert return value to XLOPER12.");
				}
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

		// Returns the entry point address of the UDF wrapper. This function
		// must be called somewhere to emit the entry point into object file.
		static inline EntryPointType* GetEntryPoint()
		{
			return EntryPointHolder < Call, LPXLOPER12,
				typename ArgumentMarshaler<TArgs>::WireType... > ::EntryPoint;
		}

		// Register the wrapper in the XLL's list of UDFs.
		static inline FunctionInfoBuilder Register(LPCWSTR entryPoint, LPCWSTR name)
		{
			//Reference(GetEntryPoint());
			xll::FunctionInfo& info = GetFunctionInfo();
			info.entryPoint = entryPoint;
			info.name = name;
			return xll::FunctionInfoBuilder(info);
		}

	private:
		static inline FunctionInfo& GetFunctionInfo()
		{
			static FunctionInfo& s_info = FunctionInfo::Create(GetEntryPoint());
			return s_info;
		}
	};
}

//
// Macro to create and export a wrapper for UDF. Requirements:
//
// *) The macro must be placed in a source file.
// *) The macro must refer to a UDF accessible from that source file.
// *) The UDF must have external linkage.
// *) The macro may be put in any namespace.
//

#define XLL_CONCAT_(x,y) x##y
#define XLL_CONCAT(x,y) XLL_CONCAT_(x,y)
#define XLL_QUOTE_(x) #x
#define XLL_QUOTE(x) XLL_QUOTE_(x)
#define XLL_LQUOTE(x) XLL_CONCAT(L,XLL_QUOTE(x))

#if 0
#define EXPORT_XLL_FUNCTION(f) \
	template <void *func, typename WireRet, typename... WireArgs> \
	struct EntryPoint_##f \
	{ \
		static inline WireRet __stdcall EntryPoint(WireArgs... args) \
		{ \
			__pragma(comment(linker, "/export:" XLL_QUOTE(XLL_WRAPPER_PREFIX) XLL_QUOTE(f) "=" __FUNCDNAME__)) \
			return static_cast<WireRet(__stdcall*)(WireArgs...)>(func)(args...); \
		} \
	}; \
	static auto XLL_CONCAT(XLFIB_,f) = ::XLL_NAMESPACE::XLWrapper \
		< XLL_CONCAT(EntryPoint_,f), decltype(f), f > \
		::Register(XLL_LQUOTE(XLL_WRAPPER_PREFIX) XLL_LQUOTE(f), XLL_LQUOTE(f))
#else
#define EXPORT_XLL_FUNCTION(f) \
	template <void *func, typename WireRet, typename... WireArgs> \
	struct EntryPoint_##f \
	{ \
		static __declspec(dllexport) WireRet __stdcall EntryPoint(WireArgs... args) \
		{ \
			return static_cast<WireRet(__stdcall*)(WireArgs...)>(func)(args...); \
		} \
	}; \
	static auto XLL_CONCAT(XLFIB_,f) = ::XLL_NAMESPACE::XLWrapper \
		< XLL_CONCAT(EntryPoint_,f), decltype(f), f > \
		::Register(nullptr, XLL_LQUOTE(f))
#endif
