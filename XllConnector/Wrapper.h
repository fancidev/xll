////////////////////////////////////////////////////////////////////////////
// Wrapper.h -- type-safe wrapper to expose user-defined functions to Excel

#pragma once

#include "xlldef.h"
#include "FunctionInfo.h"
#include "Conversion.h"
#include "Marshal.h"

//
// strip_cc, strip_cc_t
//
// Removes explicit calling convention from a function type.
//
// strip_cc<Func>::type removes any explicit calling convention from
// the function type Func, producing a function type with the default
// calling convention. The default calling convention can be altered
// by compiler switches /Gd, /Gr, /Gv, or /Gz.
//
// strip_cc_t<Func> is shorthand for strip_cc<Func>::type.
//

namespace XLL_NAMESPACE
{
	template <typename Func, typename = void> struct strip_cc;

	template <int> struct Placeholder;

	template <typename TRet, typename... TArgs>
	struct strip_cc < TRet(TArgs...), void >
	{
		typedef TRet type(TArgs...);
	};

#define XLL_DEFINE_STRIP_CC(n, cc) \
	template <typename TRet, typename... TArgs> \
	struct strip_cc <TRet cc(TArgs...), std::conditional_t< \
		std::is_same< TRet cc(TArgs...), TRet(TArgs...)>::value, \
		Placeholder<n>, void > > \
		{ \
		typedef TRet type(TArgs...); \
		}

	XLL_DEFINE_STRIP_CC(0, __cdecl);
	XLL_DEFINE_STRIP_CC(1, __stdcall);
	XLL_DEFINE_STRIP_CC(2, __fastcall);
	XLL_DEFINE_STRIP_CC(3, __vectorcall);

	template <typename Func>
	using strip_cc_t = typename strip_cc<Func>::type;
}

#if XLL_SUPPORT_THREAD_LOCAL
namespace XLL_NAMESPACE
{
	__declspec(thread) extern XLOPER12 xllReturnValue;
}
#endif

//
// XLWrapper
//
// Template-based, type-safe UDF wrapper that
//   1) handles argument and return value marshalling, and
//   2) provides a __stdcall entry point for Excel to call.
//
// By default, the entry point is exported by its decorated name,
// which contains the name of the underlying function as well as
// the name of its argument types. If this is a concern to you, 
// wrap your function in another function that doesn't contain
// sensitive names.
//
// NOTE: In 32-bit build, we may create and export a naked function
// that contains a single jmp instruction to jump to the true entry
// point. In 64-bit build, this is not supported by Visual C++.
// 

namespace XLL_NAMESPACE
{
	// TODO: Find some way to have one less template parameter (to make
	// export table prettier. Might need to use tuples.
	template <typename Func, Func *func, int Attributes = 0, 
		      typename = strip_cc_t<Func> >
	struct XLWrapper;

	template <typename Func, Func *func, int Attributes,
		      typename TRet, typename... TArgs>
	struct XLWrapper < Func, func, Attributes, TRet(TArgs...) > 
		: FunctionAttributes<Attributes>
	{
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
		// If the UDF is not threadsafe, then Excel will never call it
		// from more than one thread at any time. In this case, we
		// allocate the return value statically to save a few instructions.
		//
		// Note that we only do this if TLS is supported, because otherwise
		// there is no way for xlAutoFree to know whether it should free
		// the LPXLOPER12 pointer.
		//

#if XLL_SUPPORT_THREAD_LOCAL
		template <bool Pred> 
		static inline LPXLOPER12 getReturnValue(
			std::enable_if_t<Pred, int> = 0)
		{
			return &xllReturnValue;
		}

		template <bool Pred>
		static inline LPXLOPER12 getReturnValue(
			std::enable_if_t<!Pred, int> = 0)
		{
			static XLOPER12 xllReturnValue;
			return &xllReturnValue;
		}
#else
		template <bool Pred>
		static inline LPXLOPER12 getReturnValue()
		{
			LPXLOPER12 p = (LPXLOPER12)malloc(sizeof(XLOPER12));
			if (p == nullptr)
				throw std::bad_alloc();
			return p;
		}
#endif

		//
		// EntryPoint
		//
		// Actual entry point called by Excel.
		//

		static __declspec(dllexport) LPXLOPER12 __stdcall 
		EntryPoint(typename ArgumentMarshaler<TArgs>::WireType... args) 
		XLL_NOEXCEPT
		{
			try
			{
				LPXLOPER12 pvRetVal = getReturnValue<IsThreadSafe>();
				HRESULT hr = SetValue(pvRetVal,
					func(ArgumentMarshaler<TArgs>::MarshalIn(args)...));
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

		static inline FunctionInfo& GetFunctionInfo()
		{
			static FunctionInfo& s_info = FunctionInfo::Create<Attributes>(EntryPoint);
			return s_info;
		}
	};

	template <typename Func, Func *func, int Attributes>
	struct XLWrapper < Func, func, Attributes, void() >
		: FunctionAttributes<Attributes>
	{
		template <typename T> struct always_false : std::false_type {};
		static_assert(always_false<Func>::value,
			"A void-returning function must take its at least "
			"one modified-in-place argument.");
	};

	template <typename Func, Func *func, int Attributes, typename TArg1, typename... TArgs>
	struct XLWrapper < Func, func, Attributes, void(TArg1, TArgs...) >
		: FunctionAttributes<Attributes>
	{
		//
		// EntryPoint
		//
		// Actual entry point called by Excel.
		//

		static __declspec(dllexport) void __stdcall
		EntryPoint(typename ArgumentMarshaler<TArg1>::WireType arg1,
				   typename ArgumentMarshaler<TArgs>::WireType... args)
		XLL_NOEXCEPT
		{
			try
			{
				func(ArgumentMarshaler<TArg1>::MarshalInOut(arg1),
					 ArgumentMarshaler<TArgs>::MarshalIn(args)...);
			}
			catch (const std::exception &)
			{
				// todo: report exception
			}
			catch (...)
			{
				// todo: report exception
			}
		}

		static inline FunctionInfo& GetFunctionInfo()
		{
			static FunctionInfo& s_info = FunctionInfo::Create<Attributes>(EntryPoint);
			return s_info;
		}
	};
}

//
// XLLocalWrapper
//
// Helper class to instantiate an XLWrapper specialization.
//
// This class is used by the EXPORT_XLL_FUNCTION() macro to instantiate
// a specialization of xll::XLWrapper. It employs a few tricks:
//
//  i) It is defined as a template with a static member variable so that
//     the macro doesn't have to make up a static variable name;
//  2) It is enclosed in an anonymous namespace so that the macro doesn't
//     have to qualify the namespace; and
//  3) It normalizes the Attributes template argument passed to XLWrapper
//     so that there is less code to write in the macro.
//

namespace
{
	template <typename Func, Func *func, int Attributes = 0>
	struct XLLocalWrapper
	{
		static ::XLL_NAMESPACE::FunctionInfoBuilder functionInfoBuilder;

		static inline ::XLL_NAMESPACE::FunctionInfoBuilder 
			BuildFunctionInfo(LPCWSTR name)
		{
			return ::XLL_NAMESPACE::FunctionInfoBuilder(
				::XLL_NAMESPACE::XLWrapper<Func, func, 
				::XLL_NAMESPACE::NormalizeAttributes<Attributes>::value>
				::GetFunctionInfo()).Name(name);
		}
	};
}

//
// EXPORT_XLL_FUNCTION, EXPORT_XLL_FUNCTION_AS
//
// Macro to create and export a wrapper for a given UDF.
//
// Requirements:
//
//   *) The macro must be placed in a source file.
//   *) The macro must refer to a declared UDF.
//   *) The UDF must have external linkage.
//   *) The macro may be put in any namespace.
//

#define EXPORT_XLL_FUNCTION_AS(f, name, ...) \
	::XLL_NAMESPACE::FunctionInfoBuilder \
		XLLocalWrapper<decltype(f), f, __VA_ARGS__>::functionInfoBuilder = \
		XLLocalWrapper<decltype(f), f, __VA_ARGS__>::BuildFunctionInfo(L##name)

#define EXPORT_XLL_FUNCTION(f,...) EXPORT_XLL_FUNCTION_AS(f, #f, __VA_ARGS__)
