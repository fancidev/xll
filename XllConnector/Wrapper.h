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

// TODO: the following two functions should be moved to a separate
// header file, probably merge with Invoke.

//
// AddInName(name)
//
// Gets or sets the name of this XLL add-in.
//
// If name is NULL, returns the add-in name. If name is not NULL,
// sets the add-in name and returns name. name must be a string
// constant. 
// 
// To set the add-in name, this function must be called at static
// initialization time, because Excel queries the add-in name as
// soon as it loads the DLL. 
//
// Use the XLL_ADDIN_NAME() macro to set add-in name.
//

namespace XLL_NAMESPACE
{
	LPCWSTR AddInName(LPCWSTR name = NULL);
}

#define XLL_ADDIN_NAME(name) \
	static LPCWSTR _XllConnector_AddInName = ::XLL_NAMESPACE::AddInName(name)

//
// AllocateReturnValue
//
// Returns a pointer an XLOPER12 used to hold the return value of a
// wrapper function.
//
// When wrapping a thread-unsafe UDF, we return a pointer to a global
// variable. This is safe because Excel always calls thread-unsafe
// UDFs from the main thread, and XLL Connector does not call any
// wrapper code (directly or indirectly) after filling the return
// value and before returning it to Excel.
//
// When wrapping a thread-safe UDF, we return a pointer to a thread-
// local variable if supported, or allocate on the heap otherwise.
// TLS is properly supported starting from Windows Vista. For more
// info on TLS support, see http://www.nynaeve.net/?p=190.
//
// The global/thread-local variables are defined in Addin.cpp. They
// are freed in xlAutoFree12().
// 

namespace XLL_NAMESPACE
{
	inline LPXLOPER12 AllocateReturnValue(bool isThreadSafe)
	{
		if (isThreadSafe)
		{
#if XLL_SUPPORT_THREAD_LOCAL
			__declspec(thread) extern XLOPER12 threadReturnValue;
			return &threadReturnValue;
#else
			// TODO: set xlbitDLLFree to free the variable
			LPXLOPER12 p = (LPXLOPER12)malloc(sizeof(XLOPER12));
			if (p == nullptr)
				throw std::bad_alloc();
			return p;
#endif
		}
		else
		{
			extern XLOPER12 globalReturnValue;
			return &globalReturnValue;
		}
	}
}

//
// XLWrapper
//
// Template-based, type-safe UDF wrapper that
//   1) handles argument and return value marshalling, and
//   2) provides a __stdcall entry point for Excel to call.
//
// If XLL_GENERATE_STUB is not defined or set to zero, XLL Connector
// exports the entry point by its decorated name, which contains the
// name of the underlying function as well as its argument types. If
// this is not desired, wrap your function in another function that
// doesn't contain sensitive names, or define XLL_GENERATE_STUB to 1.
//
// When XLL_GENERATE_STUB is defined to non-zero, XLL Connector
// creates a stub for each wrapper function. This stub contains a
// single JMP instruction to jump to the actual start of the wrapper
// function. This makes it easier to identify the wrapper function
// in the DLL's export table when debugging.
//
// The XLL_GENERATE_STUB macro can be defined differently in each
// translation unit.
//
// Implementation Note: Because Visual C++ does not support inline
// assembly in 64-bit mode, the stub is directly emitted as machine
// code. 
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
		// EntryPoint
		//
		// Actual entry point called by Excel.
		//

#if !XLL_GENERATE_WRAPPER_STUB
		__declspec(dllexport)
#endif
		static LPXLOPER12 __stdcall 
		EntryPoint(typename ArgumentMarshaler<TArgs>::WireType... args) 
		XLL_NOEXCEPT
		{
			try
			{
				LPXLOPER12 pvRetVal = AllocateReturnValue(IsThreadSafe);
				HRESULT hr = CreateValue(pvRetVal,
					func(ArgumentMarshaler<TArgs>::Marshal(args)...));
				if (FAILED(hr))
				{
					throw std::invalid_argument(
						"Cannot convert return value to XLOPER12.");
				}
				// TODO: delete malloc-ed return value on return
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
			return const_cast<LPXLOPER12>(&Constants::ErrValue);
		}

		static inline FunctionInfo& GetFunctionInfo(FARPROC stub = 0)
		{
			static FunctionInfo& s_info = 
				FunctionInfo::Create<Attributes>(EntryPoint, stub);
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
			BuildFunctionInfo(LPCWSTR name, FARPROC stub = 0)
		{
			return ::XLL_NAMESPACE::FunctionInfoBuilder(
				::XLL_NAMESPACE::XLWrapper<Func, func, 
				::XLL_NAMESPACE::NormalizeAttributes<Attributes>::value>
				::GetFunctionInfo(stub)).Name(name);
		}
	};
}

#define XLL_CONCAT_(x,y) x##y
#define XLL_CONCAT(x,y) XLL_CONCAT_(x,y)

#define XLL_QUOTE_(x) #x
#define XLL_QUOTE(x) XLL_QUOTE_(x)

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

#if XLL_GENERATE_WRAPPER_STUB

#define XLL_STUB_NAME(name) XLL_CONCAT(XLL_WRAPPER_STUB_PREFIX,name)

#if 0
#define EXPORT_XLL_FUNCTION_AS(f, name, ...) \
	extern "C" __declspec(dllexport, naked) void XLL_STUB_NAME(name)() \
	{ \
		static const auto proc = ::XLL_NAMESPACE::XLWrapper < decltype(f), f, \
			::XLL_NAMESPACE::NormalizeAttributes<__VA_ARGS__>::value > ::EntryPoint; \
		__asm { jmp [proc] } \
	} \
	::XLL_NAMESPACE::FunctionInfoBuilder \
		XLLocalWrapper<decltype(f), f, __VA_ARGS__>::functionInfoBuilder = \
		XLLocalWrapper<decltype(f), f, __VA_ARGS__>::BuildFunctionInfo( \
		XLL_CONCAT(L,XLL_QUOTE(name)), (FARPROC)(XLL_STUB_NAME(name)))
#endif

#pragma pack(push, 1)
struct JmpInstruction
{
	unsigned char opcode[2]; // FF 25: Jump near, absolute indirect
	void *ptr;
};
#pragma pack(pop)

#pragma code_seg(push, ".xllstub")
#pragma code_seg(pop)

#define XLL_WRAPPER(f, ...) \
	::XLL_NAMESPACE::XLWrapper <decltype(f), f, \
	::XLL_NAMESPACE::NormalizeAttributes<__VA_ARGS__>::value >

#define EXPORT_XLL_FUNCTION_AS(f, name, ...) \
	static auto XLL_CONCAT(XLL_STUB_NAME(name),_EntryPoint) = \
		XLL_WRAPPER(f, __VA_ARGS__)::EntryPoint; \
	extern "C" __declspec(dllexport, allocate(".xllstub")) \
	JmpInstruction XLL_STUB_NAME(name) = { { 0xFF, 0x25 }, \
		& XLL_CONCAT(XLL_STUB_NAME(name),_EntryPoint) }; \
	::XLL_NAMESPACE::FunctionInfoBuilder \
		XLLocalWrapper<decltype(f), f, __VA_ARGS__>::functionInfoBuilder = \
		XLLocalWrapper<decltype(f), f, __VA_ARGS__>::BuildFunctionInfo( \
		XLL_CONCAT(L,XLL_QUOTE(name)), (FARPROC)&XLL_STUB_NAME(name))

#else

#define EXPORT_XLL_FUNCTION_AS(f, name, ...) \
	::XLL_NAMESPACE::FunctionInfoBuilder \
		XLLocalWrapper<decltype(f), f, __VA_ARGS__>::functionInfoBuilder = \
		XLLocalWrapper<decltype(f), f, __VA_ARGS__>::BuildFunctionInfo( \
		XLL_CONCAT(L,XLL_QUOTE(name)))

#endif


#define EXPORT_XLL_FUNCTION(f,...) EXPORT_XLL_FUNCTION_AS(f, f, __VA_ARGS__)
