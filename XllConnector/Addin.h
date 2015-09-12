////////////////////////////////////////////////////////////////////////////
// Addin.h -- defines macros to export functions to Excel

#pragma once

#include "xlldef.h"
#include "FunctionInfo.h"
#include "Conversion.h"
#include <cstdint>

XLL_BEGIN_NAMEPSACE

FunctionInfoBuilder AddFunction(FunctionInfo &f);

#define EXPORT_UNDECORATED_NAME comment(linker, "/export:" __FUNCTION__ "=" __FUNCDNAME__)


#if 0
template <typename T> struct fake_dependency : public std::false_type {};
template <typename T> struct ArgumentTypeText
{
	static_assert(fake_dependency<T>::value, "Does not support marshalling of the supplied type.");
};

#define DECLARE_ARGUMENT_TYPE_TEXT(type, text) \
template <> struct ArgumentTypeText < type > { \
	static const wchar_t * getTypeText() { return L##text; } \
	};

#define GET_ARGUMENT_TYPE_TEXT(type) ArgumentTypeText<type>::getTypeText()
#else
template <typename T> inline const wchar_t * GetTypeText()
{
	static_assert(false, "The supplied type is not a valid XLL argument type.");
}
#define DEFINE_TYPE_TEXT(type, text) \
template<> inline const wchar_t * GetTypeText<type>() { return L##text; }
#endif

DEFINE_TYPE_TEXT(bool, "A");
DEFINE_TYPE_TEXT(bool*, "L");
DEFINE_TYPE_TEXT(double, "B");
DEFINE_TYPE_TEXT(double*, "E");
DEFINE_TYPE_TEXT(char*, "C");
DEFINE_TYPE_TEXT(const char*, "C");
DEFINE_TYPE_TEXT(uint16_t, "H");
DEFINE_TYPE_TEXT(int16_t, "I");
DEFINE_TYPE_TEXT(int16_t*, "M");
DEFINE_TYPE_TEXT(int32_t, "J");
DEFINE_TYPE_TEXT(int32_t*, "N");
DEFINE_TYPE_TEXT(wchar_t*, "C%");
DEFINE_TYPE_TEXT(const wchar_t*, "C%");
DEFINE_TYPE_TEXT(LPXLOPER12, "Q");

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

#if !defined(_WIN64)
// On 32-bit platform, we use naked function to emit a JMP instruction,
// and export this function. The function must be __cdecl.

template <typename Func, Func *func, typename = typename StripCallingConvention<Func>::type>
struct XLWrapper;

template <typename Func, Func *func, typename TRet, typename... TArgs>
struct XLWrapper < Func, func, TRet(TArgs...) >
{
	static LPXLOPER12 __stdcall Call(typename ArgumentWrapper<TArgs>::wrapped_type... args)
	{
		try
		{
			LPXLOPER12 pvRetVal = getReturnValue();
			*pvRetVal = make<XLOPER12>(func(ArgumentWrapper<TArgs>::unwrap(args)...));
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
};

XLL_END_NAMESPACE

#define EXPORT_DLL_FUNCTION(name, implementation) \
	extern "C" __declspec(naked, dllexport) void name() \
		{ \
		static const void *fp = static_cast<const void *>(implementation); \
		__asm { jmp [fp] } \
		}

#define EXPORT_XLL_FUNCTION(f) \
	EXPORT_DLL_FUNCTION(XL##f, (::XLL_NAMESPACE::XLWrapper<decltype(f), f>::Call)) \
	static ::XLL_NAMESPACE::FunctionInfoBuilder XLFun_##f = ::XLL_NAMESPACE::AddFunction(\
		::XLL_NAMESPACE::FunctionInfoFactory<decltype(f)>::Create(L"XL" L#f)).Name(L#f)

#else
	// Code for WIN64. Because naked function and inline assembly is not 
	// supported, we have to define a macro to export the proper symbol.
#define EXPOSE_FUNCTION_RENAME(Function, Name) \
	template <typename> struct XL##Name; \
	template <typename TRet, typename... TArgs> \
	struct XL##Name < TRet(TArgs...) > \
			{ \
		static __declspec(dllexport) typename Boxed<TRet>::type __stdcall \
			Call(typename Boxed<TArgs>::type... args) \
				{ \
			__pragma(comment(linker, "/export:XL" #Name "=" __FUNCDNAME__)) \
			return Function(args...); \
				} \
			}; \
	static auto XL##Name##Temp = XL##Name<decltype(Function)>::Call;

#define EXPOSE_FUNCTION(Function) EXPOSE_FUNCTION_RENAME(Function, Function)
#endif
