#pragma once

#include <Windows.h>
#include "XLCALL.H"
#include <string>
#include <vector>
#include <exception>
#include <cstdint>
#include <array>
#include <type_traits>

#define XS(s) L##s

// Excel defines a variant-like XLOPER type. We wrap this data type in
// ExcelVariant to simplify operations.
#pragma region Excel data types

class ExcelRef : public xlref12
{
};

void XLOPER12_Clear(LPXLOPER12 p);

template <typename T>
void XLOPER12_Create(LPXLOPER12 pv, const T &value)
{
	static_assert(false, "Don't know how to convert the specified type to XLOPER12. "
						 "Overload XLOPER12_Create() to fix this issue.");
}

inline void XLOPER12_Create(LPXLOPER12 pv, double value)
{
	pv->xltype = xltypeNum;
	pv->val.num = value;
}

inline void XLOPER12_Create(LPXLOPER12 pv, unsigned long value)
{
	XLOPER12_Create(pv, static_cast<double>(value));
}

//ExcelVariant(const char *s);

void XLOPER12_Create(LPXLOPER12 pv, const wchar_t *s, size_t len);

inline void XLOPER12_Create(LPXLOPER12 pv, const wchar_t *s)
{
	XLOPER12_Create(pv, s, s ? (size_t)lstrlenW(s) : 0);
}

inline void XLOPER12_Create(LPXLOPER12 pv, const std::wstring &value)
{
	XLOPER12_Create(pv, value.c_str(), value.size());
}

inline void XLOPER12_Create(LPXLOPER12 pv, bool value)
{
	pv->xltype = xltypeBool;
	pv->val.xbool = value;
}

inline void XLOPER12_Create(LPXLOPER12 pv, int value)
{
	pv->xltype = xltypeInt;
	pv->val.w = value;
}

// ExcelVariant() ref

// ExcelVariant() err

// ExcelVariant() flow

// ExcelVariant() array

// ExcelVariant() missing

// ExcelVariant() nil

//ExcelVariant(const ExcelRef &ref)
//{
//	xltype = xltypeSRef;
//	val.sref.count = 1;
//	val.sref.ref = ref;
//}



///////////////////////////////

// Wraps an XLOPER12 and automatically releases memory on destruction.
// Use this class when you pass arguments to an Excel function.
class ExcelVariant : public XLOPER12
{
	static ExcelVariant FromType(WORD xltype)
	{
		ExcelVariant v;
		v.xltype = xltype;
		return v;
	}

	static ExcelVariant MakeError(int err)
	{
		ExcelVariant v;
		v.xltype = xltypeErr;
		v.val.err = err;
		return v;
	}

	static void Copy(XLOPER12& to, const XLOPER12 &from);

public:
	static const ExcelVariant Empty;
	static const ExcelVariant Missing;
	static const ExcelVariant ErrValue;

	ExcelVariant()
	{
		xltype = 0;
	}

	explicit ExcelVariant(const XLOPER12 &other)
	{
		Copy(*this, other);
	}

	ExcelVariant(const ExcelVariant &other) = delete;

	ExcelVariant& operator=(const ExcelVariant &other) = delete;

	ExcelVariant& operator=(ExcelVariant &&other)
	{
		if (&other != this)
		{
			XLOPER12 tmp;
			memcpy(&tmp, &other, sizeof(XLOPER12));
			memcpy(&other, this, sizeof(XLOPER12));
			memcpy(this, &tmp, sizeof(XLOPER12));
		}
		return (*this);
	}

	ExcelVariant(ExcelVariant&& other)
	{
		if (&other != this)
		{
			memcpy(this, &other, sizeof(ExcelVariant));
			memset(&other, 0, sizeof(ExcelVariant));
		}
	}

	ExcelVariant(double value)
	{
		xltype = xltypeNum;
		val.num = value;
	}

	ExcelVariant(wchar_t c)
	{
		wchar_t *p = (wchar_t*)malloc(sizeof(wchar_t)*2);
		if (p == nullptr)
			throw std::bad_alloc();

		p[0] = 1;
		p[1] = c;
		xltype = xltypeStr;
		val.str = p;
	}

	//ExcelVariant(const char *s)
	//{
	//}

	ExcelVariant(const wchar_t *s)
	{
		if (s == nullptr)
		{
			xltype = xltypeMissing;
			return;
		}

		int len = lstrlenW(s);
		if (len < 0 || len > 65535)
			throw new std::invalid_argument("input string is too long");
		
		wchar_t *p = (wchar_t*)malloc(sizeof(wchar_t)*(len + 1));
		if (p == nullptr)
			throw std::bad_alloc();

		p[0] = (wchar_t)len;
		memcpy(&p[1], s, len*sizeof(wchar_t));

		xltype = xltypeStr;
		val.str = p;
	}

	ExcelVariant(const std::wstring &value)
		: ExcelVariant(value.c_str())
	{
	}
	
	ExcelVariant(bool value)
	{
		xltype = xltypeBool;
		val.xbool = value;
	}

	// ExcelVariant() ref

	// ExcelVariant() err

	// ExcelVariant() flow

	// ExcelVariant() array

	// ExcelVariant() missing

	// ExcelVariant() nil

	ExcelVariant(const ExcelRef &ref)
	{
		xltype = xltypeSRef;
		val.sref.count = 1;
		val.sref.ref = ref;
	}

	ExcelVariant(int value)
	{
		xltype = xltypeInt;
		val.w = value;
	}

	// ExcelVariant() xltypeBigData

	~ExcelVariant()
	{
		XLOPER12_Clear(this);
	}

	// Returns the content of this object in a heap-allocated XLOPER12 suitable
	// to be returned to Excel. The XLOPER12 has its xlbitDLLFree bit set. The
	// content of this object is cleared.
	LPXLOPER12 detach()
	{
		LPXLOPER12 p = (LPXLOPER12)malloc(sizeof(XLOPER12));
		if (p == nullptr)
			throw std::bad_alloc();
		memcpy(p, this, sizeof(XLOPER12));
		p->xltype |= xlbitDLLFree;
		xltype = 0;
		return p;
	}
};
#pragma endregion

#if 0
//ExcelVariant ExcelCall(int xlfn);
//ExcelVariant ExcelCall(int xlfn, const ExcelVariant &); // cannot pass XLOPER12 &

template <typename... T>
ExcelVariant ExcelCall(int xlfn, const T&... args)
{
	//va_list x;
	//va_start;
	// TODO: this function won't work or would invoke a copy constructor
	// if the supplied argument is of type XLOPER12. This is not what we want.
	const int NumArgs = sizeof...(args);
	ExcelVariant xlArgs[NumArgs] = { args... };
	LPXLOPER12 pArgs[NumArgs];
	for (int i = 0; i < NumArgs; i++)
		pArgs[i] = &xlArgs[i];

	XLOPER12 result;
	int ret = Excel12v(xlfn, &result, numargs, pArgs);
	if (ret != xlretSuccess)
		throw ExcelException(ret);

	ExcelVariant vResult(result);
	Excel12(xlFree, 0, 1, &result);
	return vResult;
}
#endif

// When Excel calls a UDF, it supports passing arguments of type LPXLOPER12
// as well as several other native types. We call these "wrapped arguments".
// We must "unwrap" these incoming arguments before forwarding the call to
// the UDF. Likewise, we must "wrap" the return value of the udf to one of
// the several return types supported by Excel.
#pragma region Argument marshalling

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
DEFINE_SIMPLE_ARGUMENT_WRAPPER(std::wstring, const wchar_t *);

template <> struct ArgumentWrapper<VARIANT>
{
	typedef LPXLOPER12 wrapped_type;
	static VARIANT unwrap(LPXLOPER12 v);
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
	SafeArrayWrapper(const XLOPER12 *pv);
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
	static SafeArrayWrapper unwrap(LPXLOPER12 v);
};

template <typename T> struct ArgumentWrapper<T &> : ArgumentWrapper < T > {};
template <typename T> struct ArgumentWrapper<T &&> : ArgumentWrapper < T > {};
template <typename T> struct ArgumentWrapper<T const> : ArgumentWrapper < T > {};
template <typename T> struct ArgumentWrapper<T volatile> : ArgumentWrapper < T > {};
template <typename T> struct ArgumentWrapper<T const volatile> : ArgumentWrapper < T > {};

#pragma endregion

#define EXPORT_UNDECORATED_NAME comment(linker, "/export:" __FUNCTION__ "=" __FUNCDNAME__)

class NameDescriptionPair
{
	LPCWSTR m_name;
	LPCWSTR m_description;
public:
	NameDescriptionPair(LPCWSTR name, LPCWSTR description)
		: m_name(name), m_description(description)
	{
	}
	LPCWSTR name() const { return m_name; }
	LPCWSTR description() const { return m_description; }
};

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

DEFINE_TYPE_TEXT(bool,       "A");
DEFINE_TYPE_TEXT(bool*,      "L");
DEFINE_TYPE_TEXT(double,     "B");
DEFINE_TYPE_TEXT(double*,    "E");
DEFINE_TYPE_TEXT(char*,      "C");
DEFINE_TYPE_TEXT(uint16_t,   "H");
DEFINE_TYPE_TEXT(int16_t,    "I");
DEFINE_TYPE_TEXT(int16_t*,   "M");
DEFINE_TYPE_TEXT(int32_t,    "J");
DEFINE_TYPE_TEXT(int32_t*,   "N");
DEFINE_TYPE_TEXT(wchar_t*,   "C%");
DEFINE_TYPE_TEXT(LPCWSTR,    "C%");
DEFINE_TYPE_TEXT(LPXLOPER12, "Q");

struct FunctionInfo
{
	LPCWSTR entryPoint;
	std::wstring typeText;

	LPCWSTR name;
	LPCWSTR description;
	std::vector<NameDescriptionPair> arguments;
	int macroType; // 0,1,2
	LPCWSTR category;
	LPCWSTR shortcut;
	LPCWSTR helpTopic;

	bool isPure;
	bool isThreadSafe;

	FunctionInfo(LPCWSTR typeText, LPCWSTR entryPoint)
		: typeText(typeText), entryPoint(entryPoint), name(), description(), 
		  macroType(1), category(), shortcut(), helpTopic(), isPure(), isThreadSafe()
	{
	}
};

class FunctionInfoBuilder
{
	FunctionInfo &_info;

public:
	FunctionInfoBuilder(FunctionInfo &functionInfo) 
		: _info(functionInfo)
	{
	}

	FunctionInfoBuilder& Name(LPCWSTR name)
	{
		_info.name = name;
		return (*this);
	}

	FunctionInfoBuilder& Description(LPCWSTR description)
	{
		_info.description = description;
		return (*this);
	}

	FunctionInfoBuilder& Arg(LPCWSTR name, LPCWSTR description)
	{
		_info.arguments.push_back(NameDescriptionPair(name, description));
		return (*this);
	}

	FunctionInfoBuilder& Category(LPCWSTR category)
	{
		_info.category = category;
		return (*this);
	}

	FunctionInfoBuilder& HelpTopic(LPCWSTR helpTopic)
	{
		_info.helpTopic = helpTopic;
		return (*this);
	}

	FunctionInfoBuilder& Pure()
	{
		_info.isPure = true;
		return (*this);
	}

	FunctionInfoBuilder& Volatile()
	{
		_info.isPure = false;
		return (*this);
	}

	FunctionInfoBuilder& ThreadSafe()
	{
		_info.isThreadSafe = true;
		return (*this);
	}
};

template <typename Func> class FunctionInfoFactory;

template <typename TRet, typename... TArgs>
struct FunctionInfoFactory<TRet(TArgs...)>
{
	static FunctionInfo Create(LPCWSTR entryPoint)
	{
		const int NumArgs = sizeof...(TArgs);
		std::array<LPCWSTR, NumArgs + 1> texts = {
			L"Q", // LPXLOPER12
			GetTypeText<typename ArgumentWrapper<TArgs>::wrapped_type>()...
		};
		std::wstring s;
		for (int i = 0; i <= NumArgs; i++)
		{
			s += texts[i];
		}
		return FunctionInfo(s.c_str(), entryPoint);
	}
};

template <typename TRet, typename... TArgs>
struct FunctionInfoFactory<TRet __stdcall(TArgs...)> 
	: public FunctionInfoFactory<TRet(TArgs...)>
{
};

class ExcelException : public std::exception
{
	int m_errorCode;
	char m_errorMessage[100];

	static const char* GetKnownErrorMessage(int errorCode)
	{
		switch (errorCode)
		{
		case xlretSuccess:
			return "success";
		case xlretAbort:
			return "macro halted";
		case xlretInvXlfn:
			return "invalid function number";
		case xlretInvCount:
			return "invalid number of arguments";
		case xlretInvXloper:
			return "invalid OPER structure";
		case xlretStackOvfl:
			return "stack overflow";
		case xlretFailed:
			return "command failed";
		case xlretUncalced:
			return "uncalced cell";
		case xlretNotThreadSafe:
			return "not allowed during multi-threaded calc";
		case xlretInvAsynchronousContext:
			return "invalid asynchronous function handle";
		case xlretNotClusterSafe:
			return "not supported on cluster";
		default:
			return nullptr;
		}
	}

public:
	explicit ExcelException(int errorCode)
		: m_errorCode(errorCode)
	{
		if (GetKnownErrorMessage(errorCode) == nullptr)
			sprintf_s(m_errorMessage, "xll error %d", m_errorCode);
	}

	const char* what() const override
	{
		const char *msg = GetKnownErrorMessage(m_errorCode);
		return msg? msg : m_errorMessage;
	}
};

class AddinRegistrar
{
public:
	static std::vector<FunctionInfo> & registry()
	{
		static std::vector<FunctionInfo> s_functions;
		return s_functions;
	}

	static FunctionInfoBuilder AddFunction(FunctionInfo &f)
	{
		registry().push_back(f);
		return FunctionInfoBuilder(registry().back());
	}
};

#ifndef XLL_SUPPORT_THREAD_LOCAL
#if WINVER >= _WIN32_WINNT_VISTA
#define XLL_SUPPORT_THREAD_LOCAL 1
#else
#define XLL_SUPPORT_THREAD_LOCAL 0
#endif
#endif

#if XLL_SUPPORT_THREAD_LOCAL
__declspec(thread) extern XLOPER12 xllReturnValue;
inline LPXLOPER12 getReturnValue() { return &xllReturnValue; }
#else
inline LPXLOPER12 getReturnValue() {
	LPXLOPER12 p = (LPXLOPER12)malloc(sizeof(XLOPER12));
	if (p == nullptr)
		throw std::bad_alloc();
	return p;
}
#endif

//extern const XLOPER12 xllErrorValue;

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
			XLOPER12_Create(pvRetVal, func(ArgumentWrapper<TArgs>::unwrap(args)...));
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

#define EXPORT_DLL_FUNCTION(name, implementation) \
	extern "C" __declspec(naked, dllexport) void name() \
	{ \
		static const void *fp = static_cast<const void *>(implementation); \
		__asm { jmp [fp] } \
	}

#define WRAPPER_TYPE(f) std::remove_pointer<decltype(XLWrapper<decltype(f), f, decltype(f)>::Call)>::type

#define EXPORT_XLL_FUNCTION(f) \
	EXPORT_DLL_FUNCTION(XL##f, (XLWrapper<decltype(f), f>::Call)) \
	static FunctionInfoBuilder XLFun_##f = AddinRegistrar::AddFunction(\
		FunctionInfoFactory<decltype(f)>::Create(L"XL" L#f)).Name(L#f)

//
//extern "C" __declspec(naked, dllexport) void XLSquare()
//{
//	static const void *f = static_cast<const void *>(
//		XLWrapper<decltype(Square), Square, decltype(Square)>::Call);
//	__asm {
//		jmp [f]
//	}
//}

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
