#pragma once

#include <Windows.h>
#include "XLCALL.H"
#include <string>
#include <vector>
#include <exception>
#include <cstdint>
#include <array>
#include <type_traits>

// Excel defines a variant-like XLOPER type. We wrap this data type in
// ExcelVariant to simplify operations.
#pragma region Excel data types

class ExcelRef : public xlref12
{
};

class ExcelVariant : public XLOPER12
{
	void Reset()
	{
		memset(this, 0, sizeof(*this));
	}

	static ExcelVariant FromType(WORD xltype)
	{
		ExcelVariant v;
		v.xltype = xltype;
		return v;
	}

public:

	static const ExcelVariant Empty;
	static const ExcelVariant Missing;

	ExcelVariant()
	{
		Reset();
	}

	ExcelVariant(const ExcelVariant &other) = delete;

	ExcelVariant& operator=(const ExcelVariant &other) = delete;

	ExcelVariant& operator=(ExcelVariant &&other)
	{
		XLOPER12 tmp;
		memcpy(&tmp, &other, sizeof(XLOPER12));
		memcpy(&other, this, sizeof(XLOPER12));
		memcpy(this, &tmp, sizeof(XLOPER12));
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
		switch (xltype)
		{
		case xltypeStr:
			free(val.str);
			break;
		case xltypeRef:
			free(val.mref.lpmref);
			break;
		case xltypeMulti:
			free(val.array.lparray);
			break;
		}
		Reset();
	}
};
#pragma endregion

// When Excel calls a UDF, it supports passing certain types of argument
// directly. We must "unbox" these incoming arguments to the native type.
// Likewise, we must "box" the native return value of the function to a
// data type that Excel supports.
#pragma region Argument and return value marshalling

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

template <typename T> struct ArgumentWrapper<T &> : ArgumentWrapper < T > {};
template <typename T> struct ArgumentWrapper<T &&> : ArgumentWrapper < T > {};
template <typename T> struct ArgumentWrapper<T const> : ArgumentWrapper < T > {};
template <typename T> struct ArgumentWrapper<T volatile> : ArgumentWrapper < T > {};
template <typename T> struct ArgumentWrapper<T const volatile> : ArgumentWrapper < T > {};

template <typename T>
struct ReturnValueWrapper
{
	static_assert(fake_dependency<T>::value,
		"Do not know how to wrap the supplied return value type. "
		"Specialize xll::ReturnValueWrapper<T> to support it.");
	// typedef ?? wrapped_type;
	// static inline wrapped_type wrap(const T &);
};

#define DEFINE_SIMPLE_RETURN_VALUE_WRAPPER(NativeType) \
	template <> struct ReturnValueWrapper<NativeType> \
	{ \
		typedef NativeType wrapped_type; \
		static inline wrapped_type wrap(const NativeType &v) { return v; } \
	};

DEFINE_SIMPLE_RETURN_VALUE_WRAPPER(int);
DEFINE_SIMPLE_RETURN_VALUE_WRAPPER(double);

template <> struct ReturnValueWrapper < std::wstring >
{
	typedef LPXLOPER12 wrapped_type;
	static inline LPXLOPER12 wrap(const std::wstring &s)
	{
		LPXLOPER12 op = (LPXLOPER12)malloc(sizeof(XLOPER12));
		if (op == NULL)
			throw std::bad_alloc();

		LPWSTR buffer = (LPWSTR)malloc(sizeof(wchar_t)*(s.size() + 1));
		if (buffer == NULL)
			throw std::bad_alloc();

		buffer[0] = (wchar_t)s.size();
		memcpy(&buffer[1], s.c_str(), sizeof(wchar_t)*s.size());

		op->xltype = xltypeStr;
		op->val.str = buffer;
		return op;
		// todo: set the xlDLLFree bit
	}
};

template <typename T> struct ReturnValueWrapper<T const> : ReturnValueWrapper < T >{};
template <typename T> struct ReturnValueWrapper<T volatile> : ReturnValueWrapper < T >{};
template <typename T> struct ReturnValueWrapper<T const volatile> : ReturnValueWrapper < T >{};

#pragma endregion

#define EXPORT_UNDECORATED_NAME comment(linker, "/export:" __FUNCTION__ "=" __FUNCDNAME__)

enum class FunctionAttributes
{
	Default = 0,
	Pure = 1,
	ThreadSafe = 2,
};

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

	FunctionInfo(LPCWSTR typeText, LPCWSTR entryPoint)
		: typeText(typeText), entryPoint(entryPoint), name(), description(), 
		  macroType(1), category(), shortcut(), helpTopic()
	{
	}

	FunctionInfo& Name(LPCWSTR name)
	{
		this->name = name;
		return (*this);
	}

	FunctionInfo& Description(LPCWSTR description)
	{
		this->description = description;
		return (*this);
	}

	FunctionInfo& Arg(LPCWSTR name, LPCWSTR description)
	{
		arguments.push_back(NameDescriptionPair(name, description));
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
			GetTypeText<typename ReturnValueWrapper<TRet>::wrapped_type>(),
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

	virtual const char* what() const override
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

	static FunctionInfo& AddFunction(FunctionInfo &f)
	{
		registry().push_back(f);
		return registry().back();
	}
};

#if !defined(_WIN64)
// On 32-bit platform, we use naked function to emit a JMP instruction,
// and export this function. The function must be __cdecl.

template <typename Func, Func *func, typename> struct XLWrapper;
template <typename Func, Func *func, typename TRet, typename... TArgs>
struct XLWrapper < Func, func, TRet(TArgs...) >
{ 
	static typename ReturnValueWrapper<TRet>::wrapped_type __stdcall
		Call(typename ArgumentWrapper<TArgs>::wrapped_type... args)
	{
		return ReturnValueWrapper<TRet>::wrap(
			func(ArgumentWrapper<TArgs>::unwrap(args)...));
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
	EXPORT_DLL_FUNCTION(XL##f, (XLWrapper<decltype(f), f, decltype(f)>::Call)) \
	static FunctionInfo XLFun_##f = AddinRegistrar::AddFunction(\
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
