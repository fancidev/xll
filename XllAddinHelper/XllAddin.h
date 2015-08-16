#pragma once

#include <Windows.h>
#include "XLCALL.H"
#include <string>
#include <vector>
#include <exception>
#include <cstdint>
#include <array>
#include <type_traits>

class ExcelRef : public xlref12
{
};

class ExcelVariant : public XLOPER12
{
	void Reset()
	{
		memset(this, 0, sizeof(*this));
	}

public:
	ExcelVariant()
	{
		Reset();
	}

	ExcelVariant(const ExcelVariant &other) = delete;
	//{
	//}

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
			//free(val.str);
			break;
		case xltypeRef:
			//free(val.mref.lpmref);
			break;
		case xltypeMulti:
			//free(val.array.lparray);
			break;
		}
		Reset();
	}
};

template <typename T>
struct Boxed
{
	typedef T type;
};

template <>
struct Boxed < std::string >
{
	typedef const char * type;
};

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

template <typename T> struct ArgumentTypeText;

#define DECLARE_ARGUMENT_TYPE_TEXT(type, text) \
	template <> struct ArgumentTypeText < type > { \
		static const wchar_t * getTypeText() { return L##text; } \
	};

DECLARE_ARGUMENT_TYPE_TEXT(bool,       "A");
DECLARE_ARGUMENT_TYPE_TEXT(bool*,      "L");
DECLARE_ARGUMENT_TYPE_TEXT(double,     "B");
DECLARE_ARGUMENT_TYPE_TEXT(double*,    "E");
DECLARE_ARGUMENT_TYPE_TEXT(char*,      "C");
DECLARE_ARGUMENT_TYPE_TEXT(uint16_t,   "H");
DECLARE_ARGUMENT_TYPE_TEXT(int16_t,    "I");
DECLARE_ARGUMENT_TYPE_TEXT(int16_t*,   "M");
DECLARE_ARGUMENT_TYPE_TEXT(int32_t,    "J");
DECLARE_ARGUMENT_TYPE_TEXT(int32_t*,   "N");
DECLARE_ARGUMENT_TYPE_TEXT(wchar_t*,   "C%");
DECLARE_ARGUMENT_TYPE_TEXT(LPXLOPER12, "Q");

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
	template <typename T>
	static LPCWSTR GetTypeText()
	{
		return ArgumentTypeText<std::remove_const<T>::type>::getTypeText();
	}

	static FunctionInfo Create(LPCWSTR entryPoint)
	{
		const int NumArgs = sizeof...(TArgs);
		std::array<LPCWSTR, NumArgs + 1> texts = {
			GetTypeText<TRet>(), GetTypeText<TArgs>()...
		};
		std::wstring s;
		for (int i = 0; i <= NumArgs; i++)
		{
			s += texts[i];
		}
		return FunctionInfo(s.c_str(), entryPoint);
	}
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
	ExcelException(int errorCode)
		: m_errorCode(errorCode)
	{
		if (GetKnownErrorMessage(errorCode) == nullptr)
			sprintf_s(m_errorMessage, "xll error %d", m_errorCode);
	}

	virtual const char* what() const
	{
		const char *msg = GetKnownErrorMessage(m_errorCode);
		return msg? msg : m_errorMessage;
	}
};

class AddinRegistrar
{
	static std::vector<FunctionInfo> & registry()
	{
		static std::vector<FunctionInfo> s_functions;
		return s_functions;
	}

	static int RegisterFunction(LPXLOPER12 dllName, const FunctionInfo &f)
	{
		if (f.arguments.size() > 245)
			throw std::invalid_argument("Too many arguments");

		std::wstring argumentText;
		if (f.arguments.size() > 0)
		{
			argumentText = f.arguments[0].name();
			for (size_t i = 1; i < f.arguments.size(); i++)
			{
				argumentText += L", ";
				argumentText += f.arguments[i].name();
			}
		}

		ExcelVariant opers[256];
		// opers[0] = dllName;
		opers[1] = f.entryPoint;
		opers[2] = f.typeText;
		opers[3] = f.name;
		if (argumentText.empty())
			opers[4] = (wchar_t*)nullptr;
		else
			opers[4] = argumentText;
		opers[5] = f.macroType;
		opers[6] = f.category;
		opers[7] = f.shortcut;
		opers[8] = f.helpTopic;
		opers[9] = f.description;
		for (size_t i = 0; i < f.arguments.size(); i++)
			opers[10 + i] = f.arguments[i].description();

		LPXLOPER12 popers[256];
		popers[0] = dllName;
		for (size_t i = 1; i < 10u + f.arguments.size(); i++)
			popers[i] = &opers[i];

		XLOPER12 id;
		int ret = Excel12v(xlfRegister, &id, 10 + f.arguments.size(), popers);
		return ret;
	}

public:
	static FunctionInfo& AddFunction(FunctionInfo &f)
	{
		registry().push_back(f);
		return registry().back();
	}

	static void RegisterAllFunctions()
	{
		XLOPER12 xDLL;
		Excel12(xlGetName, &xDLL, 0); // TODO: check return value
		for (FunctionInfo& f : registry())
		{
			RegisterFunction(&xDLL, f);
		}
		Excel12(xlFree, 0, 1, &xDLL);
	}
};


template <typename Func, Func *, typename> struct NamedFunction;

template <typename Func, Func *func, typename TRet, typename... TArgs>
struct NamedFunction < Func, func, TRet(TArgs...) >
{
private:
	std::string m_name;
	std::string m_description;
	std::vector<NameDescriptionPair> m_args;

	NamedFunction() {}
	NamedFunction(const NamedFunction<Func,func,Func>&) = delete;

public:

	static NamedFunction<Func,func,Func>& Instance()
	{
		static NamedFunction<Func, func, Func> s_instance;
		return s_instance;
	}

	std::string Prototype()
	{
		return Prototype(func);
	}

	typedef NamedFunction<Func, func, Func> self_type;

	self_type & Name(const std::string &name)
	{
		m_name = name;
		return (*this);
	}

	self_type & Description(const std::string &description)
	{
		m_description = description;
		return (*this);
	}

	self_type & Arg(const std::string &name, const std::string &description)
	{
		m_args.push_back(NameDescriptionPair(name, description));
		return (*this);
	}

private:

	template <typename TRet, typename... TArgs>
	std::string Prototype(TRet(*)(TArgs...))
	{
		return std::string(typeid(TRet).name()) + " " + m_name +
			"(" + FormatArgument<TArgs...>(0) + ")";
	}

	template <typename First, typename Second, typename... Rest>
	std::string FormatArgument(size_t k)
	{
		return std::string(typeid(First).name()) +
			(k < m_args.size() ? " " + m_args[k] : std::string()) + ", " +
			FormatArgument<Second, Rest...>(k + 1);
	}

	template <typename TArg>
	std::string FormatArgument(size_t k)
	{
		return std::string(typeid(TArg).name()) +
			(k < m_args.size() ? " " + m_args[k] : std::string());
	}

	std::string FormatArgument(size_t)
	{
		return "";
	}
};

#define NAMED_FUNCTION(f) NamedFunction<decltype(f), f, decltype(f)>::Instance()


#if !defined(_WIN64)
// On 32-bit platform, we use naked function to emit a JMP instruction,
// and export this function. The function must be __cdecl.

template <typename Func, Func *func, typename> struct XLWrapper { };
template <typename Func, Func *func, typename TRet, typename... TArgs>
struct XLWrapper < Func, func, TRet(TArgs...) >
{ 
	static typename Boxed<TRet>::type __stdcall Call(typename Boxed<TArgs>::type... args)
	{
		return func(args...);
	}
};

#define EXPORT_DLL_FUNCTION(name, implementation) \
	extern "C" __declspec(naked, dllexport) void name() \
	{ \
		static const void *fp = static_cast<const void *>(implementation); \
		__asm { jmp [fp] } \
	}

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
