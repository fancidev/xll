#include <Windows.h>
#include <cassert>
#include "XllAddin.h"

#define EXPORT_UNDECORATED_NAME comment(linker, "/export:" __FUNCTION__ "=" __FUNCDNAME__)

const ExcelVariant ExcelVariant::Empty(ExcelVariant::FromType(xltypeNil));
const ExcelVariant ExcelVariant::Missing(ExcelVariant::FromType(xltypeMissing));
const ExcelVariant ExcelVariant::ErrValue(ExcelVariant::MakeError(xlerrValue));

double __stdcall MyTestFunc(double x, double y)
{
#pragma EXPORT_UNDECORATED_NAME
	return x * y;
}

const wchar_t * __stdcall MyToString(double x)
{
#pragma EXPORT_UNDECORATED_NAME
	static wchar_t result[100];
	swprintf(result, 100, L"%lf", x);
	return result;
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
			argumentText += L",";
			argumentText += f.arguments[i].name();
		}
	}

	ExcelVariant opers[256];
	// opers[0] = dllName;
	opers[1] = f.entryPoint;
	opers[2] = f.typeText + (f.isPure ? L"" : L"!") + (f.isThreadSafe ? L"$" : L"");
	opers[3] = f.name;
	opers[4] = argumentText;
	opers[5] = f.macroType;
	opers[6] = f.category;
	opers[7] = f.shortcut;
	opers[8] = f.helpTopic;
	//opers[8] = L"e:\\Dev\\Repos\\Xll\\Test\\A.chm!123";
	opers[9] = f.description;
	for (size_t i = 0; i < f.arguments.size(); i++)
	{
		// Excel sometimes truncates the last one or two characters of the
		// last argument description. Therefore we need to append two spaces
		// to the last argument description to counter this behavior. See 
		// https://msdn.microsoft.com/en-us/library/office/bb687841.aspx
		if (i == f.arguments.size() - 1 && f.arguments[i].description() != nullptr)
			opers[10 + i] = std::wstring(f.arguments[i].description()) + L"  ";
		else
			opers[10 + i] = f.arguments[i].description();
	}

	LPXLOPER12 popers[256];
	popers[0] = dllName;
	for (size_t i = 1; i < 10u + f.arguments.size(); i++)
		popers[i] = &opers[i];

	// If opers[9] is supplied, regardless of its value, Excel will not
	// automatically fill in argument text. So we do not supply it unless
	// user has specified something.
	int n;
	if (f.description == nullptr && f.arguments.size() == 0)
		n = 9;
	else
		n = 10 + f.arguments.size();

	XLOPER12 id;
	int ret = Excel12v(xlfRegister, &id, n, popers);
	return ret;
}

static void RegisterAllFunctions()
{
	XLOPER12 xDLL;
	Excel12(xlGetName, &xDLL, 0); // TODO: check return value

#if 0
	Excel12(xlfRegister, 0, 6,
		&xDLL,
		&ExcelVariant(L"MyToString"),
		&ExcelVariant(L"C%B"),
		&ExcelVariant(L"MyToString"),
		&ExcelVariant(L"x"),
		&ExcelVariant(1.0));
#endif

#if 0
	Excel12(xlfRegister, 0, 10,
		&xDLL,
		&ExcelVariant(L"MyTestFunc"),
		&ExcelVariant(L"BBB"),
		&ExcelVariant(L"MyTestFunc"),
		&ExcelVariant(L"a,b,c"), // 4: argumentText; extra arguments are shown but if you fill them, you get error
		&ExcelVariant(1.0),
		&ExcelVariant(L""),
		&ExcelVariant(L""), 
		&ExcelVariant(L""), // 8
		//&ExcelVariant((wchar_t*)nullptr)  // 9
		&ExcelVariant::Missing // 9 -- if this argument is supplied, regardless of
								// what value it is, Excel will not automatically
								// fill in argument text. You must supply it in
								// argumentText.
		);
#endif

	for (FunctionInfo& f : AddinRegistrar::registry())
	{
		RegisterFunction(&xDLL, f);
	}

	Excel12(xlFree, 0, 1, &xDLL);
}

int WINAPI xlAutoOpen()
{
#pragma EXPORT_UNDECORATED_NAME

#if 0
	static XLOPER12 xDLL;
	Excel12(xlGetName, &xDLL, 0);
	//MessageBoxW(NULL, L"xlAutoOpen", L"MyAddin", MB_OK);
	Excel12(xlfRegister, 0, 4,
		&xDLL,
		&ExcelVariant(L"CalcCircum"),
		&ExcelVariant(L"BB"),
		&ExcelVariant(L"CalcCircum"));

	Excel12(xlfRegister, 0, 4,
		&xDLL,
		&ExcelVariant(L"XLSquare"),
		&ExcelVariant(L"BB"),
		&ExcelVariant(L"Mysquare"));

	Excel12(xlFree, 0, 1, &xDLL);
#else
	RegisterAllFunctions();
#endif

	return 1;
}

int WINAPI xlAutoClose()
{
#pragma EXPORT_UNDECORATED_NAME

	return 1;
}

LPXLOPER12 WINAPI xlAutoRegister12(LPXLOPER12 pxName)
{
	return 0;
}

int WINAPI xlAutoAdd(void)
{
	return 1;
}

int WINAPI xlAutoRemove(void)
{
	return 1;
}

LPXLOPER12 WINAPI xlAddInManagerInfo12(LPXLOPER12 xAction)
{
	return 0;
}

#if XLL_SUPPORT_THREAD_LOCAL
__declspec(thread) XLOPER12 xllReturnValue;
#endif

void WINAPI xlAutoFree12(LPXLOPER12 p)
{
#pragma EXPORT_UNDECORATED_NAME
	if (p)
	{
#if XLL_SUPPORT_THREAD_LOCAL
		assert(p == &xllReturnValue);
		XLOPER12_Clear(p);
#else
		XLOPER12_Clear(p);
		free(p);
#endif
	}
}

void XLOPER12_Create(LPXLOPER12 pv, const wchar_t *s, size_t len)
{
	if (s == nullptr)
	{
		pv->xltype = xltypeMissing;
		return;
	}

	if (len > 32767)
		throw new std::invalid_argument("input string is too long");

	wchar_t *p = (wchar_t*)malloc(sizeof(wchar_t)*(len + 1));
	if (p == nullptr)
		throw std::bad_alloc();

	p[0] = (wchar_t)len;
	memcpy(&p[1], s, len*sizeof(wchar_t));

	pv->xltype = xltypeStr | xlbitDLLFree;
	pv->val.str = p;
}

void XLOPER12_Clear(XLOPER12 *p)
{
	if (p == nullptr)
		return;

	switch (p->xltype & ~xlbitDLLFree)
	{
	case xltypeStr:
		free(p->val.str);
		break;
	case xltypeRef:
		free(p->val.mref.lpmref);
		break;
	case xltypeMulti:
		if (p->val.array.lparray != nullptr)
		{
			int nr = p->val.array.rows;
			int nc = p->val.array.columns;
			int count = nr*nc;
			if (nr > 0 && nc > 0 && count > 0)
			{
				for (int i = 0; i < count; i++)
					XLOPER12_Clear(&p->val.array.lparray[i]);
			}
			free(p->val.array.lparray);
		}
		break;
	}
	p->xltype = 0;
}

void ExcelVariant::Copy(XLOPER12 &to, const XLOPER12 &from)
{
	if (&from == &to)
		return;

	memcpy(&to, &from, sizeof(XLOPER12));

	switch (from.xltype)
	{
	case xltypeStr:
		if (from.val.str != nullptr)
		{
			int len = (unsigned short)from.val.str[0];
			to.val.str = (wchar_t*)malloc(sizeof(wchar_t)*(len + 1));
			if (to.val.str == nullptr)
				throw std::bad_alloc();
			memcpy(to.val.str, from.val.str, sizeof(wchar_t)*(len + 1));
		}
		break;
	case xltypeRef:
		if (from.val.mref.lpmref != nullptr)
		{
			int count = from.val.mref.lpmref->count;
			if (count == 0)
			{
				LPXLMREF12 p = (LPXLMREF12)malloc(sizeof(XLMREF12));
				if (p == nullptr)
					throw std::bad_alloc();
				p->count = (WORD)count;
				to.val.mref.lpmref = p;
			}
			else
			{
				LPXLMREF12 p = (LPXLMREF12)malloc(sizeof(XLMREF12) + sizeof(XLREF12)*(count - 1));
				if (p == nullptr)
					throw std::bad_alloc();
				p->count = (WORD)count;
				memcpy(p->reftbl, from.val.mref.lpmref->reftbl, sizeof(XLREF12)*count);
				to.val.mref.lpmref = p;
			}
		}
		break;
	case xltypeMulti:
		if (from.val.array.lparray != nullptr)
		{
			int count = from.val.array.rows * from.val.array.columns;
			LPXLOPER12 p = (LPXLOPER12)malloc(sizeof(XLOPER12)*count); // todo: free if exception
			if (p == nullptr)
				throw std::bad_alloc();
			for (int i = 0; i < count; i++)
			{
				Copy(p[i], from.val.array.lparray[i]);
			}
			to.val.array.lparray = p;
		}
		break;
	case xltypeBigData:
		if (from.val.bigdata.h.lpbData != nullptr && from.val.bigdata.cbData > 0)
		{
			size_t numBytes = from.val.bigdata.cbData;
			BYTE *p = (BYTE*)malloc(numBytes);
			if (p == nullptr)
				throw std::bad_alloc();
			memcpy(p, from.val.bigdata.h.lpbData, numBytes);
			to.val.bigdata.h.lpbData = p;
		}
		else
		{
			to.xltype = 0;
		}
		break;
	}
}

static HRESULT Copy(VARIANT &v, const XLOPER12 &from, bool allowArray)
{
	HRESULT hr = S_OK;
	VariantInit(&v);
	switch (from.xltype)
	{
	case xltypeNum:
		V_VT(&v) = VT_R8;
		V_R8(&v) = from.val.num;
		break;
	case xltypeStr:
		if (from.val.str != nullptr)
		{
			BSTR s = SysAllocStringLen(&from.val.str[1], from.val.str[0]);
			if (s == nullptr)
				hr = E_OUTOFMEMORY;
			else
			{
				V_VT(&v) = VT_BSTR;
				V_BSTR(&v) = s;
			}
		}
		break;
	case xltypeBool:
		V_VT(&v) = VT_BOOL;
		V_BOOL(&v) = from.val.xbool;
		break;
	case xltypeErr:
		V_VT(&v) = VT_ERROR;
		V_ERROR(&v) = 0x800A07D0 + from.val.err;
		break;
	case xltypeMissing:
		V_VT(&v) = VT_ERROR;
		V_ERROR(&v) = 0x80020004;
		break;
	case xltypeNil:
		V_VT(&v) = VT_EMPTY;
		break;
	case xltypeInt:
		V_VT(&v) = VT_I4;
		V_I4(&v) = from.val.w;
		break;
	case xltypeMulti:
		if (!allowArray)
			hr = E_INVALIDARG;
		if (SUCCEEDED(hr) &&
			from.val.array.lparray != nullptr &&
			from.val.array.rows > 0 &&
			from.val.array.columns > 0)
		{
			int nr = from.val.array.rows;
			int nc = from.val.array.columns;
			LPXLOPER12 src = from.val.array.lparray;

			SAFEARRAYBOUND bounds[2];
			bounds[0].cElements = nc;
			bounds[0].lLbound = 1;
			bounds[1].cElements = nr;
			bounds[1].lLbound = 1;

			SAFEARRAY *psa = SafeArrayCreate(VT_VARIANT, 2, bounds);
			if (psa == nullptr)
				hr = E_OUTOFMEMORY;
			if (SUCCEEDED(hr))
			{
				VARIANT *dest;
				hr = SafeArrayAccessData(psa, (void**)&dest);
				if (SUCCEEDED(hr))
				{
					int count = nr*nc;
					for (int i = 0; i < count; i++)
					{
						hr = Copy(dest[i], src[i], false);
						if (FAILED(hr))
						{
							for (int j = 0; j < i; j++)
								VariantClear(&dest[j]);
							break;
						}
					}
					SafeArrayUnaccessData(psa);
					if (SUCCEEDED(hr))
					{
						V_VT(&v) = VT_ARRAY | VT_VARIANT;
						V_ARRAY(&v) = psa;
					}
				}
			}
		}
		break;
	default:
	case xltypeBigData:
	case xltypeFlow:
	case xltypeRef:
	case xltypeSRef:
		hr = E_NOTIMPL;
	}
	return hr;
}

// TODO: free memory
VARIANT ArgumentWrapper<VARIANT>::unwrap(LPXLOPER12 p)
{
	VARIANT v;
	HRESULT hr = Copy(v, *p, true);
	if (FAILED(hr))
		throw std::invalid_argument("Cannot convert XLOPER12 to VARIANT.");
	return v;
}