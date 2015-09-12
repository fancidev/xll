////////////////////////////////////////////////////////////////////////////
// Conversion.cpp -- helper functions to convert between data types

#include "Conversion.h"
#include <new>

XLL_BEGIN_NAMEPSACE

////////////////////////////////////////////////////////////////////////////
// Conversions to XLOPER12

HRESULT SetValue(LPXLOPER12 dest, const XLOPER12 &from)
{
#if _DEBUG
	if (dest == nullptr)
		return E_INVALIDARG;
#endif

	memcpy(dest, &from, sizeof(XLOPER12));
	switch (from.xltype)
	{
	case xltypeStr:
		if (from.val.str != nullptr)
		{
			int len = (unsigned short)from.val.str[0];
			dest->val.str = (wchar_t*)malloc(sizeof(wchar_t)*(len + 1));
			if (dest->val.str == nullptr)
				return E_OUTOFMEMORY;
			memcpy(dest->val.str, from.val.str, sizeof(wchar_t)*(len + 1));
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
					return E_OUTOFMEMORY;
				p->count = (WORD)count;
				dest->val.mref.lpmref = p;
			}
			else
			{
				LPXLMREF12 p = (LPXLMREF12)malloc(sizeof(XLMREF12) + sizeof(XLREF12)*(count - 1));
				if (p == nullptr)
					return E_OUTOFMEMORY;
				p->count = (WORD)count;
				memcpy(p->reftbl, from.val.mref.lpmref->reftbl, sizeof(XLREF12)*count);
				dest->val.mref.lpmref = p;
			}
		}
		break;
	case xltypeMulti:
		if (from.val.array.lparray != nullptr)
		{
			int count = from.val.array.rows * from.val.array.columns;
			LPXLOPER12 p = (LPXLOPER12)malloc(sizeof(XLOPER12)*count);
			if (p == nullptr)
				return E_OUTOFMEMORY;

			for (int i = 0; i < count; i++)
			{
				HRESULT hr = SetValue(&p[i], from.val.array.lparray[i]);
				if (FAILED(hr))
				{
					free(p);
					return hr;
				}
			}
			dest->val.array.lparray = p;
		}
		break;
	case xltypeBigData:
		if (from.val.bigdata.h.lpbData != nullptr && from.val.bigdata.cbData > 0)
		{
			size_t numBytes = from.val.bigdata.cbData;
			BYTE *p = (BYTE*)malloc(numBytes);
			if (p == nullptr)
				return E_OUTOFMEMORY;
			memcpy(p, from.val.bigdata.h.lpbData, numBytes);
			dest->val.bigdata.h.lpbData = p;
		}
		else
		{
			dest->xltype = 0;
		}
		break;
	}
	return S_OK;
}

HRESULT SetValue(LPXLOPER12 dest, double value)
{
#if _DEBUG
	if (dest == nullptr)
		return E_POINTER;
#endif
	dest->xltype = xltypeNum;
	dest->val.num = value;
	return S_OK;
}

HRESULT SetValue(LPXLOPER12 dest, int value)
{
#if _DEBUG
	if (dest == nullptr)
		return E_POINTER;
#endif
	dest->xltype = xltypeInt;
	dest->val.w = value;
	return S_OK;
}

HRESULT SetValue(LPXLOPER12 dest, unsigned long value)
{
	return SetValue(dest, static_cast<double>(value));
}

HRESULT SetValue(LPXLOPER12 dest, bool value)
{
#if _DEBUG
	if (dest == nullptr)
		return E_POINTER;
#endif
	dest->xltype = xltypeBool;
	dest->val.xbool = value;
	return S_OK;
}

HRESULT SetValue(LPXLOPER12 dest, const wchar_t *s, size_t len)
{
#if _DEBUG
	if (dest == nullptr)
		return E_POINTER;
#endif
	if (s == nullptr)
	{
		dest->xltype = xltypeMissing;
		return S_OK;
	}

	if (len > 32767u)
		return E_INVALIDARG;

	wchar_t *p = (wchar_t*)malloc(sizeof(wchar_t)*(len + 1));
	if (p == nullptr)
		return E_OUTOFMEMORY;

	p[0] = (wchar_t)len;
	memcpy(&p[1], s, len*sizeof(wchar_t));

	dest->xltype = xltypeStr | xlbitDLLFree;
	dest->val.str = p;
	return S_OK;
}

HRESULT SetValue(LPXLOPER12 dest, const wchar_t *s)
{
	return SetValue(dest, s, (s == nullptr) ? 0 : lstrlenW(s));
}

HRESULT SetValue(LPXLOPER12 dest, const std::wstring &s)
{
	return SetValue(dest, s.c_str(), s.size());
}

////////////////////////////////////////////////////////////////////////////
// Conversions to VARIANT

template <> VARIANT make<VARIANT>(const XLOPER12 &src)
{
	VARIANT dest;
	VariantInit(&dest);

	switch (src.xltype & ~(xlbitDLLFree | xlbitXLFree))
	{
	case xltypeNum:
		V_VT(&dest) = VT_R8;
		V_R8(&dest) = src.val.num;
		break;
	case xltypeStr:
		if (src.val.str != nullptr)
		{
			BSTR s = SysAllocStringLen(&src.val.str[1], src.val.str[0]);
			if (s == nullptr)
				throw std::bad_alloc();
			V_VT(&dest) = VT_BSTR;
			V_BSTR(&dest) = s;
		}
		break;
	case xltypeBool:
		V_VT(&dest) = VT_BOOL;
		V_BOOL(&dest) = src.val.xbool;
		break;
	case xltypeErr:
		V_VT(&dest) = VT_ERROR;
		V_ERROR(&dest) = 0x800A07D0 + src.val.err;
		break;
	case xltypeMissing:
		V_VT(&dest) = VT_ERROR;
		V_ERROR(&dest) = 0x80020004;
		break;
	case xltypeNil:
		V_VT(&dest) = VT_EMPTY;
		break;
	case xltypeInt:
		V_VT(&dest) = VT_I4;
		V_I4(&dest) = src.val.w;
		break;
	case xltypeMulti:
		V_VT(&dest) = VT_ARRAY | VT_VARIANT;
		V_ARRAY(&dest) = make<LPSAFEARRAY>(src);
		break;
	default:
	case xltypeBigData:
	case xltypeFlow:
	case xltypeRef:
	case xltypeSRef:
		throw std::invalid_argument("Not supported");
		break;
	}
	return dest;
}

////////////////////////////////////////////////////////////////////////////
// Conversions to LPSAFEARRAY

template <> LPSAFEARRAY make<LPSAFEARRAY>(const XLOPER12 &src)
{
	const XLOPER12 *pSrc;
	int nr, nc;
	switch (src.xltype & ~(xlbitDLLFree | xlbitXLFree))
	{
	case xltypeMissing:
		nr = 0;
		nc = 0;
		pSrc = nullptr;
		break;
	case xltypeMulti:
		nr = src.val.array.rows;
		nc = src.val.array.columns;
		pSrc = src.val.array.lparray;
		break;
	case xltypeNil:
	default:
		nr = 1;
		nc = 1;
		pSrc = &src;
		break;
	}

	if (nr < 0 || nr > 0x100000 || nc < 0 || nc > 0x10000)
		throw std::invalid_argument("invalid array size");
	if (nr != 0 && nc != 0 && pSrc == nullptr)
		throw std::invalid_argument("invalid input pointer");

	SAFEARRAYBOUND bounds[2];
	bounds[0].cElements = nr;
	bounds[0].lLbound = 1;
	bounds[1].cElements = nc;
	bounds[1].lLbound = 1;

	SAFEARRAY *psa = SafeArrayCreate(VT_VARIANT, 2, bounds);
	if (psa == nullptr)
		throw std::bad_alloc();

	if (nr == 0 || nc == 0)
		return psa;

	VARIANT *dest;
	HRESULT hr = SafeArrayAccessData(psa, (void**)&dest);
	if (FAILED(hr))
	{
		SafeArrayDestroy(psa);
		throw std::invalid_argument("cannot access array");
	}

	int count = nr*nc;
	for (int i = 0; i < count; i++)
	{
		try
		{
			dest[i] = make<VARIANT>(pSrc[i]);
		}
		catch (...)
		{
			for (int j = 0; j < i; j++)
			{
				VariantClear(&dest[j]);
			}
			SafeArrayUnaccessData(psa);
			SafeArrayDestroy(psa);
			throw;
		}
	}
	SafeArrayUnaccessData(psa);
	return psa;
}

XLL_END_NAMESPACE