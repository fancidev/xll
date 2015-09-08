////////////////////////////////////////////////////////////////////////////
// Conversion.cpp -- helper functions to convert between data types

#include "Conversion.h"
#include <new>

XLL_BEGIN_NAMEPSACE

////////////////////////////////////////////////////////////////////////////
// Conversions to XLOPER12

template <> XLOPER12 make<XLOPER12>(const XLOPER12 &from)
{
	XLOPER12 to = from;

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
				p[i] = make<XLOPER12>(from.val.array.lparray[i]);
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
	return to;
}

template <> XLOPER12 make<XLOPER12>(double value)
{
	XLOPER12 dest;
	dest.xltype = xltypeNum;
	dest.val.num = value;
	return dest;
}

template <> XLOPER12 make<XLOPER12>(bool value)
{
	XLOPER12 dest;
	dest.xltype = xltypeBool;
	dest.val.xbool = value;
	return dest;
}

template <> XLOPER12 make<XLOPER12>(int value)
{
	XLOPER12 dest;
	dest.xltype = xltypeInt;
	dest.val.w = value;
	return dest;
}

static XLOPER12 make_XLOPER12_from_string(const wchar_t *s, size_t len)
{
	XLOPER12 dest;

	if (s == nullptr)
	{
		dest.xltype = xltypeMissing;
		return dest;
	}

	if (len > 32767)
		throw new std::invalid_argument("input string is too long");

	wchar_t *p = (wchar_t*)malloc(sizeof(wchar_t)*(len + 1));
	if (p == nullptr)
		throw std::bad_alloc();

	p[0] = (wchar_t)len;
	memcpy(&p[1], s, len*sizeof(wchar_t));

	dest.xltype = xltypeStr | xlbitDLLFree;
	dest.val.str = p;
	return dest;
}

template <> XLOPER12 make<XLOPER12>(const wchar_t *s)
{
	return make_XLOPER12_from_string(s, (s == nullptr) ? 0 : lstrlenW(s));
}

template <> XLOPER12 make<XLOPER12>(const std::wstring &s)
{
	return make_XLOPER12_from_string(s.c_str(), s.size());
}

template <> XLOPER12 make<XLOPER12>(unsigned long value)
{
	return make<XLOPER12>(static_cast<double>(value));
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