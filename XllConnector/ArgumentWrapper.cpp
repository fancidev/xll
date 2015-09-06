#include "ArgumentWrapper.h"
#include <comdef.h>

using namespace XLL_NAMESPACE;

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


HRESULT SafeArrayCopyFrom(_In_ const XLOPER12 *src, _Out_ SAFEARRAY ** ppsa);

HRESULT VariantSet(VARIANTARG *dest, const XLOPER12 *src, bool allowArray)
{
	if (dest == nullptr || src == nullptr)
		return E_POINTER;

	HRESULT hr = S_OK;
	VariantInit(dest);
	switch (src->xltype & ~(xlbitDLLFree | xlbitXLFree))
	{
	case xltypeNum:
		V_VT(dest) = VT_R8;
		V_R8(dest) = src->val.num;
		break;
	case xltypeStr:
		if (src->val.str != nullptr)
		{
			BSTR s = SysAllocStringLen(&src->val.str[1], src->val.str[0]);
			if (s == nullptr)
			{
				hr = E_OUTOFMEMORY;
			}
			else
			{
				V_VT(dest) = VT_BSTR;
				V_BSTR(dest) = s;
			}
		}
		break;
	case xltypeBool:
		V_VT(dest) = VT_BOOL;
		V_BOOL(dest) = src->val.xbool;
		break;
	case xltypeErr:
		V_VT(dest) = VT_ERROR;
		V_ERROR(dest) = 0x800A07D0 + src->val.err;
		break;
	case xltypeMissing:
		V_VT(dest) = VT_ERROR;
		V_ERROR(dest) = 0x80020004;
		break;
	case xltypeNil:
		V_VT(dest) = VT_EMPTY;
		break;
	case xltypeInt:
		V_VT(dest) = VT_I4;
		V_I4(dest) = src->val.w;
		break;
	case xltypeMulti:
		if (!allowArray)
		{
			hr = E_INVALIDARG;
		}
		else
		{
			SAFEARRAY *psa;
			hr = SafeArrayCopyFrom(src, &psa);
			if (SUCCEEDED(hr))
			{
				V_VT(dest) = VT_ARRAY | VT_VARIANT;
				V_ARRAY(dest) = psa;
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

HRESULT SafeArrayCopyFrom(_In_ const XLOPER12 *src, _Out_ SAFEARRAY ** ppsa)
{
	if (ppsa == nullptr)
		return E_POINTER;
	if (src == nullptr)
		return E_INVALIDARG;

	*ppsa = nullptr;

	int nr, nc;
	switch (src->xltype & ~(xlbitDLLFree | xlbitXLFree))
	{
	case xltypeMissing:
		nr = 0;
		nc = 0;
		break;
	case xltypeMulti:
		nr = src->val.array.rows;
		nc = src->val.array.columns;
		src = src->val.array.lparray;
		break;
	case xltypeNil:
	default:
		nr = 1;
		nc = 1;
		break;
	}

	if (nr < 0 || nr > 0x100000 || nc < 0 || nc > 0x10000)
		return E_INVALIDARG;
	if (nr != 0 && nc != 0 && src == nullptr)
		return E_INVALIDARG;

	SAFEARRAYBOUND bounds[2];
	bounds[0].cElements = nr;
	bounds[0].lLbound = 1;
	bounds[1].cElements = nc;
	bounds[1].lLbound = 1;

	SAFEARRAY *psa = SafeArrayCreate(VT_VARIANT, 2, bounds);
	if (psa == nullptr)
		return E_OUTOFMEMORY;

	*ppsa = psa;
	if (nr == 0 || nc == 0)
		return S_OK;

	VARIANT *dest;
	HRESULT hr = SafeArrayAccessData(psa, (void**)&dest);
	if (SUCCEEDED(hr))
	{
		int count = nr*nc;
		for (int i = 0; i < count; i++)
		{
			hr = VariantSet(&dest[i], &src[i], false);
			if (FAILED(hr))
			{
				for (int j = 0; j < i; j++)
					VariantClear(&dest[j]);
				break;
			}
		}
		SafeArrayUnaccessData(psa);
	}
	if (FAILED(hr))
	{
		SafeArrayDestroy(psa);
		*ppsa = nullptr;
	}
	return hr;
}

SafeArrayWrapper::SafeArrayWrapper(const XLOPER12 *pv)
{
	HRESULT hr = SafeArrayCopyFrom(pv, &psa);
	if (FAILED(hr))
		throw _com_error(hr);
}

SafeArrayWrapper ArgumentWrapper<SAFEARRAY*>::unwrap(LPXLOPER12 pv)
{
	return SafeArrayWrapper(pv);
}
