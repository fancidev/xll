#include "ArgumentWrapper.h"
#include <comdef.h>

using namespace XLL_NAMESPACE;

template <> VARIANT value_cast<VARIANT>(const XLOPER12 &src)
{
	VARIANT ret;
	Copy(ret, src, true);
	return ret;
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
