#include "XllAddin.h"

std::wstring VariantType(VARIANTARG *pv)
{
	switch (V_VT(pv))
	{
	case VT_BSTR:
		return L"string";
	case VT_VARIANT|VT_ARRAY:
		return L"array";
	case VT_R8:
		return L"double";
	default:
		return L"other";
	}
}

EXPORT_XLL_FUNCTION(VariantType);

int BigFunc(
	int, double, int, double, int,
	VARIANT*, VARIANT*, VARIANT*, VARIANT*, VARIANT*,
	const char *, const wchar_t *, std::wstring, const std::wstring &)
{
	return 7;
}

EXPORT_XLL_FUNCTION(BigFunc);
