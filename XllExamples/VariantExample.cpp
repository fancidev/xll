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