////////////////////////////////////////////////////////////////////////////
// TypeConversion.cpp -- helper routines to convert between types

#include "ExcelVariant.h"

using namespace XLL_NAMESPACE;

const ExcelVariant ExcelVariant::Empty(ExcelVariant::FromType(xltypeNil));
const ExcelVariant ExcelVariant::Missing(ExcelVariant::FromType(xltypeMissing));
const ExcelVariant ExcelVariant::ErrValue(ExcelVariant::MakeError(xlerrValue));

//// Provides conversion routines for an encapsulated XLOPER12 reference.
//struct XLOper12Ref
//{
//	XLOPER12 &x;
//	XLOper12Ref(XLOPER12 &y) :x(y){}
//	XLOper12Ref& operator=(double value){
//		x.xltype = xltypeNum;
//		x.val.num = value;
//		return *this;
//	}
//	void Clear()
//	{
//		// TO BE FILLED
//		const int n = sizeof(XLOPER12);
//		const int m = sizeof(VARIANT);
//		const int l = sizeof(SAFEARRAY);
//		const int k = sizeof(XLOPER);
//	}
//
//};


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
