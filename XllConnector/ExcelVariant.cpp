#include "ExcelVariant.h"

using namespace XLL_NAMESPACE;

const ExcelVariant ExcelVariant::Empty(ExcelVariant::FromType(xltypeNil));
const ExcelVariant ExcelVariant::Missing(ExcelVariant::FromType(xltypeMissing));
const ExcelVariant ExcelVariant::ErrValue(ExcelVariant::MakeError(xlerrValue));

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

