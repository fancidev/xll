#pragma once

#include "xlldef.h"
#include <exception>
#include <string>

namespace XLL_NAMESPACE{}

class ExcelRef : public xlref12
{
};

void XLOPER12_Clear(LPXLOPER12 p);
//
//template <typename T>
//void XLOPER12_Create(LPXLOPER12 pv, const T &value)
//{
//	static_assert(false, "Don't know how to convert the specified type to XLOPER12. "
//		"Overload XLOPER12_Create() to fix this issue.");
//}
//
//inline void XLOPER12_Create(LPXLOPER12 pv, double value)
//{
//	pv->xltype = xltypeNum;
//	pv->val.num = value;
//}
//
//inline void XLOPER12_Create(LPXLOPER12 pv, unsigned long value)
//{
//	XLOPER12_Create(pv, static_cast<double>(value));
//}
//
////ExcelVariant(const char *s);
//
//void XLOPER12_Create(LPXLOPER12 pv, const wchar_t *s, size_t len);
//
//inline void XLOPER12_Create(LPXLOPER12 pv, const wchar_t *s)
//{
//	XLOPER12_Create(pv, s, s ? (size_t)lstrlenW(s) : 0);
//}
//
//inline void XLOPER12_Create(LPXLOPER12 pv, const std::wstring &value)
//{
//	XLOPER12_Create(pv, value.c_str(), value.size());
//}
//
//inline void XLOPER12_Create(LPXLOPER12 pv, bool value)
//{
//	pv->xltype = xltypeBool;
//	pv->val.xbool = value;
//}
//
//inline void XLOPER12_Create(LPXLOPER12 pv, int value)
//{
//	pv->xltype = xltypeInt;
//	pv->val.w = value;
//}

// ExcelVariant() ref

// ExcelVariant() err

// ExcelVariant() flow

// ExcelVariant() array

// ExcelVariant() missing

// ExcelVariant() nil

//ExcelVariant(const ExcelRef &ref)
//{
//	xltype = xltypeSRef;
//	val.sref.count = 1;
//	val.sref.ref = ref;
//}




// Wraps an XLOPER12 and automatically releases memory on destruction.
// Use this class when you pass arguments to an Excel function.
class ExcelVariant : public XLOPER12
{
	static ExcelVariant FromType(WORD xltype)
	{
		ExcelVariant v;
		v.xltype = xltype;
		return v;
	}

	static ExcelVariant MakeError(int err)
	{
		ExcelVariant v;
		v.xltype = xltypeErr;
		v.val.err = err;
		return v;
	}


public:
	static const ExcelVariant Empty;
	static const ExcelVariant Missing;
	static const ExcelVariant ErrValue;

	ExcelVariant()
	{
		xltype = 0;
	}

	/*explicit ExcelVariant(const XLOPER12 &other)
	{
		Copy(*this, other);
	}*/

	ExcelVariant(const ExcelVariant &other) = delete;

	ExcelVariant& operator=(const ExcelVariant &other) = delete;

	ExcelVariant& operator=(ExcelVariant &&other)
	{
		if (&other != this)
		{
			XLOPER12 tmp;
			memcpy(&tmp, &other, sizeof(XLOPER12));
			memcpy(&other, this, sizeof(XLOPER12));
			memcpy(this, &tmp, sizeof(XLOPER12));
		}
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
		wchar_t *p = (wchar_t*)malloc(sizeof(wchar_t) * 2);
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
			throw std::invalid_argument("input string is too long");

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

	//ExcelVariant(const ExcelRef &ref)
	//{
	//	xltype = xltypeSRef;
	//	val.sref.count = 1;
	//	val.sref.ref = ref;
	//}

	ExcelVariant(int value)
	{
		xltype = xltypeInt;
		val.w = value;
	}

	// ExcelVariant() xltypeBigData

	~ExcelVariant()
	{
		XLOPER12_Clear(this);
	}

	// Returns the content of this object in a heap-allocated XLOPER12 suitable
	// to be returned to Excel. The XLOPER12 has its xlbitDLLFree bit set. The
	// content of this object is cleared.
	LPXLOPER12 detach()
	{
		LPXLOPER12 p = (LPXLOPER12)malloc(sizeof(XLOPER12));
		if (p == nullptr)
			throw std::bad_alloc();
		memcpy(p, this, sizeof(XLOPER12));
		p->xltype |= xlbitDLLFree;
		xltype = 0;
		return p;
	}
};