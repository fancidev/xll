#pragma once

#include "xlldef.h"
#include "Conversion.h"
#include <exception>
#include <string>

namespace XLL_NAMESPACE
{
	class ExcelRef : public xlref12
	{
	};

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
			CreateValue(this, value);
		}

		ExcelVariant(unsigned long value)
		{
			CreateValue(this, value);
		}

#if 0
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
#endif

		//ExcelVariant(const char *s)
		//{
		//}

		ExcelVariant(const wchar_t *s)
		{
			HRESULT hr = CreateValue(this, s);
			if (FAILED(hr))
				throw std::invalid_argument("Cannot convert string to ExcelVariant");
		}

		ExcelVariant(const std::wstring &value)
			: ExcelVariant(value.c_str())
		{
		}

		ExcelVariant(bool value)
		{
			CreateValue(this, value);
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
			CreateValue(this, value);
		}

		// ExcelVariant() xltypeBigData

		~ExcelVariant()
		{
			DeleteValue(this);
		}

#if 0
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
#endif
	};
}