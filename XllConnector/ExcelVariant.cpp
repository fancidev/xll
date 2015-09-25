////////////////////////////////////////////////////////////////////////////
// TypeConversion.cpp -- helper routines to convert between types

#include "ExcelVariant.h"

//
// INT2DBL
//
// Helper macro to convert an integer bit-wise to a double.
//
// This is equivalent to reinterpret_cast<const double &>(n), except
// that n is not required to be an lvalue.
//

namespace
{
	union IntegerToDouble
	{
		__int64 i;
		double d;
	};
}
#define INT2DBL(n) (IntegerToDouble{n}.d)

//
// XLOPER12 Constants
//
// These constants are constructed at compile time and stored in the
// image directly.
//

namespace XLL_NAMESPACE
{

#define XLCONST(type, value) { { INT2DBL(value) }, type }
	const XLOPER12 Constants::Empty = XLCONST(xltypeNil, 0);
	const XLOPER12 Constants::Missing = XLCONST(xltypeMissing, 0);

#define XLERROR(err) XLCONST(xltypeErr, err)
	const XLOPER12 Constants::ErrNull = XLERROR(xlerrNull);
	const XLOPER12 Constants::ErrDiv0 = XLERROR(xlerrDiv0);
	const XLOPER12 Constants::ErrValue = XLERROR(xlerrValue);
	const XLOPER12 Constants::ErrRef = XLERROR(xlerrRef);
	const XLOPER12 Constants::ErrName = XLERROR(xlerrName);
	const XLOPER12 Constants::ErrNum = XLERROR(xlerrNum);
	const XLOPER12 Constants::ErrNA = XLERROR(xlerrNA);
	const XLOPER12 Constants::ErrGettingData = XLERROR(xlerrGettingData);
}
