////////////////////////////////////////////////////////////////////////////
// ArrayExample.cpp
//
// This file demonstrates how to take array arguments from and return
// array to Excel. 
//
// The following types are supported as arguments:
//
//   SAFEARRAY *
//
// An argument passed from Excel is always converted to a two-dimensional
// SAFEARRAY with elements of type VARIANT. If a scalar (which may be
// empty) is passed from Excel, it is converted to a 1-by-1 array. If a
// missing value is passed (i.e. when a parameter is not specified), it
// is converted to a 0-by-0 array.
//
// The lower bound of both dimensions is always set to 1.
//
// When the array contains more than one element, the elements are stored
// in row-major order. Specifically,
//
//   1, The size of the first dimension is equal to the number of rows;
//      the size of the second dimension is equal to the number of columns.
//   2, Elements in the same row are stored in contiguous memory.
//
// (fortran array)
//
// In Excel 2003 and earlier, array size is limited to 65,536 rows by
// 256 columns. 

#include "XllAddin.h"
#include <cassert>
#include <comdef.h>
#include <comutil.h>
#include <algorithm>
#include <random>

// Helper class to access a two-dimensional SAFEARRAY with element type
// VARIANT. The wrapper always creates such an array when the UDF expects
// an argument of type SAFEARRAY*.
class VariantMatrixAccessor
{
	SAFEARRAY *m_psa;
	VARIANT *m_data;
public:
	explicit VariantMatrixAccessor(SAFEARRAY *psa)
	{
		assert(psa != nullptr);
		assert(SafeArrayGetDim(psa) == 2);

		VARTYPE vt;
		HRESULT hr = SafeArrayGetVartype(psa, &vt);
		if (FAILED(hr))
			throw _com_error(hr);
		assert(vt == VT_VARIANT);

		VARIANT *data;
		hr = SafeArrayAccessData(psa, (void**)&data);
		if (FAILED(hr))
			throw std::invalid_argument("Cannot access data");

		m_psa = psa;
		m_data = data;
	}

	~VariantMatrixAccessor()
	{
		if (m_psa != nullptr)
		{
			SafeArrayUnaccessData(m_psa);
			m_data = nullptr;
			m_psa = nullptr;
		}
	}

	size_t rows() const { return m_psa->rgsabound[0].cElements; }

	size_t columns() const { return m_psa->rgsabound[1].cElements; }

	size_t size() const { return rows()*columns(); }

	VARIANT& operator[](size_t index)
	{
		return m_data[index];
	}

	VARIANT& operator()(size_t row, size_t column) // zero-based
	{
		return m_data[row*columns() + column];
	}
};

template <typename T> T variant_cast(const VARIANT &);

template <> double variant_cast<double>(const VARIANT &v)
{
	VARIANT result;
	VariantInit(&result);
	HRESULT hr = VariantChangeType(&result, &v, 0, VT_R8);
	if (FAILED(hr))
		throw std::bad_cast();
	return V_R8(&result);
}

double Trace(SAFEARRAY *matrix)
{
	VariantMatrixAccessor mat(matrix);
	if (mat.rows() != mat.columns())
		throw std::invalid_argument("Only supports square matrix.");

	double sum = 0.0;
	size_t n = mat.rows();
	for (size_t i = 0; i < n; i++)
	{
		sum += variant_cast<double>(mat(i, i));
	}
	return sum;
}

EXPORT_XLL_FUNCTION(Trace)
.Description(L"Returns the sum of the diagonal elements of a square matrix.");

double PartialSum(SAFEARRAY *matrix, int count)
{
	if (count < 0)
		throw std::invalid_argument("Count must be greater than or equal to zero.");

	VariantMatrixAccessor mat(matrix);

	size_t n = mat.size();
	double sum = 0.0;
	for (size_t i = 0; i < n && i < (size_t)count; i++)
	{
		sum += variant_cast<double>(mat[i]);
	}
	return sum;
}

EXPORT_XLL_FUNCTION(PartialSum);

// ShuffleColumns -- reorder the columns in a matrix randomly.
// 
// This example shows how to use FP12* type argument, and how to return
// value in-place. To do this, define the function as void.
//
// This example also shows that the FP12* array is passed in row-major.

void ShuffleColumns(FP12 *mat)
{
	if (mat && mat->rows>0 && mat->columns > 0)
	{
		std::vector<int> order(mat->columns);
		for (int j = 0; j < mat->columns; j++)
			order[j] = j;

		std::default_random_engine rng;
		std::shuffle(std::begin(order), std::end(order), rng);

		for (int i = 0; i < mat->rows; i++)
		{
			for (int j = 0; j < mat->columns; j++)
			{
				std::swap(
					mat->array[i*mat->columns + j],
					mat->array[i*mat->columns + order[j]]);
			}
		}
	}
}

EXPORT_XLL_FUNCTION(ShuffleColumns, XLL_THREADSAFE);
