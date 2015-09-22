#pragma once

#include "xlldef.h"
#include <cstdint>

// 
// NormalizeAttributes
//
// Helper class to combine default attributes and explicit attributes
// to form effective attributes.
//

#define XLL_NO_MORE_THAN_ONE_BIT_SET(x) ( ((x) & ((x)-1)) == 0 )

namespace XLL_NAMESPACE
{
	template <int Attributes = 0>
	struct NormalizeAttributes
	{
		static_assert((Attributes & ~(
			XLL_VOLATILE | XLL_NOT_VOLATILE |
			XLL_THREADSAFE | XLL_NOT_THREADSAFE)) == 0,
			"Unknown attributes specified.");

		static_assert(XLL_NO_MORE_THAN_ONE_BIT_SET(
			Attributes & (XLL_VOLATILE | XLL_NOT_VOLATILE)),
			"Only one of XLL_VOLATILE and XLL_NOT_VOLATILE may be set.");

		static_assert(XLL_NO_MORE_THAN_ONE_BIT_SET(
			Attributes & (XLL_THREADSAFE | XLL_NOT_THREADSAFE)),
			"Only one of XLL_THREADSAFE and XLL_NOT_THREADSAFE may be set.");

		enum
		{
			volatility_value =
			(Attributes & XLL_VOLATILE) ? XLL_VOLATILE :
			(Attributes & XLL_NOT_VOLATILE) ? 0 :
			(XLL_DEFAULT_VOLATILE) ? XLL_VOLATILE : 0
		};

		enum
		{
			threadsafe_value =
			(Attributes & XLL_THREADSAFE) ? XLL_THREADSAFE :
			(Attributes & XLL_NOT_THREADSAFE) ? 0 :
			(XLL_DEFAULT_THREADSAFE) ? XLL_THREADSAFE : 0
		};

		enum { value = volatility_value | threadsafe_value };
	};
}

// 
// FunctionAttributes
//
// Helper class to interpret and validate function attribute constants.
//

namespace XLL_NAMESPACE
{
	template <int Attributes>
	struct FunctionAttributes
	{
		static_assert((Attributes & ~(XLL_VOLATILE | XLL_THREADSAFE)) == 0,
			"Invalid attributes specified.");

		enum { IsVolatile = (Attributes & XLL_VOLATILE) ? 1 : 0 };

		enum { IsThreadSafe = (Attributes & XLL_THREADSAFE) ? 1 : 0 };
	};
}

namespace XLL_NAMESPACE
{
	// 
	// Compile-time sequence and concatenation
	//

	template <typename T, T... Elem> struct Sequence
	{
		static const T * ToArray()
		{
			static const T array[] = { Elem... };
			return array;
		}
	};

	template <typename T, typename...> struct Concat;

	template <typename T>
	struct Concat < T >
	{
		typedef Sequence<T> type;
	};

	template <typename T, T... Elem>
	struct Concat < T, Sequence<T, Elem...> >
	{
		typedef Sequence<T, Elem...> type;
	};

	template <typename T, T... Elem1, T... Elem2, typename... Rest>
	struct Concat < T, Sequence<T, Elem1...>, Sequence<T, Elem2...>, Rest... >
	{
		typedef typename Concat<T, Sequence < T, Elem1..., Elem2... >, Rest...>::type type;
	};

	//
	// Mapping from wire type to Excel type text. For a full list of
	// supported wire types and their type text, see
	// https://msdn.microsoft.com/en-us/library/office/bb687900.aspx
	// 

	template <typename T, typename Char> struct TypeText
	{
		template <typename U> struct always_false : std::false_type {};
		static_assert(always_false<T>::value,
			"The supplied type is not a supported XLL wire type.");
	};

#define DEFINE_TYPE_TEXT(type, ...) \
	template <typename Char> struct TypeText<type, Char> { \
		typedef Sequence<Char, __VA_ARGS__> SeqType; \
	}

	DEFINE_TYPE_TEXT(bool, 'A');
	DEFINE_TYPE_TEXT(bool*, 'L');
	DEFINE_TYPE_TEXT(double, 'B');
	DEFINE_TYPE_TEXT(double*, 'E');
	DEFINE_TYPE_TEXT(char*, 'C'); // nul-terminated
	DEFINE_TYPE_TEXT(const char*, 'C'); // nul-terminated
	DEFINE_TYPE_TEXT(wchar_t*, 'C', '%'); // nul-terminated
	DEFINE_TYPE_TEXT(const wchar_t*, 'C', '%'); // nul-terminated
	DEFINE_TYPE_TEXT(uint16_t, 'H');
	DEFINE_TYPE_TEXT(int16_t, 'I');
	DEFINE_TYPE_TEXT(int16_t*, 'M');
	DEFINE_TYPE_TEXT(int32_t, 'J');
	DEFINE_TYPE_TEXT(int32_t*, 'N');
	DEFINE_TYPE_TEXT(FP12*, 'K', '%');
	DEFINE_TYPE_TEXT(LPXLOPER12, 'Q');

	template <typename Char, int Attributes>
	struct TypeText < FunctionAttributes<Attributes>, Char >
	{
		typedef std::conditional_t <
			FunctionAttributes<Attributes>::IsVolatile,
			Sequence<Char, '!'>,
			Sequence < Char >> VolatileText;
		typedef std::conditional_t <
			FunctionAttributes<Attributes>::IsThreadSafe,
			Sequence<Char, '$'>,
			Sequence < Char >> ThreadSafeText;
	};

	template <typename Char, int Attributes, typename TRet, typename... TArgs>
	inline const Char * GetTypeTextImpl(TRet(__stdcall*)(TArgs...))
	{
		typedef typename TypeText < FunctionAttributes<Attributes>, Char >
			AttributeTypeText;
		typedef typename Concat < Char,
			typename TypeText<typename TRet, Char>::SeqType,
			typename TypeText<typename TArgs, Char>::SeqType...,
			typename AttributeTypeText::VolatileText,
			typename AttributeTypeText::ThreadSafeText,
			Sequence<Char, 0> > ::type SeqType;
		return SeqType::ToArray();
	}
}