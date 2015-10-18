////////////////////////////////////////////////////////////////////////////
// XLString.h -- Excel length-prefixed string operations

#pragma once

#include <type_traits>
#include <string>

namespace xll
{
	template <class Char>
	struct XLString
	{
		Char m_buffer[1];

	public:
		XLString() = delete;
		XLString(const XLString &) = delete;
		XLString(XLString &&) = delete;
		XLString& operator=(const XLString &) = delete;
		XLString& operator=(XLString&&) = delete;

		bool operator==(const Char *s) const
		{
			const Char *p = &m_buffer[1];
			size_t i;
			for (i = 0; i < m_buffer[0]; i++)
			{
				if (p[i] != s[i])
					return false;
			}
			return s[i] == 0;
		}

		size_t length() const 
		{
			return static_cast<std::make_unsigned<Char>::type>(m_buffer[0]);
		}

		const Char* cbegin() const
		{
			return &m_buffer[1]; 
		}

		const Char* cend() const
		{
			return cbegin() + length();
		}

		operator std::basic_string<Char>() const
		{
			return std::basic_string<Char>(cbegin(), length());
		}

	public:
		static const XLString& FromBuffer(const Char *buffer)
		{
			return reinterpret_cast<const XLString &>(*buffer);
		}
	};

	typedef XLString<char> XLStringA;
	typedef XLString<wchar_t> XLStringW;

	struct LengthPrefixedStringLiteral
	{
		const wchar_t * s;
	};
	//typedef LengthPrefixedConstString XLPCSTR

	
	template <class T, size_t Length>
	struct LengthPrefixedString
	{
		T len;
		T str[Length + 1];
	};

	//template <size_t N>
	//LengthPrefixedString<wchar_t,N> MakeLPS(const wchar_t(&s)[N])
	//{
	//	LengthPrefixedString<wchar_t,N> buffer = { static_cast<wchar_t>(N), s };
	//	return buffer;
	//}
	
	// s must be a string literal
#define XLSTR(s) LengthPrefixedString<char,sizeof(s)-1>{sizeof(s)-1,s}
//#define XLSTR(s) std::array<char,sizeof(s)> {sizeof(s)-1,s}


}