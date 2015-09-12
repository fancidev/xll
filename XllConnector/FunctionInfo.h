////////////////////////////////////////////////////////////////////////////
// FunctionInfo.h -- structs to hold UDF description and attributes

#pragma once

#include "xlldef.h"
#include <Windows.h>
#include <string>
#include <vector>
#include <array>

namespace XLL_NAMESPACE
{
	class NameDescriptionPair
	{
		LPCWSTR m_name;
		LPCWSTR m_description;
	public:
		NameDescriptionPair(LPCWSTR name, LPCWSTR description)
			: m_name(name), m_description(description)
		{
		}
		LPCWSTR name() const { return m_name; }
		LPCWSTR description() const { return m_description; }
	};

	struct FunctionInfo
	{
		LPCWSTR entryPoint;
		std::wstring typeText;

		LPCWSTR name;
		LPCWSTR description;
		std::vector<NameDescriptionPair> arguments;
		int macroType; // 0,1,2
		LPCWSTR category;
		LPCWSTR shortcut;
		LPCWSTR helpTopic;

		bool isPure;
		bool isThreadSafe;

		FunctionInfo(LPCWSTR typeText, LPCWSTR entryPoint)
			: typeText(typeText), entryPoint(entryPoint), name(), description(),
			macroType(1), category(), shortcut(), helpTopic(), isPure(), isThreadSafe()
		{
		}
	};

	class FunctionInfoBuilder
	{
		FunctionInfo &_info;

	public:
		FunctionInfoBuilder(FunctionInfo &functionInfo)
			: _info(functionInfo)
		{
		}

		FunctionInfoBuilder& Name(LPCWSTR name)
		{
			_info.name = name;
			return (*this);
		}

		FunctionInfoBuilder& Description(LPCWSTR description)
		{
			_info.description = description;
			return (*this);
		}

		FunctionInfoBuilder& Arg(LPCWSTR name, LPCWSTR description)
		{
			_info.arguments.push_back(NameDescriptionPair(name, description));
			return (*this);
		}

		FunctionInfoBuilder& Category(LPCWSTR category)
		{
			_info.category = category;
			return (*this);
		}

		FunctionInfoBuilder& HelpTopic(LPCWSTR helpTopic)
		{
			_info.helpTopic = helpTopic;
			return (*this);
		}

		FunctionInfoBuilder& Pure()
		{
			_info.isPure = true;
			return (*this);
		}

		FunctionInfoBuilder& Volatile()
		{
			_info.isPure = false;
			return (*this);
		}

		FunctionInfoBuilder& ThreadSafe()
		{
			_info.isThreadSafe = true;
			return (*this);
		}
	};

	template <typename Func> class FunctionInfoFactory;

	template <typename TRet, typename... TArgs>
	struct FunctionInfoFactory<TRet(TArgs...)>
	{
		template <typename T>
		inline static const wchar_t * GetTypeText()
		{
			return ArgumentMarshaler<T>::GetTypeText();
		}

		static FunctionInfo Create(LPCWSTR entryPoint)
		{
			const int NumArgs = sizeof...(TArgs);
			std::array<LPCWSTR, NumArgs + 1> texts = {
				L"Q", // LPXLOPER12
				GetTypeText<TArgs>()...
			};
			std::wstring s;
			for (int i = 0; i <= NumArgs; i++)
			{
				s += texts[i];
			}
			return FunctionInfo(s.c_str(), entryPoint);
		}
	};

	template <typename TRet, typename... TArgs>
	struct FunctionInfoFactory<TRet __stdcall(TArgs...)>
		: public FunctionInfoFactory<TRet(TArgs...)>
	{
	};
}
