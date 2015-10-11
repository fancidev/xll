////////////////////////////////////////////////////////////////////////////
// FunctionInfo.h -- structs to hold UDF description and attributes

#pragma once

#include "xlldef.h"
#include <Windows.h>
#include <string>
#include <vector>
#include <array>
#include "TypeText.h"

// TODO: it makes more sense to put typetext inside this file?

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
		FARPROC entryPoint;
		//LPCWSTR entryPoint;
		LPCWSTR typeText;

		LPCWSTR name;
		LPCWSTR description;
		std::vector<NameDescriptionPair> arguments;
		int macroType; // 0,1,2
		LPCWSTR category;
		LPCWSTR shortcut;
		LPCWSTR helpTopic;

		double registerId;

		//bool isPure;
		//bool isThreadSafe;

		FunctionInfo(FARPROC entryPoint, LPCWSTR typeText)
			: entryPoint(entryPoint), typeText(typeText),
			name(), description(), macroType(1), category(), 
			shortcut(), helpTopic(), registerId()
		{
		}

		static std::vector<FunctionInfo> & registry()
		{
			static std::vector<FunctionInfo> s_functions;
			return s_functions;
		}

		template <int Attributes, typename TRet, typename... TArgs>
		static FunctionInfo& Create(TRet(__stdcall *func)(TArgs...), FARPROC stub = 0)
		{
			const wchar_t *typeText = GetTypeTextImpl<wchar_t, Attributes>(func);
			registry().emplace_back((stub == 0)? (FARPROC)func : stub, typeText);
			return registry().back();
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
	};
}