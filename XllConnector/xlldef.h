////////////////////////////////////////////////////////////////////////////
// xlldef.h -- macros that customize the behavior of XLL Connector

#pragma once

#include <Windows.h>
#include "XLCALL.H"

#ifndef XLL_NAMESPACE
#define XLL_NAMESPACE xll
#endif

#define XLL_BEGIN_NAMEPSACE namespace XLL_NAMESPACE {
#define XLL_END_NAMESPACE }

#ifndef XLL_SUPPORT_THREAD_LOCAL
#if WINVER >= _WIN32_WINNT_VISTA
#define XLL_SUPPORT_THREAD_LOCAL 1
#else
#define XLL_SUPPORT_THREAD_LOCAL 0
#endif
#endif

// 
// XLL_MAX_ARG_COUNT
//
// Maximum number of UDF arguments supported in Excel 2007 and later.
// This is a limit imposed by Excel; there is no point changing this
// value.
//

#define XLL_MAX_ARG_COUNT 245
