////////////////////////////////////////////////////////////////////////////
// xlldef.h -- defines macros to customize the behavior of this library

#pragma once

#include <Windows.h>
#include "XLCALL.H"

#ifndef XLL_NAMESPACE
#define XLL_NAMESPACE xll
#endif

#ifndef XLL_SUPPORT_THREAD_LOCAL
#if WINVER >= _WIN32_WINNT_VISTA
#define XLL_SUPPORT_THREAD_LOCAL 1
#else
#define XLL_SUPPORT_THREAD_LOCAL 0
#endif
#endif