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

#ifndef XLL_WRAPPER_PREFIX
#define XLL_WRAPPER_PREFIX XL12
#endif

#ifndef XLL_SUPPORT_THREAD_LOCAL
#if WINVER >= _WIN32_WINNT_VISTA
#define XLL_SUPPORT_THREAD_LOCAL 1
#else
#define XLL_SUPPORT_THREAD_LOCAL 0
#endif
#endif