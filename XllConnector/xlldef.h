////////////////////////////////////////////////////////////////////////////
// xlldef.h -- macros that customize the behavior of XLL Connector

#pragma once

#include <Windows.h>
#include "XLCALL.H"

#ifndef XLL_NAMESPACE
#define XLL_NAMESPACE xll
#endif

//
// Function Attributes
//
// Defines attributes of a UDF, which determines how the wrapper is
// generated and how it is called by Excel.
//
// You may override XLL_DEFAULT_VOLATILE and XLL_DEFAULT_THREAD_SAFE
// at translation unit level BEFORE #include <XllDef.h> to alter the
// defaults for functions exported from that translation unit.
//
// For each individual export, you may set XLL_[NOT_]VOLATILE and/or
// XLL_[NOT_]THREAD_SAFE to specify the behavior of that UDF.
//

#define XLL_VOLATILE       1
#define XLL_THREADSAFE     2
#define XLL_NOT_VOLATILE   0x10
#define XLL_NOT_THREADSAFE 0x20

#ifndef XLL_DEFAULT_VOLATILE
#define XLL_DEFAULT_VOLATILE 1
#endif

#ifndef XLL_DEFAULT_THREADSAFE
#define XLL_DEFAULT_THREADSAFE 0
#endif

//
// ALL THE FOLLOWING ARE IMPLEMENTATION DETAILS THAT YOU SHOULDN'T ALTER.
//

// to be removed

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

//
// XLL_NOEXCEPT
//
// Indicates that a function does not throw C++ exception.
//

#define XLL_NOEXCEPT throw()
