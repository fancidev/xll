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
// You may override XLL_DEFAULT_VOLATILE/THREADSAFE/HEAVY at
// translation unit level BEFORE #include <XllDef.h> to alter the
// defaults attributes for functions exported from that translation
// unit.
//
// For each individual export, you may set XLL_[NOT_]VOLATILE and/or
// XLL_[NOT_]THREADSAFE to specify the behavior of that UDF.
//

//
// XLL_VOLATILE, XLL_NOT_VOLATILE, XLL_DEFAULT_VOLATILE
//
// Specifies whether Excel should recalculate the function even if
// its input arguments have not changed.
//
// XLL Connector exposes a UDF as volatile by default.
//

#define XLL_VOLATILE         1
#define XLL_NOT_VOLATILE     0x100

#ifndef XLL_DEFAULT_VOLATILE
#define XLL_DEFAULT_VOLATILE 1
#endif

//
// XLL_THREADSAFE, XLL_NOT_THREADSAFE, XLL_DEFAULT_THREADSAFE
//
// Specifies whether Excel may call the function from a background
// thread at the same time of calling other UDF functions. All
// thread-unsafe functions are called from the main thread.
//
// XLL Connector exposes a UDF as thread-unsafe by default.
//

#define XLL_THREADSAFE         2
#define XLL_NOT_THREADSAFE     0x200

#ifndef XLL_DEFAULT_THREADSAFE
#define XLL_DEFAULT_THREADSAFE 0
#endif

// Not implemented
// #define XLL_CLUSTERSAFE     4 
// #define XLL_NOT_CLUSTERSAFE 0x400
// #ifndef XLL_DEFAULT_CLUSTERSAFE
// #define XLL_DEFAULT_CLUSTERSAFE 0
// #endif

//
// XLL_HEAVY, XLL_LIGHT, XLL_DEFAULT_HEAVY
//
// Specifies whether to evaluate the function in Function Wizard. 
// If the function is marked as heavy, XLL Connector returns Empty
// if called from the Function Wizard.
//
// XLL Connector treats UDFs as heavy by default.
//

#define XLL_HEAVY         8
#define XLL_LIGHT         0x800

#ifndef XLL_DEFAULT_HEAVY
#define XLL_DEFAULT_HEAVY 1
#endif

//
// XLL_GENERATE_WRAPPER_STUB, XLL_WRAPPER_STUB_PREFIX
//
// Controls the name of UDF wrappers Xll Connector export from
// the DLL.
//
// If XLL_GENERATE_WRAPPER_STUB is defined to a non-zero value,
// Xll Connector generates a stub function for each UDF wrapper
// and exports this stub from the DLL. The name of the stub is
// XLL_WRAPPER_STUB_PREFIX followed the Excel-visible name of
// the UDF. Xll Connector also registers these stubs by name to
// Excel; if stubs are not generated, Xll Connector registers
// wrappers by ordinal.
//
// Generating a stub may be useful if you want to debug the wrapper
// or to hide the signature text of your UDF. On the downside, it
// incurs an extra indirect jump instruction for each UDF call.
//

#ifndef XLL_GENERATE_WRAPPER_STUB
#define XLL_GENERATE_WRAPPER_STUB 1
#endif

#ifndef XLL_WRAPPER_STUB_PREFIX
#define XLL_WRAPPER_STUB_PREFIX XL12
#endif

//
// ALL THE FOLLOWING ARE IMPLEMENTATION DETAILS THAT YOU SHOULDN'T ALTER.
//

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
