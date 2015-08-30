# Using XLL Add-In Helper

This document explains how to make your C++ functions available for use in Excel.

## Quick Start

To build an XLL, do the following:

1. Start Visual Studio 2013 or later (required).
2. Build a regular DLL that contains your C++ functions. You don't need to export the functions.
3. Add these files into your source: XllAddin.cpp, XllAddin.h, XLCALL.cpp, XLCALL.H.
4. For each function you want to expose to Excel, add the following line in a source file:
`EXPORT_XLL_FUNCTION(NameOfYourFunction);`
5. Build the DLL.

To debug the XLL, do the following:

1. Set a breakpoint in your UDF.
2. Change your Debug settings to launch Excel with the XLL.
3. Hit F5. If Visual C++ asks you whether to proceed without loading the symbols for EXCEL.EXE, choose Yes.
4. Excel now starts. Create a new workbook. Enter a formula that uses your UDF. When the formula is evaluated, your breakpoint will be triggered and you can debug it.

To provide descriptions for functions and arguments, use the following syntax:
```c++
EXPORT_XLL_FUNCTION(Plus)
  .Description(L"Returns the sum of two numbers.")
  .Arg(L"a", L"first number")
  .Arg(L"b", L"second number");
```

## Background

It is useful to know how an XLL works to make your work smooth.

An XLL is a DLL that exports certain functions to support usage in Excel. The high-level workflow is as follows:

1. The Excel executable exports a procedure named `MdCallBack12`. All calls from the XLL to Excel goes through this point. 

2. When Excel loads an XLL, it calls `xlAutoOpen`, a routine exported by the XLL.

3. In this routine, it calls the Excel entry point to register each UDF available for use in Excel. The registration passes in the export symbol of the function in the DLL, function and argument descriptions, and the types of argument and return values.

4. Each UDF is a regular C++ function except that its parameters must be understood by Excel. Primitive types such as `double` and `const char *` are naturally supported. For more complex types such as arrays, an Excel-specific Variant-like structure is used.

This XLL Addin Helper helps you with all the above steps. You only have to write the functions and declare each function to export (as well as optionally provide their descriptions).

## Parameter Marshalling

Only certain types of arguments are supported. As a rule of thumb, most "value-type" arguments are supported. In particular, user-defined structures are not supported because Excel doesn't understand them. Vectors and matrices are supported by the Windows `VARIANT` structure.

All arguments and return values are passed by value. For large matrices, this incurs some runtime overhead in copying the buffer. However, this choice is made to ensure memory safety.

## Exception Handling

Currently the library does not handle exceptions. If your code throws an exception, Excel will crash.

If memory allocation fails at the wrapper level, it throws `std::bad_alloc` and crashes Excel.

## Volatile Functions

All functions are assumed to be volatile and registered as such unless you explicitly state otherwise using `EXPORT_XLL_FUNCTION(...).Pure()`.

## Thread Safety

All functions are assumed to be thread-unsafe and registered as such unless you explicitly state otherwise using `EXPORT_XLL_FUNCTION(...).ThreadSafe()`.
