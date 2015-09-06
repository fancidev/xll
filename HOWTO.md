# Using XLL Connector

This document explains some advanced topis in using XLL Connector.

## Parameter Marshalling

Only certain types of arguments are supported. As a rule of thumb, most "value-type" arguments are supported. In particular, user-defined structures are not supported because Excel doesn't understand them. Vectors and matrices are supported by the Windows `VARIANT` structure.

All arguments and return values are passed by value. For large matrices, this incurs some runtime overhead in copying the buffer. However, this choice is made to ensure memory safety.

## Exception Handling

XLL Connector handles C++ exceptions. If an exception is thrown by your code, XLL Connector silently catches the exception and returns #VALUE! to Excel.

XLL Connector does not handle SEH exceptions (such as division by zero or access violation) unless you explicitly set the /EHa compiler switch in the DLL's property page. Unhandled SEH exceptions will cause Excel to crash.

## Volatile Functions

All functions are assumed to be volatile and registered as such unless you explicitly specify otherwise using `EXPORT_XLL_FUNCTION(...).Pure()`.

## Thread Safety

All functions are assumed to be thread-unsafe and registered as such unless you explicitly state otherwise using `EXPORT_XLL_FUNCTION(...).ThreadSafe()`.

## Design

Excel supports calling user-defined functions (UDFs) defined in a dll. However, some boilerplate code is needed to register the UDFs and to marshal parameters and return values. There are several ways to do this:

* Manually write the necessary boilerplate code. Advantage: full control. Drawback: tedious and error prone.

* Use an external script to generate the boilerplate code as a pre-build step. This is the approach taken by XLW (http://xlw.sourceforge.net/). Advantage: flexible, can make use of comments, integrate with documentation, etc. Drawback: need to format source code according to generator, not transparent, more involved set-up.

* Use C++ magic to generate the boilerplate code. This is the approach taken by the Excel xll add-in library (http://xll.codeplex.com/). Advantage: simple and transparent. Disadvantage: somewhat verbose.

* For .NET programs, exposing the functions are much easier because its reflection support. See for example Excel-DNA (http://excel-dna.net/). Advantage: almost seemless integration. Disadvantage: only supports .NET; can be a mess when mixing .NET 2.0 and .NET 4.0 assemblies.

This project provides a simple way to generate such boilerplate code using C++ templates. It aims to be as simple to set up and use as possible. Just 4 files need to be added to your existing project, and just a one-liner is needed to expose each of your existing functions. Nothing in your existing source file has to be changed. This makes it particularly suitable for quick starts where it does not make much sense to spend much effort in maintaining the XLL interface.

## Background

It is useful to know how an XLL works to make your work smooth.

An XLL is a DLL that exports certain functions to support usage in Excel. The high-level workflow is as follows:

1. The Excel executable exports a procedure named `MdCallBack12`. All calls from the XLL to Excel goes through this point. 

2. When Excel loads an XLL, it calls `xlAutoOpen`, a routine exported by the XLL.

3. In this routine, it calls the Excel entry point to register each UDF available for use in Excel. The registration passes in the export symbol of the function in the DLL, function and argument descriptions, and the types of argument and return values.

4. Each UDF is a regular C++ function except that its parameters must be understood by Excel. Primitive types such as `double` and `const char *` are naturally supported. For more complex types such as arrays, an Excel-specific Variant-like structure is used.

XLL Connector helps you with all the above steps. You only have to write the functions and declare each function to export (as well as optionally provide their descriptions).
