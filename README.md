# xll

This project makes it extremely easy to make existing C++ functions available for use in Excel spreadsheets. All you need to do is to add two source files and two header files to your project, and then add a one-liner to export each existing C++ function. A minimal working example is as simple as follows:

```c++
#include "XllAddin.h"

double Square(double x)
{
	return x * x;
}

EXPORT_XLL_FUNCTION(Square);

BOOL WINAPI DllMain(HANDLE hInstance, ULONG fdwReason, LPVOID lpReserved)
{
	return TRUE;
}
```

## Introduction

Excel supports calling user-defined functions (UDFs) defined in a dll. However, some boilerplate code is needed to register the UDFs and to marshal parameters and return values. There are several ways to do this:

* Manually write the necessary boilerplate code. Advantage: full control. Drawback: tedious and error prone.

* Use an external script to generate the boilerplate code as a pre-build step. This is the approach taken by XLW (http://xlw.sourceforge.net/). Advantage: flexible, can make use of comments, integrate with documentation, etc. Drawback: need to format source code according to generator, not transparent, more involved set-up.

* Use C++ magic to generate the boilerplate code. This is the approach taken by the Excel xll add-in library (http://xll.codeplex.com/). Advantage: simple and transparent. Disadvantage: somewhat verbose.

* For .NET programs, exposing the functions are much easier because its reflection support. See for example Excel-DNA (http://excel-dna.net/). Advantage: almost seemless integration. Disadvantage: only supports .NET; can be a mess when mixing .NET 2.0 and .NET 4.0 assemblies.

This project provides a simple way to generate such boilerplate code using C++ templates. It aims to be as simple to set up and use as possible. Just 4 files need to be added to your existing project, and just a one-liner is needed to expose each of your existing functions. Nothing in your existing source file has to be changed. This makes it particularly suitable for quick starts where it does not make much sense to spend much effort in maintaining the XLL interface.

## Design

(to be written)

## Limitations

Being a one-man project, there are a few limitations:

- Visual Studio 2013 or higher is required, because this library makes use of variadic templates to generate xll wrappers.

- Only Excel 2007 and higher is supported as they expose a different interface than prior versions of Excel, which supports more parameters, longer strings, and larger worksheet range. Support for prior versions of Excel may be added later.

- Only 32-bit Excel is supported at the moment. 64-bit support will be added later.

You are more than welcome to help!

## License

This project uses Apache License V2. See the LICENSE file for details.
