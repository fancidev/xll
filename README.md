# XLL Connector

XLL Connector is a modern C++ library that makes existing functions available for use in Excel spreadsheets. All you need to do is to add one line for each function you want to export to Excel. A minimal working example is as simple as follows:

```c++
#include "XllAddin.h"

std::wstring ReverseString(const std::wstring &s)
{
	return std::wstring(s.crbegin(), s.crend());
}

EXPORT_XLL_FUNCTION(ReverseString);

BOOL WINAPI DllMain(HANDLE hInstance, ULONG fdwReason, LPVOID lpReserved)
{
	return TRUE;
}
```

## Quick Start

The following assumes you have already built a DLL that contains functions you want to make available for use in Excel. Follow four simple steps to make this happen:
1. Build the XLL Connector static library
2. Link the static library into your DLL
3. Mark the functions you want to export
4. Test the XLL in Excel

### Building XLL Connector

1. Start Visual Studio 2013 or later. This is required because XLL Connector relies on C++11 features to ensure type safety.
2. Download the source code from github. The source code contains two projects. The XllConnector project is the actual library. XllExamples.dll is a sample DLL that you can load into Excel to see it in action.
2. Build the solution XLLConnector.sln. This produces XllConnector.lib in the output directory. 

### Linking XLL Connector 

The static library, XllConnector.lib, must be linked into your DLL to make it available to Excel. There are two alternatives:

* To use XllConnector as a standalone library, add it to the "Additional Dependencies" list in the DLL's linker options.
* To customize XllConnector, add the XllConnector project into your solution, and add it into the "Reference" section of the DLL project.

### Marking Functions to Export

You don't need to export your functions or change any existing code. Just add one line (in a source file) for each function you want to expose to Excel. Suppose you have written a function `Plus` with the following signature:
```
double Plus(double, double);
```
To expose this function to Excel, simply add the following line:
```c++
EXPORT_XLL_FUNCTION(Plus);
```
This automatically generates and exports a wrapper function named `XLPlus` that interfaces with Excel.

You can provide more information about the UDF by the following:
```c++
EXPORT_XLL_FUNCTION(Plus)
  .Description(L"Returns the sum of two numbers.")
  .Arg(L"a", L"first number")
  .Arg(L"b", L"second number")
  .ThreadSafe();
```

### Testing the XLL in Excel

To load the XLL into Excel automatically when you run the project, do the following:

1. Change your Debug settings to launch Excel with the XLL.
2. Hit Ctrl+F5 to run the project.
3. Excel now starts. Create a new workbook. Enter a formula that uses your UDF.

To debug the XLL, do the following:

1. Set a breakpoint in your UDF.
2. Hit F5. If Visual C++ asks you whether to proceed without loading the symbols for EXCEL.EXE, choose Yes.
3. Excel now starts. Create a new workbook. Enter a formula that uses your UDF. When the formula is evaluated, your breakpoint will be triggered and you can debug it.

## Limitations

Being a one-man project, there are a few limitations:

- Visual Studio 2013 or higher is required, because this library makes use of variadic templates to generate xll wrappers.

- Only Excel 2007 and higher is supported as they expose a different interface than prior versions of Excel, which supports more parameters, longer strings, and larger worksheet range. Support for prior versions of Excel may be added later.

- Only 32-bit Excel is supported at the moment. 64-bit support will be added later.

## License

This project uses Apache License V2. See the [LICENSE](LICENSE) for details.
