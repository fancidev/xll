#include "XllAddin.h"

BOOL WINAPI DllMain(HANDLE hInstance, ULONG fdwReason, LPVOID lpReserved)
{
	return TRUE;
}

//template <typename Func> void XLMyFunc();
//template <typename TRet, typename... TArgs>
//TRet XLMyFunc<TRet(TArgs...)> XLMyFunc(TArgs... args)
//{
//	return TRet();
//}



//
//template <typename Func, Func *func, typename TRet, typename... TArgs>
//struct XLWrapper < Func, func, TRet(TArgs...) >
//{
//	static LPXLOPER12 __stdcall Call(typename ArgumentWrapper<TArgs>::wrapped_type... args)
//	{
//		try
//		{
//			LPXLOPER12 pvRetVal = getReturnValue();
//			XLOPER12_Create(pvRetVal, func(ArgumentWrapper<TArgs>::unwrap(args)...));
//			return pvRetVal;
//		}
//		catch (const std::exception &)
//		{
//			// todo: report exception
//		}
//		catch (...)
//		{
//			// todo: report exception
//		}
//		return const_cast<ExcelVariant*>(&ExcelVariant::ErrValue);
//	}
//};