#ifndef MAIN_H
#define MAIN_H

#include <Windows.h>
#include <tchar.h>

extern HINSTANCE g_hinstThisDll;        // Our DLL's module handle
extern UINT      g_uDllLockCount;       // # of COM objects in existence
extern TCHAR     szNtuWebmailGUID[];

struct INoAggregationUnknown{
	virtual HRESULT __stdcall QueryInterface_NoAggregation(REFIID riid, void** ppv)=0;
	virtual ULONG __stdcall AddRef_NoAggregation()=0;
	virtual ULONG __stdcall Release_NoAggregation()=0;
};

#endif //MAIN_H