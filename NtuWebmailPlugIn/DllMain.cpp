#include "Mainh.h"
#include "ntuwebmail.h"
#include "ntuWebmailClassFactory.h"

HINSTANCE g_hinstThisDll = NULL;        // Our DLL's module handle
UINT      g_uDllLockCount = 0;          // # of COM objects in existence
TCHAR     szNtuWebmailGUID[] = TEXT("2C5A4662-F642-40E4-9F1A-61852714ECDD");

BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, void* pv){
	g_hinstThisDll = hInstance;
	return TRUE;
}

HRESULT __stdcall DllGetClassObject(REFCLSID clsid, REFIID riid, void** ppv){
	if(clsid != __uuidof(NtuWebmail))
		return CLASS_E_CLASSNOTAVAILABLE;
	NtuWebmailClassFactory* pFactory = new NtuWebmailClassFactory;
	if(pFactory == NULL)
		return E_OUTOFMEMORY;
	// riid is probably IID_IClassFactory.
	HRESULT hr = pFactory->QueryInterface(riid, ppv);
	pFactory->Release();
	return hr;
}

HRESULT __stdcall DllCanUnloadNow(){
	if(g_uDllLockCount == 0)
		return S_OK;
	else
		return S_FALSE;
	__debugbreak();
}

HRESULT __stdcall DllRegisterServer(){
	HKEY  hCLSIDKey = NULL, hInProcSvrKey = NULL, hAGKey = NULL;
	LONG  lRet;
	TCHAR szModulePath [MAX_PATH];
	TCHAR szClassDescription[] = TEXT("NtuWebmail class");
	TCHAR szThreadingModel[]   = TEXT("Apartment");

	TCHAR clsidslash[45];
	ZeroMemory(clsidslash, 45*sizeof(TCHAR));
	wmemcpy(clsidslash, TEXT("CLSID\\{"), 7);
	wmemcpy(clsidslash+7, szNtuWebmailGUID, 36);
	clsidslash[43] = TEXT('}');

__try
    {
    // Create a key under CLSID for our COM server.

    lRet = RegCreateKeyEx ( HKEY_CLASSES_ROOT, clsidslash/*"CLSID\\{7D51904E-1645-4a8c-BDE0-0F4A44FC38C4}"*/,
                            0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY,
                            NULL, &hCLSIDKey, NULL );
    
    if ( ERROR_SUCCESS != lRet ) 
        return HRESULT_FROM_WIN32(lRet);

    // The default value of the key is a human-readable description of the coclass.

    lRet = RegSetValueEx ( hCLSIDKey, NULL, 0, REG_SZ, (const BYTE*) szClassDescription,
                           sizeof(szClassDescription) );

    if ( ERROR_SUCCESS != lRet )
        return HRESULT_FROM_WIN32(lRet);
    
    // Create the InProcServer32 key, which holds info about our coclass.

    lRet = RegCreateKeyEx ( hCLSIDKey, _T("InProcServer32"), 0, NULL, REG_OPTION_NON_VOLATILE,
                            KEY_SET_VALUE, NULL, &hInProcSvrKey, NULL );

    if ( ERROR_SUCCESS != lRet )
        return HRESULT_FROM_WIN32(lRet);

    // The default value of the InProcServer32 key holds the full path to our DLL.

    GetModuleFileName ( g_hinstThisDll, szModulePath, MAX_PATH );

    lRet = RegSetValueEx ( hInProcSvrKey, NULL, 0, REG_SZ, (const BYTE*) szModulePath, 
                           sizeof(TCHAR) * (lstrlen(szModulePath)+1) );

    if ( ERROR_SUCCESS != lRet )
        return HRESULT_FROM_WIN32(lRet);

    // The ThreadingModel value tells COM how it should handle threads in our DLL.
    // The concept of apartments is beyond the scope of this article, but for
    // simple, single-threaded DLLs, use Apartment.

    lRet = RegSetValueEx ( hInProcSvrKey, _T("ThreadingModel"), 0, REG_SZ,
                           (const BYTE*) szThreadingModel, 
                           sizeof(szThreadingModel) );

    if ( ERROR_SUCCESS != lRet )
        return HRESULT_FROM_WIN32(lRet);

	
	//Fumin's NtuWebmailPlugIn IE registration
	//1. Create {another guid} under HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Extensions
	lRet = RegCreateKeyEx ( HKEY_LOCAL_MACHINE, TEXT("Software\\Microsoft\\Internet Explorer\\Extensions\\{4F45CB9F-BA41-4352-8CA8-38B12C8A9CA0}"),
                            0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY,
                            NULL, &hAGKey, NULL );
	//Set ButtonText
	TCHAR szButtonText[]   = TEXT("Go to Ntu Webmail");
	lRet = RegSetValueEx ( hAGKey, _T("ButtonText"), 0, REG_SZ,
                           (const BYTE*) szButtonText, 
                           sizeof(szButtonText) );
	//Set CLSID_Shell_ToolbarExtExec = {1FBA04EE-3024-11d2-8F1F-0000F87ABD16}
	TCHAR szCLSID_Shell_ToolbarExtExec[] = TEXT("{1FBA04EE-3024-11d2-8F1F-0000F87ABD16}");
	lRet = RegSetValueEx ( hAGKey, _T("CLSID"), 0, REG_SZ,
                           (const BYTE*) szCLSID_Shell_ToolbarExtExec, 
                           sizeof(szCLSID_Shell_ToolbarExtExec) );
	TCHAR CLSIDExtension[100]; ZeroMemory(CLSIDExtension, 100*sizeof(TCHAR)); 
	*CLSIDExtension = TEXT('{');
	wmemcpy(CLSIDExtension+1, szNtuWebmailGUID, 36);
	*(CLSIDExtension+37) = TEXT('}');
	lRet = RegSetValueEx ( hAGKey, _T("CLSIDExtension"), 0, REG_SZ,
                           (const BYTE*) CLSIDExtension, 
                           39*sizeof(TCHAR) );
	TCHAR szHotIcon[] = TEXT("D:\\Desktop\\Exchange Student\\NtuWebmailPlugIn\\Debug\\NtuWebmailPlugIn.dll,101");
	lRet = RegSetValueEx ( hAGKey, _T("HotIcon"), 0, REG_SZ,
                           (const BYTE*) szHotIcon, 
                           sizeof(szHotIcon) );
	lRet = RegSetValueEx ( hAGKey, _T("Icon"), 0, REG_SZ,
                           (const BYTE*) szHotIcon, 
                           sizeof(szHotIcon) );
	//Completed Fumin's NtuWebmailPlugIn IE registration

	MessageBox(0, clsidslash, TEXT("REVCS32"), 0);
    }   
__finally
    {
    if ( NULL != hCLSIDKey )
        RegCloseKey ( hCLSIDKey );

    if ( NULL != hInProcSvrKey )
        RegCloseKey ( hInProcSvrKey );

	if ( NULL != hAGKey )
        RegCloseKey ( hAGKey );
    }

    return S_OK;
}

HRESULT __stdcall DllUnregisterServer()
{
    // Delete our registry entries.  Note that you must delete from the deepest
    // key and work upwards, because on NT/2K, RegDeleteKey() doesn't delete 
    // keys that have subkeys on NT/2K.

	TCHAR clsidslash[45];
	ZeroMemory(clsidslash, 45*sizeof(TCHAR));
	wmemcpy(clsidslash, TEXT("CLSID\\{"), 7);
	wmemcpy(clsidslash+7, szNtuWebmailGUID, 36);
	clsidslash[43] = TEXT('}');

	TCHAR cl[60];
	ZeroMemory(cl, 60*sizeof(TCHAR));
	wmemcpy(cl, clsidslash, 44);
	wmemcpy(cl+44, TEXT("\\InProcServer32"), 15);

    RegDeleteKey ( HKEY_CLASSES_ROOT, cl/*"CLSID\\{7D51904E-1645-4a8c-BDE0-0F4A44FC38C4}\\InProcServer32"*/);
    RegDeleteKey ( HKEY_CLASSES_ROOT, clsidslash/*"CLSID\\{7D51904E-1645-4a8c-BDE0-0F4A44FC38C4}"*/);
	//Delete NtuWebmailPlugIn Extension
	RegDeleteKey ( HKEY_LOCAL_MACHINE, TEXT("Software\\Microsoft\\Internet Explorer\\Extensions\\{4F45CB9F-BA41-4352-8CA8-38B12C8A9CA0}"));
    return S_OK;
}