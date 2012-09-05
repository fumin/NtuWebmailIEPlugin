#include "ntuWebmailClassFactory.h"
#include "ntuwebmail.h" //for NtuWebmailClassFactory::CreateInstance

ULONG __stdcall NtuWebmailClassFactory::AddRef(){
    return ++m_uRefCount;               // Increment this object's reference count.
}

ULONG __stdcall NtuWebmailClassFactory::Release(){
	if(--m_uRefCount != 0)
		return m_uRefCount;
	delete this;
	return 0;
}

HRESULT __stdcall NtuWebmailClassFactory::QueryInterface ( REFIID riid, void** ppv ){
	if(ppv == 0)
		return E_POINTER;
	if (InlineIsEqualGUID ( riid, IID_IUnknown ))
        *ppv = (IUnknown*) this;
	else if ( InlineIsEqualGUID ( riid, IID_IClassFactory ))
		*ppv = (IClassFactory*)this;
    else {
        *ppv = 0;
        return E_NOINTERFACE;
    }    
    AddRef ();
    return S_OK;
}

STDMETHODIMP NtuWebmailClassFactory::CreateInstance ( IUnknown* pUnknownOuter,REFIID riid, void** ppv ){
	if(pUnknownOuter != NULL && riid != IID_IUnknown)
		return CLASS_E_NOAGGREGATION;
	NtuWebmail* pInsideCOM = new NtuWebmail(pUnknownOuter);
	if(pInsideCOM == NULL)
		return E_OUTOFMEMORY;
	HRESULT hr = pInsideCOM->QueryInterface_NoAggregation(riid, ppv);
	pInsideCOM->Release_NoAggregation();
	return hr;
}

STDMETHODIMP NtuWebmailClassFactory::LockServer ( BOOL fLock )
{
    // Increase/decrease the DLL ref count, according to the fLock param.

    fLock ? g_uDllLockCount++ : g_uDllLockCount--;

    return S_OK;
}
