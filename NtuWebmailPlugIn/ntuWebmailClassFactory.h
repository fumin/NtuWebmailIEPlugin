#ifndef NTUWEBMAILCLASSFACTORY_H
#define NTUWEBMAILCLASSFACTORY_H

#include "Mainh.h"

struct NtuWebmailClassFactory : public IClassFactory{
public:
	NtuWebmailClassFactory(): m_uRefCount(1) {++g_uDllLockCount;};
	virtual ~NtuWebmailClassFactory() {--g_uDllLockCount;};
    // IUnknown
    virtual ULONG __stdcall AddRef();
    virtual ULONG __stdcall Release();
    virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppv);
    // IClassFactory
    virtual HRESULT __stdcall CreateInstance(IUnknown* pUnknownOuter, REFIID riid, void** ppv);
    virtual HRESULT __stdcall LockServer(BOOL fLock);

protected:
    ULONG m_uRefCount;
};

#endif //NTUWEBMAILCLASSFACTORY_H