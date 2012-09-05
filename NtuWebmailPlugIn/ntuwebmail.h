#ifndef NTUWEBMAIL_H
#define NTUWEBMAIL_H
#include "Mainh.h"
#include <DocObj.h> //IObjectWithSite, IOleCommandTarget 
#include <Exdisp.h> //IWebBrowser2
#include <Exdispid.h>//DWebBrowserEvents2::DocumentComplete Event
#include <comutil.h>
#include <ShObjIdl.h>//IShellBrowser, IShellView
#include <SHLGUID.h>//Service IDs
#include <Mshtml.h>//IHTMLWindow2
#include <string>
#include <fstream>//damn

struct __declspec(uuid("2C5A4662-F642-40E4-9F1A-61852714ECDD")) 
NtuWebmail : public INoAggregationUnknown, IDispatch, IObjectWithSite, IOleCommandTarget {//This is aggregatable inplementation
public:
    //IUnknown interface 
    virtual HRESULT __stdcall QueryInterface(REFIID riid, void **ppObj);
    virtual ULONG   __stdcall AddRef();
    virtual ULONG   __stdcall Release();
	//IDispatch
	virtual HRESULT __stdcall GetTypeInfoCount(UINT* pctinfo);
	virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);
	virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, OLECHAR** rgszNames,UINT cNames, LCID lcid, DISPID* rgDispId);
	virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);
	//IObjectWithSite
	virtual HRESULT __stdcall GetSite(REFIID riid, void **ppvSite);
	virtual HRESULT __stdcall SetSite(IUnknown *pUnkSite);
	//IOleCommandTarget
	virtual HRESULT __stdcall QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText);
	virtual HRESULT __stdcall Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut);
	
	//INoAggregationUnknown
	virtual HRESULT __stdcall QueryInterface_NoAggregation(REFIID riid, void** ppv);
	virtual ULONG __stdcall AddRef_NoAggregation();
	virtual ULONG __stdcall Release_NoAggregation();

	NtuWebmail():m_cRef(1), p_IUnknown_Site(0) {++g_uDllLockCount;}
	NtuWebmail(IUnknown*);//for aggregation
	virtual ~NtuWebmail(){--g_uDllLockCount;}

protected:
	IWebBrowser2 *m_spWebBrowser, *m_NtuWebBrowser;
	IConnectionPoint *m_pConnectionPoint;
	DWORD m_dwCookie;
	IUnknown* p_IUnknown_Site; //Implementation of IObjectWithSite
	IUnknown* m_pUnknownOuter; //provide aggregatable mechanisms
	ULONG m_cRef; // The reference counting variable
	int logout_before;
};

#endif //NTUWEBMAIL_H