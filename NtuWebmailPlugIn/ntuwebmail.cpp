// The while(!m_NtuWebBrowser) loop in NtuWebmail::Exec() is really driving me crazy, because
// once immediately firing IWebBrowser::Navigate(), IE hasn't have to register the new tab
// as a shellwindow (a window that can be reached by IShellWindows, which is an object upon 
// creation will have a SYSTEM wide <including both internet and windows explorer> access to all shell windows!)
// Therefore my way to circumvent this timing complication is to while loop until IE registers 
// the new tab so that I can retrieve it and its interfaces.

// Another bustard complication is when I place a SINGLE __debugbreak() inside NtuWebmail::Invoke(),
// perhaps because IE runs in several confounded different threads, IE tends to kick into my memory layout
// igniting a damn heap corruption! Therefore the correct way to stop inside NtuWebmail::Invoke() is to
// put __debugbreak() ahead say in NtuWebmail::Exec(), and then put a toggle point inside NtuWebmail::Invoke()

#include "ntuwebmail.h"

NtuWebmail::NtuWebmail(IUnknown* pUnknownOuter) : m_cRef(1), logout_before(0){
	p_IUnknown_Site = 0;
	g_uDllLockCount++;
	if(pUnknownOuter != NULL)
		// We're being aggregated.
		// No AddRef!
		m_pUnknownOuter = pUnknownOuter;
	else
		// Standard usage
		m_pUnknownOuter = (IUnknown*)(INoAggregationUnknown*)this;
}

ULONG __stdcall NtuWebmail::AddRef(){
	return m_pUnknownOuter->AddRef();
}

ULONG __stdcall NtuWebmail::Release(){
	return m_pUnknownOuter->Release();
}

HRESULT __stdcall NtuWebmail::QueryInterface(REFIID riid, void** ppv){
	return m_pUnknownOuter->QueryInterface(riid, ppv);
}

HRESULT __stdcall NtuWebmail::QueryInterface_NoAggregation(REFIID riid, void** ppv){
	if(ppv == 0)
		return E_POINTER;
	if (InlineIsEqualGUID ( riid, IID_IUnknown ))
        *ppv = (INoAggregationUnknown*) this;
	else if (InlineIsEqualGUID ( riid, __uuidof(NtuWebmail) ))
		*ppv = (NtuWebmail*)this;
	else if (InlineIsEqualGUID ( riid, IID_IObjectWithSite ))
        *ppv = (IObjectWithSite*) this;
	else if (InlineIsEqualGUID ( riid, IID_IOleCommandTarget ))
        *ppv = (IOleCommandTarget*) this;
	else if (InlineIsEqualGUID ( riid, IID_IDispatch )){
		*ppv = (IDispatch*)this;
	}
    else {
        *ppv = 0;
        return E_NOINTERFACE;
    }    
	((IUnknown*)(*ppv))->AddRef();
	return S_OK;
}

ULONG __stdcall NtuWebmail::AddRef_NoAggregation(){
	return ++m_cRef;
}

ULONG __stdcall NtuWebmail::Release_NoAggregation(){
	if(--m_cRef != 0)
		return m_cRef;
	m_spWebBrowser->Release();
	if(m_NtuWebBrowser)m_NtuWebBrowser->Release();
	p_IUnknown_Site->Release();
	m_pConnectionPoint->Release();
	delete this;
	return 0;
}

HRESULT __stdcall NtuWebmail::SetSite(IUnknown *pUnkSite){
	if (pUnkSite != NULL){
        // Cache the pointer to IWebBrowser2
        IServiceProvider* sp = 0;
		HRESULT hr = pUnkSite->QueryInterface(IID_IServiceProvider, (void**)&sp);
        hr = sp->QueryService(IID_IWebBrowserApp, IID_IWebBrowser2, (void**)&m_spWebBrowser);
		sp->Release();

		//Usual SetSite Procedures
		if(p_IUnknown_Site == 0){
			p_IUnknown_Site = pUnkSite;
			p_IUnknown_Site->AddRef();
		}
		else{
			pUnkSite->AddRef();
			p_IUnknown_Site->Release();
			p_IUnknown_Site = pUnkSite;
		}
    }
    else{
        // Release pointer
        m_spWebBrowser->Release();
    }
	return S_OK;
}

HRESULT __stdcall NtuWebmail::GetSite(REFIID riid, void **ppvSite){
	if(p_IUnknown_Site == 0){
		*ppvSite = 0;
		return E_FAIL;
	}
	else{
		if(p_IUnknown_Site->QueryInterface(riid, ppvSite) == S_OK){
			p_IUnknown_Site->AddRef();
			return S_OK;
		}
		else{
			*ppvSite = 0;
			return E_NOINTERFACE;
		}
	}
}

HRESULT __stdcall NtuWebmail::QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText){
	if(prgCmds == 0)
		return E_POINTER;

	prgCmds->cmdf = OLECMDF_SUPPORTED;
	ZeroMemory(pCmdText->rgwz, pCmdText->cwBuf);
	int NcwActual, ScwActual;
	switch(pCmdText->cmdtextf){
	case OLECMDTEXTF_NAME:
		NcwActual = (pCmdText->cwBuf > 10) ? 10 : pCmdText->cwBuf - 1;
		wmemcpy(pCmdText->rgwz, TEXT("NTU Webmail"), NcwActual);
		pCmdText->cwActual = NcwActual;
		break;
	case OLECMDTEXTF_STATUS:
		ScwActual = (pCmdText->cwBuf > 7) ? 7 : pCmdText->cwBuf - 1;
		wmemcpy(pCmdText->rgwz, TEXT("Enabled"), ScwActual);
		pCmdText->cwActual = ScwActual;
		break;
	}
	return  S_OK ;
}

HRESULT CoCreateInstanceAsAdmin(HWND hwnd, REFCLSID rclsid, REFIID riid, DWORD dwClassContext, __out void ** ppv)
{
    BIND_OPTS3 bo;
    WCHAR  wszCLSID[50]; char szCLSID[50];
    WCHAR  wszMonikerName[300];

    StringFromGUID2(rclsid, wszCLSID, sizeof(wszCLSID)/sizeof(wszCLSID[0])); 
	WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, wszCLSID, 50, szCLSID, 50, 0, 0);
	std::string str("Elevation:Administrator!clsid:");
	str += szCLSID; //str += "}";
	MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, str.c_str(), -1, wszMonikerName, 300);
    
    memset(&bo, 0, sizeof(bo));
    bo.cbStruct = sizeof(bo);
    bo.hwnd = hwnd;
    bo.dwClassContext  = dwClassContext;
    return CoGetObject(wszMonikerName, &bo, riid, ppv);
}

HRESULT __stdcall NtuWebmail::Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut){
	BSTR Url = SysAllocString(TEXT("https://wmail1.cc.ntu.edu.tw/imp/login.php")), LocationURL;
	if (m_spWebBrowser != NULL){
		VARIANT flags, noArg;
        flags.vt = VT_I4;
        flags.lVal = navOpenInNewTab;
        noArg.vt = VT_EMPTY;
		HRESULT hr = m_spWebBrowser->Navigate(Url, &flags, &noArg, &noArg, &noArg);
		//IStream* pStm;
		//hr = CoMarshalInterThreadInterfaceInStream(IID_IUnknown, p_IUnknown_Site, &pStm);
		//HANDLE hThread =  (HANDLE)_beginthreadex(0/*Security Attr*/, 0/*Stack Size*/, EnterPW, pStm/*Arglist*/, 0/*suspend flags*/, 0/*ThreadID*/);
        m_NtuWebBrowser = 0; //__debugbreak();
		IShellWindows *up_psw;
		hr = CoCreateInstance(CLSID_ShellWindows, NULL, CLSCTX_ALL, IID_IShellWindows, (void**)&up_psw);
		if(hr == E_ACCESSDENIED){
			MessageBox(0, TEXT("Couldn't activate ShellWindows due to Access denied. \nPlease run IE as Administrator."), TEXT("NtuWebmailPlugIn.dll Error"), 0);
			return E_ACCESSDENIED;
		}
		long up_count, count;
		hr = up_psw->get_Count(&up_count);
		while(!m_NtuWebBrowser || wcscmp(LocationURL, Url) || m_NtuWebBrowser == m_spWebBrowser || up_count + 1 != count){
			IShellWindows *psw; HWND hwnd;
			//Sleep(600);//I don't why I have to sleep, but if don't sleep, Invoke won't be called by IE!
			hr = m_spWebBrowser->get_HWND((SHANDLE_PTR*)&hwnd);
			//hr = CoCreateInstanceAsAdmin(hwnd, CLSID_ShellWindows, IID_IShellWindows, CLSCTX_ALL, (void**)&psw);
		    hr = CoCreateInstance(CLSID_ShellWindows, NULL, CLSCTX_ALL, IID_IShellWindows, (void**)&psw);
		    if(hr == E_ACCESSDENIED){
			    MessageBox(hwnd, TEXT("Couldn't activate ShellWindows due to Access denied. \nPlease run IE as Administrator."), TEXT("NtuWebmailPlugIn.dll Error"), 0);
			    return E_ACCESSDENIED;
		    }
			hr = psw->get_Count(&count);
			VARIANT va2; va2.vt = VT_I4; va2.lVal = count - 1;
			IDispatch* pFolder2;
			hr = psw->Item(va2, &pFolder2);
			hr = pFolder2->QueryInterface(IID_IWebBrowser2, (void**)&m_NtuWebBrowser);
			
			hr = m_NtuWebBrowser->get_LocationURL(&LocationURL);
			
			++m_cRef;//Add a reference to NtuWebmail
			pFolder2->Release();
			psw->Release();
		}

		IConnectionPointContainer* pCPContainer;
		// Step 1: Get a pointer to the connection point container.
		hr = m_NtuWebBrowser->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPContainer);
		if (SUCCEEDED(hr)) {
			// m_pConnectionPoint is defined like this:
			// IConnectionPoint* m_pConnectionPoint;
			// Step 2: Find the connection point.
			hr = pCPContainer->FindConnectionPoint(DIID_DWebBrowserEvents2, &m_pConnectionPoint);
			if (SUCCEEDED(hr)){
				// Step 3: Advise the connection point that you
				// want to sink its events.
				hr = m_pConnectionPoint->Advise((IDispatch*)this, &m_dwCookie);
				if (FAILED(hr))
					MessageBox(NULL, TEXT("Failed to Advise"), TEXT("C++ Event Sink"), MB_OK);
			}
			pCPContainer->Release();
		}
	}
    else
        MessageBox(NULL, TEXT("No Web browser pointer"), TEXT("Oops"), 0); 
	SysFreeString(Url); SysFreeString(LocationURL);
	m_NtuWebBrowser->Refresh();
    return S_OK;
}

HRESULT __stdcall NtuWebmail::GetTypeInfoCount(UINT* pctinfo){
	if (pctinfo == NULL)
      return E_INVALIDARG;
	*pctinfo = 0; //we do not provide type information
	__debugbreak();
	return NOERROR;
}

HRESULT __stdcall NtuWebmail::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo){
	__debugbreak();return E_NOTIMPL;
}

HRESULT __stdcall NtuWebmail::GetIDsOfNames(REFIID riid, OLECHAR** rgszNames,UINT cNames, LCID lcid, DISPID* rgDispId){
	__debugbreak();return E_NOTIMPL;
}

HRESULT __stdcall NtuWebmail::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr){
	BSTR Url, imapuser, pass, r98723077, s5jm8ouxc, loginbutton, tabtitle, CerError; HRESULT hr; IWebBrowser2* test_NtuWB; IHTMLElement *pelUser, *pelPW, *pellb; 
	IHTMLDocument3* pIHTMLDocument3; IDispatch* pIDispatch; IHTMLInputElement *pUser, *pPW, *pLB; IHTMLFormElement* pFE; IHTMLDocument2* pIHTMLDocument2;
	std::ifstream file; TCHAR wszBuffer[100]; char szBuffer[300]; std::string str;
	BSTR LocationURL;
	if(dispIdMember!=DISPID_COMMANDSTATECHANGE /*&& dispIdMember!=DISPID_WINDOWSTATECHANGED*/ && dispIdMember!=DISPID_STATUSTEXTCHANGE && dispIdMember!=DISPID_SETSECURELOCKICON && dispIdMember!=DISPID_SETPHISHINGFILTERSTATUS && dispIdMember!=DISPID_PROGRESSCHANGE && dispIdMember!=DISPID_PROPERTYCHANGE){
	switch(dispIdMember){
	case 113:
	case DISPID_DOCUMENTCOMPLETE:
		if(logout_before == 1)
			break;
		// pDispParams->cArgs == 2, two arguments received
		// ((pDispParams->rgvarg) + 1)->vt  == VT_DISPATCH
		Url = SysAllocString(TEXT("https://wmail1.cc.ntu.edu.tw/imp/login.php")); CerError = SysAllocString(TEXT("Certificate Error: Navigation Blocked"));
		//hr = ((pDispParams->rgvarg) + 1)->pdispVal->QueryInterface(IID_IWebBrowser2, (void**)&test_NtuWB);		
		test_NtuWB = m_NtuWebBrowser;
		hr = test_NtuWB->get_Document(&pIDispatch);
		hr = pIDispatch->QueryInterface(IID_IHTMLDocument2, (void**)&pIHTMLDocument2);
		hr = pIHTMLDocument2->get_title(&tabtitle);
		pIHTMLDocument2->Release();
		hr = m_NtuWebBrowser->get_LocationURL(&LocationURL);
		if(wcscmp(tabtitle, CerError) && !wcscmp(LocationURL, Url)/*!wcscmp(((VARIANT*)(pDispParams->rgvarg->byref))->bstrVal, Url) && test_NtuWB == m_NtuWebBrowser*/){
			GetModuleFileNameA(g_hinstThisDll, szBuffer, 300); 
			str = szBuffer; str.replace(str.end()-3, str.end(), "txt");
			file.open(str);
			if(file.is_open()){
				std::getline(file, str);
			    MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, str.c_str(), -1, wszBuffer, 100);
			    r98723077 = SysAllocString(wszBuffer);
			    ZeroMemory(szBuffer, sizeof(szBuffer));
			    std::getline(file, str);
			    MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, str.c_str(), -1, wszBuffer, 100);
			    s5jm8ouxc = SysAllocString(wszBuffer);
				imapuser = SysAllocString(TEXT("imapuser"));
				pass = SysAllocString(TEXT("pass"));
				loginbutton = SysAllocString(TEXT("loginbutton"));	

				hr = pIDispatch->QueryInterface(IID_IHTMLDocument3, (void**)&pIHTMLDocument3);
			    hr = pIHTMLDocument3->getElementById(imapuser, &pelUser);
			    hr = pelUser->QueryInterface(IID_IHTMLInputElement, (void**)&pUser);
			    hr = pUser->put_value(r98723077);
			    hr = pIHTMLDocument3->getElementById(pass, &pelPW);
			    hr = pelPW->QueryInterface(IID_IHTMLInputElement, (void**)&pPW);
			    hr = pPW->put_value(s5jm8ouxc);
			    hr = pIHTMLDocument3->getElementById(loginbutton, &pellb);
			    hr = pellb->QueryInterface(IID_IHTMLInputElement, (void**)&pLB);
			    hr = pLB->get_form(&pFE);
			    hr = pFE->submit();
				SysFreeString(r98723077); SysFreeString(s5jm8ouxc); SysFreeString(imapuser);  SysFreeString(pass); SysFreeString(loginbutton);
			    file.close();
			}
			else{
				MessageBox(0, TEXT("Can't open NtuWebmailPlugIn.txt"), TEXT("Error"), 0);
				return E_FAIL;
			}// file.open()	
		}//if (test_NtuWB == m_NtuWebBrowser && url == test_NtuWB.url())
		SysFreeString(Url); SysFreeString(tabtitle); SysFreeString(CerError); SysFreeString(LocationURL);
		logout_before = 1;
		break;
	case DISPID_ONQUIT:
		// Step 5: Unadvise. Note that m_pConnectionPoint should
		// be released in the destructor for CSomeClass.
		//if (m_pConnectionPoint){
			//HRESULT hr = m_pConnectionPoint->Unadvise(m_dwCookie);
			//if (FAILED(hr))
			//	MessageBox(NULL, TEXT("Failed to Unadvise"), TEXT("C++ Event Sink"), MB_OK);
		//}
		//m_pConnectionPoint->Release();
		--m_cRef;
		break;
	}}
	return S_OK;
}