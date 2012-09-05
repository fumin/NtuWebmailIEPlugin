#include <Windows.h>
#include <tchar.h>
#include "testmidl.h"
#include <iostream>
#include <string>

#define COMRet(hr) if(hr != S_OK) {__debugbreak(); return hr;}

void TExpandPointer(TYPEDESC typedesc, int indentation);// function to show the type of a paramater of a function, ex. VT_PTR
int VarTranslate(int VT_, TCHAR* szBuffer, int Blength);//Translate the enum VT_BOOL to the string "VT_BOOL"
void wParamFlagsTranslate(const USHORT wParamFlags, TCHAR* szBuffer, const int Blength);//Translates the enum PARAMFLAG_FRETVAL to the string "retval"

HRESULT PrintTInfo(ITypeInfo* pTinfo, int indentation){
	TYPEATTR* pTypeAttr;
	HRESULT hr = pTinfo->GetTypeAttr(&pTypeAttr); COMRet(hr);
	for(int inden = 0; inden != indentation; ++inden) std::wcout<<TEXT("    "); 
	LPOLESTR guid_str;
	hr = StringFromCLSID(pTypeAttr->guid, &guid_str);
	std::wcout<<guid_str<<std::endl;

	// Inherited Interfaces, therefore we recursively call PrintInfo
	for(int i = 0; i != pTypeAttr->cImplTypes; ++i){
		HREFTYPE RefType;
		hr = pTinfo->GetRefTypeOfImplType(i, &RefType); COMRet(hr);
		ITypeInfo* pImplTinfo;
		hr = pTinfo->GetRefTypeInfo(RefType, &pImplTinfo); COMRet(hr);
		hr = PrintTInfo(pImplTinfo, indentation + 1); if(hr != S_OK && hr != TYPE_E_BADMODULEKIND) return hr; // Because this ITypeInfo is retrieved
		pImplTinfo->Release();                                                                                // directly from a .tlb file instead
		                                                                                                      // of a .dll, AddressofMember fails
	}

	//member functions
	for(int i = 0; i != pTypeAttr->cFuncs; ++i){
		FUNCDESC* pFuncDesc;
		hr = pTinfo->GetFuncDesc(i, &pFuncDesc); COMRet(hr);
		const UINT cMaxNames = 10; UINT cNames;
		BSTR rgBstrNames[cMaxNames];
		hr = pTinfo->GetNames(pFuncDesc->memid, rgBstrNames, cMaxNames, &cNames); COMRet(hr);
		void* pv;
		hr = pTinfo->AddressOfMember(pFuncDesc->memid, pFuncDesc->invkind, &pv); if(hr != S_OK && hr != TYPE_E_BADMODULEKIND) return hr;
		for(int inden = 0; inden != indentation; ++inden) std::wcout<<TEXT("    ");
		std::wcout<<TEXT("Func memid = ")<<pFuncDesc->memid<<TEXT(" Name = ")<<*rgBstrNames<<TEXT(" DllAddress = ")<<pv<<std::endl;
		for(int j = 0; j != pFuncDesc->cParams; ++j){
			TCHAR szBuffer[30];
			wParamFlagsTranslate(pFuncDesc->lprgelemdescParam->paramdesc.wParamFlags, szBuffer, 30);
			for(int inden = 0; inden != indentation; ++inden) std::wcout<<TEXT("    ");
			std::wcout<<TEXT(" ")<<szBuffer<<TEXT(" ")<<*(rgBstrNames+j+1)<<TEXT(": ")<<std::endl;
			TExpandPointer(pFuncDesc->lprgelemdescParam->tdesc, 2);
		}
		pTinfo->ReleaseFuncDesc(pFuncDesc);
	}

	//member variables
	for(int i = 0; i != pTypeAttr->cVars; ++i){
		VARDESC* pVarDesc;
		hr = pTinfo->GetVarDesc(i, &pVarDesc); COMRet(hr);
		for(int inden = 0; inden != indentation; ++inden) std::wcout<<TEXT("    ");
		std::wcout<<TEXT("Var memid = ")<<pVarDesc->memid<<TEXT("Varkind = ")<<pVarDesc->varkind<<std::endl;
		TCHAR szBuffer[30];
		wParamFlagsTranslate(pVarDesc->elemdescVar.paramdesc.wParamFlags, szBuffer, 30);
		for(int inden = 0; inden != indentation; ++inden) std::wcout<<TEXT("    ");
		std::wcout<<"  Variable wParamFlags: "<<szBuffer<<std::endl;
		TExpandPointer(pVarDesc->elemdescVar.tdesc, 1);
		pTinfo->ReleaseVarDesc(pVarDesc);
	}
	pTinfo->ReleaseTypeAttr(pTypeAttr);
	return hr;
}

int _tmain(int argc, _TCHAR* argv[]){
	const OLECHAR szFile[] = TEXT("testmidl.tlb");//TEXT("D:\\Desktop\\COM_Type_Library_Viewer\\COM_Type_Library_Viewer\\testmidl.tlb");
	ITypeLib* p_tlib;  
	ITypeInfo* pTinfo;
	HRESULT hr = LoadTypeLib(szFile, &p_tlib); 
	//hr = p_tlib->GetTypeInfoOfGuid(IID_ISum, &pTinfo);//Succeeded
	//hr = p_tlib->GetTypeInfoOfGuid(LIBID_Component, &pTinfo);
	hr = p_tlib->GetTypeInfoOfGuid(CLSID_InsideCOM, &pTinfo);//Succeeded!!
	hr = PrintTInfo(pTinfo, 0);
	pTinfo->Release();
	p_tlib->Release();
	std::getchar();
	return 1;
}

//Actually this function is more appropriate to call the function that describes a particular parameter
//in a given function. It does its work by chasing down the pointers inside the struct "tagTYPEDESC"
void TExpandPointer(TYPEDESC typedesc, int indentation){
	for(int i = 0; i != indentation; ++i)std::wcout<<TEXT("    ");
	TCHAR szBuffer[20];
	VarTranslate(typedesc.vt, szBuffer, 20);
	std::wcout<<"  typedesc.vt = "<<szBuffer<<" typedesc.hreftype = "<<typedesc.hreftype<<std::endl;
	if(typedesc.lptdesc != 0 && (typedesc.lptdesc->lpadesc != 0 || typedesc.lptdesc->lptdesc != 0))
		TExpandPointer(*(typedesc.lptdesc), indentation + 1);
	if(typedesc.lpadesc != 0 && ((*(typedesc.lpadesc)).tdescElem.lptdesc != 0 || (*(typedesc.lpadesc)).tdescElem.lpadesc != 0)){
		int cDims = (*(typedesc.lpadesc)).cDims;
		for(int i = 0; i != cDims; ++i)
			TExpandPointer((*(typedesc.lpadesc)).tdescElem, indentation);
	}
}

bool WriteBuffer(TCHAR* szBuffer, TCHAR* insz, const int Blength){
	TCHAR *p = szBuffer, *q = insz; int b = Blength - 1; if(b<=0)return 1;
	while(*p++ = *q++)
		if(--b == 0 && *q != 0) return 1;
	return 0;
}

void wParamFlagsTranslate(const USHORT wParamFlags, TCHAR* szBuffer, const int Blength){
	ZeroMemory(szBuffer, Blength); std::string str = "[";
	if(wParamFlags & PARAMFLAG_NONE) str += ", ";
	if(wParamFlags & PARAMFLAG_FIN) str += "in, ";
	if(wParamFlags & PARAMFLAG_FOUT) str += "out, ";
	if(wParamFlags & PARAMFLAG_FLCID) str += "lcid, ";
	if(wParamFlags & PARAMFLAG_FRETVAL) str += "retval, ";
	if(wParamFlags & PARAMFLAG_FOPT) str += "optional, ";
	if(wParamFlags & PARAMFLAG_FHASDEFAULT) str += "defaultvalue(x), ";
	if(wParamFlags & PARAMFLAG_FHASCUSTDATA) str += "custom, ";
	if(str.size() > 2){str.pop_back(); str.pop_back();} str += "]";
    MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, str.c_str(), -1, szBuffer, Blength);
}

int VarTranslate(const int VT_, TCHAR* szBuffer, const int Blength){
	ZeroMemory(szBuffer, Blength);
	switch(VT_){
	case VT_I2:
		return WriteBuffer(szBuffer, TEXT("VT_I2"), Blength);
	case VT_I4:
		return WriteBuffer(szBuffer, TEXT("VT_I4"), Blength);
	case VT_R4:
		return WriteBuffer(szBuffer, TEXT("VT_R4"), Blength); 
	case VT_R8:
		return WriteBuffer(szBuffer, TEXT("VT_R8"), Blength); 
	case VT_CY:
		return WriteBuffer(szBuffer, TEXT("VT_CY"), Blength); 
	case VT_DATE:
		return WriteBuffer(szBuffer, TEXT("VT_DATE"), Blength); 
	case VT_BSTR:
		return WriteBuffer(szBuffer, TEXT("VT_BSTR"), Blength); 
	case VT_DISPATCH:
		return WriteBuffer(szBuffer, TEXT("VT_DISPATCH"), Blength); 
	case VT_ERROR:
		return WriteBuffer(szBuffer, TEXT("VT_ERROR"), Blength); 
	case VT_BOOL:
		return WriteBuffer(szBuffer, TEXT("VT_BOOL"), Blength); 
	case VT_VARIANT:
		return WriteBuffer(szBuffer, TEXT("VT_VARIANT"), Blength); 
	case VT_UNKNOWN:
		return WriteBuffer(szBuffer, TEXT("VT_UNKNOWN"), Blength); 
	case VT_DECIMAL:
		return WriteBuffer(szBuffer, TEXT("VT_DECIMAL"), Blength); 
	case VT_I1:
		return WriteBuffer(szBuffer, TEXT("VT_I1"), Blength); 
	case VT_UI1:
		return WriteBuffer(szBuffer, TEXT("VT_UI1"), Blength); 
	case VT_UI2:
		return WriteBuffer(szBuffer, TEXT("VT_UI2"), Blength); 
	case VT_UI4:
		return WriteBuffer(szBuffer, TEXT("VT_UI4"), Blength); 
	case VT_I8:
		return WriteBuffer(szBuffer, TEXT("VT_I8"), Blength); 
	case VT_UI8:
		return WriteBuffer(szBuffer, TEXT("VT_UI8"), Blength); 
	case VT_INT:
		return WriteBuffer(szBuffer, TEXT("VT_INT"), Blength);
	case VT_UINT:
		return WriteBuffer(szBuffer, TEXT("VT_UINT"), Blength); 
	case VT_INT_PTR:
		return WriteBuffer(szBuffer, TEXT("VT_INT_PTR"), Blength); 
	case VT_UINT_PTR:
		return WriteBuffer(szBuffer, TEXT("VT_UINT_PTR"), Blength); 
	case VT_VOID:
		return WriteBuffer(szBuffer, TEXT("VT_VOID"), Blength); 
	case VT_HRESULT:
		return WriteBuffer(szBuffer, TEXT("VT_HRESULT"), Blength); 
	case VT_PTR:
		return WriteBuffer(szBuffer, TEXT("VT_PTR"), Blength); 
	case VT_SAFEARRAY:
		return WriteBuffer(szBuffer, TEXT("VT_SAFEARRAY"), Blength); 
	case VT_CARRAY:
		return WriteBuffer(szBuffer, TEXT("VT_CARRAY"), Blength); 
	case VT_USERDEFINED:
		return WriteBuffer(szBuffer, TEXT("VT_USERDEFINED"), Blength); 
	case VT_LPSTR:
		return WriteBuffer(szBuffer, TEXT("VT_LPSTR"), Blength); 
	case VT_LPWSTR:
		return WriteBuffer(szBuffer, TEXT("VT_LPWSTR"), Blength); 
	default:
		return -1;
	}
}



