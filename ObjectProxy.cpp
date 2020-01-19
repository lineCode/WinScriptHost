// ObjectProxy.cpp : Implementation of CWinScriptHostApp and DLL registration.

#include "stdafx.h"
#include "WinScriptHost.h"
#include "ObjectProxy.h"
#include "ScriptEngine.h"

/////////////////////////////////////////////////////////////////////////////
//

using namespace PaintsNow;

// static int init = 0;
CObjectProxy::CObjectProxy() : engine(NULL) {
	// init++;
	memset(&typeAttr, 0, sizeof(typeAttr));
}

CObjectProxy::~CObjectProxy() {
	// init--;
	// printf("INIT %d", init);
}

STDMETHODIMP CObjectProxy::InterfaceSupportsErrorInfo(REFIID riid) {
	static const IID* arr[] =
	{
		&IID_IObjectProxy,
	};

	for (int i = 0; i < sizeof(arr) / sizeof(arr[0]); i++) {
		if (IsEqualGUID(*arr[i], riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CObjectProxy::GetType(BSTR *type) {
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	*type = ::SysAllocString(L"Object");
	return S_OK;
}

STDMETHODIMP CObjectProxy::SetAttrib(BSTR key, VARIANT *var) {
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	std::map<CComBSTR, MEMBERID>::iterator p = funcIds.find(key);
	std::map<CComBSTR, MEMBERID>::iterator q = attrIds.find(key);

	const int WRAPPER_MASK = (VT_BYREF | VT_UI1);
	if (p != funcIds.end()) {
		MEMBERID id = p->second;
		if ((var->vt & WRAPPER_MASK) == WRAPPER_MASK) {
			funcDescs[id].wrapper = reinterpret_cast<const PaintsNow::IScript::Request::AutoWrapperBase*>(var->punkVal);
			return S_OK;
		} else {
			assert(false);
			funcDescs[id].wrapper = NULL;
			funcIds.erase(p);
			typeAttr.cFuncs--;
		}
	} else if ((var->vt & WRAPPER_MASK) == WRAPPER_MASK) {
		if (q != attrIds.end()) {
			assert(false);
			attrIds.erase(q);
			typeAttr.cVars--;
		}

		Function func;
		func.name = key;
		func.wrapper = reinterpret_cast<const PaintsNow::IScript::Request::AutoWrapperBase*>(var->punkVal)->Clone();
		funcIds[key] = (MEMBERID)funcDescs.size();
		funcDescs.push_back(func);
		typeAttr.cFuncs++;
		return S_OK;
	}

	if (q != attrIds.end()) {
		MEMBERID ida = q->second;
		attrDescs[ida].value = *var;
	} else {
		Attribute attr;
		attr.name = key;
		attr.value = *var;
		attrIds[key] = (MEMBERID)attrDescs.size();
		attrDescs.push_back(attr);
		typeAttr.cVars++;
	}

	return S_OK;
}

STDMETHODIMP CObjectProxy::GetAttrib(BSTR key, VARIANT *var) {
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	std::map<CComBSTR, MEMBERID>::iterator p = funcIds.find(key);
	std::map<CComBSTR, MEMBERID>::iterator q = attrIds.find(key);
	if (p != funcIds.end()) {
		CComVariant v;
		v.vt = VT_UI1 | VT_BYREF;
		v.pbVal = ((BYTE*)funcDescs[p->second].wrapper);
		v.Detach(var);
	} else if (q != attrIds.end()) {
		CComVariant v(attrDescs[q->second].value);
		v.Detach(var);
	} else {
		var->vt = VT_EMPTY;
	}

	return S_OK;
}


// Override IDispatch interfaces

STDMETHODIMP CObjectProxy::GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo** pptinfo) {
	*pptinfo = this;
	return S_OK;
}

STDMETHODIMP CObjectProxy::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgdispid) {
	return GetIDsOfNames(rgszNames, cNames, rgdispid);
	// return IDispatchImpl<IObjectProxy, &IID_IObjectProxy, &LIBID_WINSCRIPTHOSTLib>::GetIDsOfNames(riid, rgszNames, cNames, lcid, rgdispid);
}

STDMETHODIMP CObjectProxy::Invoke(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr) {
	return Invoke(this, dispidMember, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr);
}

//

CObjectProxy::Function::Function() : wrapper(NULL) {
	memset(&desc, 0, sizeof(desc));
}

CObjectProxy::Function::~Function() {
	if (wrapper != NULL) {
		delete wrapper;
	}
}


CObjectProxy::Function::Function(const Function& rhs) {
	*this = rhs;
}

CObjectProxy::Function& CObjectProxy::Function::operator = (const Function& rhs) {
	desc = rhs.desc;
	wrapper = rhs.wrapper->Clone();

	return *this;
}

CObjectProxy::Attribute::Attribute() {
	memset(&desc, 0, sizeof(desc));
}

CObjectProxy::Attribute::~Attribute() {}


// Override ITypeInfo interfaces

HRESULT STDMETHODCALLTYPE CObjectProxy::GetTypeAttr(
	/* [out] */ TYPEATTR **ppTypeAttr) {
	*ppTypeAttr = &typeAttr;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CObjectProxy::GetTypeComp(
	/* [out] */ ITypeComp **ppTComp) {
	*ppTComp = NULL;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CObjectProxy::GetFuncDesc(
	/* [in] */ UINT index,
	/* [out] */ FUNCDESC **ppFuncDesc) {
	if (index >= 0 && index < funcDescs.size()) {
		*ppFuncDesc = &funcDescs[index].desc;
		return S_OK;
	} else {
		return DISP_E_BADINDEX;
	}
}

HRESULT STDMETHODCALLTYPE CObjectProxy::GetVarDesc(
	/* [in] */ UINT index,
	/* [out] */ VARDESC **ppVarDesc) {
	if (index >= 0 && index < funcDescs.size()) {
		*ppVarDesc = &attrDescs[index].desc;
		return S_OK;
	} else {
		return DISP_E_BADINDEX;
	}

	return S_OK;
}

const MEMBERID MAX_MEMBER_ID = 65534;

HRESULT STDMETHODCALLTYPE CObjectProxy::GetNames(
	/* [in] */ MEMBERID memid,
	/* [length_is][size_is][out] */ BSTR *rgBstrNames,
	/* [in] */ UINT cMaxNames,
	/* [out] */ UINT *pcNames) {
	if (memid < (MEMBERID)attrDescs.size()) {
		*rgBstrNames = attrDescs[memid].name.Copy();
		*pcNames = 1;
	} else if (MAX_MEMBER_ID - memid < (MEMBERID)funcDescs.size()) {
		*rgBstrNames = funcDescs[MAX_MEMBER_ID - memid].name.Copy();
		*pcNames = 1;
	} else {
		*pcNames = 0;
	}

	/*
	for (std::map<CComBSTR, MEMBERID>::const_iterator p = funcIds.begin(); p != funcIds.end() && cMaxNames != 0; ++p) {
		*rgBstrNames++ = p->first.Copy();
		cMaxNames--;
	}

	for (std::map<CComBSTR, MEMBERID>::const_iterator q = attrIds.begin(); q != attrIds.end() && cMaxNames != 0; ++q) {
		*rgBstrNames++ = q->first.Copy();
		cMaxNames--;
	}*/

	return S_OK;
}

HRESULT STDMETHODCALLTYPE CObjectProxy::GetRefTypeOfImplType(
	/* [in] */ UINT index,
	/* [out] */ HREFTYPE *pRefType) {

	assert(false);
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CObjectProxy::GetImplTypeFlags(
	/* [in] */ UINT index,
	/* [out] */ INT *pImplTypeFlags) {

	assert(false);
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CObjectProxy::GetIDsOfNames(
	/* [size_is][in] */ LPOLESTR *rgszNames,
	/* [in] */ UINT cNames,
	/* [size_is][out] */ MEMBERID *pMemId) {
	while (cNames-- > 0) {
		if (wcscmp(*rgszNames, L"length") == 0) {
			*pMemId = -1;
		} else {
			std::map<CComBSTR, MEMBERID>::const_iterator p = funcIds.find(*rgszNames);
			if (p != funcIds.end()) {
				*pMemId = MAX_MEMBER_ID - p->second;
			} else {
				std::map<CComBSTR, MEMBERID>::const_iterator p = attrIds.find(*rgszNames);
				if (p != attrIds.end()) {
					*pMemId = p->second;
				} else {
					*pMemId = MAX_MEMBER_ID + 1;
				}
			}
		}

		rgszNames++;
		pMemId++;
	}

	return S_OK;
}

HRESULT STDMETHODCALLTYPE CObjectProxy::Invoke(
	/* [in] */ PVOID pvInstance,
	/* [in] */ MEMBERID memid,
	/* [in] */ WORD wFlags,
	/* [out][in] */ DISPPARAMS *pDispParams,
	/* [out] */ VARIANT *pVarResult,
	/* [out] */ EXCEPINFO *pExcepInfo,
	/* [out] */ UINT *puArgErr) {
	if (wFlags & DISPATCH_METHOD) {
		if (memid == -1) { // get length
			CComVariant v((int)attrDescs.size());
			v.Detach(pVarResult);

			return S_OK;
		} else if (memid != MAX_MEMBER_ID + 1 && MAX_MEMBER_ID - memid < (MEMBERID)funcDescs.size()) {
			CScriptEngine::Request& request = static_cast<CScriptEngine::Request&>(engine->GetDefaultRequest());
			request.Push();
			for (size_t i = 0; i < pDispParams->cArgs; i++) {
				request << pDispParams->rgvarg[pDispParams->cArgs - i - 1];
			}

			(funcDescs[MAX_MEMBER_ID - memid].wrapper->Execute)(request);

			// collect result
			// request >> skip(pDispParams->cArgs);
			if (pVarResult != NULL) {
				CComVariant res;
				request >> res;
				res.Detach(pVarResult);
			}

			request.Pop();

			return S_OK;
		}
	}
	
	if (wFlags & DISPATCH_PROPERTYPUT) {
		if (memid < (MEMBERID)attrDescs.size()) {
			attrDescs[memid].value = *pDispParams[0].rgvarg;
		}

		if (pVarResult != NULL) {
			pVarResult->vt = VT_EMPTY;
		}
	} else if (wFlags & DISPATCH_PROPERTYGET) {
		if (memid < (MEMBERID)attrDescs.size()) {
			CComVariant v;
			v.Copy(&attrDescs[memid].value);
			v.Detach(pVarResult);
		} else {
			pVarResult->vt = VT_EMPTY;
		}
	}

	return S_OK;
}

HRESULT STDMETHODCALLTYPE CObjectProxy::GetDocumentation(
	/* [in] */ MEMBERID memid,
	/* [out] */ BSTR *pBstrName,
	/* [out] */ BSTR *pBstrDocString,
	/* [out] */ DWORD *pdwHelpContext,
	/* [out] */ BSTR *pBstrHelpFile) {

	return S_OK;
}

HRESULT STDMETHODCALLTYPE CObjectProxy::GetDllEntry(
	/* [in] */ MEMBERID memid,
	/* [in] */ INVOKEKIND invKind,
	/* [out] */ BSTR *pBstrDllName,
	/* [out] */ BSTR *pBstrName,
	/* [out] */ WORD *pwOrdinal) {
	assert(false);
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CObjectProxy::GetRefTypeInfo(
	/* [in] */ HREFTYPE hRefType,
	/* [out] */ ITypeInfo **ppTInfo) {
	assert(false);
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CObjectProxy::AddressOfMember(
	/* [in] */ MEMBERID memid,
	/* [in] */ INVOKEKIND invKind,
	/* [out] */ PVOID *ppv) {
	assert(false);
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CObjectProxy::CreateInstance(
	/* [in] */ IUnknown *pUnkOuter,
	/* [in] */ REFIID riid,
	/* [iid_is][out] */ PVOID *ppvObj) {
	assert(false);
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CObjectProxy::GetMops(
	/* [in] */ MEMBERID memid,
	/* [out] */ BSTR *pBstrMops) {
	assert(false);
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CObjectProxy::GetContainingTypeLib(
	/* [out] */ ITypeLib **ppTLib,
	/* [out] */ UINT *pIndex) {
	assert(false);
	return S_OK;
}

void STDMETHODCALLTYPE CObjectProxy::ReleaseTypeAttr(
	/* [in] */ TYPEATTR *pTypeAttr) {}

void STDMETHODCALLTYPE CObjectProxy::ReleaseFuncDesc(
	/* [in] */ FUNCDESC *pFuncDesc) {}

void STDMETHODCALLTYPE CObjectProxy::ReleaseVarDesc(
	/* [in] */ VARDESC *pVarDesc) {}

STDMETHODIMP CObjectProxy::SetHost(LONGLONG base)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	engine = reinterpret_cast<CScriptEngine*>(base);
	return S_OK;
}
