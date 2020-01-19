// ObjectProxy.h: Definition of the CObjectProxy class
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_OBJECTPROXY_H__2F8E55FA_710F_45FA_BD9C_2109A82DB836__INCLUDED_)
#define AFX_OBJECTPROXY_H__2F8E55FA_710F_45FA_BD9C_2109A82DB836__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CObjectProxy

namespace PaintsNow {
	class CScriptEngine;
}

class CObjectProxy :
	public IDispatchImpl<IObjectProxy, &IID_IObjectProxy, &LIBID_WINSCRIPTHOSTLib>,
	public ISupportErrorInfo,
	public CComObjectRoot,
	public CComCoClass<CObjectProxy, &CLSID_ObjectProxy>,
	public ITypeInfo
{
public:
	CObjectProxy();
	virtual ~CObjectProxy();

BEGIN_COM_MAP(CObjectProxy)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IObjectProxy)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()
//DECLARE_NOT_AGGREGATABLE(CObjectProxy) 
// Remove the comment from the line above if you don't want your object to 
// support aggregation. 

DECLARE_REGISTRY_RESOURCEID(IDR_ObjectProxy)
// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);


public:
	// IDispatch interfaces
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo** pptinfo);
	virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgdispid);
	virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr);

public:
	// ITypeInfo interfaces
	virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetTypeAttr(
		/* [out] */ TYPEATTR **ppTypeAttr);

	virtual HRESULT STDMETHODCALLTYPE GetTypeComp(
		/* [out] */ ITypeComp **ppTComp);

	virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetFuncDesc(
		/* [in] */ UINT index,
		/* [out] */ FUNCDESC **ppFuncDesc);

	virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetVarDesc(
		/* [in] */ UINT index,
		/* [out] */ VARDESC **ppVarDesc);

	virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetNames(
		/* [in] */ MEMBERID memid,
		/* [length_is][size_is][out] */ BSTR *rgBstrNames,
		/* [in] */ UINT cMaxNames,
		/* [out] */ UINT *pcNames);

	virtual HRESULT STDMETHODCALLTYPE GetRefTypeOfImplType(
		/* [in] */ UINT index,
		/* [out] */ HREFTYPE *pRefType);

	virtual HRESULT STDMETHODCALLTYPE GetImplTypeFlags(
		/* [in] */ UINT index,
		/* [out] */ INT *pImplTypeFlags);

	virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetIDsOfNames(
		/* [size_is][in] */ LPOLESTR *rgszNames,
		/* [in] */ UINT cNames,
		/* [size_is][out] */ MEMBERID *pMemId);

	virtual /* [local] */ HRESULT STDMETHODCALLTYPE Invoke(
		/* [in] */ PVOID pvInstance,
		/* [in] */ MEMBERID memid,
		/* [in] */ WORD wFlags,
		/* [out][in] */ DISPPARAMS *pDispParams,
		/* [out] */ VARIANT *pVarResult,
		/* [out] */ EXCEPINFO *pExcepInfo,
		/* [out] */ UINT *puArgErr);

	virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetDocumentation(
		/* [in] */ MEMBERID memid,
		/* [out] */ BSTR *pBstrName,
		/* [out] */ BSTR *pBstrDocString,
		/* [out] */ DWORD *pdwHelpContext,
		/* [out] */ BSTR *pBstrHelpFile);

	virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetDllEntry(
		/* [in] */ MEMBERID memid,
		/* [in] */ INVOKEKIND invKind,
		/* [out] */ BSTR *pBstrDllName,
		/* [out] */ BSTR *pBstrName,
		/* [out] */ WORD *pwOrdinal);

	virtual HRESULT STDMETHODCALLTYPE GetRefTypeInfo(
		/* [in] */ HREFTYPE hRefType,
		/* [out] */ ITypeInfo **ppTInfo);

	virtual /* [local] */ HRESULT STDMETHODCALLTYPE AddressOfMember(
		/* [in] */ MEMBERID memid,
		/* [in] */ INVOKEKIND invKind,
		/* [out] */ PVOID *ppv);

	virtual /* [local] */ HRESULT STDMETHODCALLTYPE CreateInstance(
		/* [in] */ IUnknown *pUnkOuter,
		/* [in] */ REFIID riid,
		/* [iid_is][out] */ PVOID *ppvObj);

	virtual HRESULT STDMETHODCALLTYPE GetMops(
		/* [in] */ MEMBERID memid,
		/* [out] */ BSTR *pBstrMops);

	virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetContainingTypeLib(
		/* [out] */ ITypeLib **ppTLib,
		/* [out] */ UINT *pIndex);

	virtual /* [local] */ void STDMETHODCALLTYPE ReleaseTypeAttr(
		/* [in] */ TYPEATTR *pTypeAttr);

	virtual /* [local] */ void STDMETHODCALLTYPE ReleaseFuncDesc(
		/* [in] */ FUNCDESC *pFuncDesc);

	virtual /* [local] */ void STDMETHODCALLTYPE ReleaseVarDesc(
		/* [in] */ VARDESC *pVarDesc);

// IObjectProxy
public:
	STDMETHOD(SetHost)(/*[in]*/ LONGLONG base);
	STDMETHOD(GetAttrib)(/*[in]*/ BSTR key, /*[out, retval]*/ VARIANT* var);
	STDMETHOD(SetAttrib)(/*[in]*/ BSTR, /*[in]*/ VARIANT* var);
	STDMETHOD(GetType)(/*[out, retval]*/ BSTR* type);

	TYPEATTR typeAttr;
	struct Function {
		Function();
		~Function();
		Function(const Function& rhs);
		Function& operator = (const Function& rhs);
		CComBSTR name;
		FUNCDESC desc;
		const PaintsNow::IScript::Request::AutoWrapperBase* wrapper;
	};

	std::vector<Function> funcDescs;
	std::map<CComBSTR, MEMBERID> funcIds;

	struct Attribute {
		Attribute();
		~Attribute();
		CComBSTR name;
		VARDESC desc;
		CComVariant value;
	};

	std::vector<Attribute> attrDescs;
	std::map<CComBSTR, MEMBERID> attrIds;
	PaintsNow::CScriptEngine* engine;
};

#endif // !defined(AFX_OBJECTPROXY_H__2F8E55FA_710F_45FA_BD9C_2109A82DB836__INCLUDED_)
