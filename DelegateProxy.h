// DelegateProxy.h: Definition of the DelegateProxy class
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DELEGATEPROXY_H__51A4497E_4945_4591_9F78_339E7994ACC8__INCLUDED_)
#define AFX_DELEGATEPROXY_H__51A4497E_4945_4591_9F78_339E7994ACC8__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// DelegateProxy

namespace PaintsNow {
	class CScriptEngine;
}

class DelegateProxy : 
	public IDispatchImpl<IDelegateProxy, &IID_IDelegateProxy, &LIBID_WINSCRIPTHOSTLib>, 
	public ISupportErrorInfo,
	public CComObjectRoot,
	public CComCoClass<DelegateProxy,&CLSID_DelegateProxy>
{
public:
	DelegateProxy();
	virtual ~DelegateProxy();

BEGIN_COM_MAP(DelegateProxy)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IDelegateProxy)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()
//DECLARE_NOT_AGGREGATABLE(DelegateProxy) 
// Remove the comment from the line above if you don't want your object to 
// support aggregation. 

DECLARE_REGISTRY_RESOURCEID(IDR_DelegateProxy)
// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IDelegateProxy
public:
	STDMETHOD(SetHost)(/*[in]*/ LONGLONG host);
	STDMETHOD(GetProxy)(/*[out, retval]*/ LONGLONG* base);
	STDMETHOD(SetProxy)(/*[in]*/ LONGLONG base);
	STDMETHOD(GetType)(/*[out, retval]*/ BSTR* type);

protected:
	PaintsNow::IScript::Object* proxy;
	PaintsNow::CScriptEngine* engine;
};

#endif // !defined(AFX_DELEGATEPROXY_H__51A4497E_4945_4591_9F78_339E7994ACC8__INCLUDED_)
