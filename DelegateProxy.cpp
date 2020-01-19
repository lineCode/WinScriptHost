// DelegateProxy.cpp : Implementation of CWinScriptHostApp and DLL registration.

#include "stdafx.h"
#include "WinScriptHost.h"
#include "DelegateProxy.h"
#include "ScriptEngine.h"

/////////////////////////////////////////////////////////////////////////////
//

using namespace PaintsNow;

DelegateProxy::DelegateProxy() : proxy(NULL), engine(NULL) {}
DelegateProxy::~DelegateProxy() {
	if (proxy != NULL) {
		engine->RemoveDelegate(proxy);
		IScript::Request& request = engine->GetDefaultRequest();
		proxy->ScriptUninitialize(request);
	}
}

STDMETHODIMP DelegateProxy::InterfaceSupportsErrorInfo(REFIID riid) {
	static const IID* arr[] =
	{
		&IID_IDelegateProxy,
	};

	for (int i = 0; i < sizeof(arr) / sizeof(arr[0]); i++) {
		if (IsEqualGUID(*arr[i], riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP DelegateProxy::GetType(BSTR *type) {
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	*type = ::SysAllocString(L"Delegate");
	return S_OK;
}

STDMETHODIMP DelegateProxy::SetProxy(LONGLONG base) {
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	assert(proxy == NULL);
	proxy = reinterpret_cast<IScript::Object*>((void*)base);
	return S_OK;
}

STDMETHODIMP DelegateProxy::GetProxy(LONGLONG *base) {
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	*base = (LONGLONG)proxy;
	return S_OK;
}

STDMETHODIMP DelegateProxy::SetHost(LONGLONG host)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	engine = reinterpret_cast<CScriptEngine*>((void*)host);

	return S_OK;
}
