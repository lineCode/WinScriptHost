// PluginWinScript.cpp : Implementation of CWinScriptHostApp and DLL registration.

#include "stdafx.h"
#include "WinScriptHost.h"
#include "PluginWinScript.h"

/////////////////////////////////////////////////////////////////////////////
//

STDMETHODIMP CPluginWinScript::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IPluginWinScript,
	};

	for (int i=0;i<sizeof(arr)/sizeof(arr[0]);i++)
	{
		if (IsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CPluginWinScript::PluginMain(LPDISPATCH host, LONGLONG vm, BSTR *pluginName)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here

	return S_OK;
}

STDMETHODIMP CPluginWinScript::PluginConfig(LONGLONG hWnd)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	// TODO: Add your implementation code here

	return S_OK;
}
