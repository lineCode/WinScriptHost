// PluginWinScript.h: Definition of the CPluginWinScript class
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_PLUGINWINSCRIPT_H__95D2BE6E_CBF3_467A_AF49_616AA068FC96__INCLUDED_)
#define AFX_PLUGINWINSCRIPT_H__95D2BE6E_CBF3_467A_AF49_616AA068FC96__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CPluginWinScript

class CPluginWinScript : 
	public IDispatchImpl<IPluginWinScript, &IID_IPluginWinScript, &LIBID_WINSCRIPTHOSTLib>, 
	public ISupportErrorInfo,
	public CComObjectRoot,
	public CComCoClass<CPluginWinScript,&CLSID_PluginWinScript>
{
public:
	CPluginWinScript() {}
BEGIN_COM_MAP(CPluginWinScript)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IPluginWinScript)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()
//DECLARE_NOT_AGGREGATABLE(CPluginWinScript) 
// Remove the comment from the line above if you don't want your object to 
// support aggregation. 

DECLARE_REGISTRY_RESOURCEID(IDR_PluginWinScript)
// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IPluginWinScript
public:
	STDMETHOD(PluginConfig)(/*[in]*/ LONGLONG hWnd);
	STDMETHOD(PluginMain)(/*[in]*/ LPDISPATCH host, /*[in]*/ LONGLONG vm, /*[out, retval]*/ BSTR* pluginName);
};

#endif // !defined(AFX_PLUGINWINSCRIPT_H__95D2BE6E_CBF3_467A_AF49_616AA068FC96__INCLUDED_)
