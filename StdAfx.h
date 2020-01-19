// stdafx.h : include file for standard system include files,
//      or project specific include files that are used frequently,
//      but are changed infrequently


#if !defined(AFX_STDAFX_H__A9DA4493_B92D_4BD0_AB7F_EE2E908E6F0A__INCLUDED_)
#define AFX_STDAFX_H__A9DA4493_B92D_4BD0_AB7F_EE2E908E6F0A__INCLUDED_
#define WINVER 0x501
#define _WIN32_WINNT 0x501

#if _MSC_VER > 1200
#pragma warning(disable:4996)
#endif 

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define STRICT
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0400
#endif
#define _ATL_APARTMENT_THREADED

#include <afxwin.h>
#include <afxdisp.h>

#include <atlbase.h>
extern CComModule _Module;
//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module
#include <atlcom.h>
#include <vector>

#define PAINTSNOW_NO_PREIMPORT_LIBS
#include "Core/Interface/IScript.h"
#include "Core/Driver/Thread/Pthread/ZThreadPthread.h"

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__A9DA4493_B92D_4BD0_AB7F_EE2E908E6F0A__INCLUDED)
