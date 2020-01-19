// stdafx.cpp : source file that includes just the standard includes
//  stdafx.pch will be the pre-compiled header
//  stdafx.obj will contain the pre-compiled type information

#include "stdafx.h"

#ifdef _ATL_STATIC_REGISTRY
#include <statreg.h>
#include <statreg.cpp>
#endif

#if _MSC_VER <= 1200
#include <atlimpl.cpp>
#endif

#if defined(_DEBUG) && _MSC_VER > 1200
#pragma comment(lib, "iconv32D.lib")
#else
#pragma comment(lib, "iconv32.lib")
#endif
