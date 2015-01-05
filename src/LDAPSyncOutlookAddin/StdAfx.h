// stdafx.h : include file for standard system include files,
//      or project specific include files that are used frequently,
//      but are changed infrequently

#if !defined(AFX_STDAFX_H__6A4E951E_A471_4D04_95D2_73B98E912FF3__INCLUDED_)
#define AFX_STDAFX_H__6A4E951E_A471_4D04_95D2_73B98E912FF3__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define _ATL_APARTMENT_THREADED

#include <afxwin.h>
#include <afxdisp.h>
# include <afxcmn.h>
#include <afxdialogex.h>

#include <atlbase.h>
//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module
extern CComModule _Module;
#include <atlcom.h>
#include <atlhost.h>
#include <atlctl.h>

#include <afxdtctl.h>

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52" rename_namespace("Office") rename("RGB", "MsoRGB"), named_guids
// #import "C:\Archivos de programa\Microsoft Office\Office\mso9.dll" rename_namespace("Office"), named_guids
using namespace Office;

#import "libid:00062FFF-0000-0000-C000-000000000046" rename_namespace("Outlook") rename("CopyFile", "MsoCopyFile") rename("GetOrganizer", "GetOrganizer1"), named_guids // , raw_interfaces_only
// #import "c:\Archivos de programa\Microsoft Office\Office\MSOUTL9.olb" rename_namespace("Outlook"), named_guids // , raw_interfaces_only
using namespace Outlook;

# import "libid:AC0714F2-3D04-11D1-AE7D-00A0C90F26F4"  raw_interfaces_only, raw_native_types, no_namespace, named_guids 
// #import "C:\Archivos de programa\Archivos comunes\designer\MSADDNDR.dll" raw_interfaces_only, raw_native_types, no_namespace, named_guids 

# ifdef REDEMPTION
	#import "libid:2D5E2D34-BED5-4B9F-9793-A31E26E6806E" redemptionNamespace
# endif

// #ifdef _UNICODE
#if defined _M_IX86
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_IA64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='ia64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif
// #endif

#endif // !defined(AFX_STDAFX_H__6A4E951E_A471_4D04_95D2_73B98E912FF3__INCLUDED)
