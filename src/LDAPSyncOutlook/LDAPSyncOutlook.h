
// LDAPSyncOutlook.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


// CLDAPSyncOutlookApp:
// See LDAPSyncOutlook.cpp for the implementation of this class
//

class CLDAPSyncOutlookApp : public CWinApp
{
public:
	CLDAPSyncOutlookApp();

// Overrides
public:
	virtual BOOL InitInstance();
	virtual BOOL InitInstance(LPCTSTR lpszUID = NULL);

// Implementation

	DECLARE_MESSAGE_MAP()
protected:
	UINT	m_uMsgCheckInst;
protected:
	BOOL FindAnotherInstance(LPCTSTR lpszUID);
	BOOL PostInstanceMessage(WPARAM wParam, LPARAM lParam);
	BOOL EnableTokenPrivilege(LPCTSTR lpszSystemName, BOOL bEnable = TRUE);
	virtual BOOL PreTranslateMessage(LPMSG pMsg);
};

extern CLDAPSyncOutlookApp theApp;