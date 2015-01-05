
// LDAPSyncOutlook.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "LDAPSyncOutlook.h"
#include "LDAPSyncOutlookDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CLDAPSyncOutlookApp

BEGIN_MESSAGE_MAP(CLDAPSyncOutlookApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// CLDAPSyncOutlookApp construction

CLDAPSyncOutlookApp::CLDAPSyncOutlookApp()
{
	// support Restart Manager
	m_dwRestartManagerSupportFlags = AFX_RESTART_MANAGER_SUPPORT_RESTART;

	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}


// The one and only CLDAPSyncOutlookApp object

CLDAPSyncOutlookApp theApp;


BOOL CLDAPSyncOutlookApp::InitInstance(LPCTSTR lpszUID /*NULL*/)
{
	CString commandline = m_lpCmdLine;
	if (lpszUID != NULL)
	{
		// If two different applications register the same message string, 
		// the applications return the same message value. The message 
		// remains registered until the session ends. If the message is 
		// successfully registered, the return value is a message identifier 
		// in the range 0xC000 through 0xFFFF.

		ASSERT(m_uMsgCheckInst == NULL); // Only once
		m_uMsgCheckInst = ::RegisterWindowMessage(lpszUID);
		ASSERT(m_uMsgCheckInst >= 0xC000 && m_uMsgCheckInst <= 0xFFFF);

		// If another instance is found, pass document file name to it
		// only if command line contains FileNew or FileOpen parameters
		// and exit instance. In other cases such as FilePrint, etc., 
		// do not exit instance because framework will process the commands 
		// itself in invisible mode and will exit.

		if (this->FindAnotherInstance(lpszUID))
		{
			WPARAM wpCmdLine = NULL;
			wpCmdLine = (ATOM)::GlobalAddAtom(commandline);
			this->PostInstanceMessage(wpCmdLine, NULL);
			return FALSE;
		}
	}
	if (commandline.Find(L"-close") != -1)
		return false;
	return true;
}

// CLDAPSyncOutlookApp initialization

BOOL CLDAPSyncOutlookApp::InitInstance()
{
	// InitCommonControlsEx() is required on Windows XP if an application
	// manifest specifies use of ComCtl32.dll version 6 or later to enable
	// visual styles.  Otherwise, any window creation will fail.
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	// Set this to include all the common control classes you want to use
	// in your application.
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);

	// Search the running app
	if (!InitInstance(L"{6A9FA33B-7458-4115-B96A-B9D28D4B4DF9}"))
		return false;
	CWinApp::InitInstance();

	AfxEnableControlContainer();

	// Create the shell manager, in case the dialog contains
	// any shell tree view or shell list view controls.
	CShellManager *pShellManager = new CShellManager;

	// Activate "Windows Native" visual manager for enabling themes in MFC controls
	CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows));

	// Standard initialization
	// If you are not using these features and wish to reduce the size
	// of your final executable, you should remove from the following
	// the specific initialization routines you do not need
	// Change the registry key under which our settings are stored
	// TODO: You should modify this string to be something appropriate
	// such as the name of your company or organization
	SetRegistryKey(_T("SDG"));

	CLDAPSyncOutlookDlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with OK
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with Cancel
	}
	else if (nResponse == -1)
	{
		TRACE(traceAppMsg, 0, "Warning: dialog creation failed, so application is terminating unexpectedly.\n");
		TRACE(traceAppMsg, 0, "Warning: if you are using MFC controls on the dialog, you cannot #define _AFX_NO_MFC_CONTROLS_IN_DIALOGS.\n");
	}

	// Delete the shell manager created above.
	if (pShellManager != NULL)
	{
		delete pShellManager;
	}

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}

BOOL CLDAPSyncOutlookApp::PreTranslateMessage(LPMSG pMsg)
{
	if (pMsg->message == m_uMsgCheckInst) {
		// Get command line arguments (if any) from new instance.

		BOOL bShellOpen = FALSE;

		if (pMsg->wParam != NULL)
		{
			if (::GlobalGetAtomName((ATOM)pMsg->wParam, m_lpCmdLine, _MAX_FNAME)) {
				::GlobalDeleteAtom((ATOM)pMsg->wParam);
				bShellOpen = TRUE;
			}
		}

		// If got valid command line then try to open the document -
		// CDocManager will popup main window itself. Otherwise, we
		// have to popup the window 'manually' :
		CString commandline = m_lpCmdLine;
		if (commandline.Find(L"-close") != -1) {
			if (::IsWindow(m_pMainWnd->GetSafeHwnd()))
				m_pMainWnd->SendMessage(WM_CLOSE);
		} else if (::IsWindow(m_pMainWnd->GetSafeHwnd())) {
			CWnd* pPopupWnd = m_pMainWnd->GetLastActivePopup();
			pPopupWnd->BringWindowToTop();

			// If window is not visible then show it, else if
			// it is iconic, restore it:

			if (!m_pMainWnd->IsWindowVisible())
				m_pMainWnd->ShowWindow(SW_SHOWNORMAL);
			else if (m_pMainWnd->IsIconic())
				m_pMainWnd->ShowWindow(SW_RESTORE);

			// And finally, bring to top after showing again:

			pPopupWnd->BringWindowToTop();
			pPopupWnd->SetForegroundWindow();
		}
		return TRUE;
	}
	return CWinApp::PreTranslateMessage(pMsg);
}

BOOL CLDAPSyncOutlookApp::FindAnotherInstance(LPCTSTR lpszUID)
{
	ASSERT(lpszUID != NULL);

	// Create a new mutex. If fails means that process already exists:

	HANDLE hMutex = ::CreateMutex(NULL, FALSE, lpszUID);
	DWORD  dwError = ::GetLastError();

	if (hMutex != NULL)
	{
		// Close mutex handle
		::ReleaseMutex(hMutex);
		hMutex = NULL;

		// Another instance of application is running:

		if (dwError == ERROR_ALREADY_EXISTS || dwError == ERROR_ACCESS_DENIED)
			return TRUE;
	}

	return FALSE;
}

BOOL CLDAPSyncOutlookApp::PostInstanceMessage(WPARAM wParam, LPARAM lParam)
{
	ASSERT(m_uMsgCheckInst != NULL);

	// One process can terminate the other process by broadcasting the private message 
	// using the BroadcastSystemMessage function as follows:

	DWORD dwReceipents = BSM_APPLICATIONS;

	if (this->EnableTokenPrivilege(SE_TCB_NAME))
		dwReceipents |= BSM_ALLDESKTOPS;

	// Send the message to all other instances.
	// If the function succeeds, the return value is a positive value.
	// If the function is unable to broadcast the message, the return value is –1.

	LONG lRet = ::BroadcastSystemMessage(BSF_IGNORECURRENTTASK | BSF_FORCEIFHUNG |
		BSF_POSTMESSAGE, &dwReceipents, m_uMsgCheckInst, wParam, lParam);

	return (BOOL)(lRet != -1);
}

BOOL CLDAPSyncOutlookApp::EnableTokenPrivilege(LPCTSTR lpszSystemName,
	BOOL    bEnable /*TRUE*/)
{
	ASSERT(lpszSystemName != NULL);
	BOOL bRetVal = FALSE;

	HANDLE hToken = NULL;
	if (::OpenProcessToken(::GetCurrentProcess(),
		TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken))
	{
		TOKEN_PRIVILEGES tp = { 0 };

		if (::LookupPrivilegeValue(NULL, lpszSystemName, &tp.Privileges[0].Luid))
		{
			tp.PrivilegeCount = 1;
			tp.Privileges[0].Attributes = (bEnable ? SE_PRIVILEGE_ENABLED : 0);

			// To determine whether the function adjusted all of the 
			// specified privileges, call GetLastError:

			if (::AdjustTokenPrivileges(hToken, FALSE, &tp,
				sizeof(TOKEN_PRIVILEGES), (PTOKEN_PRIVILEGES)NULL, NULL))
			{
				bRetVal = (::GetLastError() == ERROR_SUCCESS);
			}
		}
		::CloseHandle(hToken);
	}

	return bRetVal;
}
