#pragma once
#include "afxwin.h"
# include "resource.h"

# define LDAPSyncOutlook_VERSION	L"3.2"
# define LDAPHelpLink	L"https://github.com/fquiroz01/LDAPSyncOutlook"
// CAboutDlg dialog

class CAboutDlg : public CDialog
{
	DECLARE_DYNAMIC(CAboutDlg)

public:
	CAboutDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedBAyuda();
	CStatic m_SVersion;
};
