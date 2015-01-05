// AboutDlg.cpp : implementation file
//

#include "stdafx.h"
#include "AboutDlg.h"


// CAboutDlg dialog

IMPLEMENT_DYNAMIC(CAboutDlg, CDialog)

CAboutDlg::CAboutDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CAboutDlg::IDD, pParent)
{

}

CAboutDlg::~CAboutDlg()
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_SVERSION, m_SVersion);
}


BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	ON_BN_CLICKED(IDC_BAYUDA, &CAboutDlg::OnBnClickedBAyuda)
END_MESSAGE_MAP()


// CAboutDlg message handlers

BOOL CAboutDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	CString texto;
	texto.Format(L"Versión %ls",LDAPSyncOutlook_VERSION);
	m_SVersion.SetWindowText(texto);
	
	return TRUE;
}

void CAboutDlg::OnBnClickedBAyuda() {

	ShellExecute(NULL, L"open", LDAPHelpLink, NULL, NULL, SW_SHOWNORMAL);
}
