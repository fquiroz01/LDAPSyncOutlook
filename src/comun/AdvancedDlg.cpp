// AdvancedDlg.cpp : implementation file
//

#include "stdafx.h"
// #include "LDAPSyncOutlook.h"
#include "AdvancedDlg.h"
#include "afxdialogex.h"


// CAdvancedDlg dialog

IMPLEMENT_DYNAMIC(CAdvancedDlg, CDialogEx)

CAdvancedDlg::CAdvancedDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CAdvancedDlg::IDD, pParent)
{
	m_pActual = NULL;
}

CAdvancedDlg::~CAdvancedDlg()
{
}

void CAdvancedDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EPUERTO, m_EPuerto);
	DDX_Control(pDX, IDC_EUSUARIO, m_EUsuario);
	DDX_Control(pDX, IDC_ECONTRASENA, m_EContrasena);
	DDX_Control(pDX, IDC_EACTUALIZAR, m_EActualizar);
	DDX_Control(pDX, IDC_CAUTH, m_CAuth);
}


BEGIN_MESSAGE_MAP(CAdvancedDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &CAdvancedDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &CAdvancedDlg::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_CAUTH, &CAdvancedDlg::OnBnClickedCauth)
END_MESSAGE_MAP()


// CAdvancedDlg message handlers
BOOL CAdvancedDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	CString temp;
	temp.Format(L"%d", m_pActual ? m_pActual->m_nPuerto : 389);
	m_EPuerto.SetWindowText(temp);
	m_CAuth.SetCheck(m_pActual ? m_pActual->m_bAuth : 0);
	m_EUsuario.SetWindowText(m_pActual ? m_pActual->m_szUsuario : L"");
	m_EContrasena.SetWindowText(m_pActual ? m_pActual->m_szContrasena : L"");
	temp.Format(L"%d", m_pActual ? m_pActual->m_nMinutos : 60);
	m_EActualizar.SetWindowText(temp);

	OnBnClickedCauth();

	return TRUE;
}

void CAdvancedDlg::OnBnClickedOk()
{
	if (m_pActual == NULL)
		m_pActual = &m_pSitio;
	CString temp;
	m_EPuerto.GetWindowText(temp);
	m_pActual->m_nPuerto = _wtoi(temp);
	m_pActual->m_bAuth = m_CAuth.GetCheck()==1;
	m_EUsuario.GetWindowText(m_pActual->m_szUsuario);
	m_EContrasena.GetWindowText(m_pActual->m_szContrasena);
	m_EActualizar.GetWindowText(temp);
	m_pActual->m_nMinutos = _wtoi(temp);
	OnOK();
}

void CAdvancedDlg::OnBnClickedCancel()
{
	OnCancel();
}

void CAdvancedDlg::OnBnClickedCauth()
{
	m_EUsuario.EnableWindow(m_CAuth.GetCheck());
	m_EContrasena.EnableWindow(m_CAuth.GetCheck());
}
