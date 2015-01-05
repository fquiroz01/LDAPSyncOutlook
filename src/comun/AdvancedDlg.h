#pragma once
#include "../comun/Sitio.h"
#include "resource.h"
// CAdvancedDlg dialog

class CAdvancedDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CAdvancedDlg)

public:
	CAdvancedDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CAdvancedDlg();

// Dialog Data
	enum { IDD = IDD_DADVANCED };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	CSitio m_pSitio;
	CSitio *m_pActual;
	CEdit m_EPuerto;
	CEdit m_EUsuario;
	CEdit m_EContrasena;
	CEdit m_EActualizar;
	CButton m_CAuth;

	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnBnClickedCauth();
};
