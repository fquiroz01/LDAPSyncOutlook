
// LDAPSyncOutlookDlg.h : header file
//

#pragma once
# include "ThreadOutlook.h"
# include "resource.h"
# include "akTrayIcon.h"
// CLDAPSyncOutlookDlg dialog
class CLDAPSyncOutlookDlg : public CDialogEx
{
public:
	int					m_nActual;
	int					m_nSegundos;
	CThreadOutlook		*m_pThread;
// Construction
public:
	CLDAPSyncOutlookDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_LDAPSYNCOUTLOOK_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:
	CListCtrl m_LCSitios;
public:
	CEdit m_ENombre;
public:
	CEdit m_EServidor;
public:
	CEdit m_EBase;
public:
	CEdit m_EDirectorio;
public:
	CButton m_BBase;
public:
	afx_msg void OnLvnKeydownLsitios(NMHDR *pNMHDR, LRESULT *pResult);
public:
	afx_msg void OnNMDblclkLsitios(NMHDR *pNMHDR, LRESULT *pResult);
public:
	afx_msg void OnNMRclickLsitios(NMHDR *pNMHDR, LRESULT *pResult);
public:
	afx_msg void OnBnClickedBavanzado();
public:
	afx_msg void OnBnClickedCancel();
public:
	afx_msg void OnBnClickedBaceptar();
public:
	afx_msg void OnBnClickedBayuda();
	void OnAplicar();
	CSitio *m_pActual;
	CSitio m_pAvanzado;
public:
	CButton m_BAceptar;
	afx_msg void OnEditar();
	afx_msg void OnEliminar();
	afx_msg void OnActualizarSitio();
public:
	afx_msg void OnDestroy();
public:
	afx_msg void OnBnClickedBcancelar();
	afx_msg void OnActualizar();
public:
	// CSectionSeparator	m_SMas;
public:
	CListCtrl m_LCResultados;
public:
	CProgressCtrl m_PProgreso;
	void ReloadSites();
	void Step();
	bool Update(int nSitio=-1);
public:
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	CakTrayIcon m_TrayIcon;
	LRESULT OnTrayIconNotify(WPARAM wParam, LPARAM lParam);
	void OnConfigurar();
	void OnAcercaDe();
public:
	CButton m_BExaminar;
public:
	afx_msg void OnBnClickedBexaminar();
	CString m_szUID;
	afx_msg void OnBnClickedOk();
};
