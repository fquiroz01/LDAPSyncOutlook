#pragma once
#include "afxwin.h"
#include "afxcmn.h"
# include "resource.h"
# include "../comun/Sitio.h"
# include "ContactosThread.h"

// CProgresoDlg dialog

class CProgresoDlg : public CDialog
{
	DECLARE_DYNAMIC(CProgresoDlg)
	CPtrArray			m_pSitios;
	int					m_nActual;
	int					m_nSegundos;
	CContactosThread	*m_pThread;
	bool				m_bVisible;

public:
	CProgresoDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CProgresoDlg();

// Dialog Data
	enum { IDD = IDD_PROGRESODLG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	bool Update();
	DECLARE_MESSAGE_MAP()
public:
	CStatic m_SEstado;
public:
	CProgressCtrl m_PEstado;
public:
	virtual BOOL OnInitDialog();
public:
	afx_msg void OnBnClickedOk();
public:
	afx_msg void OnBnClickedCancel();
public:
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	void Cancel();
	void Step();
	void UpdateSites();
	void ReloadSites();
public:
	afx_msg void OnDestroy();
/*public:
	CSectionSeparator m_SMas;*/
public:
	CListCtrl m_LCResultados;
};
