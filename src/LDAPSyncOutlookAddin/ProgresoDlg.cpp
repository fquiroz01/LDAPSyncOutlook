// ProgresoDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ProgresoDlg.h"
# include "../comun/AboutDlg.h"


// CProgresoDlg dialog

IMPLEMENT_DYNAMIC(CProgresoDlg, CDialog)

CProgresoDlg::CProgresoDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CProgresoDlg::IDD, pParent)
{
	m_nActual=-1;
	m_nSegundos=0;
	m_pThread= NULL;

	if(Create(CProgresoDlg::IDD, pParent)) {
#ifdef _DEBUG
		m_bVisible= true;
		ShowWindow(SW_NORMAL);
#else 
		m_bVisible= false;
		ShowWindow(SW_HIDE);
#endif
		CRect rect, rectp;
		GetWindowRect(rect);
		CWnd *parent= GetParent();
		parent->GetWindowRect(rectp);
		rectp.right-= 30;
		rectp.bottom-=130;
		rectp.left= rectp.right-rect.Width();
		rectp.top= rectp.bottom-rect.Height();
		MoveWindow(rectp);
	}
}

CProgresoDlg::~CProgresoDlg()
{
	// el hilo se libera el mismo
	if(m_pThread) 
		m_pThread->Kill(true);
	CSitio *sitio;
	for(int i=0; i<m_pSitios.GetCount(); i++) {

		sitio= (CSitio *)m_pSitios[i];
		delete sitio;
	}
	m_pSitios.RemoveAll();
}

void CProgresoDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_SPROGRESO, m_SEstado);
	DDX_Control(pDX, IDC_PESTADO, m_PEstado);
	// DDX_Control(pDX, IDC_SMAS, m_SMas);
	DDX_Control(pDX, IDC_LERRORES, m_LCResultados);
}


BEGIN_MESSAGE_MAP(CProgresoDlg, CDialog)
	ON_BN_CLICKED(IDOK, &CProgresoDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &CProgresoDlg::OnBnClickedCancel)
	ON_WM_TIMER()
	ON_WM_DESTROY()
END_MESSAGE_MAP()


// CProgresoDlg message handlers

BOOL CProgresoDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	/*m_SMas.SetNextSectionCtrl(m_PEstado);
	m_SMas.Collapse();*/

	m_LCResultados.InsertColumn(0, L"Sitio", LVCFMT_LEFT, 50);
	m_LCResultados.InsertColumn(1, L"Mensaje", LVCFMT_LEFT, 200);

	ReloadSites();
	return TRUE;
}

void CProgresoDlg::ReloadSites() {

	Cancel();
	CSitio *sitio;
	for(int i=0; i<m_pSitios.GetCount(); i++) {

		sitio= (CSitio *)m_pSitios[i];
		delete sitio;
	}
	m_pSitios.RemoveAll();

	m_PEstado.SetRange(0,100);
	m_PEstado.SetPos(0);
	CStringArray sitios;
	CSitio::GetSitesCount(&sitios);
	for(int i=0; i<sitios.GetCount(); i++) {
		CSitio *sitio= new CSitio(sitios[i]);
		sitio->m_nRemaing=1;
		m_pSitios.Add(sitio);
	}
	m_nSegundos=30;
#ifdef _DEBUG
	m_nSegundos=59;
#endif
	SetTimer(1, 1000, NULL);
}

void CProgresoDlg::UpdateSites() {
	
	if(m_nActual==-1) {
		KillTimer(1);
		m_nSegundos=59;
		for(int i=0; i<m_pSitios.GetCount(); i++) {
			CSitio *sitio= (CSitio *)m_pSitios[i];
			sitio->m_nRemaing=1;
		}
		SetTimer(1, 1000, NULL);
	}
	m_bVisible= true;
	ShowWindow(SW_NORMAL);
}

void CProgresoDlg::OnBnClickedOk()
{
	// OnOK();
}

void CProgresoDlg::OnBnClickedCancel()
{
	m_bVisible= false;
#if _DEBUG
#else
	ShowWindow(SW_HIDE);
#endif
	// OnCancel();
}

void CProgresoDlg::Cancel() {

	if(::IsWindow(*this))
		KillTimer(1);
	if(m_pThread) {
		m_pThread->Kill(true);
	}
}

void CProgresoDlg::Step() {

	CString texto;
	texto.Format(L"LDAPSyncOutlook %ls", LDAPSyncOutlook_VERSION);
	SetWindowText(texto);
	m_nActual=-1;
	CSitio *sitio;
	for(int i=0; i<m_pSitios.GetCount(); i++) {

		sitio= (CSitio *)m_pSitios[i];
		sitio->m_nRemaing--;
		if(sitio->m_nRemaing<=0) {
			if(m_nActual==-1)
				m_nActual= i;
		}
	}
	if(m_nActual==-1) {
		SetTimer(1, 1000, NULL);
		if(m_bVisible)
			PostMessage(WM_CLOSE);
	} else {
		if(!Update())
			SetTimer(1, 1000, NULL);
	}
}

bool CProgresoDlg::Update() {

	if(m_nActual==-1)
		return false;
	CSitio *sitio= (CSitio *)m_pSitios[m_nActual];
	if(sitio==NULL)
		return false;
	SetWindowText(L"Actualizando a " + sitio->m_szNombre);
	m_pThread= new CContactosThread(m_pThread);
	m_pThread->Run(this);
	return true;
}

void CProgresoDlg::OnTimer(UINT_PTR nIDEvent)
{
	if(nIDEvent==1)	{
		m_nSegundos++;
		if(m_nSegundos==60) {
			KillTimer(1);
			m_nSegundos=0;
			Step();
		}
	}
	CDialog::OnTimer(nIDEvent);
}

void CProgresoDlg::OnDestroy()
{
	KillTimer(1);
	Cancel();
	if(m_pThread) 
		m_pThread->Kill(true);
	CDialog::OnDestroy();
}
