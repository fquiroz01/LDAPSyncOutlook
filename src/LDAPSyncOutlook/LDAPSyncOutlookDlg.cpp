
// LDAPSyncOutlookDlg.cpp : implementation file
//

#include "stdafx.h"
#include "LDAPSyncOutlook.h"
#include "LDAPSyncOutlookDlg.h"
#include "afxdialogex.h"
#include "../comun/AboutDlg.h"
#include "../comun/AdvancedDlg.h"
# include "../comun/MyRegKey.h"

#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52" rename_namespace("Office") rename("RGB", "MsoRGB") , named_guids
// #import "C:\Archivos de programa\Microsoft Office\Office\mso9.dll" rename_namespace("Office"), named_guids
using namespace Office;

// #import "libid:00062FFF-0000-0000-C000-000000000046" rename_namespace("Outlook") rename("CopyFile", "MsoCopyFile") rename("GetOrganizer", "GetOrganizer1"), named_guids // , raw_interfaces_only
#import "progid:Outlook.Application" rename_namespace("Outlook") rename("CopyFile", "MsoCopyFile") rename ("GetOrganizer", "GetOrganizer1"), named_guids // , raw_interfaces_only
// #import "c:\\Program Files (x86)\\Microsoft Office\\Office15\\MSOUTL.olb" rename_namespace("Outlook"), named_guids // , raw_interfaces_only
using namespace Outlook;
// #import "C:\\Program Files (x86)\\Common Files\\Designer\\MSADDNDR.DLL" raw_interfaces_only, raw_native_types, no_namespace, named_guids

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CLDAPSyncOutlookDlg dialog
CLDAPSyncOutlookDlg::CLDAPSyncOutlookDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CLDAPSyncOutlookDlg::IDD, pParent)
{
	m_pActual = NULL;
	m_nActual = -1;
	m_nSegundos = 0;
	m_pThread = NULL;
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CLDAPSyncOutlookDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LSITIOS, m_LCSitios);
	DDX_Control(pDX, IDC_ENOMBRE, m_ENombre);
	DDX_Control(pDX, IDC_ESERVIDOR, m_EServidor);
	DDX_Control(pDX, IDC_EBASE, m_EBase);
	DDX_Control(pDX, IDC_EDIRECTORIO, m_EDirectorio);
	DDX_Control(pDX, IDC_BAVANZADO, m_BBase);
	DDX_Control(pDX, IDC_BACEPTAR, m_BAceptar);
	// DDX_Control(pDX, IDC_SMAS, m_SMas);
	DDX_Control(pDX, IDC_LRESULTADOS, m_LCResultados);
	DDX_Control(pDX, IDC_PPROGRESO, m_PProgreso);
	DDX_Control(pDX, IDC_BEXAMINAR, m_BExaminar);
}

BEGIN_MESSAGE_MAP(CLDAPSyncOutlookDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_NOTIFY(LVN_KEYDOWN, IDC_LSITIOS, &CLDAPSyncOutlookDlg::OnLvnKeydownLsitios)
	ON_NOTIFY(NM_DBLCLK, IDC_LSITIOS, &CLDAPSyncOutlookDlg::OnNMDblclkLsitios)
	ON_NOTIFY(NM_RCLICK, IDC_LSITIOS, &CLDAPSyncOutlookDlg::OnNMRclickLsitios)
	ON_BN_CLICKED(IDC_BAVANZADO, &CLDAPSyncOutlookDlg::OnBnClickedBavanzado)
	ON_BN_CLICKED(IDCANCEL, &CLDAPSyncOutlookDlg::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_BACEPTAR, &CLDAPSyncOutlookDlg::OnBnClickedBaceptar)
	ON_BN_CLICKED(IDC_BAYUDA, &CLDAPSyncOutlookDlg::OnBnClickedBayuda)
	ON_COMMAND(2000, OnEditar)
	ON_COMMAND(2001, OnEliminar)
	ON_COMMAND(2002, OnActualizarSitio)
	ON_WM_DESTROY()
	ON_BN_CLICKED(IDC_BCANCELAR, &CLDAPSyncOutlookDlg::OnBnClickedBcancelar)
	ON_WM_TIMER()
	ON_REGISTERED_MESSAGE(CakTrayIcon::WM_TRAYICONNOTIFY, OnTrayIconNotify)
	ON_COMMAND(3000, OnConfigurar)
	ON_COMMAND(3001, OnAcercaDe)
	ON_COMMAND(3002, OnBnClickedCancel)
	ON_COMMAND(3003, OnActualizar)
	ON_COMMAND(3006, OnBnClickedBayuda)
	ON_BN_CLICKED(IDC_BEXAMINAR, &CLDAPSyncOutlookDlg::OnBnClickedBexaminar)
	ON_BN_CLICKED(IDOK, &CLDAPSyncOutlookDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// CLDAPSyncOutlookDlg message handlers

BOOL CLDAPSyncOutlookDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here

	m_LCSitios.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
	m_LCSitios.InsertColumn(0, L"Nombre", LVCFMT_LEFT, 100);
	m_LCSitios.InsertColumn(1, L"Servidor", LVCFMT_LEFT, 100);
	m_LCSitios.InsertColumn(2, L"Estado", LVCFMT_LEFT, 200);

	CStringArray sitios;
	CSitio::GetSitesCount(&sitios);
	for (int i = 0; i<sitios.GetCount(); i++) {
		CSitio *sitio = new CSitio(sitios[i]);
		m_LCSitios.InsertItem(i, sitios[i]);
		m_LCSitios.SetItemText(i, 1, sitio->m_szServidor);
		m_LCSitios.SetItemData(i, (DWORD_PTR)sitio);
	}

	m_LCResultados.InsertColumn(0, L"Sitio", LVCFMT_LEFT, 50);
	m_LCResultados.InsertColumn(1, L"Mensaje", LVCFMT_LEFT, 200);

	m_LCResultados.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

	ReloadSites();
	SetTimer(2, 1000, NULL);

	m_TrayIcon.Create(this, 1001);
	m_TrayIcon.SetIconAndTip(LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_MAINFRAME)), _T("SDG - LDAPSyncOutlook"));
	m_TrayIcon.Hide();

	Update(0);
	Update(1);
	CString cmdline = AfxGetApp()->m_lpCmdLine;

	if (cmdline.Find(L"-start") != -1)
		PostMessage(WM_SYSCOMMAND, SC_MINIMIZE, 0);

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CLDAPSyncOutlookDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		if (nID == SC_MINIMIZE) {
			ShowWindow(SW_HIDE);
			m_TrayIcon.Show();
		} else CDialogEx::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CLDAPSyncOutlookDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	} else {
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CLDAPSyncOutlookDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CLDAPSyncOutlookDlg::ReloadSites() {

	// este es el cancel.
	CWaitCursor wait;
	KillTimer(1);
	if (m_pThread) {
		m_pThread->Kill(true);
	}

	m_PProgreso.SetRange(0, 100);
	m_PProgreso.SetPos(0);

	for (int i = 0; i<m_LCSitios.GetItemCount(); i++) {
		CSitio *sitio = (CSitio *)m_LCSitios.GetItemData(i);
		sitio->m_nRemaing = 1;
	}
	m_nSegundos = 30;
#ifdef _DEBUG
	m_nSegundos = 59;
#endif
	SetTimer(1, 1000, NULL);
}

void CLDAPSyncOutlookDlg::Step() {

	CString texto;
	texto.Format(L"LDAPSyncOutlook %ls", LDAPSyncOutlook_VERSION);
	SetWindowText(texto);
	m_nActual = -1;
	CSitio *sitio;
	for (int i = 0; i<m_LCSitios.GetItemCount(); i++) {

		sitio = (CSitio *)m_LCSitios.GetItemData(i);
		sitio->m_nRemaing--;
		if (sitio->m_nRemaing <= 0) {
			if (m_nActual == -1)
				m_nActual = i;
		}
	}
	if (m_nActual == -1) {
		SetTimer(1, 1000, NULL);
	}
	else {
		if (!Update())
			SetTimer(1, 1000, NULL);
	}
}

bool CLDAPSyncOutlookDlg::Update(int nSitio/* = -1*/) {

	if (nSitio != -1)
		m_nActual = nSitio;
	if (m_nActual == -1)
		return false;
	CSitio *sitio = (CSitio *)m_LCSitios.GetItemData(m_nActual);
	if (sitio == NULL)
		return false;
	SetWindowText(L"Actualizando a " + sitio->m_szNombre);
	m_pThread = new CThreadOutlook(m_pThread);
	m_pThread->Run(this);
	return true;
}

void CLDAPSyncOutlookDlg::OnAplicar() {

	CMyRegKey key(DEFAULT_REG_KEY);
	key.DeleteRegKey(L"", L"sitios");

	CSitio *info;
	for (int i = 0; i<m_LCSitios.GetItemCount(); i++) {

		info = (CSitio *)m_LCSitios.GetItemData(i);
		info->GuardarEnRegistro();
	}
	ReloadSites();
}

void CLDAPSyncOutlookDlg::OnLvnKeydownLsitios(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if (pLVKeyDow->wVKey == VK_DELETE) {
		OnEliminar();
	}
	*pResult = 0;
}

void CLDAPSyncOutlookDlg::OnNMDblclkLsitios(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;

	if (pNMListView->iItem == -1 || pNMListView->iSubItem == -1)
		return;
	OnEditar();
	*pResult = 0;
}

void CLDAPSyncOutlookDlg::OnNMRclickLsitios(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;

	if (pNMListView->iItem == -1 || pNMListView->iSubItem == -1)
		return;

	CMenu menu;

	menu.CreatePopupMenu();

	menu.AppendMenu(MF_BYPOSITION | MF_STRING | MF_DEFAULT, 2000, L"&Editar");
	menu.AppendMenu(MF_BYPOSITION | MF_STRING, 2002, L"Actualizar");
	menu.AppendMenu(MF_BYPOSITION | MF_SEPARATOR);
	menu.AppendMenu(MF_BYPOSITION | MF_STRING, 2001, L"Eliminar");

	POINT point;
	GetCursorPos(&point);
	menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON, point.x, point.y, this);
	*pResult = 0;
}

void CLDAPSyncOutlookDlg::OnBnClickedBavanzado()
{
	CAdvancedDlg dlg;
	dlg.m_pActual = &m_pAvanzado;
	if (dlg.DoModal() == IDOK) {
	}
}

void CLDAPSyncOutlookDlg::OnEditar() {

	int index = m_LCSitios.GetNextItem(-1, LVIS_SELECTED);
	if (index == -1)
		return;

	CSitio *sitio = (CSitio *)m_LCSitios.GetItemData(index);
	m_pActual = sitio;
	m_pAvanzado.m_nMinutos = m_pActual->m_nMinutos;
	m_pAvanzado.m_nPuerto = m_pActual->m_nPuerto;
	m_pAvanzado.m_bAuth = m_pActual->m_bAuth;
	m_pAvanzado.m_szUsuario = m_pActual->m_szUsuario;
	m_pAvanzado.m_szContrasena = m_pActual->m_szContrasena;

	m_ENombre.SetWindowText(sitio->m_szNombre);
	m_EServidor.SetWindowText(sitio->m_szServidor);
	m_EBase.SetWindowText(sitio->m_szBase);
	m_EDirectorio.SetWindowText(sitio->m_szDestino);
	m_szUID = m_pActual->m_szUID;

	m_BAceptar.SetWindowText(L"Editar");
}

void CLDAPSyncOutlookDlg::OnEliminar() {

	int index = m_LCSitios.GetNextItem(-1, LVIS_SELECTED);
	if (index == -1)
		return;
	if (MessageBox(L"Esta seguro de eliminar a " + m_LCSitios.GetItemText(index, 0), L"LDAPSyncOutlook", MB_YESNO | MB_ICONQUESTION) == IDYES) {

		CSitio *sitio = (CSitio *)m_LCSitios.GetItemData(index);
		m_LCSitios.DeleteItem(index);
		delete sitio;
		OnAplicar();
		if (m_pActual) {
			OnBnClickedBcancelar();
		}
	}
}

void CLDAPSyncOutlookDlg::OnBnClickedCancel()
{
#ifndef _DEBUG
	if (MessageBox(L"¿Esta seguro que desea salir de LDAPSyncOutlook?", L"LDAPSyncOutlook", MB_ICONQUESTION | MB_YESNOCANCEL) == IDYES)
#endif
		OnCancel();
}

void CLDAPSyncOutlookDlg::OnBnClickedBaceptar()
{
	CString nombre, servidor, base, directorio;

	if (m_ENombre.GetWindowTextLength() == 0) {
		MessageBox(L"Por favor ingrese un nombre");
		m_ENombre.SetFocus();
		return;
	}
	m_ENombre.GetWindowText(nombre);

	if (m_EServidor.GetWindowTextLength() == 0) {
		MessageBox(L"Por favor ingrese un servidor");
		m_EServidor.SetFocus();
		return;
	}
	m_EServidor.GetWindowText(servidor);

	if (m_EBase.GetWindowTextLength() == 0) {
		MessageBox(L"Por favor ingrese un base de busqueda");
		m_EBase.SetFocus();
		return;
	}
	m_EBase.GetWindowText(base);
	m_EDirectorio.GetWindowText(directorio);

	CSitio *sitio;
	if (m_pActual)
		sitio = m_pActual;
	else
		sitio = new CSitio(nombre);

	sitio->m_szServidor = servidor;
	sitio->m_szNombre = nombre;
	sitio->m_szBase = base;
	sitio->m_nMinutos = m_pAvanzado.m_nMinutos;
	sitio->m_nPuerto = m_pAvanzado.m_nPuerto;
	sitio->m_bAuth = m_pAvanzado.m_bAuth;
	sitio->m_szUsuario = m_pAvanzado.m_szUsuario;
	sitio->m_szContrasena = m_pAvanzado.m_szContrasena;
	sitio->m_szUID = m_szUID;
	sitio->m_szDestino = directorio;

	int pos = m_LCSitios.GetItemCount();
	if (m_pActual) {
		for (int i = 0; i<m_LCSitios.GetItemCount(); i++) {
			if (m_LCSitios.GetItemData(i) == (DWORD_PTR)m_pActual) {
				pos = i;
				break;
			}
		}
	}
	else {
		for (int i = 0; i<m_LCSitios.GetItemCount(); i++) {
			if (m_LCSitios.GetItemText(i, 0) == sitio->m_szNombre) {
				MessageBox(L"Ya existe un sitio llamado '" + sitio->m_szNombre + "'");
				delete sitio;
				return;
			}
		}
		m_LCSitios.InsertItem(pos, L"");
	}

	m_LCSitios.SetItemText(pos, 0, sitio->m_szNombre);
	m_LCSitios.SetItemText(pos, 1, sitio->m_szServidor);
	m_LCSitios.SetItemData(pos, (DWORD_PTR)sitio);

	OnAplicar();
	OnBnClickedBcancelar();
}

void CLDAPSyncOutlookDlg::OnBnClickedBayuda()
{
	ShellExecute(NULL, L"open", LDAPHelpLink, NULL, NULL, SW_SHOWNORMAL);
}

void CLDAPSyncOutlookDlg::OnDestroy()
{
	CWaitCursor wait;

	KillTimer(1);
	KillTimer(2);
	m_TrayIcon.Hide();

	m_pThread->Kill(false);
	CSitio *info;
	for (int i = 0; i<m_LCSitios.GetItemCount(); i++) {

		info = (CSitio *)m_LCSitios.GetItemData(i);
		delete info;
	}
	m_LCSitios.DeleteAllItems();
	CDialog::OnDestroy();
}

void CLDAPSyncOutlookDlg::OnBnClickedBcancelar()
{
	m_ENombre.SetWindowText(L"");
	m_EServidor.SetWindowText(L"");
	m_EBase.SetWindowText(L"");
	m_EDirectorio.SetWindowText(L"");

	m_BAceptar.SetWindowText(L"Agregar");

	m_szUID.Empty();
	m_pActual = NULL;
	m_pAvanzado.Clean();
}

void CLDAPSyncOutlookDlg::OnTimer(UINT_PTR nIDEvent)
{
	if (nIDEvent == 1)	{
		m_nSegundos++;
		if (m_nSegundos == 60) {
			KillTimer(1);
			m_nSegundos = 0;
			Step();
		}
	} else if (nIDEvent == 2) {
		CSitio *sitio;
		for (int i = 0; i<m_LCSitios.GetItemCount(); i++) {

			if (i != m_nActual) {
				sitio = (CSitio *)m_LCSitios.GetItemData(i);
				CString texto;
				texto.Format(L"Actualización en %d minuto%ls", sitio->m_nRemaing, sitio->m_nRemaing>1 ? L"s" : L"");
				m_LCSitios.SetItemText(i, 2, texto);
			}
		}
	}
	CDialog::OnTimer(nIDEvent);
}

LRESULT CLDAPSyncOutlookDlg::OnTrayIconNotify(WPARAM wParam, LPARAM lParam)
{
	if (1001 == wParam)
	{
		if (WM_LBUTTONDBLCLK == lParam)
		{
			OnConfigurar();
		}
		else if (WM_RBUTTONUP == lParam)
		{
			CPoint pt;
			GetCursorPos(&pt);
			CMenu menu;
			menu.CreatePopupMenu();
			menu.AppendMenu(MF_STRING, 3000, L"&Configurar");
			menu.AppendMenu(MF_SEPARATOR);
			menu.AppendMenu(MF_STRING, 3003, L"Ac&tualizar");
			menu.AppendMenu(MF_SEPARATOR);
			menu.AppendMenu(MF_STRING, 3006, L"A&yuda");
			menu.AppendMenu(MF_STRING, 3001, L"&Acerca de LDAPSyncOutlook");
			menu.AppendMenu(MF_SEPARATOR);
			menu.AppendMenu(MF_STRING, 3002, L"&Salir");
			menu.SetDefaultItem(3000);
			menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_LEFTBUTTON, pt.x, pt.y, this);
			PostMessage(WM_NULL);
		}
	}
	return 0;
}

void CLDAPSyncOutlookDlg::OnConfigurar() {

	m_TrayIcon.Hide();
	ShowWindow(SW_SHOW);
	SetForegroundWindow();
}

void CLDAPSyncOutlookDlg::OnAcercaDe() {

	CAboutDlg dlgAbout;
	dlgAbout.DoModal();
}


void CLDAPSyncOutlookDlg::OnBnClickedBexaminar()
{
	MAPIFolderPtr pFolder;

	_ApplicationPtr m_pApp;
	CString directorio;
	bool directorioerrado = false;
	// es necesario abrir el outlook porque de lo contrario deshabilita los complementos y es necesario reiniciar el computador.
	::ShellExecute(NULL, L"open", L"Outlook", NULL, NULL, SW_NORMAL);
	VERIFY(SUCCEEDED(CoInitialize(NULL)));

	HRESULT hr = m_pApp.CreateInstance(__uuidof(Application));
	if (FAILED(hr)) {
		AfxMessageBox(L"Es necesario tener Microsoft Outlook instalado.");
		return;
	}
	try
	{
		pFolder = m_pApp->GetNamespace(_bstr_t(L"MAPI"))->PickFolder();
		if (!pFolder)
			return;
		if (pFolder->GetDefaultItemType() != olContactItem) {
			directorioerrado = true;
		} else {
			m_szUID = (LPCTSTR)pFolder->GetEntryID();
			directorio = (LPCTSTR)pFolder->Name;
			m_EDirectorio.SetWindowText(directorio);
		}
		m_pApp->Quit();
		m_pApp.Release();
	} catch (_com_error &e) {
		AfxMessageBox(e.ErrorMessage(), MB_OK);
	}
	CoUninitialize();
	if (directorioerrado)
		AfxMessageBox(L"El Directorio seleccionado no es de contactos", MB_OK);
}

void CLDAPSyncOutlookDlg::OnActualizarSitio() {

	POSITION pos = m_LCSitios.GetFirstSelectedItemPosition();
	if (pos == NULL)
		return;
	int index = m_LCSitios.GetNextSelectedItem(pos);
	if (index == -1)
		return;
	Update(index);
}

void CLDAPSyncOutlookDlg::OnActualizar() {

	Update(0);
	OnConfigurar();
}

void CLDAPSyncOutlookDlg::OnBnClickedOk()
{
	PostMessage(WM_SYSCOMMAND, SC_MINIMIZE, 0);
	// CDialogEx::OnOK();
}
