
////////////////////////////////////////////////////////////////////////////
//	Copyright : Amit Dey 2002
//
//	Email :amitdey@joymail.com
//
//	This code may be used in compiled form in any way you desire. This
//	file may be redistributed unmodified by any means PROVIDING it is 
//	not sold for profit without the authors written consent, and 
//	providing that this notice and the authors name is included.
//
//	This file is provided 'as is' with no expressed or implied warranty.
//	The author accepts no liability if it causes any damage to your computer.
//
//	Do expect bugs.
//	Please let me know of any bugs/mods/improvements.
//	and I will try to fix/incorporate them into this file.
//	Enjoy!
//
//	Description : Outlook2K Addin 
////////////////////////////////////////////////////////////////////////////

// PropPage.cpp : Implementation of CPropPage

#include "stdafx.h"
#include "OutlookAddin.h"
# include "Addin.h"
#include "PropPage.h"
#include <afxcmn.h>        // MFC Automation classes
# include "../comun/AdvancedDlg.h"
# include "../comun/MyRegKey.h"


/////////////////////////////////////////////////////////////////////////////
// CPropPage

STDMETHODIMP  CPropPage::GetControlInfo(LPCONTROLINFO pCI ) {

	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);

	CListCtrl *lista = (CListCtrl *)CListCtrl::FromHandle(GetDlgItem(IDC_LSITIOS).m_hWnd);
	lista->SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
	lista->InsertColumn(0, L"Nombre", LVCFMT_LEFT, 100);
	lista->InsertColumn(1, L"Servidor", LVCFMT_LEFT, 100);
	lista->InsertColumn(2, L"Estado", LVCFMT_LEFT, 200);	

	CStringArray sitios;
	CSitio::GetSitesCount(&sitios);
	for(int i=0; i<sitios.GetCount(); i++) {
		CSitio *sitio=  new CSitio(sitios[i]);
		lista->InsertItem(i, sitios[i]);
		lista->SetItemText(i, 1, sitio->m_szServidor);
		lista->SetItemData(i, (DWORD_PTR)sitio);
	}
	SetTimer(1, 1000, NULL);
	return S_OK;
};

// IOleObject
STDMETHODIMP CPropPage::SetClientSite(IOleClientSite *pClientSite){

	HRESULT result;
	// call default ATL implementation
	result = IOleObjectImpl<CPropPage>::SetClientSite(pClientSite);
	if (result != S_OK) return result;
	// pClientSite may be NULL when container has being destructed
	if (pClientSite != NULL) {
		CComQIPtr<PropertyPageSite> pPropertyPageSite(pClientSite);
		result = pPropertyPageSite.CopyTo(&m_pPropPageSite);
	}
	m_fDirty = false;
	return result;
}

STDMETHODIMP CPropPage::raw_GetPageInfo(BSTR * HelpFile,long * HelpContext ) {
	
#ifdef _DEBUG
	AfxMessageBox(L"raw_getpageinfo");
#endif
	return S_OK;
}

STDMETHODIMP CPropPage::raw_Apply() {

#ifdef _DEBUG
	AfxMessageBox(L"aplicar");
#endif
	return S_OK;
}

STDMETHODIMP CPropPage::get_Dirty(VARIANT_BOOL *Dirty)
{
#ifdef _DEBUG
	AfxMessageBox(L"dirty");
#endif
	*Dirty = m_fDirty.boolVal;
	return S_OK;
}
STDMETHODIMP CPropPage::Apply()
{
	KillTimer(1);

	CMyRegKey key(DEFAULT_REG_KEY);
	key.DeleteRegKey(L"", L"sitios");

	CSitio *info;
	CListCtrl *lista = (CListCtrl *)CListCtrl::FromHandle(GetDlgItem(IDC_LSITIOS).m_hWnd);

	for(int i=0; i<lista->GetItemCount(); i++) {

		info= (CSitio *)lista->GetItemData(i);
		info->GuardarEnRegistro();
	}
	if(CAddin::m_pDlgProgreso)
		CAddin::m_pDlgProgreso->ReloadSites();
	SetTimer(1, 1000, NULL);
	return S_OK;
}

STDMETHODIMP CPropPage::GetPageInfo(BSTR *HelpFile,LONG *HelpContext)
{
	if (HelpFile == NULL)
		return E_POINTER;
	if (HelpContext == NULL)
		return E_POINTER;
	return E_NOTIMPL;
}

LRESULT CPropPage::OnBnClickedBaceptar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	CString temp;
	CString nombre,servidor,base;

	CEdit *edit= (CEdit *)CWnd::FromHandle(GetDlgItem(IDC_ENOMBRE).m_hWnd);
	if(edit->GetWindowTextLength()==0) {
		MessageBox(L"Por favor ingrese un nombre");
		edit->SetFocus();
		return 0;
	}
	edit->GetWindowText(nombre);

	edit= (CEdit *)CWnd::FromHandle(GetDlgItem(IDC_ESERVIDOR).m_hWnd);
	if(edit->GetWindowTextLength()==0) {
		MessageBox(L"Por favor ingrese un servidor");
		edit->SetFocus();
		return 0;
	}
	edit->GetWindowText(servidor);

	edit= (CEdit *)CWnd::FromHandle(GetDlgItem(IDC_EBASE).m_hWnd);
	if(edit->GetWindowTextLength()==0) {
		MessageBox(L"Por favor ingrese un base de busqueda");
		edit->SetFocus();
		return 0;
	}
	edit->GetWindowText(base);

	edit= (CEdit *)CWnd::FromHandle(GetDlgItem(IDC_EDIRECTORIO).m_hWnd);
	if(edit->GetWindowTextLength()==0) {
		BOOL bAlgo;
		OnBnClickedBexplorar(0,0,0,bAlgo);
		if(edit->GetWindowTextLength()==0) {
			MessageBox(L"Por favor ingrese un directorio.");
			edit->SetFocus();
			return 0;
		}
	}
	edit->GetWindowText(temp);

	CSitio *sitio;
	if(m_pActual)
		sitio= m_pActual;
	else 
		sitio= new CSitio(temp);

	sitio->m_szServidor= servidor;
	sitio->m_szNombre= nombre;
	sitio->m_szBase= base;
	sitio->m_szDestino=temp;
	sitio->m_nMinutos= m_pAvanzado.m_nMinutos;
	sitio->m_nPuerto= m_pAvanzado.m_nPuerto;
	sitio->m_bAuth= m_pAvanzado.m_bAuth;
	sitio->m_szUsuario= m_pAvanzado.m_szUsuario;
	sitio->m_szContrasena= m_pAvanzado.m_szContrasena;
	sitio->m_szUID=m_szUID;


	CListCtrl *lista = (CListCtrl *)CListCtrl::FromHandle(GetDlgItem(IDC_LSITIOS).m_hWnd);

	int pos= lista->GetItemCount();
	if(m_pActual) {
		for(int i=0; i<lista->GetItemCount(); i++) {
			if(lista->GetItemData(i)==(DWORD_PTR)m_pActual) {
				pos=i;
				break;
			}
		}
	} else {
		for(int i=0; i<lista->GetItemCount(); i++) {
			if(lista->GetItemText(i, 0)==sitio->m_szNombre) {
				MessageBox(L"Ya existe un sitio llamado '" + sitio->m_szNombre + L"'");
				delete sitio;
				return 0;
			}
		}
		lista->InsertItem(pos, L"");
	}

	lista->SetItemText(pos, 0, sitio->m_szNombre);
	lista->SetItemText(pos, 1, sitio->m_szServidor);
	lista->SetItemData(pos, (DWORD_PTR)sitio);

	Apply();
	BOOL btemp= false;
	OnBnClickedBcancelar(0,0,0,btemp);

	return 0;
}

LRESULT CPropPage::OnBnClickedBcancelar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	CEdit *edit= (CEdit *)CWnd::FromHandle(GetDlgItem(IDC_ENOMBRE).m_hWnd);
	edit->SetWindowText(L"");
	edit= (CEdit *)CWnd::FromHandle(GetDlgItem(IDC_ESERVIDOR).m_hWnd);
	edit->SetWindowText(L"");
	edit= (CEdit *)CWnd::FromHandle(GetDlgItem(IDC_EBASE).m_hWnd);
	edit->SetWindowText(L"");
	edit= (CEdit *)CWnd::FromHandle(GetDlgItem(IDC_EDIRECTORIO).m_hWnd);
	edit->SetWindowText(L"");
	m_szUID.Empty();

	CButton *aceptar=(CButton *)CWnd::FromHandle(GetDlgItem(IDC_BACEPTAR).m_hWnd);
	aceptar->SetWindowText(L"Agregar");

	m_pActual= NULL;
	m_pAvanzado.Clean();

	return 0;
}

LRESULT CPropPage::OnBnClickedBexplorar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	CEdit *edit= (CEdit *)CWnd::FromHandle(GetDlgItem(IDC_EDIRECTORIO).m_hWnd);

	MAPIFolderPtr pFolder;
	CString szAux;

	_ApplicationPtr m_pApp= CAddin::m_pParentApp;

	try
	{
		pFolder= m_pApp->GetNamespace(_bstr_t("MAPI"))->PickFolder();
		if(!pFolder)
			return 0;
		if (pFolder->GetDefaultItemType()!=olContactItem) {
			AfxMessageBox(L"El Folder seleccionado no es de contactos",MB_OK);
			return 0;
		}
		m_szUID= (LPCTSTR) pFolder->GetEntryID();
		szAux= (LPCTSTR) pFolder->Name;
		edit->SetWindowText(szAux);
	}
	catch(_com_error &e)
	{
		AfxMessageBox(e.ErrorMessage(), MB_OK);
		return NULL;
	}
	return 0;
}

LRESULT CPropPage::OnBnClickedBavanzado(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	CAdvancedDlg dlg;
	dlg.m_pActual= &m_pAvanzado;
	if(dlg.DoModal()==IDOK) {
	}
	return 0;
}

LRESULT CPropPage::OnEditar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/) {

	CListCtrl *lista = (CListCtrl *)CListCtrl::FromHandle(GetDlgItem(IDC_LSITIOS).m_hWnd);
	int index= lista->GetNextItem(-1, LVIS_SELECTED);
	if(index==-1)
		return 0;

	CSitio *sitio= (CSitio *)lista->GetItemData(index);
	m_pActual= sitio;
	m_pAvanzado.m_nMinutos= m_pActual->m_nMinutos;
	m_pAvanzado.m_nPuerto= m_pActual->m_nPuerto;
	m_pAvanzado.m_bAuth= m_pActual->m_bAuth;
	m_pAvanzado.m_szUsuario= m_pActual->m_szUsuario;
	m_pAvanzado.m_szContrasena= m_pActual->m_szContrasena;
	m_szUID= m_pActual->m_szUID;

	CEdit *edit= (CEdit *)CWnd::FromHandle(GetDlgItem(IDC_ENOMBRE).m_hWnd);
	edit->SetWindowText(sitio->m_szNombre);
	edit= (CEdit *)CWnd::FromHandle(GetDlgItem(IDC_ESERVIDOR).m_hWnd);
	edit->SetWindowText(sitio->m_szServidor);
	edit= (CEdit *)CWnd::FromHandle(GetDlgItem(IDC_EBASE).m_hWnd);
	edit->SetWindowText(sitio->m_szBase);
	edit= (CEdit *)CWnd::FromHandle(GetDlgItem(IDC_EDIRECTORIO).m_hWnd);
	edit->SetWindowText(sitio->m_szDestino);

	CButton *aceptar=(CButton *)CWnd::FromHandle(GetDlgItem(IDC_BACEPTAR).m_hWnd);
	aceptar->SetWindowText(L"Editar");

	return 0;
}

LRESULT CPropPage::OnEliminar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/) {

	CListCtrl *lista = (CListCtrl *)CListCtrl::FromHandle(GetDlgItem(IDC_LSITIOS).m_hWnd);
	int index= lista->GetNextItem(-1, LVIS_SELECTED);
	if(index==-1)
		return 0;
	if(MessageBox(L"Esta seguro de eliminar a " + lista->GetItemText(index, 0), L"LDAPSyncOutlook", MB_YESNO | MB_ICONQUESTION)==IDYES) {

		CSitio *sitio= (CSitio *)lista->GetItemData(index);
		lista->DeleteItem(index);
		delete sitio;
		Apply();
		if(m_pActual) {
			BOOL bFalse;
			OnBnClickedBcancelar(0,0,0,bFalse);
		}
	}
	return 0;
}

LRESULT CPropPage::OnNMDblclkLsitios(int /*idCtrl*/, LPNMHDR pNMHDR, BOOL& /*bHandled*/)
{
	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;

	if( pNMListView->iItem == -1 || pNMListView->iSubItem == -1) 
		return 0;

	BOOL bFalse=1;
	OnEditar(0,0,0,bFalse);
	return 0;
}

LRESULT CPropPage::OnNMRclickLsitios(int /*idCtrl*/, LPNMHDR pNMHDR, BOOL& /*bHandled*/)
{
	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;

	if( pNMListView->iItem == -1 || pNMListView->iSubItem == -1) 
		return 0;

	CMenu menu;

	menu.CreatePopupMenu();
	menu.InsertMenu(0, MF_BYPOSITION | MF_STRING, 2000, L"&Editar");
	menu.InsertMenu(0, MF_BYPOSITION | MF_STRING, 2001, L"Eliminar");

	POINT point;
	GetCursorPos( &point );
	menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON, point.x, point.y, CWnd::FromHandle(m_hWnd));

	return 0;
}

LRESULT CPropPage::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	KillTimer(1);
	CSitio *info;
	CListCtrl *lista = (CListCtrl *)CListCtrl::FromHandle(GetDlgItem(IDC_LSITIOS).m_hWnd);

	for(int i=0; i<lista->GetItemCount(); i++) {

		info= (CSitio *)lista->GetItemData(i);
		delete info;
	}
	lista->DeleteAllItems();
	return 0;
}

LRESULT CPropPage::OnLvnKeydownLsitios(int /*idCtrl*/, LPNMHDR pNMHDR, BOOL& /*bHandled*/)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey==VK_DELETE) {
		BOOL bFalse=1;
		OnEliminar(0,0,0,bFalse);		
	}
	return 0;
}

LRESULT CPropPage::OnBnClickedBayuda(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	return 0;
}


LRESULT CPropPage::OnTimer(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	if(CAddin::m_pDlgProgreso) {

		CListCtrl *lista = (CListCtrl *)CListCtrl::FromHandle(GetDlgItem(IDC_LSITIOS).m_hWnd);
		if(lista) {
		
			CSitio *sitio;
			for(int i=0; i<CAddin::m_pDlgProgreso->m_pSitios.GetCount(); i++) {

				if(i!=CAddin::m_pDlgProgreso->m_nActual) {
					sitio= (CSitio *)CAddin::m_pDlgProgreso->m_pSitios[i];
					CString texto= sitio->GetMensaje();
					if(texto.IsEmpty())
						texto.Format(L"Actualización en %d minuto%ls", sitio->m_nRemaing, sitio->m_nRemaing>1 ? L"s" : L"");
					for(int j=0; j<lista->GetItemCount(); j++) {
						if(lista->GetItemText(j, 0)==sitio->m_szNombre) {
							lista->SetItemText(j, 2, texto);
						}
					}
				}
			}
		}
	}
	return 0;
}
