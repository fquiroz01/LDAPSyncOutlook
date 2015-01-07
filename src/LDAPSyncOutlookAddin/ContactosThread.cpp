#include "stdafx.h"
#include "ContactosThread.h"
#include "ProgresoDlg.h"
# include "contacto.h"

#include "OutlookAddin.h"
# include "Addin.h"

CContactosThread::CContactosThread(CContactosThread *old/*=NULL*/) : CThreadLDAP(old)
{
}

CContactosThread::~CContactosThread(void)
{
}

void CContactosThread::SetMensaje(const CString mensaje,CSitio *sitio/*=NULL*/) {

	if(sitio) {
		sitio->SetMensaje(mensaje);
		SetError(sitio, mensaje);
	} else {
		CProgresoDlg *dlg= (CProgresoDlg * )m_pParametro;
		dlg->m_SEstado.SetWindowText(mensaje);
	}
}

void CContactosThread::SetError(const CString &sitio, const CString &mensaje) {

	CProgresoDlg *dlg= CAddin::m_pDlgProgreso;
	if(dlg && ::IsWindow(*dlg)) {
		dlg->m_SEstado.SetWindowText(mensaje);
		int pos= dlg->m_LCResultados.InsertItem(0, sitio);
		dlg->m_LCResultados.SetItemText(pos, 1, mensaje);
		dlg->m_LCResultados.EnsureVisible(pos, false);
	}
}

void CContactosThread::SetError(CSitio *sitio, const CString &mensaje) {

	SetError(sitio ? sitio->m_szNombre : "", mensaje);
}

bool CContactosThread::Init() {

	m_pWrapper.Initialize();
	::CoInitialize(NULL);

	m_pMutex.Lock();
	CProgresoDlg *dlg= (CProgresoDlg *)m_pParametro;
	if(dlg) {
		CSitio *sitio= (CSitio *)dlg->m_pSitios[dlg->m_nActual];
		SetError(sitio, "Conectando con servidor LDAP");
		int actualizados= 0;
		int eliminados=0;
		int nuevos=0;
		if(!sitio)
			goto fin;
		if (sitio->m_szUID.IsEmpty()) {
			SetError(sitio, "No hay directorio definido");
			goto fin;
		}
		dlg->m_PEstado.SetPos(0);
		if(sitio && GetLDAP(sitio)) {

			_ItemsPtr pItems;
			_ContactItemPtr pContact;
			MAPIFolderPtr pFolder;

			CString szAux, value, szParametro;
			CString temp;

			if(!IsRunning())
				goto fin;

			dlg->m_PEstado.SetRange(0, m_pContactos.GetCount());
			if(!IsRunning())
				goto fin;

			SetMensaje(L"Cargando libreta de direcciones");
			try
			{
				if (CAddin::m_fVersion >= 15) { // outlook 2013 no admite llamadas desde hilos
					HRESULT hr = m_pApp.CreateInstance(__uuidof(Application));
					if (FAILED(hr)) {
						SetError(sitio, L"No se pudo iniciar la conexión con Microsoft Outlook.");
						goto fin;
					}
				} else {
					m_pApp = CAddin::m_pParentApp;
				}
				pFolder= m_pApp->GetNamespace(_bstr_t("MAPI"))->GetFolderFromID(_bstr_t(sitio->m_szUID));

				if(!pFolder) {
					SetError(sitio, "No se puede obtener directorio");
					goto fin;
				}
				if(!IsRunning())
					goto fin;
				pItems= pFolder->GetItems();
				if(!pItems) {
					SetError(sitio, "No se pueden obtener los contactos.");
					goto fin;
				}
				if(!IsRunning())
					goto fin;
				CContacto *info;
				dlg->m_PEstado.SetRange32(0, pItems->GetCount());
				dlg->m_PEstado.SetPos(0);
				int pos=1;
				pContact= pItems->GetFirst();

				while( pContact) {

					if(!IsRunning())
						goto fin;

					dlg->m_PEstado.SetPos(pos++);
					info= ((CContacto *)&m_pContactos)->Actualizar(pContact, &m_pWrapper);
					if( info && info->m_nEstado==info->modificado) {
						actualizados++;
						temp.Format(L"Actualizando a %ls", info->GetNombreCompleto());
						SetMensaje(temp, sitio);
						dlg->m_SEstado.SetWindowText(temp);
						info->Modificar(pContact, &m_pWrapper);
					} else if(info) {
						temp.Format(L"Omitiendo a %ls", info->GetNombreCompleto());
						if(info->m_nEstado!=CContacto::eliminado) 
							info->m_nEstado= CContacto::guardado;
					}
					if(!IsRunning())
						goto fin;
					pContact= pItems->GetNext();
				}

				dlg->m_PEstado.SetRange32(0, m_pContactos.GetCount());
				dlg->m_PEstado.SetPos(0);

				for( int i=0; i<m_pContactos.GetCount(); i++) {

					if(!IsRunning())
						goto fin;
					dlg->m_PEstado.SetPos(i+1);
					info= (CContacto *) m_pContactos.GetAt(i);
					// ASSERT(info->m_szNombre!="Uriel");
					switch(info->m_nEstado) {
						case CContacto::guardado: // ya se le hizo algo.
							break;
						case CContacto::eliminado:
							eliminados++;
							temp.Format(L"Eliminando a %ls", info->GetNombreCompleto());
							dlg->m_SEstado.SetWindowText(temp);
							SetMensaje(temp, sitio);
							info->DeleteContacto(pItems, sitio->m_szNombre, &m_pWrapper);
							break;
						default:
							{
								nuevos++;
								CComVariant tipo(OlItemType::olContactItem);
								temp.Format(L"Agregando a %ls", info->GetNombreCompleto());
								SetMensaje(temp, sitio);
								dlg->m_SEstado.SetWindowText(temp);
								info->Modificar(pItems->Add(tipo), &m_pWrapper);
							}
							break;
					}
					if(info->m_nEstado!=CContacto::eliminado)
						info->SaveID(sitio->m_szNombre);
					if(!IsRunning())
						goto fin;
				}
				// aqui construimos el resumen.
				if(actualizados) {
					temp.Format(L"%d Actualizado%ls", actualizados, actualizados>1 ? L"s" : L"");
					SetError(sitio, temp);
				}
				if(eliminados) {
					temp.Format(L"%d Eliminado%ls", eliminados, eliminados>1 ? L"s" : L"");
					SetError(sitio, temp);
				}
				if(nuevos) {
					temp.Format(L"%d Nuevo%ls", nuevos, nuevos>1 ? L"s" : L"");
					SetError(sitio, temp);
				}
				SetError(sitio, L"Finalizado");
				dlg->m_SEstado.SetWindowText(sitio->m_szNombre + L" Actualizado.");
				sitio->SetMensaje("");
			}
			catch(_com_error &e)
			{
				SetError(sitio, e.ErrorMessage());
			}
		}
fin:
		if(IsRunning()) {
			if(sitio) {
				sitio->SetMensaje(L"");
				sitio->m_nRemaing= sitio->m_nMinutos;
			}
			dlg->Step();
		} else if(sitio) {
			sitio->SetMensaje(L"");
		}
	}
	m_pMutex.Unlock();
	::CoUninitialize();
	m_pWrapper.Uninitialize();
	return CThreadAuto::Init();
}

