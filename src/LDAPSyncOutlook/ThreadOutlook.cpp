#include "StdAfx.h"
#include "ThreadOutlook.h"
#include "LDAPSyncOutlookDlg.h"
#include "../comun/MyRegKey.h"

// agregar esto como bandera // PR_COMPANY_NAME
enum {
	ieidPR_DISPLAY_NAME = 0,
	ieidPR_ENTRYID,
	ieidPR_OBJECT_TYPE,
	ieidPR_EMAIL_ADDRESS,
	ieidPR_GIVEN_NAME,
	ieidPR_SURNAME,
	ieidPR_COMPANY_NAME,
	ieidPR_BUSINESS_ADDRESS_STREET,
	ieidPR_BUSINESS_ADDRESS_CITY,
	ieidPR_BUSINESS_ADDRESS_COUNTRY,
	ieidPR_BUSINESS_ADDRESS_POSTAL_CODE,
	ieidPR_BUSINESS_TELEPHONE_NUMBER,
	ieidPR_BUSINESS_FAX_NUMBER,
	ieidPR_PAGER_TELEPHONE_NUMBER,
	ieidPR_MOBILE_TELEPHONE_NUMBER,
	ieidPR_HOME_TELEPHONE_NUMBER,
	ieidMax
};

static const SizedSPropTagArray(ieidMax, ptaEid)=
{
	ieidMax,
	{
		PR_DISPLAY_NAME,
			PR_ENTRYID,
			PR_OBJECT_TYPE,
			PR_EMAIL_ADDRESS,
			PR_GIVEN_NAME,
			PR_SURNAME,
			PR_COMPANY_NAME,
			PR_BUSINESS_ADDRESS_STREET,
			// nuevo en la version 2.1
			PR_BUSINESS_ADDRESS_CITY,
			PR_BUSINESS_ADDRESS_COUNTRY,
			PR_BUSINESS_ADDRESS_POSTAL_CODE,
			PR_BUSINESS_TELEPHONE_NUMBER,
			PR_BUSINESS_FAX_NUMBER,
			PR_PAGER_TELEPHONE_NUMBER,
			PR_MOBILE_TELEPHONE_NUMBER,
			PR_HOME_TELEPHONE_NUMBER,
	}
};


enum {
	iemailPR_DISPLAY_NAME = 0,
	iemailPR_ENTRYID,
	iemailPR_EMAIL_ADDRESS,
	iemailPR_OBJECT_TYPE,
	iemailMax
};

static const SizedSPropTagArray(iemailMax, ptaEmail)=
{
	iemailMax,
	{
		PR_DISPLAY_NAME,
			PR_ENTRYID,
			PR_EMAIL_ADDRESS,
			PR_OBJECT_TYPE
	}
};

CThreadOutlook::CThreadOutlook(CThreadOutlook *old/*=NULL*/) : CThreadLDAP(old)
{
}

void CThreadOutlook::Inicio() {

		// Here we load the WAB Object and initialize it
	m_bInitialized = FALSE;
	m_lpPropArray = NULL;
	m_ulcValues = 0;
	m_SB.cb = 0;
	m_SB.lpb = NULL;

	{
		TCHAR  szWABDllPath[MAX_PATH];
		DWORD  dwType = 0;
		ULONG  cbData = sizeof(szWABDllPath);
		HKEY hKey = NULL;

		*szWABDllPath = '\0';

		// First we look under the default WAB DLL path location in the
		// Registry. 
		// WAB_DLL_PATH_KEY is defined in wabapi.h
		//
		if (ERROR_SUCCESS == RegOpenKeyEx(HKEY_LOCAL_MACHINE, WAB_DLL_PATH_KEY, 0, KEY_READ, &hKey))
			RegQueryValueEx( hKey, L"", NULL, &dwType, (LPBYTE) szWABDllPath, &cbData);

		if(hKey) RegCloseKey(hKey);

		// if the Registry came up blank, we do a loadlibrary on the wab32.dll
		// WAB_DLL_NAME is defined in wabapi.h
		//
		TCHAR szPath[MAX_PATH];
		CString rutafinal= szWABDllPath;

		if(SUCCEEDED(SHGetFolderPath(NULL, 
                             CSIDL_PROGRAM_FILES_COMMON |CSIDL_FLAG_CREATE, 
                             NULL, 
                             0, 
							 szPath))) {
								 
								 rutafinal.Replace(L"%CommonProgramFiles%", szPath);
		}
		m_hinstWAB = LoadLibrary( rutafinal.IsEmpty() ? WAB_DLL_NAME : rutafinal);
	}

	if(m_hinstWAB)
	{
		// if we loaded the dll, get the entry point 
		//
		m_lpfnWABOpen = (LPWABOPEN) GetProcAddress(m_hinstWAB, "WABOpen");

		if(m_lpfnWABOpen)
		{
			LPSTR ruta = "";
			HRESULT hr = E_FAIL;
			WAB_PARAM wp = {0};
			wp.cbSize = sizeof(WAB_PARAM);
			wp.szFileName = ruta;

			// if we choose not to pass in a WAB_PARAM object, 
			// the default WAB file will be opened up
			//
			hr = m_lpfnWABOpen(&m_lpAdrBook,&m_lpWABObject,&wp,0);

			if(!hr)
				m_bInitialized = TRUE;
		}
	}
}

CThreadOutlook::~CThreadOutlook(void)
{
	ClearWABLVContents();
	if(m_SB.lpb)
		LocalFree(m_SB.lpb);

	if(m_bInitialized)
	{
		if(m_lpPropArray)
			m_lpWABObject->FreeBuffer(m_lpPropArray);

		if(m_lpAdrBook)
			m_lpAdrBook->Release();

		if(m_lpWABObject)
			m_lpWABObject->Release();

		/*if(m_hinstWAB)
		FreeLibrary(m_hinstWAB);*/
	}
}

void CThreadOutlook::SetMensaje(const CString mensaje,CSitio *sitio/*=NULL*/) {

	CLDAPSyncOutlookDlg *dlg = (CLDAPSyncOutlookDlg *)m_pParametro;
	if(sitio) {
		for(int i=0; i<dlg->m_LCSitios.GetItemCount(); i++) {
			if(dlg->m_LCSitios.GetItemData(i)==(DWORD_PTR)sitio) {
				dlg->m_LCSitios.SetItemText(i, 2, mensaje);
				break;
			}
		}
		dlg->m_LCResultados.InsertItem(0, sitio->m_szNombre);
		dlg->m_LCResultados.SetItemText(0, 1, mensaje);
	} else {

		CSitio *sitio= (CSitio *)dlg->m_LCSitios.GetItemData(dlg->m_nActual);
		for(int i=0; i<dlg->m_LCSitios.GetItemCount(); i++) {
			if(dlg->m_LCSitios.GetItemData(i)==(DWORD_PTR)sitio) {
				dlg->m_LCSitios.SetItemText(i, 2, mensaje);
				break;
			}
		}
	}
}

bool CThreadOutlook::Init() {

	CLDAPSyncOutlookDlg *dlg = (CLDAPSyncOutlookDlg *)m_pParametro;

	m_pMutex.Lock();
	if(dlg) {

		Inicio();

		CSitio *sitio= (CSitio *)dlg->m_LCSitios.GetItemData(dlg->m_nActual);
		dlg->m_PProgreso.SetPos(0);

		SetMensaje(L"Conectando a servidor LDAP", sitio);

		if(sitio && GetLDAP(sitio)) {
			
			if(!IsRunning())
				goto fin;

			SetMensaje(L"Cargando libreta de direcciones", sitio);
			if(SUCCEEDED(LoadWABContents())) {

				if(!IsRunning())
					goto fin;

				CString texto;
				dlg->m_PProgreso.SetRange32(0, m_pContactos.GetCount());

				for( int i=0; i<m_pContactos.GetCount(); i++) {

					if(!IsRunning())
						goto fin;
					dlg->m_PProgreso.SetPos(i+1);
					CContactoOutlook *info= (CContactoOutlook *)m_pContactos.GetAt(i);
					switch( info->m_nEstado) {
						case CContactoBase::eliminado: // eliminar
							texto.Format(L"Eliminando a %ls", info->GetNombreCompleto());
							SetMensaje(texto,sitio);
							Borrar(info, sitio->m_szNombre);
							break;
						case CContactoBase::modificado: // modificar
							texto.Format(L"Actualizando a %ls", info->GetNombreCompleto());
							SetMensaje(texto,sitio);
							Borrar(info);
							Adicionar(info);
							break;
						case CContactoBase::nuevo: // agregar
							texto.Format(L"Agregando a %ls", info->GetNombreCompleto());
							SetMensaje(texto,sitio);
							Adicionar(info);
							break;
					}
					if(!IsRunning())
						goto fin;
					if(info->m_nEstado!=CContactoBase::eliminado)
						info->SaveID(sitio->m_szNombre);
				}
				SetMensaje(L"Finalizado", sitio);
			}
		}
fin:
		if(IsRunning()) {
			if(sitio) 
				sitio->m_nRemaing= sitio->m_nMinutos;
			dlg->Step();
		}
	}
	m_pMutex.Unlock();
	return CThreadAuto::Init();
}

HRESULT CThreadOutlook::LoadWABContents()
{
	ULONG ulObjType =   0;
	LPMAPITABLE lpAB =  NULL;
	LPTSTR * lppszArray=NULL;
	ULONG cRows =       0;
	LPSRowSet lpRow =   NULL;
	LPSRowSet lpRowAB = NULL;
	LPABCONT  lpContainer = NULL;
	int cNumRows = 0;
	int nRows=0;

	HRESULT hr = E_FAIL;

	ULONG lpcbEID;
	LPENTRYID lpEID = NULL;

	// Get the entryid of the root PAB container
	//
	hr = m_lpAdrBook->GetPAB( &lpcbEID, &lpEID);

	ulObjType = 0;

	// Open the root PAB container
	// This is where all the WAB contents reside
	//
	hr = m_lpAdrBook->OpenEntry(lpcbEID,
		(LPENTRYID)lpEID,
		NULL,
		0,
		&ulObjType,
		(LPUNKNOWN *)&lpContainer);

	m_lpWABObject->FreeBuffer(lpEID);

	lpEID = NULL;

	if(HR_FAILED(hr))
		goto exit;

	// Get a contents table of all the contents in the
	// WABs root container
	//
	hr = lpContainer->GetContentsTable( 0,
		&lpAB);

	if(HR_FAILED(hr))
		goto exit;

	// Order the columns in the ContentsTable to conform to the
	// ones we want - which are mainly DisplayName, EntryID and
	// ObjectType
	// The table is guaranteed to set the columns in the order 
	// requested
	//
	hr =lpAB->SetColumns( (LPSPropTagArray)&ptaEid, 0 );

	if(HR_FAILED(hr))
		goto exit;


	// Reset to the beginning of the table
	//
	hr = lpAB->SeekRow( BOOKMARK_BEGINNING, 0, NULL );

	if(HR_FAILED(hr))
		goto exit;

	if(!IsRunning())
		goto exit;

	// Read all the rows of the table one by one
	//
	do {
		if(!IsRunning())
			goto exit;
		hr = lpAB->QueryRows(1,	0, &lpRowAB);

		if(HR_FAILED(hr))
			break;

		if(lpRowAB)
		{
			cNumRows = lpRowAB->cRows;
			if (cNumRows)
			{
				if(lpRowAB->aRow[0].lpProps[ieidPR_OBJECT_TYPE].Value.l == MAPI_MAILUSER)
				{
					if(!IsRunning())
						goto exit;

					CContactoOutlook *info=((CContactoOutlook *)&m_pContactos)->Actualizar(lpRowAB->aRow[0].lpProps, lpRowAB->aRow[0].cValues);
					if( info) {

						if(!IsRunning())
							goto exit;

						LPENTRYID lpEID = (LPENTRYID) lpRowAB->aRow[0].lpProps[ieidPR_ENTRYID].Value.bin.lpb;
						ULONG cbEID = lpRowAB->aRow[0].lpProps[ieidPR_ENTRYID].Value.bin.cb;
						// We will now take the entry-id of each object and cache it
						// on the listview item representing that object. This enables
						// us to uniquely identify the object later if we need to
						LPSBinary lpSB = NULL;
						m_lpWABObject->AllocateBuffer(sizeof(SBinary), (LPVOID *) &lpSB);

						if(lpSB)
						{
							m_lpWABObject->AllocateMore(cbEID, lpSB, (LPVOID *) &(lpSB->lpb));

							if(!lpSB->lpb)
							{
								m_lpWABObject->FreeBuffer(lpSB);
								continue;
							}
							CopyMemory(lpSB->lpb, lpEID, cbEID);
							lpSB->cb = cbEID;
							info->m_lpDireccion= (LPARAM) lpSB;
						}
						if(!IsRunning())
							goto exit;

						if( info && info->m_nEstado==info->modificado) {
							// actualizados++;
							// temp.Format("Actualizando a %ls %ls", info->m_szNombre, info->m_szApellido);
							// dlg->m_SEstado.SetWindowText(temp);
							info->FillContacto(lpRowAB->aRow[0].lpProps);
						} else if(info) {
							if(info->m_nEstado!=CContactoBase::eliminado) 
								info->m_nEstado= CContactoBase::guardado;
						}
						if(!IsRunning())
							goto exit;
					}
				}
			}
			FreeProws(lpRowAB );
		}
	} while ( SUCCEEDED(hr) && cNumRows && lpRowAB)  ;

exit:

	if ( lpContainer )
		lpContainer->Release();

	if ( lpAB )
		lpAB->Release();

	return hr;
}


// Clears the contents of the specified ListView
//
void CThreadOutlook::ClearWABLVContents()
{
	for(int i=0;i<m_pContactos.GetCount();i++)
	{
		LPSBinary lpSB = (LPSBinary) m_pContactos.GetAt(i)->m_lpDireccion;
		if( lpSB)
			m_lpWABObject->FreeBuffer(lpSB);
	}
}

void CThreadOutlook::FreeProws(LPSRowSet prows)
{
	ULONG		irow;
	if (!prows)
		return;
	for (irow = 0; irow < prows->cRows; ++irow)
		m_lpWABObject->FreeBuffer(prows->aRow[irow].lpProps);
	m_lpWABObject->FreeBuffer(prows);
}


enum {
	icrPR_DEF_CREATE_MAILUSER = 0,
	icrPR_DEF_CREATE_DL,
	icrMax
};

const SizedSPropTagArray(icrMax, ptaCreate)=
{
	icrMax,
	{
		PR_DEF_CREATE_MAILUSER,
			PR_DEF_CREATE_DL,
	}
};


// Gets the WABs default Template ID for MailUsers
// or DistLists. These Template IDs are needed for creating
// new mailusers and distlists
//
HRESULT CThreadOutlook::HrGetWABTemplateID(ULONG   ulObjectType,
								 ULONG * lpcbEID,
								 LPENTRYID * lppEID)
{
	LPABCONT lpContainer = NULL;
	HRESULT hr  = hrSuccess;
	SCODE sc = ERROR_SUCCESS;
	ULONG ulObjType = 0;
	ULONG cbWABEID = 0;
	LPENTRYID lpWABEID = NULL;
	LPSPropValue lpCreateEIDs = NULL;
	LPSPropValue lpNewProps = NULL;
	ULONG cNewProps;
	ULONG nIndex;

	if (    (!m_lpAdrBook) ||
		((ulObjectType != MAPI_MAILUSER) && (ulObjectType != MAPI_DISTLIST)) )
	{
		hr = MAPI_E_INVALID_PARAMETER;
		goto out;
	}

	*lpcbEID = 0;
	*lppEID = NULL;

	if (HR_FAILED(hr = m_lpAdrBook->GetPAB( &cbWABEID,
		&lpWABEID)))
	{
		goto out;
	}

	if (HR_FAILED(hr = m_lpAdrBook->OpenEntry(cbWABEID,     // size of EntryID to open
		lpWABEID,     // EntryID to open
		NULL,         // interface
		0,            // flags
		&ulObjType,
		(LPUNKNOWN *)&lpContainer)))
	{
		goto out;
	}

	// Opened PAB container OK

	// Get us the default creation entryids
	if (HR_FAILED(hr = lpContainer->GetProps(   (LPSPropTagArray)&ptaCreate,
		0,
		&cNewProps,
		&lpCreateEIDs)  )   )
	{
		goto out;
	}

	// Validate the properties
	if (    lpCreateEIDs[icrPR_DEF_CREATE_MAILUSER].ulPropTag != PR_DEF_CREATE_MAILUSER ||
		lpCreateEIDs[icrPR_DEF_CREATE_DL].ulPropTag != PR_DEF_CREATE_DL)
	{
		goto out;
	}

	if(ulObjectType == MAPI_DISTLIST)
		nIndex = icrPR_DEF_CREATE_DL;
	else
		nIndex = icrPR_DEF_CREATE_MAILUSER;

	*lpcbEID = lpCreateEIDs[nIndex].Value.bin.cb;

	m_lpWABObject->AllocateBuffer(*lpcbEID, (LPVOID *) lppEID);

	if (sc != S_OK)
	{
		hr = MAPI_E_NOT_ENOUGH_MEMORY;
		goto out;
	}
	CopyMemory(*lppEID,lpCreateEIDs[nIndex].Value.bin.lpb,*lpcbEID);

out:
	if (lpCreateEIDs)
		m_lpWABObject->FreeBuffer(lpCreateEIDs);

	if (lpContainer)
		lpContainer->Release();

	if (lpWABEID)
		m_lpWABObject->FreeBuffer(lpWABEID);

	return hr;
}


HRESULT CThreadOutlook::Adicionar( CContactoOutlook *info) { // const char * email, const char * nombre, const char * apellido, const char *descripcion) {

	ASSERT(info);
	ULONG cbEID=0;
	LPENTRYID lpEID=NULL;

	HRESULT hr = hrSuccess;
	ULONG cbTplEID = 0;
	LPENTRYID lpTplEID = NULL;

	// Get the template id which is needed to create the
	// new object
	//
	if(HR_FAILED(hr = HrGetWABTemplateID(MAPI_MAILUSER,
		&cbTplEID,
		&lpTplEID)))
		return hr;

	// Display the New Entry dialog to create the new entry
	//

	ULONG lpcbEntryID;
	ENTRYID *lpEntryID;
	LPUNKNOWN lpUnk = NULL;
	ULONG ulObjType = NULL;
	LPMAPIPROP      lpNewEntry = NULL;

	if (HR_FAILED(hr = m_lpAdrBook->GetPAB(
		&lpcbEntryID,
		&lpEntryID
		)))
		return hr;


	if (HR_FAILED(hr = m_lpAdrBook->OpenEntry(
		lpcbEntryID,
		lpEntryID,
		NULL,
		MAPI_MODIFY,
		&ulObjType,
		&lpUnk
		))) 
		return hr;

	IABContainer *lpContainer = static_cast <IABContainer *>(lpUnk);

	hr= lpContainer->CreateEntry(cbTplEID ,
		lpTplEID,
		CREATE_REPLACE,
		&lpNewEntry);

	if (S_OK == hr && lpNewEntry)
	{
		// Ok, now you need to set some properties on the new Entry.
		const unsigned long cProps = ieidMax-2;
		SPropValue   aPropsMesg[cProps];
		LPSPropProblemArray   lpPropProblems = NULL;
		info->FillContacto(aPropsMesg);

		// Set the properties on the object.
		hr = lpNewEntry->SetProps(cProps, aPropsMesg, &lpPropProblems);      

		// Explictly save the changes to the new entry.
		hr = lpNewEntry->SaveChanges(NULL);

		if (MAPI_E_COLLISION == hr)

			return hr;
	}
	return hr;
}

HRESULT CThreadOutlook::Borrar(CContactoOutlook *contacto, CString sitio) {

	ASSERT(contacto);
	LPSBinary lpSB = (LPSBinary) contacto->m_lpDireccion;
	if(m_SB.lpb)
		LocalFree(m_SB.lpb);
	if(lpSB==NULL)
		return S_FALSE;
	m_SB.cb = lpSB->cb;
	m_SB.lpb = (LPBYTE) LocalAlloc(LMEM_ZEROINIT, m_SB.cb);
	if(m_SB.lpb)
		CopyMemory(m_SB.lpb, lpSB->lpb, m_SB.cb);
	else
		m_SB.cb = 0;
	if(SUCCEEDED(DeleteEntry())) {

		if(contacto->m_nEstado==CContactoOutlook::eliminado && !sitio.IsEmpty()) {
			CMyRegKey key(DEFAULT_REG_KEY);
			key.DeleteRegKey(L"directorio\\" +  sitio, contacto->m_szLDAPCN);
		}
		return S_OK;
	}
	return S_FALSE;
}

// Delete an entry from the WAB
//
HRESULT CThreadOutlook::DeleteEntry()
{
	HRESULT hr = hrSuccess;
	ULONG cbWABEID = 0;
	LPENTRYID lpWABEID = NULL;
	LPABCONT lpWABCont = NULL;
	ULONG ulObjType;
	SBinaryArray SBA;

	hr = m_lpAdrBook->GetPAB( &cbWABEID,
		&lpWABEID);
	if(HR_FAILED(hr))
		goto out;

	hr = m_lpAdrBook->OpenEntry(  cbWABEID,     // size of EntryID to open
		lpWABEID,     // EntryID to open
		NULL,         // interface
		0,            // flags
		&ulObjType,
		(LPUNKNOWN *)&lpWABCont);

	if(HR_FAILED(hr))
		goto out;

	SBA.cValues = 1;
	SBA.lpbin = &m_SB;

	hr = lpWABCont->DeleteEntries((LPENTRYLIST) &SBA, 0);

	if(m_lpPropArray)
		m_lpWABObject->FreeBuffer(m_lpPropArray);

	m_lpPropArray = NULL;
	m_ulcValues = 0;

out:
	if(lpWABCont)
		lpWABCont->Release();

	if(lpWABEID)
		m_lpWABObject->FreeBuffer(lpWABEID);

	return hr;
}
