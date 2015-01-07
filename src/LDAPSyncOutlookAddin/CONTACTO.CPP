#include "StdAfx.h"
#include "contacto.h"
# include "ContactosThread.h"
# include "MAPIPropWrapper.h"
#include "../comun/MyRegKey.h"
#include "Addin.h"


CContacto::CContacto(void)
{
}

CContacto::~CContacto(void)
{
	
}

void CContacto::Modificar(_ContactItemPtr pContact, CMAPIPropWrapper *wrapper) {
	
	if(pContact==NULL)
		return;
	try {
		pContact->put_Email1Address((_bstr_t)m_szEmail);
		pContact->put_CompanyName((_bstr_t) m_szOrganization);
		pContact->put_LastName((_bstr_t)m_szApellido);
		pContact->put_FirstName((_bstr_t)m_szNombre);
		pContact->put_Email1AddressType((_bstr_t) _T("SMTP"));
		pContact->put_BusinessAddressCity((_bstr_t)m_szCiudad);
		pContact->put_BusinessAddressCountry((_bstr_t)m_szEstado);
		pContact->put_BusinessAddressPostalCode((_bstr_t)m_szPostalCode);
		pContact->put_BusinessTelephoneNumber((_bstr_t)m_szTelefono);
		pContact->put_BusinessFaxNumber((_bstr_t)m_szFax);
		pContact->put_BusinessAddressStreet((_bstr_t) m_szDescripcion);
		pContact->put_PagerNumber((_bstr_t)m_szBeeper);
		pContact->put_MobileTelephoneNumber((_bstr_t)m_szCelular);
		pContact->put_HomeTelephoneNumber((_bstr_t) m_szTelefonoCasa);
		if (m_szApellido.IsEmpty() && m_szNombre.IsEmpty())
			pContact->put_FullName((_bstr_t)m_szNombreParaMostrar);

		pContact->Save();

		m_nEstado= guardado;

		m_szID= GetSafeEntryID(pContact, wrapper);

	} catch(_com_error &e)
	{
		ASSERT(false);
		TRACE(e.ErrorMessage());
		CContactosThread::SetError("Modificando " + m_szNombreCompleto, e.ErrorMessage());
	}
}

CString GetRightString(CString cadena, CString separador) {

	int pos = cadena.Find(separador);
	if (pos == -1)
		return cadena;
	return cadena.Mid(pos + separador.GetLength(), cadena.GetLength() - pos).Trim(L" \t");
}

CString CContacto::GetSafeEmail(_ContactItemPtr pContact, CMAPIPropWrapper *wrapper, const CString &szEmail) {

	CString email= szEmail;

	if (CAddin::m_fVersion >= 11 && CAddin::m_fVersion<15) { // el office 2013 no admite revisar la cuenta de correo
		email= pContact->GetEmail1Address().GetBSTR();
	} else if (CAddin::m_fVersion >= 10) { // office XP
	} else if (CAddin::m_fVersion == 9) { // office 2k.
		email= GetRightString(CAddin::m_szVersion, "9.0.0.");
		if(_wtoi(email)>6000) { // service pack 3
		} else {
			email= pContact->GetEmail1Address().GetBSTR();
		}
	} else {
		email= pContact->GetEmail1Address().GetBSTR();
	}
	
	return email;
}

CString CContacto::GetSafeEntryID(_ContactItemPtr pContact, CMAPIPropWrapper *wrapper) {

	CString id;
	if (CAddin::m_fVersion<12) {
			id= pContact->GetEntryID().GetBSTR();
	} else {
		COleVariant resultado;
		VARIANT contacto;
		contacto.vt=VT_DISPATCH;
		contacto.pdispVal= pContact;
		if(SUCCEEDED(wrapper->GetOneProp(contacto, PR_ENTRYID, &resultado))) {
			CByteArray arreglo;
			resultado.GetByteArrayFromVariantArray(arreglo);
			CString temp;
			for(int i=0; i<arreglo.GetCount(); i++) {
				temp.Format(arreglo[i]>15 ? L"%X" : L"0%X", arreglo[i]);
				id+= temp;
			}
		}
	}
	return id;
}

CContacto *CContacto::Actualizar(_ContactItemPtr pContact, CMAPIPropWrapper *wrapper) {

	if( pContact==NULL)
		return NULL;

	CString resultado;
	CString id;

	try {

		id= GetSafeEntryID(pContact, wrapper);
		resultado= pContact->GetFullName().GetBSTR();

		CContacto *info= FindById(id,resultado);
		if( info) {

			CString temp;
			/*temp= pContact->GetFirstName().GetBSTR();
			info->SetValue(temp, info->m_szNombre);
			temp= pContact->GetLastName().GetBSTR();
			info->SetValue(temp, info->m_szApellido);*/
			temp= pContact->GetCompanyName().GetBSTR();
			info->SetValue(temp, info->m_szOrganization);
			temp= pContact->GetBusinessAddressStreet().GetBSTR();
			info->SetValue(temp, info->m_szDescripcion);
			temp= pContact->GetBusinessAddressCity().GetBSTR();
			info->SetValue(temp, info->m_szCiudad);
			temp= pContact->GetBusinessAddressCountry().GetBSTR();
			info->SetValue(temp, info->m_szEstado);
			temp= pContact->GetBusinessAddressPostalCode().GetBSTR();
			info->SetValue(temp, info->m_szPostalCode);
			temp= pContact->GetBusinessTelephoneNumber().GetBSTR();
			info->SetValue(temp, info->m_szTelefono);
			temp= pContact->GetBusinessFaxNumber().GetBSTR();
			info->SetValue(temp, info->m_szFax);
			temp= pContact->GetPagerNumber().GetBSTR();
			info->SetValue(temp, info->m_szBeeper);
			temp= pContact->GetMobileTelephoneNumber().GetBSTR();
			info->SetValue(temp, info->m_szCelular);
			temp= pContact->GetHomeTelephoneNumber().GetBSTR();
			info->SetValue(temp, info->m_szTelefonoCasa);
			/*temp = pContact->GetFullName().GetBSTR();
			info->SetValue(temp, info->m_szNombreParaMostrar);*/
			// el correo toca revisarlos con el safeemail para evitar letrero lata.
			info->SetValue(GetSafeEmail(pContact, wrapper, info->m_szEmail), info->m_szEmail);
		}
		return info;
	}
	catch(_com_error &e)
	{
		TRACE(e.ErrorMessage());
		CContactosThread::SetError(L"Agregando " + m_szNombreCompleto, e.ErrorMessage());
	}
	return NULL;
}

CContacto *CContacto::FindById(const CString id, const CString nombrecompleto) {

	CContacto *info= this;
	while(info) {
		
		if(id==info->m_szID) {
			if(!id.IsEmpty() && !info->m_szID.IsEmpty())
				return info;
		}
		if( info->m_szNombreCompleto.CompareNoCase(nombrecompleto)==0) {
			info->m_szID= id;
			return info;
		}
		info= (CContacto *)info->m_pSig;
	}
	return NULL;
}

void CContacto::SaveID(CString sitio) {
	
	if(m_szID.IsEmpty())
		return;
	CMyRegKey key(DEFAULT_REG_KEY);
	key.WriteString(L"directorio\\" +  sitio + L"\\" + m_szLDAPCN, L"idmoutlook", m_szID);
}

void CContacto::DeleteContacto(_ItemsPtr pItems, const CString &sitio, CMAPIPropWrapper *wrapper) {
	
	if(pItems==NULL || m_szID.IsEmpty())
		return;
	
	try {

		_ContactItemPtr pContact= pItems->GetFirst();
		CString temp;

		while( pContact) {

			temp= GetSafeEntryID(pContact, wrapper);
			if(temp.IsEmpty()) 
				continue;
			if(temp==m_szID) {
				pContact->Delete();
				return;
			}
			pContact= pItems->GetNext();
		}
	} catch(_com_error &e) {
		ASSERT(false);
		TRACE(e.ErrorMessage());
		CContactosThread::SetError(L"Eliminando " + m_szNombreCompleto, e.ErrorMessage());
		return;
	}
	CMyRegKey key(DEFAULT_REG_KEY);
	key.DeleteRegKey(L"directorio\\" +  sitio, m_szLDAPCN);
}
