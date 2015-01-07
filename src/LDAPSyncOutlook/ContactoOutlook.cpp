#include "StdAfx.h"
#include "ContactoOutlook.h"
#include "../comun/MyRegKey.h"

CContactoOutlook::CContactoOutlook(void)
{
}

CContactoOutlook::~CContactoOutlook(void)
{
}

void CContactoOutlook::DeleteContacto(const CString &sitio) {

	CMyRegKey key(DEFAULT_REG_KEY);
	key.DeleteRegKey(L"directorio\\" +  sitio, m_szLDAPCN);
}

void CContactoOutlook::FillContacto(LPSPropValue row) {

	SetValue(&row[0], m_szEmail, PR_EMAIL_ADDRESS);
	SetValue(&row[1], m_szTipo,PR_ADDRTYPE );
	SetValue(&row[2],m_szApellido ,PR_SURNAME );
	SetValue(&row[3], m_szNombre.IsEmpty() ? m_szNombreParaMostrar : m_szNombre, PR_GIVEN_NAME);
	SetValue(&row[4],m_szOrganization ,PR_COMPANY_NAME );
	SetValue(&row[5],m_szDescripcion ,PR_BUSINESS_ADDRESS_STREET );
	SetValue(&row[6], m_szCiudad, PR_BUSINESS_ADDRESS_CITY);
	SetValue(&row[7],m_szEstado , PR_BUSINESS_ADDRESS_COUNTRY);
	SetValue(&row[8],m_szPostalCode , PR_BUSINESS_ADDRESS_POSTAL_CODE);
	SetValue(&row[9],m_szTelefono , PR_BUSINESS_TELEPHONE_NUMBER);
	SetValue(&row[10],m_szFax ,PR_BUSINESS_FAX_NUMBER );
	SetValue(&row[11], m_szBeeper, PR_PAGER_TELEPHONE_NUMBER);
	SetValue(&row[12], m_szCelular, PR_MOBILE_TELEPHONE_NUMBER);
	SetValue(&row[13], m_szTelefonoCasa, PR_HOME_TELEPHONE_NUMBER);
}

CContactoOutlook *CContactoOutlook::Actualizar(LPSPropValue row, int crows) {

	if( !row || crows<16)
		return NULL;
	CString nombreCompleto(row[0].Value.lpszA);
	CContactoOutlook *info= FindByID(row[1].Value.bin.lpb,row[1].Value.bin.cb,nombreCompleto);

	if( info) {		
		// info->SetValue(info->m_szNombreCompleto, row, crows, 0);
		info->SetValue(info->m_szEmail, row, crows, 3);
		/*info->SetValue(info->m_szNombre, row, crows, 4);
		info->SetValue(info->m_szApellido, row, crows, 5);*/
		info->SetValue(info->m_szOrganization, row, crows, 6);
		info->SetValue(info->m_szDescripcion, row, crows, 7);
		info->SetValue(info->m_szCiudad, row, crows, 8);
		info->SetValue(info->m_szEstado, row, crows, 9);
		info->SetValue(info->m_szPostalCode, row, crows, 10);
		info->SetValue(info->m_szTelefono, row, crows, 11);
		info->SetValue(info->m_szFax, row, crows, 12);
		info->SetValue(info->m_szBeeper, row, crows, 13);
		info->SetValue(info->m_szCelular, row, crows, 14);
		info->SetValue(info->m_szTelefonoCasa, row, crows, 15);
	}
	return info;
}

CContactoOutlook *CContactoOutlook::FindByID(BYTE *id, int countid, const CString nombrecompleto) {

	CContactoOutlook *info= this;
	while(info) {

		if(countid>0 && countid==info->m_nBytes && info->m_pByte) {
			bool found=true;
			for( int i=0; i<countid; i++) {
				if(info->m_pByte[i]!=id[i]) {
					found=false;
					break;
				}
			}
			if(found)
				return info;
		}
		if (info->GetNombreCompleto().CompareNoCase(nombrecompleto) == 0) {
			info->m_nBytes= countid;
			if( info->m_pByte)
				delete []info->m_pByte;
			info->m_pByte= new BYTE[info->m_nBytes];
			memcpy_s(info->m_pByte,info->m_nBytes, id, countid);
			info->m_nEstado=3;
			return info;
		}
		info= (CContactoOutlook *) info->m_pSig;
	}
	return NULL;
}

void CContactoOutlook::SaveID(CString sitio) {

	if(m_nBytes==0)
		return;
	AfxGetApp()->WriteProfileBinary(L"directorio\\" +  sitio + L"\\" + m_szLDAPCN, L"idbin", m_pByte, m_nBytes);
}

void CContactoOutlook::SetValue(CString &actual, LPSPropValue row, int crows, int pos) {

	if( crows<pos)
		return;
	int valor= PROP_TYPE(row[pos].ulPropTag);
	if(PROP_TYPE(row[pos].ulPropTag)!=PT_STRING8)
		return;

	CString szNuevo(row[pos].Value.lpszA);
	if(actual.IsEmpty() && szNuevo.IsEmpty())
		return;
	if(actual==szNuevo)
		return;
	if( actual.CompareNoCase(szNuevo)) 
		m_nEstado= modificado;
}

void CContactoOutlook::SetValue(LPSPropValue row, CString value, DWORD tag) {

	row->dwAlignPad=0;
	row->ulPropTag= tag;
	row->Value.lpszW= (LPWSTR)(const wchar_t *)value;
}
