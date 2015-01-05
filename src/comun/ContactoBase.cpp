#include "stdafx.h"
#include "ContactoBase.h"
#include "MyRegKey.h"

CContactoBase::CContactoBase(void)
{
	m_pSig=NULL;
	m_lpDireccion= NULL;
	m_nEstado= nuevo;
	m_pByte= NULL;
	m_nBytes=0;
	m_szTipo="SMTP";
}

CContactoBase::~CContactoBase(void)
{
	if( m_pByte) {
		delete []m_pByte;
		m_pByte=NULL;
	}
	CContactoBase *info= m_pSig, *aux;
	while( info) {
		aux= info->m_pSig;
		info->m_pSig= NULL;
		delete info;
		info=aux;
	}
}

void CContactoBase::SetValue(CString &actual, const CString nuevo) {

	if(actual.IsEmpty() && nuevo.IsEmpty())
		return;
	if(actual==nuevo)
		return;
	if( actual.CompareNoCase(nuevo)) {
		actual= nuevo;
		m_nEstado= modificado;
	}
}

CContactoBase *CContactoBase::Add(const CString &sitio, const CString &nombre) {

	if( m_szNombreCompleto.IsEmpty()) {

		m_szLDAPCN= nombre;

		Update(sitio, false);
		return this;
	} else if(m_szLDAPCN==nombre) {
		Update(sitio, true);
		return this;
	}
	CContactoBase **info= &m_pSig;
	while( *info) {
		if((*info)->m_szLDAPCN==nombre) {
			(*info)->Update(sitio, true);
			return *info;
		}
		info= &(*info)->m_pSig;
	}
	*info= new CContactoBase();
	return (*info)->Add(sitio, nombre);
}

int CContactoBase::GetCount() {

	if( m_szNombreCompleto.IsEmpty())
		return 0;
	CContactoBase *base= this;
	int i=0;
	while(base) {
		i++;
		base= base->m_pSig;
	}
	return i;
}

CContactoBase *CContactoBase::GetAt( int indice) {

	CContactoBase *info= this;
	int i=0;
	for( int i=0; info; i++, info= info->m_pSig) 

		if( i==indice)
			return info;
	return NULL;
}

void CContactoBase::ResetContent() {

	m_nEstado= nuevo;
	if( m_pByte)
		delete [] m_pByte;
	m_pByte= NULL;
	m_nBytes=0;
	m_szLDAPCN.Empty();
	m_szNombreCompleto.Empty();
	m_szApellido.Empty();
	m_szDescripcion.Empty();
	m_szNombre.Empty();
	m_szEmail.Empty();
	m_szOrganization.Empty();
	m_szCiudad.Empty();
	m_szEstado.Empty();
	m_szPostalCode.Empty();
	m_szTelefono.Empty();
	m_szFax.Empty();
	m_szBeeper.Empty();
	m_szCelular.Empty();
	m_szTelefonoCasa.Empty();
	m_lpDireccion= NULL;
	m_szID.Empty();
	if( m_pSig) {

		delete m_pSig;
		m_pSig=NULL;
	}
}

void CContactoBase::Update(const CString &sitio, bool validate) {

	CString todo= L"directorio\\"; 
	todo+=sitio+ L"\\" + m_szLDAPCN;
	CMyRegKey key(DEFAULT_REG_KEY);

	SetValue(m_szNombre,key.GetString(todo, L"givenName", L""));
	SetValue(m_szApellido, key.GetString(todo, L"sn", L""));
	SetValue(m_szEmail, key.GetString(todo, L"mail", L""));
	SetValue(m_szDescripcion, key.GetString(todo, L"description", L""));
	SetValue(m_szCiudad, key.GetString(todo, L"l", L""));
	SetValue(m_szEstado, key.GetString(todo, L"st", L""));
	SetValue(m_szPostalCode, key.GetString(todo, L"postalCode", L""));
	SetValue(m_szTelefono, key.GetString(todo, L"telephoneNumber", L""));
	SetValue(m_szFax, key.GetString(todo, L"facsimileTelephoneNumber", L""));
	SetValue(m_szBeeper, key.GetString(todo, L"pager", L""));
	SetValue(m_szCelular, key.GetString(todo, L"mobile", L""));
	SetValue(m_szTelefonoCasa, key.GetString(todo, L"homePhone", L""));
	SetValue(m_szNombreCompleto, m_szNombre + L" " + m_szApellido);
	SetValue(m_szID, key.GetString(todo, L"idmoutlook"));

	if(m_pByte)
		delete m_pByte;
	m_pByte= new BYTE[255];
	if ((m_nBytes = key.GetBytes(todo, L"idbin", m_pByte, 255)) == 0) {
		delete []m_pByte;
		m_pByte=NULL;
	}

	if(!validate) {
		m_nEstado = key.GetInt(todo, L"estado", nuevo);
	} else {
		m_nEstado= nuevo;
	}
}

int CContactoBase::CargarTodo(CString sitio) {

	CMyRegKey key(DEFAULT_REG_KEY);
	CStringArray valores;
	key.EnumRegKey(L"directorio\\" + sitio, valores);

	for( int i=0; i<valores.GetCount(); i++) 
		Add(sitio, valores[i]);
	return GetCount();
}

