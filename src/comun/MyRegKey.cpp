#include "StdAfx.h"
#include "myregkey.h"

CMyRegKey::CMyRegKey(void)
{
	m_hKeyParent= NULL;
}

CMyRegKey::CMyRegKey(HKEY parent, CString clave) {

	m_hKeyParent= parent;
	m_szClave= clave+L"\\";
}

CMyRegKey::CMyRegKey(CString clave) {
	m_hKeyParent= HKEY_CURRENT_USER;
	m_szClave= clave+L"\\";
}

CMyRegKey::~CMyRegKey(void)
{
}

void CMyRegKey::MyOpen(HKEY parent, CString clave) {
	
	m_hKeyParent= parent;
	m_szClave= clave+L"\\";
}

CString CMyRegKey::GetString(CString clave, CString nombre, CString defecto) {

	ASSERT(!m_szClave.IsEmpty());

	if( m_hKey.Open(m_hKeyParent, m_szClave+ clave, KEY_READ)==ERROR_SUCCESS) {
	
		ULONG tamano=MAX_PATH;
		wchar_t szClave[MAX_PATH];
		if( m_hKey.QueryStringValue(nombre, szClave,&tamano)==ERROR_SUCCESS)
			defecto= szClave;
		m_hKey.Close();
	}
	return defecto;
}

bool CMyRegKey::WriteString(CString clave, CString nombre, CString valor) {

	ASSERT(!m_szClave.IsEmpty());

	bool retorno= false;
	if( m_hKey.Create(m_hKeyParent, m_szClave + clave)==ERROR_SUCCESS) {
	
		retorno= m_hKey.SetStringValue(nombre, valor)==ERROR_SUCCESS;
		m_hKey.Close();
	}
	return retorno;
}

DWORD CMyRegKey::GetInt(CString clave, CString nombre, DWORD defecto) {

	ASSERT(!m_szClave.IsEmpty());
	if( m_hKey.Open(m_hKeyParent, m_szClave + clave, KEY_READ)==ERROR_SUCCESS) {
		
		DWORD dwclave=0;
		if( m_hKey.QueryDWORDValue(nombre, dwclave)==ERROR_SUCCESS)
			defecto= dwclave;
		m_hKey.Close();
	}
	return defecto;
}

bool CMyRegKey::WriteInt(CString clave, CString nombre, DWORD valor) {

	ASSERT(!m_szClave.IsEmpty());
	
	bool retorno= false;
	if( m_hKey.Create(m_hKeyParent, m_szClave + clave)==ERROR_SUCCESS) {
	
		retorno= m_hKey.SetDWORDValue(nombre, valor)==ERROR_SUCCESS;
		m_hKey.Close();
	}
	return retorno;
}


int  CMyRegKey::EnumRegKey(CString clave, CStringArray &valores) {

	if( m_szClave.IsEmpty())
		return 0;
	valores.RemoveAll();

	if( m_hKey.Open(m_hKeyParent, m_szClave + clave, KEY_READ)==ERROR_SUCCESS) {

		wchar_t	szclave[256];
		DWORD	tamano= 255;
		int		nClave=0;

		while( ERROR_NO_MORE_ITEMS!= m_hKey.EnumKey(nClave, szclave, &tamano)) {

			valores.Add(szclave);
			tamano=255;
			nClave++;
		}
		m_hKey.Close();
	}
	return valores.GetCount();
}

int  CMyRegKey::EnumRegValues(CString clave, CStringArray &valores) {

	if( m_szClave.IsEmpty())
		return 0;
	valores.RemoveAll();

	if( m_hKey.Open(m_hKeyParent, m_szClave + clave, KEY_READ | KEY_ENUMERATE_SUB_KEYS)==ERROR_SUCCESS) {

		wchar_t	szclave[256];
		DWORD	tamano= 255;
		int		nClave=0;

		while( ERROR_NO_MORE_ITEMS!= RegEnumValue( m_hKey.m_hKey, nClave, szclave, &tamano, NULL, NULL, NULL, NULL)) {	

			valores.Add(szclave);
			tamano=255;
			nClave++;
		}
		m_hKey.Close();
	}
	return valores.GetCount();
}

bool CMyRegKey::DeleteRegKey(CString clave, CString nombre) {

	if( m_szClave.IsEmpty())
		return false;
	bool retorno=false;
	if( m_hKey.Open(m_hKeyParent, m_szClave + clave, KEY_READ | KEY_WRITE)==ERROR_SUCCESS) {
		retorno= true;
		if( m_hKey.DeleteSubKey(nombre)!=ERROR_SUCCESS)
			if( m_hKey.RecurseDeleteKey(nombre)!=ERROR_SUCCESS)
				retorno= false;
		m_hKey.Close();
	}
	return retorno;
}

UINT CMyRegKey::GetBytes(CString clave, CString nombre, BYTE *byte, ULONG tamano) {

	if( m_szClave.IsEmpty())
		return 0;
	if( m_hKey.Open(m_hKeyParent, m_szClave + clave, KEY_READ)==ERROR_SUCCESS) {
		if( m_hKey.QueryBinaryValue(nombre, byte, &tamano)!=ERROR_SUCCESS) {
			tamano=0;
		}
		m_hKey.Close();
	} else tamano=0;
	return tamano;
}

CString CMyRegKey::GetProfileString(CString clave, CString nombre, CString defecto/*=""*/) {

	CMyRegKey key(DEFAULT_REG_KEY);
	return key.GetString(clave, nombre, defecto);
}

bool CMyRegKey::WriteProfileString(CString clave, CString nombre, CString valor) {
	
	CMyRegKey key(DEFAULT_REG_KEY);
	return key.WriteString(clave, nombre, valor);
}

int CMyRegKey::GetProfileInt(CString clave, CString nombre, DWORD defecto/*=NULL*/) {

	CMyRegKey key(DEFAULT_REG_KEY);
	return key.GetInt(clave, nombre, defecto);
}

bool CMyRegKey::WriteProfileInt(CString clave, CString nombre, DWORD valor) {

	CMyRegKey key(DEFAULT_REG_KEY);
	return key.WriteInt(clave, nombre, valor);
}
