#include "stdafx.h"
#include "Sitio.h"
#include "MyRegKey.h"

CSitio::CSitio(void)
{
	Clean();
}

CSitio::CSitio(const CString &sitio) {

	Clean();
	CargarDesdeRegistro(sitio);
}

CSitio::~CSitio(void)
{
}

void CSitio::Clean() {

	m_szServidor=_T("");
	m_szNombre=_T("");
	m_szBase=_T("");
	m_szUsuario=_T("");
	m_szContrasena=_T("");
	m_szDestino=_T("");
	m_szUID= _T("");
	m_nPuerto= 389;
	m_nRemaing= m_nMinutos= 60;
	m_bAuth= false;
	m_bActualizar= true;
}

bool CSitio::GuardarEnRegistro() {

	CString szKey = DEFAULT_REG_KEY;
	CMyRegKey key(szKey + L"\\sitios");

	key.WriteString(m_szNombre, L"servidor", m_szServidor);
	key.WriteString(m_szNombre, L"base", m_szBase);
	key.WriteString(m_szNombre, L"usuario", m_szUsuario);
	key.WriteString(m_szNombre, L"contrasena", m_szContrasena);
	key.WriteString(m_szNombre, L"destino", m_szDestino);
	key.WriteString(m_szNombre, L"uid", m_szUID);
	key.WriteInt(m_szNombre, L"puerto", m_nPuerto);
	key.WriteInt(m_szNombre, L"minutos", m_nMinutos);
	key.WriteInt(m_szNombre, L"auth", m_bAuth);
	key.WriteInt(m_szNombre, L"actualizar", m_bActualizar);
	key.WriteInt(m_szNombre, L"faltante", m_nRemaing);
	return true;
}

int	CSitio::GetSitesCount(CStringArray *sitios/*=NULL*/) {

	CStringArray campos;
	if(sitios)
		sitios->RemoveAll();
	else sitios=&campos;
	CMyRegKey key(DEFAULT_REG_KEY);
	return key.EnumRegKey(L"sitios", *sitios);
}

bool CSitio::CargarDesdeRegistro(const CString &sitio) {

	Clean();

	CString szKey = DEFAULT_REG_KEY;
	CMyRegKey key(szKey + L"\\sitios");
	
	m_szNombre= sitio;
	m_szServidor = key.GetString(m_szNombre, L"servidor");
	if( m_szServidor.IsEmpty())
		return false;
	m_szBase = key.GetString(m_szNombre, L"base");
	m_szUsuario = key.GetString(m_szNombre, L"usuario");
	m_szContrasena = key.GetString(m_szNombre, L"contrasena");
	m_szDestino= key.GetString(m_szNombre, L"destino");
	m_szUID = key.GetString(m_szNombre, L"uid");
	m_nPuerto = key.GetInt(m_szNombre, L"puerto");
	m_nRemaing = m_nMinutos = key.GetInt(m_szNombre, L"minutos");
	m_bAuth = key.GetInt(m_szNombre, L"auth") == 1;
	m_bActualizar = key.GetInt(m_szNombre, L"actualizar") == 1;

	return true;
}

CString CSitio::GetMensaje() {
	
	CString szKey = DEFAULT_REG_KEY;
	CMyRegKey key(szKey + L"\\sitios");
	return key.GetString(m_szNombre, L"mensaje");
}

void CSitio::SetMensaje(const CString &mensaje) {
	
	CString szKey = DEFAULT_REG_KEY;
	CMyRegKey key(szKey + L"\\sitios");
	key.WriteString(m_szNombre, L"mensaje", mensaje);
}
