#pragma once

class CSitio
{
public:
	CString		m_szServidor;
	CString		m_szNombre;
	CString		m_szBase;
	CString		m_szUsuario;
	CString		m_szContrasena;
	CString		m_szDestino;
	int			m_nPuerto;
	int			m_nMinutos;
	bool		m_bAuth;
	bool		m_bActualizar;
	CString		m_szUID;
	int			m_nRemaing;
public:
	CSitio(void);
	CSitio(const CString &sitio);
public:
	~CSitio(void);
	void Clean();
	bool	GuardarEnRegistro();
	bool	CargarDesdeRegistro(const CString &sitio);
	static int	GetSitesCount(CStringArray *sitios=NULL);
	CString GetMensaje();
	void SetMensaje(const CString &mensaje);
};
