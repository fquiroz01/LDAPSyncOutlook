#pragma once

class CContactoBase
{
protected:
	CContactoBase *m_pSig;
	// funciones de outlook express
	void SetValue(CString &actual, const CString nuevo);
public:
	CContactoBase(void);
public:
	~CContactoBase(void);
	enum {
		nuevo,
		modificado,
		eliminado,
		actualizado,
		pendiente,
		sinid,
		guardado
	};
	CString m_szLDAPCN;
	CString m_szNombreCompleto;
	CString m_szNombre;
	CString m_szApellido;
	CString m_szEmail;
	CString m_szOrganization;
	CString m_szDescripcion;
	CString	m_szCiudad;
	CString	m_szEstado;
	CString	m_szPostalCode;
	CString m_szTelefono;
	CString m_szFax;
	CString m_szBeeper;
	CString m_szCelular;
	CString m_szTelefonoCasa;
	CString m_szNombreParaMostrar;

	CString	m_szTipo;
	CString	m_szID;

	int		m_nEstado;

	BYTE	*m_pByte;
	UINT	m_nBytes;

	LPARAM  m_lpDireccion;

public:
	CContactoBase *Add( const CString &sitio, const CString &nombre);
	int GetCount();
	CContactoBase *GetAt(int indice);
	void ResetContent();
	void Update(const CString &sitio, bool validate);
	int CargarTodo(CString sitio);
	CString GetNombreCompleto();
};
