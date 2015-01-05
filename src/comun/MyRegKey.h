#pragma once
#include "atlbase.h"

# define DEFAULT_REG_KEY	L"Software\\SDG\\LDAPSyncOutlook"


//***************************************************************************
//
// AUTHOR:  JUHU
//
// DESCRIPTION: Administrador de registro.
//
// ORIGINAL: 
//
//***************************************************************************
class CMyRegKey
{
	// Construction
public:
	//***********************************************************************
	// Name:		CMyRegKey
	// Description:	Constructor.
	// Parameters:	None.
	// Return:		None.	
	// Notes:		None.
	//***********************************************************************
	CMyRegKey(void);
	//***********************************************************************
	// Name:		CMyRegKey
	// Description:	Constructor.
	// Parameters:	HKEY parent, CString clave.
	// Return:		None.	
	// Notes:		None.
	//***********************************************************************
	CMyRegKey(HKEY parent, CString clave);
	//***********************************************************************
	// Name:		CMyRegKey
	// Description:	Constructor.
	// Parameters:	CString clave.
	// Return:		None.	
	// Notes:		None.
	//***********************************************************************
	CMyRegKey(CString clave);
	//***********************************************************************
	// Name:		CMyRegKey
	// Description:	Destructor.
	// Parameters:	None.
	// Return:		None.	
	// Notes:		None.
	//***********************************************************************
	~CMyRegKey(void);

// Attributes
protected:
	CRegKey		m_hKey;
	CString		m_szClave;
	HKEY		m_hKeyParent;

public:
	//***********************************************************************
	// Name:		MyOpen
	// Description:	Carga una llave de registro.
	// Parameters:	HKEY parent, CString clave.
	// Return:		None.	
	// Notes:		None.
	//***********************************************************************
	void MyOpen(HKEY parent, CString clave);
	//***********************************************************************
	// Name:		GetString
	// Description:	Obtiene una llave almacenda en el registro.
	// Parameters:	CString clave, CString nombre, CString defecto="".
	// Return:		CString.	
	// Notes:		None.
	//***********************************************************************
	CString GetString(CString clave, CString nombre, CString defecto=L"");
	//***********************************************************************
	// Name:		WriteString
	// Description:	Escribe una llave almacenda en el registro (text).
	// Parameters:	CString clave, CString nombre, CString valor.
	// Return:		bool.	
	// Notes:		None.
	//***********************************************************************
	bool WriteString(CString clave, CString nombre, CString valor);
	//***********************************************************************
	// Name:		GetInt
	// Description:	Obtiene una llave almacenda en el registro (entero).
	// Parameters:	CString clave, CString nombre, DWORD defecto=NULL.
	// Return:		DWORD.	
	// Notes:		None.
	//***********************************************************************
	DWORD GetInt(CString clave, CString nombre, DWORD defecto=NULL);
	//***********************************************************************
	// Name:		WriteInt
	// Description:	Escribe una llave almacenda en el registro (entero).
	// Parameters:	CString clave, CString nombre, DWORD valor.
	// Return:		bool.	
	// Notes:		None.
	//***********************************************************************
	bool WriteInt(CString clave, CString nombre, DWORD valor);
	//***********************************************************************
	// Name:		EnumRegKey
	// Description: Obtiene un cojunto de valores(Sub Carpetas) del registro 
	//				debajo de la clave suministrada.
	// Parameters:	CString clave, CStringArray &valores.
	// Return:		int.	
	// Notes:		None.
	//***********************************************************************
	int  EnumRegKey(CString clave, CStringArray &valores);
	//***********************************************************************
	// Name:		EnumRegValues
	// Description:	Obtiene un cojunto de valores(llaves) del registro debajo 
	//				de la clave suministrada.
	// Parameters:	CString clave, CStringArray &valores.
	// Return:		int.	
	// Notes:		None.
	//***********************************************************************
	int  EnumRegValues(CString clave, CStringArray &valores);
	//***********************************************************************
	// Name:		GetBytes
	// Description:	Obtiene un cojunto de valores(llaves) del registro debajo 
	//				de la clave suministrada.
	// Parameters:	CString clave,CString nombre, BYTE *byte, ULONG tamano.
	// Return:		UINT.	
	// Notes:		None.
	//***********************************************************************
	UINT GetBytes(CString clave,CString nombre, BYTE *byte, ULONG tamano);
	//***********************************************************************
	// Name:		DeleteRegKey
	// Description:	Borra un entrada de registro.
	// Parameters:	CString clave, CString nombre.
	// Return:		bool.	
	// Notes:		None.
	//***********************************************************************
	bool DeleteRegKey(CString clave, CString nombre);
	//***********************************************************************
	// Name:		GetProfileString
	// Description:	staticas entradas de registro de la aplicacion.
	// Parameters:	CString clave, CString nombre, CString defecto="".
	// Return:		bool.	
	// Notes:		None.
	//***********************************************************************
	static CString GetProfileString(CString clave, CString nombre, CString defecto=L"");
	//***********************************************************************
	// Name:		WriteProfileString
	// Description:	staticas entradas de registro de la aplicacion.
	// Parameters:	CString clave, CString nombre, CString valor.
	// Return:		bool.	
	// Notes:		None.
	//***********************************************************************
	static bool WriteProfileString(CString clave, CString nombre, CString valor);
	//***********************************************************************
	// Name:		GetProfileInt
	// Description:	staticas entradas de registro de la aplicacion.
	// Parameters:	CString clave, CString nombre, DWORD defecto=NULL.
	// Return:		bool.	
	// Notes:		None.
	//***********************************************************************
	static int GetProfileInt(CString clave, CString nombre, DWORD defecto=NULL);
	//***********************************************************************
	// Name:		WriteProfileInt
	// Description:	staticas entradas de registro de la aplicacion.
	// Parameters:	CString clave, CString nombre, DWORD valor.
	// Return:		bool.	
	// Notes:		None.
	//***********************************************************************
	static bool WriteProfileInt(CString clave, CString nombre, DWORD valor);
};
