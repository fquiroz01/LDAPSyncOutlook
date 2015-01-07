#include "stdafx.h"
#include "ThreadLDAP.h"
#include "MyRegKey.h"
# include <winldap.h>
# include <Winber.h>

CMutex CThreadLDAP::m_pMutex;
CThreadLDAP::CThreadLDAP(CThreadLDAP *old) : CThreadAuto(old)
{
	m_nPrioridad= THREAD_PRIORITY_LOWEST;
}

CThreadLDAP::~CThreadLDAP(void)
{
}

bool CThreadLDAP::GetLDAP(CSitio *sitio) {

	ASSERT(sitio);

	LDAPMessage *res;
	ULONG	lres;
	CDialog *dlg= (CDialog *)m_pParametro;
	CString szKey = DEFAULT_REG_KEY;
	CMyRegKey key(szKey + L"\\directorio");
	LDAP	*m_ld;
	CStringArray valores;

	SetMensaje(L"Conectando con " + sitio->m_szServidor);

	key.EnumRegKey(sitio->m_szNombre, valores);

	for( int i=0; i<valores.GetCount(); i++) {

		if(!IsRunning())
			return false;
		key.WriteInt(sitio->m_szNombre + L"\\" + valores[i], L"estado", CContactoBase::eliminado);
		m_pContactos.Add(sitio->m_szNombre, valores[i]);
	}

	int rc;

	m_ld = ldap_init( (const PTCHAR)(const wchar_t *)sitio->m_szServidor, sitio->m_nPuerto);
	if ( NULL != m_ld )
	{
		if( (lres= ldap_connect( m_ld, NULL))!= LDAP_SUCCESS) {

			SetMensaje(ldap_err2string( lres), sitio);
			return false;
		}
		if( sitio->m_bAuth) {

			if(!IsRunning())
				return false;
			if ( (lres= ldap_bind_s( m_ld, (const PTCHAR)(const wchar_t *)sitio->m_szUsuario, 
				(const PTCHAR)(const wchar_t *)sitio->m_szContrasena, LDAP_AUTH_SIMPLE))
				!= LDAP_SUCCESS )
			{
				SetMensaje(ldap_err2string( lres),sitio);
				return false;
			}
		}
	} else 	{

		SetMensaje(L"Error Conectando con el servidor LDAP",sitio);
		return false;
	}
	SetMensaje(L"Iniciando busqueda (&(|(objectclass=group)(objectclass=person))(mail=*))");
	// buscamos solo los que tengan el atributo mail
	if ((lres = ldap_search(m_ld, (const PTCHAR)(const wchar_t *)sitio->m_szBase, LDAP_SCOPE_SUBTREE, _T("(&(|(objectclass=group)(objectclass=person))(mail=*))"),
		NULL, FALSE)) == -1 )
	{
		SetMensaje(ldap_err2string( lres),sitio);
		return false;
	}

	if(!IsRunning())
		return false;

	while ( (rc = ldap_result( m_ld, LDAP_RES_ANY, 0, NULL, &res ))
		== LDAP_RES_SEARCH_ENTRY )
	{
		if(!IsRunning())
			return false;

		LDAPMessage *e = ldap_first_entry( m_ld, res );
		BerElement		*ber;
		CString nombre;
		CString dn= ldap_get_dn(m_ld, e);

		SetMensaje(dn);
		nombre= sitio->m_szNombre + "\\" + dn;

		CString szID= key.GetString(nombre, L"idmoutlook");
		key.DeleteRegKey(sitio->m_szNombre, dn);
		key.WriteString(nombre, L"idmoutlook", szID);

		CString valor;

		for (wchar_t *a = ldap_first_attribute(m_ld, e, &ber); a != NULL;
			a = ldap_next_attribute( m_ld, e, ber ))
		{
			if(!IsRunning())
				return false;
			struct berval **bvals;
			if ( (bvals = ldap_get_values_len( m_ld, e, a )) != NULL )
			{
				for ( int i = 0; bvals[i] != NULL; i++ )
				{
					if(!IsRunning())
						return false;
					valor= bvals[i]->bv_val;
					valor.Trim();
					key.WriteString(nombre, a, valor);
				}
				ber_bvecfree( bvals );
			}
		}
		if(!IsRunning())
			return false;
		key.WriteInt(nombre, L"estado", CContactoBase::pendiente);
		m_pContactos.Add(sitio->m_szNombre, ldap_get_dn(m_ld, e));

		if ( ber != NULL )
			ber_free( ber, 0 );
		ldap_msgfree( res );
		if(!IsRunning())
			return false;
	} 
	if ( rc == -1 )
	{
		SetMensaje(L"No se pudo conectar con el servidor LDAP",sitio);
		return false;
	} else if (( rc = ldap_result2error( m_ld, res, 0 )) != LDAP_SUCCESS ) { // el error 9 es un error reservado
		if (rc != LDAP_REFERRAL_V2) { // esto es un warning
			SetMensaje(ldap_err2string(rc), sitio);
			return false;
		}
	}
	if(!IsRunning())
		return false;
	ldap_msgfree( res );
	CString temp;
	temp.Format(L"actualizando %d Contactos", m_pContactos.GetCount());
	SetMensaje(temp);
	return true;
}
