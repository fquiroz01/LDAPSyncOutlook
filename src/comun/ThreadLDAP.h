#pragma once

# include "ContactoBase.h"
# include "Sitio.h"
#include "ThreadAuto.h"
# include "afxmt.h"

class CThreadLDAP : public CThreadAuto
{
protected:
	virtual bool	Init()=0;
	bool GetLDAP(CSitio *sitio);
	CContactoBase	m_pContactos;
	virtual void SetMensaje(const CString mensaje,CSitio *sitio=NULL)=0;
public:
	CThreadLDAP(CThreadLDAP *old=NULL);
public:
	virtual ~CThreadLDAP(void);
	static CMutex m_pMutex;
};
