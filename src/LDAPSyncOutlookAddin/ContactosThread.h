#pragma once

# include "../comun/ThreadLDAP.h"
# include "MAPIPropWrapper.h"


class CContactosThread : public CThreadLDAP
{
protected:
	CMAPIPropWrapper	m_pWrapper;
	virtual bool	Init();
	_ApplicationPtr m_pApp;
	virtual void SetMensaje(const CString mensaje,CSitio *sitio=NULL);
public:
	static void SetError(CSitio *sitio, const CString &mensaje);
	static void SetError(const CString &sitio, const CString &mensaje);
	static void SetApp();
public:
	CContactosThread(CContactosThread *old=NULL);
public:
	virtual ~CContactosThread(void);
};
