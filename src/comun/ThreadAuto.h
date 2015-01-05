#pragma once

class CThreadAuto : public CWinThread
{
	DECLARE_DYNCREATE(CThreadAuto)
public:
	CThreadAuto(CThreadAuto	*old=NULL);
	virtual ~CThreadAuto();
public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();
protected:
	DECLARE_MESSAGE_MAP()
	// aqui es la version vieja de hilos.

private:
	CThreadAuto		*m_pThreadOld;
protected:
	volatile bool	m_bRunning;
	virtual bool	Init();
	void	*m_pParametro;
	int		m_nPrioridad;
public:
	DWORD		Run(void *parametro=NULL);
	bool		Kill(bool wait=false);
	bool		IsRunning();
	int			GetThreadId();
};
