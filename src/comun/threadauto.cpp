# include "stdafx.h"
#include "ThreadAuto.h"


IMPLEMENT_DYNCREATE(CThreadAuto, CWinThread)

CThreadAuto::CThreadAuto(CThreadAuto *old/*=NULL*/)
{
	m_pThreadOld= old;
	m_bRunning= false;
	m_pParametro= NULL;
	m_nPrioridad= THREAD_PRIORITY_NORMAL;
}

CThreadAuto::~CThreadAuto()
{
}

BOOL CThreadAuto::InitInstance()
{
	if (m_pThreadOld) {

		m_pThreadOld->Kill();
		m_pThreadOld = NULL;
	}
	m_bRunning=true;
	Init();
	// m_pMutex.Lock();
	if(m_bRunning) {
		m_bRunning=false;
		// m_pMutex.Unlock();
		SuspendThread();
	} else {
		// m_pMutex.Unlock();
	}
	return FALSE;
}

int CThreadAuto::ExitInstance()
{
	return CWinThread::ExitInstance();
}

BEGIN_MESSAGE_MAP(CThreadAuto, CWinThread)
END_MESSAGE_MAP()

DWORD CThreadAuto::Run(void *parametro/*=NULL*/) {

	if(m_bRunning)
		return m_nThreadID;
	m_bRunning= true;
	m_pParametro= parametro;
	CreateThread();
	SetThreadPriority(m_nPrioridad);
	return m_nThreadID;
}

bool CThreadAuto::Kill(bool wait/*=false*/) {
	// m_pMutex.Lock();
	if(m_bRunning) 
		m_bRunning= false;
	// m_pMutex.Unlock();
	ResumeThread();
	return m_bRunning;
}

bool CThreadAuto::IsRunning() {

	return m_bRunning;
}

int	CThreadAuto::GetThreadId() {

	return m_nThreadID;
}

bool CThreadAuto::Init() {

	return false;
}
