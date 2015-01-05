#pragma once
# include "../comun/ThreadLDAP.h"
# include "ContactoOutlook.h"

class CThreadOutlook : public CThreadLDAP
{
protected:
	virtual bool	Init();
	virtual void SetMensaje(const CString mensaje,CSitio *sitio=NULL);
public:
	CThreadOutlook(CThreadOutlook *old=NULL);
public:
	~CThreadOutlook(void);
	HRESULT LoadWABContents();
	void ClearWABLVContents();
	HRESULT Adicionar(CContactoOutlook *info);
	HRESULT DeleteEntry();
private:
	BOOL        m_bInitialized;
	HINSTANCE   m_hinstWAB;
	LPWABOPEN   m_lpfnWABOpen;
	LPADRBOOK   m_lpAdrBook; 
	LPWABOBJECT m_lpWABObject;

	// Cache Proparray of currently selected item in the list view
	LPSPropValue m_lpPropArray;
	ULONG       m_ulcValues;

	// Cache entry id of currently selected item in the listview
	SBinary     m_SB;

	void FreeProws(LPSRowSet prows);
	HRESULT Borrar(CContactoOutlook *contacto, CString sitio=L"");
	HRESULT HrGetWABTemplateID(ULONG   ulObjectType, ULONG * lpcbEID, LPENTRYID * lppEID);
	void Inicio();
};
