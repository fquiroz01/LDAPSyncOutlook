#pragma once
# include "../comun/ContactoBase.h"
# include <Wab.h>

class CContactoOutlook : public CContactoBase
{
public:
	CContactoOutlook(void);
public:
	~CContactoOutlook(void);
	void DeleteContacto(const CString &sitio);
	void FillContacto(LPSPropValue row);
	CContactoOutlook *Actualizar(LPSPropValue row, int crows);
	void SaveID(CString sitio);
	void SetValue(CString &actual, LPSPropValue row, int crows, int pos);
	void SetValue(LPSPropValue row, CString value, DWORD tag);
	CContactoOutlook *FindByID(BYTE *id, int countid, const CString nombrecompleto);
};
