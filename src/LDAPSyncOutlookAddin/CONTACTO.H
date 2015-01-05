#pragma once
# include "../comun/ContactoBase.h"
# include "MAPIPropWrapper.h"

class CContacto : public CContactoBase
{
public:
	CContacto(void);
	~CContacto(void);
	
	CString GetSafeEntryID(_ContactItemPtr pContact, CMAPIPropWrapper *wrapper);
	CString GetSafeEmail(_ContactItemPtr pContact, CMAPIPropWrapper *wrapper, const CString &szEmail);
	void DeleteContacto(_ItemsPtr pItems, const CString &sitio, CMAPIPropWrapper *wrapper);
	void Modificar(_ContactItemPtr pContact, CMAPIPropWrapper *wrapper);
	CContacto *Actualizar(_ContactItemPtr pContact, CMAPIPropWrapper *wrapper);
	CContacto *FindById(const CString id, const CString nombrecompleto);
	void SaveID(CString sitio);
};
