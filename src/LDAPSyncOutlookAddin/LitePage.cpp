
////////////////////////////////////////////////////////////////////////////
//	Copyright : Amit Dey 2002
//
//	Email :amitdey@joymail.com
//
//	This code may be used in compiled form in any way you desire. This
//	file may be redistributed unmodified by any means PROVIDING it is 
//	not sold for profit without the authors written consent, and 
//	providing that this notice and the authors name is included.
//
//	This file is provided 'as is' with no expressed or implied warranty.
//	The author accepts no liability if it causes any damage to your computer.
//
//	Do expect bugs.
//	Please let me know of any bugs/mods/improvements.
//	and I will try to fix/incorporate them into this file.
//	Enjoy!
//
//	Description : Outlook2K Addin 
////////////////////////////////////////////////////////////////////////////

// LitePage.cpp : Implementation of CLitePage

#include "stdafx.h"
#include "OutlookAddin.h"
#include "LitePage.h"

/////////////////////////////////////////////////////////////////////////////
// CLitePage

STDMETHODIMP CLitePage::SetClientSite(IOleClientSite *pSite)
{
	ATLTRACE("SetClientSite");
	if(pSite!=NULL)
	IOleObjectImpl<CLitePage>::SetClientSite(pSite);
	
	return S_OK;

	
}
STDMETHODIMP CLitePage::GetControlInfo(LPCONTROLINFO pCI)
{
	return S_OK;	
}