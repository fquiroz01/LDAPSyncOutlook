
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


// Addin.cpp : Implementation of CAddin
#include "stdafx.h"
#include "OutlookAddin.h"
#include "Addin.h"
# include "../comun/AboutDlg.h"

/////////////////////////////////////////////////////////////////////////////
// CAddin
_ATL_FUNC_INFO OnClickButtonInfo ={CC_STDCALL,VT_EMPTY,2,{VT_DISPATCH,VT_BYREF | VT_BOOL}};
_ATL_FUNC_INFO OnOptionsAddPagesInfo={CC_STDCALL,VT_EMPTY,1,{VT_DISPATCH}};

LPDISPATCH  CAddin::m_pParentApp=NULL;
CProgresoDlg *CAddin::m_pDlgProgreso= NULL;
float		CAddin::m_fVersion=0;
CString		CAddin::m_szVersion = "";

STDMETHODIMP CAddin::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IAddin
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

void __stdcall CAddin::OnClickAcercaDe(IDispatch* /*Office::_CommandBarButton* */ Ctrl,VARIANT_BOOL * CancelDefault)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	CAboutDlg dlg;
	dlg.DoModal();
}

void __stdcall CAddin::OnClickActualizar(IDispatch* /*Office::_CommandBarButton* */ Ctrl,VARIANT_BOOL * CancelDefault)
{
	if (m_pDlgProgreso) {
		m_pDlgProgreso->UpdateSites();
	}
}

void __stdcall CAddin::OnOptionsAddPages(IDispatch *Ctrl)
{
	CComQIPtr<Outlook::PropertyPages> spPages(Ctrl);
	CComVariant varProgId(OLESTR("LDAPSyncOutlookAddin.PropPage"));
	CComBSTR bstrTitle(OLESTR("LDAPSyncOutlook"));
	HRESULT hr = spPages->Add((_variant_t)varProgId,(_bstr_t)bstrTitle);
	if(FAILED(hr))
		ATLTRACE("\nFailed addin propertypage");
}

void _stdcall CAddin::OnClickConfigurar(IDispatch* /*Office::_CommandBarButton* */ Ctrl, VARIANT_BOOL * CancelDefault) {
	
	AfxMessageBox(L"Configurar");
}