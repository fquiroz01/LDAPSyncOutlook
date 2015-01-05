
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

// PropPage.h : Declaration of the CPropPage

#ifndef __PROPPAGE_H_
#define __PROPPAGE_H_

//#include "resource.h"       // main symbols
//# include <comdef.h>
//#include <atlctl.h>
# include "../comun/Sitio.h"
//
//// #import "C:\Archivos de programa\Archivos comunes\designer\MSADDNDR.dll" raw_interfaces_only
//// #import "c:\Archivos de programa\Microsoft Office\Office\MSOUTL.olb" raw_interfaces_only
//// # include <atlcontrols.h>
//// using namespace ATLControls;
///////////////////////////////////////////////////////////////////////////////
//// CPropPage

class ATL_NO_VTABLE CPropPage : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public IDispatchImpl<IPropPage, &IID_IPropPage, &LIBID_OUTLOOKADDINLib>,
	public CComCompositeControl<CPropPage>,
	public IPersistStreamInitImpl<CPropPage>,
	public IOleControlImpl<CPropPage>,
	public IOleObjectImpl<CPropPage>,
	public IOleInPlaceActiveObjectImpl<CPropPage>,
	public IViewObjectExImpl<CPropPage>,
	public IOleInPlaceObjectWindowlessImpl<CPropPage>,
	public IPersistStorageImpl<CPropPage>,
	public ISpecifyPropertyPagesImpl<CPropPage>,
	public IQuickActivateImpl<CPropPage>,
	public IDataObjectImpl<CPropPage>,
	public IProvideClassInfo2Impl<&CLSID_PropPage, NULL, &LIBID_OUTLOOKADDINLib>,
	public CComCoClass<CPropPage, &CLSID_PropPage>,
	public IPersistPropertyBagImpl<CPropPage>,
	public IDispatchImpl<PropertyPage,&__uuidof(PropertyPage),&LIBID_OUTLOOKADDINLib>
{
public:
	CPropPage()
	{
		m_pActual=NULL;
		m_bWindowOnly = TRUE;
		CalcExtent(m_sizeExtent);
	}

	DECLARE_REGISTRY_RESOURCEID(IDR_PROPPAGE)
//
	DECLARE_PROTECT_FINAL_CONSTRUCT()
//
	BEGIN_COM_MAP(CPropPage)
		COM_INTERFACE_ENTRY(IPropPage)
		COM_INTERFACE_ENTRY2(IDispatch,IPropPage)
		COM_INTERFACE_ENTRY(IViewObjectEx)
		COM_INTERFACE_ENTRY(IViewObject2)
		COM_INTERFACE_ENTRY(IViewObject)
		COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
		COM_INTERFACE_ENTRY(IOleInPlaceObject)
		COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
		COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
		COM_INTERFACE_ENTRY(IOleControl)
		COM_INTERFACE_ENTRY(IOleObject)
		COM_INTERFACE_ENTRY(IPersistStreamInit)
		COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
		COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
		COM_INTERFACE_ENTRY(IQuickActivate)
		COM_INTERFACE_ENTRY(IPersistStorage)
		COM_INTERFACE_ENTRY(IDataObject)
		COM_INTERFACE_ENTRY(IProvideClassInfo)
		COM_INTERFACE_ENTRY(IProvideClassInfo2)
		COM_INTERFACE_ENTRY(IPersistPropertyBag) 
		COM_INTERFACE_ENTRY(PropertyPage)
	END_COM_MAP()
//
	BEGIN_PROP_MAP(CPropPage)
		PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
		PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
		// Example entries
		// PROP_ENTRY("Property Description", dispid, clsid)
		// PROP_PAGE(CLSID_StockColorPage)
	END_PROP_MAP()
//
	BEGIN_MSG_MAP(CPropPage)
		COMMAND_HANDLER(IDC_BACEPTAR, BN_CLICKED, OnBnClickedBaceptar)
		COMMAND_HANDLER(IDC_BCANCELAR, BN_CLICKED, OnBnClickedBcancelar)
		COMMAND_HANDLER(IDC_BEXPLORAR, BN_CLICKED, OnBnClickedBexplorar)
		COMMAND_HANDLER(IDC_BAVANZADO, BN_CLICKED, OnBnClickedBavanzado)
		COMMAND_HANDLER(2000, BN_CLICKED, OnEditar)
		COMMAND_HANDLER(2001, BN_CLICKED, OnEliminar)
		NOTIFY_HANDLER(IDC_LSITIOS, NM_DBLCLK, OnNMDblclkLsitios)
		NOTIFY_HANDLER(IDC_LSITIOS, NM_RCLICK, OnNMRclickLsitios)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		NOTIFY_HANDLER(IDC_LSITIOS, LVN_KEYDOWN, OnLvnKeydownLsitios)
		COMMAND_HANDLER(IDC_BAYUDA, BN_CLICKED, OnBnClickedBayuda)
		MESSAGE_HANDLER(WM_TIMER, OnTimer)
		CHAIN_MSG_MAP(CComCompositeControl<CPropPage>)
	END_MSG_MAP()
//	// Handler prototypes:

	//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
//
	BEGIN_SINK_MAP(CPropPage)
		//Make sure the Event Handlers have __stdcall calling convention
	END_SINK_MAP()
//
	STDMETHOD(GetControlInfo)(LPCONTROLINFO pCI);
	STDMETHOD(SetClientSite)(IOleClientSite *pSite);

	STDMETHOD(OnAmbientPropertyChange)(DISPID dispid)
	{
		if (dispid == DISPID_AMBIENT_BACKCOLOR)
		{
			SetBackgroundColorFromAmbient();
			FireViewChange();
		}
		return IOleControlImpl<CPropPage>::OnAmbientPropertyChange(dispid);
	}
//
//
//
//	// IViewObjectEx
	DECLARE_VIEW_STATUS(0)
//
//	// IPropPage
public:

	enum { IDD = IDD_PROPPAGE };
//	// PropertyPage
	STDMETHOD(GetPageInfo)(BSTR * HelpFile, LONG * HelpContext);
	STDMETHOD(get_Dirty)(VARIANT_BOOL * Dirty);
	STDMETHOD(Apply)();	

	STDMETHOD(raw_GetPageInfo)(BSTR * HelpFile,long * HelpContext );
	STDMETHOD(raw_Apply)( );
//
private:
	CComVariant m_fDirty;
	CComPtr<PropertyPageSite> m_pPropPageSite; 
	CSitio *m_pActual;
	CSitio m_pAvanzado;
//
	LRESULT OnBnClickedBaceptar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnBnClickedBcancelar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnBnClickedBexplorar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnBnClickedBavanzado(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnEditar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnEliminar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnNMDblclkLsitios(int /*idCtrl*/, LPNMHDR pNMHDR, BOOL& /*bHandled*/);
	LRESULT OnNMRclickLsitios(int /*idCtrl*/, LPNMHDR pNMHDR, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnLvnKeydownLsitios(int /*idCtrl*/, LPNMHDR pNMHDR, BOOL& /*bHandled*/);

	CString m_szUID;
	LRESULT OnBnClickedBayuda(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnTimer(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
};

#endif //__PROPPAGE_H_
