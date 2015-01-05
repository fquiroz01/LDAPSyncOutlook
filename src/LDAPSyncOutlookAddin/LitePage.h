
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

// LitePage.h : Declaration of the CLitePage

#ifndef __LITEPAGE_H_
#define __LITEPAGE_H_

#include "resource.h"       // main symbols
#include <atlctl.h>
// #import "C:\Archivos de programa\Archivos comunes\designer\MSADDNDR.dll" raw_interfaces_only
// #import "c:\Archivos de programa\Microsoft Office\Office12\MSOUTL.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids 


/////////////////////////////////////////////////////////////////////////////
// CLitePage
class ATL_NO_VTABLE CLitePage : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public IDispatchImpl<ILitePage, &IID_ILitePage, &LIBID_OUTLOOKADDINLib>,
	public CComCompositeControl<CLitePage>,
	public IPersistStreamInitImpl<CLitePage>,
	public IPersistPropertyBagImpl<CLitePage>,
	public IOleControlImpl<CLitePage>,
	public IOleObjectImpl<CLitePage>,
	public IOleInPlaceActiveObjectImpl<CLitePage>,
	public IViewObjectExImpl<CLitePage>,
	public IOleInPlaceObjectWindowlessImpl<CLitePage>,
	public ISupportErrorInfo,
	public CComCoClass<CLitePage, &CLSID_LitePage>,
	public IDispatchImpl<PropertyPage,&__uuidof(PropertyPage),&LIBID_OUTLOOKADDINLib>
{
public:
	CLitePage()
	{
		m_bWindowOnly = TRUE;
		CalcExtent(m_sizeExtent);
	}

DECLARE_REGISTRY_RESOURCEID(IDR_LITEPAGE)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CLitePage)
	COM_INTERFACE_ENTRY(ILitePage)
	COM_INTERFACE_ENTRY2(IDispatch,ILitePage)
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
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(PropertyPage)
END_COM_MAP()

BEGIN_PROP_MAP(CLitePage)
	PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
	PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
	// Example entries
	// PROP_ENTRY("Property Description", dispid, clsid)
	// PROP_PAGE(CLSID_StockColorPage)
END_PROP_MAP()

BEGIN_MSG_MAP(CLitePage)
	CHAIN_MSG_MAP(CComCompositeControl<CLitePage>)
END_MSG_MAP()
// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

BEGIN_SINK_MAP(CLitePage)
	//Make sure the Event Handlers have __stdcall calling convention
END_SINK_MAP()
STDMETHOD(SetClientSite)(IOleClientSite *pSite);
STDMETHOD(GetControlInfo)(LPCONTROLINFO pCI);

	STDMETHOD(OnAmbientPropertyChange)(DISPID dispid)
	{
		if (dispid == DISPID_AMBIENT_BACKCOLOR)
		{
			SetBackgroundColorFromAmbient();
			FireViewChange();
		}
		return IOleControlImpl<CLitePage>::OnAmbientPropertyChange(dispid);
	}



// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid)
	{
		static const IID* arr[] = 
		{
			&IID_ILitePage,
		};
		for (int i=0; i<sizeof(arr)/sizeof(arr[0]); i++)
		{
			if (InlineIsEqualGUID(*arr[i], riid))
				return S_OK;
		}
		return S_FALSE;
	}

// IViewObjectEx
	DECLARE_VIEW_STATUS(0)

// ILitePage
public:

	enum { IDD = IDD_LITEPAGE };
// PropertyPage
	STDMETHOD(GetPageInfo)(BSTR * HelpFile, LONG * HelpContext)
	{
		ATLTRACE("GetPageInfo");
		if (HelpFile == NULL)
			return E_POINTER;
			
		if (HelpContext == NULL)
			return E_POINTER;
			
		return E_NOTIMPL;
	}
	STDMETHOD(get_Dirty)(VARIANT_BOOL * Dirty)
	{
		ATLTRACE("GetDirty");
		if (Dirty == NULL)
			return E_POINTER;
			
		return E_NOTIMPL;
	}
	STDMETHOD(Apply)()
	{
		ATLTRACE("Apply");
		return E_NOTIMPL;
	}
};

#endif //__LITEPAGE_H_
