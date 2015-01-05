
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

// Addin.h : Declaration of the CAddin

#ifndef __ADDIN_H_
#define __ADDIN_H_
#include "stdafx.h"
#include "OutlookAddin.h"
// # include "contacto.h"
#include "resource.h"       // main symbols
# include "ProgresoDlg.h"

/////////////////////////////////////////////////////////////////////////////
// CAddin
extern _ATL_FUNC_INFO OnClickButtonInfo;
extern _ATL_FUNC_INFO OnOptionsAddPagesInfo;

class ATL_NO_VTABLE CAddin : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CAddin, &CLSID_Addin>,
	public ISupportErrorInfo,
	public IDispatchImpl<IAddin, &IID_IAddin, &LIBID_OUTLOOKADDINLib>,
	public IDispatchImpl<_IDTExtensibility2, &IID__IDTExtensibility2, &LIBID_AddInDesignerObjects>,
	public IDispEventSimpleImpl<1,CAddin,&__uuidof(Office::_CommandBarButtonEvents)>,
	public IDispEventSimpleImpl<2,CAddin,&__uuidof(Office::_CommandBarButtonEvents)>,
	public IDispEventSimpleImpl<3,CAddin,&__uuidof(Office::_CommandBarButtonEvents)>,
	public IDispEventSimpleImpl<4,CAddin,&__uuidof(Outlook::ApplicationEvents)>,
	public IDispEventSimpleImpl<5,CAddin,&__uuidof(Office::_CommandBarButtonEvents)>,
	public IDispEventSimpleImpl<6, CAddin, &__uuidof(Office::_CommandBarButtonEvents)>
{
public:

	typedef IDispEventSimpleImpl</*nID =*/ 1,CAddin, &__uuidof(Office::_CommandBarButtonEvents)> CommandButtonAcercaDeEvents;
	typedef IDispEventSimpleImpl</*nID =*/ 2,CAddin, &__uuidof(Office::_CommandBarButtonEvents)> CommandButtonActualizarEvents;
	typedef IDispEventSimpleImpl</*nID =*/ 3,CAddin, &__uuidof(Office::_CommandBarButtonEvents)> CommandMenuAcercaDeEvents;
	typedef IDispEventSimpleImpl</*nID =*/ 4,CAddin, &__uuidof(Outlook::ApplicationEvents)> AppEvents;
	typedef IDispEventSimpleImpl</*nID =*/ 5,CAddin, &__uuidof(Office::_CommandBarButtonEvents)> CommandMenuActualizarEvents;
	typedef IDispEventSimpleImpl</*nID =*/ 6, CAddin, &__uuidof(Office::_CommandBarButtonEvents)> CommandButtonConfigurarEvents;
	CAddin()
	{
		bConnected=false;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_ADDIN)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CAddin)
	COM_INTERFACE_ENTRY(IAddin)
//DEL 	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY2(IDispatch, IAddin)
	COM_INTERFACE_ENTRY(_IDTExtensibility2)
END_COM_MAP()

BEGIN_SINK_MAP(CAddin)
SINK_ENTRY_INFO(1, __uuidof(Office::_CommandBarButtonEvents),/*dispid*/ 0x01, OnClickAcercaDe, &OnClickButtonInfo)
SINK_ENTRY_INFO(2, __uuidof(Office::_CommandBarButtonEvents),/*dispid*/ 0x01, OnClickActualizar, &OnClickButtonInfo)
SINK_ENTRY_INFO(3, __uuidof(Office::_CommandBarButtonEvents),/*dispid*/ 0x01, OnClickAcercaDe, &OnClickButtonInfo)
SINK_ENTRY_INFO(4,__uuidof(Outlook::ApplicationEvents),/*dispinterface*/0xf005,OnOptionsAddPages,&OnOptionsAddPagesInfo)
SINK_ENTRY_INFO(5, __uuidof(Office::_CommandBarButtonEvents),/*dispid*/ 0x01, OnClickActualizar, &OnClickButtonInfo)
SINK_ENTRY_INFO(5, __uuidof(Office::_CommandBarButtonEvents),/*dispid*/ 0x01, OnClickConfigurar, &OnClickButtonInfo)
END_SINK_MAP()



// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);
	void __stdcall OnClickAcercaDe(IDispatch * /*Office::_CommandBarButton**/ Ctrl,VARIANT_BOOL * CancelDefault);
	void __stdcall OnClickActualizar(IDispatch * /*Office::_CommandBarButton**/ Ctrl,VARIANT_BOOL * CancelDefault);
	void __stdcall OnOptionsAddPages(IDispatch* /*PropertyPages**/ Ctrl);
	void __stdcall OnClickConfigurar(IDispatch * /*Office::_CommandBarButton**/ Ctrl, VARIANT_BOOL * CancelDefault);
// IAddin
public:
// _IDTExtensibility2
	STDMETHOD(OnConnection)(IDispatch * Application, ext_ConnectMode ConnectMode, IDispatch * AddInInst, SAFEARRAY * * custom)
	{
		m_pParentApp = Application;
		m_pParentApp->AddRef();

		m_szVersion = ((_ApplicationPtr)m_pParentApp)->GetVersion().GetBSTR();
		m_fVersion = (float)_wtof(m_szVersion);

		CComPtr < Office::_CommandBars> spCmdBars; 
	CComPtr < Office::CommandBar> spCmdBar;

	// QI() for _Application
	CComQIPtr <Outlook::_Application> spApp(Application); 
	ATLASSERT(spApp);
	// get the CommandBars interface that represents Outlook's 		//toolbars & menu 
	//items	
	m_spApp = spApp;

	CComPtr<Outlook::_Explorer> spExplorer; 	
	spExplorer= spApp->ActiveExplorer();
	if(spExplorer==NULL)
		return E_FAIL;
	// spApp->ActiveExplorer(&spExplorer);
	HRESULT hr = spExplorer->get_CommandBars(&spCmdBars);
	if(FAILED(hr))
		return hr;

	ATLASSERT(spCmdBars);
	// now we add a new toolband to Outlook
	// to which we'll add 2 buttons
	CComVariant vName("LDAPSyncOutlook");
	CComPtr <Office::CommandBar> spNewCmdBar;
	
	// position it below all toolbands
	//MsoBarPosition::msoBarTop = 1
	CComVariant vPos(1); 

	CComVariant vTemp(VARIANT_TRUE); // menu is temporary		
	CComVariant vEmpty(DISP_E_PARAMNOTFOUND, VT_ERROR);			
	//Add a new toolband through Add method
	// vMenuTemp holds an unspecified parameter
	//spNewCmdBar points to the newly created toolband
	spNewCmdBar = spCmdBars->Add(vName, vPos, vEmpty, vTemp);

	//now get the toolband's CommandBarControls
	CComPtr < Office::CommandBarControls> spBarControls;
	spBarControls = spNewCmdBar->GetControls();
	ATLASSERT(spBarControls);
	
	//MsoControlType::msoControlButton = 1
	CComVariant vToolBarType(1);
	//show the toolbar?
	CComVariant vShow(VARIANT_TRUE);

	// get CommandBarButton interface for each toolbar button
	// so we can specify button styles and stuff
	// each button displays a bitmap and caption next to it
	m_spAcercaDe = CComQIPtr < Office::_CommandBarButton>(spBarControls->Add(vToolBarType, vEmpty, vEmpty, vEmpty, vShow));
	m_spActualizar= CComQIPtr < Office::_CommandBarButton>(spBarControls->Add(vToolBarType, vEmpty, vEmpty, vEmpty, vShow));
	// m_spConfigurar = CComQIPtr < Office::_CommandBarButton>(spBarControls->Add(vToolBarType, vEmpty, vEmpty, vEmpty, vShow));

	// to set a bitmap to a button, load a 32x32 bitmap
	// and copy it to clipboard. Call CommandBarButton's PasteFace()
	// to copy the bitmap to the button face. to use
	// Outlook's set of predefined bitmap, set button's FaceId to 	//the
	// button whose bitmap you want to use
	//HBITMAP hBmp =(HBITMAP)::LoadImage(_Module.GetResourceInstance(),
	//MAKEINTRESOURCE(IDB_BITMAP1),IMAGE_BITMAP,0,0,LR_LOADMAP3DCOLORS);

	//// put bitmap into Clipboard
	//::OpenClipboard(NULL);
	//::EmptyClipboard();
	//::SetClipboardData(CF_BITMAP, (HANDLE)hBmp);
	//::CloseClipboard();
	//::DeleteObject(hBmp);
	// set style before setting bitmap
	m_spAcercaDe->PutStyle(Office::msoButtonIconAndCaption);

	// hr = spCmdButton->PasteFace();
	m_spAcercaDe->PutFaceId(682);
	
	m_spAcercaDe->PutVisible(VARIANT_TRUE);
	m_spAcercaDe->PutCaption(OLESTR("LDAPSyncOutlook"));
	m_spAcercaDe->PutEnabled(VARIANT_TRUE);
	m_spAcercaDe->PutTooltipText(OLESTR("Acerca de LDAPSyncOutlook"));
	m_spAcercaDe->PutTag(OLESTR(""));
	
	//show the toolband
	spNewCmdBar->PutVisible(VARIANT_TRUE); 
	
	//specify predefined bitmap
	// 1759
	m_spActualizar->PutStyle(Office::msoButtonIconAndCaption);
	m_spActualizar->PutFaceId(1759);
	m_spActualizar->PutCaption(OLESTR("Actualizar"));
	m_spActualizar->PutEnabled(VARIANT_TRUE);
	m_spActualizar->PutTooltipText(OLESTR("Actualiza los contactos por medio de LDAPSyncOutlook."));
	m_spActualizar->PutTag(OLESTR(""));
	m_spActualizar->PutVisible(VARIANT_TRUE);

	/*m_spConfigurar->PutStyle(Office::msoButtonIconAndCaption);
	m_spConfigurar->PutFaceId(1759);
	m_spConfigurar->PutCaption(OLESTR("Configurar"));
	m_spConfigurar->PutEnabled(VARIANT_TRUE);
	m_spConfigurar->PutTooltipText(OLESTR("Configura el servidor de contactos LDAPSyncOutlook."));
	m_spConfigurar->PutTag(OLESTR("Mi tag"));
	m_spConfigurar->PutVisible(VARIANT_TRUE);*/

	if (m_fVersion < 12) {// office 2010 en adelante ya no tienen menus

		_bstr_t bstrNewMenuText(OLESTR("Acerca de LDAPSyncOutlook..."));
		CComPtr < Office::CommandBarControls> spCmdCtrls;
		CComPtr < Office::CommandBarControls> spCmdBarCtrls;
		CComPtr < Office::CommandBarPopup> spCmdPopup;
		CComPtr < Office::CommandBarControl> spCmdCtrl;

		// get CommandBar that is Outlook's main menu
		hr = spCmdBars->get_ActiveMenuBar(&spCmdBar);
		if (FAILED(hr))
			return hr;
		// get menu as CommandBarControls 
		spCmdCtrls = spCmdBar->GetControls();
		ATLASSERT(spCmdCtrls);

		// we want to add a menu entry to Outlook's 6th(Tools) menu 	//item

		CComVariant vItem(7);
		spCmdCtrl = spCmdCtrls->GetItem(vItem);
		ATLASSERT(spCmdCtrl);

		IDispatchPtr spDisp;
		spDisp = spCmdCtrl->GetControl();

		// a CommandBarPopup interface is the actual menu item
		CComQIPtr < Office::CommandBarPopup> ppCmdPopup(spDisp);
		ATLASSERT(ppCmdPopup);

		spCmdBarCtrls = ppCmdPopup->GetControls();
		ATLASSERT(spCmdBarCtrls);

		CComVariant vMenuType(1); // type of control - menu
		CComVariant vMenuPos(6);
		CComVariant vMenuEmpty(DISP_E_PARAMNOTFOUND, VT_ERROR);
		CComVariant vMenuShow(VARIANT_TRUE); // menu should be visible
		CComVariant vMenuTemp(VARIANT_TRUE); // menu is temporary		


		CComPtr < Office::CommandBarControl> spNewMenu;
		// now create the actual menu item and add it
		spNewMenu = spCmdBarCtrls->Add(vMenuType, vMenuEmpty, vMenuEmpty, vMenuEmpty, vMenuTemp);
		ATLASSERT(spNewMenu);

		spNewMenu->PutCaption(bstrNewMenuText);
		spNewMenu->PutEnabled(VARIANT_TRUE);
		spNewMenu->PutVisible(VARIANT_TRUE);


		//we'd like our new menu item to look cool and display
		// an icon. Get menu item  as a CommandBarButton
		CComQIPtr < Office::_CommandBarButton> spCmdMenuButton(spNewMenu);
		ATLASSERT(spCmdMenuButton);
		spCmdMenuButton->PutStyle(Office::msoButtonIconAndCaption);

		// we want to use the same toolbar bitmap for menuitem too.
		// we grab the CommandBarButton interface so we can add
		// a bitmap to it through PasteFace().
		spCmdMenuButton->PutFaceId(682);
		// show the menu		
		spNewMenu->PutVisible(VARIANT_TRUE);
		m_spMenuAcercaDe = spCmdMenuButton;



		// menu de actualizar
		spCmdCtrl = spCmdCtrls->GetItem(6);
		ATLASSERT(spCmdCtrl);

		// a CommandBarPopup interface is the actual menu item
		CComQIPtr < Office::CommandBarPopup> ppCmdPopup1(spCmdCtrl->GetControl());
		ATLASSERT(ppCmdPopup1);

		spCmdBarCtrls = ppCmdPopup1->GetControls();
		ATLASSERT(spCmdBarCtrls);

		CComPtr < Office::CommandBarControl> spNewMenu1;
		// now create the actual menu item and add it
		spNewMenu1 = spCmdBarCtrls->Add(vMenuType, vMenuEmpty, vMenuEmpty, vMenuEmpty, vMenuTemp);
		ATLASSERT(spNewMenu1);

		spNewMenu1->PutCaption("Actualizar");
		spNewMenu1->PutEnabled(VARIANT_TRUE);
		spNewMenu1->PutVisible(VARIANT_TRUE);


		//we'd like our new menu item to look cool and display
		// an icon. Get menu item  as a CommandBarButton
		CComQIPtr < Office::_CommandBarButton> spCmdMenuButton1(spNewMenu1);
		ATLASSERT(spCmdMenuButton1);
		spCmdMenuButton1->PutStyle(Office::msoButtonIconAndCaption);

		// we want to use the same toolbar bitmap for menuitem too.
		// we grab the CommandBarButton interface so we can add
		// a bitmap to it through PasteFace().
		spCmdMenuButton1->PutFaceId(1759);
		// show the menu
		spNewMenu1->PutVisible(VARIANT_TRUE);
		m_spMenuActualizar = spCmdMenuButton1;

		hr = CommandMenuAcercaDeEvents::DispEventAdvise((IDispatch*)m_spMenuAcercaDe);
		if (FAILED(hr))
			return hr;

		hr = CommandMenuActualizarEvents::DispEventAdvise((IDispatch*)m_spMenuActualizar);
		if (FAILED(hr))
			return hr;
	}

	hr = CommandButtonAcercaDeEvents::DispEventAdvise((IDispatch*)m_spAcercaDe);
	if(FAILED(hr))
		return hr;

	hr = CommandButtonActualizarEvents::DispEventAdvise((IDispatch*)m_spActualizar);
	if(FAILED(hr))
		return hr;

	/*hr = CommandButtonConfigurarEvents::DispEventAdvise((IDispatch*)m_spConfigurar);
	if (FAILED(hr))
		return hr;*/

	hr = AppEvents::DispEventAdvise((IDispatch*)m_spApp,&__uuidof(Outlook::ApplicationEvents));
	if(FAILED(hr))
	{
		ATLTRACE("Failed advising to ApplicationEvents");
		return hr;
	}

	bConnected = true;
	
	return S_OK;

	}

	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom)
	{
		if(bConnected)
		{
			HRESULT hr = CommandButtonAcercaDeEvents::DispEventUnadvise((IDispatch*)m_spAcercaDe);
			if(FAILED(hr))
				return hr;
			hr = CommandButtonActualizarEvents::DispEventUnadvise((IDispatch*)m_spActualizar);
			if(FAILED(hr))
				return hr;

			/*hr = CommandButtonConfigurarEvents::DispEventUnadvise((IDispatch*)m_spConfigurar);
			if (FAILED(hr))
				return hr;*/

			hr = CommandMenuAcercaDeEvents::DispEventUnadvise((IDispatch*)m_spMenuAcercaDe);
			if(FAILED(hr))
				return hr;

			hr = CommandMenuActualizarEvents::DispEventUnadvise((IDispatch*)m_spMenuActualizar);
			if(FAILED(hr))
				return hr;

			hr = AppEvents::DispEventUnadvise((IDispatch*)m_spApp);
			if(FAILED(hr))
				return hr;

			m_pParentApp->Release();
			m_pParentApp = NULL;

			bConnected = false;
		}
		return S_OK;

	}

	STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom)
	{
		return E_NOTIMPL;
	}

	STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom)
	{
		AFX_MANAGE_STATE(AfxGetStaticModuleState());
		if(m_pDlgProgreso==NULL)
			m_pDlgProgreso= new CProgresoDlg();
		return S_OK;
	}
	STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom)
	{
		if(m_pDlgProgreso) {
			m_pDlgProgreso->Cancel();
			delete m_pDlgProgreso;
		}
		m_pDlgProgreso=NULL;
		return S_OK;
	}

	private:
		CComPtr<Office::_CommandBarButton> m_spAcercaDe; 
		CComPtr<Office::_CommandBarButton> m_spActualizar; 
		CComPtr<Office::_CommandBarButton> m_spMenuAcercaDe; 
		CComPtr<Office::_CommandBarButton> m_spMenuActualizar; 
		CComPtr<Office::_CommandBarButton> m_spConfigurar;
		CComPtr<Outlook::_Application> m_spApp;
		bool bConnected;
	public:
		static LPDISPATCH	m_pParentApp;
		static float		m_fVersion;
		static CString		m_szVersion;
		static CProgresoDlg *m_pDlgProgreso;
};

#endif //__ADDIN_H_
