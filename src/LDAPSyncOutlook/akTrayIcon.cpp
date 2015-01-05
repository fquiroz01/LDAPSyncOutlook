#include "StdAfx.h"

#include "akTrayIcon.h"
#include <assert.h>
#include <afxmt.h>


//
//  CakTrayIcon
//
const UINT CakTrayIcon::WM_TRAYICONNOTIFY = RegisterWindowMessage(_T("Tray.{47BCDAC1-2E6F-4f9a-9A3F-68A3B97CE33E}"));

CakTrayIcon::CakTrayIcon()
    : m_Icon(0)
    , m_Id(0)
    , m_Master(0)
{
	IsVisible= false;
}

CakTrayIcon::CakTrayIcon(CWnd *pWnd, UINT Id, HICON Icon, LPCTSTR Tip)
    : m_Icon(0)
    , m_Id(0)
    , m_Master(0)
{
	IsVisible= false;
    Create(pWnd, Id);
    SetIconAndTip(Icon, Tip);
}

void CakTrayIcon::Create(CWnd *pWnd, UINT Id)
{
    assert(pWnd);                                       //  Should be valid
    assert(Id);                                         //  Shouldn't be zero

    m_Master = pWnd->m_hWnd;
    m_Id = Id;
	Show();
}

CakTrayIcon::~CakTrayIcon()
{
	Hide();
}

void CakTrayIcon::Show() {

	if( IsVisible)
		return;

	NOTIFYICONDATA nid;
    InitializeTrayIconStruct(&nid, NIF_MESSAGE);
    nid.uCallbackMessage = WM_TRAYICONNOTIFY;
    Shell_NotifyIcon(NIM_ADD, &nid);
	IsVisible= true;
	if( m_Icon)
		SetIcon(m_Icon);
	if(m_Tip)
		SetTip(m_Tip);
}

void CakTrayIcon::Hide() {

	if(!IsVisible)
		return;

	NOTIFYICONDATA nid;
    InitializeTrayIconStruct(&nid);
    Shell_NotifyIcon(NIM_DELETE, &nid);
	IsVisible= false;
}

void CakTrayIcon::InitializeTrayIconStruct(NOTIFYICONDATA *pStruct, UINT Flags/* = 0*/, 
                                           UINT Size/* = NOTIFYICONDATA_V1_SIZE*/)
{
	if( !m_Master || !m_Id)
		return;

    assert(m_Master);                                   //  Should be initialized first
    assert(m_Id);                                       //  Should be initialized first

    pStruct->cbSize = Size;
    pStruct->hWnd = m_Master;
    pStruct->uID = m_Id;
    pStruct->uFlags = Flags;
}

void CakTrayIcon::SetIcon(HICON Icon)
{
    assert(INVALID_HANDLE_VALUE != Icon);               //  Should be valid
    m_Icon = Icon;

    NOTIFYICONDATA nid;
    InitializeTrayIconStruct(&nid, NIF_ICON);
    nid.hIcon = Icon;
    Shell_NotifyIcon(NIM_MODIFY, &nid);
}

void CakTrayIcon::SetTip(LPCTSTR Tip)
{
    assert(Tip);                                        //  Should be valid
    m_Tip = Tip;

    NOTIFYICONDATA nid;
    InitializeTrayIconStruct(&nid, NIF_TIP);
	_tcsncpy_s(nid.szTip, Tip, sizeof(nid.szTip));
    Shell_NotifyIcon(NIM_MODIFY, &nid);
}

void CakTrayIcon::SetIconAndTip(HICON Icon, LPCTSTR Tip)
{
    assert(Tip);                                        //  Should be valid
    assert(INVALID_HANDLE_VALUE != Icon);               //  Should be valid
    m_Icon = Icon;
    m_Tip = Tip;

    NOTIFYICONDATA nid;
    InitializeTrayIconStruct(&nid, NIF_ICON | NIF_TIP);
    nid.hIcon = Icon;
	_tcsncpy_s(nid.szTip, Tip, sizeof(nid.szTip));
    Shell_NotifyIcon(NIM_MODIFY, &nid);
}

void CakTrayIcon::GetIconAndTip(HICON *pIcon, CString *pTip)
{
    assert(pIcon);                                      //  Shouldn't be null
    assert(pTip);                                       //  Shouldn't be null

    *pIcon = m_Icon;
    *pTip = m_Tip;
}

UINT CakTrayIcon::GetId()
{
    return m_Id;
}

void CakTrayIcon::PopupBalloon(LPCTSTR Info, LPCTSTR InfoTitle, DWORD Flags, 
                               UINT Timeout/* = 10000*/)
{
#if (_WIN32_IE >= 0x0500)
    NOTIFYICONDATA nid;
    //  Instruct the shell to handle latest systray enhancements
    InitializeTrayIconStruct(&nid);
    nid.uVersion = NOTIFYICON_VERSION;
    Shell_NotifyIcon(NIM_SETVERSION, &nid);

    //  Show balloon
    InitializeTrayIconStruct(&nid, NIF_INFO, sizeof(NOTIFYICONDATA));
    nid.dwInfoFlags = Flags;
	_tcsncpy_s(nid.szInfo, Info, sizeof(nid.szInfo));
	_tcsncpy_s(nid.szInfoTitle, InfoTitle, sizeof(nid.szInfoTitle));
    Shell_NotifyIcon(NIM_MODIFY, &nid);

    //  Revert to old behaviour
    InitializeTrayIconStruct(&nid);
    nid.uVersion = 0;
    Shell_NotifyIcon(NIM_SETVERSION, &nid);
#endif
}


//
//  CakTrayIconAnimator
//
UINT CakTrayIconAnimator::AnimateProc(LPVOID lpVoid)
{
    AnimateParams *pAp = reinterpret_cast<AnimateParams*>(lpVoid);

    CEvent EvtStop(TRUE, TRUE, pAp->StopEventName);
    EvtStop.ResetEvent();
    while (WAIT_TIMEOUT == WaitForSingleObject(EvtStop, 0))
    {
        for (UINT i = 0; i < pAp->pLoader->GetIconsCount(); i++)
        {
            pAp->pMaster->SetIcon(pAp->pLoader->GetIcon(i));
            if (WAIT_TIMEOUT != WaitForSingleObject(EvtStop, pAp->FrameDelay))
            {
                break;
            }
        }
    }
    return 0;
}

CakTrayIconAnimator::CakTrayIconAnimator(CakTrayIcon *pMaster, UINT ResourceId, 
                                         LPCTSTR Tip/* = 0*/)
    : m_pMaster(pMaster)
    , m_Loader(ResourceId)
{
    assert(pMaster);                                    //  Shouldn't be null

    Initialize(Tip);
}

CakTrayIconAnimator::CakTrayIconAnimator(CakTrayIcon *pMaster, LPCTSTR ResourceId, 
                                         LPCTSTR Tip/* = 0*/)
    : m_pMaster(pMaster)
    , m_Loader(ResourceId)
{
    assert(pMaster);                                    //  Shouldn't be null

    Initialize(Tip);
}

CakTrayIconAnimator::~CakTrayIconAnimator()
{
    if (INVALID_HANDLE_VALUE != m_hAnimateThread)
    {
        CEvent EvtStop(TRUE, TRUE, m_StopEventName);
        EvtStop.SetEvent();
        WaitForSingleObject(m_hAnimateThread, m_FrameDelay * 2);
        CloseHandle(m_hAnimateThread);
    }

    m_pMaster->SetIconAndTip(m_OldIcon, m_OldTip);
}

void CakTrayIconAnimator::Initialize(LPCTSTR Tip)
{
    if (Tip)
        m_Tip = Tip;
    m_FrameDelay = 0;
    m_pMaster->GetIconAndTip(&m_OldIcon, &m_OldTip);
    m_pMaster->SetTip(m_Tip);
    m_hAnimateThread = INVALID_HANDLE_VALUE;
    m_StopEventName.Format(_T("stop.animation.%d"), m_pMaster->GetId());
}

void CakTrayIconAnimator::Animate(UINT FrameDelay/* = 300*/)
{
    m_FrameDelay = FrameDelay;

    m_AnimateParams.FrameDelay = FrameDelay;
    m_AnimateParams.pLoader = &m_Loader;
    m_AnimateParams.pMaster = m_pMaster;
    m_AnimateParams.StopEventName = m_StopEventName;

    //  Start animation thread
    CWinThread *pThread = AfxBeginThread(AnimateProc, &m_AnimateParams, THREAD_PRIORITY_NORMAL, 0, CREATE_SUSPENDED);
    DuplicateHandle(GetCurrentProcess(), pThread->m_hThread, GetCurrentProcess(), 
        &m_hAnimateThread, 0, FALSE, DUPLICATE_SAME_ACCESS);
    pThread->ResumeThread();
}


//
//  CakIconLoader
//
CakIconLoader::CakIconLoader(UINT ResourceId, HMODULE hModule/* = 0*/)
{
    if (!hModule)
    {
        m_hModule = reinterpret_cast<HMODULE>(AfxGetInstanceHandle());
    }

    Initialize(FindResource(m_hModule, MAKEINTRESOURCE(ResourceId), RT_GROUP_ICON));
}

CakIconLoader::CakIconLoader(LPCTSTR ResourceId, HMODULE hModule/* = 0*/)
{
    if (!hModule)
    {
        m_hModule = reinterpret_cast<HMODULE>(AfxGetInstanceHandle());
    }

    Initialize(FindResource(m_hModule, ResourceId, RT_GROUP_ICON));
}

CakIconLoader::~CakIconLoader()
{
    DeleteCurrentIcon();
}

void CakIconLoader::Initialize(HRSRC hResource)
{
    m_CurrentIcon = 0;
    if (hResource)
    {
        HGLOBAL hGlob = LoadResource(m_hModule, hResource);
        if (hGlob)
        {
            ICONDIR *pIcon = reinterpret_cast<ICONDIR*>(LockResource(hGlob));
            if (pIcon)
            {
                m_Frames.SetSize(pIcon->idCount);
                for (int i = 0; i < pIcon->idCount; i++)
                {
                    LoadFrame(pIcon->idEntries[i].nID, &m_Frames[i]);
                }
            }
            FreeResource(hGlob);
        }
        FreeResource(hResource);
    }
}

void CakIconLoader::LoadFrame(UINT Id, CArray<BYTE> *pFrame)
{
    HRSRC hRsrc = FindResource(m_hModule, MAKEINTRESOURCE(Id), RT_ICON);
    if (hRsrc)
    {
        HGLOBAL hGlobal = LoadResource(m_hModule, hRsrc);
        if (hGlobal)
        {
            BYTE *pData = reinterpret_cast<BYTE*>(LockResource(hGlobal));
            pFrame->SetSize(SizeofResource(m_hModule, hRsrc));
            memmove(pFrame->GetData(), pData, pFrame->GetSize());
            FreeResource(hGlobal);
        }
        FreeResource(hRsrc);
    }
}

void CakIconLoader::DeleteCurrentIcon()
{
    if (m_CurrentIcon)
    {
        DestroyIcon(m_CurrentIcon);
        m_CurrentIcon = 0;
    }
}

UINT CakIconLoader::GetIconsCount()
{
    return m_Frames.GetCount();
}

HICON CakIconLoader::GetIcon(UINT Index)
{
    DeleteCurrentIcon();
    if (Index < GetIconsCount())
    {
        HICON hIcon = CreateIconFromResourceEx(m_Frames[Index].GetData(), m_Frames[Index].GetSize(), 
            TRUE, 0x00030000, 16, 16, 0);
        m_CurrentIcon = hIcon;
    }

    return m_CurrentIcon;
}

#define WIDTHBYTES(bits)      ((((bits) + 31) >> 5) << 2)

WORD CakIconLoader::DIBNumColors(LPSTR lpbi)
{
    WORD wBitCount;
    DWORD dwClrUsed;

    dwClrUsed = ((LPBITMAPINFOHEADER) lpbi)->biClrUsed;

    if (dwClrUsed)
        return (WORD) dwClrUsed;

    wBitCount = ((LPBITMAPINFOHEADER) lpbi)->biBitCount;

    switch (wBitCount)
    {
        case 1:
            return 2;
        case 4:
            return 16;
        case 8:
            return 256;
        default:
            return 0;
    }
    return 0;
}

WORD CakIconLoader::PaletteSize(LPSTR lpbi)
{
    return (DIBNumColors(lpbi) * sizeof(RGBQUAD));
}

LPSTR CakIconLoader::FindDIBBits(LPSTR lpbi)
{
   return (lpbi + *(LPDWORD)lpbi + PaletteSize(lpbi));
}

DWORD CakIconLoader::BytesPerLine(BITMAPINFOHEADER *lpBMIH)
{
    return WIDTHBYTES(lpBMIH->biWidth * lpBMIH->biPlanes * lpBMIH->biBitCount);
}

bool CakIconLoader::AdjustIconImagePointers(ICONIMAGE *lpImage)
{
    //  Sanity check
    if (!lpImage)
        return false;

    //  BITMAPINFO is at beginning of bits
    lpImage->lpbi = (LPBITMAPINFO)lpImage->lpBits;
    //  Width - simple enough
    lpImage->Width = lpImage->lpbi->bmiHeader.biWidth;
    //  Icons are stored in funky format where height is doubled - account for it
    lpImage->Height = (lpImage->lpbi->bmiHeader.biHeight) / 2;
    //  How many colors?
    lpImage->Colors = lpImage->lpbi->bmiHeader.biPlanes * lpImage->lpbi->bmiHeader.biBitCount;
    //  XOR bits follow the header and color table
    lpImage->lpXOR = reinterpret_cast<BYTE*>(FindDIBBits((LPSTR)lpImage->lpbi));
    //  AND bits follow the XOR bits
    lpImage->lpAND = lpImage->lpXOR + (lpImage->Height * BytesPerLine((LPBITMAPINFOHEADER)(lpImage->lpbi)));

    return true;
}
