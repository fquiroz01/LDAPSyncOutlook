#pragma once

#include <afxtempl.h>


//
//  CakTrayIcon
//
class CakTrayIcon
{

public:
    static const UINT WM_TRAYICONNOTIFY;

public:
    CakTrayIcon();
    CakTrayIcon(CWnd *pWnd, UINT Id, HICON Icon, LPCTSTR Tip);
    ~CakTrayIcon();

    void Create(CWnd *pWnd, UINT Id);

	void Show();
	void Hide();
    void SetIcon(HICON Icon);
    void SetTip(LPCTSTR Tip);
    void SetIconAndTip(HICON Icon, LPCTSTR Tip);

    void GetIconAndTip(HICON *pIcon, CString *pTip);
    UINT GetId();

    void PopupBalloon(LPCTSTR Info, LPCTSTR InfoTitle, DWORD Flags, UINT Timeout = 10000);

private:
    HWND m_Master;
    UINT m_Id;
    HICON m_Icon;
    CString m_Tip;
	bool IsVisible;

    void InitializeTrayIconStruct(NOTIFYICONDATA *pStruct, UINT Flags = 0, 
        UINT Size = NOTIFYICONDATA_V1_SIZE);

};

//
//  CakIconLoader
//
class CakIconLoader
{

public:
    CakIconLoader(UINT ResourceId, HMODULE hModule = 0);
    CakIconLoader(LPCTSTR ResourceId, HMODULE hModule = 0);
    ~CakIconLoader();

    UINT GetIconsCount();
    HICON GetIcon(UINT Index);

private:
    CArray<CArray<BYTE> > m_Frames;
    HMODULE m_hModule;

    void Initialize(HRSRC hResource);
    void LoadFrame(UINT Id, CArray<BYTE> *pFrame);

private:
    HICON m_CurrentIcon;
    void DeleteCurrentIcon();

private:
#pragma pack(push)
#pragma pack(2)
    struct ICONDIRENTRY
    {
        BYTE bWidth;                                    // Width, in pixels, of the image
        BYTE bHeight;                                   // Height, in pixels, of the image
        BYTE bColorCount;                               // Number of colors in image (0 if >=8bpp)
        BYTE bReserved;                                 // Reserved ( must be 0)
        WORD wPlanes;                                   // Color Planes
        WORD wBitCount;                                 // Bits per pixel
        DWORD dwBytesInRes;                             // How many bytes in this resource?
        WORD nID;                                       // Where in the file is this image?
    };

    struct ICONDIR
    {
        WORD idReserved;                                // Reserved (must be 0)
        WORD idType;                                    // Resource Type (1 for icons)
        WORD idCount;                                   // How many images?
        ICONDIRENTRY idEntries[1];                      // An entry for each image (idCount of 'em)
    };
#pragma pack(pop)

    struct ICONIMAGE
    {
        UINT Width, Height, Colors;                     // Width, Height and bpp
        LPBYTE lpBits;                                  // ptr to DIB bits
        DWORD dwNumBytes;                               // how many bytes?
        LPBITMAPINFO lpbi;                              // ptr to header
        LPBYTE lpXOR;                                   // ptr to XOR image bits
        LPBYTE lpAND;                                   // ptr to AND image bits
    };

    LPSTR FindDIBBits(LPSTR lpbi);
    DWORD BytesPerLine(BITMAPINFOHEADER *lpBMIH);
    WORD DIBNumColors(LPSTR lpbi);
    WORD PaletteSize(LPSTR lpbi);
    bool AdjustIconImagePointers(ICONIMAGE *lpImage);

};


//
//  CakTrayIconAnimator
//
class CakTrayIconAnimator
{

public:
    CakTrayIconAnimator(CakTrayIcon *pMaster, UINT ResourceId, LPCTSTR Tip = 0);
    CakTrayIconAnimator(CakTrayIcon *pMaster, LPCTSTR ResourceId, LPCTSTR Tip = 0);
    ~CakTrayIconAnimator();

    void Animate(UINT FrameDelay = 300);

private:
    CakTrayIcon *m_pMaster;
    HICON m_OldIcon;
    CString m_OldTip;

    CString m_Tip;

    CakIconLoader m_Loader;

private:
    void Initialize(LPCTSTR Tip);

private:
    static UINT AnimateProc(LPVOID lpVoid);

    struct AnimateParams
    {
        CakTrayIcon *pMaster;
        CakIconLoader *pLoader;
        UINT FrameDelay;
        CString StopEventName;
    };

    AnimateParams m_AnimateParams;

    HANDLE m_hAnimateThread;
    CString m_StopEventName;
    UINT m_FrameDelay;

};
