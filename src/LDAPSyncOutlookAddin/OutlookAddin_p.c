

/* this ALWAYS GENERATED file contains the proxy stub code */


 /* File created by MIDL compiler version 7.00.0555 */
/* at Sun Jan 04 16:39:40 2015
 */
/* Compiler settings for OutlookAddin.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 7.00.0555 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#if !defined(_M_IA64) && !defined(_M_AMD64)


#pragma warning( disable: 4049 )  /* more than 64k source lines */
#if _MSC_VER >= 1200
#pragma warning(push)
#endif

#pragma warning( disable: 4211 )  /* redefine extern to static */
#pragma warning( disable: 4232 )  /* dllimport identity*/
#pragma warning( disable: 4024 )  /* array to pointer mapping*/
#pragma warning( disable: 4152 )  /* function/data pointer conversion in expression */
#pragma warning( disable: 4100 ) /* unreferenced arguments in x86 call */

#pragma optimize("", off ) 

#define USE_STUBLESS_PROXY


/* verify that the <rpcproxy.h> version is high enough to compile this file*/
#ifndef __REDQ_RPCPROXY_H_VERSION__
#define __REQUIRED_RPCPROXY_H_VERSION__ 475
#endif


#include "rpcproxy.h"
#ifndef __RPCPROXY_H_VERSION__
#error this stub requires an updated version of <rpcproxy.h>
#endif /* __RPCPROXY_H_VERSION__ */


#include "OutlookAddin.h"

#define TYPE_FORMAT_STRING_SIZE   3                                 
#define PROC_FORMAT_STRING_SIZE   1                                 
#define EXPR_FORMAT_STRING_SIZE   1                                 
#define TRANSMIT_AS_TABLE_SIZE    0            
#define WIRE_MARSHAL_TABLE_SIZE   0            

typedef struct _OutlookAddin_MIDL_TYPE_FORMAT_STRING
    {
    short          Pad;
    unsigned char  Format[ TYPE_FORMAT_STRING_SIZE ];
    } OutlookAddin_MIDL_TYPE_FORMAT_STRING;

typedef struct _OutlookAddin_MIDL_PROC_FORMAT_STRING
    {
    short          Pad;
    unsigned char  Format[ PROC_FORMAT_STRING_SIZE ];
    } OutlookAddin_MIDL_PROC_FORMAT_STRING;

typedef struct _OutlookAddin_MIDL_EXPR_FORMAT_STRING
    {
    long          Pad;
    unsigned char  Format[ EXPR_FORMAT_STRING_SIZE ];
    } OutlookAddin_MIDL_EXPR_FORMAT_STRING;


static const RPC_SYNTAX_IDENTIFIER  _RpcTransferSyntax = 
{{0x8A885D04,0x1CEB,0x11C9,{0x9F,0xE8,0x08,0x00,0x2B,0x10,0x48,0x60}},{2,0}};


extern const OutlookAddin_MIDL_TYPE_FORMAT_STRING OutlookAddin__MIDL_TypeFormatString;
extern const OutlookAddin_MIDL_PROC_FORMAT_STRING OutlookAddin__MIDL_ProcFormatString;
extern const OutlookAddin_MIDL_EXPR_FORMAT_STRING OutlookAddin__MIDL_ExprFormatString;


extern const MIDL_STUB_DESC Object_StubDesc;


extern const MIDL_SERVER_INFO IAddin_ServerInfo;
extern const MIDL_STUBLESS_PROXY_INFO IAddin_ProxyInfo;


extern const MIDL_STUB_DESC Object_StubDesc;


extern const MIDL_SERVER_INFO IPropPage_ServerInfo;
extern const MIDL_STUBLESS_PROXY_INFO IPropPage_ProxyInfo;


extern const MIDL_STUB_DESC Object_StubDesc;


extern const MIDL_SERVER_INFO ILitePage_ServerInfo;
extern const MIDL_STUBLESS_PROXY_INFO ILitePage_ProxyInfo;



#if !defined(__RPC_WIN32__)
#error  Invalid build platform for this stub.
#endif

#if !(TARGET_IS_NT50_OR_LATER)
#error You need Windows 2000 or later to run this stub because it uses these features:
#error   /robust command line switch.
#error However, your C/C++ compilation flags indicate you intend to run this app on earlier systems.
#error This app will fail with the RPC_X_WRONG_STUB_VERSION error.
#endif


static const OutlookAddin_MIDL_PROC_FORMAT_STRING OutlookAddin__MIDL_ProcFormatString =
    {
        0,
        {

			0x0
        }
    };

static const OutlookAddin_MIDL_TYPE_FORMAT_STRING OutlookAddin__MIDL_TypeFormatString =
    {
        0,
        {
			NdrFcShort( 0x0 ),	/* 0 */

			0x0
        }
    };


/* Object interface: IUnknown, ver. 0.0,
   GUID={0x00000000,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}} */


/* Object interface: IDispatch, ver. 0.0,
   GUID={0x00020400,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}} */


/* Object interface: IAddin, ver. 0.0,
   GUID={0xF8608145,0xBB5A,0x4A5A,{0xAE,0x7E,0x24,0x4C,0x6E,0x02,0x80,0x93}} */

#pragma code_seg(".orpc")
static const unsigned short IAddin_FormatStringOffsetTable[] =
    {
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    0
    };

static const MIDL_STUBLESS_PROXY_INFO IAddin_ProxyInfo =
    {
    &Object_StubDesc,
    OutlookAddin__MIDL_ProcFormatString.Format,
    &IAddin_FormatStringOffsetTable[-3],
    0,
    0,
    0
    };


static const MIDL_SERVER_INFO IAddin_ServerInfo = 
    {
    &Object_StubDesc,
    0,
    OutlookAddin__MIDL_ProcFormatString.Format,
    &IAddin_FormatStringOffsetTable[-3],
    0,
    0,
    0,
    0};
CINTERFACE_PROXY_VTABLE(7) _IAddinProxyVtbl = 
{
    0,
    &IID_IAddin,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    0 /* IDispatch::GetTypeInfoCount */ ,
    0 /* IDispatch::GetTypeInfo */ ,
    0 /* IDispatch::GetIDsOfNames */ ,
    0 /* IDispatch_Invoke_Proxy */
};


static const PRPC_STUB_FUNCTION IAddin_table[] =
{
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION
};

CInterfaceStubVtbl _IAddinStubVtbl =
{
    &IID_IAddin,
    &IAddin_ServerInfo,
    7,
    &IAddin_table[-3],
    CStdStubBuffer_DELEGATING_METHODS
};


/* Object interface: IPropPage, ver. 0.0,
   GUID={0xC76C5E4D,0xFDFF,0x4756,{0xB9,0x25,0xB2,0x98,0x17,0xCE,0x62,0x88}} */

#pragma code_seg(".orpc")
static const unsigned short IPropPage_FormatStringOffsetTable[] =
    {
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    0
    };

static const MIDL_STUBLESS_PROXY_INFO IPropPage_ProxyInfo =
    {
    &Object_StubDesc,
    OutlookAddin__MIDL_ProcFormatString.Format,
    &IPropPage_FormatStringOffsetTable[-3],
    0,
    0,
    0
    };


static const MIDL_SERVER_INFO IPropPage_ServerInfo = 
    {
    &Object_StubDesc,
    0,
    OutlookAddin__MIDL_ProcFormatString.Format,
    &IPropPage_FormatStringOffsetTable[-3],
    0,
    0,
    0,
    0};
CINTERFACE_PROXY_VTABLE(7) _IPropPageProxyVtbl = 
{
    0,
    &IID_IPropPage,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    0 /* IDispatch::GetTypeInfoCount */ ,
    0 /* IDispatch::GetTypeInfo */ ,
    0 /* IDispatch::GetIDsOfNames */ ,
    0 /* IDispatch_Invoke_Proxy */
};


static const PRPC_STUB_FUNCTION IPropPage_table[] =
{
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION
};

CInterfaceStubVtbl _IPropPageStubVtbl =
{
    &IID_IPropPage,
    &IPropPage_ServerInfo,
    7,
    &IPropPage_table[-3],
    CStdStubBuffer_DELEGATING_METHODS
};


/* Object interface: ILitePage, ver. 0.0,
   GUID={0x067CA4E2,0x556E,0x47DD,{0x9E,0xFA,0xAD,0xAA,0xBB,0x2C,0x0D,0x95}} */

#pragma code_seg(".orpc")
static const unsigned short ILitePage_FormatStringOffsetTable[] =
    {
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    0
    };

static const MIDL_STUBLESS_PROXY_INFO ILitePage_ProxyInfo =
    {
    &Object_StubDesc,
    OutlookAddin__MIDL_ProcFormatString.Format,
    &ILitePage_FormatStringOffsetTable[-3],
    0,
    0,
    0
    };


static const MIDL_SERVER_INFO ILitePage_ServerInfo = 
    {
    &Object_StubDesc,
    0,
    OutlookAddin__MIDL_ProcFormatString.Format,
    &ILitePage_FormatStringOffsetTable[-3],
    0,
    0,
    0,
    0};
CINTERFACE_PROXY_VTABLE(7) _ILitePageProxyVtbl = 
{
    0,
    &IID_ILitePage,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    0 /* IDispatch::GetTypeInfoCount */ ,
    0 /* IDispatch::GetTypeInfo */ ,
    0 /* IDispatch::GetIDsOfNames */ ,
    0 /* IDispatch_Invoke_Proxy */
};


static const PRPC_STUB_FUNCTION ILitePage_table[] =
{
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION
};

CInterfaceStubVtbl _ILitePageStubVtbl =
{
    &IID_ILitePage,
    &ILitePage_ServerInfo,
    7,
    &ILitePage_table[-3],
    CStdStubBuffer_DELEGATING_METHODS
};

static const MIDL_STUB_DESC Object_StubDesc = 
    {
    0,
    NdrOleAllocate,
    NdrOleFree,
    0,
    0,
    0,
    0,
    0,
    OutlookAddin__MIDL_TypeFormatString.Format,
    1, /* -error bounds_check flag */
    0x50002, /* Ndr library version */
    0,
    0x700022b, /* MIDL Version 7.0.555 */
    0,
    0,
    0,  /* notify & notify_flag routine table */
    0x1, /* MIDL flag */
    0, /* cs routines */
    0,   /* proxy/server info */
    0
    };

const CInterfaceProxyVtbl * const _OutlookAddin_ProxyVtblList[] = 
{
    ( CInterfaceProxyVtbl *) &_IAddinProxyVtbl,
    ( CInterfaceProxyVtbl *) &_IPropPageProxyVtbl,
    ( CInterfaceProxyVtbl *) &_ILitePageProxyVtbl,
    0
};

const CInterfaceStubVtbl * const _OutlookAddin_StubVtblList[] = 
{
    ( CInterfaceStubVtbl *) &_IAddinStubVtbl,
    ( CInterfaceStubVtbl *) &_IPropPageStubVtbl,
    ( CInterfaceStubVtbl *) &_ILitePageStubVtbl,
    0
};

PCInterfaceName const _OutlookAddin_InterfaceNamesList[] = 
{
    "IAddin",
    "IPropPage",
    "ILitePage",
    0
};

const IID *  const _OutlookAddin_BaseIIDList[] = 
{
    &IID_IDispatch,
    &IID_IDispatch,
    &IID_IDispatch,
    0
};


#define _OutlookAddin_CHECK_IID(n)	IID_GENERIC_CHECK_IID( _OutlookAddin, pIID, n)

int __stdcall _OutlookAddin_IID_Lookup( const IID * pIID, int * pIndex )
{
    IID_BS_LOOKUP_SETUP

    IID_BS_LOOKUP_INITIAL_TEST( _OutlookAddin, 3, 2 )
    IID_BS_LOOKUP_NEXT_TEST( _OutlookAddin, 1 )
    IID_BS_LOOKUP_RETURN_RESULT( _OutlookAddin, 3, *pIndex )
    
}

const ExtendedProxyFileInfo OutlookAddin_ProxyFileInfo = 
{
    (PCInterfaceProxyVtblList *) & _OutlookAddin_ProxyVtblList,
    (PCInterfaceStubVtblList *) & _OutlookAddin_StubVtblList,
    (const PCInterfaceName * ) & _OutlookAddin_InterfaceNamesList,
    (const IID ** ) & _OutlookAddin_BaseIIDList,
    & _OutlookAddin_IID_Lookup, 
    3,
    2,
    0, /* table of [async_uuid] interfaces */
    0, /* Filler1 */
    0, /* Filler2 */
    0  /* Filler3 */
};
#pragma optimize("", on )
#if _MSC_VER >= 1200
#pragma warning(pop)
#endif


#endif /* !defined(_M_IA64) && !defined(_M_AMD64)*/

