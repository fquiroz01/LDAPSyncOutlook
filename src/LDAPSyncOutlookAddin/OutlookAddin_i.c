

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


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

#pragma warning( disable: 4049 )  /* more than 64k source lines */


#ifdef __cplusplus
extern "C"{
#endif 


#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <guiddef.h>
#undef INITGUID
#else
#include <guiddef.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, IID_IAddin,0xF8608145,0xBB5A,0x4A5A,0xAE,0x7E,0x24,0x4C,0x6E,0x02,0x80,0x93);


MIDL_DEFINE_GUID(IID, IID_IPropPage,0xC76C5E4D,0xFDFF,0x4756,0xB9,0x25,0xB2,0x98,0x17,0xCE,0x62,0x88);


MIDL_DEFINE_GUID(IID, IID_ILitePage,0x067CA4E2,0x556E,0x47DD,0x9E,0xFA,0xAD,0xAA,0xBB,0x2C,0x0D,0x95);


MIDL_DEFINE_GUID(IID, LIBID_OUTLOOKADDINLib,0xEDDBDEA4,0x5C07,0x453F,0xBE,0x8C,0x81,0xD7,0x38,0x98,0x43,0x81);


MIDL_DEFINE_GUID(CLSID, CLSID_Addin,0xC704648D,0x6030,0x47E9,0xAD,0xBA,0x1E,0x13,0xB6,0xA7,0x84,0xAE);


MIDL_DEFINE_GUID(CLSID, CLSID_PropPage,0xBE1A7148,0x9F02,0x40F9,0xAB,0x5A,0x5A,0xD4,0xE7,0xD1,0x1E,0x62);


MIDL_DEFINE_GUID(CLSID, CLSID_LitePage,0x21881BCB,0x9F1E,0x42CE,0xBD,0x5B,0xED,0x0E,0x13,0x37,0x9A,0x54);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



