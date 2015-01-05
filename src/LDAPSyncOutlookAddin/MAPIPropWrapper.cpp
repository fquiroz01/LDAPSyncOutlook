#include "stdafx.h"
#include "MAPIPropWrapper.h"


#define USES_IID_IMAPIProp
#define	INITGUID
#include <initguid.h>
#include <mapiguid.h>

#define SREAM_BUF_SIZE 0xffff

CMAPIPropWrapper::CMAPIPropWrapper(void)
{
}

CMAPIPropWrapper::~CMAPIPropWrapper(void)
{
}

STDMETHODIMP CMAPIPropWrapper::Initialize()
{
	#define MAPI_NO_COINIT 0x08
	MAPIINIT_0 MAPIINIT = { 0, MAPI_MULTITHREAD_NOTIFICATIONS | MAPI_NO_COINIT };
	HRESULT hr = MAPIInitialize(&MAPIINIT);
	if(hr == MAPI_E_UNKNOWN_FLAGS)
	{
		MAPIINIT.ulFlags &= ~MAPI_NO_COINIT;
		hr = MAPIInitialize(&MAPIINIT);		
	}	
	if (FAILED(hr)) hr = MAPIInitialize (NULL);
	return hr;
}

STDMETHODIMP CMAPIPropWrapper::Uninitialize()
{
	MAPIUninitialize();
	return S_OK;
}

STDMETHODIMP CMAPIPropWrapper::ReadStreamProp(/*[in]*/ VARIANT pObject, /*[in]*/ long propTag, /*[out,retval]*/ BSTR* result)
{
	LPUNKNOWN unk;
	if (pObject.vt == (VT_DISPATCH | VT_BYREF)) unk = getMapiobjectProperty(*(pObject.ppdispVal));
	else if (pObject.vt == (VT_DISPATCH)) unk = getMapiobjectProperty(pObject.pdispVal);
	else return E_INVALIDARG;	
	if (!unk) return E_NOINTERFACE;

	HRESULT hr = getStreamProp(unk,propTag,result,false);
	unk->Release();
	return hr;
}

STDMETHODIMP CMAPIPropWrapper::ReadRTFStreamProp(/*[in]*/ VARIANT pObject, /*[in]*/ long propTag, /*[out,retval]*/ BSTR* result)
{
	LPUNKNOWN unk;
	if (pObject.vt == (VT_DISPATCH | VT_BYREF)) unk = getMapiobjectProperty(*(pObject.ppdispVal));
	else if (pObject.vt == (VT_DISPATCH)) unk = getMapiobjectProperty(pObject.pdispVal);
	else return E_INVALIDARG;	
	if (!unk) return E_NOINTERFACE;

	HRESULT hr = getStreamProp(unk,propTag,result,true);
	unk->Release();
	return hr;
}

HRESULT CMAPIPropWrapper::getStreamProp(LPUNKNOWN unk, ULONG ulPropTag, BSTR* result, BOOL bCompressed)
{
	CComQIPtr <IMAPIProp,&IID_IMAPIProp> mprop;
	mprop = unk;
	if (mprop)
	{
		HRESULT hr = E_FAIL;
		LPSTR buf = NULL;

		if (bCompressed) hr = readRTFStreamProperty(mprop,ulPropTag,&buf);
		else hr = readStreamProperty(mprop,ulPropTag,&buf);

		if (SUCCEEDED(hr) && buf)
		{
			CComBSTR res(buf);
			res.CopyTo(result);
			hr = S_OK;
		}

		if (buf) delete buf;
		return hr;
	}
	return E_NOINTERFACE;
}

HRESULT CMAPIPropWrapper::readStreamProperty(LPMAPIPROP lpMapiProp, ULONG ulPropTag, LPSTR* lpptDestination)
{
	HRESULT hr;

	if (!lpMapiProp || !lpptDestination) return E_FAIL;
	*lpptDestination = NULL;

	CComPtr<IStream> spStream;
	hr = lpMapiProp->OpenProperty(ulPropTag, &IID_IStream, 0, 0L, (LPUNKNOWN*)&spStream);
	if(FAILED(hr) || !spStream) return E_FAIL;

	return readStream(spStream, lpptDestination, false);
}

HRESULT CMAPIPropWrapper::readRTFStreamProperty(LPMAPIPROP lpMapiProp, ULONG ulPropTag, LPSTR* lpptDestination)
{
	HRESULT hr;

	if (!lpMapiProp || !lpptDestination) return E_FAIL;
	*lpptDestination = NULL;

	CComPtr<IStream> spCompressedStream;
	hr = lpMapiProp->OpenProperty(PR_RTF_COMPRESSED, &IID_IStream, 0, 0L, (LPUNKNOWN *)&spCompressedStream);
	if(FAILED(hr) || !spCompressedStream) return E_FAIL;

	CComPtr<IStream> spStream;
	hr = WrapCompressedRTFStream(spCompressedStream, 0L, &spStream);
	if(FAILED(hr) || !spStream) return E_FAIL;

	return readStream(spStream, lpptDestination, false);
}


HRESULT CMAPIPropWrapper::readStream(CComPtr<IStream> &spStream, LPSTR* lpptDestination, BOOL bHeaderOnly)
{
	HRESULT hr;
	char* buf1=NULL, *buf2=NULL;
	unsigned long sz1=0, sz2=0;

	// Read from IStream to Portions
	while(true)
	{
		// Allocate memory for buffer
		if (buf1) delete buf1;
		buf1 = new char[SREAM_BUF_SIZE];
		if (!buf1) return E_FAIL;

		hr = spStream->Read((LPVOID)buf1, SREAM_BUF_SIZE, &sz1);
		if(FAILED(hr) || hr == S_FALSE || sz1 <= 0)
		{
			delete buf1;
			buf1 = NULL;
		}

		if (buf1)
		{
			char *t = new char[sz1+sz2+1];
			if (!t) return E_FAIL;
			if (buf2)
			{
				memcpy(t,buf2,sz2);
				delete buf2;
			}
			memcpy(t+sz2,buf1,sz1);
			buf2 = t;
			sz2 = sz1+sz2;
			if (bHeaderOnly && sz2 > 0) break;
		}
		else
		{
			break;
		}
	}

	if (!buf2)
	{
		if (FAILED(hr)) return hr;
		buf2 = new char[1]; // for S_FALSE
		sz2 = 0;
	}

	if (buf2)
	{
		buf2[sz2] = 0; // null terminated
		*lpptDestination = buf2;
		return S_OK;
	}

	return E_FAIL;
}


/////////////////////////////////////////////////////////////////////////////
// ONE MAPI PROPERY

STDMETHODIMP CMAPIPropWrapper::getOneProp(LPUNKNOWN unk, long propTag, VARIANT *result)
{
	CComQIPtr <IMAPIProp,&IID_IMAPIProp> mprop;
	mprop = unk;
	if (mprop)
	{
		HRESULT hr = S_OK;
		LPSPropValue val;

		if (FAILED(HrGetOneProp(mprop,propTag,&val))) 
			return MAPI_E_NOT_FOUND;

		CComVariant v;
		VariantInit(&v);
		switch(val->ulPropTag & 0xffff)
		{
		case PT_SHORT:   v = val->Value.i; break;
		case PT_LONG:	 v = val->Value.l; break;
		case PT_FLOAT:	 v = val->Value.flt; break;
		case PT_DOUBLE:	 v = val->Value.dbl; break;
		case PT_STRING8: v = val->Value.lpszA; break;
		case PT_UNICODE: v = val->Value.lpszW; break;
		case PT_BOOLEAN: v = val->Value.b; break;
		case PT_CURRENCY:
			{
				v.cyVal.Hi = val->Value.cur.Hi;
				v.cyVal.Lo = val->Value.cur.Lo;
				v.vt = VT_CY;
			}
			break;
		case PT_I8:
			{
				v.cyVal.Hi = val->Value.li.HighPart;
				v.cyVal.Lo = val->Value.li.LowPart;
				v.vt = VT_CY;
			}
			break;
		case PT_CLSID:
			{
				LPOLESTR str;
				if (SUCCEEDED(hr = StringFromCLSID( *(val->Value.lpguid),&str)))
				{
					hr = S_OK;
					v = str;
					CoTaskMemFree(str);
				}
			}
			break;
		case PT_SYSTIME:
			{
				SYSTEMTIME st;
				double dtime;
				if (FileTimeToSystemTime( &(val->Value.ft), &st) && SystemTimeToVariantTime(&st,&dtime))
				{
					v.date = (DATE)dtime;
					v.vt = VT_DATE;
				}
				else
				{
					hr = E_FAIL;
				}
			}
			break;
		case PT_APPTIME:
			{
				v.date = (DATE)val->Value.at;
				v.vt = VT_DATE;
			}		
			break;
		case PT_BINARY:
			{
				SAFEARRAY *pSA = NULL;
				SAFEARRAYBOUND bound[1];
				LPCTSTR lpData = NULL;				

				bound[0].lLbound = 0;
				bound[0].cElements = val->Value.bin.cb;
				pSA = SafeArrayCreate(VT_UI1,1,bound);
				if (!pSA) return E_OUTOFMEMORY;

				SafeArrayAccessData(pSA, (void**)(&lpData));
				memcpy((void*)lpData,val->Value.bin.lpb,val->Value.bin.cb);

				v.vt = VT_ARRAY | VT_UI1;
				v.parray = pSA;

				SafeArrayUnaccessData(pSA);
			}
			break;
		case PT_NULL:
		case PT_OBJECT:	 v = val->Value.x; break;
		case PT_ERROR:
			{
				v.scode = val->Value.err;
				v.vt = VT_ERROR;
			}
			break;
		default:
			v.scode = E_FAIL;
			v.vt = VT_ERROR;
		}
		MAPIFreeBuffer(&val);
		v.Detach(result);

		return hr;
	}
	return E_NOINTERFACE;
}


STDMETHODIMP CMAPIPropWrapper::GetOneProp(VARIANT pObject, long propTag, VARIANT *result)
{
	LPUNKNOWN unk;
	if (pObject.vt == (VT_DISPATCH | VT_BYREF)) unk = getMapiobjectProperty(*(pObject.ppdispVal));
	else if (pObject.vt == (VT_DISPATCH)) unk = getMapiobjectProperty(pObject.pdispVal);
	else return E_INVALIDARG;	
	if (!unk) return E_NOINTERFACE;

	HRESULT hr = getOneProp(unk,propTag,result);
	unk->Release();
	return hr;
}


/////////////////////////////////////////////////////////////////////////////
// GET MAPIOBJECT THROUGH DISPATCH

LPUNKNOWN CMAPIPropWrapper::getMapiobjectProperty(LPDISPATCH disp)
{
	DISPID		rgDispId;
	DISPPARAMS	dispparams = {NULL, NULL, 0, 0};
	VARIANT		vaResult;	
	LPOLESTR	pName = L"MAPIOBJECT";
	VariantInit(&vaResult);

	if (SUCCEEDED(disp->GetIDsOfNames(IID_NULL,&pName,1,LOCALE_SYSTEM_DEFAULT,&rgDispId)))
		if (SUCCEEDED(disp->Invoke(rgDispId, IID_NULL, 0, DISPATCH_PROPERTYGET, &dispparams, &vaResult, NULL, NULL)))
			if (vaResult.vt == VT_UNKNOWN)
			{
				return vaResult.punkVal;
			}

			return NULL;
}

