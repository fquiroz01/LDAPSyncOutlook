#pragma once
#include <MAPI.H>
#include <MAPIDEFS.H>
// #include <MAPIVAL.H>
#include <MAPIUTIL.H>

class CMAPIPropWrapper
{
public:
	CMAPIPropWrapper(void);
	~CMAPIPropWrapper(void);
public:

	STDMETHODIMP Initialize();
	STDMETHODIMP Uninitialize();

	STDMETHODIMP GetOneProp(/*[in]*/ VARIANT pObject, /*[in]*/ long propTag, /*[out,retval]*/ VARIANT* result);
	STDMETHODIMP ReadStreamProp(/*[in]*/ VARIANT pObject, /*[in]*/ long propTag, /*[out,retval]*/ BSTR* result);
	STDMETHODIMP ReadRTFStreamProp(/*[in]*/ VARIANT pObject, /*[in]*/ long propTag, /*[out,retval]*/ BSTR* result);

private:

	STDMETHODIMP getOneProp(LPUNKNOWN pObject, long propTag, VARIANT* result);

	HRESULT	getStreamProp(LPUNKNOWN unk, ULONG ulPropTag, BSTR* result, BOOL bCompressed);
	HRESULT readRTFStreamProperty(LPMAPIPROP lpMapiProp, ULONG ulPropTag, LPSTR* lpptDestination);
	HRESULT readStreamProperty(LPMAPIPROP lpMapiProp, ULONG ulPropTag, LPSTR* lpptDestination);
	HRESULT readStream(CComPtr<IStream> &spStream, LPSTR* lpptDestination, BOOL bHeaderOnly);

	LPUNKNOWN getMapiobjectProperty(LPDISPATCH disp);
};
