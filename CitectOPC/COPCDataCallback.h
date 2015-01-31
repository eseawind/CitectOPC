#include <atlbase.h>
#include <atlcom.h>
#include "opcda.h"
//
// Callback.h : Declaration of callback class and definition of minor methods
//
//---------------------------------------------------------
// (c) COPYRIGHT 2003,2004 http://www.opc-china.com INC.
// ALL RIGHTS RESERVED
// Original Author:geekCarnegie	
// Original Author Email:geekcarnegie@gmail.com
//---------------------------------------------------------
class COPCDataCallback : public IOPCDataCallback,
	   public CComObjectRootEx<CComSingleThreadModel>
{

public:
	COPCDataCallback() {};
	virtual ~COPCDataCallback() { ; };

	BEGIN_COM_MAP(COPCDataCallback)
		COM_INTERFACE_ENTRY(IOPCDataCallback)
	END_COM_MAP()

	// IOPCDataCallback
	STDMETHODIMP  OnDataChange(
		/* [in] */ DWORD dwTransid,
		/* [in] */ OPCHANDLE hGroup,
		/* [in] */ HRESULT hrMasterquality,
		/* [in] */ HRESULT hrMastererror,
		/* [in] */ DWORD dwCount,
		/* [size_is][in] */ OPCHANDLE __RPC_FAR *phClientItems,
		/* [size_is][in] */ VARIANT __RPC_FAR *pvValues,
		/* [size_is][in] */ WORD __RPC_FAR *pwQualities,
		/* [size_is][in] */ FILETIME __RPC_FAR *pftTimeStamps,
		/* [size_is][in] */ HRESULT __RPC_FAR *pErrors);
	STDMETHODIMP  OnReadComplete(
		/* [in] */ DWORD dwTransid,
		/* [in] */ OPCHANDLE hGroup,
		/* [in] */ HRESULT hrMasterquality,
		/* [in] */ HRESULT hrMastererror,
		/* [in] */ DWORD dwCount,
		/* [size_is][in] */ OPCHANDLE __RPC_FAR *phClientItems,
		/* [size_is][in] */ VARIANT __RPC_FAR *pvValues,
		/* [size_is][in] */ WORD __RPC_FAR *pwQualities,
		/* [size_is][in] */ FILETIME __RPC_FAR *pftTimeStamps,
		/* [size_is][in] */ HRESULT __RPC_FAR *pErrors);
	STDMETHODIMP  OnWriteComplete(
		/* [in] */ DWORD dwTransid,
		/* [in] */ OPCHANDLE hGroup,
		/* [in] */ HRESULT hrMastererr,
		/* [in] */ DWORD dwCount,
		/* [size_is][in] */ OPCHANDLE __RPC_FAR *pClienthandles,
		/* [size_is][in] */ HRESULT __RPC_FAR *pErrors);
	STDMETHODIMP  OnCancelComplete(
		/* [in] */ DWORD dwTransid,
		/* [in] */ OPCHANDLE hGroup)
	{
		return S_OK;
	};
};