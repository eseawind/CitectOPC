/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Sun Nov 20 22:30:48 2005
 */
/* Compiler settings for opcenum.idl:
    Os (OptLev=s), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __opcenum_h__
#define __opcenum_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IOPCServerList2_FWD_DEFINED__
#define __IOPCServerList2_FWD_DEFINED__
typedef interface IOPCServerList2 IOPCServerList2;
#endif 	/* __IOPCServerList2_FWD_DEFINED__ */


#ifndef __IOPCEnumGUID_FWD_DEFINED__
#define __IOPCEnumGUID_FWD_DEFINED__
typedef interface IOPCEnumGUID IOPCEnumGUID;
#endif 	/* __IOPCEnumGUID_FWD_DEFINED__ */


#ifndef __IOPCServerList_FWD_DEFINED__
#define __IOPCServerList_FWD_DEFINED__
typedef interface IOPCServerList IOPCServerList;
#endif 	/* __IOPCServerList_FWD_DEFINED__ */


#ifndef __IEnumGUID_FWD_DEFINED__
#define __IEnumGUID_FWD_DEFINED__
typedef interface IEnumGUID IEnumGUID;
#endif 	/* __IEnumGUID_FWD_DEFINED__ */


#ifndef __OpcServerList_FWD_DEFINED__
#define __OpcServerList_FWD_DEFINED__

#ifdef __cplusplus
typedef class OpcServerList OpcServerList;
#else
typedef struct OpcServerList OpcServerList;
#endif /* __cplusplus */

#endif 	/* __OpcServerList_FWD_DEFINED__ */


void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 


#ifndef __OpcEnumLib_LIBRARY_DEFINED__
#define __OpcEnumLib_LIBRARY_DEFINED__

/* library OpcEnumLib */
/* [helpstring][version][uuid] */ 






EXTERN_C const IID LIBID_OpcEnumLib;

#ifndef __IOPCServerList2_INTERFACE_DEFINED__
#define __IOPCServerList2_INTERFACE_DEFINED__

/* interface IOPCServerList2 */
/* [object][uuid] */ 


EXTERN_C const IID IID_IOPCServerList2;


#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("9DD0B56C-AD9E-43EE-8305-487F3188BF7A")
    IOPCServerList2 : public IUnknown
    {
    public:
        virtual HRESULT __stdcall EnumClassesOfCategories( 
            /* [in] */ unsigned long cImplemented,
            /* [in] */ GUID __RPC_FAR *rgcatidImpl,
            /* [in] */ unsigned long cRequired,
            /* [in] */ GUID __RPC_FAR *rgcatidReq,
            /* [out] */ IOPCEnumGUID __RPC_FAR *__RPC_FAR *ppenumClsid) = 0;
        
        virtual HRESULT __stdcall GetClassDetails( 
            /* [in] */ GUID __RPC_FAR *clsid,
            /* [out] */ LPWSTR __RPC_FAR *ppszProgID,
            /* [out] */ LPWSTR __RPC_FAR *ppszUserType,
            /* [out] */ LPWSTR __RPC_FAR *ppszVerIndProgID) = 0;
        
        virtual HRESULT __stdcall CLSIDFromProgID( 
            /* [in] */ LPWSTR szProgId,
            /* [out] */ GUID __RPC_FAR *clsid) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IOPCServerList2Vtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IOPCServerList2 __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IOPCServerList2 __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IOPCServerList2 __RPC_FAR * This);
        
        HRESULT ( __stdcall __RPC_FAR *EnumClassesOfCategories )( 
            IOPCServerList2 __RPC_FAR * This,
            /* [in] */ unsigned long cImplemented,
            /* [in] */ GUID __RPC_FAR *rgcatidImpl,
            /* [in] */ unsigned long cRequired,
            /* [in] */ GUID __RPC_FAR *rgcatidReq,
            /* [out] */ IOPCEnumGUID __RPC_FAR *__RPC_FAR *ppenumClsid);
        
        HRESULT ( __stdcall __RPC_FAR *GetClassDetails )( 
            IOPCServerList2 __RPC_FAR * This,
            /* [in] */ GUID __RPC_FAR *clsid,
            /* [out] */ LPWSTR __RPC_FAR *ppszProgID,
            /* [out] */ LPWSTR __RPC_FAR *ppszUserType,
            /* [out] */ LPWSTR __RPC_FAR *ppszVerIndProgID);
        
        HRESULT ( __stdcall __RPC_FAR *CLSIDFromProgID )( 
            IOPCServerList2 __RPC_FAR * This,
            /* [in] */ LPWSTR szProgId,
            /* [out] */ GUID __RPC_FAR *clsid);
        
        END_INTERFACE
    } IOPCServerList2Vtbl;

    interface IOPCServerList2
    {
        CONST_VTBL struct IOPCServerList2Vtbl __RPC_FAR *lpVtbl;
    };
//#endif /* interface IOPCServerList2 */ÊýÓà¶àÓà

    

#ifdef COBJMACROS


#define IOPCServerList2_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IOPCServerList2_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IOPCServerList2_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IOPCServerList2_EnumClassesOfCategories(This,cImplemented,rgcatidImpl,cRequired,rgcatidReq,ppenumClsid)	\
    (This)->lpVtbl -> EnumClassesOfCategories(This,cImplemented,rgcatidImpl,cRequired,rgcatidReq,ppenumClsid)

#define IOPCServerList2_GetClassDetails(This,clsid,ppszProgID,ppszUserType,ppszVerIndProgID)	\
    (This)->lpVtbl -> GetClassDetails(This,clsid,ppszProgID,ppszUserType,ppszVerIndProgID)

#define IOPCServerList2_CLSIDFromProgID(This,szProgId,clsid)	\
    (This)->lpVtbl -> CLSIDFromProgID(This,szProgId,clsid)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT __stdcall IOPCServerList2_EnumClassesOfCategories_Proxy( 
    IOPCServerList2 __RPC_FAR * This,
    /* [in] */ unsigned long cImplemented,
    /* [in] */ GUID __RPC_FAR *rgcatidImpl,
    /* [in] */ unsigned long cRequired,
    /* [in] */ GUID __RPC_FAR *rgcatidReq,
    /* [out] */ IOPCEnumGUID __RPC_FAR *__RPC_FAR *ppenumClsid);


void __RPC_STUB IOPCServerList2_EnumClassesOfCategories_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT __stdcall IOPCServerList2_GetClassDetails_Proxy( 
    IOPCServerList2 __RPC_FAR * This,
    /* [in] */ GUID __RPC_FAR *clsid,
    /* [out] */ LPWSTR __RPC_FAR *ppszProgID,
    /* [out] */ LPWSTR __RPC_FAR *ppszUserType,
    /* [out] */ LPWSTR __RPC_FAR *ppszVerIndProgID);


void __RPC_STUB IOPCServerList2_GetClassDetails_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT __stdcall IOPCServerList2_CLSIDFromProgID_Proxy( 
    IOPCServerList2 __RPC_FAR * This,
    /* [in] */ LPWSTR szProgId,
    /* [out] */ GUID __RPC_FAR *clsid);


void __RPC_STUB IOPCServerList2_CLSIDFromProgID_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IOPCServerList2_INTERFACE_DEFINED__ */


#ifndef __IOPCEnumGUID_INTERFACE_DEFINED__
#define __IOPCEnumGUID_INTERFACE_DEFINED__

/* interface IOPCEnumGUID */
/* [object][uuid] */ 


EXTERN_C const IID IID_IOPCEnumGUID;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("55C382C8-21C7-4E88-96C1-BECFB1E3F483")
    IOPCEnumGUID : public IUnknown
    {
    public:
        virtual HRESULT __stdcall Next( 
            /* [in] */ unsigned long celt,
            /* [out] */ GUID __RPC_FAR *rgelt,
            /* [out] */ unsigned long __RPC_FAR *pceltFetched) = 0;
        
        virtual HRESULT __stdcall Skip( 
            /* [in] */ unsigned long celt) = 0;
        
        virtual HRESULT __stdcall Reset( void) = 0;
        
        virtual HRESULT __stdcall Clone( 
            /* [out] */ IOPCEnumGUID __RPC_FAR *__RPC_FAR *ppenum) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IOPCEnumGUIDVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IOPCEnumGUID __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IOPCEnumGUID __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IOPCEnumGUID __RPC_FAR * This);
        
        HRESULT ( __stdcall __RPC_FAR *Next )( 
            IOPCEnumGUID __RPC_FAR * This,
            /* [in] */ unsigned long celt,
            /* [out] */ GUID __RPC_FAR *rgelt,
            /* [out] */ unsigned long __RPC_FAR *pceltFetched);
        
        HRESULT ( __stdcall __RPC_FAR *Skip )( 
            IOPCEnumGUID __RPC_FAR * This,
            /* [in] */ unsigned long celt);
        
        HRESULT ( __stdcall __RPC_FAR *Reset )( 
            IOPCEnumGUID __RPC_FAR * This);
        
        HRESULT ( __stdcall __RPC_FAR *Clone )( 
            IOPCEnumGUID __RPC_FAR * This,
            /* [out] */ IOPCEnumGUID __RPC_FAR *__RPC_FAR *ppenum);
        
        END_INTERFACE
    } IOPCEnumGUIDVtbl;

    interface IOPCEnumGUID
    {
        CONST_VTBL struct IOPCEnumGUIDVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IOPCEnumGUID_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IOPCEnumGUID_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IOPCEnumGUID_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IOPCEnumGUID_Next(This,celt,rgelt,pceltFetched)	\
    (This)->lpVtbl -> Next(This,celt,rgelt,pceltFetched)

#define IOPCEnumGUID_Skip(This,celt)	\
    (This)->lpVtbl -> Skip(This,celt)

#define IOPCEnumGUID_Reset(This)	\
    (This)->lpVtbl -> Reset(This)

#define IOPCEnumGUID_Clone(This,ppenum)	\
    (This)->lpVtbl -> Clone(This,ppenum)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT __stdcall IOPCEnumGUID_Next_Proxy( 
    IOPCEnumGUID __RPC_FAR * This,
    /* [in] */ unsigned long celt,
    /* [out] */ GUID __RPC_FAR *rgelt,
    /* [out] */ unsigned long __RPC_FAR *pceltFetched);


void __RPC_STUB IOPCEnumGUID_Next_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT __stdcall IOPCEnumGUID_Skip_Proxy( 
    IOPCEnumGUID __RPC_FAR * This,
    /* [in] */ unsigned long celt);


void __RPC_STUB IOPCEnumGUID_Skip_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT __stdcall IOPCEnumGUID_Reset_Proxy( 
    IOPCEnumGUID __RPC_FAR * This);


void __RPC_STUB IOPCEnumGUID_Reset_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT __stdcall IOPCEnumGUID_Clone_Proxy( 
    IOPCEnumGUID __RPC_FAR * This,
    /* [out] */ IOPCEnumGUID __RPC_FAR *__RPC_FAR *ppenum);


void __RPC_STUB IOPCEnumGUID_Clone_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IOPCEnumGUID_INTERFACE_DEFINED__ */


#ifndef __IOPCServerList_INTERFACE_DEFINED__
#define __IOPCServerList_INTERFACE_DEFINED__

/* interface IOPCServerList */
/* [object][uuid] */ 


EXTERN_C const IID IID_IOPCServerList;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("13486D50-4821-11D2-A494-3CB306C10000")
    IOPCServerList : public IUnknown
    {
    public:
        virtual HRESULT __stdcall EnumClassesOfCategories( 
            /* [in] */ unsigned long cImplemented,
            /* [in] */ GUID __RPC_FAR *rgcatidImpl,
            /* [in] */ unsigned long cRequired,
            /* [in] */ GUID __RPC_FAR *rgcatidReq,
            /* [out] */ IEnumGUID __RPC_FAR *__RPC_FAR *ppenumClsid) = 0;
        
        virtual HRESULT __stdcall GetClassDetails( 
            /* [in] */ GUID __RPC_FAR *clsid,
            /* [out] */ LPWSTR __RPC_FAR *ppszProgID,
            /* [out] */ LPWSTR __RPC_FAR *ppszUserType) = 0;
        
        virtual HRESULT __stdcall CLSIDFromProgID( 
            /* [in] */ LPWSTR szProgId,
            /* [out] */ GUID __RPC_FAR *clsid) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IOPCServerListVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IOPCServerList __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IOPCServerList __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IOPCServerList __RPC_FAR * This);
        
        HRESULT ( __stdcall __RPC_FAR *EnumClassesOfCategories )( 
            IOPCServerList __RPC_FAR * This,
            /* [in] */ unsigned long cImplemented,
            /* [in] */ GUID __RPC_FAR *rgcatidImpl,
            /* [in] */ unsigned long cRequired,
            /* [in] */ GUID __RPC_FAR *rgcatidReq,
            /* [out] */ IEnumGUID __RPC_FAR *__RPC_FAR *ppenumClsid);
        
        HRESULT ( __stdcall __RPC_FAR *GetClassDetails )( 
            IOPCServerList __RPC_FAR * This,
            /* [in] */ GUID __RPC_FAR *clsid,
            /* [out] */ LPWSTR __RPC_FAR *ppszProgID,
            /* [out] */ LPWSTR __RPC_FAR *ppszUserType);
        
        HRESULT ( __stdcall __RPC_FAR *CLSIDFromProgID )( 
            IOPCServerList __RPC_FAR * This,
            /* [in] */ LPWSTR szProgId,
            /* [out] */ GUID __RPC_FAR *clsid);
        
        END_INTERFACE
    } IOPCServerListVtbl;

    interface IOPCServerList
    {
        CONST_VTBL struct IOPCServerListVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IOPCServerList_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IOPCServerList_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IOPCServerList_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IOPCServerList_EnumClassesOfCategories(This,cImplemented,rgcatidImpl,cRequired,rgcatidReq,ppenumClsid)	\
    (This)->lpVtbl -> EnumClassesOfCategories(This,cImplemented,rgcatidImpl,cRequired,rgcatidReq,ppenumClsid)

#define IOPCServerList_GetClassDetails(This,clsid,ppszProgID,ppszUserType)	\
    (This)->lpVtbl -> GetClassDetails(This,clsid,ppszProgID,ppszUserType)

#define IOPCServerList_CLSIDFromProgID(This,szProgId,clsid)	\
    (This)->lpVtbl -> CLSIDFromProgID(This,szProgId,clsid)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT __stdcall IOPCServerList_EnumClassesOfCategories_Proxy( 
    IOPCServerList __RPC_FAR * This,
    /* [in] */ unsigned long cImplemented,
    /* [in] */ GUID __RPC_FAR *rgcatidImpl,
    /* [in] */ unsigned long cRequired,
    /* [in] */ GUID __RPC_FAR *rgcatidReq,
    /* [out] */ IEnumGUID __RPC_FAR *__RPC_FAR *ppenumClsid);


void __RPC_STUB IOPCServerList_EnumClassesOfCategories_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT __stdcall IOPCServerList_GetClassDetails_Proxy( 
    IOPCServerList __RPC_FAR * This,
    /* [in] */ GUID __RPC_FAR *clsid,
    /* [out] */ LPWSTR __RPC_FAR *ppszProgID,
    /* [out] */ LPWSTR __RPC_FAR *ppszUserType);


void __RPC_STUB IOPCServerList_GetClassDetails_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT __stdcall IOPCServerList_CLSIDFromProgID_Proxy( 
    IOPCServerList __RPC_FAR * This,
    /* [in] */ LPWSTR szProgId,
    /* [out] */ GUID __RPC_FAR *clsid);


void __RPC_STUB IOPCServerList_CLSIDFromProgID_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IOPCServerList_INTERFACE_DEFINED__ */


#ifndef __IEnumGUID_INTERFACE_DEFINED__
#define __IEnumGUID_INTERFACE_DEFINED__

/* interface IEnumGUID */
/* [object][uuid] */ 


EXTERN_C const IID IID_IEnumGUID;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("0002E000-0000-0000-C000-000000000046")
    IEnumGUID : public IUnknown
    {
    public:
        virtual HRESULT __stdcall Next( 
            /* [in] */ unsigned long celt,
            /* [out] */ GUID __RPC_FAR *rgelt,
            /* [out] */ unsigned long __RPC_FAR *pceltFetched) = 0;
        
        virtual HRESULT __stdcall Skip( 
            /* [in] */ unsigned long celt) = 0;
        
        virtual HRESULT __stdcall Reset( void) = 0;
        
        virtual HRESULT __stdcall Clone( 
            /* [out] */ IEnumGUID __RPC_FAR *__RPC_FAR *ppenum) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IEnumGUIDVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IEnumGUID __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IEnumGUID __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IEnumGUID __RPC_FAR * This);
        
        HRESULT ( __stdcall __RPC_FAR *Next )( 
            IEnumGUID __RPC_FAR * This,
            /* [in] */ unsigned long celt,
            /* [out] */ GUID __RPC_FAR *rgelt,
            /* [out] */ unsigned long __RPC_FAR *pceltFetched);
        
        HRESULT ( __stdcall __RPC_FAR *Skip )( 
            IEnumGUID __RPC_FAR * This,
            /* [in] */ unsigned long celt);
        
        HRESULT ( __stdcall __RPC_FAR *Reset )( 
            IEnumGUID __RPC_FAR * This);
        
        HRESULT ( __stdcall __RPC_FAR *Clone )( 
            IEnumGUID __RPC_FAR * This,
            /* [out] */ IEnumGUID __RPC_FAR *__RPC_FAR *ppenum);
        
        END_INTERFACE
    } IEnumGUIDVtbl;

    interface IEnumGUID
    {
        CONST_VTBL struct IEnumGUIDVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IEnumGUID_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IEnumGUID_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IEnumGUID_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IEnumGUID_Next(This,celt,rgelt,pceltFetched)	\
    (This)->lpVtbl -> Next(This,celt,rgelt,pceltFetched)

#define IEnumGUID_Skip(This,celt)	\
    (This)->lpVtbl -> Skip(This,celt)

#define IEnumGUID_Reset(This)	\
    (This)->lpVtbl -> Reset(This)

#define IEnumGUID_Clone(This,ppenum)	\
    (This)->lpVtbl -> Clone(This,ppenum)

#endif /* COBJMACROS */


#endif 	/* C style interface */



HRESULT __stdcall IEnumGUID_Next_Proxy( 
    IEnumGUID __RPC_FAR * This,
    /* [in] */ unsigned long celt,
    /* [out] */ GUID __RPC_FAR *rgelt,
    /* [out] */ unsigned long __RPC_FAR *pceltFetched);


void __RPC_STUB IEnumGUID_Next_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT __stdcall IEnumGUID_Skip_Proxy( 
    IEnumGUID __RPC_FAR * This,
    /* [in] */ unsigned long celt);


void __RPC_STUB IEnumGUID_Skip_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT __stdcall IEnumGUID_Reset_Proxy( 
    IEnumGUID __RPC_FAR * This);


void __RPC_STUB IEnumGUID_Reset_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


HRESULT __stdcall IEnumGUID_Clone_Proxy( 
    IEnumGUID __RPC_FAR * This,
    /* [out] */ IEnumGUID __RPC_FAR *__RPC_FAR *ppenum);


void __RPC_STUB IEnumGUID_Clone_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IEnumGUID_INTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_OpcServerList;

#ifdef __cplusplus

class DECLSPEC_UUID("13486D51-4821-11D2-A494-3CB306C10000")
OpcServerList;
#endif
#endif /* __OpcEnumLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
