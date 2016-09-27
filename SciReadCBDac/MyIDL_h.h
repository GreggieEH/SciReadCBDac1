

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Tue Sep 27 11:15:18 2016
 */
/* Compiler settings for MyIDL.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.00.0603 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__


#ifndef __MyIDL_h_h__
#define __MyIDL_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __ISciReadCBDac_FWD_DEFINED__
#define __ISciReadCBDac_FWD_DEFINED__
typedef interface ISciReadCBDac ISciReadCBDac;

#endif 	/* __ISciReadCBDac_FWD_DEFINED__ */


#ifndef ___SciReadCBDac_FWD_DEFINED__
#define ___SciReadCBDac_FWD_DEFINED__
typedef interface _SciReadCBDac _SciReadCBDac;

#endif 	/* ___SciReadCBDac_FWD_DEFINED__ */


#ifndef __SciReadCBDac_FWD_DEFINED__
#define __SciReadCBDac_FWD_DEFINED__

#ifdef __cplusplus
typedef class SciReadCBDac SciReadCBDac;
#else
typedef struct SciReadCBDac SciReadCBDac;
#endif /* __cplusplus */

#endif 	/* __SciReadCBDac_FWD_DEFINED__ */


#ifndef __IMyContainer_FWD_DEFINED__
#define __IMyContainer_FWD_DEFINED__
typedef interface IMyContainer IMyContainer;

#endif 	/* __IMyContainer_FWD_DEFINED__ */


#ifndef __IMyExtended_FWD_DEFINED__
#define __IMyExtended_FWD_DEFINED__
typedef interface IMyExtended IMyExtended;

#endif 	/* __IMyExtended_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_MyIDL_0000_0000 */
/* [local] */ 

#pragma once


extern RPC_IF_HANDLE __MIDL_itf_MyIDL_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_MyIDL_0000_0000_v0_0_s_ifspec;


#ifndef __SciReadCBDac_LIBRARY_DEFINED__
#define __SciReadCBDac_LIBRARY_DEFINED__

/* library SciReadCBDac */
/* [version][helpstring][uuid] */ 


EXTERN_C const IID LIBID_SciReadCBDac;

#ifndef __ISciReadCBDac_DISPINTERFACE_DEFINED__
#define __ISciReadCBDac_DISPINTERFACE_DEFINED__

/* dispinterface ISciReadCBDac */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID_ISciReadCBDac;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("8244515C-54A3-4A64-A260-8C9BD432DF5D")
    ISciReadCBDac : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct ISciReadCBDacVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ISciReadCBDac * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ISciReadCBDac * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ISciReadCBDac * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ISciReadCBDac * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ISciReadCBDac * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ISciReadCBDac * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ISciReadCBDac * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } ISciReadCBDacVtbl;

    interface ISciReadCBDac
    {
        CONST_VTBL struct ISciReadCBDacVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ISciReadCBDac_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ISciReadCBDac_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ISciReadCBDac_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ISciReadCBDac_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ISciReadCBDac_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ISciReadCBDac_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ISciReadCBDac_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __ISciReadCBDac_DISPINTERFACE_DEFINED__ */


#ifndef ___SciReadCBDac_DISPINTERFACE_DEFINED__
#define ___SciReadCBDac_DISPINTERFACE_DEFINED__

/* dispinterface _SciReadCBDac */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__SciReadCBDac;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("C831A675-3AE1-4B34-9DE3-BE58E9CD7B55")
    _SciReadCBDac : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _SciReadCBDacVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _SciReadCBDac * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _SciReadCBDac * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _SciReadCBDac * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _SciReadCBDac * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _SciReadCBDac * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _SciReadCBDac * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _SciReadCBDac * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } _SciReadCBDacVtbl;

    interface _SciReadCBDac
    {
        CONST_VTBL struct _SciReadCBDacVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _SciReadCBDac_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _SciReadCBDac_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _SciReadCBDac_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _SciReadCBDac_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _SciReadCBDac_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _SciReadCBDac_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _SciReadCBDac_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___SciReadCBDac_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_SciReadCBDac;

#ifdef __cplusplus

class DECLSPEC_UUID("DF890F52-1AC3-42E3-9DB6-958AB48B87AE")
SciReadCBDac;
#endif

#ifndef __IMyContainer_DISPINTERFACE_DEFINED__
#define __IMyContainer_DISPINTERFACE_DEFINED__

/* dispinterface IMyContainer */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID_IMyContainer;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("1AB5F9FD-EC5C-4E75-8FA1-8D586232951E")
    IMyContainer : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IMyContainerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IMyContainer * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IMyContainer * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IMyContainer * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IMyContainer * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IMyContainer * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IMyContainer * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IMyContainer * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } IMyContainerVtbl;

    interface IMyContainer
    {
        CONST_VTBL struct IMyContainerVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IMyContainer_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IMyContainer_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IMyContainer_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IMyContainer_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IMyContainer_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IMyContainer_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IMyContainer_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IMyContainer_DISPINTERFACE_DEFINED__ */


#ifndef __IMyExtended_DISPINTERFACE_DEFINED__
#define __IMyExtended_DISPINTERFACE_DEFINED__

/* dispinterface IMyExtended */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID_IMyExtended;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("44E8D280-12DB-438F-B246-13BA19F66B7D")
    IMyExtended : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IMyExtendedVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IMyExtended * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IMyExtended * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IMyExtended * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IMyExtended * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IMyExtended * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IMyExtended * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IMyExtended * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } IMyExtendedVtbl;

    interface IMyExtended
    {
        CONST_VTBL struct IMyExtendedVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IMyExtended_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IMyExtended_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IMyExtended_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IMyExtended_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IMyExtended_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IMyExtended_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IMyExtended_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IMyExtended_DISPINTERFACE_DEFINED__ */

#endif /* __SciReadCBDac_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


