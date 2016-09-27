#include "StdAfx.h"
#include "MyObject.h"
#include "Server.h"
#include "dispids.h"
#include "MyGuids.h"
#include "MyReadCBDac.h"
#include "DlgRunCBDac.h"
#include "AutoTimeConstant.h"

CMyObject::CMyObject(IUnknown * pUnkOuter) :
	m_pImpIDispatch(NULL),
	m_pImpIConnectionPointContainer(NULL),
	m_pImpIProvideClassInfo2(NULL),
	m_pImpIPersistStreamInit(NULL),
	m_pImp_IReadAnalog(NULL),
	// outer unknown for aggregation
	m_pUnkOuter(pUnkOuter),
	// object reference count
	m_cRefs(0),
	// _IReadAnalog interface
	m_iid_IReadAnalog(IID_NULL),
	// flag am scanning
	m_fAmScanning(FALSE),
	// flag allow am scanning change
	m_fAllowChangeAmScanning(TRUE),
	// start scanning flag
	m_fStartScanning(FALSE),
	m_pMyReadCBDac(NULL),
	// flag want signal integration
	m_fWantIntegrateSignal(TRUE),
	// auto time constant
//	m_fWantAutoTimeConstant(FALSE),
	m_pAutoTimeConstant(NULL),
	// array scan
	m_nArray(0),
	m_pArray(NULL)
{
	if (NULL == this->m_pUnkOuter) this->m_pUnkOuter = this;
	for (ULONG i=0; i<MAX_CONN_PTS; i++)
		this->m_paConnPts[i]	= NULL;
}

CMyObject::~CMyObject(void)
{
	Utils_DELETE_POINTER(this->m_pImpIDispatch);
	Utils_DELETE_POINTER(this->m_pImpIConnectionPointContainer);
	Utils_DELETE_POINTER(this->m_pImpIProvideClassInfo2);
	Utils_DELETE_POINTER(this->m_pImpIPersistStreamInit);
	Utils_DELETE_POINTER(this->m_pImp_IReadAnalog);
	for (ULONG i=0; i<MAX_CONN_PTS; i++)
		Utils_RELEASE_INTERFACE(this->m_paConnPts[i]);
	Utils_DELETE_POINTER(this->m_pMyReadCBDac);
	Utils_DELETE_POINTER(this->m_pAutoTimeConstant);
	this->ClearArrayScan();
}

// clear the array scan
void CMyObject::ClearArrayScan()
{
	if (NULL != this->m_pArray)
	{
		delete[] this->m_pArray;
	}
	this->m_pArray = NULL;
	this->m_nArray = 0;
}


// IUnknown methods
STDMETHODIMP CMyObject::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	*ppv = NULL;
	if (IID_IUnknown == riid)
		*ppv = this;
	else if (IID_IDispatch == riid || MY_IID == riid)
		*ppv = this->m_pImpIDispatch;
	else if (IID_IConnectionPointContainer == riid)
		*ppv = this->m_pImpIConnectionPointContainer;
	else if (IID_IProvideClassInfo == riid || IID_IProvideClassInfo2 == riid)
		*ppv = this->m_pImpIProvideClassInfo2;
	else if (IID_IPersist == riid || IID_IPersistStreamInit == riid)
		*ppv = this->m_pImpIPersistStreamInit;
	else if (riid == this->m_iid_IReadAnalog)
		*ppv = this->m_pImp_IReadAnalog;
	if (NULL != *ppv)
	{
		((IUnknown*)*ppv)->AddRef();
		return S_OK;
	}
	else
	{
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CMyObject::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyObject::Release()
{
	ULONG			cRefs;
	cRefs = --m_cRefs;
	if (0 == cRefs)
	{
		m_cRefs++;
		GetServer()->ObjectsDown();
		delete this;
	}
	return cRefs;
}

// initialization
HRESULT CMyObject::Init()
{
	HRESULT						hr;

	this->m_pImpIDispatch					= new CImpIDispatch(this, this->m_pUnkOuter);
	this->m_pImpIConnectionPointContainer	= new CImpIConnectionPointContainer(this, this->m_pUnkOuter);
	this->m_pImpIProvideClassInfo2			= new CImpIProvideClassInfo2(this, this->m_pUnkOuter);
	this->m_pImpIPersistStreamInit			= new CImpIPersistStreamInit(this, this->m_pUnkOuter);
	this->m_pImp_IReadAnalog				= new CImp_IReadAnalog(this, this->m_pUnkOuter);
	this->m_pMyReadCBDac					= new CMyReadCBDac(this);
	if (NULL != this->m_pImpIDispatch					&&
		NULL != this->m_pImpIConnectionPointContainer	&&
		NULL != this->m_pImpIProvideClassInfo2			&&
		NULL != this->m_pImpIPersistStreamInit			&&
		NULL != this->m_pImp_IReadAnalog				&&
		NULL != this->m_pMyReadCBDac)
	{
		// create connection points
		hr = Utils_CreateConnectionPoint(this, MY_IIDSINK, &(this->m_paConnPts[CONN_PT_CUSTOMSINK]));
		if (SUCCEEDED(hr)) hr = Utils_CreateConnectionPoint(this, IID_IPropertyNotifySink, &(this->m_paConnPts[CONN_PT_PROPNOTIFY]));
	}
	else hr = E_OUTOFMEMORY;
	return hr;
}

HRESULT CMyObject::GetClassInfo(
									ITypeInfo	**	ppTI)
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;
	*ppTI		= NULL;
	hr = GetServer()->GetTypeLib(&pTypeLib);
	if (SUCCEEDED(hr))
	{
		hr = pTypeLib->GetTypeInfoOfGuid(MY_CLSID, ppTI);
		pTypeLib->Release();
	}
	return hr;
}

HRESULT CMyObject::GetRefTypeInfo(
									LPCTSTR			szInterface,
									ITypeInfo	**	ppTypeInfo)
{
	HRESULT			hr;
	ITypeInfo	*	pClassInfo;
	BOOL			fSuccess	= FALSE;
	*ppTypeInfo	= NULL;
	hr = this->GetClassInfo(&pClassInfo);
	if (SUCCEEDED(hr))
	{
		fSuccess = Utils_FindImplClassName(pClassInfo, szInterface, ppTypeInfo);
		pClassInfo->Release();
	}
	return fSuccess ? S_OK : E_FAIL;
}

// sink events
HWND CMyObject::FireQueryMainWindow()
{
	VARIANTARG		varg;
	long			lval = 0;
	VariantInit(&varg);
	varg.vt = VT_BYREF | VT_I4;
	varg.plVal = &lval;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_QueryMainWindow, &varg, 1);
	return (HWND)lval;
}

BOOL CMyObject::FireRequestConfigFile(
	LPTSTR			szConfigFile,
	UINT			nBufferSize)
{
	VARIANTARG		varg;
	BSTR			bstr = NULL;
	BOOL			fSuccess = FALSE;
	szConfigFile[0] = '\0';
	VariantInit(&varg);
	varg.vt = VT_BYREF | VT_BSTR;
	varg.pbstrVal = &bstr;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestConfigFile, &varg, 1);
	if (NULL != bstr)
	{
		StringCchCopy(szConfigFile, nBufferSize, bstr);
		fSuccess = TRUE;
		SysFreeString(bstr);
	}
	return fSuccess;
}

BOOL CMyObject::FireRequestIniFile(
	LPTSTR			szIniFile,
	UINT			nBufferSize)
{
	VARIANTARG		varg;
	BSTR			bstr = NULL;
	BOOL			fSuccess = FALSE;
	szIniFile[0] = '\0';
	VariantInit(&varg);
	varg.vt = VT_BYREF | VT_BSTR;
	varg.pbstrVal = &bstr;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestIniFile, &varg, 1);
	if (NULL != bstr)
	{
		StringCchCopy(szIniFile, nBufferSize, bstr);
		fSuccess = TRUE;
		SysFreeString(bstr);
	}
	return fSuccess;
}

// get the working directory
BOOL CMyObject::GetWorkingDirectory(
	LPTSTR			szWorkingDirectory,
	UINT			nBufferSize)
{
	TCHAR			szConfig[MAX_PATH];
	BOOL			fSuccess = FALSE;
	szWorkingDirectory[0] = '\0';
	if (this->FireRequestConfigFile(szConfig, MAX_PATH))
	{
		PathRemoveFileSpec(szConfig);
		StringCchCopy(szWorkingDirectory, nBufferSize, szConfig);
		fSuccess = TRUE;
	}
	return fSuccess;
}

BOOL CMyObject::FireRequestCurrentPosition(
	double		*	position)
{
	VARIANTARG		avarg[2];
	VARIANT_BOOL	handled = VARIANT_FALSE;
	*position = 0.0;
	VariantInit(&avarg[1]);
	avarg[1].vt = VT_BYREF | VT_R8;
	avarg[1].pdblVal = position;
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_BOOL;
	avarg[0].pboolVal = &handled;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestCurrentPosition, avarg, 2);
	return VARIANT_TRUE == handled;
}

BOOL CMyObject::FireRequestSetPosition(
	double			newPosition)
{
	VARIANTARG		avarg[2];
	VARIANT_BOOL	handled = VARIANT_FALSE;
	InitVariantFromDouble(newPosition, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_BOOL;
	avarg[0].pboolVal = &handled;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestSetPosition, avarg, 2);
	return VARIANT_TRUE == handled;
}

void CMyObject::FireError(
	LPCTSTR			Error)
{
	VARIANTARG			varg;
	InitVariantFromString(Error, &varg);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_Error, &varg, 1);
	VariantClear(&varg);
}

BOOL CMyObject::FireRequestScanningStatus()
{
	VARIANT_BOOL	amScanning = VARIANT_FALSE;
	VARIANTARG		varg;
	VariantInit(&varg);
	varg.vt = VT_BYREF | VT_BOOL;
	varg.pboolVal = &amScanning;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestScanningStatus, &varg, 1);
	return VARIANT_TRUE == amScanning;
}

BOOL CMyObject::FireRequestWorkingDirectory(
	LPTSTR			szWorkingDirectory,
	UINT			nBufferSize)
{
	VARIANTARG		avarg[2];
	VARIANT_BOOL	fHandled = VARIANT_FALSE;
	BSTR			bstr = NULL;
	szWorkingDirectory[0] = '\0';
	VariantInit(&avarg[1]);
	avarg[1].vt = VT_BYREF | VT_BSTR;
	avarg[1].pbstrVal = &bstr;
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_BOOL;
	avarg[0].pboolVal = &fHandled;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestWorkingDirectory, avarg, 2);
	if (NULL != bstr)
	{
		StringCchCopy(szWorkingDirectory, nBufferSize, bstr);
		SysFreeString(bstr);
	}
	return VARIANT_TRUE == fHandled;
}

void CMyObject::FireHaveArrayScan(
	unsigned long	nArray,
	unsigned short*	arr)
{
	VARIANTARG		varg;
	unsigned long	i;
//	double		*	pSignal = NULL;
	InitVariantFromUInt16Array(arr, nArray, &varg);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_HaveArrayScan, &varg, 1);
	VariantClear(&varg);
	// convert scan to double array
	if (nArray != this->m_nArray)
	{
		this->ClearArrayScan();
		this->m_nArray = nArray;
		this->m_pArray = new double[this->m_nArray];
	}
//	pSignal = new double[nArray];
	for (i = 0; i < nArray; i++)
		this->m_pMyReadCBDac->CountsToVoltage(arr[i], &this->m_pArray[i]);
	this->FireHaveVoltageScan(nArray, this->m_pArray);
}

void CMyObject::FireHaveVoltageScan(
	unsigned long	nArray,
	double		*	arr)
{
	VARIANTARG		varg;
	InitVariantFromDoubleArray(arr, nArray, &varg);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_HaveVoltageScan, &varg, 1);
	VariantClear(&varg);
	// signal average
	unsigned long	i;
	double			sum = 0.0;
	for (i = 0; i < nArray; i++)
		sum += arr[i];
	if (nArray < 0) nArray = 1;
	this->m_pMyReadCBDac->SetSignalAverage(sum / nArray);
}

BOOL CMyObject::FirehaveIntSignal(
	unsigned long	nArray,
	double		*	intSignal,
	long			nAttempts)
{
	if (NULL != this->m_pAutoTimeConstant && this->m_pAutoTimeConstant->GetScriptLoaded())
	{
		return this->m_pAutoTimeConstant->HaveIntSignal(nArray, intSignal, (short int) nAttempts);
	}

	VARIANTARG		avarg[3];
	SAFEARRAYBOUND	sab;
	SAFEARRAY	*	pSA;
	double		*	arr;
	unsigned long	i;
	VARIANT_BOOL	fContinue = VARIANT_FALSE;

	// form the safearray
	sab.lLbound = 0;
	sab.cElements = nArray;
	pSA = SafeArrayCreate(VT_R8, 1, &sab);
	SafeArrayAccessData(pSA, (void**)&arr);
	for (i = 0; i < nArray; i++)
	{
		arr[i] = intSignal[i];
	}
	SafeArrayUnaccessData(pSA);
	VariantInit(&avarg[2]);
	avarg[2].vt = VT_ARRAY | VT_BYREF | VT_R8;
	avarg[2].pparray = &pSA;
	VariantInit(&avarg[1]);
	avarg[1].vt = VT_BYREF | VT_I4;
	avarg[1].plVal = &nAttempts;
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_BOOL;
	avarg[0].pboolVal = &fContinue;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_haveIntSignal, avarg, 3);
	SafeArrayDestroy(pSA);
	return VARIANT_TRUE == fContinue;
}

// property change notification
void CMyObject::OnPropChange(
	DISPID			dispid)
{
	Utils_OnPropChanged(this, dispid);
}

CMyObject::CImpIDispatch::CImpIDispatch(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter),
	m_pTypeInfo(NULL)
{
}

CMyObject::CImpIDispatch::~CImpIDispatch()
{
	Utils_RELEASE_INTERFACE(this->m_pTypeInfo);
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIDispatch::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIDispatch::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIDispatch::Release()
{
	return this->m_punkOuter->Release();
}

// IDispatch methods
STDMETHODIMP CMyObject::CImpIDispatch::GetTypeInfoCount( 
									PUINT			pctinfo)
{
	*pctinfo	= 1;
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIDispatch::GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo)
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;

	*ppTInfo	= NULL;
	if (0 != iTInfo) return DISP_E_BADINDEX;
	if (NULL == this->m_pTypeInfo)
	{
		hr = GetServer()->GetTypeLib(&pTypeLib);
		if (SUCCEEDED(hr))
		{
			hr = pTypeLib->GetTypeInfoOfGuid(MY_IID, &(this->m_pTypeInfo));
			pTypeLib->Release();
		}
	}
	else hr = S_OK;
	if (SUCCEEDED(hr))
	{
		*ppTInfo	= this->m_pTypeInfo;
		this->m_pTypeInfo->AddRef();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIDispatch::GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId)
{
	HRESULT					hr;
	ITypeInfo			*	pTypeInfo;
	hr = this->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgDispId);
		pTypeInfo->Release();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIDispatch::Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr) 
{
	switch (dispIdMember)
	{
	case DISPID_BoardNumber:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetBoardNumber(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetBoardNumber(pDispParams);
		break;
	case DISPID_AmInitialized:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetAmInitialized(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetAmInitialized(pDispParams);
		break;
	case DISPID_ChannelNumber:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetChannelNumber(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetChannelNumber(pDispParams);
		break;
	case DISPID_IntegratingSignal:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetIntegratingSignal(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetIntegratingSignal(pDispParams);
		break;
	case DISPID_IntegrationTime:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetIntegrationTime(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetIntegrationTime(pDispParams);
		break;
	case DISPID_ScanRate:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetScanRate(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetScanRate(pDispParams);
		break;
	case DISPID_ADResolution:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetADResolution(pVarResult);
		break;
	case DISPID_VoltageRange:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetVoltageRange(pDispParams, pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetVoltageRange(pDispParams);
		break;
	case DISPID_VoltageRanges:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetVoltageRanges(pDispParams, pVarResult);
		break;
	case DISPID_GetSignalReading:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetSingleReading(pVarResult);
		break;
	case DISPID_ConfigDigitalBit:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ConfigDigitalBit(pDispParams, pVarResult);
		break;
	case DISPID_WriteDigitalBit:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->WriteDigitalBit(pDispParams, pVarResult);
		break;
	case DISPID_ReadDigitalBit:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ReadDigitalBit(pDispParams, pVarResult);
		break;
	case DISPID_OutputAnalog:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->OutputAnalog(pDispParams, pVarResult);
		break;
	case DISPID_ReadAnalogInput:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ReadAnalogInput(pDispParams, pVarResult);
		break;
	case DISPID_CountsToVoltage:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->CountsToVoltage(pDispParams, pVarResult);
		break;
	case DISPID_VoltageToCounts:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->VoltageToCounts(pDispParams, pVarResult);
		break;
	case DISPID_Setup:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->Setup(pDispParams);
		break;
	case DISPID_WantIntegrateSignal:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetWantIntegrateSignal(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetWantIntegrateSignal(pDispParams);
		break;
	case DISPID_WantAutoTimeConstant:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetWantAutoTimeConstant(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetWantAutoTimeConstant(pDispParams);
		break;
	case DISPID_SetupAutoTimeConstant:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SetupAutoTimeConstant(pDispParams);
		break;
	case DISPID_TestAutoTimeConstant:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->TestAutoTimeConstant(pDispParams, pVarResult);
		break;
	default:
		break;
	}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CMyObject::CImpIDispatch::GetBoardNumber(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pBackObj->m_pMyReadCBDac->GetBoardNumber(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetBoardNumber(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyReadCBDac->SetBoardNumber(varg.lVal);
	Utils_OnPropChanged(this, DISPID_BoardNumber);
	return S_OK;

	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetAmInitialized(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pMyReadCBDac->GetAmInitialized(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetAmInitialized(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyReadCBDac->SetAmInitialized(VARIANT_TRUE == varg.boolVal);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetChannelNumber(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pBackObj->m_pMyReadCBDac->GetChannelNumber(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetChannelNumber(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyReadCBDac->SetChannelNumber(varg.lVal);
	Utils_OnPropChanged(this->m_pBackObj, DISPID_ChannelNumber);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetIntegratingSignal(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pMyReadCBDac->GetIntegratingSignal(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetIntegratingSignal(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyReadCBDac->SetIntegratingSignal(VARIANT_TRUE == varg.boolVal);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetIntegrationTime(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pBackObj->m_pMyReadCBDac->GetIntegrationTime(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetIntegrationTime(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyReadCBDac->SetIntegrationTime(varg.lVal);
	Utils_OnPropChanged(this->m_pBackObj, DISPID_IntegrationTime);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetScanRate(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pBackObj->m_pMyReadCBDac->GetScanRate(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetScanRate(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyReadCBDac->SetScanRate(varg.lVal);
	Utils_OnPropChanged(this, DISPID_ScanRate);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetVoltageRange(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_R8) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	double			minVoltage;
	double			maxVoltage;
	this->m_pBackObj->m_pMyReadCBDac->GetVoltageRange(&minVoltage, &maxVoltage);
	*(pDispParams->rgvarg[0].pdblVal) = minVoltage;
	InitVariantFromDouble(maxVoltage, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetVoltageRange(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;

	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_R8) != pDispParams->rgvarg[1].vt) return DISP_E_TYPEMISMATCH;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyReadCBDac->SetVoltageRange(*(pDispParams->rgvarg[1].pdblVal), varg.dblVal);
	Utils_OnPropChanged(this->m_pBackObj, DISPID_VoltageRange);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetADResolution(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pBackObj->m_pMyReadCBDac->GetADResolution(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetVoltageRanges(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_ARRAY | VT_R8) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	long			nRanges;
	double		*	minValues;
	double		*	maxValues;
	SAFEARRAYBOUND	sab;
	long			i;
	SAFEARRAY	*	pSA;
	double		*	arr;

	if (this->m_pBackObj->m_pMyReadCBDac->GetVoltageRanges(&nRanges, &minValues, &maxValues))
	{
		sab.lLbound = 0;
		sab.cElements = nRanges;
		pSA = SafeArrayCreate(VT_R8, 1, &sab);
		SafeArrayAccessData(pSA, (void**)&arr);
		for (i = 0; i < nRanges; i++)
			arr[i] = minValues[i];
		SafeArrayUnaccessData(pSA);
		*(pDispParams->rgvarg[0].pparray) = pSA;
		InitVariantFromDoubleArray(maxValues, (ULONG)nRanges, pVarResult);
		delete[] minValues;
		delete[] maxValues;
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetSingleReading(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	unsigned short	dataValue;
	long			channel = this->m_pBackObj->m_pMyReadCBDac->GetChannelNumber();
	double			voltage;
	if (this->m_pBackObj->m_pMyReadCBDac->ReadAnalogInput(channel, &dataValue))
	{
		this->m_pBackObj->m_pMyReadCBDac->CountsToVoltage(dataValue, &voltage);
		InitVariantFromDouble(voltage, pVarResult);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::ConfigDigitalBit(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	short			bitNumber;
	BOOL			fOutput;
	BOOL			fSuccess = FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I2, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	bitNumber = varg.iVal;
	hr = DispGetParam(pDispParams, 1, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	fOutput = VARIANT_TRUE == varg.boolVal;
	fSuccess = this->m_pBackObj->m_pMyReadCBDac->ConfigDigitalBit(bitNumber, fOutput);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::WriteDigitalBit(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	short			bitNumber;
	BOOL			fHi;
	BOOL			fSuccess = FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I2, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	bitNumber = varg.iVal;
	hr = DispGetParam(pDispParams, 1, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	fHi = VARIANT_TRUE == varg.boolVal;
	fSuccess = this->m_pBackObj->m_pMyReadCBDac->WriteDigitalBit(bitNumber, fHi);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::ReadDigitalBit(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	short			bitNumber;
	BOOL			fHi;
	BOOL			fSuccess = FALSE;
	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_BOOL) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I2, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	bitNumber = varg.iVal;
	if (this->m_pBackObj->m_pMyReadCBDac->ReadDigitalBit(bitNumber, &fHi))
	{
		*(pDispParams->rgvarg[0].pboolVal) = fHi ? VARIANT_TRUE : VARIANT_FALSE;
		fSuccess = TRUE;
	}
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::OutputAnalog(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	short int		channel;
	double			value;
	BOOL			fSuccess = FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I2, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	channel = varg.iVal;
	hr = DispGetParam(pDispParams, 1, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	value = varg.dblVal;
	fSuccess = this->m_pBackObj->m_pMyReadCBDac->OutputAnalog(channel, value);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::ReadAnalogInput(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	short int			channel;
	unsigned short		dataValue;
	BOOL				fSuccess = FALSE;
	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_UI2) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I2, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	channel = varg.iVal;
	if (this->m_pBackObj->m_pMyReadCBDac->ReadAnalogInput(channel, &dataValue))
	{
		*(pDispParams->rgvarg[0].puiVal) = dataValue;
		fSuccess = TRUE;
	}
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::CountsToVoltage(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	unsigned short	dataValue;
	double			voltage;
	BOOL			fSuccess = FALSE;
	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_R8) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_UI2, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	dataValue = varg.uiVal;
	if (this->m_pBackObj->m_pMyReadCBDac->CountsToVoltage(dataValue, &voltage))
	{
		*(pDispParams->rgvarg[0].pdblVal) = voltage;
		fSuccess = TRUE;
	}
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::VoltageToCounts(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	double			voltage;
	unsigned short	dataValue;
	BOOL			fSuccess = FALSE;
	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_UI2) != pDispParams->rgvarg[1].vt) return DISP_E_TYPEMISMATCH;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	voltage = varg.dblVal;
	if (this->m_pBackObj->m_pMyReadCBDac->VoltageToCounts(voltage, &dataValue))
	{
		*(pDispParams->rgvarg[0].puiVal) = dataValue;
		fSuccess = TRUE;
	}
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::Setup(
	DISPPARAMS	*	pDispParams)
{
	CDlgRunCBDac		dlg(this->m_pBackObj);
	dlg.DoOpenDialog(GetServer()->GetInstance(), NULL);

	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetWantIntegrateSignal(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_fWantIntegrateSignal, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetWantIntegrateSignal(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_fWantIntegrateSignal = VARIANT_TRUE == varg.boolVal;
	Utils_OnPropChanged(this->m_pBackObj, DISPID_WantIntegrateSignal);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetWantAutoTimeConstant(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
//	InitVariantFromBoolean(this->m_pBackObj->m_fWantAutoTimeConstant, pVarResult);
	BOOL			fAutoTimeConstant = FALSE;
	if (NULL != this->m_pBackObj->m_pAutoTimeConstant)
		fAutoTimeConstant = this->m_pBackObj->m_pAutoTimeConstant->GetScriptLoaded();
	InitVariantFromBoolean(fAutoTimeConstant, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetWantAutoTimeConstant(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	if (VARIANT_TRUE == varg.boolVal)
	{
		if (NULL == this->m_pBackObj->m_pAutoTimeConstant)
		{
			this->m_pBackObj->m_pAutoTimeConstant = new CAutoTimeConstant(this->m_pBackObj);
		}
//		this->m_pBackObj->m_fWantAutoTimeConstant = TRUE;
		this->m_pBackObj->m_pAutoTimeConstant->SetScriptLoaded(TRUE);
	}
	else
	{
//		this->m_pBackObj->m_fWantAutoTimeConstant = FALSE;
		if (NULL != this->m_pBackObj->m_pAutoTimeConstant)
		{
			this->m_pBackObj->m_pAutoTimeConstant->SetScriptLoaded(FALSE);
		}
	}
	Utils_OnPropChanged(this->m_pBackObj, DISPID_WantAutoTimeConstant);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetupAutoTimeConstant(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	if (NULL != this->m_pBackObj->m_pAutoTimeConstant)
	{
		this->m_pBackObj->m_pAutoTimeConstant->DisplaySetup((HWND)varg.lVal);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::TestAutoTimeConstant(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
//	HRESULT			hr;
//	VARIANTARG		varg;
//	UINT			uArgErr;
//	HWND			hwnd;
	TCHAR			szResult[MAX_PATH];
	long			nAttempts = 1;
	BOOL			fContinue;
	BOOL			fDone = FALSE;
	double			signal;
	// make sure want auto time constant
	Utils_SetBoolProperty(this, DISPID_WantAutoTimeConstant, TRUE);
	while (!fDone)
	{
		signal = this->m_pBackObj->m_pMyReadCBDac->IntegrateSignal();
		fContinue = this->m_pBackObj->FirehaveIntSignal(
			this->m_pBackObj->m_nArray, this->m_pBackObj->m_pArray, nAttempts);
		if (fContinue || nAttempts >= 5)
			fDone = TRUE;
		else
		{
			nAttempts++;
		}
	}
	// form the message string
	StringCchPrintf(szResult, MAX_PATH, L"AutoTimeConstant test\nNumber of Attempts = %1d\nCurrent integration time = %1d",
		nAttempts, Utils_GetLongProperty(this, DISPID_IntegrationTime));
//	MessageBox(NULL, szResult, L"SciReadCBDac", MB_OK);
	if (NULL != pVarResult) InitVariantFromString(szResult, pVarResult);
	return S_OK;
}

CMyObject::CImpIProvideClassInfo2::CImpIProvideClassInfo2(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIProvideClassInfo2::~CImpIProvideClassInfo2()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIProvideClassInfo2::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIProvideClassInfo2::Release()
{
	return this->m_punkOuter->Release();
}

// IProvideClassInfo method
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::GetClassInfo(
									ITypeInfo	**	ppTI) 
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;
	*ppTI		= NULL;
	hr = GetServer()->GetTypeLib(&pTypeLib);
	if (SUCCEEDED(hr))
	{
		hr = pTypeLib->GetTypeInfoOfGuid(MY_CLSID, ppTI);
		pTypeLib->Release();
	}
	return hr;
}

// IProvideClassInfo2 method
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::GetGUID(
									DWORD			dwGuidKind,  //Desired GUID
									GUID		*	pGUID)
{
	if (GUIDKIND_DEFAULT_SOURCE_DISP_IID == dwGuidKind)
	{
		*pGUID	= MY_IIDSINK;
		return S_OK;
	}
	else
	{
		*pGUID	= GUID_NULL;
		return E_INVALIDARG;	
	}
}

CMyObject::CImpIConnectionPointContainer::CImpIConnectionPointContainer(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIConnectionPointContainer::~CImpIConnectionPointContainer()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIConnectionPointContainer::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIConnectionPointContainer::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIConnectionPointContainer::Release()
{
	return this->m_punkOuter->Release();
}

// IConnectionPointContainer methods
STDMETHODIMP CMyObject::CImpIConnectionPointContainer::EnumConnectionPoints(
									IEnumConnectionPoints **ppEnum)
{
	return Utils_CreateEnumConnectionPoints(this, MAX_CONN_PTS, this->m_pBackObj->m_paConnPts,
		ppEnum);
}

STDMETHODIMP CMyObject::CImpIConnectionPointContainer::FindConnectionPoint(
									REFIID			riid,  //Requested connection point's interface identifier
									IConnectionPoint **ppCP)
{
	HRESULT					hr		= CONNECT_E_NOCONNECTION;
	IConnectionPoint	*	pConnPt	= NULL;
	*ppCP	= NULL;
	if (MY_IIDSINK == riid)
		pConnPt = this->m_pBackObj->m_paConnPts[CONN_PT_CUSTOMSINK];
	else if (IID_IPropertyNotifySink == riid)
		pConnPt = this->m_pBackObj->m_paConnPts[CONN_PT_PROPNOTIFY];
	if (NULL != pConnPt)
	{
		*ppCP = pConnPt;
		pConnPt->AddRef();
		hr = S_OK;
	}
	return hr;
}

CMyObject::CImpIPersistStreamInit::CImpIPersistStreamInit(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{

}

CMyObject::CImpIPersistStreamInit::~CImpIPersistStreamInit()
{

}
// IUnknown methods
STDMETHODIMP CMyObject::CImpIPersistStreamInit::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}
STDMETHODIMP_(ULONG) CMyObject::CImpIPersistStreamInit::AddRef()
{
	return this->m_punkOuter->AddRef();
}
STDMETHODIMP_(ULONG) CMyObject::CImpIPersistStreamInit::Release()
{
	return this->m_punkOuter->Release();
}
// IPersist method
STDMETHODIMP CMyObject::CImpIPersistStreamInit::GetClassID(
	CLSID		*	pClassID)
{
	*pClassID = MY_CLSID;
	return S_OK;
}
// IPersistStreamInit methods
STDMETHODIMP CMyObject::CImpIPersistStreamInit::GetSizeMax(
	ULARGE_INTEGER *pCbSize)
{
	pCbSize->QuadPart = 100;
	return S_OK;
}
STDMETHODIMP CMyObject::CImpIPersistStreamInit::InitNew()
{
	return S_OK;
}
STDMETHODIMP CMyObject::CImpIPersistStreamInit::IsDirty()
{
	return S_FALSE;
}
STDMETHODIMP CMyObject::CImpIPersistStreamInit::Load(
	LPSTREAM		pStm)
{
	return S_OK;
}
STDMETHODIMP CMyObject::CImpIPersistStreamInit::Save(
	LPSTREAM		pStm,
	BOOL			fClearDirty)
{
	return S_OK;
}

CMyObject::CImp_IReadAnalog::CImp_IReadAnalog(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter),
	m_pTypeInfo(NULL),
	m_dispidReadINI(DISPID_UNKNOWN),
	m_dispidWriteINI(DISPID_UNKNOWN),
	m_dispidReadCFG(DISPID_UNKNOWN),
	m_dispidInit(DISPID_UNKNOWN),
	m_dispidAmIOpen(DISPID_UNKNOWN),
	m_dispidDoSetup(DISPID_UNKNOWN),
	m_dispidGetVoltageRange(DISPID_UNKNOWN),
	m_dispidSetOutVolts(DISPID_UNKNOWN),
	m_dispidSetOutputLine(DISPID_UNKNOWN),
	m_dispidSetInputLine(DISPID_UNKNOWN),
	m_dispidGetMaximumCounts(DISPID_UNKNOWN),
	m_dispidCheckWarnings(DISPID_UNKNOWN),
	m_dispidSignalType(DISPID_UNKNOWN),
	m_dispidReadSignal(DISPID_UNKNOWN),
	m_dispidGetInfoString(DISPID_UNKNOWN),
	m_dispidAllowSignalAveraging(DISPID_UNKNOWN),
	m_dispidAllowScanAveraging(DISPID_UNKNOWN),
	m_dispidAllowMultipleScans(DISPID_UNKNOWN),
	m_dispidSetDwellTime(DISPID_UNKNOWN),
	m_dispidSignalIntegrationAvailable(DISPID_UNKNOWN),
	m_dispidIntegrationTime(DISPID_UNKNOWN),
	m_dispidAutoIntegrationTime(DISPID_UNKNOWN),
	m_dispidSetupAutoIntegrationTime(DISPID_UNKNOWN)
{
	this->Init();
}

CMyObject::CImp_IReadAnalog::~CImp_IReadAnalog()
{
	Utils_RELEASE_INTERFACE(this->m_pTypeInfo);
}

// IUnknown methods
STDMETHODIMP CMyObject::CImp_IReadAnalog::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImp_IReadAnalog::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImp_IReadAnalog::Release()
{
	return this->m_punkOuter->Release();
}

// IDispatch methods
STDMETHODIMP CMyObject::CImp_IReadAnalog::GetTypeInfoCount(
	PUINT			pctinfo)
{
	*pctinfo = 1;
	return S_OK;
}

STDMETHODIMP CMyObject::CImp_IReadAnalog::GetTypeInfo(
	UINT			iTInfo,
	LCID			lcid,
	ITypeInfo	**	ppTInfo)
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;

	*ppTInfo = NULL;
	if (0 != iTInfo) return DISP_E_BADINDEX;
	if (NULL == this->m_pTypeInfo)
	{
		hr = GetServer()->GetTypeLib(&pTypeLib);
		if (SUCCEEDED(hr))
		{
			hr = pTypeLib->GetTypeInfoOfGuid(
				this->m_pBackObj->m_iid_IReadAnalog, &(this->m_pTypeInfo));
			pTypeLib->Release();
		}
	}
	else hr = S_OK;
	if (SUCCEEDED(hr))
	{
		*ppTInfo = this->m_pTypeInfo;
		this->m_pTypeInfo->AddRef();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImp_IReadAnalog::GetIDsOfNames(
	REFIID			riid,
	OLECHAR		**  rgszNames,
	UINT			cNames,
	LCID			lcid,
	DISPID		*	rgDispId)
{
	HRESULT					hr;
	ITypeInfo			*	pTypeInfo;
	hr = this->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgDispId);
		pTypeInfo->Release();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImp_IReadAnalog::Invoke(
	DISPID			dispIdMember,
	REFIID			riid,
	LCID			lcid,
	WORD			wFlags,
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult,
	EXCEPINFO	*	pExcepInfo,
	PUINT			puArgErr)
{
	if (dispIdMember == this->m_dispidReadINI)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->ReadINI(pDispParams);
		}
	}
	else if (dispIdMember == this->m_dispidWriteINI)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->WriteINI(pDispParams);
		}
	}
	else if (dispIdMember == this->m_dispidReadCFG)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->ReadCFG(pDispParams);
		}
	}
	else if (dispIdMember == this->m_dispidInit)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->intInit();
		}
	}
	else if (dispIdMember == this->m_dispidAmIOpen)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->AmIOpen(pVarResult);
		}
	}
	else if (dispIdMember == this->m_dispidDoSetup)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->DoSetup();
		}
	}
	else if (dispIdMember == this->m_dispidGetVoltageRange)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->GetVoltageRange(pDispParams);
		}
	}
	else if (dispIdMember == this->m_dispidSetOutVolts)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->SetOutVolts(pDispParams, pVarResult);
		}
	}
	else if (dispIdMember == this->m_dispidSetOutputLine)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->SetOutputLine(pDispParams);
		}
	}
	else if (dispIdMember == this->m_dispidSetInputLine)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->SetInputLine(pDispParams);
		}
	}
	else if (dispIdMember == this->m_dispidGetMaximumCounts)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->GetMaximumCounts(pVarResult);
		}
	}
	else if (dispIdMember == this->m_dispidCheckWarnings)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->CheckWarnings(pDispParams, pVarResult);
		}
	}
	else if (dispIdMember == this->m_dispidSignalType)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->SignalType(pVarResult);
		}
	}
	else if (dispIdMember == this->m_dispidReadSignal)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->ReadSignal(pVarResult);
		}
	}
	else if (dispIdMember == this->m_dispidGetInfoString)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->GetInfoString(pVarResult);
		}
	}
	else if (dispIdMember == this->m_dispidAllowSignalAveraging)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetAllowSignalAveraging(pVarResult);
		}
	}
	else if (dispIdMember == this->m_dispidAllowScanAveraging)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetAllowScanAveraging(pVarResult);
		}
	}
	else if (dispIdMember == this->m_dispidAllowMultipleScans)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetAllowMultipleScans(pVarResult);
		}
	}
	else if (dispIdMember == this->m_dispidSetDwellTime)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->SetDwellTime(pDispParams, pVarResult);
		}
	}
	else if (dispIdMember == this->m_dispidSignalIntegrationAvailable)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetSignalIntegrationAvailable(pVarResult);
		}
	}
	else if (dispIdMember == this->m_dispidIntegrationTime)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetIntegrationTime(pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			return this->SetIntegrationTime(pDispParams);
		}
	}
	else if (dispIdMember == this->m_dispidAutoIntegrationTime)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetAutoIntegrationTime(pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			return this->SetAutoIntegrationTime(pDispParams);
		}
	}
	else if (dispIdMember == this->m_dispidSetupAutoIntegrationTime)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->SetupAutoIntegrationTime(pDispParams);
		}
	}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CMyObject::CImp_IReadAnalog::Init()
{
	HRESULT				hr = E_FAIL;
	ITypeInfo		*	pTypeInfo;
	TYPEATTR		*	pTypeAttr;

	hr = this->m_pBackObj->GetRefTypeInfo(L"_IReadAnalog", &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		// store the type info
		this->m_pTypeInfo = pTypeInfo;
		this->m_pTypeInfo->AddRef();
		// get the interface ID
		hr = this->m_pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			this->m_pBackObj->m_iid_IReadAnalog = pTypeAttr->guid;
			this->m_pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		// dispatch ids
		Utils_GetMemid(this->m_pTypeInfo, TEXT("ReadINI"), &m_dispidReadINI);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("WriteINI"), &m_dispidWriteINI);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("ReadCFG"), &m_dispidReadCFG);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("Init"), &m_dispidInit);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("AmIOpen"), &m_dispidAmIOpen);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("DoSetup"), &m_dispidDoSetup);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("GetVoltageRange"), &m_dispidGetVoltageRange);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("SetOutVolts"), &m_dispidSetOutVolts);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("SetOutputLine"), &m_dispidSetOutputLine);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("SetInputLine"), &m_dispidSetInputLine);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("GetMaximumCounts"), &m_dispidGetMaximumCounts);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("CheckWarnings"), &m_dispidCheckWarnings);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("SignalType"), &m_dispidSignalType);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("ReadSignal"), &m_dispidReadSignal);
		Utils_GetMemid(pTypeInfo, L"GetInfoString", &m_dispidGetInfoString);
		Utils_GetMemid(pTypeInfo, L"AllowSignalAveraging", &m_dispidAllowSignalAveraging);
		Utils_GetMemid(pTypeInfo, L"AllowScanAveraging", &m_dispidAllowScanAveraging);
		Utils_GetMemid(pTypeInfo, L"AllowMultipleScans", &m_dispidAllowMultipleScans);
		Utils_GetMemid(pTypeInfo, L"SetDwellTime", &m_dispidSetDwellTime);
		Utils_GetMemid(pTypeInfo, L"SignalIntegrationAvailable", &m_dispidSignalIntegrationAvailable);
		Utils_GetMemid(pTypeInfo, L"IntegrationTime", &m_dispidIntegrationTime);
		Utils_GetMemid(pTypeInfo, L"AutoIntegrationTime", &m_dispidAutoIntegrationTime);
		Utils_GetMemid(pTypeInfo, L"SetupAutoIntegrationTime", &m_dispidSetupAutoIntegrationTime);

		pTypeInfo->Release();
	}
	return hr;
}

HRESULT CMyObject::CImp_IReadAnalog::ReadINI(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	LPTSTR				szIniFile = NULL;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	SHStrDup(varg.bstrVal, &szIniFile);
	VariantClear(&varg);
	this->m_pBackObj->ReadIni(szIniFile);
	CoTaskMemFree((LPVOID)szIniFile);
	return S_OK;
}

HRESULT	CMyObject::CImp_IReadAnalog::WriteINI(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	LPTSTR				szIniFile = NULL;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	SHStrDup(varg.bstrVal, &szIniFile);
	VariantClear(&varg);
	this->m_pBackObj->WriteIni(szIniFile);
	CoTaskMemFree((LPVOID)szIniFile);
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::ReadCFG(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	LPTSTR				szCFGFile = NULL;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	SHStrDup(varg.bstrVal, &szCFGFile);
	VariantClear(&varg);
	this->m_pBackObj->ReadConfig(szCFGFile);
	CoTaskMemFree((LPVOID)szCFGFile);
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::intInit()		// internal initialization
{
//	this->m_pBackObj->SetAmScanning(TRUE);
	this->m_pBackObj->m_pMyReadCBDac->SetAmInitialized(TRUE);
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::AmIOpen(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
//	InitVariantFromBoolean(this->m_pBackObj->m_pThreadManager->GetAmInit(), pVarResult);
	InitVariantFromBoolean(this->m_pBackObj->m_pMyReadCBDac->GetAmInitialized(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::DoSetup()
{
//	CDlgReadAlphalas	dlg(this->m_pBackObj);
//	this->m_pBackObj->m_fAllowChangeAmScanning = FALSE;
//	dlg.DoOpenDialog(this->m_pBackObj->FireQueryMainWindow(), GetServer()->GetInstance());
//	this->m_pBackObj->m_fAllowChangeAmScanning = TRUE;
	CDlgRunCBDac		dlg(this->m_pBackObj);
	dlg.DoOpenDialog(GetServer()->GetInstance(), this->m_pBackObj->FireQueryMainWindow());
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::GetVoltageRange(
	DISPPARAMS	*	pDispParams)
{
	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_R8) != pDispParams->rgvarg[0].vt ||
		(VT_BYREF | VT_R8) != pDispParams->rgvarg[1].vt)
	{
		return DISP_E_TYPEMISMATCH;
	}
	double		minValue;
	double		maxValue;
	this->m_pBackObj->m_pMyReadCBDac->GetVoltageRange(&minValue, &maxValue);
	*(pDispParams->rgvarg[1].pdblVal) = minValue;
	*(pDispParams->rgvarg[0].pdblVal) = maxValue;			// this->m_pBackObj->m_MaxVoltageLevel;
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::SetOutVolts(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::SetOutputLine(
	DISPPARAMS	*	pDispParams)
{
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::SetInputLine(
	DISPPARAMS	*	pDispParams)
{
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::GetMaximumCounts(
	VARIANT		*	pVarResult)
{
	long		resolution = this->m_pBackObj->m_pMyReadCBDac->GetADResolution();
	if (NULL == pVarResult) return E_INVALIDARG;
	if (16 == resolution)
		InitVariantFromInt32(65535, pVarResult);
	else if (12 == resolution)
		InitVariantFromInt32(4095, pVarResult);
	return S_OK;
}

HRESULT	CMyObject::CImp_IReadAnalog::CheckWarnings(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::SignalType(
	VARIANT		*	pVarResult)
{
	// signal type is volts == 1
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(1, pVarResult);
	return S_OK;
}

// return the last signal reading
HRESULT CMyObject::CImp_IReadAnalog::ReadSignal(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	double		signal = 0.0;
	// if not scanning or not want signal integration then single reading
	if (!this->m_pBackObj->m_fWantIntegrateSignal || 
		!this->m_pBackObj->FireRequestScanningStatus())
	{
		ULONG		counts = this->m_pBackObj->m_pMyReadCBDac->GetSignalReading();
		this->m_pBackObj->m_pMyReadCBDac->CountsToVoltage((unsigned short) counts, &signal);
	}
	else
	{
		// scanning integrate signal
		// check want auto time constant
//		if (this->m_pBackObj->m_fWantAutoTimeConstant)
		if (NULL != this->m_pBackObj->m_pAutoTimeConstant && this->m_pBackObj->m_pAutoTimeConstant->GetScriptLoaded())
		{
			long		nAttempts = 0;
			BOOL		fContinue;
			BOOL		fDone = FALSE;
			while (!fDone)
			{
				signal = this->m_pBackObj->m_pMyReadCBDac->IntegrateSignal();
				fContinue = this->m_pBackObj->FirehaveIntSignal(
					this->m_pBackObj->m_nArray, this->m_pBackObj->m_pArray, nAttempts);
				if (fContinue || nAttempts >= 5)
					fDone = TRUE;
				else
				{
					nAttempts++;
				}
			}
		}
		else
		{
			// just integrate signal
			signal = this->m_pBackObj->m_pMyReadCBDac->IntegrateSignal();
		}
	}
	InitVariantFromDouble(signal, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::GetInfoString(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromString(L"SciReadCBDac", pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::GetAllowSignalAveraging(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(TRUE, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::GetAllowScanAveraging(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(TRUE, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::GetAllowMultipleScans(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(TRUE, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::SetDwellTime(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	double			dwellTime = 55.0;			// minimum dwell time in ms
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	if (varg.dblVal > dwellTime) dwellTime = varg.dblVal;
	InitVariantFromDouble(dwellTime, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::GetSignalIntegrationAvailable(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(TRUE, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::GetIntegrationTime(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	long			integrationTime = this->m_pBackObj->m_pMyReadCBDac->GetIntegrationTime();
//	InitVariantFromDouble(integrationTime / 1000.0, pVarResult);
	InitVariantFromInt32(integrationTime, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::SetIntegrationTime(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	IDispatch	*	pdisp;
	long			intTime;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	// use automation call for property change notifications
	hr = this->m_pBackObj->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		// convert to ms
//		intTime = (long)floor(varg.dblVal * 1000.0);
		intTime = varg.lVal;
		Utils_SetLongProperty(pdisp, DISPID_IntegrationTime, intTime);
		pdisp->Release();
	}
	return S_OK;
}

HRESULT CMyObject::CImp_IReadAnalog::GetAutoIntegrationTime(
	VARIANT		*	pVarResult)
{
	IDispatch		*	pdisp;
	HRESULT				hr;
	hr = this->m_pBackObj->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		hr = pdisp->Invoke(DISPID_WantAutoTimeConstant, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET,
			NULL, pVarResult, NULL, NULL);
		pdisp->Release();
	}
	return hr;
}
HRESULT CMyObject::CImp_IReadAnalog::SetAutoIntegrationTime(
	DISPPARAMS	*	pDispParams)
{
	IDispatch		*	pdisp;
	HRESULT				hr;
	hr = this->m_pBackObj->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		hr = pdisp->Invoke(DISPID_WantAutoTimeConstant, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT,
			pDispParams, NULL, NULL, NULL);
		pdisp->Release();
	}
	return hr;
}
HRESULT CMyObject::CImp_IReadAnalog::SetupAutoIntegrationTime(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	if (NULL == this->m_pBackObj->m_pAutoTimeConstant)
	{
		this->m_pBackObj->m_pAutoTimeConstant = new CAutoTimeConstant(this->m_pBackObj);
	}
	if (NULL != this->m_pBackObj->m_pAutoTimeConstant)
	{
		this->m_pBackObj->m_pAutoTimeConstant->DisplaySetup((HWND)varg.lVal);
	}
	return S_OK;
}


void CMyObject::ReadConfig(
	LPCTSTR			szConfigFile)
{

/*
	CIniFile		iniFile;
	long			nGratings;
	long			i;
	TCHAR			szInfo[MAX_PATH];
	double			nmPerPixel;
	double			quadCorrection;
	long			lval;
	TCHAR			szValue[16];
	// pixel truncation
	long			lo, hi;
	float			fval;
	iniFile.Init(szConfigFile);
	iniFile.ReadIntValue(L"SciReadAlphalas", L"NumGratings", 1, &nGratings);
	// positive sense flag
	iniFile.ReadIntValue(L"SciReadAlphalas", L"PositiveSense", 0, &lval);
	this->m_PositiveSense = 0 != lval;
	// pixel truncation
	iniFile.ReadIntValue(L"SciReadAlphalas", L"LowPixelTruncation", -1, &lval);
	lo = lval;
	iniFile.ReadIntValue(L"SciReadAlphalas", L"HighPixelTruncation", -1, &lval);
	hi = lval;
	this->m_pThreadManager->SetPixelTruncation(lo, hi);
	// grating info
	this->m_GratingInfo.clear();
	for (i = 0; i<nGratings; i++)
	{
		StringCchPrintf(szInfo, MAX_PATH, L"Grating%1dNMPerPixel", i + 1);
		iniFile.ReadDoubleValue(L"SciReadAlphalas", szInfo, 0.03, &nmPerPixel);
		StringCchPrintf(szInfo, MAX_PATH, L"Grating%1dQuadFactor", i + 1);
//		iniFile.ReadDoubleValue(L"SciReadAlphalas", szInfo, 0.0, &quadCorrection);
		quadCorrection = 0.0;
		if (iniFile.ReadStringValue(L"SciReadAlphalas", szInfo, L"", szValue, 16))
		{
			if (1 == _stscanf_s(szValue, L"%f", &fval))
			{
				quadCorrection = fval;
			}
		}

		CGratingInfo			gratInfo(i + 1, nmPerPixel, quadCorrection);
		this->m_GratingInfo.push_back(gratInfo);
	}
*/
}

void CMyObject::ReadIni(
	LPCTSTR			szIniFile)
{
/*
	CIniFile		iniFile;
	long			i;
	TCHAR			szInfo[MAX_PATH];
	size_t			nGratings = this->m_GratingInfo.size();
	double			nmPerPixel;
	double			quadCorrection;
	long			lval;
	long			lo, hi;
	TCHAR			szValue[16];
	float			fval;
	iniFile.Init(szIniFile);
	iniFile.ReadIntValue(L"SciReadAlphalas", L"FrameAveraging", 1, &lval);
//	this->m_FrameAveraging = (short int)lval;
	this->m_pThreadManager->SetFrameAveraging(lval);
	// positive sense flag
	iniFile.ReadIntValue(L"SciReadAlphalas", L"PositiveSense", 0, &lval);
	this->m_PositiveSense = 0 != lval;
	// pixel offset
	iniFile.ReadIntValue(L"SciReadAlphalas", L"PixelOffset", 0, &lval);
	this->m_PixelOffset = lval;
	// Hardware Dark Correction
	iniFile.ReadIntValue(L"SciReadAlphalas", L"HardwareDarkCorrection", 1, &lval);
	this->m_pThreadManager->SetHardwareDarkCorrection(1 == lval);
	// pixel truncation
	this->m_pThreadManager->GetPixelTruncation(&lo, &hi);
	iniFile.ReadIntValue(L"SciReadAlphalas", L"LowPixelTruncation", lo, &lval);
	lo = lval;
	iniFile.ReadIntValue(L"SciReadAlphalas", L"HighPixelTruncation", hi, &lval);
	hi = lval;
	this->m_pThreadManager->SetPixelTruncation(lo, hi);

	for (i = 0; i < nGratings; i++)
	{
		StringCchPrintf(szInfo, MAX_PATH, L"Grating%1dNMPerPixel", this->m_GratingInfo.at(i).GetGrating());
		if (iniFile.ReadStringValue(L"SciReadAlphalas", szInfo, L"", szValue, 16))
		{
			if (1 == _stscanf_s(szValue, L"%f", &fval))
			{
				this->m_GratingInfo.at(i).SetNMPerPixel(fval);
			}
		}
	//	iniFile.ReadDoubleValue(L"SciReadAlphalas", szInfo, this->m_GratingInfo.at(i).GetNMPerPixel(), &nmPerPixel);
		StringCchPrintf(szInfo, MAX_PATH, L"Grating%1dQuadFactor", this->m_GratingInfo.at(i).GetGrating());
		if (iniFile.ReadStringValue(L"SciReadAlphalas", szInfo, L"", szValue, 16))
		{
			if (1 == _stscanf_s(szValue, L"%f", &fval))
			{
				this->m_GratingInfo.at(i).SetquadraticFactor(fval);
			}
		}
	//	iniFile.ReadDoubleValue(L"SciReadAlphalas", szInfo, this->m_GratingInfo.at(i).GetquadraticFactor(), &quadCorrection);
	//	this->m_GratingInfo.at(i).SetNMPerPixel(nmPerPixel);
	//	this->m_GratingInfo.at(i).SetquadraticFactor(quadCorrection);
	}
*/
}

void CMyObject::WriteIni(
	LPCTSTR			szIniFile)
{
/*
	CIniFile		iniFile;
	long			i;
	TCHAR			szInfo[MAX_PATH];
	size_t			nGratings = this->m_GratingInfo.size();
	TCHAR			szValue[MAX_PATH];
	long			lo, hi;				// pixel truncation
	iniFile.Init(szIniFile);
	iniFile.WriteIntValue(L"SciReadAlphalas", L"FrameAveraging", 
		this->m_pThreadManager->GetFrameAveraging());
	iniFile.WriteIntValue(L"SciReadAlphalas", L"PositiveSense", this->m_PositiveSense ? 1 : 0);
	// pixel offset
	iniFile.WriteIntValue(L"SciReadAlphalas", L"PixelOffset", this->m_PixelOffset);
	// Hardware Dark Correction
	iniFile.WriteIntValue(L"SciReadAlphalas", L"HardwareDarkCorrection", this->m_pThreadManager->GetHardwareDarkCorrection() ? 1 : 0);
	this->m_pThreadManager->GetPixelTruncation(&lo, &hi);
	iniFile.WriteIntValue(L"SciReadAlphalas", L"LowPixelTruncation", lo);
	iniFile.WriteIntValue(L"SciReadAlphalas", L"HighPixelTruncation", hi);

	// loop over the gratings
	for (i = 0; i < nGratings; i++)
	{
		StringCchPrintf(szInfo, MAX_PATH, L"Grating%1dNMPerPixel", this->m_GratingInfo.at(i).GetGrating());
		_stprintf_s(szValue, L"%8.3g", this->m_GratingInfo.at(i).GetNMPerPixel());
		iniFile.WriteStringValue(L"SciReadAlphalas", szInfo, szValue);
//		iniFile.WriteDoubleValue(L"SciReadAlphalas", szInfo, this->m_GratingInfo.at(i).GetNMPerPixel());
		StringCchPrintf(szInfo, MAX_PATH, L"Grating%1dQuadFactor", this->m_GratingInfo.at(i).GetGrating());
		_stprintf_s(szValue, L"%8.3g", this->m_GratingInfo.at(i).GetquadraticFactor());
		iniFile.WriteStringValue(L"SciReadAlphalas", szInfo, szValue);
//		iniFile.WriteDoubleValue(L"SciReadAlphalas", szInfo, this->m_GratingInfo.at(i).GetquadraticFactor());
	}
*/
}
