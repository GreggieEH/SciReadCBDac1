#include "stdafx.h"
#include "AutoTimeConstant.h"
#include "MyObject.h"
#include "dispids.h"

CAutoTimeConstant::CAutoTimeConstant(CMyObject * pMyObject) :
	m_pMyObject(pMyObject),
	m_pdisp(NULL),
	m_iidSink(IID_NULL),
	m_dwCookie(0)
{
	HRESULT			hr;
	CLSID			clsid;
	CImpISink	*	pSink;
	IUnknown	*	pUnkSink;

	hr = CLSIDFromProgID(L"Sciencetech.SSAutoTimeConstant.1", &clsid);
	if (SUCCEEDED(hr)) hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (LPVOID*)&this->m_pdisp);
	if (SUCCEEDED(hr))
	{
		// connect the custom sink
		pSink = new CImpISink(this);
		hr = pSink->QueryInterface(this->m_iidSink, (LPVOID*)&pUnkSink);
		if (SUCCEEDED(hr))
		{
			Utils_ConnectToConnectionPoint(this->m_pdisp, pUnkSink, this->m_iidSink, TRUE, &this->m_dwCookie);
			pUnkSink->Release();
		}
		else
		{
			delete pSink;
		}
	}
}

CAutoTimeConstant::~CAutoTimeConstant()
{
	if (NULL != this->m_pdisp)
	{
		Utils_ConnectToConnectionPoint(this->m_pdisp, NULL, this->m_iidSink, FALSE, &this->m_dwCookie);
		Utils_RELEASE_INTERFACE(this->m_pdisp);
	}
}

BOOL CAutoTimeConstant::GetScriptLoaded()
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, L"ScriptLoaded", &dispid);
	return Utils_GetBoolProperty(this->m_pdisp, dispid);
}

void CAutoTimeConstant::SetScriptLoaded(
	BOOL		ScriptLoaded)
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, L"ScriptLoaded", &dispid);
	Utils_SetBoolProperty(this->m_pdisp, dispid, ScriptLoaded);
}

void CAutoTimeConstant::GetScriptFile(
	LPTSTR		szScriptFile,
	UINT		nBufferSize)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANT			varResult;
	szScriptFile[0] = '\0';
	Utils_GetMemid(this->m_pdisp, L"ScriptFile", &dispid);
	hr = Utils_InvokePropertyGet(this->m_pdisp, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr))
	{
		VariantToString(varResult, szScriptFile, nBufferSize);
		VariantClear(&varResult);
	}
}

void CAutoTimeConstant::SetScriptFile(
	LPCTSTR		szScriptFile)
{
	DISPID			dispid;
	VARIANTARG		varg;
	Utils_GetMemid(this->m_pdisp, L"ScriptFile", &dispid);
	InitVariantFromString(szScriptFile, &varg);
	Utils_InvokePropertyPut(this->m_pdisp, dispid, &varg, 1);
	VariantClear(&varg);
}

BOOL CAutoTimeConstant::DisplaySetup(
	HWND		hwndParent)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	Utils_GetMemid(this->m_pdisp, L"DisplaySetup", &dispid);
	InitVariantFromInt32((long)hwndParent, &varg);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

BOOL CAutoTimeConstant::HaveIntSignal(
	ULONG		nArray,
	double	*	arr,
	short int	numberOfAttempts)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANTARG		avarg[2];
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	Utils_GetMemid(this->m_pdisp, L"HaveIntSignal", &dispid);
	InitVariantFromDoubleArray(arr, nArray, &avarg[1]);
	InitVariantFromInt16(numberOfAttempts, &avarg[0]);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, avarg, 2, &varResult);
	VariantClear(&avarg[1]);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

// sink events
BOOL CAutoTimeConstant::OnRequestWorkingDirectory(
	LPTSTR		workingDirectory,
	UINT		nBufferSize)
{
	// use the same directory as the config file
	if (this->m_pMyObject->FireRequestConfigFile(workingDirectory, nBufferSize))
	{
		PathRemoveFileSpec(workingDirectory);
		return TRUE;
	}
	else
	{
		workingDirectory[0] = '\0';
		return FALSE;
	}
//
//	return this->m_pMyObject->FireRequestWorkingDirectory(workingDirectory, nBufferSize);
}

void CAutoTimeConstant::OnHaveError(
	LPCTSTR		Error)
{
	TCHAR		szString[MAX_PATH];
	StringCchPrintf(szString, MAX_PATH, L"Script Error: %s", Error);
	this->m_pMyObject->FireError(szString);
}

BOOL CAutoTimeConstant::OnRequestGetIntegrationTime(
	long	*	integrationTime)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	BOOL			fSuccess = FALSE;
	*integrationTime = 0;
	hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		*integrationTime = Utils_GetLongProperty(pdisp, DISPID_IntegrationTime);
		fSuccess = TRUE;
		pdisp->Release();
	}
	return fSuccess;
}

BOOL CAutoTimeConstant::OnRequestSetIntegrationTime(
	long		integrationTime)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	BOOL			fSuccess = FALSE;
	hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		fSuccess = TRUE;
		Utils_SetLongProperty(pdisp, DISPID_IntegrationTime, integrationTime);
		pdisp->Release();
	}
	return fSuccess;
}

void CAutoTimeConstant::OnRequestVoltageRange(
	double	*	minVoltage,
	double	*	maxVoltage)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	VARIANTARG		varg;
	VARIANT			varResult;
	*minVoltage = 0.0;
	*maxVoltage = 0.0;
	hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		VariantInit(&varg);
		varg.vt = VT_BYREF | VT_R8;
		varg.pdblVal = minVoltage;
		hr = Utils_InvokePropertyGet(pdisp, DISPID_VoltageRange, &varg, 1, &varResult);
		if (SUCCEEDED(hr)) VariantToDouble(varResult, maxVoltage);
		pdisp->Release();
	}
}

CAutoTimeConstant::CImpISink::CImpISink(CAutoTimeConstant *pAutoTimeConstant) :
	m_pAutoTimeConstant(pAutoTimeConstant),
	m_cRefs(0),
	m_dispidRequestWorkingDirectory(DISPID_UNKNOWN),
	m_dispidHaveError(DISPID_UNKNOWN),
	m_dispidRequestGetIntegrationTime(DISPID_UNKNOWN),
	m_dispidRequestSetIntegrationTime(DISPID_UNKNOWN),
	m_dispidRequestVoltageRange(DISPID_UNKNOWN)
{
	HRESULT			hr;
	ITypeInfo	*	pTypeInfo;
	TYPEATTR	*	pTypeAttr;

	if (NULL == this->m_pAutoTimeConstant->m_pdisp) return;
	hr = Utils_GetSinkInterfaceID(this->m_pAutoTimeConstant->m_pdisp, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			this->m_pAutoTimeConstant->m_iidSink = pTypeAttr->guid;
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		Utils_GetMemid(pTypeInfo, L"RequestWorkingDirectory", &m_dispidRequestWorkingDirectory);
		Utils_GetMemid(pTypeInfo, L"HaveError", &m_dispidHaveError);
		Utils_GetMemid(pTypeInfo, L"RequestGetIntegrationTime", &m_dispidRequestGetIntegrationTime);
		Utils_GetMemid(pTypeInfo, L"RequestSetIntegrationTime", &m_dispidRequestSetIntegrationTime);
		Utils_GetMemid(pTypeInfo, L"RequestVoltageRange", &m_dispidRequestVoltageRange);
		pTypeInfo->Release();
	}
}

CAutoTimeConstant::CImpISink::~CImpISink()
{

}

// IUnknown methods
STDMETHODIMP CAutoTimeConstant::CImpISink::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IDispatch == riid || riid == this->m_pAutoTimeConstant->m_iidSink)
	{
		*ppv = this;
		this->AddRef();
		return S_OK;
	}
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CAutoTimeConstant::CImpISink::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CAutoTimeConstant::CImpISink::Release()
{
	ULONG		cRefs = --m_cRefs;
	if (0 == cRefs) delete this;
	return cRefs;
}

// IDispatch methods
STDMETHODIMP CAutoTimeConstant::CImpISink::GetTypeInfoCount(
	PUINT			pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CAutoTimeConstant::CImpISink::GetTypeInfo(
	UINT			iTInfo,
	LCID			lcid,
	ITypeInfo	**	ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CAutoTimeConstant::CImpISink::GetIDsOfNames(
	REFIID			riid,
	OLECHAR		**  rgszNames,
	UINT			cNames,
	LCID			lcid,
	DISPID		*	rgDispId)
{
	return E_NOTIMPL;
}

STDMETHODIMP CAutoTimeConstant::CImpISink::Invoke(
	DISPID			dispIdMember,
	REFIID			riid,
	LCID			lcid,
	WORD			wFlags,
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult,
	EXCEPINFO	*	pExcepInfo,
	PUINT			puArgErr)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	if (dispIdMember == this->m_dispidRequestWorkingDirectory)
	{
		if (2 == pDispParams->cArgs && (VT_BYREF | VT_BSTR) == pDispParams->rgvarg[1].vt &&
			(VT_BYREF | VT_BOOL) == pDispParams->rgvarg[0].vt &&
			VARIANT_FALSE == *(pDispParams->rgvarg[0].pboolVal))
		{
			TCHAR			workingDirectory[MAX_PATH];
			if (this->m_pAutoTimeConstant->OnRequestWorkingDirectory(workingDirectory, MAX_PATH))
			{
				*(pDispParams->rgvarg[1].pbstrVal) = SysAllocString(workingDirectory);
				*(pDispParams->rgvarg[0].pboolVal) = VARIANT_TRUE;
			}
		}
	}
	else if (dispIdMember == this->m_dispidHaveError)
	{
		hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
		if (SUCCEEDED(hr))
		{
			this->m_pAutoTimeConstant->OnHaveError(varg.bstrVal);
			VariantClear(&varg);
		}
	}
	else if (dispIdMember == this->m_dispidRequestGetIntegrationTime)
	{
		if (2 == pDispParams->cArgs && (VT_BYREF | VT_I4) == pDispParams->rgvarg[1].vt &&
			(VT_BYREF | VT_BOOL) == pDispParams->rgvarg[0].vt &&
			VARIANT_FALSE == *(pDispParams->rgvarg[0].pboolVal))
		{
			*(pDispParams->rgvarg[0].pboolVal) = this->m_pAutoTimeConstant->OnRequestGetIntegrationTime(
				pDispParams->rgvarg[1].plVal) ? VARIANT_TRUE : VARIANT_FALSE;
		}
	}
	else if (dispIdMember == this->m_dispidRequestSetIntegrationTime)
	{
		if (2 == pDispParams->cArgs && (VT_BYREF | VT_BOOL) == pDispParams->rgvarg[0].vt &&
			VARIANT_FALSE == *(pDispParams->rgvarg[0].pboolVal))
		{
			hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
			if (SUCCEEDED(hr))
			{
				*(pDispParams->rgvarg[0].pboolVal) = this->m_pAutoTimeConstant->OnRequestSetIntegrationTime(
					varg.lVal) ? VARIANT_TRUE : VARIANT_FALSE;
			}
		}
	}
	else if (dispIdMember == this->m_dispidRequestVoltageRange)
	{
		if (2 == pDispParams->cArgs && (VT_BYREF | VT_R8) == pDispParams->rgvarg[1].vt &&
			(VT_BYREF | VT_R8) == pDispParams->rgvarg[0].vt)
		{
			this->m_pAutoTimeConstant->OnRequestVoltageRange(
				pDispParams->rgvarg[1].pdblVal, pDispParams->rgvarg[0].pdblVal);
		}
	}
	return S_OK;
}
