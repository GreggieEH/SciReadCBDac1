#pragma once

class CMyObject;

class CAutoTimeConstant
{
public:
	CAutoTimeConstant(CMyObject * pMyObject);
	~CAutoTimeConstant();
	BOOL			GetScriptLoaded();
	void			SetScriptLoaded(
						BOOL		ScriptLoaded);
	void			GetScriptFile(
						LPTSTR		szScriptFile,
						UINT		nBufferSize);
	void			SetScriptFile(
						LPCTSTR		szScriptFile);
	BOOL			DisplaySetup(
						HWND		hwndParent);
	BOOL			HaveIntSignal(
						ULONG		nArray,
						double	*	arr,
						short int	numberOfAttempts);
protected:
	// sink events
	BOOL			OnRequestWorkingDirectory(
						LPTSTR		workingDirectory,
						UINT		nBufferSize);
	void			OnHaveError(
						LPCTSTR		Error);
	BOOL			OnRequestGetIntegrationTime(
						long	*	integrationTime);
	BOOL			OnRequestSetIntegrationTime(
						long		integrationTime);
	void			OnRequestVoltageRange(
						double	*	minVoltage,
						double	*	maxVoltage);
private:
	CMyObject	*	m_pMyObject;
	IDispatch	*	m_pdisp;
	IID				m_iidSink;
	DWORD			m_dwCookie;

	// sink events
	class CImpISink : public IDispatch
	{
	public:
		CImpISink(CAutoTimeConstant *pAutoTimeConstant);
		~CImpISink();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IDispatch methods
		STDMETHODIMP			GetTypeInfoCount(
									PUINT			pctinfo);
		STDMETHODIMP			GetTypeInfo(
									UINT			iTInfo,
									LCID			lcid,
									ITypeInfo	**	ppTInfo);
		STDMETHODIMP			GetIDsOfNames(
									REFIID			riid,
									OLECHAR		**  rgszNames,
									UINT			cNames,
									LCID			lcid,
									DISPID		*	rgDispId);
		STDMETHODIMP			Invoke(
									DISPID			dispIdMember,
									REFIID			riid,
									LCID			lcid,
									WORD			wFlags,
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult,
									EXCEPINFO	*	pExcepInfo,
									PUINT			puArgErr);
	private:
		CAutoTimeConstant	*	m_pAutoTimeConstant;
		ULONG					m_cRefs;
		DISPID					m_dispidRequestWorkingDirectory;
		DISPID					m_dispidHaveError;
		DISPID					m_dispidRequestGetIntegrationTime;
		DISPID					m_dispidRequestSetIntegrationTime;
		DISPID					m_dispidRequestVoltageRange;
	};
	friend CImpISink;
};

