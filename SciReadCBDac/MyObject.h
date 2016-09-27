#pragma once

class CMyReadCBDac;
class CAutoTimeConstant;

class CMyObject : public IUnknown
{
public:
	CMyObject(IUnknown * pUnkOuter);
	~CMyObject(void);
	// IUnknown methods
	STDMETHODIMP				QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
	STDMETHODIMP_(ULONG)		AddRef();
	STDMETHODIMP_(ULONG)		Release();
	// initializationFireRequestWorkingDirectory
	HRESULT						Init();
	// sink events
	HWND						FireQueryMainWindow();
	BOOL						FireRequestConfigFile(
									LPTSTR			szConfigfile,
									UINT			nBufferSize);
	BOOL						FireRequestIniFile(
									LPTSTR			szIniFile,
									UINT			nBufferSize);
	BOOL						FireRequestCurrentPosition(
									double		*	position);
	BOOL						FireRequestSetPosition(
									double			newPosition);
	void						FireError(
									LPCTSTR			Error);
	BOOL						FireRequestScanningStatus();
	BOOL						FireRequestWorkingDirectory(
									LPTSTR			szWorkingDirectory,
									UINT			nBufferSize);
	void						FireHaveArrayScan(
									unsigned long	nArray,
									unsigned short*	arr);
	void						FireHaveVoltageScan(
									unsigned long	nArray,
									double		*	arr);
	BOOL						FirehaveIntSignal(
									unsigned long	nArray,
									double		*	arr,
									long			nAttempts);
	// property change notification
	void						OnPropChange(
									DISPID			dispid);
//	BOOL						GetAmScanning();
//	// allow change in am scanning
//	void						SetAllowChangeAmScanning(
//									BOOL			fAllow);
	// get the working directory
	BOOL						GetWorkingDirectory(
									LPTSTR			szWorkingDirectory,
									UINT			nBufferSize);
protected:
	HRESULT						GetClassInfo(
									ITypeInfo	**	ppTI);
	HRESULT						GetRefTypeInfo(
									LPCTSTR			szInterface,
									ITypeInfo	**	ppTypeInfo);
	void						ReadConfig(
									LPCTSTR			szConfigFile);
	void						ReadIni(
									LPCTSTR			szIniFile);
	void						WriteIni(
									LPCTSTR			szIniFile);
//	long						GetCurrentGratingIndex();
//	// get/set scanning
//	void						SetAmScanning(
//									BOOL			fAmScanning);
//	// convert count to signal
//	double						ConvertCountsToSignal(
//									long			counts);
//	BOOL						CountsToVolts(
//									long			nArray, 
//									long		*	aCounts, 
//									double		*	aVolts);
//	BOOL						VoltsToCounts(
//									long			nArray, 
//									double		*	aVolts, 
//									long		*	aCounts);
//	// frame data
//	void						ResetFrameData();
//	void						ClearFrameData();
//	BOOL						AddFrameData(
//									long			nData,
//									long		*	pBinnedData);
//	BOOL						FrameAverageComplete();
	// 
//	double						GetCenterWavelength();
//	double						GetNMPerPixel();
//	long						GetPixelOffset();
//	// form array data from binned data
//	void						FormArrayData(
//									long			nData,
//									long		*	pBinnedData);
	// clear the array scan
	void						ClearArrayScan();
private:
	class CImpIDispatch : public IDispatch
	{
	public:
		CImpIDispatch(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIDispatch();
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
	protected:
		HRESULT					GetBoardNumber(
									VARIANT		*	pVarResult);
		HRESULT					SetBoardNumber(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetAmInitialized(
									VARIANT		*	pVarResult);
		HRESULT					SetAmInitialized(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetChannelNumber(
									VARIANT		*	pVarResult);
		HRESULT					SetChannelNumber(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetIntegratingSignal(
									VARIANT		*	pVarResult);
		HRESULT					SetIntegratingSignal(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetIntegrationTime(
									VARIANT		*	pVarResult);
		HRESULT					SetIntegrationTime(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetScanRate(
									VARIANT		*	pVarResult);
		HRESULT					SetScanRate(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetVoltageRange(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					SetVoltageRange(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetADResolution(
									VARIANT		*	pVarResult);
		HRESULT					GetVoltageRanges(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetSingleReading(
									VARIANT		*	pVarResult);
		HRESULT					ConfigDigitalBit(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					WriteDigitalBit(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					ReadDigitalBit(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					OutputAnalog(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					ReadAnalogInput(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					CountsToVoltage(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					VoltageToCounts(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					Setup(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetWantIntegrateSignal(
									VARIANT		*	pVarResult);
		HRESULT					SetWantIntegrateSignal(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetWantAutoTimeConstant(
									VARIANT		*	pVarResult);
		HRESULT					SetWantAutoTimeConstant(
									DISPPARAMS	*	pDispParams);
		HRESULT					SetupAutoTimeConstant(
									DISPPARAMS	*	pDispParams);
		HRESULT					TestAutoTimeConstant(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
		ITypeInfo			*	m_pTypeInfo;
	};
	class CImpIProvideClassInfo2 : public IProvideClassInfo2
	{
	public:
		CImpIProvideClassInfo2(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIProvideClassInfo2();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IProvideClassInfo method
		STDMETHODIMP			GetClassInfo(
									ITypeInfo	**	ppTI);  
		// IProvideClassInfo2 method
		STDMETHODIMP			GetGUID(
									DWORD			dwGuidKind,  //Desired GUID
									GUID		*	pGUID);       
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImpIConnectionPointContainer : public IConnectionPointContainer
	{
	public:
		CImpIConnectionPointContainer(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIConnectionPointContainer();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IConnectionPointContainer methods
		STDMETHODIMP			EnumConnectionPoints(
									IEnumConnectionPoints **ppEnum);
		STDMETHODIMP			FindConnectionPoint(
									REFIID			riid,  //Requested connection point's interface identifier
									IConnectionPoint **ppCP);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImpIPersistStreamInit : public IPersistStreamInit
	{
	public:
		CImpIPersistStreamInit(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIPersistStreamInit();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IPersist method
		STDMETHODIMP			GetClassID(
									CLSID		*	pClassID);
		// IPersistStreamInit methods
		STDMETHODIMP			GetSizeMax(
									ULARGE_INTEGER *pCbSize);
		STDMETHODIMP			InitNew();
		STDMETHODIMP			IsDirty();
		STDMETHODIMP			Load(
									LPSTREAM		pStm);
		STDMETHODIMP			Save(
									LPSTREAM		pStm,
									BOOL			fClearDirty);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImp_IReadAnalog : public IDispatch
	{
	public:
		CImp_IReadAnalog(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImp_IReadAnalog();
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
		HRESULT					Init();
	protected:
		HRESULT					ReadINI(
									DISPPARAMS	*	pDispParams);
		HRESULT					WriteINI(
									DISPPARAMS	*	pDispParams);
		HRESULT					ReadCFG(
									DISPPARAMS	*	pDispParams);
		HRESULT					intInit();				// internal initialization
		HRESULT					AmIOpen(
									VARIANT		*	pVarResult);
		HRESULT					DoSetup();
		HRESULT					GetVoltageRange(
									DISPPARAMS	*	pDispParams);
		HRESULT					SetOutVolts(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					SetOutputLine(
									DISPPARAMS	*	pDispParams);
		HRESULT					SetInputLine(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetMaximumCounts(
									VARIANT		*	pVarResult);
		HRESULT					CheckWarnings(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					SignalType(
									VARIANT		*	pVarResult);
		HRESULT					ReadSignal(
									VARIANT		*	pVarResult);
		HRESULT					GetInfoString(
									VARIANT		*	pVarResult);
		HRESULT					GetAllowSignalAveraging(
									VARIANT		*	pVarResult);
		HRESULT					GetAllowScanAveraging(
									VARIANT		*	pVarResult);
		HRESULT					GetAllowMultipleScans(
									VARIANT		*	pVarResult);
		HRESULT					SetDwellTime(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetSignalIntegrationAvailable(
									VARIANT		*	pVarResult);
		HRESULT					GetIntegrationTime(
									VARIANT		*	pVarResult);
		HRESULT					SetIntegrationTime(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetAutoIntegrationTime(
									VARIANT		*	pVarResult);
		HRESULT					SetAutoIntegrationTime(
									DISPPARAMS	*	pDispParams);
		HRESULT					SetupAutoIntegrationTime(
									DISPPARAMS	*	pDispParams);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
		ITypeInfo			*	m_pTypeInfo;
		DISPID					m_dispidReadINI;
		DISPID					m_dispidWriteINI;
		DISPID					m_dispidReadCFG;
		DISPID					m_dispidInit;
		DISPID					m_dispidAmIOpen;
		DISPID					m_dispidDoSetup;
		DISPID					m_dispidGetVoltageRange;
		DISPID					m_dispidSetOutVolts;
		DISPID					m_dispidSetOutputLine;
		DISPID					m_dispidSetInputLine;
		DISPID					m_dispidGetMaximumCounts;
		DISPID					m_dispidCheckWarnings;
		DISPID					m_dispidSignalType;
		DISPID					m_dispidReadSignal;
		DISPID					m_dispidGetInfoString;
		DISPID					m_dispidAllowSignalAveraging;
		DISPID					m_dispidAllowScanAveraging;
		DISPID					m_dispidAllowMultipleScans;
		DISPID					m_dispidSetDwellTime;
		DISPID					m_dispidSignalIntegrationAvailable;
		DISPID					m_dispidIntegrationTime;
		DISPID					m_dispidAutoIntegrationTime;
		DISPID					m_dispidSetupAutoIntegrationTime;
	};
	friend CImpIDispatch;
	friend CImpIConnectionPointContainer;
	friend CImpIProvideClassInfo2;
	friend CImpIPersistStreamInit;
	friend CImp_IReadAnalog;
	// data members
	CImpIDispatch			*	m_pImpIDispatch;
	CImpIConnectionPointContainer*	m_pImpIConnectionPointContainer;
	CImpIProvideClassInfo2	*	m_pImpIProvideClassInfo2;
	CImpIPersistStreamInit	*	m_pImpIPersistStreamInit;
	CImp_IReadAnalog		*	m_pImp_IReadAnalog;
	// outer unknown for aggregation
	IUnknown				*	m_pUnkOuter;
	// object reference count
	ULONG						m_cRefs;
	// connection points array
	IConnectionPoint		*	m_paConnPts[MAX_CONN_PTS];
	// _IReadAnalog interface
	IID							m_iid_IReadAnalog;
	// flag am scanning
	BOOL						m_fAmScanning;
	// flag allow am scanning change
	BOOL						m_fAllowChangeAmScanning;
	// start scanning flag
	BOOL						m_fStartScanning;
	CMyReadCBDac			*	m_pMyReadCBDac;
	// flag want signal integration
	BOOL						m_fWantIntegrateSignal;
	// auto time constant
//	BOOL						m_fWantAutoTimeConstant;
	CAutoTimeConstant		*	m_pAutoTimeConstant;
	// array scan
	unsigned long				m_nArray;
	double					*	m_pArray;
};
