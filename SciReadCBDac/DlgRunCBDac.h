#pragma once

class CMyObject;
class CPlotWindow;

class CDlgRunCBDac
{
public:
	CDlgRunCBDac(CMyObject * pMyObject);
	~CDlgRunCBDac();
	BOOL			DoOpenDialog(
						HINSTANCE		hInst,
						HWND			hwndParent);
	// handler for plot window mouse move
	void			OnMouseMove(
						double			x,
						double			y);
protected:
	BOOL			DlgProc(
						UINT			uMsg,
						WPARAM			wParam,
						LPARAM			lParam);
	BOOL			OnInitDialog();
	BOOL			OnCommand(
						WORD			wmID,
						WORD			wmEvent);
	BOOL			OnNotify(
						LPNMHDR			pnmh);
	BOOL			OnGetDefID();
	// handler for WM_SIZE
	void			OnSize(
						WPARAM			wParam,
						LPARAM			lParam);
	// reading timer
	void			OnReadingTimer();
	// 
	void			DisplayInitialized();
	void			SetInitialized();
	// am integrating flag
	BOOL			GetAmIntegrating();
	void			SetAmIntegrating(
						BOOL			amIntegrating);
	void			OnClickedIntegrationOn();
	// integration time
	void			DisplayIntegrationTime();
	void			SetIntegrationTime();
	// AD range
	void			GetADRange(
						double		*	minVoltage,
						double		*	maxVoltage);
	void			SetADRange();
	void			SetADRange(
						double			minVoltage,
						double			maxVoltage);
	void			DisplayADRange();
	void			OnReturnClickedMinVoltage();
	void			OnIncrementVoltageRange(
						int				iDelta);
	BOOL			GetVoltageRanges(
						long		*	nArray,
						double		**	ppMin,
						double		**	ppMax);
	// want signal integration
	void			DisplayWantIntegrateSignal();
	void			SetWantIntegrateSignal(
						BOOL			fWantIntegrateSignal);
	BOOL			GetWantIntegrateSignal();
	// AD channel
	void			DisplayADChannel();
	void			SetADChannel();
	// remove first half of plot
	void			OnRemoveFirstHalf();
	// display current wavelength
	void			DisplayCurrentWavelength();
	void			SetCurrentWavelength();
	// auto time constant
	void			DisplayAutoTimeConstant();
	BOOL			GetWantAutoTimeConstant();
	void			SetWantAutoTimeConstant(
						BOOL			WantAutoTimeConstant);
	void			SetupAutoTimeConstant();
	void			TestAutoTimeConstant();
	// get our object
	BOOL			GetOurObject(
						IDispatch	**	ppdisp);
	// sink events
	void			OnHaveArrayScan(
						long			nArray,
						unsigned short*	arr);
	void			OnHaveVoltageScan(
						long			nArray,
						double		*	arr);
	// property change notification
	void			OnPropChanged(
						DISPID			dispid);
private:
	CMyObject	*	m_pMyObject;
	HWND			m_hwndDlg;
	BOOL			m_fAllowClose;
	CPlotWindow	*	m_pPlotWindow;
	// bottom right corner for plot window
	long			m_marginRight;
	long			m_marginBottom;
	// sinks
	DWORD			m_dwCookie;
	DWORD			m_dwPropNotify;
	// is the reading timer running?
	BOOL			m_fReadingTimerRunning;

	friend LRESULT CALLBACK DlgProcRunCBDac(HWND, UINT, WPARAM, LPARAM);

	// sink events
	class CImpISink : public IDispatch
	{
	public:
		CImpISink(CDlgRunCBDac *pDlgRunCBDac);
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
		CDlgRunCBDac		*	m_pDlgRunCBDac;
		ULONG					m_cRefs;
	};
	class CImpIPropNotifySink : public IPropertyNotifySink
	{
	public:
		CImpIPropNotifySink(CDlgRunCBDac *pDlgRunCBDac);
		~CImpIPropNotifySink();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IPropertyNotifySink methods
		STDMETHODIMP			OnChanged(
									DISPID			dispID);
		STDMETHODIMP			OnRequestEdit(
									DISPID			dispID);
	private:
		CDlgRunCBDac		*	m_pDlgRunCBDac;
		ULONG					m_cRefs;
	};
	friend CImpISink;
	friend CImpIPropNotifySink;
};

