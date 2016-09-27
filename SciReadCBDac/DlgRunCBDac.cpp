#include "stdafx.h"
#include "DlgRunCBDac.h"
#include "MyObject.h"
#include "dispids.h"
#include "PlotWindow.h"
#include "MyGuids.h"

#define			TMR_READING					0x0001				// reading timer
#define			TMR_CHECKCURRENTPOSITION	0x0002				// monitor the current position

CDlgRunCBDac::CDlgRunCBDac(CMyObject * pMyObject) :
	m_pMyObject(pMyObject),
	m_hwndDlg(NULL),
	m_fAllowClose(TRUE),
	// plot window
	m_pPlotWindow(NULL),
	// sinks
	m_dwCookie(0),
	m_dwPropNotify(0),
	// is the reading timer running?
	m_fReadingTimerRunning(FALSE)
{
}

CDlgRunCBDac::~CDlgRunCBDac()
{
	if (NULL != this->m_pPlotWindow)
		this->m_pPlotWindow->ClosePlotWindow();
	Utils_RELEASE_INTERFACE(this->m_pPlotWindow);
}

BOOL CDlgRunCBDac::DoOpenDialog(
	HINSTANCE		hInst,
	HWND			hwndParent)
{
	return IDOK == DialogBoxParam(hInst, MAKEINTRESOURCE(IDD_DialogRunCBDac), hwndParent, (DLGPROC)DlgProcRunCBDac, (LPARAM) this);
}

LRESULT CALLBACK DlgProcRunCBDac(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CDlgRunCBDac	*	pDlg = NULL;
	if (WM_INITDIALOG == uMsg)
	{
		pDlg = (CDlgRunCBDac*)lParam;
		pDlg->m_hwndDlg = hwndDlg;
		SetWindowLongPtr(hwndDlg, GWLP_USERDATA, lParam);
	}
	else
	{
		pDlg = (CDlgRunCBDac *)GetWindowLongPtr(hwndDlg, GWLP_USERDATA);
	}
	if (NULL != pDlg)
		return pDlg->DlgProc(uMsg, wParam, lParam);
	else
		return FALSE;
}

BOOL CDlgRunCBDac::DlgProc(
	UINT			uMsg,
	WPARAM			wParam,
	LPARAM			lParam)
{
	switch (uMsg)
	{
	case WM_INITDIALOG:
		return this->OnInitDialog();
	case WM_COMMAND:
		return this->OnCommand(LOWORD(wParam), HIWORD(wParam));
	case WM_NOTIFY:
		return this->OnNotify((LPNMHDR)lParam);
	case DM_GETDEFID:
		return this->OnGetDefID();
	case WM_SIZE:
		this->OnSize(wParam, lParam);
		break;
	case WM_TIMER:
		if (TMR_READING == wParam)
		{
			this->OnReadingTimer();
			return TRUE;
		}
		else if (TMR_CHECKCURRENTPOSITION == wParam)
		{
			// display the current wavelength if the edit box doesnt have focus
			if (GetFocus() != GetDlgItem(this->m_hwndDlg, IDC_EDITWAVELENGTH))
			{
				this->DisplayCurrentWavelength();
			}
			return TRUE;
		}
		break;
	default:
		break;
	}
	return FALSE;
}

BOOL CDlgRunCBDac::OnInitDialog()
{
	RECT		rcClient;
	RECT		rcPlot;
	HRESULT		hr;
	CImpIPropNotifySink* pPropNotify;
	IUnknown*	pUnkSink;
	CImpISink*	pSink;

	// connect sinks
	pPropNotify = new CImpIPropNotifySink(this);
	hr = pPropNotify->QueryInterface(IID_IPropertyNotifySink, (LPVOID*)&pUnkSink);
	if (SUCCEEDED(hr))
	{
		Utils_ConnectToConnectionPoint(this->m_pMyObject, pUnkSink, IID_IPropertyNotifySink, TRUE, &this->m_dwPropNotify);
		pUnkSink->Release();
	}
	else delete pPropNotify;
	pSink = new CImpISink(this);
	hr = pSink->QueryInterface(MY_IIDSINK, (LPVOID*)&pUnkSink);
	if (SUCCEEDED(hr))
	{
		Utils_ConnectToConnectionPoint(this->m_pMyObject, pUnkSink, MY_IIDSINK, TRUE, &this->m_dwCookie);
		pUnkSink->Release();
	}
	else delete pSink;

	// plot window
	this->m_pPlotWindow = new CPlotWindow(this);
	this->m_pPlotWindow->AddRef();
	this->m_pPlotWindow->CreatePlotWindow(GetDlgItem(this->m_hwndDlg, IDC_PLOT));
	this->m_pPlotWindow->SetTitles(L"CB Dac Scans", L"Step Number", L"Signal in Volts");
	this->m_pPlotWindow->AddDataSet(L"DAC scan", RGB(0, 0, 255));
	this->m_pPlotWindow->SetCurveVisible(L"DAC scan", TRUE);
	this->m_pPlotWindow->SetYRange(-5.0, 5.0);
	// store the margins for the plot window
	GetClientRect(this->m_hwndDlg, &rcClient);
	GetWindowRect(this->m_pPlotWindow->GetControlWindow(), &rcPlot);
	MapWindowPoints(NULL, this->m_hwndDlg, (LPPOINT)&rcPlot, 2);
	this->m_marginRight = rcClient.right - rcPlot.right;
	this->m_marginBottom = rcClient.bottom - rcPlot.bottom;
	// display values
	this->DisplayInitialized();
	this->DisplayIntegrationTime();
	this->DisplayADRange();
	this->DisplayWantIntegrateSignal();
	this->DisplayADChannel();
	this->DisplayCurrentWavelength();
	// set the timer to monitor the current position
	SetTimer(this->m_hwndDlg, TMR_CHECKCURRENTPOSITION, 500, NULL);
	// auto time constant
	this->DisplayAutoTimeConstant();

	return TRUE;
}

BOOL CDlgRunCBDac::OnCommand(
	WORD			wmID,
	WORD			wmEvent)
{
	switch (wmID)
	{
	case IDOK:
		if (!this->m_fAllowClose)
		{
			this->m_fAllowClose = TRUE;
			return TRUE;
		}
		// fall through
	case IDCANCEL:
		EndDialog(this->m_hwndDlg, wmID);
		// disconnect the sinks
		Utils_ConnectToConnectionPoint(this->m_pMyObject, NULL, MY_IIDSINK, FALSE, &this->m_dwCookie);
		Utils_ConnectToConnectionPoint(this->m_pMyObject, NULL, IID_IPropertyNotifySink, FALSE, &this->m_dwPropNotify);
		return TRUE;
	case IDC_CHKINITIALIZED:
		if (BN_CLICKED == wmEvent)
		{
			this->SetInitialized();
			return TRUE;
		}
		break;
	case IDC_CHKINTEGRATIONON:
		if (BN_CLICKED == wmEvent)
		{
			this->OnClickedIntegrationOn();
			return TRUE;
		}
		break;
	case IDC_CHKSIGNALINTEGRATION:
		if (BN_CLICKED == wmEvent)
		{
			this->SetWantIntegrateSignal(TRUE);
			return TRUE;
		}
		break;
	case IDC_CHKSINGLEREADINGS:
		if (BN_CLICKED == wmEvent)
		{
			this->SetWantIntegrateSignal(FALSE);
			return TRUE;
		}
		break;
	case IDC_REMOVEFIRSTHALF:
		this->OnRemoveFirstHalf();
		return TRUE;
	case IDC_CHKSELECTEDAUTOTIMECONSTANT:
		if (BN_CLICKED == wmEvent)
		{
			// toggle on/off
			BOOL		fWant = this->GetWantAutoTimeConstant();
			this->SetWantAutoTimeConstant(!fWant);
			return TRUE;
		}
		break;
	case IDC_SETUPAUTOTIMECONSTANT:
		this->SetupAutoTimeConstant();
		return TRUE;
	case IDC_TESTAUTOTIMECONSTANT:
		this->TestAutoTimeConstant();
		return TRUE;
	case IDC_EDITINTEGRATIONTIME:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->SetIntegrationTime();
			return TRUE;
		}
		break;
	case IDC_EDITMINVOLTAGE:
	case IDC_EDITMAXVOLTAGE:
	default:
		break;
	}
	return FALSE;
}

BOOL CDlgRunCBDac::OnNotify(
	LPNMHDR			pnmh)
{
	if (UDN_DELTAPOS == pnmh->code && IDC_UPDADRANGE == pnmh->idFrom)
	{
		LPNMUPDOWN		pnmud = (LPNMUPDOWN)pnmh;
		this->OnIncrementVoltageRange(-pnmud->iDelta);
//		this->OnDeltaPosUpdFrameAveraging(pnmud->iDelta);
		return TRUE;
	}
	return FALSE;
}

BOOL CDlgRunCBDac::OnGetDefID()
{
	SHORT		shKeyState;
	HWND		hwndFocus;
	UINT		nID;
	shKeyState = GetKeyState(VK_RETURN);
	if (0 != (0x8000 & shKeyState))
	{
		hwndFocus = GetFocus();
		nID = GetDlgCtrlID(hwndFocus);
		switch (nID)
		{
		case IDC_EDITINTEGRATIONTIME:
			this->m_fAllowClose = FALSE;
			this->SetIntegrationTime();
			return TRUE;
		case IDC_EDITMINVOLTAGE:
			this->m_fAllowClose = FALSE;
			this->OnReturnClickedMinVoltage();
			return TRUE;
		case IDC_EDITMAXVOLTAGE:
			this->m_fAllowClose = FALSE;
			this->SetADRange();
			return TRUE;
		case IDC_EDITCHANNEL:
			this->m_fAllowClose = FALSE;
			this->SetADChannel();
			return TRUE;
		case IDC_EDITWAVELENGTH:
			this->m_fAllowClose = FALSE;
			this->SetCurrentWavelength();
			return TRUE;
		default:
			break;
		}
	}
	return FALSE;
}

// handler for WM_SIZE
void CDlgRunCBDac::OnSize(
	WPARAM			wParam,
	LPARAM			lParam)
{
	RECT			rcClient;
	RECT			rcPlot;
	HWND			hwndPlot;

	if (SIZE_MAXIMIZED == wParam || SIZE_RESTORED == wParam)
	{
		hwndPlot = this->m_pPlotWindow->GetControlWindow();
		GetWindowRect(hwndPlot, &rcPlot);
		// convert to dialog coordinates
		MapWindowPoints(NULL, this->m_hwndDlg, (LPPOINT)&rcPlot, 2);
		GetClientRect(this->m_hwndDlg, &rcClient);
		// determine the new client rectangle for the plot window
		rcPlot.right = rcClient.right - this->m_marginRight;
		rcPlot.bottom = rcClient.bottom - this->m_marginBottom;
		MoveWindow(hwndPlot, rcPlot.left, rcPlot.top, rcPlot.right - rcPlot.left, rcPlot.bottom - rcPlot.top, TRUE);
	}
}

// reading timer
void CDlgRunCBDac::OnReadingTimer()
{
	if (!this->GetWantIntegrateSignal())
	{
		IDispatch		*	pdisp;
		HRESULT				hr;
		VARIANT				varResult;
		double				signal = 0.0;
		long				nArray;
		double			*	pX;
		double				xValue;

		if (this->GetOurObject(&pdisp))
		{
			hr = Utils_InvokeMethod(pdisp, DISPID_GetSignalReading, NULL, 0, &varResult);
			if (SUCCEEDED(hr))VariantToDouble(varResult, &signal);
			pdisp->Release();
		}
		// add the value to the end of the active curve
		if (this->m_pPlotWindow->GetCurveXData(0, &nArray, &pX))
		{
			xValue = pX[nArray - 1] + 1.0;
			delete[] pX;
		}
		else
		{
			xValue = 1.0;
		}
		this->m_pPlotWindow->AddDataValue(0, xValue, signal);
		this->m_pPlotWindow->ZoomOut();
		this->m_pPlotWindow->Refresh();
	}
}

// 
void CDlgRunCBDac::DisplayInitialized()
{
	IDispatch	*	pdisp;
	if (this->GetOurObject(&pdisp))
	{
		Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKINITIALIZED), 
			Utils_GetBoolProperty(pdisp, DISPID_AmInitialized) ? BST_CHECKED : BST_UNCHECKED);
		pdisp->Release();
	}
}

void CDlgRunCBDac::SetInitialized()
{
	IDispatch	*	pdisp;
	if (this->GetOurObject(&pdisp))
	{
		Utils_SetBoolProperty(pdisp, DISPID_AmInitialized, TRUE);
		pdisp->Release();
	}
}

// am integrating flag
BOOL CDlgRunCBDac::GetAmIntegrating()
{
	return BST_CHECKED == Button_GetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKINTEGRATIONON));
}

void CDlgRunCBDac::SetAmIntegrating(
	BOOL			amIntegrating)
{
	Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKINTEGRATIONON),
		amIntegrating ? BST_CHECKED : BST_UNCHECKED);
}

void CDlgRunCBDac::OnClickedIntegrationOn()
{
	// toggle on/off
	BOOL		AmIntegrating = this->GetAmIntegrating();
	this->SetAmIntegrating(!AmIntegrating);
	if (this->GetAmIntegrating())
	{
		// start the reading timer
		SetTimer(this->m_hwndDlg, TMR_READING, 250, NULL);
		this->m_fReadingTimerRunning = TRUE;
		if (this->GetWantIntegrateSignal())
		{
			IDispatch		*	pdisp;
			if (this->GetOurObject(&pdisp))
			{
				Utils_SetBoolProperty(pdisp, DISPID_IntegratingSignal, TRUE);
				pdisp->Release();
			}
		}
	}
	else
	{
		// stop the reading timer
		KillTimer(this->m_hwndDlg, TMR_READING);
		this->m_fReadingTimerRunning = FALSE;
	}
}

// integration time
void CDlgRunCBDac::DisplayIntegrationTime()
{
	IDispatch	*	pdisp;
	long			integrationTime = 1000;
	if (this->GetOurObject(&pdisp))
	{
		integrationTime = Utils_GetLongProperty(pdisp, DISPID_IntegrationTime);
		pdisp->Release();
	}
	SetDlgItemInt(this->m_hwndDlg, IDC_EDITINTEGRATIONTIME, integrationTime, FALSE);
}
void CDlgRunCBDac::SetIntegrationTime()
{
	long		integrationTime = GetDlgItemInt(this->m_hwndDlg, IDC_EDITINTEGRATIONTIME, NULL, FALSE);
	IDispatch*	pdisp;
	if (this->GetOurObject(&pdisp))
	{
		Utils_SetLongProperty(pdisp, DISPID_IntegrationTime, integrationTime);
		pdisp->Release();
	}
}

// AD range
void CDlgRunCBDac::GetADRange(
	double		*	minVoltage,
	double		*	maxVoltage)
{
	IDispatch	*	pdisp;
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	*minVoltage = 0.0;
	*maxVoltage = 0.0;
	if (this->GetOurObject(&pdisp))
	{
		VariantInit(&varg);
		varg.vt = VT_BYREF | VT_R8;
		varg.pdblVal = minVoltage;
		hr = Utils_InvokePropertyGet(pdisp, DISPID_VoltageRange, &varg, 1, &varResult);
		if (SUCCEEDED(hr)) VariantToDouble(varResult, maxVoltage);
		pdisp->Release();
	}
}

void CDlgRunCBDac::SetADRange()
{
	TCHAR			szString[MAX_PATH];
	float			fval;
	double			minVoltage = 0.0;
	double			maxVoltage = 0.0;

	if (GetDlgItemText(this->m_hwndDlg, IDC_EDITMINVOLTAGE, szString, MAX_PATH) > 0)
	{
		if (1 == _stscanf_s(szString, L"%f", &fval))
		{
			minVoltage = fval;
		}
	}
	if (GetDlgItemText(this->m_hwndDlg, IDC_EDITMAXVOLTAGE, szString, MAX_PATH) > 0)
	{
		if (1 == _stscanf_s(szString, L"%f", &fval))
		{
			maxVoltage = fval;
		}
	}
	this->SetADRange(minVoltage, maxVoltage);
}

void CDlgRunCBDac::SetADRange(
	double			minVoltage,
	double			maxVoltage)
{
	IDispatch	*	pdisp;
	VARIANTARG		avarg[2];
	if (this->GetOurObject(&pdisp))
	{
		VariantInit(&avarg[1]);
		avarg[1].vt = VT_BYREF | VT_R8;
		avarg[1].pdblVal = &minVoltage;
		InitVariantFromDouble(maxVoltage, &avarg[0]);
		Utils_InvokePropertyPut(pdisp, DISPID_VoltageRange, avarg, 2);
		pdisp->Release();
	}
}

void CDlgRunCBDac::DisplayADRange()
{
	double		minVoltage;
	double		maxVoltage;
	TCHAR		szString[MAX_PATH];
	this->GetADRange(&minVoltage, &maxVoltage);
	_stprintf_s(szString, L"%5.1f", minVoltage);
	SetDlgItemText(this->m_hwndDlg, IDC_EDITMINVOLTAGE, szString);
	_stprintf_s(szString, L"%4.1f", maxVoltage);
	SetDlgItemText(this->m_hwndDlg, IDC_EDITMAXVOLTAGE, szString);
}

void CDlgRunCBDac::OnReturnClickedMinVoltage()
{
	double			minVoltage = 0.0;
	float			fval;
	TCHAR			szString[MAX_PATH];
	if (GetDlgItemText(this->m_hwndDlg, IDC_EDITMINVOLTAGE, szString, MAX_PATH) > 0)
	{
		if (1 == _stscanf_s(szString, L"%f", &fval))
		{
			minVoltage = abs(1.0 * fval);
			_stprintf_s(szString, L"%4.1f", (float) minVoltage);
			SetDlgItemText(this->m_hwndDlg, IDC_EDITMAXVOLTAGE, szString);
			SetFocus(GetDlgItem(this->m_hwndDlg, IDC_EDITMAXVOLTAGE));
		}
	}
}

void CDlgRunCBDac::OnIncrementVoltageRange(
	int				iDelta)
{
	long			index	= 0;			// current index of the voltage range
	long			nArray;
	double		*	pMin = NULL;
	double		*	pMax = NULL;
	double			minValue = 0.0;
	double			maxValue = 0.0;
	double			tolerance;
	BOOL			fDone = FALSE;

	// get the available ranges
	if (this->GetVoltageRanges(&nArray, &pMin, &pMax))
	{
		// get the current range
		this->GetADRange(&minValue, &maxValue);
		// find the current index in the array
		index = 0;
		fDone = FALSE;
		while (index < nArray && !fDone)
		{
			tolerance = pMax[index] / 100.0;
			if (maxValue >= (pMax[index] - tolerance) && maxValue <= (pMax[index] + tolerance))
			{
				fDone = TRUE;
			}
			if (!fDone) index++;
		}
		if (!fDone) index = 0;
		index = iDelta > 0 ? index + 1 : index - 1;
		if (index < 0) index = 0;
		if (index >= nArray) index = nArray - 1;
		this->SetADRange(pMin[index], pMax[index]);
		// cleanup
		delete[] pMin;
		delete[] pMax;
	}
}

BOOL CDlgRunCBDac::GetVoltageRanges(
	long		*	nArray,
	double		**	ppMin,
	double		**	ppMax)
{
	HRESULT				hr;
	IDispatch		*	pdisp;
	VARIANTARG			varg;
	VARIANT				varResult;
	SAFEARRAY		*	pSAMins	= NULL;
	long				ubMins;
	long				ubMaxs;
	long				i;
	BOOL				fSuccess = FALSE;
	double			*	arrMins;
	double			*	arrMaxs;
	*nArray = 0;
	*ppMin = NULL;
	*ppMax = NULL;
	if (this->GetOurObject(&pdisp))
	{
		VariantInit(&varg);
		varg.vt = VT_BYREF | VT_ARRAY | VT_R8;
		varg.pparray = &pSAMins;
		hr = Utils_InvokePropertyGet(pdisp, DISPID_VoltageRanges, &varg, 1, &varResult);
		if (SUCCEEDED(hr))
		{
			if ((VT_ARRAY | VT_R8) == varResult.vt && NULL != varResult.parray)
			{
				if (NULL != pSAMins)
				{
					// check array sizes
					SafeArrayGetUBound(pSAMins, 1, &ubMins);
					SafeArrayGetUBound(varResult.parray, 1, &ubMaxs);
					if (ubMins == ubMaxs)
					{
						fSuccess = TRUE;
						*nArray = ubMins + 1;
						*ppMin = new double[*nArray];
						*ppMax = new double[*nArray];
						SafeArrayAccessData(pSAMins, (void**)&arrMins);
						SafeArrayAccessData(varResult.parray, (void**)&arrMaxs);
						for (i = 0; i < (*nArray); i++)
						{
							(*ppMin)[i] = arrMins[i];
							(*ppMax)[i] = arrMaxs[i];
						}
						SafeArrayUnaccessData(pSAMins);
						SafeArrayUnaccessData(varResult.parray);
					}
				}
			}
			VariantClear(&varResult);
		}
		if (NULL != pSAMins) SafeArrayDestroy(pSAMins);
		pdisp->Release();
	}
	return fSuccess;
}

// want signal integration
void CDlgRunCBDac::DisplayWantIntegrateSignal()
{
	if (this->GetWantIntegrateSignal())
	{
		// integrate signal
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_EDITINTEGRATIONTIME), TRUE);
		Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKSIGNALINTEGRATION), BST_CHECKED);
		Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKSINGLEREADINGS), BST_UNCHECKED);
	}
	else
	{
		// dont integrate signal
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_EDITINTEGRATIONTIME), FALSE);
		Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKSIGNALINTEGRATION), BST_UNCHECKED);
		Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKSINGLEREADINGS), BST_CHECKED);
	}
}

void CDlgRunCBDac::SetWantIntegrateSignal(
	BOOL			fWantIntegrateSignal)
{
	IDispatch	*	pdisp;
	if (this->GetOurObject(&pdisp))
	{
		Utils_SetBoolProperty(pdisp, DISPID_WantIntegrateSignal, fWantIntegrateSignal);
		pdisp->Release();
	}
}

BOOL CDlgRunCBDac::GetWantIntegrateSignal()
{
	IDispatch	*	pdisp;
	BOOL			WantIntegrateSignal = FALSE;
	if (this->GetOurObject(&pdisp))
	{
		WantIntegrateSignal = Utils_GetBoolProperty(pdisp, DISPID_WantIntegrateSignal);
		pdisp->Release();
	}
	return WantIntegrateSignal;
}

// AD channel
void CDlgRunCBDac::DisplayADChannel()
{
	IDispatch	*	pdisp;
	if (this->GetOurObject(&pdisp))
	{
		SetDlgItemInt(this->m_hwndDlg, IDC_EDITCHANNEL, Utils_GetLongProperty(pdisp, DISPID_ChannelNumber), TRUE);
		pdisp->Release();
	}
}

void CDlgRunCBDac::SetADChannel()
{
	long			channel = GetDlgItemInt(this->m_hwndDlg, IDC_EDITCHANNEL, NULL, TRUE);
	IDispatch	*	pdisp;
	if (this->GetOurObject(&pdisp))
	{
		Utils_SetLongProperty(pdisp, DISPID_ChannelNumber, channel);
		pdisp->Release();
	}
}

// remove first half of plot
void CDlgRunCBDac::OnRemoveFirstHalf()
{
	long			nX;
	double		*	pX;
	long			nY;
	double		*	pY;
	long			nRemove;
	long			i;
	double		*	tX;
	double		*	tY;
	long			tN;

	if (this->m_pPlotWindow->GetCurveXData(0, &nX, &pX))
	{
		if (this->m_pPlotWindow->GetCurveYData(0, &nY, &pY))
		{
			if (nX == nY)
			{
				nRemove = nX / 2;
				tN = nX - nRemove;
				tX = new double[tN];
				tY = new double[tN];
				for (i = 0; i < tN; i++)
				{
					tX[i] = (i + 1) * 1.0;
					tY[i] = pY[i + nRemove];
				}
				this->m_pPlotWindow->SetCurveXData(0, tN, tX);
				this->m_pPlotWindow->SetCurveYData(0, tN, tY);
				delete[] tX;
				delete[] tY;
				this->m_pPlotWindow->ZoomOut();
				this->m_pPlotWindow->Refresh();
			}
			delete[] pY;
		}
		delete[] pX;
	}
}

// display current wavelength
void CDlgRunCBDac::DisplayCurrentWavelength()
{
	double		position;
	TCHAR		szString[MAX_PATH];

	if (this->m_pMyObject->FireRequestCurrentPosition(&position))
	{
		_stprintf_s(szString, L"%5.1f", position);
		SetDlgItemText(this->m_hwndDlg, IDC_EDITWAVELENGTH, szString);
	}
	else
	{
		SetDlgItemText(this->m_hwndDlg, IDC_EDITWAVELENGTH, L"");
	}
}

void CDlgRunCBDac::SetCurrentWavelength()
{
	TCHAR		szString[MAX_PATH];
	float		fval;

	if (GetDlgItemText(this->m_hwndDlg, IDC_EDITWAVELENGTH, szString, MAX_PATH) > 0)
	{
		if (1 == _stscanf_s(szString, L"%f", &fval))
		{
			this->m_pMyObject->FireRequestSetPosition(fval);
		}
	}
}

// auto time constant
void CDlgRunCBDac::DisplayAutoTimeConstant()
{
	Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKSELECTEDAUTOTIMECONSTANT),
		this->GetWantAutoTimeConstant() ? BST_CHECKED : BST_UNCHECKED);
}

BOOL CDlgRunCBDac::GetWantAutoTimeConstant()
{
	IDispatch	*	pdisp;
	BOOL			fWant = FALSE;
	if (this->GetOurObject(&pdisp))
	{
		fWant = Utils_GetBoolProperty(pdisp, DISPID_WantAutoTimeConstant);
		pdisp->Release();
	}
	return fWant;
}

void CDlgRunCBDac::SetWantAutoTimeConstant(
	BOOL			WantAutoTimeConstant)
{
	IDispatch	*	pdisp;
	if (this->GetOurObject(&pdisp))
	{
		Utils_SetBoolProperty(pdisp, DISPID_WantAutoTimeConstant, WantAutoTimeConstant);
		pdisp->Release();
	}
}

void CDlgRunCBDac::SetupAutoTimeConstant()
{
	IDispatch	*	pdisp;
	VARIANTARG		varg;
	if (this->GetOurObject(&pdisp))
	{
		InitVariantFromInt32((long) this->m_hwndDlg, &varg);
		Utils_InvokeMethod(pdisp, DISPID_SetupAutoTimeConstant, &varg, 1, NULL);
		pdisp->Release();
	}
}

void CDlgRunCBDac::TestAutoTimeConstant()
{
	IDispatch	*	pdisp;
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	TCHAR			szString[MAX_PATH];
	if (this->GetOurObject(&pdisp))
	{
		InitVariantFromInt32((long) this->m_hwndDlg, &varg);
		hr = Utils_InvokeMethod(pdisp, DISPID_TestAutoTimeConstant, &varg, 1, &varResult);
		if (SUCCEEDED(hr))
		{
			VariantToString(varResult, szString, MAX_PATH);
			MessageBox(this->m_hwndDlg, szString, L"Auto Time Constant Test", MB_OK);
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
}


// handler for plot window mouse move
void CDlgRunCBDac::OnMouseMove(
	double			x,
	double			y)
{
	TCHAR			szString[MAX_PATH];
	_stprintf_s(szString, L"x = %4.0f  y = %7.4g", x, y);
	SetDlgItemText(this->m_hwndDlg, IDC_STATUS, szString);
}

// get our object
BOOL CDlgRunCBDac::GetOurObject(
	IDispatch	**	ppdisp)
{
	HRESULT		hr;
	hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*)ppdisp);
	return SUCCEEDED(hr);
}

// sink events
void CDlgRunCBDac::OnHaveArrayScan(
	long			nArray,
	unsigned short*	arr)
{

}

void CDlgRunCBDac::OnHaveVoltageScan(
	long			nArray,
	double		*	arr)
{
	if (NULL != this->m_pPlotWindow)
	{
		double	*	pX = new double[nArray];
		long		i;
		for (i = 0; i < nArray; i++)
			pX[i] = (i + 1) / 1000.0;
		this->m_pPlotWindow->SetCurveData(L"DAC scan", nArray, pX, arr);
		delete[] pX;
		this->m_pPlotWindow->Refresh();
	}
}

// property change notification
void CDlgRunCBDac::OnPropChanged(
	DISPID			dispid)
{
	if (DISPID_AmInitialized == dispid)
	{
		this->DisplayInitialized();
	}
	else if (DISPID_IntegratingSignal == dispid)
	{
		if (this->GetAmIntegrating())
		{
			IDispatch	*	pdisp;
			if (this->GetOurObject(&pdisp))
			{
				if (!Utils_GetBoolProperty(pdisp, DISPID_IntegratingSignal))
				{
					Utils_SetBoolProperty(pdisp, DISPID_IntegratingSignal, TRUE);
				}
				pdisp->Release();
			}
		}
	}
	else if (DISPID_IntegrationTime == dispid)
	{
		this->DisplayIntegrationTime();
	}
	else if (DISPID_VoltageRange == dispid)
	{
		this->DisplayADRange();
	}
	else if (DISPID_WantIntegrateSignal == dispid)
	{
		this->DisplayWantIntegrateSignal();
	}
	else if (DISPID_ChannelNumber == dispid)
	{
		this->DisplayADChannel();
	}
	else if (DISPID_WantAutoTimeConstant == dispid)
	{
		this->DisplayAutoTimeConstant();
	}
}

CDlgRunCBDac::CImpISink::CImpISink(CDlgRunCBDac *pDlgRunCBDac) :
	m_pDlgRunCBDac(pDlgRunCBDac),
	m_cRefs(0)
{
}

CDlgRunCBDac::CImpISink::~CImpISink()
{
}

// IUnknown methods
STDMETHODIMP CDlgRunCBDac::CImpISink::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IDispatch == riid || MY_IIDSINK == riid)
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

STDMETHODIMP_(ULONG) CDlgRunCBDac::CImpISink::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CDlgRunCBDac::CImpISink::Release()
{
	ULONG		cRefs = --m_cRefs;
	if (0 == cRefs) delete this;
	return cRefs;
}

// IDispatch methods
STDMETHODIMP CDlgRunCBDac::CImpISink::GetTypeInfoCount(
	PUINT			pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CDlgRunCBDac::CImpISink::GetTypeInfo(
	UINT			iTInfo,
	LCID			lcid,
	ITypeInfo	**	ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CDlgRunCBDac::CImpISink::GetIDsOfNames(
	REFIID			riid,
	OLECHAR		**  rgszNames,
	UINT			cNames,
	LCID			lcid,
	DISPID		*	rgDispId)
{
	return E_NOTIMPL;
}

STDMETHODIMP CDlgRunCBDac::CImpISink::Invoke(
	DISPID			dispIdMember,
	REFIID			riid,
	LCID			lcid,
	WORD			wFlags,
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult,
	EXCEPINFO	*	pExcepInfo,
	PUINT			puArgErr)
{
	if (DISPID_HaveArrayScan == dispIdMember)
	{
		if (1 == pDispParams->cArgs && (VT_ARRAY | VT_UI2) == pDispParams->rgvarg[0].vt)
		{
			long		ub;
			unsigned short*	arr;
			SafeArrayGetUBound(pDispParams->rgvarg[0].parray, 1, &ub);
			SafeArrayAccessData(pDispParams->rgvarg[0].parray, (void**)&arr);
			this->m_pDlgRunCBDac->OnHaveArrayScan(ub + 1, arr);
			SafeArrayUnaccessData(pDispParams->rgvarg[0].parray);
		}
	}
	else if (DISPID_HaveVoltageScan == dispIdMember)
	{
		if (1 == pDispParams->cArgs && (VT_ARRAY | VT_R8) == pDispParams->rgvarg[0].vt)
		{
			long		ub;
			double	*	arr;
			SafeArrayGetUBound(pDispParams->rgvarg[0].parray, 1, &ub);
			SafeArrayAccessData(pDispParams->rgvarg[0].parray, (void**)&arr);
			this->m_pDlgRunCBDac->OnHaveVoltageScan(ub + 1, arr);
			SafeArrayUnaccessData(pDispParams->rgvarg[0].parray);
		}
	}
	return S_OK;
}

CDlgRunCBDac::CImpIPropNotifySink::CImpIPropNotifySink(CDlgRunCBDac *pDlgRunCBDac) :
	m_pDlgRunCBDac(pDlgRunCBDac),
	m_cRefs(0)
{

}

CDlgRunCBDac::CImpIPropNotifySink::~CImpIPropNotifySink()
{

}

// IUnknown methods
STDMETHODIMP CDlgRunCBDac::CImpIPropNotifySink::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IPropertyNotifySink == riid)
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

STDMETHODIMP_(ULONG) CDlgRunCBDac::CImpIPropNotifySink::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CDlgRunCBDac::CImpIPropNotifySink::Release()
{
	ULONG		cRefs = --m_cRefs;
	if (0 == cRefs) delete this;
	return cRefs;
}

// IPropertyNotifySink methods
STDMETHODIMP CDlgRunCBDac::CImpIPropNotifySink::OnChanged(
	DISPID			dispID)
{
	this->m_pDlgRunCBDac->OnPropChanged(dispID);
	return S_OK;
}

STDMETHODIMP CDlgRunCBDac::CImpIPropNotifySink::OnRequestEdit(
	DISPID			dispID)
{
	return S_OK;
}