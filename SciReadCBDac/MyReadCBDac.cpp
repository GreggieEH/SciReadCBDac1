#include "stdafx.h"
#include "MyReadCBDac.h"
#include "MyObject.h"
#include "G:\Users\Public\Documents\Measurement Computing\DAQ\CWin\cbw.h"
#include "Server.h"
#include "dispids.h"

// private windows messages
#define			WM_CHECKADPROCESS			WM_APP + 0x0103
#define			WM_THREADCLOSED				WM_APP + 0x0104

CMyReadCBDac::CMyReadCBDac(CMyObject * pMyObject) :
	m_pMyObject(pMyObject),
	// properties
	m_boardNumber(0),
	m_SystemInitialized(FALSE),
	m_ChannelNumber(0),
	m_ADRange(BIP5VOLTS),
	m_ScanRate(10000),
	m_integrationTime(1000),			// integration time in ms
	m_fIntegratingSignal(FALSE),
	m_hwndThreadMonitor(NULL),
	m_fStopIntegration(FALSE),			// flag stop integration
	m_hMemory(NULL),					// windows memory handle
	m_nPoints(0),						// number of points in the memory
	m_hThread(NULL),					// thread handle
	// signal average
	m_signalAverage(0.0)
{
	// critical section object
	InitializeCriticalSection(&(this->m_CriticalSection));
}

CMyReadCBDac::~CMyReadCBDac()
{
	// make sure that the thread is closed
	if (this->GetIntegratingSignal())
	{
		this->SetIntegratingSignal(FALSE);
	}
	// delete the critical section
	DeleteCriticalSection(&(this->m_CriticalSection));
}

long CMyReadCBDac::GetBoardNumber()
{
	return this->m_boardNumber;
}

void CMyReadCBDac::SetBoardNumber(
	long		boardNumber)
{
	this->m_boardNumber = boardNumber;
}

BOOL CMyReadCBDac::GetAmInitialized()
{
	return this->m_SystemInitialized;
}

void CMyReadCBDac::SetAmInitialized(
	BOOL		fAmInitialized)
{
	if (fAmInitialized)
	{
		if (!this->m_SystemInitialized)
		{
			float	RevNum = (float) CURRENTREVNUM;
			int		ulStat = cbDeclareRevision(&RevNum);
			if (0 == ulStat)
			{
				cbErrHandling(DONTPRINT, DONTSTOP);
				this->m_SystemInitialized = TRUE;
				Utils_OnPropChanged(this->m_pMyObject, DISPID_AmInitialized);
			}
			else
			{
				this->m_pMyObject->FireError(L"Failed to initialize Measurement Computing");
			}
		}
	}
}

long CMyReadCBDac::GetChannelNumber()
{
	return this->m_ChannelNumber;
}

void CMyReadCBDac::SetChannelNumber(
	long		channelNumber)
{
	this->m_ChannelNumber = channelNumber;
}

BOOL CMyReadCBDac::GetIntegratingSignal()
{
	return this->m_fIntegratingSignal;
}

void CMyReadCBDac::SetIntegratingSignal(
	BOOL		integratingSignal)
{
	int			ulStat;

	if (integratingSignal)
	{
		if (!this->m_fIntegratingSignal)
		{
			int			options = BACKGROUND;
			// allocate space
			this->m_nPoints = (this->m_integrationTime * this->m_ScanRate) / 1000;
			this->m_hMemory = cbWinBufAlloc(this->m_nPoints);
			ulStat = cbAInScan(this->m_boardNumber, this->m_ChannelNumber, this->m_ChannelNumber, this->m_nPoints, &this->m_ScanRate,
				this->m_ADRange, this->m_hMemory, options);
			if (0 == ulStat)
			{
				this->StartThreadMonitor();
				this->m_fIntegratingSignal = this->GetThreadRunning();
				Utils_OnPropChanged(this->m_pMyObject, DISPID_IntegratingSignal);

			}
		}
	}
	else
	{
		if (this->m_fIntegratingSignal)
		{
			this->SetStopThread(TRUE);
		}
	}
}

// integrate signal
double CMyReadCBDac::IntegrateSignal()
{
	this->SetIntegratingSignal(TRUE);
	BOOL		fDone = FALSE;
	while (!fDone)
	{
		MyYield();
		fDone = !this->GetIntegratingSignal();
	}
	// at this point, the signal average should have been updated
	return this->m_signalAverage;
}

void CMyReadCBDac::SetSignalAverage(
	double			signalAverage)
{
	this->m_signalAverage = signalAverage;
}


// integration time in ms
long CMyReadCBDac::GetIntegrationTime()
{
	return this->m_integrationTime;
}

void CMyReadCBDac::SetIntegrationTime(
	long		integrationTime)
{
	this->m_integrationTime = integrationTime;
}

// scan rate data per second
long CMyReadCBDac::GetScanRate()
{
	return this->m_ScanRate;
}

void CMyReadCBDac::SetScanRate(
	long		scanRate)
{
	this->m_ScanRate = scanRate;
}

// last signal reading
ULONG CMyReadCBDac::GetSignalReading()
{
	unsigned short	signal = 0;
	this->ReadAnalogInput(this->m_ChannelNumber, &signal);
	return signal;
}

// get scan array
BOOL CMyReadCBDac::GetScanArray(
	long	*	nValues,
	ULONG	**	ppScanData)
{
	*nValues = 0;
	*ppScanData = NULL;
	return FALSE;
}

// voltage range
void CMyReadCBDac::GetVoltageRange(
	double	*	minVoltage,
	double	*	maxVoltage)
{
	if (BIP1VOLTS == this->m_ADRange)
	{
		*minVoltage = -1.0;
		*maxVoltage = 1.0;
	}
	else if (BIP2VOLTS == this->m_ADRange)
	{
		*minVoltage = -2.0;
		*maxVoltage = 2.0;
	}
	else if (BIP5VOLTS == this->m_ADRange)
	{
		*minVoltage = -5.0;
		*maxVoltage = 5.0;
	}
	else if (BIP10VOLTS == this->m_ADRange)
	{
		*minVoltage = -10.0;
		*maxVoltage = 10.0;
	}
}

void CMyReadCBDac::SetVoltageRange(
	double		minVoltage,
	double		maxVoltage)
{
	if (maxVoltage >= 0.99 && maxVoltage <= 1.01)
	{
		this->m_ADRange = BIP1VOLTS;
	}
	else if (maxVoltage >= 1.98 && maxVoltage <= 2.02)
	{
		this->m_ADRange = BIP2VOLTS;
	}
	else if (maxVoltage > 4.95 && maxVoltage <= 5.05)
	{
		this->m_ADRange = BIP5VOLTS;
	}
	else if (maxVoltage >= 9.90 && maxVoltage <= 10.10)
	{
		this->m_ADRange = BIP10VOLTS;
	}
	else
	{
		this->m_pMyObject->FireError(L"Desired max voltage out of range");
	}
}

// AD resolution
long CMyReadCBDac::GetADResolution()
{
	// for USB-1608FS Plus
	return 16;
}

// configure a digital bit for input or output
BOOL CMyReadCBDac::ConfigDigitalBit(
	short		bitNumber,
	BOOL		forOutput)
{
	int		ulStat = cbDConfigBit(this->m_boardNumber, AUXPORT, bitNumber, forOutput ? DIGITALOUT : DIGITALIN);
	return 0 == ulStat;
}

BOOL CMyReadCBDac::WriteDigitalBit(
	short		bitNumber,
	BOOL		fHi)
{
	int			ulStat = cbDBitOut(this->m_boardNumber, AUXPORT, bitNumber, fHi ? 1 : 0);
	return 0 == ulStat;
}

BOOL CMyReadCBDac::ReadDigitalBit(
	short		bitNumber,
	BOOL	*	fHi)
{
	int				ulStat;
	unsigned short	bitValue;
	*fHi = FALSE;
	ulStat = cbDBitIn(this->m_boardNumber, AUXPORT, bitNumber, &bitValue);
	if (0 == ulStat)
	{
		*fHi = 1 == bitValue;
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

// output analog signal
BOOL CMyReadCBDac::OutputAnalog(
	int			channel,
	double		value)
{
	if (!this->GetIntegratingSignal())
	{
		unsigned short		dataValue;
		int		ulStat;
		
		ulStat = cbFromEngUnits(this->m_boardNumber, this->m_ADRange, (float)value, &dataValue);
		if (0 == ulStat) ulStat = cbAOut(this->m_boardNumber, channel, this->m_ADRange, dataValue);
		return 0 == ulStat;
	}
	else
	{
		return FALSE;
	}
}

BOOL CMyReadCBDac::ReadAnalogInput(
	int			channel,
	unsigned short*	dataValue)
{
	*dataValue = 0;
	if (!this->GetIntegratingSignal())
	{
		int		ulStat = cbAIn(this->m_boardNumber, channel, this->m_ADRange, dataValue);
		return 0 == ulStat;
	}
	else
	{
		return FALSE;
	}
}

// convert to engineering units
BOOL CMyReadCBDac::CountsToVoltage(
	unsigned short	dataValue,
	double		*	voltage)
{
	*voltage = 0.0;
	float		fval;
	int		ulStat = cbToEngUnits(this->m_boardNumber, this->m_ADRange, dataValue, &fval);
	if (0 == ulStat)
	{
		*voltage = fval;
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

BOOL CMyReadCBDac::VoltageToCounts(
	double		voltage,
	unsigned short*	dataValue)
{
	*dataValue = 0;
	int  ulStat = cbFromEngUnits(this->m_boardNumber, this->m_ADRange, (float)voltage, dataValue);
	return 0 == ulStat;
}

// available voltage ranges  - set for USB-1608FS Plus 
BOOL CMyReadCBDac::GetVoltageRanges(
	long		*	nRanges,
	double		**	ppMinVoltage,
	double		**	ppMaxVoltage)
{
	*nRanges = 4;
	*ppMinVoltage = new double[*nRanges];
	*ppMaxVoltage = new double[*nRanges];
	(*ppMinVoltage)[0] = -1.0;			// BIP1VOLTS
	(*ppMaxVoltage)[0] = 1.0;
	(*ppMinVoltage)[1] = -2.0;			// BIP2VOLTS
	(*ppMaxVoltage)[1] = 2.0;
	(*ppMinVoltage)[2] = -5.0;			// BIP5VOLTS
	(*ppMaxVoltage)[2] = 5.0;
	(*ppMinVoltage)[3] = -10.0;			// BIP10VOLTS
	(*ppMaxVoltage)[3] = 10.0;
	return TRUE;
}

BOOL CMyReadCBDac::StartThreadMonitor()
{
	if (this->GetThreadRunning()) return FALSE;
	// create the monitor dialog
	this->m_hwndThreadMonitor = CreateDialogParam(GetServer()->GetInstance(), MAKEINTRESOURCE(IDD_DIALOGThreadMonitor), NULL, (DLGPROC)DlgProcThreadMonitor, (LPARAM)this);
	// clear the flag
	this->SetStopThread(FALSE);
	// create the thread
	this->m_hThread = CreateThread(NULL, 0, (LPTHREAD_START_ROUTINE)ThreadProcADScan, (LPVOID) this, 0, NULL);
	return this->GetThreadRunning();
}

void CMyReadCBDac::StopThreadMonitor()
{
	if (this->GetThreadRunning())
	{
		this->SetStopThread(TRUE);
	}
}

void CMyReadCBDac::OnCheckADProcess()
{
	int			ulStat;
	short		Status;
	long		CurCount;
	long		CurIndex;

	ulStat = cbGetStatus(this->m_boardNumber, &Status, &CurCount, &CurIndex, AIFUNCTION);
	if (0 == ulStat && NULL != this->m_hMemory)
	{
		// check if background process complete
		if (0 == Status)
		{
			// stop the thread
			this->SetStopThread(TRUE);
			// 
			cbStopBackground(this->m_boardNumber, AIFUNCTION);
			// sink notification
			this->m_pMyObject->FireHaveArrayScan(CurCount, (unsigned short*)this->m_hMemory);
			// cleanup
			cbWinBufFree(this->m_hMemory);
			this->m_hMemory = NULL;
		}
	}
	else
	{
		this->SetStopThread(TRUE);
	}
}

// synchronization
BOOL CMyReadCBDac::OwnThis()
{
	EnterCriticalSection(&(this->m_CriticalSection));
	return TRUE;
}

void CMyReadCBDac::UnOwnThis()
{
	LeaveCriticalSection(&(this->m_CriticalSection));
}

BOOL CMyReadCBDac::GetThreadRunning()
{
	if (NULL != this->m_hThread)
	{
		DWORD		dwExitCode;
		GetExitCodeThread(this->m_hThread, &dwExitCode);
		return STILL_ACTIVE == dwExitCode;
	}
	else
	{
		return FALSE;
	}
}

BOOL CMyReadCBDac::GetStopThread()
{
	BOOL		fStop = TRUE;
	if (this->OwnThis())
	{
		fStop = this->m_fStopIntegration;
		this->UnOwnThis();
	}
	return fStop;
}

void CMyReadCBDac::SetStopThread(
	BOOL			fStopThread)
{
	if (this->OwnThis())
	{
		this->m_fStopIntegration = fStopThread;
		this->UnOwnThis();
	}
}

// thread monitor dialog procedure
LRESULT CALLBACK DlgProcThreadMonitor(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CMyReadCBDac	*	pReadCBDac = NULL;
	switch (uMsg)
	{
	case WM_INITDIALOG:
		SetWindowLongPtr(hwndDlg, GWLP_USERDATA, lParam);
		return TRUE;
	case WM_DESTROY:
		pReadCBDac = (CMyReadCBDac *)GetWindowLongPtr(hwndDlg, GWLP_USERDATA);
		pReadCBDac->m_hwndThreadMonitor = NULL;
		pReadCBDac->m_fIntegratingSignal = FALSE;
		Utils_OnPropChanged(pReadCBDac->m_pMyObject, DISPID_IntegratingSignal);
		break;
	case WM_CHECKADPROCESS:
		pReadCBDac = (CMyReadCBDac *)GetWindowLongPtr(hwndDlg, GWLP_USERDATA);
		pReadCBDac->OnCheckADProcess();
		return TRUE;
	case WM_THREADCLOSED:
		DestroyWindow(hwndDlg);
		return TRUE;
	default:
		break;
	}
	return FALSE;
}

// thread procedure
DWORD WINAPI ThreadProcADScan(LPVOID pv)
{
	CMyReadCBDac*	pReadCBDac = (CMyReadCBDac*)pv;
	BOOL			fDone = FALSE;
	while (!fDone)
	{
		if (!pReadCBDac->GetStopThread())
		{
			Sleep(25);
			PostMessage(pReadCBDac->m_hwndThreadMonitor, WM_CHECKADPROCESS, 0, 0);
		}
		else fDone = TRUE;
	}
	PostMessage(pReadCBDac->m_hwndThreadMonitor, WM_THREADCLOSED, 0, 0);
	return 0;
}
