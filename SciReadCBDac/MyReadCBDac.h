#pragma once

class CMyObject;

class CMyReadCBDac
{
public:
	CMyReadCBDac(CMyObject * pMyObject);
	~CMyReadCBDac();
	long			GetBoardNumber();
	void			SetBoardNumber(
						long		boardNumber);
	BOOL			GetAmInitialized();
	void			SetAmInitialized(
						BOOL		fAmInitialized);
	long			GetChannelNumber();
	void			SetChannelNumber(
						long		channelNumber);
	BOOL			GetIntegratingSignal();
	void			SetIntegratingSignal(
						BOOL		integratingSignal);
	// integration time in ms
	long			GetIntegrationTime();
	void			SetIntegrationTime(
						long		integrationTime);
	// scan rate data per second
	long			GetScanRate();
	void			SetScanRate(
						long		scanRate);
	// last signal reading
	ULONG			GetSignalReading();
	// get scan array
	BOOL			GetScanArray(
						long	*	nValues,
						ULONG	**	ppScanData);
	// voltage range
	void			GetVoltageRange(
						double	*	minVoltage,
						double	*	maxVoltage);
	void			SetVoltageRange(
						double		minVoltage,
						double		maxVoltage);
	// AD resolution
	long			GetADResolution();
	// configure a digital bit for input or output
	BOOL			ConfigDigitalBit(
						short		bitNumber,
						BOOL		forOutput);
	BOOL			WriteDigitalBit(
						short		bitNumber,
						BOOL		fHi);
	BOOL			ReadDigitalBit(
						short		bitNumber,
						BOOL	*	fHi);
	// output analog signal
	BOOL			OutputAnalog(
						int			channel,
						double		value);
	BOOL			ReadAnalogInput(
						int			channel,
						unsigned short*	dataValue);
	// convert to engineering units
	BOOL			CountsToVoltage(
						unsigned short	dataValue,
						double		*	voltage);
	BOOL			VoltageToCounts(
						double		voltage,
						unsigned short*	dataValue);
	// available voltage ranges
	BOOL			GetVoltageRanges(
						long		*	nRanges,
						double		**	ppMinVoltage,
						double		**	ppMaxVoltage);
	// integrate signal
	double			IntegrateSignal();
	void			SetSignalAverage(
						double			signalAverage);
protected:
	// synchronization
	BOOL			OwnThis();
	void			UnOwnThis();

	BOOL			StartThreadMonitor();
	void			StopThreadMonitor();
	void			OnCheckADProcess();
	BOOL			GetThreadRunning();
	BOOL			GetStopThread();
	void			SetStopThread(
						BOOL			fStopThread);

private:
	CMyObject	*	m_pMyObject;
	// properties
	long			m_boardNumber;
	BOOL			m_SystemInitialized;
	long			m_ChannelNumber;
	int				m_ADRange;
	long			m_ScanRate;
	// parameters needed for signal integration:
	long			m_integrationTime;			// integration time in ms
	BOOL			m_fIntegratingSignal;
	HWND			m_hwndThreadMonitor;
	BOOL			m_fStopIntegration;			// flag stop integration
	HANDLE			m_hMemory;					// windows memory handle
	long			m_nPoints;					// number of points in the memory
	HANDLE			m_hThread;					// thread handle
	// critical section object
	CRITICAL_SECTION	m_CriticalSection;
	// signal average
	double			m_signalAverage;

	// thread monitor dialog procedure
	friend LRESULT CALLBACK DlgProcThreadMonitor(HWND, UINT, WPARAM, LPARAM);
	// thread procedure
	friend DWORD WINAPI		ThreadProcADScan(LPVOID pv);

};

