#pragma once

class CMyObject;

class CThreadMonitor
{
public:
	CThreadMonitor(CMyObject * pMyObject);
	~CThreadMonitor();
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
protected:
	BOOL			OwnThis();
	void			UnOwnThis();
private:
	CMyObject	*	m_pMyObject;
	HINSTANCE		m_hInstance;
	HWND			m_hwndMessage;			// message window
	HWND			m_hwndThread;			// thread window
	HANDLE			m_hThread;

};

