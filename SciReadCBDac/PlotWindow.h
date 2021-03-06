#pragma once
#include "MyContainer.h"

class CDlgRunCBDac;

class CPlotWindow : public CMyContainer
{
public:
	CPlotWindow(CDlgRunCBDac * pDlgRunCBDac);
	virtual ~CPlotWindow();
	BOOL						CreatePlotWindow(
									HWND			hwndParent);
	void						ClosePlotWindow();
	BOOL						AddDataSet(
									LPCTSTR			szName,
									unsigned long	curveColor);
	void						SetCurveData(
									LPCTSTR			szName,
									long			nValues,
									double		*	pX,
									double		*	pY);
	BOOL						ZoomOut();
	long						GetNumCurves();
	void						Refresh();
	void						SetXRange(
									double			xMin,
									double			xMax);
	void						GetYRange(
									double		*	yMin,
									double		*	yMax);
	void						SetYRange(
									double			yMin,
									double			yMax);
	void						GetDataXRange(
									double		*	xMin,
									double		*	xMax);
	void						GetDataYRange(
									double			xMin,
									double			xMax,
									double		*	yMin,
									double		*	yMax);
	BOOL						TryMNemonic(
									LPMSG			pmsg);
	void						SetTitles(
									LPCTSTR			szTitle,
									LPCTSTR			szXTitle,
									LPCTSTR			szYTitle);
	void						SetCurveVisible(
									LPCTSTR			szName,
									BOOL			fVisible);
	HWND						GetControlWindow();
	BOOL						GetZoomedIn();
	BOOL						GetCurveXData(
									long			curveIndex,
									long		*	nArray,
									double		**	ppX);
	void						SetCurveXData(
									long			curveIndex,
									long			nArray,
									double		*	pX);
	BOOL						GetCurveYData(
									long			curveIndex,
									long		*	nArray,
									double		**	ppY);
	void						SetCurveYData(
									long			curveIndex,
									long			nArray,
									double		*	pY);
	BOOL						AddDataValue(
									long			Curve,
									double			xValue,
									double			yValue);
protected:
	virtual HWND				GetMainWindow();				// pure virtual
	virtual void				OnObjectCreated(
									LPCTSTR			szObject,
									IUnknown	*	pUnkSite);
	virtual BOOL				OnPropRequestEdit(
									LPCTSTR			szObject,
									DISPID			dispid);
	virtual void				OnPropChanged(
									LPCTSTR			szObject,
									DISPID			dispid);
	virtual void				OnSinkEvent(
									LPCTSTR			ObjectName,
									long			Dispid,
									long			nParams,
									VARIANT		*	Parameters);
	virtual HRESULT				MySetActiveObject(
									IOleInPlaceActiveObject *pActiveObject);
	BOOL						GetOurObject(
									IDispatch	**	ppdisp);
private:
	CDlgRunCBDac			*	m_pDlgRunCBDac;
	HWND						m_hwndParent;
	long						m_nextCurveIndex;
};

