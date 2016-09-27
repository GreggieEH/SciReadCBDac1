// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files:
#include <windows.h>
#include <Windowsx.h>
#include <ole2.h>
#include <olectl.h>

#include <shlwapi.h>
#include <stdio.h>
#include <tchar.h>
#include <math.h>

#include <propvarutil.h>
#include <Shobjidl.h>
#include <ShlObj.h>
#include <Shellapi.h>

#include <vector>

#include <strsafe.h>
#include "G:\Users\Greg\Documents\Visual Studio 2015\Projects\MyUtils\MyUtils\myutils.h"
#include "resource.h"

class CServer;
CServer * GetServer();

// yield for messages
void MyYield();

// definitions
#define				MY_CLSID					CLSID_SciReadCBDac
#define				MY_LIBID					LIBID_SciReadCBDac
#define				MY_IID						IID_ISciReadCBDac
#define				MY_IIDSINK					IID__SciReadCBDac

#define				MAX_CONN_PTS				2
#define				CONN_PT_CUSTOMSINK			0
#define				CONN_PT_PROPNOTIFY			1

#define				FRIENDLY_NAME				TEXT("SciReadCBDac")
#define				PROGID						TEXT("Sciencetech.SciReadCBDac.1")
#define				VERSIONINDEPENDENTPROGID	TEXT("Sciencetech.SciReadCBDac")

// object names
#define				PLOT_WINDOW					L"SciPlotWindow"