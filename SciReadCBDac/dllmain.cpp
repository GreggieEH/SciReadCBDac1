// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"
#include "Server.h"
#include <initguid.h>
#include "MyGuids.h"

CServer * g_pServer = NULL;

BOOL WINAPI DllMain(
	HINSTANCE hinstDLL,
	DWORD fdwReason,
	LPVOID lpvReserved
	)
{
	switch (fdwReason)
	{
	case DLL_PROCESS_ATTACH:
		g_pServer = new CServer(hinstDLL);
		break;
	case DLL_PROCESS_DETACH:
		Utils_DELETE_POINTER(g_pServer);
		break;
	default:
		break;
	}
	return TRUE;
}

CServer * GetServer()
{
	return g_pServer;
}

// yield for messages
void MyYield()
{
	MSG					msg;			// message structure
	while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
}

