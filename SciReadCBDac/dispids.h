#pragma once

// dispatch ids

// properties
#define			DISPID_BoardNumber					0x0100
#define			DISPID_AmInitialized				0x0101
#define			DISPID_ChannelNumber				0x0102
#define			DISPID_IntegratingSignal			0x0103
#define			DISPID_IntegrationTime				0x0104
#define			DISPID_ScanRate						0x0105
#define			DISPID_VoltageRange					0x0106
#define			DISPID_ADResolution					0x0107
#define			DISPID_VoltageRanges				0x0108
#define			DISPID_WantIntegrateSignal			0x0109
#define			DISPID_WantAutoTimeConstant			0x010a

// methods
#define			DISPID_GetSignalReading				0x0120
#define			DISPID_ConfigDigitalBit				0x0121
#define			DISPID_WriteDigitalBit				0x0122
#define			DISPID_ReadDigitalBit				0x0123
#define			DISPID_OutputAnalog					0x0124
#define			DISPID_ReadAnalogInput				0x0125
#define			DISPID_CountsToVoltage				0x0126
#define			DISPID_VoltageToCounts				0x0127
#define			DISPID_Setup						0x0128
#define			DISPID_SetupAutoTimeConstant		0x0129
#define			DISPID_TestAutoTimeConstant			0x012a

// events
#define			DISPID_QueryMainWindow				0x0140
#define			DISPID_RequestConfigFile			0x0141
#define			DISPID_RequestCurrentPosition		0x0142
#define			DISPID_RequestSetPosition			0x0143
#define			DISPID_Error						0x0144
#define			DISPID_RequestIniFile				0x0145
#define			DISPID_RequestScanningStatus		0x0146
#define			DISPID_RequestWorkingDirectory		0x0147
#define			DISPID_HaveArrayScan				0x0148
#define			DISPID_HaveVoltageScan				0x0149
#define			DISPID_haveIntSignal				0x014a

// container
#define			DISPID_Container_CreateObject		0x0180
#define			DISPID_Container_OnPropRequestEdit	0x0181
#define			DISPID_Container_OnPropChanged		0x0182
#define			DISPID_Container_OnSinkEvent		0x0183
#define			DISPID_Container_CloseObject		0x0184
#define			DISPID_Container_AmbientProperty	0x0185
#define			DISPID_Container_GetNamedObject		0x0186
#define			DISPID_Container_CloseAllObjects	0x0187
#define			DISPID_Container_TryMNemonic		0x0188

// extended object
#define			DISPID_Extended_DoCreateObject		0x0200
#define			DISPID_Extended_DoCloseObject		0x0201
#define			DISPID_Extended_TryMNemonic			0x0202
