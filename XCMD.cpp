


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Header
#include "Main.h"
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// DISPlay setting
HRESULT DISP(PCTSTR ptzCmd)
{
	DEVMODE dmOld;
	EnumDisplaySettings(NULL, ENUM_CURRENT_SETTINGS, &dmOld);

	DEVMODE dmNew = dmOld;
	dmNew.dmPelsWidth = UStrToInt(ptzCmd);
	ptzCmd = UStrChr(ptzCmd, ',');
	if (ptzCmd)
	{
		dmNew.dmPelsHeight = UStrToInt(++ptzCmd);
		ptzCmd = UStrChr(ptzCmd, ',');
		if (ptzCmd)
		{			
			dmNew.dmBitsPerPel = UStrToInt(++ptzCmd);
			ptzCmd = UStrChr(ptzCmd, ',');
			if (ptzCmd)
			{
				dmNew.dmDisplayFrequency = UStrToInt(++ptzCmd);
			}
		}
	}

	if ((dmNew.dmPelsWidth == dmOld.dmPelsWidth) && (dmNew.dmPelsHeight == dmOld.dmPelsHeight) &&
		(dmNew.dmBitsPerPel == dmOld.dmBitsPerPel) && (dmNew.dmDisplayFrequency == dmOld.dmDisplayFrequency))
	{
		return S_OK;
	}

	LONG lResult = ChangeDisplaySettings(&dmNew, CDS_UPDATEREGISTRY);
	if (lResult != DISP_CHANGE_SUCCESSFUL)
	{
		ChangeDisplaySettings(&dmOld, CDS_UPDATEREGISTRY);
	}

	return lResult;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// PAGE file
#include <NTSecAPI.h>
#define REG_PageFile TEXT("PagingFiles")
#define REG_MemMgr TEXT("SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management")
HRESULT PAGE(PTSTR ptzCmd)
{
	// Parse size
	UINT uMin = 0;
	UINT uMax = 0;
	PTSTR p = UStrChr(ptzCmd, ' ');
	if (p)
	{
		*p++ = 0;
		uMin = UStrToInt(p);
		p = UStrChr(p, ' ');
		if (p)
		{
			*p++ = 0;
			uMax = UStrToInt(p);
		}
	}

	if (uMax<uMin)
	{
		uMax=uMin;
	}

	// Get DOS device name for page file
	TCHAR tzDrive[16];
	TCHAR tzDos[MAX_PATH];
	TCHAR tzFile[MAX_PATH];
	tzDrive[0] = ptzCmd[0]; tzDrive[1] = ptzCmd[1]; tzDrive[2] = 0;
	UStrCopy(tzFile, ptzCmd + 2);
	QueryDosDevice(tzDrive, tzDos, MAX_PATH);
	UStrCat(tzDos, tzFile);

	WCHAR wzPath[MAX_PATH];
	UStrToWStr(wzPath, tzDos, MAX_PATH);

	UNICODE_STRING sPath;
	sPath.Length = UWStrLen(wzPath) * sizeof(WCHAR);
	sPath.MaximumLength = sPath.Length + sizeof(WCHAR);
	sPath.Buffer = wzPath;

	// Fill size param
	ULARGE_INTEGER ulMax, ulMin;
	ulMin.QuadPart = uMin * 1024 * 1024;
	ulMax.QuadPart = uMax * 1024 * 1024;

	// Get function address
	typedef NTSTATUS (NTAPI* PNtCreatePagingFile)(PUNICODE_STRING sPath, PULARGE_INTEGER puInitSize, PULARGE_INTEGER puMaxSize, ULONG uPriority);
	PNtCreatePagingFile NtCreatePagingFile = (PNtCreatePagingFile) GetProcAddress(GetModuleHandle(TEXT("NTDLL")), "NtCreatePagingFile");
	if (!NtCreatePagingFile)
	{
		return E_FAIL;
	}

	// Create page file
	Priv(SE_CREATE_PAGEFILE_NAME);
	HRESULT hResult = NtCreatePagingFile(&sPath, &ulMin, &ulMax, 0);
	if (hResult == S_OK)
	{
		// Log to Windows Registry
		TCHAR tzStr[MAX_PATH];
		DWORD i = sizeof(tzStr);
		if (SHGetValue(HKEY_LOCAL_MACHINE, REG_MemMgr, REG_PageFile, NULL, tzStr, &i) != S_OK)
		{
			i = 0;
		}
		else
		{
			i = (i / sizeof(TCHAR)) - 1;
		}

		i += UStrPrint(tzStr + i, TEXT("%s %d %d"), ptzCmd, uMin, uMax);
		tzStr[++i] = 0;
		SHSetValue(HKEY_LOCAL_MACHINE, REG_MemMgr, REG_PageFile, REG_MULTI_SZ, tzStr, i * sizeof(TCHAR));
	}

	return hResult;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// SERVice control
HRESULT SERV(PCTSTR ptzCmd)
{
	BOOL bResult = FALSE;
	SC_HANDLE hManager = OpenSCManager(NULL, NULL, SC_MANAGER_ALL_ACCESS);
	if (hManager)
	{
		BOOL bStop = (*ptzCmd == '!');
		if (bStop) ptzCmd++;

		SC_HANDLE hService = OpenService(hManager, ptzCmd, SERVICE_START | SERVICE_STOP);
		if (hService)
		{
			if (bStop)
			{
				SERVICE_STATUS ss;
				bResult = ControlService(hService, SERVICE_CONTROL_STOP, &ss);
			}
			else
			{
				bResult = StartService(hService, 0, NULL);
			}
			CloseServiceHandle(hService);
		}
		CloseServiceHandle(hManager);
	}
	return bResult ? S_OK : S_FALSE;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Test string equal
template <typename T1, typename T2> inline BOOL TStrEqual(T1 ptStr1, T2 ptStr2, UINT uLen)
{
	for (UINT i = 0; i < uLen; i++)
	{
		if (!UChrMatch((TCHAR) ptStr1[i], (TCHAR) ptStr2[i]))
		{
			return FALSE;
		}
	}
	return TRUE;
}

// Search string
template <typename T1, typename T2> PVOID TStrStr(T1 ptStr1, T2 ptStr2, UINT uLen)
{
	for (T1 p = ptStr1; *p; p++)
	{
		if (TStrEqual(p, ptStr2, uLen))
		{
			return p;
		}
	}
	return NULL;
}

// Lookup device ID from INFs
PCTSTR DevGetInf(PCTSTR ptzDevID, PCTSTR ptzSrcInf)
{
	PVOID pvFile = UFileLoad(ptzSrcInf, NULL, NULL);
	if (pvFile)
	{
		BOOL bASCII = ((PBYTE) pvFile)[3];//!IsTextUnicode(pvFile, -1, NULL); // Trick for UNICODE
		for (ptzDevID++; ptzDevID[-1]; ptzDevID += UStrLen(ptzDevID) + 2)
		{
			if (bASCII ? TStrStr((PSTR) pvFile, ptzDevID, ptzDevID[-1]) : TStrStr((PWSTR) pvFile, ptzDevID, ptzDevID[-1]))
			{
				Log(IDS_FoundDriver, ptzDevID, ptzSrcInf);
				UMemFree(pvFile);
				return ptzDevID;
			}
		}
		UMemFree(pvFile);
	}
	return NULL;
}

// Update device driver
#ifdef _UNICODE
#define STR_UpdateDriverForPlugAndPlayDevices "UpdateDriverForPlugAndPlayDevicesW"
#else
#define STR_UpdateDriverForPlugAndPlayDevices "UpdateDriverForPlugAndPlayDevicesA"
#endif
typedef BOOL (WINAPI* PUPNP)(HWND hWnd, PCTSTR ptzDevID, PCTSTR ptzPath, DWORD dwFlags, PBOOL pbReboot);
BOOL DevIns(PCTSTR ptzDevID, PCTSTR ptzInfPath, DWORD dwForce = 3)
{
	BOOL bResult = FALSE;
	HMODULE hLib = LoadLibrary(TEXT("NewDev.dll"));
	if (hLib)
	{
		// Install INF
		PUPNP p = (PUPNP) GetProcAddress(hLib, STR_UpdateDriverForPlugAndPlayDevices);
		if (p)
		{
			BOOL bReboot = FALSE;
			bResult = p(NULL, ptzDevID, ptzInfPath, dwForce, &bReboot);
		}
		FreeLibrary(hLib);
	}
	return bResult;
}

// Install driver from DIR
#include <SetupAPI.h>
#pragma comment(lib, "SetupAPI.lib")
UINT DevDir(PCTSTR ptzDir, PCTSTR ptzDevID, PCTSTR ptzClass)
{
	TCHAR tzPath[MAX_NAME];
	if (ptzDir[0] == '\\')
	{
		ptzDir++;
		TCHAR tzDrives[MAX_NAME];
		GetLogicalDriveStrings(MAX_NAME, tzDrives);
		for (PTSTR p = tzDrives; *p; p += UStrLen(p) + 1)
		{
			UStrPrint(tzPath, TEXT("%s%s"), p, ptzDir);
			DevDir(tzPath, ptzDevID, ptzClass);
		}
		return S_OK;
	}

	WIN32_FIND_DATA fd;
	UStrPrint(tzPath, TEXT("%s\\INF\\*.INF"), ptzDir);
	HANDLE hFind = FindFirstFile(tzPath, &fd);
	if (hFind == INVALID_HANDLE_VALUE)
	{
		return ERROR_FILE_NOT_FOUND;
	}

	do
	{
		UStrPrint(tzPath, TEXT("%s\\INF\\%s"), ptzDir, fd.cFileName);
		if (ptzClass)
		{
			GUID idClass;
			TCHAR tzClass[MAX_NAME];
			SetupDiGetINFClass(tzPath, &idClass, tzClass, MAX_NAME, NULL);
			if (UStrCmpI(tzClass, ptzClass))
			{
				continue;
			}
		}
		PCTSTR p = DevGetInf(ptzDevID, tzPath);
		if (p)
		{
			DevIns(p, tzPath, 0);
		}
	}
	while (FindNextFile(hFind, &fd));

	FindClose(hFind);
	return S_OK;
}

// Extract driver from CAB
PTSTR g_ptzDevInf = NULL;
UINT CALLBACK DevCab(PVOID pvContext, UINT uMsg, UINT_PTR upParam1, UINT_PTR upParam2)
{
	static UINT s_uExtract = 0;

	if (uMsg == SPFILENOTIFY_FILEINCABINET)
	{
		PTSTR ptzTarget = ((FILE_IN_CABINET_INFO*) upParam1)->FullTargetName;
		PCTSTR ptzName = ((FILE_IN_CABINET_INFO*) upParam1)->NameInCabinet;

		PCTSTR p = UStrRChr(ptzName, '\\');
		if (p)
		{
			ptzName = p + 1;
		}

		// Extract INF or driver file
		p = ptzName + UStrLen(ptzName) - 4;
		if (UStrCmpI(p, TEXT(".INF")) == 0)
		{
			p = TEXT("%SystemRoot%\\INF\\");
		}
		else if (s_uExtract)
		{
			if (UStrCmpI(p, TEXT(".SYS")) == 0)
			{
				p = TEXT("%SystemRoot%\\SYSTEM32\\DRIVERS\\");
			}
			else if (UStrCmpI(p, TEXT(".DLL")) == 0)
			{
				p = TEXT("%SystemRoot%\\SYSTEM32\\");
			}
			else
			{
				p = TEXT("%SystemRoot%\\");
			}
		}
		else
		{
			// Skip
			return FILEOP_SKIP;
		}

		ExpandEnvironmentStrings(p, ptzTarget, MAX_PATH);
		UStrCat(ptzTarget, ptzName);
		UStrRep(ptzTarget, '#', '\\');
		UDirCreate(ptzTarget);
		return FILEOP_DOIT;
	}
	else if (uMsg == SPFILENOTIFY_FILEEXTRACTED)
	{
		PCTSTR ptzTarget = ((FILEPATHS*) upParam1)->Target;
		if (UStrCmpI(ptzTarget + UStrLen(ptzTarget) - 4, TEXT(".INF")))
		{
			// Not INF
			s_uExtract++;
			return NO_ERROR;
		}

		// Get Device from INF
		PCTSTR ptzDevID = DevGetInf((PCTSTR) pvContext, ptzTarget);
		if (ptzDevID)
		{
			// Found Driver
			s_uExtract = 1;
			do {*g_ptzDevInf++ = *ptzDevID;} while (*ptzDevID++);
			do {*g_ptzDevInf++ = *ptzTarget;} while (*ptzTarget++);
			return NO_ERROR;
		}

		// Delete INF
		if (s_uExtract != 1)
		{
			// Driver has been extracted completely.
			s_uExtract = 0;
			UFileDelete(ptzTarget);
		}
	}
	return NO_ERROR;
}

// DEVIce driver installation
#if (_MSC_VER < 1300)
#define DN_HAS_PROBLEM 0x00000400
#define CM_PROB_NOT_CONFIGURED 0x00000001
typedef DWORD (WINAPI* PCM_Get_DevNode_Status)(OUT PULONG pulStatus, OUT PULONG pulProblemNumber, IN  DWORD dnDevInst, IN  ULONG ulFlags);
#else
#include <CfgMgr32.h>
#endif
#define MAX_DevID 2048
#define NUM_DevID ((UINT) ((PBYTE) p - (PBYTE) tzDevID))
HRESULT DEVI(PTSTR ptzCmd)
{
	// Lookup device
	HDEVINFO hDev = SetupDiGetClassDevs(NULL, NULL, 0, DIGCF_ALLCLASSES);
	if (hDev == INVALID_HANDLE_VALUE)
	{
		return E_FAIL;
	}

	// Build SPDRP_HARDWAREID list
	TCHAR tzDevID[MAX_DevID];
	PTSTR p = tzDevID + 1;
	SP_DEVINFO_DATA sdDev = {sizeof(SP_DEVINFO_DATA)};
	for (UINT i = 0; (NUM_DevID < MAX_DevID) && SetupDiEnumDeviceInfo(hDev, i, &sdDev); i++)
	{
#if (_MSC_VER < 1300)
		PCM_Get_DevNode_Status CM_Get_DevNode_Status = (PCM_Get_DevNode_Status) GetProcAddress(GetModuleHandle(TEXT("SetupAPI")), "CM_Get_DevNode_Status");
		if (CM_Get_DevNode_Status)
#endif
		{
			// Exclude configured device
			ULONG uProblem = 0;
			ULONG uStatus = DN_HAS_PROBLEM;
			CM_Get_DevNode_Status(&uStatus, &uProblem, sdDev.DevInst, 0);
			if (uProblem != CM_PROB_NOT_CONFIGURED)
			{
#ifndef _DEBUG
				continue;
#endif
			}
		}

		// Get device ID
		if (SetupDiGetDeviceRegistryProperty(hDev, &sdDev, SPDRP_HARDWAREID, NULL, (PBYTE) p, MAX_DevID - NUM_DevID, NULL) && UStrCmpNI(p, TEXT("ACPI"), 4))
		{
			Log(IDS_FoundDevice, p);

			// Trim some stuff for quick search
			UINT j = 0;
			for (UINT k = 0; p[j]; j++)
			{
				if ((p[j] == '&') && (++k == 2))
				{
					break;
				}
			}
			p[-1] = j;
			for (p += j; *p; p++);
			p += 2;
		}
	}
	p[-1] = 0;

	SetupDiDestroyDeviceInfoList(hDev);
	if (tzDevID[0] == 0)
	{
		// No device
		return ERROR_NO_MATCH;
	}

	// Parse param
	BOOL bInstall = (ptzCmd[0] == '$');
	if (bInstall) ptzCmd++;
	PTSTR ptzClass = UStrChr(ptzCmd, ',');
	if (ptzClass) *ptzClass++ = 0;

	if (UStrCmpI(ptzCmd + UStrLen(ptzCmd) - 4, TEXT(".CAB")))
	{
		// Lookup driver from directory
		return DevDir(ptzCmd, tzDevID, ptzClass);
	}
	else
	{
		// Lookup CAB file
		TCHAR tzDevInf[MAX_PATH * 16];
		g_ptzDevInf = tzDevInf;
		HRESULT hResult = SetupIterateCabinet(ptzCmd, 0, (PSP_FILE_CALLBACK) DevCab, tzDevID) ? S_OK : E_FAIL;
		if (bInstall)
		{
			for (PTSTR p = tzDevInf; p < g_ptzDevInf; p += UStrLen(p) + 1)
			{
				PTSTR ptzDevID = p;
				p += UStrLen(p) + 1;
				DevIns(ptzDevID, p);
			}
		}
		return hResult;
	}
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  Folder macro
#include <ShlObj.h>
const struct {INT iFolder; PCTSTR ptzMacro;} c_sMacro[] =
{
	{CSIDL_FAVORITES,		TEXT("Favorites")},
	{CSIDL_DESKTOPDIRECTORY,TEXT("Desktop")},
	{CSIDL_STARTMENU,		TEXT("StartMenu")},
	{CSIDL_STARTUP,			TEXT("Startup")},
	{CSIDL_PROGRAMS,		TEXT("Programs")},
	{CSIDL_SENDTO,			TEXT("SendTo")},
	{CSIDL_PERSONAL,		TEXT("Personal")},
	{CSIDL_APPDATA,			TEXT("QuickLaunch")},
};

//  ENVIronment
HRESULT ENVI(PTSTR ptzEnv, BOOL bSystem)
{
	TCHAR tzStr[MAX_STR];
	if (!ptzEnv[0] || (ptzEnv[0] == '='))
	{
		for (UINT i = 0; i < _NumOf(c_sMacro); i++)
		{
			if (ptzEnv[0]/* == '='*/)
			{
				ENVI((PTSTR) c_sMacro[i].ptzMacro, bSystem);
				continue;
			}

			PTSTR p = tzStr + UStrPrint(tzStr, TEXT("%s="), c_sMacro[i].ptzMacro);
			if (SHGetSpecialFolderPath(NULL, p, c_sMacro[i].iFolder, TRUE))
			{
				if (c_sMacro[i].iFolder == CSIDL_APPDATA)
				{
					// Trick
					UStrCat(p, TEXT("\\Microsoft\\Internet Explorer\\Quick Launch\\"));
					UDirCreate(p);
					CreateDirectory(p, NULL);
				}
				ENVI(tzStr, bSystem);
			}
		}

		if (ptzEnv[0]/* == '='*/)
		{
			return ENVI((PTSTR) STR_AppName, bSystem);
		}
		else
		{
			UStrPrint(tzStr, TEXT("%s=%s/%s"), STR_AppName, STR_VersionStamp, STR_BuildStamp);
			return ENVI(tzStr, bSystem);
		}
	}

	if (bSystem)
	{
		UStrPrint(tzStr, TEXT("HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment\\%s"), ptzEnv);
		REGX(tzStr);
	}

	PTSTR ptzValue = UStrChr(ptzEnv, '=');
	if (ptzValue)
	{
		// Trick : If no '=', ptzEnv can be a constant string.
		*ptzValue++ = 0;
	}
	HRESULT hResult = SetEnvironmentVariable(ptzEnv, ptzValue) ? S_OK : GetLastError();

	if (bSystem)
	{
		if (!ptzValue)
		{
			ptzValue = ptzEnv + UStrLen(ptzEnv) + 1;
		}
		DWORD dwResult = 0;
		SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0, (LPARAM) TEXT("Environment"), SMTO_ABORTIFHUNG, 5000, &dwResult);
	}
	return hResult;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
