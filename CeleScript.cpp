


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Header
#include "Main.h"

#define CC_SEP 1
#define CC_LEN 4

enum
{
	CMD_LOAD, CMD_BATC, CMD_IFEX, CMD_ELSE, 
	CMD_CALL, CMD_GOTO, CMD_PROC, CMD_ENDP,
	CMD_PLAY, CMD_BEEP, CMD_MSGX, CMD_DLGX,
	CMD_LINK, CMD_FILE, CMD_REGX, CMD_ENVI,
	CMD_SEND, CMD_WAIT, CMD_KILL, CMD_SHUT,
	CMD_EXEC, CMD_CDLL, CMD_EVAL, CMD_ASOC, 
	CMD_DEVI, CMD_SERV, CMD_PAGE, CMD_DISP, 
	CMD_CCUI, 	
};

const TCHAR c_tzCmd[][CC_LEN + 1] =
{
	TEXT("LOAD"), TEXT("BATC"), TEXT("IFEX"), TEXT("ELSE"), 
	TEXT("CALL"), TEXT("GOTO"), TEXT("PROC"), TEXT("ENDP"),
	TEXT("PLAY"), TEXT("BEEP"), TEXT("MSGX"), TEXT("DLGX"),
	TEXT("LINK"), TEXT("FILE"), TEXT("REGX"), TEXT("ENVI"),
	TEXT("SEND"), TEXT("WAIT"), TEXT("KILL"), TEXT("SHUT"),
	TEXT("EXEC"), TEXT("CDLL"), TEXT("EVAL"), TEXT("ASOC"),
	TEXT("DEVI"), TEXT("SERV"), TEXT("PAGE"), TEXT("DISP"), 
	TEXT("CCUI"), 
};

typedef struct _CXT
{
	PCTSTR ptzCur;
	PCTSTR ptzNext;
	PCTSTR ptzFile;
	HRESULT hXVar;
}
CXT, *PCXT;

TCHAR g_tzXVar[10][MAX_PATH] = {0};
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Get Registry root
HKEY GetRootKey(PTSTR& ptzKey)
{
	STATIC struct {HKEY hKey; TCHAR tzKey[20];} c_sRegRoot[] =
	{
		{HKEY_USERS, TEXT("HKU\\")},
		{HKEY_CURRENT_USER, TEXT("HKCU\\")},
		{HKEY_CLASSES_ROOT, TEXT("HKCR\\")},
		{HKEY_LOCAL_MACHINE, TEXT("HKLM\\")},
		{HKEY_USERS, TEXT("HKEY_USERS\\")},
		{HKEY_CURRENT_USER, TEXT("HKEY_CURRENT_USER\\")},
		{HKEY_CLASSES_ROOT, TEXT("HKEY_CLASSES_ROOT\\")},
		{HKEY_LOCAL_MACHINE, TEXT("HKEY_LOCAL_MACHINE\\")},
	};

	for (UINT i = 0; i < _NumOf(c_sRegRoot); i++)
	{
		UINT n = UStrMatch(c_sRegRoot[i].tzKey, ptzKey);
		if (c_sRegRoot[i].tzKey[n] == 0)
		{
			ptzKey += n;
			return c_sRegRoot[i].hKey;
		}
	}
	return NULL;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Import Windows Registry Script
#ifndef REG_MUI_SZ
#define REG_MUI_SZ 21
#endif
#define _SkipBlank(p) while ((*p == '\r') || (*p == '\n') || (*p == '\t') || (*p == ' ')) p++;
HRESULT Import(PTSTR ptzScript, PCTSTR* pptzStop = NULL)
{
	HKEY hKey = NULL;
	_SkipBlank(ptzScript);
	HRESULT hResult = S_OK;
	PTSTR ptzEnd = ptzScript;
	PTSTR ptzStart = ptzScript;
	while (*ptzScript)
	{
		switch (*ptzScript)
		{
		case '"':
			if (*ptzStart != '[')
			{
				ptzScript++;
				*ptzEnd++ = CC_SEP;
				continue;
			}
			break;

		case '\\':
			if (*ptzStart != '[')
			{
				ptzScript++;
				_SkipBlank(ptzScript);
				*ptzEnd++ = *ptzScript++;
				continue;
			}
			break;

		case '\r':
		case '\n':
			*ptzEnd = 0;
			if (*ptzStart == '[')
			{
				if (hKey)
				{
					RegCloseKey(hKey);
					hKey = NULL;
				}
				ptzScript[-1] = 0;
				if (*++ptzStart == '-')
				{
					HKEY hRoot = GetRootKey(++ptzStart);
					if (hRoot)
					{
						SHDeleteKey(hRoot, ptzStart);
					}
				}
				else
				{
					HKEY hRoot = GetRootKey(ptzStart);
					if (hRoot)
					{
						RegCreateKeyEx(hRoot, ptzStart, 0, NULL, 0, KEY_ALL_ACCESS, NULL, &hKey, NULL);
					}
				}
			}
			else if (hKey && ((*ptzStart == '@') || (*ptzStart == CC_SEP)))
			{
				PTSTR ptzVal;
				if (*ptzStart++ == '@')
				{
					ptzVal = UStrChr(ptzStart, '=');
					ptzStart = NULL;
				}
				else if (ptzVal = UStrChr(ptzStart, CC_SEP))
				{
					*ptzVal++ = 0;
					ptzVal = UStrChr(ptzVal, '=');
				}

				if (ptzVal)
				{
					PTSTR ptzTemp;
					if (ptzTemp = UStrStrI(ptzVal, TEXT("DWORD:")))
					{
						ptzVal = ptzTemp + 4;
						ptzVal[0] = '0';
						ptzVal[1] = 'x';
						DWORD dwData = UStrToInt(ptzVal);
						RegSetValueEx(hKey, ptzStart, 0, REG_DWORD, (PBYTE) &dwData, sizeof(DWORD));
					}
					else if (ptzTemp = UStrStrI(ptzVal, TEXT("HEX")))
					{
						DWORD dwType = REG_BINARY;
						if (ptzTemp[3] == '(')
						{
							ptzVal[2] = '0';
							ptzVal[3] = 'x';
							dwType = UStrToInt(ptzTemp + 2);
						}

						if (ptzVal = UStrChr(ptzTemp, ':'))
						{
							ptzVal++;
							UINT i = 0;
							PBYTE pbVal = (PBYTE) ptzVal;
							while (*ptzVal)
							{
								BYTE bVal = (UChrToHex(ptzVal[0]) << 4) | UChrToHex(ptzVal[1]);
								while (*ptzVal && (*ptzVal++ != ','));
								pbVal[i++] = bVal;
							}
							RegSetValueEx(hKey, ptzStart, 0, dwType, pbVal, i);
						}
					}
					else if (ptzTemp = UStrStrI(ptzVal, TEXT("MULTI_SZ:")))
					{
						if (ptzVal = UStrChr(ptzTemp, CC_SEP))
						{
							ptzEnd = ++ptzVal;
							ptzTemp = ptzVal;
							while (*ptzTemp)
							{
								if (*ptzTemp == CC_SEP)
								{
									ptzTemp++;
									*ptzEnd++ = 0;
									while (*ptzTemp && (*ptzTemp++ != CC_SEP));
								}
								else
								{
									*ptzEnd++ =*ptzTemp++;
								}
							}
							*ptzEnd++ = 0;
							RegSetValueEx(hKey, ptzStart, 0, REG_MULTI_SZ, (PBYTE) ptzVal, (DWORD) (ptzEnd - ptzVal) * sizeof(TCHAR));
						}
					}
					else
					{
						DWORD dwType = UStrStrI(ptzVal, TEXT("MUI_SZ:")) ? REG_MUI_SZ : REG_SZ;
						if (ptzTemp = UStrChr(ptzVal, CC_SEP))
						{
							ptzVal = ptzTemp + 1;
							if (ptzTemp = UStrChr(ptzVal, CC_SEP))
							{
								*ptzTemp++ = 0;
								RegSetValueEx(hKey, ptzStart, 0, dwType, (PBYTE) ptzVal, (DWORD) (ptzTemp - ptzVal) * sizeof(TCHAR));
							}
						}
						else if (ptzVal[1] == '-')
						{
							RegDeleteValue(hKey, ptzStart);
						}
					}
				}
			}
			else if (*ptzStart != ';')
			{
				hResult = ERROR_NOT_REGISTRY_FILE;
				if (pptzStop)
				{
					*ptzScript = '\n';
					ptzScript = ptzStart;
					goto _Stop;
				}				
			}
			ptzScript++;
			_SkipBlank(ptzScript);
			ptzStart = ptzEnd = ptzScript;
			continue;
		}

		*ptzEnd++ = *ptzScript++;
	}

_Stop:
	if (pptzStop)
	{
		*pptzStop = ptzScript;
	}
	if (hKey)
	{
		RegCloseKey(hKey);
	}
	return hResult;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// LOAD script file
#define CC_REG TEXT("REGEDIT4")
#define CC_REX TEXT("Windows Registry Editor Version 5.00")
HRESULT LOAD(PCTSTR ptzPath)
{
	// LOAD file
	UINT uSize = -1;
	PBYTE pbFile = (PBYTE) UFileLoad(ptzPath, NULL, &uSize);
	if (!pbFile)
	{
		return ERROR_OPEN_FAILED;
	}

	// Convert ASCII <=> UNICODE
	PTSTR ptzScript = (PTSTR) pbFile;
#ifdef _UNICODE
	if (pbFile[3])	// !IsTextUnicode(pbFile, -1, NULL)
	{
		uSize += 16;
		ptzScript = (PTSTR) UMemAlloc(uSize * sizeof(TCHAR));
		UAStrToWStr(ptzScript, (PCSTR) pbFile, uSize);
		UMemFree(pbFile);
		pbFile = (PBYTE) ptzScript;
	}
#else
	if (!pbFile[3])	// IsTextUnicode(pbFile, -1, NULL)
	{
		uSize += 16;
		ptzScript = (PTSTR) UMemAlloc(uSize * sizeof(TCHAR));
		UWStrToAStr(ptzScript, (PCWSTR) pbFile, uSize);
		UMemFree(pbFile);
		pbFile = (PBYTE) ptzScript;
	}
#endif

	// Skip UNICODE BOM and white space
	while ((*ptzScript == '\r') || (*ptzScript == '\n') || (*ptzScript == '\t') || (*ptzScript == ' ') || ((UTCHAR) *ptzScript >= 0x7F))
	{
		ptzScript++;
	}

	// Select process engine
	HRESULT hResult;
	if (UStrMatch(ptzScript, CC_REG) >= _NumOf(CC_REG) - 1)
	{
		hResult = Import(ptzScript + _NumOf(CC_REG));
	}
	else if (UStrMatch(ptzScript, CC_REX) >= _NumOf(CC_REX) - 1)
	{
		hResult = Import(ptzScript + _NumOf(CC_REX));
	}
	else
	{
		hResult = Process(ptzScript, ptzPath);
	}

	UMemFree(pbFile);
	return hResult;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// BATch Command
HRESULT Dispatch(PTSTR ptzCmd, CXT& XT);
HRESULT BATC(PTSTR ptzCmd, CXT& XT)
{
	for (PTSTR p = ptzCmd; *p; p++)
	{
		if (*p == ';')
		{
			*p = 0;
			Dispatch(ptzCmd, XT);
			ptzCmd = p + 1;
		}
	}
	return Dispatch(ptzCmd, XT);
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// IF condition is true, EXecute command
BOOL g_bResult = FALSE;
HRESULT IFEX(PTSTR ptzCmd)
{
	BOOL bNot = (*ptzCmd == '!');
	if (bNot)
	{
		ptzCmd++;
	}

	BOOL bResult;
	PTSTR p = ptzCmd;
	for (INT iVal = UStrToInt(ptzCmd); *p; p++)
	{
		switch (*p)
		{
		case '=':
			if (p[1] == '=')
			{
				bResult = (iVal == UStrToInt(p + 2));
			}
			else
			{
				*p++ = 0;
				bResult = (UStrCmpI(ptzCmd, p) == 0);
			}
			break;

		case '!':
			if (p[1] == '=')
			{
				bResult = (iVal != UStrToInt(p + 2));
			}
			else
			{
				*p++ = 0;
				bResult = (UStrCmpI(ptzCmd, p) != 0);
			}
			break;

		case '>':
			if (p[1] == '=')
			{
				bResult = (iVal >= UStrToInt(p + 2));
			}
			else
			{
				bResult = (iVal > UStrToInt(p + 1));
			}
			break;

		case '<':
			if (p[1] == '=')
			{
				bResult = (iVal <= UStrToInt(p + 2));
			}
			else
			{
				bResult = (iVal < UStrToInt(p + 1));
			}
			break;

		case '&':
			if (p[1] == '&')
			{
				bResult = (iVal && UStrToInt(p + 2));
			}
			else
			{
				bResult = (iVal & UStrToInt(p + 1));
			}
			break;

		case '|':
			if (p[1] == '|')
			{
				bResult = (iVal || UStrToInt(p + 2));
			}
			else
			{
				bResult = (iVal | UStrToInt(p + 1));
			}
			break;

		default:
			continue;
		}
		break;
	}

	if (!*p)
	{
		bResult = (GetFileAttributes(ptzCmd) != -1);
	}

	return g_bResult = (bResult && !bNot) || (!bResult && bNot);	
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// EVALuate variable
HRESULT EVAL(PTSTR ptzCmd)
{
	if (UChrIsNum(*ptzCmd))
	{
		UINT i = (*ptzCmd++ - '0');
		if (*ptzCmd++ == '=')
		{
			UStrCopyN(g_tzXVar[i], ptzCmd, MAX_PATH);
			return S_OK;
		}
		else
		{
			return UStrToInt(g_tzXVar[i]);
		}
	}
	else if ((*ptzCmd == 'x') || (*ptzCmd == 'X'))
	{
		return UStrToInt(ptzCmd + 1 + (ptzCmd[1] == '='));
	}
	else
	{
		_Zero(g_tzXVar);
		return S_OK;
	}
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// PLAY sound
#include <MMSystem.h>
#pragma comment(lib, "WinMM.lib")
HRESULT PLAY(PTSTR ptzCmd)
{
	DWORD dwFlags = SND_FILENAME;
	while (TRUE)
	{
		if (*ptzCmd == '!') dwFlags |= SND_ASYNC;
		else if (*ptzCmd == '$') dwFlags |= SND_ALIAS;
		else if (*ptzCmd == '*') dwFlags |= SND_LOOP;
		else break;
		ptzCmd++;
	}
	return !PlaySound(ptzCmd[0] ? ptzCmd : NULL, NULL, dwFlags);
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// MesSaGe boX
HRESULT MSGX(PTSTR ptzCmd)
{
	UINT uType = MB_ICONINFORMATION;
	PTSTR ptzCaption = UStrChr(ptzCmd, CC_SEP);
	if (ptzCaption)
	{
		*ptzCaption++ = 0;
		PTSTR ptzType = UStrChr(ptzCaption, CC_SEP);
		if (ptzType)
		{
			*ptzType++ = 0;
			uType = UStrToInt(ptzType);
		}
	}
	else
	{
		ptzCaption = STR_AppName;
	}
	return MessageBox(g_hWnd, ptzCmd, ptzCaption, uType | MB_SETFOREGROUND);
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// DiaLoG boX
#define CC_CTL 4531
#define CC_CHECK (WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX)
#define CC_RADIO (WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON)
#define CC_GROUP (CC_RADIO | WS_GROUP)
//#define PIX_PAD 10
#define PIX_PAD (GetSystemMetrics(SM_CXVSCROLL) + 4)
INT_PTR CALLBACK DLGX(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	STATIC UINT s_uCount = 0;
	switch (uMsg)
	{
	case WM_INITDIALOG:
		s_uCount = 0;
		if (lParam)
		{
			DWORD dwGroup = CC_CTL;
			PTSTR p = (PTSTR) lParam;
			UINT u = GetSystemMetrics(SM_CXSMICON) + 5;
			for (PTSTR q = p; s_uCount < 32; q++)
			{
				TCHAR t = *q;
				if ((t == 0) || (t == CC_SEP))
				{
					*q = 0;
					DWORD dwStyle;
					PCTSTR ptzClass;
					BOOL bCheck = FALSE;
					switch (*p)
					{
					case '{':
						dwGroup = (CC_CTL + s_uCount);

					case '[':
					case '<':
						TCHAR x;
						x = *p + 2;
						dwStyle = (x == '}') ? CC_GROUP : ((x == '>') ? CC_RADIO : CC_CHECK);
						ptzClass = TEXT("BUTTON");
						bCheck = UStrToInt(p + 1);
						while (*p && (*p++ != x));
						break;

					case '$':
						p++;
						dwStyle = WS_CHILD | WS_VISIBLE | SS_LEFT;
						ptzClass = TEXT("STATIC");
						break;

					default:
						if (s_uCount == 0)
						{
							SetWindowText(hWnd, p);
							p = q + 1;
							continue;
						}
						dwStyle = WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | WS_BORDER;
						ptzClass = TEXT("EDIT");
						break;
					}
					

					HWND hCtrl = CreateWindow(ptzClass, p, dwStyle, 10, 5 + s_uCount * u, 280, u, hWnd, (HMENU) (UINT_PTR) (CC_CTL + s_uCount), g_hInst, 0);
					SendMessage(hCtrl, WM_SETFONT, SendMessage(hWnd, WM_GETFONT, 0, 0), FALSE);
					if (bCheck)
					{
						if (dwStyle == CC_CHECK)
						{
							CheckDlgButton(hWnd, (CC_CTL + s_uCount), BST_CHECKED);
						}
						else
						{
							CheckRadioButton(hWnd, dwGroup, (CC_CTL + s_uCount), (CC_CTL + s_uCount));
						}
					}
					s_uCount++;

					p = q + 1;
					if (t == 0)
					{
						break;
					}
				}
			}
		}
		return TRUE;

	case  WM_SIZE:
		for (UINT i = 0; i < s_uCount; i++)
		{
			RECT rt;
			HWND hCtrl = GetDlgItem(hWnd, CC_CTL + i);
			GetWindowRect(hCtrl, &rt);
			MapWindowPoints(NULL, hWnd, (PPOINT) &rt, 2);
			UINT uWidth = LOWORD(lParam) - (rt.left + PIX_PAD);
			MoveWindow(hCtrl, rt.left, rt.top, uWidth, _RectHeight(rt), TRUE);
		}
		break;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK)
		{
			UINT j = 0;
			UINT k = 0;
			for (UINT i = 0; i < s_uCount; i++)
			{
				if (GetWindowLong(GetDlgItem(hWnd, CC_CTL + i), GWL_STYLE) & WS_BORDER)
				{
					if (k < _NumOf(g_tzXVar))
					{
						GetDlgItemText(hWnd, CC_CTL + i, g_tzXVar[k++], MAX_PATH);
					}
				}
				else if (IsDlgButtonChecked(hWnd, (CC_CTL + i)))
				{
					j |= (1 << i);
				}
			}
			return EndDialog(hWnd, j);
		}
		else if (LOWORD(wParam) == IDCANCEL)
		{
			return EndDialog(hWnd, -1);
		}
		break;
	}

	return FALSE;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// LINK shortcut to target
#include <ShlObj.h>
HRESULT LINK(PTSTR ptzCmd)
{
	// Parse Shortcut,Target,Param,IconPath,IconIndex
	PTSTR ptzTarget = UStrChr(ptzCmd, CC_SEP);
	if (ptzTarget == NULL)
	{
		return ERROR_PATH_NOT_FOUND;
	}

	INT iIcon = 0;
	PTSTR ptzIcon = NULL;

	*ptzTarget++ = 0;
	PTSTR ptzParam = UStrChr(ptzTarget, CC_SEP);
	if (ptzParam)
	{
		*ptzParam++ = 0;
		ptzIcon = UStrChr(ptzParam, CC_SEP);
		if (ptzIcon)
		{
			*ptzIcon++ = 0;
			PTSTR ptzIndex = UStrChr(ptzIcon, CC_SEP);
			if (ptzIndex)
			{
				*ptzIndex++ = 0;
				iIcon = UStrToInt(ptzIndex);
			}
		}
	}

	// Search target
	if (*ptzCmd == '*')
	{
		ptzCmd++;
	}
	else
	{
		TCHAR tzTarget[MAX_PATH];
		if (SearchPath(NULL, ptzTarget, NULL, MAX_PATH, tzTarget, NULL))
		{
			ptzTarget = tzTarget;
		}
		else if (!UDirExist(ptzTarget))
		{
			return ERROR_PATH_NOT_FOUND;
		}
	}

	// Create shortcut
	IShellLink *pLink;
	CoInitialize(NULL);
	HRESULT hResult = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (PVOID *) &pLink);
	if (hResult == S_OK)
	{
		IPersistFile *pFile;
		hResult = pLink->QueryInterface(IID_IPersistFile, (PVOID *) &pFile);
		if (hResult == S_OK)
		{
			if (*ptzCmd == '!')
			{
				ptzCmd++;
				hResult = pLink->SetShowCmd(SW_SHOWMINIMIZED);
			}

			// Shortcut settings
			hResult = pLink->SetPath(ptzTarget);
			hResult = pLink->SetArguments(ptzParam);
			hResult = pLink->SetIconLocation(ptzIcon, iIcon);
			if (UPathSplit(ptzTarget) != ptzTarget)
			{
				hResult = pLink->SetWorkingDirectory(ptzTarget);
			}

			// Save link
			WCHAR wzLink[MAX_PATH];
			if (UStrCmpI(ptzCmd + UStrLen(ptzCmd) - 4, TEXT(".LNK")))
			{
				UStrCat(ptzCmd, TEXT(".LNK"));
			}
			UStrToWStr(wzLink, ptzCmd, MAX_PATH);
			UDirCreate(ptzCmd);
			hResult = pFile->Save(wzLink, FALSE);

			pFile->Release();
		}
		pLink->Release();
	}

	CoUninitialize();
	return hResult;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Copy directory
BOOL CopyDir(PCTSTR ptzSrc, PCTSTR ptzDst)
{
	WIN32_FIND_DATA fd;
	TCHAR tzDst[MAX_PATH];
	TCHAR tzSrc[MAX_PATH];
	TCHAR tzFind[MAX_PATH];
	UStrPrint(tzFind, TEXT("%s\\*"), ptzSrc);
	HANDLE hFind = FindFirstFile(tzFind, &fd);
	if (hFind != INVALID_HANDLE_VALUE)
	{
		do
		{
			if (fd.cFileName[0] != '.')
			{
				UStrPrint(tzDst, TEXT("%s\\%s"), ptzDst, fd.cFileName);
				UStrPrint(tzSrc, TEXT("%s\\%s"), ptzSrc, fd.cFileName);
				if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
				{
					CopyDir(tzSrc, tzDst);
				}
				else
				{
					UDirCreate(tzDst);
					CopyFile(tzSrc, tzDst, FALSE);
				}
			}
		}
		while (FindNextFile(hFind, &fd));
		FindClose(hFind);
	}
	UStrPrint(tzDst, TEXT("%s\\"), ptzDst);
	UDirCreate(tzDst);
	return TRUE;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// FILE or directory operation
HRESULT FILe(PTSTR ptzCmd, UINT uLen = -1)
{
	SHFILEOPSTRUCT so = {0};
	so.pFrom = ptzCmd;
	so.wFunc = FO_DELETE;
	so.fFlags = FOF_NOCONFIRMATION | FOF_NOCONFIRMMKDIR | FOF_SILENT;

	BOOL bAppend;
	PTSTR p = ptzCmd;
	for (; *p; p++)
	{
		switch (*p)
		{
		case ';':
			*p = 0;
			break;

		case '\\':
			if (!p[1])
			{
				return !UDirCreate(ptzCmd);
			}
			break;

		case '>':
			if (p[-1] == '=')
			{
				p[-1] = 0;
				so.pTo = p + 1;
				so.wFunc = FO_COPY;
			}
			else if (p[-1] == '-')
			{
				p[-1] = 0;
				so.pTo = p + 1;
				so.wFunc = FO_MOVE;
			}
			*p = 0;
			break;

		case '<':
			*p++ = 0;
			bAppend = (*p++ == '=');
			uLen = ((uLen == -1) ? UStrLen(p) : (uLen - (UINT) (p - ptzCmd)));
#ifdef _UNICODE
			CHAR szFile[MAX_STR];
			uLen = WideCharToMultiByte(CP_ACP, 0, p, uLen, szFile, sizeof(szFile), NULL, NULL);
			p = (PTSTR) szFile;
#endif
			return !UFileSave(ptzCmd, p, uLen, bAppend);

		case '{':
			*p++ = 0;
			bAppend = (*p++ == '=');
			uLen = ((uLen == -1) ? UStrLen(p) : (uLen - (UINT) (p - ptzCmd)));
#ifndef _UNICODE
			WCHAR wzFile[MAX_STR];
			uLen = MultiByteToWideChar(CP_ACP, 0, p, uLen, wzFile, _NumOf(wzFile));
			p = (PTSTR) wzFile;
#endif
			return !UFileSave(ptzCmd, p, uLen * sizeof(WCHAR), bAppend);
		}
	}

#ifdef _COPY
	if ((so.wFunc == FO_COPY) && UDirExist(so.pFrom))
	{
		return !CopyDir(so.pFrom, so.pTo);
	}
#endif

	TCHAR t = p[1];
	p[1] = 0;
	HRESULT hResult = SHFileOperation(&so);
	p[1] = t;
	return hResult;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// REGistry eXecution
HRESULT REGX(PTSTR ptzCmd)
{
	HKEY hRoot = GetRootKey(ptzCmd);
	if (!hRoot)
	{
		return E_INVALIDARG;
	}

	PTSTR ptzVal = UStrChr(ptzCmd, '=');
	if (ptzVal)
	{
		*ptzVal++ = 0;
	}

	PTSTR ptzName = UStrRChr(ptzCmd, '\\');
	if (!ptzName)
	{
		return E_INVALIDARG;
	}
	else
	{
		*ptzName++ = 0;
	}

	HKEY hKey;
	HRESULT hResult = RegCreateKeyEx(hRoot, ptzCmd, 0, NULL, 0, KEY_ALL_ACCESS, NULL, &hKey, NULL);
	if (hResult != S_OK)
	{
		return hResult;
	}

	if (ptzVal)
	{
		if (*ptzName == '#')
		{
			DWORD dwData = UStrToInt(ptzVal);
			hResult = RegSetValueEx(hKey, ptzName + 1, 0, REG_DWORD, (PBYTE) &dwData, sizeof(DWORD));
		}
		else if (*ptzName == '@')
		{
			UINT i = 0;
			PBYTE pbVal = (PBYTE) ptzVal;
			while (*ptzVal)
			{
				pbVal[i++] = (UChrToHex(ptzVal[0]) << 4) | UChrToHex(ptzVal[1]);
				while (*ptzVal && (*ptzVal++ != ','));
			}
			hResult = RegSetValueEx(hKey, ptzName + 1, 0, REG_BINARY, pbVal, i);
		}
		else
		{
			hResult = RegSetValueEx(hKey, ptzName, 0, REG_SZ, (PBYTE) ptzVal, (UStrLen(ptzVal) + 1) * sizeof(TCHAR));
		}
	}
	else
	{
		if (*ptzName == '-')
		{
			if (ptzName[1])
			{
				hResult = RegDeleteValue(hKey, ptzName + 1);
			}
			else
			{
				RegCloseKey(hKey);
				return SHDeleteKey(hRoot, ptzCmd);
			}
		}
		else if (*ptzName == '#')
		{
			DWORD dwSize = sizeof(hResult);
			RegQueryValueEx(hKey, ptzName + 1, NULL, NULL, (PBYTE) &hResult, &dwSize);
		}
		else
		{
			g_tzXVar[0][0] = 0;
			DWORD dwSize = sizeof(g_tzXVar[0]);
			hResult = RegQueryValueEx(hKey, ptzName, NULL, NULL, (PBYTE) g_tzXVar[0], &dwSize);
		}
	}

	RegCloseKey(hKey);
	return hResult;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// SEND keystroke
HRESULT SEND(PCTSTR ptzCmd)
{
	INT i = 0;
	PCTSTR p = ptzCmd;
	do
	{
		if ((*p == CC_SEP) || (*p == 0))
		{
			i = UStrToInt(ptzCmd);
			if (*(p - 1) != '^')
			{
				keybd_event(i, 0, 0, 0);
			}
			if (*(p - 1) != '_')
			{
				keybd_event(i, 0, KEYEVENTF_KEYUP, 0);
			}
			ptzCmd = p + 1;
		}
	}
	while (*p++);
	return i ? S_OK : S_FALSE;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// KILL process
#include <TLHelp32.h>
HRESULT KILL(PCTSTR ptzCmd)
{
	HRESULT hResult = S_FALSE;
	HANDLE hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
	if (hSnap != INVALID_HANDLE_VALUE)
	{
		PROCESSENTRY32 pe;
		pe.dwSize = sizeof(PROCESSENTRY32);
		for (BOOL b = Process32First(hSnap, &pe); b; b = Process32Next(hSnap, &pe))
		{
			if (!ptzCmd[UStrMatchI(ptzCmd, pe.szExeFile)])
			{
				HANDLE hProcess = OpenProcess(PROCESS_TERMINATE, FALSE, pe.th32ProcessID);
				if (hProcess)
				{
					if (TerminateProcess(hProcess, 0))
					{
						hResult = S_OK;
					}
					CloseHandle(hProcess);
				}
			}
		}
		CloseHandle(hSnap);
	}

	return hResult;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// SHUTdown
HRESULT SHUT(PCTSTR ptzCmd)
{
	Priv(SE_SHUTDOWN_NAME);
	BOOL bReboot = ((*ptzCmd) == 'R') || ((*ptzCmd) == 'r');
	if (ExitWindowsEx(bReboot ? EWX_REBOOT : EWX_POWEROFF, 0))
	{
		return S_OK;
	}

	// End session
	DWORD dwResult;
	SendMessageTimeout(HWND_BROADCAST, WM_QUERYENDSESSION, 0, 0, 0, 2000, &dwResult);
	SendMessageTimeout(HWND_BROADCAST, WM_ENDSESSION, 0, 0, 0, 2000, &dwResult);
	//SendMessageTimeout(HWND_BROADCAST, WM_CLOSE, 0, 0, 0, 2000, &dwResult);
	SendMessageTimeout(HWND_BROADCAST, WM_DESTROY, 0, 0, 0, 2000, &dwResult);

	// Get function address
	typedef DWORD (NTAPI *PNtShutdownSystem)(DWORD dwAction);
	PNtShutdownSystem NtShutdownSystem = (PNtShutdownSystem) GetProcAddress(GetModuleHandle(TEXT("NTDLL")), "NtShutdownSystem");
	if (!NtShutdownSystem)
	{
		return E_FAIL;
	}

	// Shutdown
	return NtShutdownSystem(bReboot ? 1: 2);
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// EXECute executable file
HRESULT EXEC(PTSTR ptzCmd)
{
	STARTUPINFO si = {0};
	PROCESS_INFORMATION pi;
	si.cb = sizeof(STARTUPINFO);
	si.lpDesktop = TEXT("WinSta0\\Default");

	BOOL bWait = FALSE;
	BOOL bRunKey = FALSE;
	while (TRUE)
	{
		if (*ptzCmd == '!') si.dwFlags = STARTF_USESHOWWINDOW;
		else if (*ptzCmd == '@') si.lpDesktop = TEXT("WinSta0\\WinLogon");
		else if (*ptzCmd == '=') bWait = TRUE;
		else if (*ptzCmd == '&') bRunKey = TRUE;
		else break;
		ptzCmd++;
	}

	if (bRunKey)
	{
		PTSTR ptzName = UStrRChr(ptzCmd, '\\');
		ptzName = ptzName ? (ptzName + 1) : ptzCmd;
		HKEY hKey = bWait ? HKEY_LOCAL_MACHINE : HKEY_CURRENT_USER;
		return SHSetValue(hKey, TEXT("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run"), ptzName, REG_SZ, ptzCmd, UStrLen(ptzCmd) * sizeof(TCHAR));
	}

	TCHAR tzCurDir[MAX_PATH];
	UStrCopy(tzCurDir, ptzCmd);
	PTSTR ptzEnd;
	if (tzCurDir[0] == '"')
	{
		ptzEnd = UStrChr(tzCurDir + 1, '"');
	}
	else if ((ptzEnd = UStrChr(tzCurDir + 1, ' ')) == NULL)
	{
		ptzEnd = tzCurDir + UStrLen(tzCurDir);
	}
	while ((ptzEnd >= tzCurDir) && (*ptzEnd != '\\')) ptzEnd--;
	*ptzEnd = 0;

	BOOL bResult = CreateProcess(NULL, ptzCmd, NULL, NULL, FALSE, 0, NULL, tzCurDir, &si, &pi);
	if (bResult)
	{
		if (bWait)
		{
			WaitForSingleObject(pi.hProcess, INFINITE);
		}
		CloseHandle(pi.hThread);
		CloseHandle(pi.hProcess);
	}

	return bResult ? S_OK : S_FALSE;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Call DLL function
typedef HRESULT (WINAPI *PROC1)(PVOID pv0);
typedef HRESULT (WINAPI *PROC2)(PVOID pv0, PVOID pv1);
typedef HRESULT (WINAPI *PROC3)(PVOID pv0, PVOID pv1, PVOID pv2);
typedef HRESULT (WINAPI *PROC4)(PVOID pv0, PVOID pv1, PVOID pv2, PVOID pv3);
HRESULT CDLL(PTSTR ptzCmd)
{
	UINT uArg = 0;
	PTSTR ptzArg[4];
	PTSTR ptzProc = UStrChr(ptzCmd, CC_SEP);
	if (ptzProc)
	{
		*ptzProc++ = 0;
		for (PTSTR p = ptzProc; (uArg < 4) && (p = UStrChr(p, CC_SEP)); uArg++)
		{
			*p++ = 0;
			ptzArg[uArg] = p;
			if (*p == '#')
			{
				ptzArg[uArg] = (PTSTR) (INT_PTR) UStrToInt(p + 1);
			}
		}
	}
	else
	{
		ptzProc = TEXT("DllRegisterServer");
	}

	// Prepare
	TCHAR tzOrg[MAX_PATH];
	TCHAR tzCur[MAX_PATH];
	UDirGetCurrent(tzOrg);
	UStrCopy(tzCur, ptzCmd);
	UPathSplit(tzCur);
	SetCurrentDirectory(tzCur);
	CoInitialize(NULL);

	HRESULT hResult = E_NOINTERFACE;
	HMODULE hLib = LoadLibrary(ptzCmd);
	if (hLib)
	{
		CHAR szProc[MAX_NAME];
		UStrToAStr(szProc, ptzProc, MAX_NAME);
		PROC f = GetProcAddress(hLib, szProc);
		if (f)
		{
			static DWORD s_dwESP;
			__asm MOV s_dwESP, ESP;
			switch (uArg)
			{
			case 0: hResult = f(); break;
			case 1: hResult = ((PROC1) f)(ptzArg[0]); break;
			case 2: hResult = ((PROC2) f)(ptzArg[0], ptzArg[1]); break;
			case 3: hResult = ((PROC3) f)(ptzArg[0], ptzArg[1], ptzArg[2]); break;
			case 4: hResult = ((PROC4) f)(ptzArg[0], ptzArg[1], ptzArg[2], ptzArg[3]); break;
			}
			__asm MOV ESP, s_dwESP;
		}
		FreeLibrary(hLib);
	}

	// Restore
	CoUninitialize();
	SetCurrentDirectory(tzOrg);

	return hResult;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Associate file
#include "Reg.h"
HRESULT ASOC(PCTSTR ptzCmd)
{
	if (*ptzCmd == '!')
	{
		return  SHDeleteKey(HKEY_CLASSES_ROOT, (ptzCmd[1] ? &ptzCmd[1] : STR_AppName));
	}

	if (*ptzCmd)
	{
		CReg reg1(ptzCmd, HKEY_CLASSES_ROOT);
		reg1.SetStr(NULL, STR_AppName);
	}

	CReg reg2(STR_AppName, HKEY_CLASSES_ROOT);
	reg2.SetStr(NULL, STR_AppName);

	TCHAR tzPath[MAX_PATH];
	PTSTR p = tzPath;
	*p++ = '"';
	p += GetModuleFileName(NULL, p, MAX_PATH);
	CReg reg3(STR_AppName TEXT("\\DefaultIcon"), HKEY_CLASSES_ROOT);
	UStrCopy(p, TEXT(",1"));
	reg3.SetStr(NULL, tzPath + 1);

	UStrCopy(p, TEXT("\" %1"));
	CReg reg4(STR_AppName TEXT("\\Shell\\Open\\Command"), HKEY_CLASSES_ROOT);
	reg4.SetStr(NULL, tzPath);
	return S_OK;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// OPEN file or shell command
#include <SetupAPI.h>
HRESULT OPEN(PCTSTR ptzDir, PCTSTR ptzFile, BOOL bSubDir = FALSE)
{
	HRESULT hResult = ERROR_FILE_NOT_FOUND;

	// Lookup
	WIN32_FIND_DATA fd;
	TCHAR tzPath[MAX_PATH];
	UStrPrint(tzPath, TEXT("%s\\%s"), ptzDir, ptzFile);
	HANDLE hFind = FindFirstFile(tzPath, &fd);
	if (hFind != INVALID_HANDLE_VALUE)
	{
		do
		{
			if (!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
			{
				UStrPrint(tzPath, TEXT("%s\\%s"), ptzDir, fd.cFileName);

				// Take action
				PTSTR ptzExt = UStrRChr(fd.cFileName, '.');
				if (ptzExt)
				{
					if ((UStrCmpI(ptzExt, TEXT(".csc")) == 0) || (UStrCmpI(ptzExt, TEXT(".reg")) == 0))
					{
						hResult = LOAD(tzPath);
					}
					else if ((UStrCmpI(ptzExt, TEXT(".dll")) == 0) || (UStrCmpI(ptzExt, TEXT(".ocx")) == 0) || (UStrCmpI(ptzExt, TEXT(".ax")) == 0))
					{
						hResult = CDLL(tzPath);
					}
					else if (UStrCmpI(ptzExt, TEXT(".inf")) == 0)
					{
						TCHAR tzCmd[MAX_PATH];
						UStrPrint(tzCmd, TEXT("DefaultInstall 132 %s"), tzPath);
						InstallHinfSection(NULL, NULL, tzCmd, 0);
						hResult = S_OK;
					}
					else
					{
						// Pass to shell to execute it
						hResult = !ShellOpen(tzPath, NULL, (fd.cFileName[0] == '!') ? SW_HIDE : SW_NORMAL);
					}
				}		
			}
		}
		while (FindNextFile(hFind, &fd));
		FindClose(hFind);
	}

	if (bSubDir)
	{
		UStrPrint(tzPath, TEXT("%s\\*"), ptzDir);
		hFind = FindFirstFile(tzPath, &fd);
		if (hFind != INVALID_HANDLE_VALUE)
		{
			do
			{
				if ((fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) && (fd.cFileName[0] != '.'))
				{
					UStrPrint(tzPath, TEXT("%s\\%s"), ptzDir, fd.cFileName);
					if (hResult == ERROR_FILE_NOT_FOUND)
					{
						hResult = OPEN(tzPath, ptzFile, bSubDir);
					}
					else
					{
						OPEN(tzPath, ptzFile, bSubDir);
					}
				}
			}
			while (FindNextFile(hFind, &fd));
			FindClose(hFind);
		}
	}
	return hResult;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Dispatch command
//#pragma comment(lib, "ComCtl32.lib")
HRESULT Dispatch(PTSTR ptzCmd, CXT& XT)
{
	// Get command ID
	UINT uCmd;
	PTSTR p = ptzCmd;
	for (uCmd = 0; uCmd < _NumOf(c_tzCmd); uCmd++)
	{
		if (UStrMatchI(p, c_tzCmd[uCmd]) >= CC_LEN)
		{
			// skip white space
			for (p += CC_LEN; (*p == ' ') || (*p == '\t'); p++);
			break;
		}
	}

	switch (uCmd)
	{
	case CMD_LOAD:
		return LOAD(p);

	case CMD_BATC:
		return BATC(p, XT);

	case CMD_IFEX:
		if (*p)
		{
			if (PTSTR ptzArg = UStrChr(p, CC_SEP))
			{
				*ptzArg++ = 0;
				if (IFEX(p))
				{
					Dispatch(ptzArg, XT);
				}
				return XT.hXVar;
			}
			else if (!IFEX(p))
			{
				if (p = UStrStr(XT.ptzNext, TEXT("IFEX\r\n")))
				{
					XT.ptzNext = p + CC_LEN + 2;
					return S_OK;
				}
				else if (p = UStrStr(XT.ptzNext, TEXT("IFEX\n")))
				{
					XT.ptzNext = p + CC_LEN + 1;
					return S_OK;
				}
				else
				{
					XT.ptzNext = TEXT("");
					return S_FALSE;
				}
			}
		}
		return S_OK;

	case CMD_ELSE:
		if (!g_bResult) Dispatch(p, XT);
		return XT.hXVar;

	case CMD_EVAL:
		return EVAL(p);

	case CMD_LINK:
		return LINK(p);

	case CMD_FILE:
		return FILe(p);

	case CMD_REGX:
		return REGX(p);

	case CMD_ENVI:
		return (*p == '$') ? ENVI(p + 1, TRUE) : ENVI(p);

	case CMD_SEND:
		return SEND(p);

	case CMD_WAIT:
		Sleep(UStrToInt(p));
		return S_OK;

	case CMD_KILL:
		return KILL(p);

	case CMD_SHUT:
		return SHUT(p);

	case CMD_PLAY:
		return PLAY(p);

	case CMD_BEEP:
		return !MessageBeep(UStrToInt(p));

	case CMD_MSGX:
		return MSGX(p);

	case CMD_DLGX:
		return (HRESULT) DialogBoxParam(g_hInst, MAKEINTRESOURCE(IDD_Dialog), NULL, DLGX, (LPARAM) p);

	case CMD_EXEC:
		return EXEC(p);

	case CMD_CDLL:
		return CDLL(p);

	case CMD_CCUI:
		return (HRESULT) DialogBoxParam(g_hInst, MAKEINTRESOURCE(IDD_Main), NULL, MAIN, (LPARAM) XT.ptzFile);

	case CMD_CALL:
	case CMD_GOTO:
	case CMD_PROC:
		UMemCopy(ptzCmd, c_tzCmd[(uCmd == CMD_PROC) ? CMD_ENDP : CMD_PROC], CC_LEN * sizeof(TCHAR));
		if (p = UStrStr(XT.ptzNext, ptzCmd))
		{
			p += CC_LEN;
			while (*p && (*p++ != '\n'));
			if (uCmd == CMD_CALL)
			{
				return Process(p, XT.ptzFile);
			}
			else
			{
				XT.ptzNext = p;
			}
			return S_OK;
		}
		else if (uCmd == CMD_PROC)
		{
			XT.ptzNext = TEXT("");
		}
		return S_FALSE;

	case CMD_ENDP:
		XT.ptzNext = TEXT("");
		return S_OK;

	case CMD_ASOC:
		return ASOC(p);

	case CMD_DEVI:
		return SERV(p);
		
	case CMD_SERV:
		return SERV(p);

	case CMD_PAGE:
		return PAGE(p);

	case CMD_DISP:
		return DISP(p);

	default:
		PTSTR ptzFile;
		if (ptzFile = UStrRChr(p, '\\'))
		{
			*ptzFile++ = 0;
		}
		return (*p == '!') ? OPEN(p + 1, ptzFile, TRUE) : OPEN(p, ptzFile);
	}
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Expand macro
#include <ShlObj.h>
PTSTR Expand(PTSTR ptzDst, TCHAR tMacro, CXT& XT)
{
	INT i;
	switch (tMacro)
	{
	case 'E':
		for (PCTSTR p = XT.ptzFile; *p; *ptzDst++ = *p++);
		return ptzDst;

	case 'C':
		if (PCTSTR q = UStrRChr(XT.ptzFile, '\\'))
		{
			for (PCTSTR p = XT.ptzFile; p < q; *ptzDst++ = *p++);
		}
		return ptzDst;

	case 'T':
		ptzDst += GetTimeFormat(LOCALE_USER_DEFAULT, 0, NULL, NULL, ptzDst, MAX_NAME) - 1;
		return ptzDst;

	case 'D':
		ptzDst += GetDateFormat(LOCALE_USER_DEFAULT, 0, NULL, NULL, ptzDst, MAX_NAME) - 1;
		return ptzDst;

	case 'X':
	case 'x':
		ptzDst += UStrPrint(ptzDst,  (tMacro == 'X') ? TEXT("%#X") : TEXT("%d"), XT.hXVar);
		return ptzDst;

	case '0':
	case '1':
	case '2':
	case '3':
	case '4':
	case '5':
	case '6':
	case '7':
	case '8':
	case '9':
		for (PTSTR p = g_tzXVar[tMacro - '0']; *p; p++) *ptzDst++ = *p;
		return ptzDst;

	case 'A': i = CSIDL_APPDATA; break;
	case 'Y': i = CSIDL_MYDOCUMENTS; break;
	case 'S': i = CSIDL_STARTUP; break;
	case 'M': i = CSIDL_STARTMENU; break;
	case 'P': i = CSIDL_PROGRAMS; break;
	case 'V': i = CSIDL_FAVORITES; break;
	case 'Z': i = CSIDL_MYPICTURES; break;
	case 'U': i = CSIDL_MYMUSIC; break;
	case 'I': i = CSIDL_MYVIDEO; break;
	case 'F': i = CSIDL_PROGRAM_FILES; break;
	case 'O': i = CSIDL_SENDTO; break;
	case 'o': i = CSIDL_DESKTOPDIRECTORY; break;

	case 'd': i = CSIDL_COMMON_DESKTOPDIRECTORY; break;
	case 'a': i = CSIDL_COMMON_APPDATA; break;
	case 'y': i = CSIDL_COMMON_DOCUMENTS; break;
	case 's': i = CSIDL_COMMON_STARTUP; break;
	case 'm': i = CSIDL_COMMON_STARTMENU; break;
	case 'p': i = CSIDL_COMMON_PROGRAMS; break;
	case 'v': i = CSIDL_COMMON_FAVORITES; break;
	case 'z': i = CSIDL_COMMON_PICTURES; break;
	case 'u': i = CSIDL_COMMON_MUSIC; break;
	case 'i': i = CSIDL_COMMON_VIDEO; break;
	case 'f': i = CSIDL_PROGRAM_FILES_COMMON; break;

	case 'W': case 'w': i = CSIDL_WINDOWS; break;

	case 'R': *ptzDst++ = '\r'; return ptzDst;
	case 'N': *ptzDst++ = '\n'; return ptzDst;
	default: *ptzDst++ = tMacro; return ptzDst;
	}

	SHGetSpecialFolderPath(NULL, ptzDst, i, TRUE);
	ptzDst += UStrLen(ptzDst);
	return ptzDst;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Execute command
HRESULT Execute(PTSTR ptzCmd, CXT& XT)
{
	while ((*ptzCmd == ' ') || (*ptzCmd == '\t')) ptzCmd++;
	if (*ptzCmd == '[')
	{
		XT.hXVar = Import((PTSTR) XT.ptzCur, &XT.ptzNext);
	}
	else if (*ptzCmd && (*ptzCmd != ';'))
	{
		if (g_hWnd)
		{
			Log(ptzCmd);
			Log(TEXT("\r\n"));
			XT.hXVar = Dispatch(ptzCmd, XT);

			TCHAR tzStr[MAX_PATH];;
			UINT i = UStrPrint(tzStr, TEXT("     %08X "), XT.hXVar);
			i += FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, NULL, XT.hXVar, 0, tzStr + i, _NumOf(tzStr), NULL);
			UStrCopy(tzStr + i, (tzStr[i - 1] == '\n') ? TEXT("\r\n") : TEXT("\r\n\r\n"));

			Log(tzStr);
		}
		else
		{
			XT.hXVar = Dispatch(ptzCmd, XT);
		}
	}
	return XT.hXVar;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Process script
HRESULT Process(PCTSTR ptzScript, PCTSTR ptzFile)
{
	TCHAR tzCmd[MAX_STR];
	PTSTR ptzDst = tzCmd;
	CXT XT = {ptzScript, ptzScript, ptzFile};
	while (TRUE)
	{
		TCHAR t = *XT.ptzNext++;
		switch (t)
		{
		case '\r':
		case '\n':
			*ptzDst = 0;
			Execute(tzCmd, XT);
			ptzDst = tzCmd;
			XT.ptzCur = XT.ptzNext;
			break;

		case 0:
			*ptzDst = 0;
			return Execute(tzCmd, XT);

		case ',':
			*ptzDst++ = CC_SEP;
			break;

		case '%':
			t = *XT.ptzNext++;
			if (t == '#')
			{
				UINT j = 0;
				for (INT i = 2 * sizeof(TCHAR) - 1; i >= 0; i--)
				{
					j += (UChrToHex(*XT.ptzNext++) << (4 * i));
				}
				*ptzDst++ = j;
			}
			else
			{
				ptzDst = Expand(ptzDst, t, XT);
			}
			break;

		default:
			*ptzDst++ = t;
			break;
		}
	}
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
