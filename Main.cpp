


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Header
#include "Main.h"
#include "Reg.h"

#pragma comment(lib, "ShLwApi.lib")

HWND g_hWnd = NULL;
HINSTANCE g_hInst = NULL;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Browse file path
BOOL BrowseFile(HWND hWnd, UINT uCtrl, UINT uFilter, BOOL bSave = FALSE)
{
	TCHAR tzPath[MAX_PATH];
	GetDlgItemText(hWnd, uCtrl, tzPath, MAX_PATH);

	TCHAR tzFilter[MAX_PATH];
	_LoadStr(uFilter, tzFilter);
	UStrRep(tzFilter, '|', 0);

	OPENFILENAME ofn = {0};
	ofn.hwndOwner = hWnd;
	ofn.hInstance = g_hInst;
	ofn.lpstrFile = tzPath;
	ofn.lpstrFilter = tzFilter;
	ofn.nMaxFile = MAX_PATH;
	ofn.lStructSize = sizeof(OPENFILENAME);

	BOOL bRet;
	if (bSave)
	{
		ofn.Flags = OFN_ENABLESIZING | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
		bRet = GetSaveFileName(&ofn);
	}
	else
	{
		ofn.Flags = OFN_ENABLESIZING | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY | OFN_FILEMUSTEXIST;
		bRet = GetOpenFileName(&ofn);
	}

	if (bRet)
	{
		SetDlgItemText(hWnd, uCtrl, tzPath);
	}

	return bRet;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// MAIN dialog
#pragma warning(disable:4244)
INT_PTR CALLBACK MAIN(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	TCHAR tzStr[MAX_STR];
	STATIC PCTSTR s_ptzCurFile = NULL;

	switch (uMsg)
	{
	case WM_INITDIALOG:
		g_hWnd = hWnd;
		s_ptzCurFile = (PCTSTR) lParam;
		MAIN(hWnd, WM_COMMAND, MAKELONG(IDC_Log, EN_KILLFOCUS), 0);
		SetClassLongPtr(hWnd, GCL_HICON, (LONG_PTR) LoadIcon(GetModuleHandle(NULL), _MakeIntRes(IDI_Main)));
		if (HFONT hFont = (HFONT) SendMessage(hWnd, WM_GETFONT, 0, 0))
		{
			LOGFONT lf;
			GetObject(hFont, sizeof(LOGFONT), &lf);
			lf.lfHeight *= 2;
			lf.lfItalic = TRUE;
			hFont = CreateFontIndirect(&lf);
			SendDlgItemMessage(hWnd, IDC_Brand, WM_SETFONT, (WPARAM) hFont, 0);
		}
		return TRUE;

	case WM_DROPFILES:
		DragQueryFile((HDROP) wParam, 0, tzStr, MAX_PATH);
		DragFinish((HDROP) wParam);
		SetDlgItemText(hWnd, IDC_Path, tzStr);
		break;

	case WM_SYSCOMMAND:
		if (LOWORD(wParam) == SC_CONTEXTHELP)
		{
			TCHAR tzHelp[MAX_PATH];
			TCHAR tzPath[MAX_PATH];
			UDirGetAppPath(tzPath);
			UStrPrint(tzHelp, TEXT("res://%s/HELP"), tzPath);
			ShellOpen(tzHelp);
			return TRUE;
		}
		break;

	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDC_Browse:
			BrowseFile(hWnd, IDC_Path, IDS_Filter);
			break;

		case IDC_Path:
			if (HIWORD(wParam) == EN_CHANGE)
			{
				EnableWindow(GetDlgItem(hWnd, IDOK), GetWindowTextLength(GetDlgItem(hWnd, IDC_Path)));
			}
			break;

		case IDOK:
			if (HIWORD(wParam) == 0)
			{
				SetDlgItemText(g_hWnd, IDC_Log, NULL);
				SetFocus(GetDlgItem(g_hWnd, IDC_Log));
				GetDlgItemText(hWnd, IDC_Path, tzStr, MAX_PATH);
				Process(tzStr, s_ptzCurFile);
				break;
			}

		case IDCANCEL:
			EndDialog(hWnd, S_OK);
			break;

		case IDC_About:
			MessageBox(hWnd, STR_AppName TEXT(" ") STR_VersionStamp TEXT("\r\n\r\n") STR_BuildStamp TEXT("\r\n\r\n") STR_Copyright, STR_AppName, MB_ICONINFORMATION);
			break;

		case IDC_Log:
			if ((HIWORD(wParam) == EN_KILLFOCUS) || (HIWORD(wParam) == EN_SETFOCUS))
			{
				static INT h = 0;
				static RECT rt = {0};
				if (h == 0)
				{
					GetWindowRect(GetDlgItem(hWnd, IDC_Log), &rt);
					h = rt.top;
					GetWindowRect(hWnd, &rt);
					h -= rt.top + 2;
					rt.right -= rt.left;
					rt.bottom -= rt.top;
				}
				MoveWindow(hWnd, rt.left, rt.top, rt.right, (HIWORD(wParam) == EN_KILLFOCUS) ? h : rt.bottom, TRUE);
			}
			break;

		case IDC_Logo:
			if (HIWORD(wParam) == STN_CLICKED)
			{
				SetDlgItemText(hWnd, IDC_Log, NULL);
			}
			else if (HIWORD(wParam) == STN_DBLCLK)
			{
				HKEY hKey;
				TCHAR tzStr[MAX_PATH];
				if (RegOpenKeyEx(HKEY_CLASSES_ROOT, STR_AppName, 0, KEY_ALL_ACCESS, &hKey) == S_OK)
				{
					RegCloseKey(hKey);
					ASOC(TEXT("!"));
					ASOC(TEXT("!.csc"));
					_GetStr(IDS_UnAssoc);
				}
				else
				{
					ASOC(TEXT(".csc"));
					_GetStr(IDS_Assoc);
				}
				MessageBox(hWnd, tzStr, STR_AppName, MB_ICONINFORMATION);
			}
			break;

		case IDC_Cmd:
			if (HIWORD(wParam) == STN_DBLCLK)
			{
				TCHAR tzPath[MAX_PATH];
				if (GetDlgItemText(hWnd, IDC_Path, tzPath, MAX_PATH))
				{
					ShellOpen(TEXT("NOTEPAD.EXE"), tzPath);
				}
			}
			break;
		}
		break;
	}

	return FALSE;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// EXE Entry
INT APIENTRY _tWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PTSTR ptzCmdLine, INT iCmdShow)
{
	TCHAR tzPath[MAX_PATH];
	GetModuleFileName(NULL, tzPath, MAX_PATH);
	g_hInst = hInstance;
	if (*ptzCmdLine)
	{
		return Process(UStrTrim(ptzCmdLine), tzPath);
	}
	else
	{
		HRSRC hRsrc = FindResource(g_hInst, STR_AppName, RT_RCDATA);
		if (hRsrc)
		{
			HGLOBAL hGlobal = LoadResource(g_hInst, hRsrc);
			if (hGlobal)
			{
				return Process((PTSTR) LockResource(hGlobal), tzPath);
			}
		}
		return Process(TEXT("CCUI"), tzPath);
	}
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
