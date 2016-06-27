


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Header
#pragma once
#define REG_AppKey				TEXT("Software\\") STR_AppName
#define REG_AppSubKey(x)		REG_AppKey TEXT("\\") TEXT(x)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// CReg class
class CReg
{
protected:
	HKEY m_hKey;

public:
	CReg(PCTSTR ptzKey = REG_AppKey, HKEY hKey = HKEY_CURRENT_USER, REGSAM regSam = KEY_ALL_ACCESS)
	{
		RegCreateKeyEx(hKey, ptzKey, 0, NULL, 0, regSam, NULL, &m_hKey, NULL);
	}

	virtual ~CReg()
	{
		RegCloseKey(m_hKey);
	}

	operator HKEY()
	{
		return m_hKey;
	}

public:
	INT GetInt(PCTSTR ptzName, INT iDef = 0)
	{
		DWORD dwSize = sizeof(DWORD);
		RegQueryValueEx(m_hKey, ptzName, NULL, NULL, (PBYTE) &iDef, &dwSize);
		return iDef;
	}

	BOOL SetInt(PCTSTR ptzName, INT iVal = 0)
	{
		return RegSetValueEx(m_hKey, ptzName, 0, REG_DWORD, (PBYTE) &iVal, sizeof(DWORD)) == ERROR_SUCCESS;
	}

	UINT GetStr(PCTSTR ptzName, PTSTR ptzStr, UINT uLen = MAX_PATH, PCTSTR ptzDef = NULL)
	{
		ptzStr[0] = 0;
		uLen = uLen * sizeof(TCHAR);
		if (RegQueryValueEx(m_hKey, ptzName, NULL, NULL, (PBYTE) ptzStr, (PDWORD) &uLen) == ERROR_SUCCESS)
		{
			return (uLen / sizeof(TCHAR)) - 1;
		}
		else if (ptzDef)
		{
			PCTSTR p = ptzDef;
			while (*ptzStr++ = *p++);
			return (UINT) (p - ptzDef);
		}
		else
		{
			return 0;
		}
	}

	BOOL SetStr(PCTSTR ptzName, PCTSTR ptzStr = TEXT(""))
	{
		return RegSetValueEx(m_hKey, ptzName, 0, REG_SZ, (PBYTE) ptzStr, (lstrlen(ptzStr) + 1) * sizeof(TCHAR)) == ERROR_SUCCESS;
	}

	BOOL GetStruct(PCTSTR ptzName, PVOID pvStruct, UINT uSize)
	{
		return RegQueryValueEx(m_hKey, ptzName, NULL, NULL, (PBYTE) pvStruct, (PDWORD) &uSize) == ERROR_SUCCESS;
	}

	BOOL SetStruct(PCTSTR ptzName, PVOID pvStruct, UINT uSize, DWORD dwType = REG_BINARY)
	{
		return RegSetValueEx(m_hKey, ptzName, 0, dwType, (PBYTE) pvStruct, uSize) == ERROR_SUCCESS;
	}

public:
	BOOL DelVal(PCTSTR ptzName)
	{
		return RegDeleteValue(m_hKey, ptzName) == ERROR_SUCCESS;
	}

	static BOOL DelKey(PCTSTR ptzKey = REG_AppKey, HKEY hKey = HKEY_CURRENT_USER)
	{
		return RegDeleteKey(hKey, ptzKey) == ERROR_SUCCESS;
	}

	BOOL EnumVal(DWORD dwIndex, PTSTR ptzName, PBYTE pbVal, PDWORD pdwSize, PDWORD pdwType = NULL)
	{
		DWORD dwSize = MAX_PATH;
		return RegEnumValue(m_hKey, dwIndex, ptzName, &dwSize, NULL, pdwType, pbVal, pdwSize) == ERROR_SUCCESS;
	}

	BOOL EnumKey(DWORD dwIndex, PTSTR ptzKeyName)
	{
		FILETIME ftTime;
		DWORD dwSize = MAX_PATH;
		return RegEnumKeyEx(m_hKey, dwIndex, ptzKeyName, &dwSize, NULL, NULL, 0, &ftTime) == ERROR_SUCCESS;
	}

public:
	UINT GetSize(PCTSTR ptzName)
	{
		UINT uSize = 0;
		RegQueryValueEx(m_hKey, ptzName, NULL, NULL, NULL, (PDWORD) &uSize);
		return uSize;
	}

	BOOL ExistVal(PCTSTR ptzName)
	{
		UINT uSize;
		return RegQueryValueEx(m_hKey, ptzName, NULL, NULL, NULL, (PDWORD) &uSize) == ERROR_SUCCESS;
	}
};
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
