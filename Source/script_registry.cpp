////////////////////////////////////////////////////////////
// This file has been adapted to work with this application.
///////////////////////////////////////////////////////////////////////////////
//
// AutoIt
//
// Copyright (C)1999-2003:
//		- Jonathan Bennett <jon@hiddensoft.com>
//		- See "AUTHORS.txt" for contributors.
//
// This program is free software; you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation; either version 2 of the License, or
// (at your option) any later version.
//
///////////////////////////////////////////////////////////////////////////////
//
// script_registry.cpp
//
// Contains registry handling routines.  Part of script.cpp
//
///////////////////////////////////////////////////////////////////////////////


// Includes
#include "StdAfx.h"								// Pre-compiled headers

#ifndef _MSC_VER								// Includes for non-MS compilers
//
#endif

#include "script.h"
#include "util.h" // for strlcpy()
#include "globaldata.h"



ResultType Line::IniRead(char *aFilespec, char *aSection, char *aKey, char *aDefault)
{
	if (!aDefault || !*aDefault)
		aDefault = "ERROR";  // This mirrors what AutoIt2 does for its default value.
	char	szFileTemp[_MAX_PATH+1];
	char	*szFilePart;
	char	szBuffer[65535] = "";					// Max ini file size is 65535 under 95
	// Get the fullpathname (ini functions need a full path):
	GetFullPathName(aFilespec, _MAX_PATH, szFileTemp, &szFilePart);
	GetPrivateProfileString(aSection, aKey, aDefault, szBuffer, sizeof(szBuffer), szFileTemp);
	// The above function is supposed to set szBuffer to be aDefault if it can't find the
	// file, section, or key.  In other words, it always changes the contents of szBuffer.
	return OUTPUT_VAR->Assign(szBuffer);
}



ResultType Line::IniWrite(char *aValue, char *aFilespec, char *aSection, char *aKey)
{
	char	szFileTemp[_MAX_PATH+1];
	char	*szFilePart;
	// Get the fullpathname (ini functions need a full path) 
	GetFullPathName(aFilespec, _MAX_PATH, szFileTemp, &szFilePart);
	WritePrivateProfileString(aSection, aKey, aValue, szFileTemp);  // Returns zero on failure.
	WritePrivateProfileString(NULL, NULL, NULL, szFileTemp);	// Flush
	return OK;  // For now, this always returns OK.
}



ResultType Line::IniDelete(char *aFilespec, char *aSection, char *aKey)
{
	char	szFileTemp[_MAX_PATH+1];
	char	*szFilePart;
	// Get the fullpathname (ini functions need a full path) 
	GetFullPathName(aFilespec, _MAX_PATH, szFileTemp, &szFilePart);
	WritePrivateProfileString(aSection, aKey, NULL, szFileTemp);  // Returns zero on failure.
	WritePrivateProfileString(NULL, NULL, NULL, szFileTemp);	// Flush
	return OK;  // For now, this always returns OK.
}



ResultType Line::RegRead(char *aValueType, char *aRegKey, char *aRegSubkey, char *aValueName)
{
	// $var = RegRead(key, valuename)

	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	OUTPUT_VAR->Assign(); // Init.  Tell it not to free the memory by not calling with "".

	HKEY	hRegKey, hMainKey;
	DWORD	dwRes, dwBuf, dwType;
	// My: Seems safest to keep the limit just below 64K in case Win95 has problems with larger values.
	char	szRegBuffer[65535];					// Only allow reading of 64Kb from a key

	// Get the main key name
	if (   !(hMainKey = RegConvertMainKey(aRegKey))   )
		return OK;  // Let ErrorLevel tell the story.

	// Open the registry key
	if ( RegOpenKeyEx(hMainKey, aRegSubkey, 0, KEY_READ, &hRegKey) != ERROR_SUCCESS )
		return OK;  // Let ErrorLevel tell the story.

	// Read the value and determine the type
	if ( RegQueryValueEx(hRegKey, aValueName, NULL, &dwType, NULL, NULL) != ERROR_SUCCESS )
		return OK;  // Let ErrorLevel tell the story.

	// The way we read is different depending on the type of the key
	switch (dwType)
	{
		case REG_SZ:
		case REG_MULTI_SZ:
		case REG_EXPAND_SZ:
			dwRes = sizeof(szRegBuffer);
			RegQueryValueEx(hRegKey, aValueName, NULL, NULL, (LPBYTE)szRegBuffer, &dwRes);
			RegCloseKey(hRegKey);
			g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
			return OUTPUT_VAR->Assign(szRegBuffer);

		case REG_DWORD:
			dwRes = sizeof(dwBuf);
			RegQueryValueEx(hRegKey, aValueName, NULL, NULL, (LPBYTE)&dwBuf, &dwRes);
			RegCloseKey(hRegKey);
			g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
			return OUTPUT_VAR->Assign((int)dwBuf);
	}

	// Since above didn't return, this is an unsupported value type.
	return OK;  // Let ErrorLevel tell the story.
} // RegRead()



ResultType Line::RegWrite(char *aValueType, char *aRegKey, char *aRegSubkey, char *aValueName, char *aValue)
{
	// $var = RegWrite(key, valuename, type, value)

	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.

	HKEY	hRegKey, hMainKey;
	DWORD	dwRes, dwBuf;
	// My: Seems safest to keep the limit just below 64K in case Win95 has problems with larger values.
	char	szRegBuffer[65535];					// Only allow writing of 64Kb to a key

	// Get the main key name
	if (   !(hMainKey = RegConvertMainKey(aRegKey))   )
		return OK;  // Let ErrorLevel tell the story.

	// Open/Create the registry key
	if (RegCreateKeyEx(hMainKey, aRegSubkey, 0, "", REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hRegKey, &dwRes)
		!= ERROR_SUCCESS)
		return OK;  // Let ErrorLevel tell the story.

	// Write the registry differently depending on type of variable we are writing
	if (!stricmp(aValueType, "REG_EXPAND_SZ"))
	{
		strlcpy(szRegBuffer, aValue, sizeof(szRegBuffer));
		if (RegSetValueEx(hRegKey, aValueName, 0, REG_EXPAND_SZ, (CONST BYTE *)szRegBuffer
			, (DWORD)strlen(szRegBuffer)+1 ) == ERROR_SUCCESS)
			g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		RegCloseKey(hRegKey);
		return OK;
	}

	if (!stricmp(aValueType, "REG_SZ"))
	{
		strlcpy(szRegBuffer, aValue, sizeof(szRegBuffer));
		if ( RegSetValueEx(hRegKey, aValueName, 0, REG_SZ, (CONST BYTE *)szRegBuffer
			, (DWORD)strlen(szRegBuffer)+1 ) == ERROR_SUCCESS )
			g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		RegCloseKey(hRegKey);
		return OK;
	}

	if (!stricmp(aValueType, "REG_DWORD"))
	{
		sscanf(aValue, "%u", &dwBuf);
		if (RegSetValueEx(hRegKey, aValueName, 0, REG_DWORD, (CONST BYTE *)&dwBuf, sizeof(dwBuf) ) == ERROR_SUCCESS)
			g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		RegCloseKey(hRegKey);
		return OK;
	}

	// If we reached here then the requested type was not known
	RegCloseKey(hRegKey);
	return OK;
} // RegWrite()



bool Line::RegRemoveSubkeys(HKEY hRegKey)
{
	// Removes all subkeys to the given key.  Will not touch the given key.
	CHAR Name[256];
	DWORD dwNameSize;
	FILETIME ftLastWrite;
	HKEY hSubKey;
	bool Success;

	for (;;) 
	{ // infinite loop 
		dwNameSize=255;
		if (RegEnumKeyEx(hRegKey, 0, Name, &dwNameSize, NULL, NULL, NULL, &ftLastWrite) == ERROR_NO_MORE_ITEMS)
			break;
		if ( RegOpenKeyEx(hRegKey, Name, 0, KEY_READ, &hSubKey) != ERROR_SUCCESS )
			return false;
		
		Success=RegRemoveSubkeys(hSubKey);
		RegCloseKey(hSubKey);
		if (!Success)
			return false;
		else if (RegDeleteKey(hRegKey, Name) != ERROR_SUCCESS)
			return false;
	}
	return true;
}



ResultType Line::RegDelete(char *aRegKey, char *aRegSubkey, char *aValueName)
{
	// $var = RegDelete(key[, valuename])

	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.

	HKEY	hRegKey, hMainKey;

	// Get the main key name
	if (   !(hMainKey = RegConvertMainKey(aRegKey))   )
		return OK;  // Let ErrorLevel tell the story.

	// Open the key we want
	if ( RegOpenKeyEx(hMainKey, aRegSubkey, 0, KEY_READ | KEY_WRITE, &hRegKey) != ERROR_SUCCESS )
		return OK;  // Let ErrorLevel tell the story.

	if (!aValueName || !*aValueName)
	{
		// Remove Key
		bool success = RegRemoveSubkeys(hRegKey);
		RegCloseKey(hRegKey);
		if (!success)
			return OK;  // Let ErrorLevel tell the story.
		if (RegDeleteKey(hMainKey, aRegSubkey) != ERROR_SUCCESS) 
			return OK;  // Let ErrorLevel tell the story.
	}
	else
	{
		// Remove Value
		LONG lRes = RegDeleteValue(hRegKey, aValueName);
		if (lRes != ERROR_SUCCESS)
			return OK;  // Let ErrorLevel tell the story.
		RegCloseKey(hRegKey);
	}

	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return OK;

} // RegDelete()
