/*
AutoHotkey

Copyright 2003 Chris Mallett

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
*/

#include "StdAfx.h"  // Pre-compiled headers
#ifndef _MSC_VER  // For non-MS compilers:
	#include <stdio.h>
	#include <stdlib.h>
#endif
#include <stdarg.h>
#include "util.h"


ResultType FileSetDateModified(char *aFilespec, char *aYYYYMMDD)
{
	if (!aFilespec || !*aFilespec) return FAIL;
	if (!aYYYYMMDD) aYYYYMMDD = "";  // In case caller explicitily called it with NULL.

	FILETIME ftLastWriteTime, ftLastWriteTimeUTC;
	if (*aYYYYMMDD)
	{
		// Convert the arg into the time struct as local (non-UTC) time:
		if (!YYYYMMDDToFileTime(aYYYYMMDD, &ftLastWriteTime))
			return FAIL;
		// Convert from local to UTC:
		LocalFileTimeToFileTime(&ftLastWriteTime, &ftLastWriteTimeUTC);
	}
	else
		// Put the current UTC time into the struct.
		GetSystemTimeAsFileTime(&ftLastWriteTimeUTC);

	// Open existing file.  Uses CreateFile() rather than OpenFile for an expectation
	// of greater compatibility for all files, and folder support too.
	// FILE_FLAG_NO_BUFFERING might improve performance because all we're doing is
	// changing one of the file's attributes.  FILE_FLAG_BACKUP_SEMANTICS must be
	// used, otherwise changing the time of a directory under NT and beyond will
	// not succeed.  Win95 (not sure about Win98/ME) does not support this, but it
	// should be harmless to specify it even if the OS is Win95:
	HANDLE hFile = CreateFile(aFilespec, GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE
		, (LPSECURITY_ATTRIBUTES)NULL, OPEN_EXISTING, FILE_FLAG_NO_BUFFERING | FILE_FLAG_BACKUP_SEMANTICS
		, NULL);

	if (hFile == INVALID_HANDLE_VALUE)
		return FAIL;

	BOOL result = SetFileTime(hFile, NULL, NULL, &ftLastWriteTimeUTC);
	CloseHandle(hFile);
	return result ? OK : FAIL;
}



ResultType YYYYMMDDToFileTime(char *aYYYYMMDD, FILETIME *pftDateTime)
{
	if (!aYYYYMMDD || !pftDateTime) return FAIL;
	if (!*aYYYYMMDD) return FAIL;
 
	SYSTEMTIME st;
	ZeroMemory(&st, sizeof(st));
	int nAssigned = sscanf(aYYYYMMDD, "%4d%2d%2d%2d%2d%2d", &st.wYear, &st.wMonth, &st.wDay
		, &st.wHour, &st.wMinute, &st.wSecond);

	switch (nAssigned)
	{
	case 0: return FAIL;
	case 1: st.wMonth = 1; // None of these cases does a "break", they just fall through to set default values.
	case 2: st.wDay = 1;   // Day of month.
	case 3: st.wHour = 0;  // Midnight.
	case 4: st.wMinute = 0;
	case 5: st.wSecond = 0;
	}

	if (!st.wMonth)
		st.wMonth = 1;
	if (!st.wDay)
		st.wDay = 1;
	if (!st.wHour)
	st.wMilliseconds = 0;
	if (!SystemTimeToFileTime(&st, pftDateTime)) // The wDayOfWeek member of the <st> structure is ignored.
		return FAIL;
	return OK;
}



int snprintfcat(char *aBuf, size_t aBufSize, const char *aFormat, ...)
// The caller must have ensured that aBuf contains a valid string
// (i.e. that it is null-terminated somewhere within the limits of aBufSize).
// Follows the xprintf() convention to "return the number of characters
// printed (not including the trailing '\0' used to end output to strings)
// or a negative value if an output error occurs, except for snprintf()
// and vsnprintf(), which return the number of characters that would have
// been printed if the size were unlimited (again, not including the final
// '\0')."
{
	if (!aBuf || !aFormat) return -1;  // But allow aBufSize to be zero so that the proper return value is provided.
	size_t length = strlen(aBuf);  // This could crash if caller didn't initialize it.
	__int64 space_remaining = (__int64)(aBufSize - length);
	if (space_remaining < 0) // But allow it to be zero so that the proper return value is provided.
		return -1;
	va_list ap;
	va_start(ap, aFormat);
	// Must use vsnprintf() not snprintf() because of the way va_list is handled:
	return vsnprintf(aBuf + length, (size_t)space_remaining, aFormat, ap);
}



int strlcmp(char *aBuf1, char *aBuf2, UINT aLength1, UINT aLength2)
// Similar to strncmp but considers each aBuf to be a string of length aLength
// if aLength was specified.  In other words, unlike strncmp() which would
// consider strncmp("ab", "abc", 2) to be a match, this function would consider
// them to be a mismatch.  Another way of looking at it: aBuf1 and aBuf2 will
// be directly compared to one another as though they were actually of length
// aLength1 and aLength2, respectively and then passed to strcmp() as those
// shorter strings.  This behavior is useful for cases where you don't want
// to have to bother with temporarily terminating a string so you can compare
// only a substring to something else.  The return value meaning is the
// same as strncmp().  If either aLength param is UINT_MAX (via the default
// parameters or via explicit call), it will be assumed that the entire
// length of the respective aBuf will be used.
{
	if (!aBuf1 || !aBuf2) return 0;
	if (aLength1 == UINT_MAX) aLength1 = (UINT)strlen(aBuf1);
	if (aLength2 == UINT_MAX) aLength2 = (UINT)strlen(aBuf2);
	UINT least_length = aLength1 < aLength2 ? aLength1 : aLength2;
	int diff;
	for (UINT i = 0; i < least_length; ++i)
		if (   diff = (int)((UCHAR)aBuf1[i] - (UCHAR)aBuf2[i])   ) // Use unsigned chars like strcmp().
			return diff;
	// Since the above didn't return, the strings are equal if they're the same length.
	// Otherwise, the longer one is considered greater than the shorter one since the
	// longer one's next character is by definition something non-zero.  I'm not completely
	// sure that this is the same policy followed by ANSI strcmp():
	return (int)(aLength1 - aLength2);
}	



int strlicmp(char *aBuf1, char *aBuf2, UINT aLength1, UINT aLength2)
// Case insensitive version.  See strlcmp() comments above.
{
	if (!aBuf1 || !aBuf2) return 0;
	if (aLength1 == UINT_MAX) aLength1 = (UINT)strlen(aBuf1);
	if (aLength2 == UINT_MAX) aLength2 = (UINT)strlen(aBuf2);
	UINT least_length = aLength1 < aLength2 ? aLength1 : aLength2;
	int diff;
	for (UINT i = 0; i < least_length; ++i)
		if (   diff = (int)((UCHAR)toupper(aBuf1[i]) - (UCHAR)toupper(aBuf2[i]))   )
			return diff;
	return (int)(aLength1 - aLength2);
}	



char *strrstr(char *aStr, char *aPattern, bool aCaseSensitive)
// Not an efficient implementation, but a simple one.
{
	if (!aStr || !aPattern) return NULL;
	if (!*aPattern) return aStr + strlen(aStr); // The empty string is found in every string.
	char *found, *found_prev;
	if (aCaseSensitive)
	{
		for (found_prev = found = strstr(aStr, aPattern); found; found = strstr(++found, aPattern))
			found_prev = found;
		return found_prev;
	}
	else // Not case sensitive.
	{
		for (found_prev = found = stristr(aStr, aPattern); found; found = stristr(++found, aPattern))
			found_prev = found;
		return found_prev;
	}
}



char *stristr(char *aStr, char *aPattern)
// Case insensitive version of strstr().
{
	if (!aStr || !aPattern) return NULL;
	if (!*aPattern) return aStr; // The empty string is found in every string.

	char *pptr, *sptr;
	size_t slen, plen;
	
	for (pptr  = aPattern, slen  = strlen(aStr), plen  = strlen(aPattern);
	     // while string length not shorter than aPattern length
	     slen >= plen;
	     ++aStr, --slen)
	{
		// find start of aPattern in string
		// MY: There's no need to check for end of "aStr" because slen keeps getting decremented.
		while (toupper(*aStr) != toupper(*aPattern))
		{
			++aStr;
			--slen;
			if (slen < plen) // aPattern longer than remaining part of string
				return NULL;
		}
		
		// MY: Can't start at "aStr + 1" because simple case where aPattern of length 1 would not be found.
		sptr = aStr;
		pptr = aPattern;
		
		// MY: No need to check for end of "sptr" because previous loop already verified that
		// there are enough characters remaining in "sptr" to support all those in "pptr".
		while (toupper(*sptr) == toupper(*pptr))
		{
			++sptr;
			++pptr;
			if (!*pptr) // Pattern was found.
				return aStr;
		}
	}
	return NULL;
}


// Note: This function is from someone/somewhere, I don't remember who to give credit to.
// 
// StrReplace: Replace OldStr by NewStr in string Str.
//
// Str should have enough allocated space for the replacement, no check
// is made for this. Str and OldStr/NewStr should not overlap.
// The empty string ("") is found at the beginning of every string.
//
// Returns: pointer to first location behind where NewStr was inserted
// or NULL if OldStr was not found.
// This is useful for multiple replacements, see example in main() below
// (be careful not to replace the empty string this way !)
char *StrReplace(char *Str, char *OldStr, char *NewStr, bool aCaseSensitive)
{
	if (!Str || !*Str) return NULL; // Nothing to do in this case.
	if (!OldStr || !*OldStr) return NULL;  // Don't even think about it.
	if (!NewStr) NewStr = "";  // In case it's called explicitly with a NULL.
	size_t OldLen, NewLen;
	char *p, *q;
	if (   NULL == (p = (aCaseSensitive ? strstr(Str, OldStr) : stristr(Str, OldStr)))   )
		return p;
	OldLen = strlen(OldStr);
	NewLen = strlen(NewStr);
	memmove(q = p+NewLen, p+OldLen, strlen(p+OldLen)+1);
	memcpy(p, NewStr, NewLen);
	return q;
}



char *StrReplaceAll(char *Str, char *OldStr, char *NewStr, bool aCaseSensitive)
{
	if (!Str || !*Str) return NULL; // Nothing to do in this case.
	if (!OldStr || !*OldStr) return NULL;  // Avoid replacing the empty string: probably infinite loop.
	if (!NewStr) NewStr = "";  // In case it's called explicitly with a NULL.
	char *ptr;
	//This doesn't quite work when doing a simple replacement such as ".." with ". .".
	//In the above example, "..." would be changed to ". .." rather than ". . ." as it should be.
	//Therefore, use a less efficient,but more accurate method instead.  UPDATE: But this method
	//can cause an infinite loop if the new string is a superset of the old string, so don't use
	//it after all.
	//for ( ; ptr = StrReplace(Str, OldStr, NewStr); );
	for (ptr = Str; ptr = StrReplace(ptr, OldStr, NewStr, aCaseSensitive); );
	return Str;
}


/*
int os_is_windows9x()
{
	OSVERSIONINFO osvi;
	ZeroMemory(&osvi, sizeof(OSVERSIONINFO));
	osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
	if (!GetVersionEx(&osvi))
		// Safer to assume true on failure, and this can also indicate a failure code if caller wants:
		return -1;
	return (osvi.dwPlatformId == VER_PLATFORM_WIN32_WINDOWS) ? 1 : 0;
}
*/

/*
void DisplayError (DWORD Error)
// Seems safer to have the caller send LastError to us, just in case any of the strange calls
// here ever wipe out the value of LastError before we would be able to properly use it.
{
	LPVOID lpMsgBuf;
	if (   !FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM
		, NULL, Error, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
		(LPTSTR) &lpMsgBuf,	0, NULL)   )
		return;
	// Display the string.  Uses NULL HWND for the first param, for now.
	MessageBox(NULL, (char *)lpMsgBuf, "Error", MB_OK | MB_ICONINFORMATION | MB_SYSTEMMODAL | MB_SETFOREGROUND);
	// Free the buffer.
	LocalFree(lpMsgBuf);
}
*/

/*
UINT FileTimeTenthsOfSecondUntil (FILETIME *pftStart, FILETIME *pftEnd)
// Returns the number of tenths of a second from pftStart until pftEnd.
// It will return a negative if pftStart is after pftEnd.
// One of the reasons for deci rather than milli is that an UINT (32-bit) can
// only handle about 49 days worth of milliseconds.  So if the difference is greater than
// that, it would be truncated.  I'm not sure which compilers can support long ints or
// long long ints that are greater than 32-bits in size.
{
	if (!pftStart || !pftEnd) return 0;

	// The calculation is done this way for compilers that don't support 64-bit math operations (not sure which):
	LARGE_INTEGER liStart, liEnd;
	liStart.LowPart = pftStart->dwLowDateTime;
	liStart.HighPart = pftStart->dwHighDateTime;
	liEnd.LowPart = pftEnd->dwLowDateTime;
	liEnd.HighPart = pftEnd->dwHighDateTime;
	if ((liEnd.QuadPart - liStart.QuadPart) / 1000000 > 0xFFFFFFFF)
		return 0xFFFFFFFF;
	return (UINT)((liEnd.QuadPart - liStart.QuadPart) / 1000000); // Convert from tenths-of-microsecond.
}
*/

bool DoesFilePatternExist(char *aFilePattern)
{
	// Taken from the AutoIt3 source:
	if (!aFilePattern || !*aFilePattern) return false;
	if (strchr(aFilePattern, '*') || strchr(aFilePattern, '?'))
	{
		WIN32_FIND_DATA wfd;
		HANDLE hFile = FindFirstFile(aFilePattern, &wfd);
		if (hFile == INVALID_HANDLE_VALUE)
			return false;
		FindClose(hFile);
		return true;
	}
    else
	{
		DWORD dwTemp = GetFileAttributes(aFilePattern);
		return dwTemp != 0xFFFFFFFF;
	}
}


int FileAppend(char *aFilespec, char *aLine, bool aAppendNewline)
{
	if (!aFilespec || !aLine) return FAIL;
	if (!*aFilespec) return FAIL;
	FILE *fp = fopen(aFilespec, "a");
	if (fp == NULL)
		return FAIL;
	fprintf(fp, aLine);
	if (aAppendNewline)
		putc('\n', fp);
	fclose(fp);
	return OK;
}



char *ConvertFilespecToCorrectCase(char *aFullFileSpec)
// aFullFileSpec must be a modifiable string since it will be converted to proper case.
// Returns aFullFileSpec, the contents of which have been converted to the case used by the
// file system.  Note: The trick of changing the current directory to be that of
// aFullFileSpec and then calling GetFullPathName() doesn't always work.  So perhaps the
// only easy way is to call FindFirstFile() on each directory that composes aFullFileSpec,
// which is what is done here.
{
	// Longer in case of UNCs and such, which might be longer than MAX_PATH:
	#define WORK_AREA_SIZE (MAX_PATH * 2)
	if (!aFullFileSpec || !*aFullFileSpec) return aFullFileSpec;
	size_t length = strlen(aFullFileSpec);
	if (length < 2 || length >= WORK_AREA_SIZE) return aFullFileSpec;
	// Start with something easy, the drive letter:
	if (aFullFileSpec[1] == ':')
		aFullFileSpec[0] = toupper(aFullFileSpec[0]);
	// else it might be a UNC that has no drive letter.
	char built_filespec[WORK_AREA_SIZE], *dir_start, *dir_end;
	if (dir_start = strchr(aFullFileSpec, ':'))
		// MSDN: "To examine any directory other than a root directory, use an appropriate
		// path to that directory, with no trailing backslash. For example, an argument of
		// "C:\windows" will return information about the directory "C:\windows", not about
		// any directory or file in "C:\windows". An attempt to open a search with a trailing
		// backslash will always fail."
		dir_start += 2; // Skip over the first backslash that goes with the drive letter.
	else // it's probably a UNC
	{
		if (strncmp(aFullFileSpec, "\\\\", 2))
			// It doesn't appear to be a UNC either, so not sure how to deal with it.
			return aFullFileSpec;
		// I think MS says you can't use FindFirstFile() directly on a share name, so we
		// want to omit both that and the server name from consideration (i.e. we don't attempt
		// to find their proper case).  MSDN: "Similarly, on network shares, you can use an
		// lpFileName of the form "\\server\service\*" but you cannot use an lpFileName that
		// points to the share itself, such as "\\server\service".
		dir_start = aFullFileSpec + 2;
		char *end_of_server_name = strchr(dir_start, '\\');
		if (end_of_server_name)
		{
			dir_start = end_of_server_name + 1;
			char *end_of_share_name = strchr(dir_start, '\\');
			if (end_of_share_name)
				dir_start = end_of_share_name + 1;
		}
	}
	// Init the new string (the filespec we're building), e.g. copy just the "c:\\" part.
	strlcpy(built_filespec, aFullFileSpec, dir_start - aFullFileSpec + 1);
	WIN32_FIND_DATA found_file;
	HANDLE file_search;
	for (dir_end = dir_start; dir_end = strchr(dir_end, '\\'); ++dir_end)
	{
		*dir_end = '\0';  // Temporarily terminate.
		file_search = FindFirstFile(aFullFileSpec, &found_file);
		*dir_end = '\\'; // Restore it before we do anything else.
		if (file_search == INVALID_HANDLE_VALUE)
			return aFullFileSpec;
		FindClose(file_search);
		// Append the case-corrected version of this directory name:
		snprintfcat(built_filespec, sizeof(built_filespec), "%s\\", found_file.cFileName);
	}
	// Now do the filename itself:
	if (   (file_search = FindFirstFile(aFullFileSpec, &found_file)) == INVALID_HANDLE_VALUE   )
		return aFullFileSpec;
	FindClose(file_search);
	snprintfcat(built_filespec, sizeof(built_filespec), "%s", found_file.cFileName);
	// It might be possible for the new one to be longer than the old, e.g. if some 8.3 short
	// names were converted to long names by the process.  Thus, the caller should ensure that
	// aFullFileSpec is large enough:
	strcpy(aFullFileSpec, built_filespec);
	return aFullFileSpec;
}
