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

#include "stdafx.h" // pre-compiled headers
#include "util.h"


char *FileAttribToStr(char *aBuf, DWORD aAttr)
{
	if (!aBuf) return aBuf;
	int length = 0;
	if (aAttr & FILE_ATTRIBUTE_READONLY)
		aBuf[length++] = 'R';
	if (aAttr & FILE_ATTRIBUTE_ARCHIVE)
		aBuf[length++] = 'A';
	if (aAttr & FILE_ATTRIBUTE_SYSTEM)
		aBuf[length++] = 'S';
	if (aAttr & FILE_ATTRIBUTE_HIDDEN)
		aBuf[length++] = 'H';
	if (aAttr & FILE_ATTRIBUTE_NORMAL)
		aBuf[length++] = 'N';
	if (aAttr & FILE_ATTRIBUTE_DIRECTORY)
		aBuf[length++] = 'D';
	if (aAttr & FILE_ATTRIBUTE_OFFLINE)
		aBuf[length++] = 'O';
	if (aAttr & FILE_ATTRIBUTE_COMPRESSED)
		aBuf[length++] = 'C';
	if (aAttr & FILE_ATTRIBUTE_TEMPORARY)
		aBuf[length++] = 'T';
	aBuf[length] = '\0';  // Perform the final termination.
	return aBuf;
}



ResultType YYYYMMDDToFileTime(char *aYYYYMMDD, FILETIME *pftDateTime)
{
	if (!aYYYYMMDD || !pftDateTime) return FAIL;
	if (!*aYYYYMMDD) return FAIL;
 
	SYSTEMTIME st = {0};
	int nAssigned = sscanf(aYYYYMMDD, "%4d%2d%2d%2d%2d%2d", &st.wYear, &st.wMonth, &st.wDay
		, &st.wHour, &st.wMinute, &st.wSecond);

	switch (nAssigned) // Supply default values for those components entirely omitted.
	{
	case 0: return FAIL;
	case 1: st.wMonth = 1; // None of these cases does a "break", they just fall through to set default values.
	case 2: st.wDay = 1;   // Day of month.
	case 3: st.wHour = 0;  // Midnight.
	case 4: st.wMinute = 0;
	case 5: st.wSecond = 0;
	}

	st.wMilliseconds = 0;

	// This will return failure if aYYYYMMDD contained any invalid elements, such as an
	// explicit zero for the day of the month:
	if (!SystemTimeToFileTime(&st, pftDateTime)) // The wDayOfWeek member of the <st> structure is ignored.
		return FAIL;
	return OK;
}



char *FileTimeToYYYYMMDD(char *aBuf, FILETIME &aTime, bool aConvertToLocalTime)
{
	SYSTEMTIME st;
	if (FileTimeToSystemTime(&aTime, &st))
		return SystemTimeToYYYYMMDD(aBuf, st, aConvertToLocalTime);
	*aBuf = '\0';
	return aBuf;
}



char *SystemTimeToYYYYMMDD(char *aBuf, SYSTEMTIME &aTime, bool aConvertToLocalTime)
{
	SYSTEMTIME st;
	if (aConvertToLocalTime)
	{
		if (!SystemTimeToTzSpecificLocalTime(NULL, &aTime, &st)) // realistically: probably never fails
		{
			*aBuf = '\0';
			return aBuf;
		}
	}
	else
		CopyMemory(&st, &aTime, sizeof(st));
	sprintf(aBuf, "%04d%02d%02d" "%02d%02d%02d"
		, st.wYear, st.wMonth, st.wDay
		, st.wHour, st.wMinute, st.wSecond);
	return aBuf;
}



__int64 YYYYMMDDSecondsUntil(char *aYYYYMMDDStart, char *aYYYYMMDDEnd, bool &aFailed)
// Returns the number of seconds from aYYYYMMDDStart until aYYYYMMDDEnd.
// If aYYYYMMDDStart is blank, the current time will be used in its place.
{
	aFailed = true;  // Set default for output parameter, in case of early return.
	if (!aYYYYMMDDStart || !aYYYYMMDDEnd) return 0;

	FILETIME ftStart, ftEnd, ftNowUTC;

	if (*aYYYYMMDDStart)
	{
		if (!YYYYMMDDToFileTime(aYYYYMMDDStart, &ftStart))
			return 0;
	}
	else // Use the current time in its place.
	{
		GetSystemTimeAsFileTime(&ftNowUTC);
		FileTimeToLocalFileTime(&ftNowUTC, &ftStart);  // Convert UTC to local time.
	}
	if (*aYYYYMMDDEnd)
	{
		if (!YYYYMMDDToFileTime(aYYYYMMDDEnd, &ftEnd))
			return 0;
	}
	else // Use the current time in its place.
	{
		GetSystemTimeAsFileTime(&ftNowUTC);
		FileTimeToLocalFileTime(&ftNowUTC, &ftEnd);  // Convert UTC to local time.
	}
	aFailed = false;  // Indicate success.
	return FileTimeSecondsUntil(&ftStart, &ftEnd);
}



__int64 FileTimeSecondsUntil(FILETIME *pftStart, FILETIME *pftEnd)
// Returns the number of seconds from pftStart until pftEnd.
{
	if (!pftStart || !pftEnd) return 0;

	// The calculation is done this way for compilers that don't support 64-bit math operations (not sure which):
	// Note: This must be LARGE vs. ULARGE because we want the calculation to be signed for cases where
	// pftStart is greater than pftEnd:
	ULARGE_INTEGER uiStart, uiEnd;
	uiStart.LowPart = pftStart->dwLowDateTime;
	uiStart.HighPart = pftStart->dwHighDateTime;
	uiEnd.LowPart = pftEnd->dwLowDateTime;
	uiEnd.HighPart = pftEnd->dwHighDateTime;
	// Must do at least the inner cast to avoid losing negative results:
	return (__int64)((__int64)(uiEnd.QuadPart - uiStart.QuadPart) / 10000000); // Convert from tenths-of-microsecond.
}



//unsigned __int64 GetFileSize64(HANDLE aFileHandle)
//// Code adapted from someone's on the Usenet.
//{
//    ULARGE_INTEGER ui = {0};
//    ui.LowPart = GetFileSize(aFileHandle, &ui.HighPart);
//    if(ui.LowPart == MAXDWORD && GetLastError() != NO_ERROR)
//    {
//        // the caller should use GetLastError() to test for error
//        return ULLONG_MAX;
//    }
//    return (unsigned __int64)ui.QuadPart;
//}



int snprintf(char *aBuf, size_t aBufSize, const char *aFormat, ...)
// _snprintf() seems to copy one more character into the buffer than it should, causing overflow.
// So until that's been explained or fixed, reduce buffer size by 1 for safety.
// Follows the xprintf() convention to "return the number of characters
// printed (not including the trailing '\0' used to end output to strings)
// or a negative value if an output error occurs, except for snprintf()
// and vsnprintf(), which return the number of characters that would have
// been printed if the size were unlimited (again, not including the final
// '\0')."
{
	if (!aBuf || !aFormat) return -1;  // But allow aBufSize to be zero so that the proper return value is provided.
	if (aBufSize) // avoid underflow
		--aBufSize; // See above for explanation.
	va_list ap;
	va_start(ap, aFormat);
	// Must use _vsnprintf() not _snprintf() because of the way va_list is handled:
	return _vsnprintf(aBuf, aBufSize, aFormat, ap);
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
	__int64 space_remaining = (__int64)(aBufSize - length) - 1; // -1 for the overflow bug, see snprintf() comments.
	if (space_remaining < 0) // But allow it to be zero so that the proper return value is provided.
		return -1;
	va_list ap;
	va_start(ap, aFormat);
	// Must use vsnprintf() not snprintf() because of the way va_list is handled:
	return _vsnprintf(aBuf + length, (size_t)space_remaining, aFormat, ap);
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



char *strrstr(char *aStr, char *aPattern, bool aCaseSensitive, int aOccurrence)
// Returns NULL if not found, otherwise the address of the found string.
{
	if (!aStr || !aPattern) return NULL;
	if (aOccurrence <= 0) return NULL;
	size_t aStr_length = strlen(aStr);
	if (!*aPattern)
		// The empty string is found in every string, and since we're searching from the right, return
		// the position of the zero terminator to indicate the situation:
		return aStr + aStr_length;

	size_t aPattern_length = strlen(aPattern);
	char aPattern_last_char = aPattern[aPattern_length - 1];
	char aPattern_last_char_upper = toupper(aPattern_last_char);

	int occurrence = 0;
	char *match_starting_pos = aStr + aStr_length - 1;

	// Keep finding matches from the right until the Nth occurrence (specified by the caller) is found.
	for (;;)
	{
		if (match_starting_pos < aStr)
			return NULL;  // No further matches are possible.
		// Find (from the right) the first occurrence of aPattern's last char:
		char *last_char_match;
		for (last_char_match = match_starting_pos; last_char_match >= aStr; --last_char_match)
		{
			if (aCaseSensitive)
			{
				if (*last_char_match == aPattern_last_char)
					break;
			}
			else
			{
				if (toupper(*last_char_match) == aPattern_last_char_upper)
					break;
			}
		}

		if (last_char_match < aStr) // No further matches are possible.
			return NULL;

		// Now that aPattern's last character has been found in aStr, ensure the rest of aPattern
		// exists in aStr to the left of last_char_match:
		char *full_match, *cp;
		bool found;
		for (found = false, cp = aPattern + aPattern_length - 2, full_match = last_char_match - 1;; --cp, --full_match)
		{
			if (cp < aPattern) // The complete pattern has been found at the position in full_match + 1.
			{
				++full_match; // Adjust for the prior iteration's decrement.
				if (++occurrence == aOccurrence)
					return full_match;
				found = true;
				break;
			}
			if (full_match < aStr) // Only after checking the above is this checked.
				break;
			if (aCaseSensitive)
			{
				if (*full_match != *cp)
					break;
			}
			else
				if (toupper(*full_match) != toupper(*cp))
					break;
		} // for() innermost
		if (found) // Although the above found a match, it wasn't the right one, so resume searching.
			match_starting_pos = full_match - 1;
		else // the pattern broke odwn, so resume searching at THIS position.
			match_starting_pos = last_char_match - 1;  // Don't go back by more than 1.
	} // while() find next match
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
		// For compatibility with stricmp() and the fact that IFEQUAL, IFINSTRING and some other
		// commands all use the C libraries for insensitive searches, it seems best not to do
		// this to support accented letters and other letters whose ASCII value is above 127.
		// Some reasons for this:
		// 1) It might break existing scripts that rely on old behavior of StringReplace and Input
		//    and any other command that calls this function.
		// 2) Performance is worse, perhaps a lot due to API locale-detection overhead.
		//    (this is true at least for lstricmp() vs. stricmp(), but maybe not the below)
		//while (   (char)CharUpper((LPTSTR)(UCHAR)*sptr) == (char)CharUpper((LPTSTR)(UCHAR)*pptr)   )
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



char *StrReplaceAllSafe(char *Str, size_t Str_size, char *OldStr, char *NewStr, bool aCaseSensitive)
// Similar to above but checks to ensure that the size of the buffer isn't exceeded.
{
	if (!Str || !*Str) return NULL; // Nothing to do in this case.
	if (!OldStr || !*OldStr) return NULL;  // Avoid replacing the empty string: probably infinite loop.
	if (!NewStr) NewStr = "";  // In case it's called explicitly with a NULL.
	char *ptr;
	int length_increase = (int)(strlen(NewStr) - strlen(OldStr));  // Can be negative.
	for (ptr = Str;; )
	{
		if (length_increase > 0) // Make sure there's enough room in Str first.
			if ((int)(Str_size - strlen(Str) - 1) < length_increase)
				break;  // Not enough room to do the next replacement.
		if (   !(ptr = StrReplace(ptr, OldStr, NewStr, aCaseSensitive))   )
			break;
	}
	return Str;
}



bool DoesFilePatternExist(char *aFilePattern)
// Adapted from the AutoIt3 source:
{
	if (!aFilePattern || !*aFilePattern) return false;
	if (StrChrAny(aFilePattern, "?*"))
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


ResultType FileAppend(char *aFilespec, char *aLine, bool aAppendNewline)
{
	if (!aFilespec || !aLine) return FAIL;
	if (!*aFilespec) return FAIL;
	FILE *fp = fopen(aFilespec, "a");
	if (fp == NULL)
		return FAIL;
	fputs(aLine, fp);
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



COLORREF ColorNameToBGR(char *aColorName)
// These are the main HTML color names.  Returns CLR_DEFAULT if a matching HTML color name can't be found.
{
	if (!aColorName || !*aColorName) return CLR_DEFAULT;
	if (!stricmp(aColorName, "Black"))  return 0x000000;  // These colors are all in BGR format, not RGB.
	if (!stricmp(aColorName, "Silver")) return 0xC0C0C0;
	if (!stricmp(aColorName, "Gray"))   return 0x808080;
	if (!stricmp(aColorName, "White"))  return 0xFFFFFF;
	if (!stricmp(aColorName, "Maroon")) return 0x000080;
	if (!stricmp(aColorName, "Red"))    return 0x0000FF;
	if (!stricmp(aColorName, "Purple")) return 0x800080;
	if (!stricmp(aColorName, "Fuchsia"))return 0xFF00FF;
	if (!stricmp(aColorName, "Green"))  return 0x008000;
	if (!stricmp(aColorName, "Lime"))   return 0x00FF00;
	if (!stricmp(aColorName, "Olive"))  return 0x008080;
	if (!stricmp(aColorName, "Yellow")) return 0x00FFFF;
	if (!stricmp(aColorName, "Navy"))   return 0x800000;
	if (!stricmp(aColorName, "Blue"))   return 0xFF0000;
	if (!stricmp(aColorName, "Teal"))   return 0x808000;
	if (!stricmp(aColorName, "Aqua"))   return 0xFFFF00;
	return CLR_DEFAULT;
}



char *ConvertEscapeSequences(char *aBuf, char aEscapeChar)
{
	char *cp, *cp1;
	for (cp = aBuf; ; ++cp)  // Increment to skip over the symbol just found by the inner for().
	{
		for (; *cp && *cp != aEscapeChar; ++cp);  // Find the next escape char.
		if (!*cp) // end of string.
			break;
		cp1 = cp + 1;
		switch (*cp1)
		{
			// Only lowercase is recognized for these:
			case 'a': *cp1 = '\a'; break;  // alert (bell) character
			case 'b': *cp1 = '\b'; break;  // backspace
			case 'f': *cp1 = '\f'; break;  // formfeed
			case 'n': *cp1 = '\n'; break;  // newline
			case 'r': *cp1 = '\r'; break;  // carriage return
			case 't': *cp1 = '\t'; break;  // horizontal tab
			case 'v': *cp1 = '\v'; break;  // vertical tab
			// Otherwise, if it's not one of the above, the escape-char is considered to
			// mark the next character as literal, regardless of what it is. Examples:
			// `` -> `
			// `:: -> :: (effectively)
			// `; -> ;
			// `c -> c (i.e. unknown escape sequences resolve to the char after the `)
		}
		// Below has a final +1 to include the terminator:
		MoveMemory(cp, cp1, strlen(cp1) + 1);
	}
	return aBuf;
}



bool IsStringInList(char *aStr, char *aList, bool aCaseSensitive)
// Checks if aStr exists in aList (which is a comma-separated list).
// If aStr is blank, aList must start with a delimiting comma for there to be a match.
{
	// Must use a temp. buffer because otherwise there's no easy way to properly match upon strings
	// such as the following:
	// if var in string,,with,,literal,,commas
	char buf[LINE_SIZE];
    char *this_field = aList, *next_field, *cp;
	size_t length;

	while (*this_field)  // For each field in aList.
	{
		// To avoid the need to constantly check for buffer overflow (i.e. to keep it simple),
		// just copy up to the limit of the buffer:
		strlcpy(buf, this_field, sizeof(buf));
		// Find the end of the field inside buf.  In keeping with the tradition set by the Input command,
		// this always uses comma rather than g_delimiter.
		for (cp = buf, next_field = this_field; *cp; ++cp, ++next_field)
		{
			if (*cp == ',')
			{
				if (*(cp + 1) == ',') // Make this pair into a single literal comma.
				{
					memmove(cp, cp + 1, strlen(cp + 1) + 1);  // +1 to include the zero terminator.
					++next_field;  // An extra increment since the source string still has both commas of the pair.
				}
				else // this comma marks the end of the field.
				{
					*cp = '\0';  // Terminate the buffer to isolate just the current field.
					break;
				}
			}
		}

		if (*next_field)  // The end of the field occurred prior to the end of aList.
			++next_field; // Point it to the character after the delimiter (otherwise, leave it where it is).

		length = strlen(buf);
		if (length) // It is possible for this to be zero only for the first field.  Example: if var in ,abc
		{
			if (aCaseSensitive)
			{
				if (!strcmp(aStr, buf)) // Match found
					return true;
			}
			else // Not case sensitive
				if (!stricmp(aStr, buf)) // Match found
					return true;
		}
		else // First item in the list is the empty string, so this is a match if aStr is also blank:
			if (!*aStr)
				return true;
		this_field = next_field;
	} // while()

	return false;  // No match found.
}