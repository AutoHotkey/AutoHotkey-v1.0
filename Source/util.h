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

#ifndef util_h
#define util_h

#include "defines.h"
#include "limits.h"  // for UINT_MAX


#define IS_SPACE_OR_TAB(c) (c == ' ' || c == '\t')


inline size_t strnlen(char *aBuf, size_t aMax)
// Returns the length of aBuf or aMax, whichever is least.
// But it does so efficiently, in case aBuf is huge.
{
	if (!aMax || !aBuf || !*aBuf) return 0;
	size_t i;
	for (i = 0; aBuf[i] && i < aMax; ++i);
	return i;
}



inline char *StrChrAny(char *aStr, char *aCharList)
// Returns the position of the first char in aStr that is of any one of
// the characters listed in aCharList.  Returns NULL if not found.
{
	if (aStr == NULL || aCharList == NULL) return NULL;
	if (!*aStr || !*aCharList) return NULL;
	// Don't use strchr() because that would just find the first occurrence
	// of the first search-char, which is not necessarily the first occurrence
	// of *any* search-char:
	char *look_for_this_char;
	for (; *aStr; ++aStr) // It's safe to use the value-parameter itself.
		// If *aStr is any of the search char's, we're done:
		for (look_for_this_char = aCharList; *look_for_this_char; ++look_for_this_char)
			if (*aStr == *look_for_this_char)
				return aStr;  // Match found.
	return NULL; // No match.
}



inline char *omit_leading_whitespace(char *aBuf)
// While aBuf points to a whitespace, moves to the right and returns the first non-whitespace
// encountered.
{
	for (; IS_SPACE_OR_TAB(*aBuf); ++aBuf);
	return aBuf;
}



inline char *omit_trailing_whitespace(char *aBuf, char *aBuf_marker)
// aBuf_marker must be a position in aBuf (to the right of it).
// Starts at aBuf_marker and keeps moving to the left until a non-whitespace
// char is encountered.  Returns the position of that char.
{
	for (; aBuf_marker > aBuf && IS_SPACE_OR_TAB(*aBuf_marker); --aBuf_marker);
	return aBuf_marker;  // Can equal aBuf.
}



inline char *ltrim(char *aStr)
// NOTE: THIS VERSION trims only tabs and spaces.  It specifically avoids
// trimming newlines because some callers want to retain those.
{
	if (!aStr || !*aStr) return aStr;
	char *ptr;
	// Find the first non-whitespace char (which might be the terminator):
	for (ptr = aStr; IS_SPACE_OR_TAB(*ptr); ++ptr);
	memmove(aStr, ptr, strlen(ptr) + 1); // +1 to include the '\0'.  memmove() permits source & dest to overlap.
	return aStr;
}

inline char *rtrim(char *aStr)
// NOTE: THIS VERSION trims only tabs and spaces.  It specifically avoids
// trimming newlines because some callers want to retain those.
{
	if (!aStr || !*aStr) return aStr; // Must do this prior to below.
	// It's done this way in case aStr just happens to be address 0x00 (probably not possible
	// on Intel & Intel-clone hardware) because otherwise --cp would decrement, causing an
	// underflow since pointers are probably considered unsigned values, which would
	// probably cause an infinite loop.  Extremely unlikely, but might as well try
	// to be thorough:
	for (char *cp = aStr + strlen(aStr) - 1; ; --cp)
	{
		if (!IS_SPACE_OR_TAB(*cp))
		{
			*(cp + 1) = '\0';
			return aStr;
		}
		if (cp == aStr)
		{
			if (IS_SPACE_OR_TAB(*cp))
				*cp = '\0';
			return aStr;
		}
	}
}

inline char *trim (char *aStr)
// NOTE: THIS VERSION trims only tabs and spaces.  It specifically avoids
// trimming newlines because some callers want to retain those.
{
	return ltrim(rtrim(aStr));
}



inline bool IsPureNumeric(char *aBuf, bool aAllowNegative = false)
// String can contain whitespace.
{
	if (!aBuf || !*aBuf) return true;
	aBuf = omit_leading_whitespace(aBuf);
	if (aAllowNegative && *aBuf == '-')
		++aBuf;
	for (; *aBuf >= '0' && *aBuf <= '9'; ++aBuf);
	aBuf = omit_leading_whitespace(aBuf);
	return *aBuf == '\0';  // true if all chars in string are digits.
}



inline int PureNumberToInt(char *aBuf)
// If aBuf doesn't contain something purely numeric, zero will be returned.
// This is how AutoIt2 treats add/subtract/divide/multiply when the target
// is something non-numeric (it takes it to zero).
{
	if (!aBuf || !*aBuf) return 0;
	if (IsPureNumeric(aBuf, true))
		return atoi(aBuf);
	return 0;
}



inline size_t strlcpy (char *aDst, const char *aSrc, size_t aDstSize)
// Same as strncpy() but guarantees null-termination of aDst upon return.
// No more than aDstSize - 1 characters will be copied from aSrc into aDst
// (leaving room for the zero terminator).
// This function is defined in some Unices but is not standard.  But unlike
// other versions, this one returns the number of unused chars remaining
// in aDst after the copy.
{
	if (!aDst || !aSrc || !aDstSize) return aDstSize;  // aDstSize must not be zero due to the below method.
	strncpy(aDst, aSrc, aDstSize - 1);
	aDst[aDstSize - 1] = '\0';
	return aDstSize - strlen(aDst) - 1; // -1 because the zero terminator is defined as taking up 1 space.
}



inline char *strcatmove(char *aDst, char *aSrc)
// Same as strcat() but allows aSrc and aDst to overlap.
// Unlike strcat(), it doesn't return aDst.  Instead, it returns the position
// in aDst where aSrc was appended.
{
	if (!aDst || !aSrc || !*aSrc) return aDst;
	char *aDst_end = aDst + strlen(aDst);
	return (char *)memmove(aDst_end, aSrc, strlen(aSrc) + 1);  // Add 1 to include aSrc's terminator.
}



#define DATE_FORMAT "YYYYMMDDHHMISS"
ResultType FileSetDateModified(char *aFilespec, char *aYYYYMMDD = "");
ResultType YYYYMMDDToFileTime(char *YYYYMMDD, FILETIME *pftDateTime);

int snprintfcat(char *aBuf, size_t aBufSize, const char *aFormat, ...);
int strlcmp (char *aBuf1, char *aBuf2, UINT aLength1 = UINT_MAX, UINT aLength2 = UINT_MAX);
int strlicmp(char *aBuf1, char *aBuf2, UINT aLength1 = UINT_MAX, UINT aLength2 = UINT_MAX);
char *strrstr(char *aStr, char *aPattern, bool aCaseSensitive = true);
char *stristr(char *aStr, char *aPattern);
char *StrReplace(char *Str, char *OldStr, char *NewStr = "", bool aCaseSensitive = true);
char *StrReplaceAll(char *Str, char *OldStr, char *NewStr = "", bool aCaseSensitive = true);
//void DisplayError (DWORD Error);
// UINT FileTimeTenthsOfSecondUntil (FILETIME *pftStart, FILETIME *pftEnd);
bool DoesFilePatternExist(char *aFilePattern);
ResultType FileAppend(char *aFilespec, char *aLine, bool aAppendNewline = true);
char *ConvertFilespecToCorrectCase(char *aFullFileSpec);

#endif
