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

#include "stdafx.h" // pre-compiled headers
#include "defines.h"

#define IS_SPACE_OR_TAB(c) (c == ' ' || c == '\t')


//inline int iround(double x)  // Taken from someone's "Snippets".
//{
//	return (int)floor(x + ((x >= 0) ? 0.5 : -0.5));
//}

inline char *strupr(char *str)
{
	char *temp = str;
	if (str)
		for ( ; *str; ++str)
			*str = toupper(*str);
	return temp;
}


inline char *strlwr(char *str)
{
	char *temp = str;
	if (str)
		for ( ; *str; ++str)
			*str = tolower(*str);
	return temp;
}


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


// Callers rely on PURE_NOT_NUMERIC being zero/false, so order is important:
enum pure_numeric_type {PURE_NOT_NUMERIC, PURE_INTEGER, PURE_FLOAT};
inline pure_numeric_type IsPureNumeric(char *aBuf, bool aAllowNegative = false
	, bool aAllowAllWhitespace = true, bool aAllowFloat = false, bool aAllowImpure = false)
// String can contain whitespace.
{
	aBuf = omit_leading_whitespace(aBuf); // i.e. caller doesn't have to have ltrimmed, only rtrimmed.
	if (!*aBuf) // The string is empty or consists entirely of whitespace.
		return aAllowAllWhitespace ? PURE_INTEGER : PURE_NOT_NUMERIC;
	if (*aBuf == '-')
		if (aAllowNegative)
			++aBuf;
		else
			return PURE_NOT_NUMERIC;
	if (*aBuf < '0' || *aBuf > '9') // Regardless of aAllowImpure, first char must always be non-numeric.
		return PURE_NOT_NUMERIC;
	bool is_float;
	for (is_float = false; *aBuf && !IS_SPACE_OR_TAB(*aBuf); ++aBuf)
	{
		if (*aBuf == '.')
			if (!aAllowFloat || is_float) // if aBuf contains 2 decimal points, it can't be a valid number.
				return PURE_NOT_NUMERIC;
			else
				is_float = true;
		else
			if (*aBuf < '0' || *aBuf > '9')
				if (aAllowImpure) // Since aStr starts with a number (as verified above), it is considered a number.
					return is_float ? PURE_FLOAT : PURE_INTEGER;
				else
					return PURE_NOT_NUMERIC;
	}
	if (*aBuf && *omit_leading_whitespace(aBuf))
		// The loop was broken because a space or tab was encountered, and since there's
		// something other than whitespace at the end of the number, it can't be a pure number
		// (e.g. "123 456" is not a valid number):
		return PURE_NOT_NUMERIC;
	// Since the above didn't return, it must be a float or an integer:
	return is_float ? PURE_FLOAT : PURE_INTEGER;
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


char *FileAttribToStr(char *aBuf, DWORD aAttr);

#define DATE_FORMAT "YYYYMMDDHHMISS"
ResultType YYYYMMDDToFileTime(char *YYYYMMDD, FILETIME *pftDateTime);
char *FileTimeToYYYYMMDD(char *aYYYYMMDD, FILETIME *pftDateTime, bool aConvertToLocalTime = false);
__int64 YYYYMMDDSecondsUntil(char *aYYYYMMDDStart, char *aYYYYMMDDEnd, bool &aFailed);
__int64 FileTimeSecondsUntil(FILETIME *pftStart, FILETIME *pftEnd);

unsigned __int64 GetFileSize64(HANDLE aFileHandle);
int snprintfcat(char *aBuf, size_t aBufSize, const char *aFormat, ...);
int strlcmp (char *aBuf1, char *aBuf2, UINT aLength1 = UINT_MAX, UINT aLength2 = UINT_MAX);
int strlicmp(char *aBuf1, char *aBuf2, UINT aLength1 = UINT_MAX, UINT aLength2 = UINT_MAX);
char *strrstr(char *aStr, char *aPattern, bool aCaseSensitive = true);
char *stristr(char *aStr, char *aPattern);
char *StrReplace(char *Str, char *OldStr, char *NewStr = "", bool aCaseSensitive = true);
char *StrReplaceAll(char *Str, char *OldStr, char *NewStr = "", bool aCaseSensitive = true);
bool DoesFilePatternExist(char *aFilePattern);
ResultType FileAppend(char *aFilespec, char *aLine, bool aAppendNewline = true);
char *ConvertFilespecToCorrectCase(char *aFullFileSpec);

#endif
