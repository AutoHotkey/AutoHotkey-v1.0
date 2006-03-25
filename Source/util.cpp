/*
AutoHotkey

Copyright 2003-2006 Chris Mallett (support@autohotkey.com)

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
#include <olectl.h> // for OleLoadPicture()
#include <Gdiplus.h> // Used by LoadPicture().
#include "util.h"
#include "globaldata.h"


int GetYDay(int aMon, int aDay, bool aIsLeapYear)
// Returns a number between 1 and 366.
// Caller must verify that aMon is a number between 1 and 12, and aDay is a number between 1 and 31.
{
	--aMon;  // Convert to zero-based.
	if (aIsLeapYear)
	{
		int leap_offset[12] = {0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335};
		return leap_offset[aMon] + aDay;
	}
	else
	{
		int normal_offset[12] = {0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334};
		return normal_offset[aMon] + aDay;
	}
}



int GetISOWeekNumber(char *aBuf, int aYear, int aYDay, int aWDay)
// Caller must ensure that aBuf is of size 7 or greater, that aYear is a valid year (e.g. 2005),
// that aYDay is between 1 and 366, and that aWDay is between 0 and 6 (day of the week).
// Produces the week number in YYYYNN format, e.g. 200501.
// Note that year is also returned because it isn't necessarily the same as aTime's calendar year.
// Based on Linux glibc source code (GPL).
{
	--aYDay;  // Convert to zero based.
	#define ISO_WEEK_START_WDAY 1 // Monday
	#define ISO_WEEK1_WDAY 4      // Thursday
	#define ISO_WEEK_DAYS(yday, wday) (yday - (yday - wday + ISO_WEEK1_WDAY + ((366 / 7 + 2) * 7)) % 7 \
		+ ISO_WEEK1_WDAY - ISO_WEEK_START_WDAY);

	int year = aYear;
	int days = ISO_WEEK_DAYS(aYDay, aWDay);

	if (days < 0) // This ISO week belongs to the previous year.
	{
		--year;
		days = ISO_WEEK_DAYS(aYDay + (365 + IS_LEAP_YEAR(year)), aWDay);
	}
	else
	{
		int d = ISO_WEEK_DAYS(aYDay - (365 + IS_LEAP_YEAR(year)), aWDay);
		if (0 <= d) // This ISO week belongs to the next year.
		{
			++year;
			days = d;
		}
	}

	// Use snprintf() for safety; that is, in case year contains a value longer than 4 digits.
	// This also adds the leading zeros in front of year and week number, if needed.
	snprintf(aBuf, 7, "%04d%02d", year, (days / 7) + 1);
	return 6; // The length of the string produced.
}



ResultType YYYYMMDDToFileTime(char *aYYYYMMDD, FILETIME &aFileTime)
{
	SYSTEMTIME st;
	YYYYMMDDToSystemTime(aYYYYMMDD, st, false);  // "false" because it's validated below.
	// This will return failure if aYYYYMMDD contained any invalid elements, such as an
	// explicit zero for the day of the month.  It also reports failure if st.wYear is
	// less than 1601, which for simplicity is enforced globally throughout the program
	// since none of the Windows API calls seem to support earlier years.
	return SystemTimeToFileTime(&st, &aFileTime) ? OK : FAIL; // The st.wDayOfWeek member is ignored.
}



DWORD YYYYMMDDToSystemTime2(char *aYYYYMMDD, SYSTEMTIME *aSystemTime)
// Calls YYYYMMDDToSystemTime() to fill up to two elements of the aSystemTime array.
// Returns a GDTR bitwise combination to indicate which of the two elements, or both, are valid.
// Caller must ensure that aYYYYMMDD is a modifiable string since it's temporarily altered and restored here.
{
	DWORD gdtr = 0;
	if (!*aYYYYMMDD)
		return gdtr;
	if (*aYYYYMMDD != '-') // Since first char isn't a dash, there is a minimum present.
	{
		char *cp;
		if (cp = strchr(aYYYYMMDD + 1, '-'))
			*cp = '\0'; // Temporarily terminate in case only the leading part of the YYYYMMDD format is present.  Otherwise, the dash and other chars would be considered invalid fields.
		if (YYYYMMDDToSystemTime(aYYYYMMDD, aSystemTime[0], true)) // Date string is valid.
			gdtr |= GDTR_MIN; // Indicate that minimum is present.
		if (cp)
		{
			*cp = '-'; // Undo the temp. termination.
			aYYYYMMDD = cp + 1; // Set it to the maximum's position for use below.
		}
		else // No dash, so there is no maximum.  Indicate this by making aYYYYMMDD empty.
			aYYYYMMDD = "";
	}
	else // *aYYYYMMDD=='-', so only the maximum is present; thus there will be no minimum.
		++aYYYYMMDD; // Skip over the dash to set it to the maximum's position.
	if (*aYYYYMMDD) // There is a maximum.
	{
		if (YYYYMMDDToSystemTime(aYYYYMMDD, aSystemTime[1], true)) // Date string is valid.
			gdtr |= GDTR_MAX; // Indicate that maximum is present.
	}
	return gdtr;
}



ResultType YYYYMMDDToSystemTime(char *aYYYYMMDD, SYSTEMTIME &aSystemTime, bool aDoValidate)
// Although aYYYYMMDD need not be terminated at the end of the YYYYMMDDHH24MISS string (as long as
// the string's capacity is at least 14), it should be terminated if only the leading part
// of the YYYYMMDDHH24MISS format is present.
// Caller must ensure that aYYYYMMDD is non-NULL.  If aDoValidate is false, OK is always
// returned and aSystemTime might contain invalid elements.  Otherwise, FAIL will be returned
// if the date and time contains any invalid elements, or if the year is less than 1601
// (Windows generally does not support earlier years).
{
	// sscanf() is avoided because it adds 2 KB to the compressed EXE size.
	char temp[16];
	size_t length = strlen(aYYYYMMDD); // Use this rather than incrementing the pointer in case there are ever partial fields such as 20051 vs. 200501.

	strlcpy(temp, aYYYYMMDD, 5);
	aSystemTime.wYear = atoi(temp);

	if (length > 4) // It has a month component.
	{
		strlcpy(temp, aYYYYMMDD + 4, 3);
		aSystemTime.wMonth = atoi(temp);  // Unlike "struct tm", SYSTEMTIME uses 1 for January, not 0.
	}
	else
		aSystemTime.wMonth = 1;

	if (length > 6) // It has a day-of-month component.
	{
		strlcpy(temp, aYYYYMMDD + 6, 3);
		aSystemTime.wDay = atoi(temp);
	}
	else
		aSystemTime.wDay = 1;

	if (length > 8) // It has an hour component.
	{
		strlcpy(temp, aYYYYMMDD + 8, 3);
		aSystemTime.wHour = atoi(temp);
	}
	else
		aSystemTime.wHour = 0;   // Midnight.

	if (length > 10) // It has a minutes component.
	{
		strlcpy(temp, aYYYYMMDD + 10, 3);
		aSystemTime.wMinute = atoi(temp);
	}
	else
		aSystemTime.wMinute = 0;

	if (length > 12) // It has a seconds component.
	{
		strlcpy(temp, aYYYYMMDD + 12, 3);
		aSystemTime.wSecond = atoi(temp);
	}
	else
		aSystemTime.wSecond = 0;

	aSystemTime.wMilliseconds = 0;  // Always set to zero in this case.

	// Day-of-week code by Tomohiko Sakamoto:
	static int t[] = {0, 3, 2, 5, 0, 3, 5, 1, 4, 6, 2, 4};
	int y = aSystemTime.wYear;
	y -= aSystemTime.wMonth < 3;
	aSystemTime.wDayOfWeek = (y + y/4 - y/100 + y/400 + t[aSystemTime.wMonth-1] + aSystemTime.wDay) % 7;

	if (aDoValidate)
	{
		FILETIME ft;
		// This will return failure if aYYYYMMDD contained any invalid elements, such as an
		// explicit zero for the day of the month.  It also reports failure if st.wYear is
		// less than 1601, which for simplicity is enforced globally throughout the program
		// since none of the Windows API calls seem to support earlier years.
		return SystemTimeToFileTime(&aSystemTime, &ft) ? OK : FAIL;
		// Above: The st.wDayOfWeek member is ignored, but that's okay since only the YYYYMMDDHH24MISS part
		// needs validation.
	}
	return OK;
}



char *FileTimeToYYYYMMDD(char *aBuf, FILETIME &aTime, bool aConvertToLocalTime)
{
	FILETIME ft;
	if (aConvertToLocalTime)
		FileTimeToLocalFileTime(&aTime, &ft); // MSDN says that target cannot be the same var as source.
	else
		memcpy(&ft, &aTime, sizeof(FILETIME));  // memcpy() might be less code size that a struct assignment, ft = aTime.
	SYSTEMTIME st;
	if (FileTimeToSystemTime(&ft, &st))
		return SystemTimeToYYYYMMDD(aBuf, st);
	*aBuf = '\0';
	return aBuf;
}



char *SystemTimeToYYYYMMDD(char *aBuf, SYSTEMTIME &aTime)
// Remember not to offer a "aConvertToLocalTime" option, because calling SystemTimeToTzSpecificLocalTime()
// on Win9x apparently results in an invalid time because the function is implemented only as a stub on
// those OSes.
{
	sprintf(aBuf, "%04d%02d%02d" "%02d%02d%02d"
		, aTime.wYear, aTime.wMonth, aTime.wDay
		, aTime.wHour, aTime.wMinute, aTime.wSecond);
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
		if (!YYYYMMDDToFileTime(aYYYYMMDDStart, ftStart))
			return 0;
	}
	else // Use the current time in its place.
	{
		GetSystemTimeAsFileTime(&ftNowUTC);
		FileTimeToLocalFileTime(&ftNowUTC, &ftStart);  // Convert UTC to local time.
	}
	if (*aYYYYMMDDEnd)
	{
		if (!YYYYMMDDToFileTime(aYYYYMMDDEnd, ftEnd))
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



SymbolType IsPureNumeric(char *aBuf, bool aAllowNegative, bool aAllowAllWhitespace
	, bool aAllowFloat, bool aAllowImpure)
// String can contain whitespace.
// If aBuf doesn't contain something purely numeric, PURE_NOT_NUMERIC is returned.  The same happens if
// aBuf contains a float but aAllowFloat is false.  Otherwise, PURE_INTEGER or PURE_FLOAT is returned.
// If aAllowAllWhitespace==true and the string is blank or all whitespace, PURE_INTEGER is returned.
// Obsolete comment: Making this non-inline reduces the size of the compressed EXE by only 2K.  Since this
// function is called so often, it seems preferable to keep it inline for performance.
{
	aBuf = omit_leading_whitespace(aBuf); // i.e. caller doesn't have to have ltrimmed, only rtrimmed.
	if (!*aBuf) // The string is empty or consists entirely of whitespace.
		return aAllowAllWhitespace ? PURE_INTEGER : PURE_NOT_NUMERIC;

	if (*aBuf == '-')
	{
		if (aAllowNegative)
			++aBuf;
		else
			return PURE_NOT_NUMERIC;
	}
	else if (*aBuf == '+')
		++aBuf;

	// Relies on short circuit boolean order to prevent reading beyond the end of the string:
	bool is_hex = IS_HEX(aBuf);
	if (is_hex)
		aBuf += 2;  // Skip over the 0x prefix.

	// Set defaults:
	bool has_decimal_point = false;
	bool has_at_least_one_digit = false; // i.e. a string consisting of only "+", "-" or "." is not considered numeric.

	for (; *aBuf && !IS_SPACE_OR_TAB(*aBuf); ++aBuf)
	{
		if (*aBuf == '.')
		{
			if (!aAllowFloat || has_decimal_point || is_hex)
				// i.e. if aBuf contains 2 decimal points, it can't be a valid number.
				// Note that decimal points are allowed in hexadecimal strings, e.g. 0xFF.EE.
				// But since that format doesn't seem to be supported by VC++'s atof() and probably
				// related functions, and since it's extremely rare, it seems best not to support it.
				return PURE_NOT_NUMERIC;
			else
				has_decimal_point = true;
		}
		else
		{
			if (is_hex ? !isxdigit(*aBuf) : (*aBuf < '0' || *aBuf > '9')) // And since we're here, it's not '.' either.
				if (aAllowImpure) // Since aStr starts with a number (as verified above), it is considered a number.
				{
					if (has_at_least_one_digit)
						return has_decimal_point ? PURE_FLOAT : PURE_INTEGER;
					else // i.e. the strings "." and "-" are not considered to be numeric by themselves.
						return PURE_NOT_NUMERIC;
				}
				else
					return PURE_NOT_NUMERIC;
			else
				has_at_least_one_digit = true;
		}
	}
	if (*aBuf) // The loop was broken because a space or tab was encountered.
		if (*omit_leading_whitespace(aBuf)) // But that space or tab is followed by something other than whitespace.
			if (!aAllowImpure) // e.g. "123 456" is not a valid pure number.
				return PURE_NOT_NUMERIC;
			// else fall through to the bottom logic.
		// else since just whitespace at the end, the number qualifies as pure, so fall through.
		// (it would already have returned in the loop if it was impure)
	// else since end of string was encountered, the number qualifies as pure, so fall through.
	// (it would already have returned in the loop if it was impure).
	if (has_at_least_one_digit)
		return has_decimal_point ? PURE_FLOAT : PURE_INTEGER;
	else
		return PURE_NOT_NUMERIC; // i.e. the strings "+" "-" and "." are not numeric by themselves.
}



void strlcpy(char *aDst, const char *aSrc, size_t aDstSize) // Non-inline because it benches slightly faster that way.
// Caller must ensure that aDstSize is greater than 0.
// Caller must ensure that the entire capacity of aDst is writable, EVEN WHEN it knows that aSrc is much shorter
// than the aDstSize.  This is because the call to strncpy (which is used for its superior performance) zero-fills
// any unused portion of aDst.
// Description:
// Same as strncpy() but guarantees null-termination of aDst upon return.
// No more than aDstSize - 1 characters will be copied from aSrc into aDst
// (leaving room for the zero terminator, which is always inserted).
// This function is defined in some Unices but is not standard.  But unlike
// other versions, this one uses void for return value for reduced code size
// (since it's called in so many places).
{
	// Disabled for performance and reduced code size:
	//if (!aDst || !aSrc || !aDstSize) return aDstSize;  // aDstSize must not be zero due to the below method.
	// It might be worthwhile to have a custom char-copying-loop here someday so that number of characters
	// actually copied (not including the zero terminator) can be returned to callers who want it.
	--aDstSize; // Convert from size to length (caller has ensured that aDstSize > 0).
	strncpy(aDst, aSrc, aDstSize); // NOTE: In spite of its zero-filling, strncpy() benchmarks considerably faster than a custom loop, probably because it uses 32-bit memory operations vs. 8-bit.
	aDst[aDstSize] = '\0';
}



int snprintf(char *aBuf, int aBufSize, const char *aFormat, ...)
// aBufSize is an int so that any negative values passed in from caller are not lost.
// aBuf will always be terminated here except when aBufSize is <= zero (in which case the caller should
// already have terminated it).  If aBufSize is greater than zero but not large enough to hold the
// entire result, as much of the result as possible is copied and the return value is aBufSize - 1.
// Returns the exact number of characters written, not including the zero terminator.  A negative
// number is never returned, even if aBufSize is <= zero (which means there isn't even enough space left
// to write a zero terminator), under the assumption that the caller has already terminated the string
// and thus prefers to have 0 rather than -1 returned in such cases.
// MSDN says (about _snprintf(), and testing shows that it applies to _vsnprintf() too): "This function
// does not guarantee NULL termination, so ensure it is followed by sz[size - 1] = 0".
{
	// The following should probably never be changed without a full suite of tests to ensure the
	// change doesn't cause the finicky _vsnprintf() to break something.
	if (aBufSize < 1 || !aBuf || !aFormat) return 0; // It's called from so many places that the extra checks seem warranted.
	va_list ap;
	va_start(ap, aFormat);
	// Must use _vsnprintf() not _snprintf() because of the way va_list is handled:
	int result = _vsnprintf(aBuf, aBufSize, aFormat, ap); // "returns the number of characters written, not including the terminating null character, or a negative value if an output error occurs"
	aBuf[aBufSize - 1] = '\0'; // Confirmed through testing: Must terminate at this exact spot because _vsnprintf() doesn't always do it.
	// Fix for v1.0.34: If result==aBufSize, must reduce result by 1 to return an accurate result to the
	// caller.  In other words, if the line above turned the last character into a terminator, one less character
	// is now present in aBuf.
	if (result == aBufSize)
		--result;
	return result > -1 ? result : aBufSize - 1; // Never return a negative value.  See comment under function definition, above.
}



int snprintfcat(char *aBuf, int aBufSize, const char *aFormat, ...)
// aBufSize is an int so that any negative values passed in from caller are not lost.
// aBuf will always be terminated here except when the amount of space left in the buffer is zero or less.
// (in which case the caller should already have terminated it).  If aBufSize is greater than zero but not
// large enough to hold the entire result, as much of the result as possible is copied and the return value
// is space_remaining - 1.
// The caller must have ensured that aBuf and aFormat are non-NULL and that aBuf contains a valid string
// (i.e. that it is null-terminated somewhere within the limits of aBufSize).
// Returns the exact number of characters written, not including the zero terminator.  A negative
// number is never returned, even if aBufSize is <= zero (which means there isn't even enough space left
// to write a zero terminator), under the assumption that the caller has already terminated the string
// and thus prefers to have 0 rather than -1 returned in such cases.
{
	// The following should probably never be changed without a full suite of tests to ensure the
	// change doesn't cause the finicky _vsnprintf() to break something.
	size_t length = strlen(aBuf);
	int space_remaining = (int)(aBufSize - length); // Must cast to int to avoid loss of negative values.
	if (space_remaining < 1) // Can't even terminate it (no room) so just indicate that no characters were copied.
		return 0;
	aBuf += length;  // aBuf is now the spot where the new text will be written.
	va_list ap;
	va_start(ap, aFormat);
	// Must use vsnprintf() not snprintf() because of the way va_list is handled:
	int result = _vsnprintf(aBuf, (size_t)space_remaining, aFormat, ap); // "returns the number of characters written, not including the terminating null character, or a negative value if an output error occurs"
	aBuf[space_remaining - 1] = '\0'; // Confirmed through testing: Must terminate at this exact spot because _vsnprintf() doesn't always do it.
	return result > -1 ? result : space_remaining - 1; // Never return a negative value.  See comment under function definition, above.
}



// Not currently used by anything, so commented out to possibly reduce code size:
//int strlcmp(char *aBuf1, char *aBuf2, UINT aLength1, UINT aLength2)
//// Case sensitive version.  See strlicmp() comments below.
//{
//	if (!aBuf1 || !aBuf2) return 0;
//	if (aLength1 == UINT_MAX) aLength1 = (UINT)strlen(aBuf1);
//	if (aLength2 == UINT_MAX) aLength2 = (UINT)strlen(aBuf2);
//	UINT least_length = aLength1 < aLength2 ? aLength1 : aLength2;
//	int diff;
//	for (UINT i = 0; i < least_length; ++i)
//		if (   diff = (int)((UCHAR)aBuf1[i] - (UCHAR)aBuf2[i])   ) // Use unsigned chars like strcmp().
//			return diff;
//	return (int)(aLength1 - aLength2);
//}	



int strlicmp(char *aBuf1, char *aBuf2, UINT aLength1, UINT aLength2)
// Similar to strnicmp but considers each aBuf to be a string of length aLength if aLength was
// specified.  In other words, unlike strnicmp() which would consider strnicmp("ab", "abc", 2)
// [example verified correct] to be a match, this function would consider them to be
// a mismatch.  Another way of looking at it: aBuf1 and aBuf2 will be directly
// compared to one another as though they were actually of length aLength1 and
// aLength2, respectively and then passed to stricmp() (not strnicmp) as those
// shorter strings.  This behavior is useful for cases where you don't want
// to have to bother with temporarily terminating a string so you can compare
// only a substring to something else.  The return value meaning is the
// same as strnicmp().  If either aLength param is UINT_MAX (via the default
// parameters or via explicit call), it will be assumed that the entire
// length of the respective aBuf will be used.
{
	if (!aBuf1 || !aBuf2) return 0;
	if (aLength1 == UINT_MAX) aLength1 = (UINT)strlen(aBuf1);
	if (aLength2 == UINT_MAX) aLength2 = (UINT)strlen(aBuf2);
	UINT least_length = aLength1 < aLength2 ? aLength1 : aLength2;
	int diff;
	for (UINT i = 0; i < least_length; ++i)
		if (   diff = (int)((UCHAR)toupper(aBuf1[i]) - (UCHAR)toupper(aBuf2[i]))   )
			return diff;
	// Since the above didn't return, the strings are equal if they're the same length.
	// Otherwise, the longer one is considered greater than the shorter one since the
	// longer one's next character is by definition something non-zero.  I'm not completely
	// sure that this is the same policy followed by ANSI strcmp():
	return (int)(aLength1 - aLength2);
}	



char *strrstr(char *aStr, char *aPattern, bool aCaseSensitive, int aOccurrence)
// Returns NULL if not found, otherwise the address of the found string.
{
	if (aOccurrence < 1) return NULL;
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
		else // the pattern broke down, so resume searching at THIS position.
			match_starting_pos = last_char_match - 1;  // Don't go back by more than 1.
	} // while() find next match
}



char *strcasestr(const char *phaystack, const char *pneedle)
	// To make this work with MS Visual C++, this version uses tolower/toupper() in place of
	// _tolower/_toupper(), since apparently in GNU C, the underscore macros are identical
	// to the non-underscore versions; but in MS the underscore ones do an unconditional
	// conversion (mangling non-alphabetic characters such as the zero terminator).  MSDN:
	// tolower: Converts c to lowercase if appropriate
	// _tolower: Converts c to lowercase

	// Return the offset of one string within another.
	// Copyright (C) 1994,1996,1997,1998,1999,2000 Free Software Foundation, Inc.
	// This file is part of the GNU C Library.

	// The GNU C Library is free software; you can redistribute it and/or
	// modify it under the terms of the GNU Lesser General Public
	// License as published by the Free Software Foundation; either
	// version 2.1 of the License, or (at your option) any later version.

	// The GNU C Library is distributed in the hope that it will be useful,
	// but WITHOUT ANY WARRANTY; without even the implied warranty of
	// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
	// Lesser General Public License for more details.

	// You should have received a copy of the GNU Lesser General Public
	// License along with the GNU C Library; if not, write to the Free
	// Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA
	// 02111-1307 USA.

	// My personal strstr() implementation that beats most other algorithms.
	// Until someone tells me otherwise, I assume that this is the
	// fastest implementation of strstr() in C.
	// I deliberately chose not to comment it.  You should have at least
	// as much fun trying to understand it, as I had to write it :-).
	// Stephen R. van den Berg, berg@pool.informatik.rwth-aachen.de

	// Faster looping by precalculating bl, bu, cl, cu before looping.
	// 2004 Apr 08	Jose Da Silva, digital@joescat@com
{
	register const unsigned char *haystack, *needle;
	register unsigned bl, bu, cl, cu;
	
	haystack = (const unsigned char *) phaystack;
	needle = (const unsigned char *) pneedle;

	bl = tolower (*needle);
	if (bl != '\0')
	{
		// Scan haystack until the first character of needle is found:
		bu = toupper (bl);
		haystack--;				/* possible ANSI violation */
		do
		{
			cl = *++haystack;
			if (cl == '\0')
				goto ret0;
		}
		while ((cl != bl) && (cl != bu));

		// See if the rest of needle is a one-for-one match with this part of haystack:
		cl = tolower (*++needle);
		if (cl == '\0')  // Since needle consists of only one character, it is already a match as found above.
			goto foundneedle;
		cu = toupper (cl);
		++needle;
		goto jin;
		
		for (;;)
		{
			register unsigned a;
			register const unsigned char *rhaystack, *rneedle;
			do
			{
				a = *++haystack;
				if (a == '\0')
					goto ret0;
				if ((a == bl) || (a == bu))
					break;
				a = *++haystack;
				if (a == '\0')
					goto ret0;
shloop:
				;
			}
			while ((a != bl) && (a != bu));

jin:
			a = *++haystack;
			if (a == '\0')  // Remaining part of haystack is shorter than needle.  No match.
				goto ret0;

			if ((a != cl) && (a != cu)) // This promising candidate is not a complete match.
				goto shloop;            // Start looking for another match on the first char of needle.
			
			rhaystack = haystack-- + 1;
			rneedle = needle;
			a = tolower (*rneedle);
			
			if (tolower (*rhaystack) == (int) a)
			do
			{
				if (a == '\0')
					goto foundneedle;
				++rhaystack;
				a = tolower (*++needle);
				if (tolower (*rhaystack) != (int) a)
					break;
				if (a == '\0')
					goto foundneedle;
				++rhaystack;
				a = tolower (*++needle);
			}
			while (tolower (*rhaystack) == (int) a);
			
			needle = rneedle;		/* took the register-poor approach */
			
			if (a == '\0')
				break;
		} // for(;;)
	} // if (bl != '\0')
foundneedle:
	return (char*) haystack;
ret0:
	return 0;
}



char *StrReplace(char *aBuf, char *aOld, char *aNew, bool aCaseSensitive)
// Replaces first occurrence of aOld with aNew in string aBuf.  Caller must ensure that
// all parameters are non-NULL (though they can be the empty string).  It must also ensure that
// aBuf has enough allocated space for the replacement since no check is made for this.
// Finally, aOld/aNew should not be inside the same memory area as aBuf.
// The empty string ("") is found at the beginning of every string.
// Returns NULL if aOld was not found.  Otherwise, returns address of first location
// behind where aNew was inserted.  This is useful for multiple replacements (though performance
// would be bad since length of old/new and Buf has to be recalculated with strlen() for each
// call, rather than just once for all calls).
{
	// Nothing to do if aBuf is blank.  If aOld is blank, that is not supported because it
	// would be an infinite loop.
	if (!*aBuf || !*aOld)
		return NULL;
	char *found = (aCaseSensitive ? strstr(aBuf, aOld) : strcasestr(aBuf, aOld));
	if (!found)
		return NULL;
	size_t aOld_length = strlen(aOld);
	size_t aNew_length = strlen(aNew);
	char *the_part_of_aBuf_to_remain_unaltered = found + aNew_length;
	// The check below can greatly improve performance if old and new strings happen to be same length
	// (especially if this function is called thousands of times in a loop to replace multiple strings):
	if (aOld_length != aNew_length)
		// Since new string can't fit exactly in place of old string, adjust the target area to
		// accept exactly the right length so that the rest of the string stays unaltered:
		memmove(the_part_of_aBuf_to_remain_unaltered, found + aOld_length, strlen(found + aOld_length) + 1); // +1 to include the terminator.
	memcpy(found, aNew, aNew_length); // Perform the replacement.
	return the_part_of_aBuf_to_remain_unaltered;
}



char *StrReplaceAll(char *aBuf, char *aOld, char *aNew, bool aAlwaysUseSlow, bool aCaseSensitive, DWORD aReplacementsNeeded)
// Replaces all occurrences of aOld with aNew inside aBuf, and returns aBuf.
{
	// Nothing to do if aBuf is blank.  If aOld is blank, that is not supported because it
	// would be an infinite loop.
	if (!*aBuf || !*aOld)
		return aBuf;
	// When the replacement string is equal in length to the search string, StrReplace()
	// already contains an enhancement to avoid the memmove() (though some implementations of
	// memmove() might already check if source and dest are the same and if so, return
	// immediately, in which case the extra check wouldn't be needed).
	// When replacement string is longer than search string, I don't know if there is
	// a simple algorithm to avoid a memmove() for each replacement.   One obvious way would
	// be to allocate a temporary memory area and build an entirely new string in place of
	// the old one.  Obviously that would only be better if the system had the memory to spare
	// and didn't have to swap to allocate it.  So such an algorithm should probably be offered
	// only as an option (in the future), especially since it wouldn't help performance much
	// unless many replacements are needed within the target string.
	size_t aOld_length = strlen(aOld);
	size_t aNew_length = strlen(aNew);
	size_t aBuf_length = strlen(aBuf);
	if (aOld_length == aNew_length)
		aAlwaysUseSlow = true; // Use slow mode because it's just as fast in this case, while avoiding the extra memory allocation.

	if (!aAlwaysUseSlow) // Auto-detect whether to use fast mode.
	{
		// The fast method is much faster when many replacements are needed because it avoids
		// a memmove() to shift the remainder of the buffer up against the area of memory that
		// will be replaced (which for the slow method is done for every replacement).  The savings
		// can be enormous if aBuf is very large, assuming the system has the memory to spare.
		char *cp, *dp;
		UINT replacements_needed;
		if (aReplacementsNeeded < UINT_MAX) // Caller provided the count to avoid us having to calculate it again.
			replacements_needed = aReplacementsNeeded;
		else
			for (replacements_needed = 0, cp = aBuf; cp = (aCaseSensitive ? strstr(cp, aOld) : strcasestr(cp, aOld)); cp += aOld_length)
				++replacements_needed;
		if (!replacements_needed) // aBuf already contains the correct contents.
			return aBuf;
		// Fall back to the slow method unless the fast method's performance gain is worth the risk of
		// stressing the system (in case it's low on memory).
		// Testing shows that the below cutoffs seem approximately correct based on the fact that the
		// slow method's execution time is proportional to both aBuf_length and replacements_needed.
		// By contrast, the fast method takes about the same amount of time for a given aBuf_length
		// regardless of how many replacements are needed.  Here are the test results (not rigorously
		// conducted -- they were done only to get a idea of the shape of the curve):
		//1415 KB with 5 evenly spaced replacements: slow and fast are about the same speed.
		//2800 KB with 10 replacements: 110ms vs. 90ms
		//5600 KB with 10 replacements: 220ms vs. 175ms
		//Same but with 20 replacements: 335 vs. 195 (at this point, the slow method is taking more than 50% longer)
		//Same but with 40 replacements: 525 vs. 185 (fast method is more than twice as fast)
		//475KB with 6200 replacements: 3055 vs. 25  (the difference becomes dramatic for thousands of replacements).
		//150 KB with 2000 replacements: 91 vs. 0ms
		//293 KB with 100,000 replacements: 14000 vs. 30
		//146 KB with 50,000 replacements: 2954 vs 20
		//30KB with 10000 replacements: 60 vs. 0
		if (replacements_needed > 20 && aBuf_length > 5000) // Also fall back to slow method if buffer is of trivial size, since there isn't much difference in that case.
		{
			// Allocate the memory:
			size_t buf_size = aBuf_length + replacements_needed*(aNew_length - aOld_length) + 1; // +1 for zero terminator.
			char *buf = (char *)malloc(buf_size);
			if (buf)
			{
				// Perform the replacement:
				cp = aBuf; // Source.
				dp = buf;  // Destination.
				size_t chars_to_copy;
				for (char *found = aBuf; found = (aCaseSensitive ? strstr(found, aOld) : strcasestr(found, aOld));)
				{
					// memcpy() might contain optimizations that make it faster than a char-moving loop such as memmove().
					// This is because memmove() or a simple *dp++ = *cp++ loop of our own allows the source and dest
					// to overlap.  Since they don't in this case, a CPU instruction might be used to copy a block of
					// memory much more quickly:
					chars_to_copy = found - cp;
					if (chars_to_copy) // Copy the part of the source string up to the position of this match.
					{
						memcpy(dp, cp, chars_to_copy);
						dp += chars_to_copy;
					}
					if (aNew_length) // Insert the replacement string in place of the old string.
					{
						memcpy(dp, aNew, aNew_length);
						dp += aNew_length;
					}
					//else omit it altogether; that is, replace every aOld with the empty string.

					// Set up "found" for the next search to be done by the for-loop.  For consistency, like
					// the "slow" method, overlapping matches are not detected.  For example, the replacement
					// of all occurrences of ".." with ". ." would transform an aBuf of "..." into ". ..",
					// not ". . .":
					found += aOld_length;
					cp = found; // Since "found" is about to be altered by strstr, cp serves as a placeholder for use by the next iteration.
				}

				// Fixed for v1.0.25.11: Copy the remaining characters after the last replacement (if any).
				// Due to the check above, it is known there was more than zero replacements.
				memcpy(dp, cp, aBuf_length - (cp - aBuf) + 1); // Should perform better than strcpy(). +1 to include the terminator.

				// Copy the result back into the caller's original buf (with added complexity by the caller, this step
				// could be avoided):
				memcpy(aBuf, buf, buf_size); // Include the zero terminator. Caller has already ensured that aBuf is large enough.
				free(buf);
				return aBuf;
			}
			//else not enough memory, so fall through to the slow method below.
		}
		// else only 1 replacement needed, so the slow mode should be about the same speed while requiring less memory.
	}

	// Since the above didn't return, either the slow method was originally in effect or it's being
	// used because the fast method could not allocate enough memory or isn't worth the cost.
	// The below doesn't quite work when doing a simple replacement such as ".." with ". .".
	// In the above example, "..." would be changed to ". .." rather than ". . ." as it should be.
	// Therefore, use a less efficient, but more accurate method instead.  UPDATE: But this method
	// can cause an infinite loop if the new string is a superset of the old string, so don't use
	// it after all.
	//for ( ; ptr = StrReplace(aBuf, aOld, aNew, aCaseSensitive); ); // Note that this very different from the below.

	// Don't call StrReplace() because its call of strlen() within the call to memmove() signficantly
	// reduces performance (a typical replacement of \r\n with \n in an aBuf of size 2 MB is close to
	// twice as fast due to the avoidance of a call to strlen() inside the loop).  Reproducing the code
	// here has the additional advantage that the lengths of aOld and aNew need be calculated only once
	// at the beginning rather than having StrReplace() do it every time, which can be a very large
	// savings if aOld and/or aNew happen to be very long strings (unusual).
	int length_to_add = (int)(aNew_length - aOld_length);  // Can be negative.
	char *found, *search_area;
	for (search_area = aBuf; found = aCaseSensitive ? strstr(search_area, aOld) : strcasestr(search_area, aOld);)
	{
		search_area = found + aNew_length;  // The next search should start at this position when all is adjusted below.
		// The check below can greatly improve performance if old and new strings happen to be same length:
		if (aOld_length != aNew_length)
		{
			// Since new string can't fit exactly in place of old string, adjust the target area to
			// accept exactly the right length so that the rest of the string stays unaltered:
			memmove(search_area, found + aOld_length
				, aBuf_length - (found - aBuf) - aOld_length + 1); // +1 to include zero terminator.
			// Above: Calculating length vs. using strlen() makes overall speed of the operation about
			// twice as fast for some typical test cases in a 2 MB buffer such as replacing \r\n with \n.
		}
		memcpy(found, aNew, aNew_length); // Perform the replacement.
		// Must keep aBuf_length updated as we go, for use with memmove() above:
		aBuf_length += length_to_add; // Note that length_to_add will be negative if aNew is shorter than aOld.
	}

	return aBuf;
}



int StrReplaceAllSafe(char *aBuf, size_t aBufSize, char *aOld, char *aNew, bool aCaseSensitive)
// Similar to above but checks to ensure that the size of the buffer isn't exceeded.
// Returns how many replacements were done.
{
	// Nothing to do if aBuf is blank.  If aOld is blank, that is not supported because it
	// would be an infinite loop.
	if (!*aBuf || !*aOld)
		return 0;
	char *ptr;
	int replacment_count;
	int length_increase = (int)(strlen(aNew) - strlen(aOld));  // Can be negative.
	for (replacment_count = 0, ptr = aBuf;; )
	{
		if (length_increase > 0) // Make sure there's enough room in aBuf first.
			if ((int)(aBufSize - strlen(aBuf) - 1) < length_increase)
				break;  // Not enough room to do the next replacement.
		if (   !(ptr = StrReplace(ptr, aOld, aNew, aCaseSensitive))   )
			break;
		// Otherwise, it did actually replace one item, so increment:
		++replacment_count;
	}
	return replacment_count;
}



char *TranslateLFtoCRLF(char *aString)
// Translates any naked LFs in aString to CRLF.  If there are non, the original string is returned.
// Otherwise, the translated versionis copied into a malloc'd buffer, which the caller must free
// when it's done with it).  Any CRLFs originally present in aString are not changed (i.e. they
// don't become CRCRLF).
{
	UINT naked_LF_count = 0;
	size_t length = 0;
	char *cp;

	for (cp = aString; *cp; ++cp)
	{
		++length;
		if (*cp == '\n' && (cp == aString || *(cp - 1) != '\r')) // Relies on short-circuit boolean order.
			++naked_LF_count;
	}

	if (!naked_LF_count)
		return aString;  // The original string is returned, which the caller must check for (vs. new string).

	// Allocate the new memory that will become the caller's responsibility:
	char *buf = (char *)malloc(length + naked_LF_count + 1);  // +1 for zero terminator.
	if (!buf)
		return NULL;

	// Now perform the translation.
	char *dp = buf; // Destination.
	for (cp = aString; *cp; ++cp)
	{
		if (*cp == '\n' && (cp == aString || *(cp - 1) != '\r')) // Relies on short-circuit boolean order.
			*dp++ = '\r';  // Insert an extra CR here, then insert the '\n' normally below.
		*dp++ = *cp;
	}
	*dp = '\0';  // Final terminator.

	return buf;  // Caller must free it when it's done with it.
}



bool DoesFilePatternExist(char *aFilePattern, DWORD *aFileAttr)
// Returns true if the file/folder exists or false otherwise.
// If non-NULL, aFileAttr's DWORD is set to the attributes of the file/folder if a match is found.
// If there is no match, its contents are undefined.
{
	if (!aFilePattern || !*aFilePattern) return false;
	// Fix for v1.0.35.12: Don't consider the question mark in "\\?\Volume{GUID}\" to be a wildcard.
	// Such volume names are obtained from GetVolumeNameForVolumeMountPoint() and perhaps other functions.
	// However, testing shows that wildcards beyond that first one should be seen as real wildcards
	// because the wildcard match-method below does work for such volume names.
	char *cp = strncmp(aFilePattern, "\\\\?\\", 4) ? aFilePattern : aFilePattern + 4;
	if (StrChrAny(cp, "?*"))
	{
		WIN32_FIND_DATA wfd;
		HANDLE hFile = FindFirstFile(aFilePattern, &wfd);
		if (hFile == INVALID_HANDLE_VALUE)
			return false;
		FindClose(hFile);
		if (aFileAttr)
			*aFileAttr = wfd.dwFileAttributes;
		return true;
	}
    else
	{
		DWORD attr = GetFileAttributes(aFilePattern);
		if (aFileAttr)
			*aFileAttr = attr;
		return attr != 0xFFFFFFFF;
	}
}



#ifdef _DEBUG
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
#endif



char *ConvertFilespecToCorrectCase(char *aFullFileSpec)
// aFullFileSpec must be a modifiable string since it will be converted to proper case.
// Also, it should be at least MAX_PATH is size because if it contains any short (8.3)
// components, they will be converted into their long names.
// Returns aFullFileSpec, the contents of which have been converted to the case used by the
// file system.  Note: The trick of changing the current directory to be that of
// aFullFileSpec and then calling GetFullPathName() doesn't always work.  So perhaps the
// only easy way is to call FindFirstFile() on each directory that composes aFullFileSpec,
// which is what is done here.  I think there's another way involving some PIDL Explorer
// function, but that might not support UNCs correctly.
{
	if (!aFullFileSpec || !*aFullFileSpec) return aFullFileSpec;
	size_t length = strlen(aFullFileSpec);
	if (length < 2 || length >= MAX_PATH) return aFullFileSpec;
	// Start with something easy, the drive letter:
	if (aFullFileSpec[1] == ':')
		aFullFileSpec[0] = toupper(aFullFileSpec[0]);
	// else it might be a UNC that has no drive letter.
	char built_filespec[MAX_PATH], *dir_start, *dir_end;
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



char *FileAttribToStr(char *aBuf, DWORD aAttr)
// Caller must ensure that aAttr is valid (i.e. that it's not 0xFFFFFFFF).
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



unsigned __int64 GetFileSize64(HANDLE aFileHandle)
// Returns ULLONG_MAX on failure.  Otherwise, it returns the actual file size.
{
    ULARGE_INTEGER ui = {0};
    ui.LowPart = GetFileSize(aFileHandle, &ui.HighPart);
    if (ui.LowPart == MAXDWORD && GetLastError() != NO_ERROR)
        return ULLONG_MAX;
    return (unsigned __int64)ui.QuadPart;
}



char *GetLastErrorText(char *aBuf, int aBufSize, bool aUpdateLastError)
// aBufSize is an int to preserve any negative values the caller might pass in.
{
	if (aBufSize < 1)
		return aBuf;
	if (aBufSize == 1)
	{
		*aBuf = '\0';
		return aBuf;
	}
	DWORD last_error = GetLastError();
	if (aUpdateLastError)
		g.LastError = last_error;
	FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, NULL, last_error, 0, aBuf, (DWORD)aBufSize - 1, NULL);
	return aBuf;
}



void AssignColor(char *aColorName, COLORREF &aColor, HBRUSH &aBrush)
// Assign the color indicated in aColorName (either a name or a hex RGB value) to both
// aColor and aBrush, deleting any prior handle in aBrush first.  If the color cannot
// be determined, it will always be set to CLR_DEFAULT (and aBrush set to NULL to match).
// It will never be set to CLR_NONE.
{
	COLORREF color;
	if (!*aColorName)
		color = CLR_DEFAULT;
	else
	{
		color = ColorNameToBGR(aColorName);
		if (color == CLR_NONE) // A matching color name was not found, so assume it's a hex color value.
			// It seems strtol() automatically handles the optional leading "0x" if present:
			color = rgb_to_bgr(strtol(aColorName, NULL, 16));
			// if aColorName does not contain something hex-numeric, black (0x00) will be assumed,
			// which seems okay given how rare such a problem would be.
	}
	if (color != aColor) // It's not already the right color.
	{
		if (aBrush) // Free the resources of the old brush.
			DeleteObject(aBrush);
		if (color == CLR_DEFAULT) // Caller doesn't need brush for CLR_DEFAULT, assuming that's even possible.
			aBrush = NULL;
		else
		{
			if (aBrush = CreateSolidBrush(color)) // Assign.  Failure should be very rare.
				aColor = color;
			else
				aColor = CLR_DEFAULT; // A NULL HBRUSH should always corresponds to CLR_DEFAULT.
		}
	}
}



COLORREF ColorNameToBGR(char *aColorName)
// These are the main HTML color names.  Returns CLR_NONE if a matching HTML color name can't be found.
// Returns CLR_DEFAULT only if aColorName is the word Default.
{
	if (!aColorName || !*aColorName) return CLR_NONE;
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
	if (!stricmp(aColorName, "Default"))return CLR_DEFAULT;
	return CLR_NONE;
}



POINT CenterWindow(int aWidth, int aHeight)
// Given a the window's width and height, calculates where to position its upper-left corner
// so that it is centered EVEN IF the task bar is on the left side or top side of the window.
// This does not currently handle multi-monitor systems explicitly, since those calculations
// require API functions that don't exist in Win95/NT (and thus would have to be loaded
// dynamically to allow the program to launch).  Therefore, windows will likely wind up
// being centered across the total dimensions of all monitors, which usually results in
// half being on one monitor and half in the other.  This doesn't seem too terrible and
// might even be what the user wants in some cases (i.e. for really big windows).
{
	RECT rect;
	SystemParametersInfo(SPI_GETWORKAREA, 0, &rect, 0);  // Get desktop rect excluding task bar.
	// Note that rect.left will NOT be zero if the taskbar is on docked on the left.
	// Similarly, rect.top will NOT be zero if the taskbar is on docked at the top of the screen.
	POINT pt;
	pt.x = rect.left + (((rect.right - rect.left) - aWidth) / 2);
	pt.y = rect.top + (((rect.bottom - rect.top) - aHeight) / 2);
	return pt;
}



bool FontExist(HDC aHdc, char *aTypeface)
{
	LOGFONT lf;
	lf.lfCharSet = DEFAULT_CHARSET;  // Enumerate all char sets.
	lf.lfPitchAndFamily = 0;  // Must be zero.
	strlcpy(lf.lfFaceName, aTypeface, LF_FACESIZE);
	bool font_exists = false;
	EnumFontFamiliesEx(aHdc, &lf, (FONTENUMPROC)FontEnumProc, (LPARAM)&font_exists, 0);
	return font_exists;
}



int CALLBACK FontEnumProc(ENUMLOGFONTEX *lpelfe, NEWTEXTMETRICEX *lpntme, DWORD FontType, LPARAM lParam)
{
	*((bool *)lParam) = true; // Indicate to the caller that the font exists.
	return 0;  // Stop the enumeration after the first, since even one match means the font exists.
}



void WindowToScreen(int &aX, int &aY)
// aX and aY are assumed to be relative to the currently active window.  Here they are converted to
// screen coordinates based on the position of the active window upper-left corner (not its client area).
{
	RECT rect;
	HWND fore_hwnd = GetForegroundWindow();
	if (fore_hwnd && !IsIconic(fore_hwnd) && GetWindowRect(fore_hwnd, &rect))
	{
		aX += rect.left;
		aY += rect.top;
	}
	//else no active window per se, so don't convert the coordinates.  Leave them as-is as desired by the
	// caller.  More details:
	// Revert to screen coordinates if the foreground window is minimized.  Although it might be
	// impossible for a visible window to be both foreground and minmized, it seems that hidden
	// windows -- such as the script's own main window when activated for the purpose of showing
	// a popup menu -- can be foreground while simultaneously being minimized.  This fixes an
	// issue where the mouse will move to the upper-left corner of the screen rather than the
	// intended coordinates (v1.0.17):
}



void GetVirtualDesktopRect(RECT &aRect)
{
	aRect.right = GetSystemMetrics(SM_CXVIRTUALSCREEN);
	if (aRect.right) // A non-zero value indicates the OS supports multiple monitors or at least SM_CXVIRTUALSCREEN.
	{
		aRect.left = GetSystemMetrics(SM_XVIRTUALSCREEN);  // Might be negative or greater than zero.
		aRect.right += aRect.left;
		aRect.top = GetSystemMetrics(SM_YVIRTUALSCREEN);   // Might be negative or greater than zero.
		aRect.bottom = aRect.top + GetSystemMetrics(SM_CYVIRTUALSCREEN);
	}
	else // Win95/NT do not support SM_CXVIRTUALSCREEN and such, so zero was returned.
		GetWindowRect(GetDesktopWindow(), &aRect);
}



ResultType RegReadString(HKEY aRootKey, char *aSubkey, char *aValueName, char *aBuf, size_t aBufSize)
{
	*aBuf = '\0'; // Set default output parameter.  Some callers rely on this being set even if failure occurs.
	HKEY hkey;
	if (RegOpenKeyEx(aRootKey, aSubkey, 0, KEY_QUERY_VALUE, &hkey) != ERROR_SUCCESS)
		return FAIL;
	DWORD buf_size = (DWORD)aBufSize; // Caller's value might be a constant memory area, so need a modifiable copy.
	LONG result = RegQueryValueEx(hkey, aValueName, NULL, NULL, (LPBYTE)aBuf, &buf_size);
	RegCloseKey(hkey);
	return (result == ERROR_SUCCESS) ? OK : FAIL;
}	



LPVOID AllocInterProcMem(HANDLE &aHandle, DWORD aSize, HWND aHwnd)
// aHandle is an output parameter that holds the mapping for Win9x and the process handle for NT.
// Returns NULL on failure (in which case caller should ignore the value of aHandle).
{
	// ALLOCATE APPROPRIATE TYPE OF MEMORY (depending on OS type)
	LPVOID mem;
	if (g_os.IsWin9x()) // Use file-mapping method.
	{
		if (   !(aHandle = CreateFileMapping(INVALID_HANDLE_VALUE, NULL, PAGE_READWRITE, 0, aSize, NULL))   )
			return NULL;
		mem = MapViewOfFile(aHandle, FILE_MAP_ALL_ACCESS, 0, 0, 0);
	}
	else // NT/2k/XP/2003 or later.  Use the VirtualAllocEx() so that caller can use Read/WriteProcessMemory().
	{
		DWORD pid;
		GetWindowThreadProcessId(aHwnd, &pid);
		// Even if the PID is our own, open the process anyway to simplify the code. After all, it would be
		// pretty silly for a script to access its own ListViews via this method.
		if (   !(aHandle = OpenProcess(PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_VM_WRITE, FALSE, pid))   )
			return NULL; // Let ErrorLevel tell the story.
		// Load function dynamically to allow program to launch on win9x:
		typedef LPVOID (WINAPI *MyVirtualAllocExType)(HANDLE, LPVOID, SIZE_T, DWORD, DWORD);
		static MyVirtualAllocExType MyVirtualAllocEx = (MyVirtualAllocExType)GetProcAddress(GetModuleHandle("kernel32")
			, "VirtualAllocEx");
		// Reason for using VirtualAllocEx(): When sending LVITEM structures to a control in a remote process, the
		// structure and its pszText buffer must both be memory inside the remote process rather than in our own.
		// Allocate both the LVITEM struct and its internal string buffer in one go because MyVirtualAllocEx() is
		// probably a high overhead call:
		mem = MyVirtualAllocEx(aHandle, NULL, aSize, MEM_RESERVE | MEM_COMMIT, PAGE_READWRITE);
	}
	if (!mem)
		CloseHandle(aHandle); // Closes the mapping for Win9x and the process handle for other OSes. Caller should ignore the value of aHandle when return value is NULL.
	//else leave the handle open (required for both methods).  It's the caller's responsibility to close it.
	return mem;
}



void FreeInterProcMem(HANDLE aHandle, LPVOID aMem)
// Caller has ensured that aMem is a file-mapping for Win9x and a VirtualAllocEx block for NT/2k/XP+.
// Similarly, it has ensured that aHandle is a file-mapping handle for Win9x and a process handle for NT/2k/XP+.
{
	if (g_os.IsWin9x())
		UnmapViewOfFile(aMem);
	else
	{
		// Load function dynamically to allow program to launch on win9x:
		typedef BOOL (WINAPI *MyVirtualFreeExType)(HANDLE, LPVOID, SIZE_T, DWORD);
		static MyVirtualFreeExType MyVirtualFreeEx = (MyVirtualFreeExType)GetProcAddress(GetModuleHandle("kernel32")
			, "VirtualFreeEx");
		MyVirtualFreeEx(aHandle, aMem, 0, MEM_RELEASE); // Size 0 is used with MEM_RELEASE.
	}
	// The following closes either the mapping or the process handle, depending on OS type.
	// But close it only after the above is done using it.
	CloseHandle(aHandle);
}



HBITMAP LoadPicture(char *aFilespec, int aWidth, int aHeight, int &aImageType, int aIconIndex, bool aUseGDIPlusIfAvailable)
// Loads a JPG/GIF/BMP/ICO/etc. and returns an HBITMAP or HICON to the caller (which it may call
// DeleteObject()/DestroyIcon() upon, though upon program termination all such handles are freed
// automatically).  The image is scaled to the specified width and height.  If zero is specified
// for either, the image's actual size will be used for that dimension.  If -1 is specified for one,
// that dimension will be kept proportional to the other dimension's size so that the original aspect
// ratio is retained.  Caller should specify -1 for aIconIndex unless an icon other than 1 is desired.
// .ico/.cur/.ani files are normally loaded as HICON (unless aUseGDIPlusIfAvailable is true of something
// else unusual happened such as file contents not matching file's extension).  This is done to preserve
// any properties that HICONs have but HBITMAPs lack, namely the ability to be animated and perhaps other things.
// Returns NULL on failure.
{
	HBITMAP hbitmap = NULL;
	aImageType = -1; // The type of image currently inside hbitmap.  Set default value for output parameter as "unknown".

	if (!*aFilespec) // Allow blank filename to yield NULL bitmap (and currently, some callers do call it this way).
		return NULL;

	char *file_ext = strrchr(aFilespec, '.');
	if (file_ext)
		++file_ext;

	// Must use ExtractIcon() if either of the following is true:
	// 1) Caller gave an explicit icon index, i.e. it wants us to use ExtractIcon() even for the first icon.
	// 2) The target file is an EXE or DLL (LoadImage() is documented not to work on those file types).
	bool ExtractIcon_was_used = aIconIndex > -1 || (file_ext && (!stricmp(file_ext, "exe") || !stricmp(file_ext, "dll")));
	if (ExtractIcon_was_used)
	{
		aImageType = IMAGE_ICON;
		if (aIconIndex < 0)
			aIconIndex = 0;  // Use the default, which is the first icon.
		hbitmap = (HBITMAP)ExtractIcon(g_hInstance, aFilespec, aIconIndex); // Return value of 1 means "incorrect file type".
		if (!hbitmap || hbitmap == (HBITMAP)1 || (!aWidth && !aHeight)) // Couldn't load icon, or could but no resizing is needed.
			return hbitmap;
		//else continue on below so that the icon can be resized to the caller's specified dimensions.
	}

	// Make an initial guess of the type of image if the above didn't already determine the type:
	if (aImageType < 0)
	{
		if (file_ext) // Assume generic file-loading method if there's no file extension.
		{
			if (!stricmp(file_ext, "ico"))
				aImageType = IMAGE_ICON;
			else if (!stricmp(file_ext, "cur") || !stricmp(file_ext, "ani"))
				aImageType = IMAGE_CURSOR;
			else if (!stricmp(file_ext, "bmp"))
				aImageType = IMAGE_BITMAP;
			//else for other extensions, leave set to "unknown" so that the below knows to use IPic or GDI+ to load it.
		}
		//else same comment as above.
	}

	if ((aWidth == -1 || aHeight == -1) && (!aWidth || !aHeight))
		aWidth = aHeight = 0; // i.e. One dimension is zero and the other is -1, which resolves to the same as "keep original size".
	bool keep_aspect_ratio = (aWidth == -1 || aHeight == -1);

	HINSTANCE hinstGDI = NULL;
	if (aUseGDIPlusIfAvailable && !(hinstGDI = LoadLibrary("gdiplus")))
		aUseGDIPlusIfAvailable = false; // Override any original "true" value as a signal for the section below.

	if (!hbitmap && aImageType > -1 && !aUseGDIPlusIfAvailable)
	{
		// Since image hasn't yet be loaded and since the file type appears to be one supported by
		// LoadImage() [icon/cursor/bitmap], attempt that first.  If it fails, fall back to the other
		// methods below in case the file's internal contents differ from what the file extension indicates.
		int desired_width, desired_height;
		if (keep_aspect_ratio)
			desired_width = desired_height = 0; // Load image at its actual size.  It will be rescaled to retain aspect ratio later below.
		else
		{
			desired_width = aWidth;
			desired_height = aHeight;
		}
		// LR_CREATEDIBSECTION applies only when aImageType == IMAGE_BITMAP, but seems appropriate in that case:
		if (hbitmap = (HBITMAP)LoadImage(NULL, aFilespec, aImageType, desired_width, desired_height
			, LR_LOADFROMFILE | LR_CREATEDIBSECTION))
		{
			// The above might have loaded an HICON vs. an HBITMAP (it has been confirmed that LoadImage()
			// will return an HICON vs. HBITMAP is aImageType is IMAGE_ICON/CURSOR).  Note that HICON and
			// HCURSOR are identical for most/all Windows API uses.  Also note that LoadImage() will load
			// an icon as a bitmap if the file contains an icon but IMAGE_BITMAP was passed in (at least
			// on Windows XP).
			if (!keep_aspect_ratio) // No further resizing is needed.
				return hbitmap;
			// Otherwise, continue on so that the image can be resized via a second call to LoadImage().
		}
		// v1.0.40.10: Abort if file doesn't exist so that GDIPlus isn't even attempted. This is done because
		// loading GDIPlus apparently disrupts the color palette of certain games, at least old ones that use
		// DirectDraw in 256-color depth.
		else if (GetFileAttributes(aFilespec) == 0xFFFFFFFF) // For simplicity, we don't check if it's a directory vs. file, since that should be too rare.
			return NULL;
		//else file exists, so continue on so that the other methods are attempted in case file's contents
		// differ from what the file extension indicates, or in case the other methods can be successful
		// even when the above failed.
	}

	IPicture *pic = NULL; // Also used to detect whether IPic method was used to load the image.

	if (!hbitmap) // Above hasn't loaded the image yet, so use the fall-back methods.
	{
		// At this point, regardless of the image type being loaded (even an icon), it will
		// definitely be converted to a Bitmap below.  So set the type:
		aImageType = IMAGE_BITMAP;
		// Find out if this file type is supported by the non-GDI+ method.  This check is not foolproof
		// since all it does is look at the file's extension, not its contents.  However, it doesn't
		// need to be 100% accurate because its only purpose is to detect whether the higher-overhead
		// calls to GdiPlus can be avoided.
		if (aUseGDIPlusIfAvailable || !file_ext || (stricmp(file_ext, "jpg")
			&& stricmp(file_ext, "jpeg") && stricmp(file_ext, "gif"))) // Non-standard file type (BMP is already handled above).
			if (!hinstGDI) // We don't yet have a handle from an earlier call to LoadLibary().
				hinstGDI = LoadLibrary("gdiplus");
		// If it is suspected that the file type isn't supported, try to use GdiPlus if available.
		// If it's not available, fall back to the old method in case the filename doesn't properly
		// reflect its true contents (i.e. in case it really is a JPG/GIF/BMP internally).
		// If the below LoadLibrary() succeeds, either the OS is XP+ or the GdiPlus extensions have been
		// installed on an older OS.
		if (hinstGDI)
		{
			// LPVOID and "int" are used to avoid compiler errors caused by... namespace issues?
			typedef int (WINAPI *GdiplusStartupType)(ULONG_PTR*, LPVOID, LPVOID);
			typedef VOID (WINAPI *GdiplusShutdownType)(ULONG_PTR);
			typedef int (WINGDIPAPI *GdipCreateBitmapFromFileType)(LPVOID, LPVOID);
			typedef int (WINGDIPAPI *GdipCreateHBITMAPFromBitmapType)(LPVOID, LPVOID, DWORD);
			typedef int (WINGDIPAPI *GdipDisposeImageType)(LPVOID);
			GdiplusStartupType DynGdiplusStartup = (GdiplusStartupType)GetProcAddress(hinstGDI, "GdiplusStartup");
  			GdiplusShutdownType DynGdiplusShutdown = (GdiplusShutdownType)GetProcAddress(hinstGDI, "GdiplusShutdown");
  			GdipCreateBitmapFromFileType DynGdipCreateBitmapFromFile = (GdipCreateBitmapFromFileType)GetProcAddress(hinstGDI, "GdipCreateBitmapFromFile");
  			GdipCreateHBITMAPFromBitmapType DynGdipCreateHBITMAPFromBitmap = (GdipCreateHBITMAPFromBitmapType)GetProcAddress(hinstGDI, "GdipCreateHBITMAPFromBitmap");
  			GdipDisposeImageType DynGdipDisposeImage = (GdipDisposeImageType)GetProcAddress(hinstGDI, "GdipDisposeImage");

			ULONG_PTR token;
			Gdiplus::GdiplusStartupInput gdi_input;
			Gdiplus::GpBitmap *pgdi_bitmap;
			if (DynGdiplusStartup && DynGdiplusStartup(&token, &gdi_input, NULL) == Gdiplus::Ok)
			{
				wchar_t filespec_wide[MAX_PATH];
				mbstowcs(filespec_wide, aFilespec, sizeof(filespec_wide));
				if (DynGdipCreateBitmapFromFile(filespec_wide, &pgdi_bitmap) == Gdiplus::Ok)
				{
					if (DynGdipCreateHBITMAPFromBitmap(pgdi_bitmap, &hbitmap, CLR_DEFAULT) != Gdiplus::Ok)
						hbitmap = NULL; // Set to NULL to be sure.
					DynGdipDisposeImage(pgdi_bitmap); // This was tested once to make sure it really returns Gdiplus::Ok.
				}
				// The current thought is that shutting it down every time conserves resources.  If so, it
				// seems justified since it is probably called infrequently by most scripts:
				DynGdiplusShutdown(token);
			}
			FreeLibrary(hinstGDI);
		}
		else // Using old picture loading method.
		{
			// Based on code sample at http://www.codeguru.com/Cpp/G-M/bitmap/article.php/c4935/
			HANDLE hfile = CreateFile(aFilespec, GENERIC_READ, 0, NULL, OPEN_EXISTING, 0, NULL);
			if (hfile == INVALID_HANDLE_VALUE)
				return NULL;
			DWORD size = GetFileSize(hfile, NULL);
			HGLOBAL hglobal = GlobalAlloc(GMEM_MOVEABLE, size);
			if (!hglobal)
			{
				CloseHandle(hfile);
				return NULL;
			}
			LPVOID hlocked = GlobalLock(hglobal);
			if (!hlocked)
			{
				CloseHandle(hfile);
				GlobalFree(hglobal);
				return NULL;
			}
			// Read the file into memory:
			ReadFile(hfile, hlocked, size, &size, NULL);
			GlobalUnlock(hglobal);
			CloseHandle(hfile);
			LPSTREAM stream;
			if (FAILED(CreateStreamOnHGlobal(hglobal, FALSE, &stream)) || !stream)  // Relies on short-circuit boolean order.
			{
				GlobalFree(hglobal);
				return NULL;
			}
			// Specify TRUE to have it do the GlobalFree() for us.  But since the call might fail, it seems best
			// to free the mem ourselves to avoid uncertainy over what it does on failure:
			if (FAILED(OleLoadPicture(stream, 0, FALSE, IID_IPicture, (void **)&pic)))
				pic = NULL;
			stream->Release();
			GlobalFree(hglobal);
			if (!pic)
				return NULL;
			pic->get_Handle((OLE_HANDLE *)&hbitmap);
			// Above: MSDN: "The caller is responsible for this handle upon successful return. The variable is set
			// to NULL on failure."
			if (!hbitmap)
			{
				pic->Release();
				return NULL;
			}
			// Don't pic->Release() yet because that will also destroy/invalidate hbitmap handle.
		} // IPicture method was used.
	} // IPicture or GDIPlus was used to load the image, not a simple LoadImage() or ExtractIcon().

	// Above has ensured that hbitmap is now not NULL.
	// Adjust things if "keep aspect ratio" is in effect:
	if (keep_aspect_ratio)
	{
		HBITMAP hbitmap_to_analyze;
		ICONINFO ii; // Must be declared at this scope level.
		if (aImageType == IMAGE_BITMAP)
			hbitmap_to_analyze = hbitmap;
		else // icon or cursor
		{
			if (GetIconInfo((HICON)hbitmap, &ii)) // Works on cursors too.
				hbitmap_to_analyze = ii.hbmMask; // Use Mask because MSDN implies hbmColor can be NULL for monochrome cursors and such.
			else
			{
				DestroyIcon((HICON)hbitmap);
				return NULL; // No need to call pic->Release() because since it's an icon, we know IPicture wasn't used (it only loads bitmaps).
			}
		}
		// Above has ensured that hbitmap_to_analyze is now not NULL.  Find bitmap's dimensions.
		BITMAP bitmap;
		GetObject(hbitmap_to_analyze, sizeof(BITMAP), &bitmap); // Realistically shouldn't fail at this stage.
		if (aHeight == -1)
		{
			// Caller wants aHeight calculated based on the specified aWidth (keep aspect ratio).
			if (bitmap.bmWidth) // Avoid any chance of divide-by-zero.
				aHeight = (int)(((double)bitmap.bmHeight / bitmap.bmWidth) * aWidth + .5); // Round.
		}
		else
		{
			// Caller wants aWidth calculated based on the specified aHeight (keep aspect ratio).
			if (bitmap.bmHeight) // Avoid any chance of divide-by-zero.
				aWidth = (int)(((double)bitmap.bmWidth / bitmap.bmHeight) * aHeight + .5); // Round.
		}
		if (aImageType != IMAGE_BITMAP)
		{
			// It's our reponsibility to delete these two when they're no longer needed:
			DeleteObject(ii.hbmColor);
			DeleteObject(ii.hbmMask);
			// If LoadImage() vs. ExtractIcon() was used originally, call LoadImage() again because
			// I haven't found any other way to retain an animated cursor's animation (and perhaps
			// other icon/cursor attributes) when resizing the icon/cursor (CopyImage() doesn't
			// retain animation):
			if (!ExtractIcon_was_used)
			{
				DestroyIcon((HICON)hbitmap); // Destroy the original HICON.
				// Load a new one, but at the size newly calculated above.
				// Due to an apparent bug in Windows 9x (at least Win98se), the below call will probably
				// crash the program with a "divide error" if the specified aWidth and/or aHeight are
				// greater than 90.  Since I don't know whether this affects all versions of Windows 9x, and
				// all animated cursors, it seems best just to document it here and in the help file rather
				// than limiting the dimensions of .ani (and maybe .cur) files for certain operating systems.
				return (HBITMAP)LoadImage(NULL, aFilespec, aImageType, aWidth, aHeight, LR_LOADFROMFILE);
			}
		}
	}

	HBITMAP hbitmap_new; // To hold the scaled image (if scaling is needed).
	if (pic) // IPicture method was used.
	{
		// The below statement is confirmed by having tested that DeleteObject(hbitmap) fails
		// if called after pic->Release():
		// "Copy the image. Necessary, because upon pic's release the handle is destroyed."
		// MSDN: CopyImage(): "[If either width or height] is zero, then the returned image will have the
		// same width/height as the original."
		// Note also that CopyImage() seems to provide better scaling quality than using MoveWindow()
		// (followed by redrawing the parent window) on the static control that contains it:
		hbitmap_new = (HBITMAP)CopyImage(hbitmap, IMAGE_BITMAP, aWidth, aHeight // We know it's IMAGE_BITMAP in this case.
			, (aWidth || aHeight) ? 0 : LR_COPYRETURNORG); // Produce original size if no scaling is needed.
		pic->Release();
		// No need to call DeleteObject(hbitmap), see above.
	}
	else // GDIPlus or a simple method such as LoadImage or ExtractIcon was used.
	{
		if (!aWidth && !aHeight) // No resizing needed.
			return hbitmap;
		// The following will also handle HICON/HCURSOR correctly if aImageType == IMAGE_ICON/CURSOR.
		// Also, LR_COPYRETURNORG|LR_COPYDELETEORG is used because it might allow the animation of
		// a cursor to be retained if the specified size happens to match the actual size of the
		// cursor.  This is because normally, it seems that CopyImage() omits cursor animation
		// from the new object.  MSDN: "LR_COPYRETURNORG returns the original hImage if it satisfies
		// the criteria for the copy—that is, correct dimensions and color depth—in which case the
		// LR_COPYDELETEORG flag is ignored. If this flag is not specified, a new object is always created."
		hbitmap_new = (HBITMAP)CopyImage(hbitmap, aImageType, aWidth, aHeight, LR_COPYRETURNORG | LR_COPYDELETEORG);
		// Above's LR_COPYDELETEORG deletes the original to avoid cascading resource usage.  MSDN's
		// LoadImage() docs say:
		// "When you are finished using a bitmap, cursor, or icon you loaded without specifying the
		// LR_SHARED flag, you can release its associated memory by calling one of [the three functions]."
		// Therefore, it seems best to call the right function even though DeleteObject might work on
		// all of them on some or all current OSes.  UPDATE: Evidence indicates that DestroyIcon()
		// will also destroy cursors, probably because icons and cursors are literally identical in
		// every functional way.  One piece of evidence:
		//> No stack trace, but I know the exact source file and line where the call
		//> was made. But still, it is annoying when you see 'DestroyCursor' even though
		//> there is 'DestroyIcon'.
		// "Can't be helped. Icons and cursors are the same thing" (Tim Robinson (MVP, Windows SDK)).
		//
		// Finally, the reason this is important is that it eliminates one handle type
		// that we would otherwise have to track.  For example, if a gui window is destroyed and
		// and recreated multiple times, its bitmap and icon handles should all be destroyed each time.
		// Otherwise, resource usage would cascade upward until the script finally terminated, at
		// which time all such handles are freed automatically.
	}
	return hbitmap_new;
}



HBITMAP IconToBitmap(HICON ahIcon, bool aDestroyIcon)
// Converts HICON to an HBITMAP that has ahIcon's actual dimensions.
// The incoming ahIcon will be destroyed if the caller passes true for aDestroyIcon.
// Returns NULL on failure, in which case aDestroyIcon will still have taken effect.
// If the icon contains any transparent pixels, they will be mapped to CLR_NONE within
// the bitmap so that the caller can detect them.
{
	if (!ahIcon)
		return NULL;

	HBITMAP hbitmap = NULL;  // Set default.  This will be the value returned.

	HDC hdc_desktop = GetDC(HWND_DESKTOP);
	HDC hdc = CreateCompatibleDC(hdc_desktop); // Don't pass NULL since I think that would result in a monochrome bitmap.
	if (hdc)
	{
		ICONINFO ii;
		if (GetIconInfo(ahIcon, &ii))
		{
			BITMAP icon_bitmap;
			// Find out how big the icon is and create a bitmap compatible with the desktop DC (not the memory DC,
			// since its bits per pixel (color depth) is probably 1.
			if (GetObject(ii.hbmColor, sizeof(BITMAP), &icon_bitmap)
				&& (hbitmap = CreateCompatibleBitmap(hdc_desktop, icon_bitmap.bmWidth, icon_bitmap.bmHeight))) // Assign
			{
				// To retain maximum quality in case caller needs to resize the bitmap we return, convert the
				// icon to a bitmap that matches the icon's actual size:
				HGDIOBJ old_object = SelectObject(hdc, hbitmap);
				if (old_object) // Above succeeded.
				{
					// Use DrawIconEx() vs. DrawIcon() because someone said DrawIcon() always draws 32x32
					// regardless of the icon's actual size.
					// If it's ever needed, this can be extended so that the caller can pass in a background
					// color to use in place of any transparent pixels within the icon (apparently, DrawIconEx()
					// skips over transparent pixels in the icon when drawing to the DC and its bitmap):
					RECT rect = {0, 0, icon_bitmap.bmWidth, icon_bitmap.bmHeight}; // Left, top, right, bottom.
					HBRUSH hbrush = CreateSolidBrush(CLR_DEFAULT);
					FillRect(hdc, &rect, hbrush);
					DeleteObject(hbrush);
					// Probably something tried and abandoned: FillRect(hdc, &rect, (HBRUSH)GetStockObject(NULL_BRUSH));
					DrawIconEx(hdc, 0, 0, ahIcon, icon_bitmap.bmWidth, icon_bitmap.bmHeight, 0, NULL, DI_NORMAL);
					// Debug: Find out properties of new bitmap.
					//BITMAP b;
					//GetObject(hbitmap, sizeof(BITMAP), &b);
					SelectObject(hdc, old_object); // Might be needed (prior to deleting hdc) to prevent memory leak.
				}
			}
			// It's our reponsibility to delete these two when they're no longer needed:
			DeleteObject(ii.hbmColor);
			DeleteObject(ii.hbmMask);
		}
		DeleteDC(hdc);
	}
	ReleaseDC(HWND_DESKTOP, hdc_desktop);
	if (aDestroyIcon)
		DestroyIcon(ahIcon);
	return hbitmap;
}



HRESULT MySetWindowTheme(HWND hwnd, LPCWSTR pszSubAppName, LPCWSTR pszSubIdList)
{
	// The library must be loaded dynamically, otherwise the app will not launch on OSes older than XP.
	// Theme DLL is normally available only on XP+, but an attempt to load it is made unconditionally
	// in case older OSes can ever have it.
	HRESULT hresult = !S_OK; // Set default as "failure".
	HINSTANCE hinstTheme = LoadLibrary("uxtheme");
	if (hinstTheme)
	{
		typedef HRESULT (WINAPI *MySetWindowThemeType)(HWND, LPCWSTR, LPCWSTR);
  		MySetWindowThemeType DynSetWindowTheme = (MySetWindowThemeType)GetProcAddress(hinstTheme, "SetWindowTheme");
		if (DynSetWindowTheme)
			hresult = DynSetWindowTheme(hwnd, pszSubAppName, pszSubIdList);
		FreeLibrary(hinstTheme);
	}
	return hresult;
}



//HRESULT MyEnableThemeDialogTexture(HWND hwnd, DWORD dwFlags)
//{
//	// The library must be loaded dynamically, otherwise the app will not launch on OSes older than XP.
//	// Theme DLL is normally available only on XP+, but an attempt to load it is made unconditionally
//	// in case older OSes can ever have it.
//	HRESULT hresult = !S_OK; // Set default as "failure".
//	HINSTANCE hinstTheme = LoadLibrary("uxtheme");
//	if (hinstTheme)
//	{
//		typedef HRESULT (WINAPI *MyEnableThemeDialogTextureType)(HWND, DWORD);
//  		MyEnableThemeDialogTextureType DynEnableThemeDialogTexture = (MyEnableThemeDialogTextureType)GetProcAddress(hinstTheme, "EnableThemeDialogTexture");
//		if (DynEnableThemeDialogTexture)
//			hresult = DynEnableThemeDialogTexture(hwnd, dwFlags);
//		FreeLibrary(hinstTheme);
//	}
//	return hresult;
//}



char *ConvertEscapeSequences(char *aBuf, char aEscapeChar, bool aAllowEscapedSpace)
// Replaces any escape sequences in aBuf with their reduced equivalent.  For example, if aEscapeChar
// is accent, Each `n would become a literal linefeed.  aBuf's length should always be the same or
// lower than when the process started, so there is no chance of overflow.
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
			case 's': // space (not always allowed for backward compatibility reasons).
				if (aAllowEscapedSpace)
					*cp1 = ' ';
				//else do nothing extra, just let the standard action for unrecognized escape sequences.
				break;
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



bool IsStringInList(char *aStr, char *aList, bool aFindExactMatch, bool aCaseSensitive)
// Checks if aStr exists in aList (which is a comma-separated list).
// If aStr is blank, aList must start with a delimiting comma for there to be a match.
{
	// Must use a temp. buffer because otherwise there's no easy way to properly match upon strings
	// such as the following:
	// if var in string,,with,,literal,,commas
	char buf[LINE_SIZE];
    char *this_field = aList, *next_field, *cp;

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
				if (cp[1] == ',') // Make this pair into a single literal comma.
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

		if (*buf) // It is possible for this to be blank only for the first field.  Example: if var in ,abc
		{
			if (aFindExactMatch)
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
			else // Substring match
			{
				if (aCaseSensitive)
				{
					if (strstr(aStr, buf)) // Match found
						return true;
				}
				else // Not case sensitive
					if (strcasestr(aStr, buf)) // Match found
						return true;
			}
		}
		else // First item in the list is the empty string.
			if (aFindExactMatch) // In this case, this is a match if aStr is also blank.
			{
				if (!*aStr)
					return true;
			}
			else // Empty string is always found as a substring in any other string.
				return true;
		this_field = next_field;
	} // while()

	return false;  // No match found.
}