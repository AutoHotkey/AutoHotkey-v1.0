/*
AutoHotkey

Copyright 2003-2005 Chris Mallett

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
#include "var.h"
#include "globaldata.h" // for g_script


// Init static vars:
char Var::sEmptyString[] = ""; // Used to as a non-constant empty string so that callers can write to it.


ResultType Var::Assign(int aValueToAssign) // For some reason, these functions are actually faster when not inline.
{
	char value_string[256];
	// ITOA() seems to perform quite a bit better than sprintf() in this case:
	ITOA(aValueToAssign, value_string);
	return Assign(value_string);
	//snprintf(value_string, sizeof(value_string), "%d", aValueToAssign);
	//return Assign(value_string);
}



ResultType Var::Assign(DWORD aValueToAssign)
// Returns OK or FAIL.
{
	char value_string[256];
	UTOA(aValueToAssign, value_string)
	return Assign(value_string);
}



ResultType Var::Assign(__int64 aValueToAssign)
// Returns OK or FAIL.
{
	char value_string[256];
	ITOA64(aValueToAssign, value_string);
	return Assign(value_string);
}


// Currently not needed:
//ResultType Var::Assign(unsigned __int64 aValueToAssign)
//// Since most script features that "read in" a number from a string (such as addition and comparison)
//// read the strings as signed values, this function here should only be used when it is documented
//// that math and comparisons should not be performed on such values.
//// Returns OK or FAIL.
//{
//	char value_string[256];
//	UTOA64(aValueToAssign, value_string);
//	return Assign(value_string);
//}



ResultType Var::Assign(double aValueToAssign)
// It's best to call this method -- rather than manually converting to double -- so that the
// digits/formatting/precision is consistent throughout the program.
// Returns OK or FAIL.
{
	char value_string[MAX_FORMATTED_NUMBER_LENGTH + 1];
	snprintf(value_string, sizeof(value_string), g.FormatFloat, aValueToAssign); // "%0.6f"; %f can handle doubles in MSVC++.
	return Assign(value_string);
}



ResultType Var::Assign(char *aBuf, VarSizeType aLength, bool aTrimIt)
// Returns OK or FAIL.
// Any existing contents of this variable will be destroyed regardless of whether aBuf is NULL.
// Caller can omit both params to set a var to be empty-string, but in that case, if the variable
// is of large capacity, its memory will not be freed.  This is by design because it allows the
// caller to exploit its knowledge of whether the var's large capacity is likely to be needed
// again in the near future, thus reducing the expected amount of memory fragmentation.
// To explicitly free the memory, use Assign("").
// If aBuf is NULL, the variable will be set up to handle a string of at least aLength
// in length.  In addition, if the var is the clipboard, it will be prepared for writing.
{
	bool do_assign = true; // Set default.
	bool free_it_if_large = true;  // Default.
	if (!aBuf)
		if (aLength == VARSIZE_MAX) // Caller omitted this param too, so it wants to assign empty string.
		{
			free_it_if_large = false;
			aLength = 0;
		}
		else // Caller gave a NULL buffer to signal us to ensure the var is at least aLength in capacity.
			do_assign = false;
	else // Caller provided a non-NULL buffer.
		if (aLength == VARSIZE_MAX) // Caller wants us to determine its length.
			aLength = (VarSizeType)strlen(aBuf);
	if (!aBuf)
		aBuf = "";  // From here on, make sure it's the empty string for all uses.

	size_t space_needed = aLength + 1; // For the zero terminator.

	if (mType == VAR_CLIPBOARD)
		if (do_assign)
			// Just return the result of this.  Note: The clipboard var's attributes,
			// such as mLength, are not maintained because it's a variable whose
			// contents usually aren't under our control.  UPDATE: aTrimIt isn't
			// needed because the clipboard is never assigned something that needs
			// to be trimmed in this way (i.e. PerformAssign handles the trimming
			// on its own for the clipboard, due to the fact that dereferencing
			// into the temp buf is unnecessary when the clipboard is the target):
			return g_clip.Set(aBuf, aLength); //, aTrimIt);
		else
			// We open it for write now, because some caller's don't call
			// this function to write to the contents of the var, they
			// do it themselves.  Note: Below call will have displayed
			// any error that occurred:
			return g_clip.PrepareForWrite(space_needed) ? OK : FAIL;

	if (space_needed > g_MaxVarCapacity)
		return g_script.ScriptError(ERR_MEM_LIMIT_REACHED);

	if (space_needed <= 1) // Variable is being assigned the empty string (or a deref that resolves to it).
	{
		mLength = 0;
		switch (mHowAllocated)
		{
		case ALLOC_NONE:
			// Some things may rely on this being empty-string vs. NULL.
			// In addition, don't make it "" so that the caller's code can be simpler and
			// more maintainable, i.e. the caller can always do *buf = '\0' if it wants:
			mContents = sEmptyString;
			*mContents = '\0';  // Just in case someone else changed it to be non-zero.
			return OK;
		case ALLOC_SIMPLE:
			// Don't set to "" because then we'd have a memory leak.  i.e. once a var
			// becomes ALLOC_SIMPLE, it should never become ALLOC_NONE again:
			*mContents = '\0';
			return OK;
		case ALLOC_MALLOC:
			// Setting a var whose contents are very large to be nothing or blank is currently the
			// only way to free up the memory of that var.  Shrinking it dynamically seems like it
			// might introduce too much memory fragmentation and overhead (since in many cases,
			// it would likely need to grow back to its former size in the near future).  So we
			// only free relatively large vars:
			if (mCapacity > (4 * 1024) && free_it_if_large)
			{
				free(mContents);
				mCapacity = 0;
				mContents = sEmptyString; // See ALLOC_NONE above for explanation.
				*mContents = '\0';  // Just in case someone else changed it to be non-zero.
			}
			else
				if (mCapacity)
					// If capacity small, don't bother: It's probably not worth the added mem fragmentation.
					*mContents = '\0';
				// else it's the empty string (a constant), so don't attempt to free it or assign to it.
			// But do not change mHowAllocated to be ALLOC_NONE because it would cause a
			// a memory leak in this sequence of events:
			// var1 is assigned something short enough to make it ALLOC_SIMPLE
			// var1 is assigned something large enough to require malloc()
			// var1 is set to empty string and its mem is thus free()'d by the above.
			// var1 is once again assigned something short enough to make it ALLOC_SIMPLE
			// The last step above would be a problem because the 2nd ALLOC_SIMPLE can't
			// reclaim the spot in SimpleHeap that had been in use by the first.  In other
			// words, when a var makes the transition from ALLOC_SIMPLE to ALLOC_MALLOC,
			// its ALLOC_SIMPLE memory is lost to the system until the program exits.
			// But since this loss occurs at most once per distinct variable name,
			// it's not considered a memory leak because the loss can't exceed a fixed
			// amount regardless of how long the program runs.  The reason for all of this
			// is that allocating dynamic memory is costly: it causes system memory fragmentation,
			// (especially if a var were to be malloc'd and free'd thousands of times in a loop)
			// and small-sized mallocs have a large overhead: it's been said that every block
			// of dynamic mem, even those of size 1 or 2, incurs about 40 bytes of overhead.
			return OK;
		} // switch()
	} // if

	if (space_needed > mCapacity)
	{
		switch (mHowAllocated)
		{
		case ALLOC_NONE:
			if (space_needed <= MAX_ALLOC_SIMPLE)
			{
				if (   !(mContents = SimpleHeap::Malloc(MAX_ALLOC_SIMPLE))   )
					return FAIL; // It already displayed the error.
				mHowAllocated = ALLOC_SIMPLE;
				mCapacity = MAX_ALLOC_SIMPLE;
				break;
			}
			// else don't break, just fall through to the next case.
		case ALLOC_SIMPLE:
		case ALLOC_MALLOC:
		{
			// This case can happen even if space_needed is less than MAX_ALLOC_SIMPLE
			// because once a var becomes ALLOC_MALLOC, it should never change to
			// one of the other alloc modes.  See above comments for explanation.
			if (mCapacity && mHowAllocated == ALLOC_MALLOC)
				free(mContents);
			// else it's the empty string (a constant) or it points to memory on
			// SimpleHeap, so don't attempt to free it.
			size_t alloc_size = space_needed;
			// Allow a little room for future expansion to cut down on the number of
			// free's and malloc's we expect to have to do in the future for this var:
			if (alloc_size < MAX_PATH)
				alloc_size = MAX_PATH;  // An amount that will fit all standard filenames seems good.
			else if (alloc_size < (160 * 1024)) // MAX_PATH to 160 KB or less -> 10% extra.
				alloc_size = (size_t)(alloc_size * 1.1);
			else if (alloc_size < (1600 * 1024))  // 160 to 1600 KB -> 16 KB extra
				alloc_size += (16 * 1024); 
			else if (alloc_size < (6400 * 1024)) // 1600 to 6400 KB -> 1% extra
				alloc_size = (size_t)(alloc_size * 1.01);
			else  // 6400 KB or more: Cap the extra margin at some reasonable compromise of speed vs. mem usage: 64 KB
				alloc_size += (64 * 1024);
			if (alloc_size > g_MaxVarCapacity)
				alloc_size = g_MaxVarCapacity;  // which has already been verified to be enough.
			mHowAllocated = ALLOC_MALLOC; // This should be done first because even if the below fails, the old value of mContents (e.g. SimpleHeap) is lost.
			if (   !(mContents = (char *)malloc(alloc_size))   )
			{
				mCapacity = 0;  // Added for v1.0.25.
				return g_script.ScriptError(ERR_OUTOFMEM ERR_ABORT); // ERR_ABORT since an error is most likely to occur at runtime.
			}
			mCapacity = (VarSizeType)alloc_size;
			break;
		}
		} // switch()
	} // if

	if (do_assign)
	{
		// Use strlcpy() so that only a substring is copied, in case caller specified that:
		strlcpy(mContents, aBuf, space_needed);
		// Calculate the length explicitly in case aBuf was shorter than expected,
		// perhaps due to one of the functions that built part of it failing.
		// UPDATE: Sometimes, such as for StringMid, the caller explicitly wants
		// us to be able to handle the case where the actual length of aBuf is less
		// that aLength.
		// Also, use mContents rather than aBuf to calculate the length, in case
		// caller only wanted a substring of aBuf copied:

		// Only trim when the caller told us to, rather than always if g_script.mIsAutoIt2
		// is true, since AutoIt2 doesn't always trim things (e.g. FileReadLine probably
		// does not trim the line that was read into its output var).  UPDATE: This is
		// no longer needed because I think AutoIt2 only auto-trims when SetEnv is used.
		// UPDATE #2: Yes, it's needed in one case currently, so it's enabled again:
		if (aTrimIt)
			trim(mContents);
		mLength = (VarSizeType)strlen(mContents);
	}
	else
	{
		// We've done everything except the actual assignment.  Let the caller handle that.
		// Even the length will be set below to the expected length in case the caller
		// doesn't override this.
		// Below: Already verified that the length value will fit into VarSizeType.
		// Also, -1 because length is defined as not including zero terminator:
		mLength = (VarSizeType)(space_needed - 1);
		// Init for greater safety.  This is safe because in this case mContents is not a
		// constant memory area:
		*mContents = '\0';
	}

	return OK;
}



VarSizeType Var::Get(char *aBuf)
// Returns the length of this var's contents.  In addition, if aBuf isn't NULL, it will
// copy the contents into aBuf.
{
	// For v1.0.25, don't do this because in some cases the existing contents of aBuf will not
	// be altered.  Instead, it will be set to blank as needed further below.
	//if (aBuf) *aBuf = '\0';  // Init early to get it out of the way, in case of early return.

	char *aBuf_orig = aBuf;

	DWORD result;
	static DWORD timestamp_tick = 0, now_tick; // static should be thread + recursion safe in this case.
	static SYSTEMTIME st_static = {0};

	// Just a fake buffer to pass to some API functions in lieu of a NULL, to avoid
	// any chance of misbehavior:
	char buf_temp[1];  // Keep the size at 1 so that API functions will always fail to copy to buf.

	switch(mType)
	{
	case VAR_NORMAL:
		if (!mLength) // If the var is empty, check to see if it's really an env. var.
		{
			// Regardless of whether aBuf is NULL or not, we don't know at this stage
			// whether mName is the name of a valid environment variable.  Therefore,
			// GetEnvironmentVariable() is currently called once in the case where
			// aBuf is NULL and twice in the case where it's not.  There may be some
			// way to reduce it to one call always, but that is an optimization for
			// the future.  Another reason: Calling it twice seems safer, because we can't
			// be completely sure that it wouldn't act up our (possibly undersized) aBuf
			// buffer even if it doesn't find the env. var:
			if (   result = GetEnvironmentVariable(mName, buf_temp, sizeof(buf_temp))   )
			{
				// This env. var exists.
				if (!aBuf)
					return result - 1;  // since GetEnvironmentVariable() returns total size needed in this case.
				// The caller has ensured, probably via previous call to this function with aBuf == NULL,
				// that aBuf is large enough to hold the result.  Also, don't use a size greater than
				// 32767 because that will cause it to fail on Win95 (tested by Robert Yalkin).
				// (size probably must be under 64K, but that is untested.  Just stick with 32767)
				// According to MSDN, 32767 is exactly large enough to handle the largest variable plus
				// its zero terminator:
				aBuf += GetEnvironmentVariable(mName, aBuf, 32767);
				break;
			}
			else // No matching env. var.
			{
				if (aBuf)
					*aBuf = '\0';
				return 0;
			}
		}
		// otherwise, it has non-zero length (set by user), so it takes precedence over any existing env. var.
		if (!aBuf)
			return mLength;
		// Copy the var contents into aBuf.  Although a little bit slower than CopyMemory() for large
		// variables (say, over 100K), this loop seems much faster for small ones, which is the typical
		// case.  Also of note is that this code section is the main bottleneck for scripts that manipulate
		// large variables, such as this:
		//start_time = %A_TICKCOUNT%
		//my = 0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
		//tempvar =
		//loop, 2000
		//	tempvar = %tempvar%%My%
		//elapsed_time = %A_TICKCOUNT%
		//elapsed_time -= %start_time%
		//msgbox, elapsed_time = %elapsed_time%
		//return
		if (aBuf == mContents)
			// When we're called from ExpandArg() that was called from PerformAssign(), PerformAssign()
			// relies on this check to avoid the overhead of copying a variables contents onto itself.
			aBuf += mLength;
		else if (mLength > 100000)
		{
			// Faster for large vars, but large vars aren't typical:
			CopyMemory(aBuf, mContents, mLength);
			aBuf += mLength;
		}
		else
			for (char *cp = mContents; *cp; *aBuf++ = *cp++);
		break;

	// Built-in vars with volatile contents:
	case VAR_CLIPBOARD:
	{
		VarSizeType size = (VarSizeType)g_clip.Get(aBuf); // It will also copy into aBuf if it's non-NULL.
		if (size == CLIPBOARD_FAILURE)
		{
			// Above already displayed the error, so just return.
			// If we were called only to determine the size, the
			// next call to g_clip.Get() will not put anything into
			// aBuf (it will either fail again, or it will return
			// a length of zero due to the clipboard not already
			// being open & ready), so there's no danger of future
			// buffer overflows as a result of returning zero here.
			// Also, due to this function's return type, there's
			// no easy way to terminate the current hotkey
			// subroutine (or entire script) due to this error.
			// However, due to the fact that multiple attempts
			// are made to open the clipboard, failure should
			// be extremely rare.  And the user will be notified
			// with a MsgBox anyway, during which the subroutine
			// will be suspended:
			if (aBuf)
				*aBuf = '\0';  // Init early to get it out of the way, in case of early return.
			return 0;
		}
		if (!aBuf)
			return size;
		aBuf += size;
		break;
	}

	case VAR_TRUE:
	case VAR_FALSE:
		if (aBuf)
			*aBuf++ = (mType == VAR_TRUE) ? '1': '0'; // It will be terminated later below.
		else
			return 1; // The length of the value.
		break;

	case VAR_WORKINGDIR:
		// Use GetCurrentDirectory() vs. g_WorkingDir because any in-progrses FileSelectFile()
		// dialog is able to keep functioning even when it's quasi-thread is suspended.  The
		// dialog can thus change the current directory as seen by the active quasi-thread even
		// though g_WorkingDir hasn't been updated.  It might also be possible for the working
		// directory to change in unusual circumstances such as a network drive being lost):
		if (!aBuf)
		{
			result = GetCurrentDirectory(sizeof(buf_temp), buf_temp);
			if (!result)
			{
				// This is just a warning because this function isn't set up to cause a true
				// failure.  So don't append ERR_ABORT to the below string:
				g_script.ScriptError("GetCurrentDirectory."); // Short msg since probably can't realistically fail.
				// Probably safer to return something so that caller reserves enough space for it
				// in case the call works the next time?:
				return MAX_PATH;
			}
			// Return values are done according to the documented behavior of GetCurrentDirectory():
			if (result > sizeof(buf_temp))
				return result - 1;
			return result;  // Should never be reached.
		}
		// Otherwise:
		aBuf += GetCurrentDirectory(9999, aBuf);  // Caller has already ensured it's large enough.
		break;

	case VAR_BATCHLINES: if (!aBuf) return g_script.GetBatchLines(); aBuf += g_script.GetBatchLines(aBuf); break;
	case VAR_TITLEMATCHMODE: if (!aBuf) return g_script.GetTitleMatchMode(); else aBuf += g_script.GetTitleMatchMode(aBuf); break;
	case VAR_TITLEMATCHMODESPEED: if (!aBuf) return g_script.GetTitleMatchModeSpeed(); else aBuf += g_script.GetTitleMatchModeSpeed(aBuf); break;
	case VAR_DETECTHIDDENWINDOWS: if (!aBuf) return g_script.GetDetectHiddenWindows(); else aBuf += g_script.GetDetectHiddenWindows(aBuf); break;
	case VAR_DETECTHIDDENTEXT: if (!aBuf) return g_script.GetDetectHiddenText(); else aBuf += g_script.GetDetectHiddenText(aBuf); break;
	case VAR_AUTOTRIM: if (!aBuf) return g_script.GetAutoTrim(); else aBuf += g_script.GetAutoTrim(aBuf); break;
	case VAR_STRINGCASESENSE: if (!aBuf) return g_script.GetStringCaseSense(); else aBuf += g_script.GetStringCaseSense(aBuf); break;
	case VAR_FORMATINTEGER: if (!aBuf) return g_script.GetFormatInteger(); else aBuf += g_script.GetFormatInteger(aBuf); break;
	case VAR_FORMATFLOAT: if (!aBuf) return g_script.GetFormatFloat(); else aBuf += g_script.GetFormatFloat(aBuf); break;
	case VAR_KEYDELAY: if (!aBuf) return g_script.GetKeyDelay(); else aBuf += g_script.GetKeyDelay(aBuf); break;
	case VAR_WINDELAY: if (!aBuf) return g_script.GetWinDelay(); else aBuf += g_script.GetWinDelay(aBuf); break;
	case VAR_CONTROLDELAY: if (!aBuf) return g_script.GetControlDelay(); else aBuf += g_script.GetControlDelay(aBuf); break;
	case VAR_MOUSEDELAY: if (!aBuf) return g_script.GetMouseDelay(); else aBuf += g_script.GetMouseDelay(aBuf); break;
	case VAR_DEFAULTMOUSESPEED: if (!aBuf) return g_script.GetDefaultMouseSpeed(); else aBuf += g_script.GetDefaultMouseSpeed(aBuf); break;

	case VAR_ICONHIDDEN: if (!aBuf) return g_script.GetIconHidden(); else aBuf += g_script.GetIconHidden(aBuf); break;
	case VAR_ICONTIP: if (!aBuf) return g_script.GetIconTip(); else aBuf += g_script.GetIconTip(aBuf); break;
	case VAR_ICONFILE: if (!aBuf) return g_script.GetIconFile(); else aBuf += g_script.GetIconFile(aBuf); break;
	case VAR_ICONNUMBER: if (!aBuf) return g_script.GetIconNumber(); else aBuf += g_script.GetIconNumber(aBuf); break;

	case VAR_EXITREASON: if (!aBuf) return g_script.GetExitReason(); aBuf += g_script.GetExitReason(aBuf); break;

	case VAR_OSTYPE: if (!aBuf) return g_script.GetOSType(); aBuf += g_script.GetOSType(aBuf); break;
	case VAR_OSVERSION: if (!aBuf) return g_script.GetOSVersion(); aBuf += g_script.GetOSVersion(aBuf); break;
	case VAR_LANGUAGE: if (!aBuf) return g_script.GetLanguage(); aBuf += g_script.GetLanguage(aBuf); break;
	case VAR_COMPUTERNAME: if (!aBuf) return g_script.GetUserOrComputer(false); aBuf += g_script.GetUserOrComputer(false, aBuf); break;
	case VAR_USERNAME: if (!aBuf) return g_script.GetUserOrComputer(true); aBuf += g_script.GetUserOrComputer(true, aBuf); break;

	case VAR_WINDIR: if (!aBuf) return GetWindowsDirectory(buf_temp, 0) - 1; aBuf += GetWindowsDirectory(aBuf, MAX_PATH); break;  // Sizes/lengths/-1/etc. verified correct.
	case VAR_PROGRAMFILES: if (!aBuf) return g_script.GetProgramFiles(); aBuf += g_script.GetProgramFiles(aBuf); break;
	case VAR_DESKTOP: if (!aBuf) return g_script.GetDesktop(false); aBuf += g_script.GetDesktop(false, aBuf); break;
	case VAR_DESKTOPCOMMON: if (!aBuf) return g_script.GetDesktop(true); aBuf += g_script.GetDesktop(true, aBuf); break;
	case VAR_STARTMENU: if (!aBuf) return g_script.GetStartMenu(false); aBuf += g_script.GetStartMenu(false, aBuf); break;
	case VAR_STARTMENUCOMMON: if (!aBuf) return g_script.GetStartMenu(true); aBuf += g_script.GetStartMenu(true, aBuf); break;
	case VAR_PROGRAMS: if (!aBuf) return g_script.GetPrograms(false); aBuf += g_script.GetPrograms(false, aBuf); break;
	case VAR_PROGRAMSCOMMON: if (!aBuf) return g_script.GetPrograms(true); aBuf += g_script.GetPrograms(true, aBuf); break;
	case VAR_STARTUP: if (!aBuf) return g_script.GetStartup(false); aBuf += g_script.GetStartup(false, aBuf); break;
	case VAR_STARTUPCOMMON: if (!aBuf) return g_script.GetStartup(true); aBuf += g_script.GetStartup(true, aBuf); break;
	case VAR_MYDOCUMENTS: if (!aBuf) return g_script.GetMyDocuments(); aBuf += g_script.GetMyDocuments(aBuf); break;

	case VAR_ISADMIN: if (!aBuf) return g_script.GetIsAdmin(); aBuf += g_script.GetIsAdmin(aBuf); break;
	case VAR_CURSOR: if (!aBuf) return g_script.ScriptGetCursor(); aBuf += g_script.ScriptGetCursor(aBuf); break;
	case VAR_CARETX: if (!aBuf) return g_script.ScriptGetCaret(VAR_CARETX); aBuf += g_script.ScriptGetCaret(VAR_CARETX, aBuf); break;
	case VAR_CARETY: if (!aBuf) return g_script.ScriptGetCaret(VAR_CARETY); aBuf += g_script.ScriptGetCaret(VAR_CARETY, aBuf); break;
	case VAR_SCREENWIDTH: if (!aBuf) return g_script.GetScreenWidth(); aBuf += g_script.GetScreenWidth(aBuf); break;
	case VAR_SCREENHEIGHT: if (!aBuf) return g_script.GetScreenHeight(); aBuf += g_script.GetScreenHeight(aBuf); break;

	case VAR_IPADDRESS1:
	case VAR_IPADDRESS2:
	case VAR_IPADDRESS3:
	case VAR_IPADDRESS4:
		if (!aBuf) return g_script.GetIP(mType - VAR_IPADDRESS1); aBuf += g_script.GetIP(mType - VAR_IPADDRESS1, aBuf); break;

	case VAR_SCRIPTNAME: if (!aBuf) return g_script.GetFilename(); else aBuf += g_script.GetFilename(aBuf); break;
	case VAR_SCRIPTDIR: if (!aBuf) return g_script.GetFileDir(); else aBuf += g_script.GetFileDir(aBuf); break;
	case VAR_SCRIPTFULLPATH: if (!aBuf) return g_script.GetFilespec(); else aBuf += g_script.GetFilespec(aBuf); break;

	case VAR_LOOPFILENAME: if (!aBuf) return g_script.GetLoopFileName(); else aBuf += g_script.GetLoopFileName(aBuf); break;
	case VAR_LOOPFILESHORTNAME: if (!aBuf) return g_script.GetLoopFileShortName(); else aBuf += g_script.GetLoopFileShortName(aBuf); break;
	case VAR_LOOPFILEDIR: if (!aBuf) return g_script.GetLoopFileDir(); else aBuf += g_script.GetLoopFileDir(aBuf); break;
	case VAR_LOOPFILEFULLPATH: if (!aBuf) return g_script.GetLoopFileFullPath(); else aBuf += g_script.GetLoopFileFullPath(aBuf); break;
	case VAR_LOOPFILELONGPATH: if (!aBuf) return g_script.GetLoopFileLongPath(); else aBuf += g_script.GetLoopFileLongPath(aBuf); break;
	case VAR_LOOPFILESHORTPATH: if (!aBuf) return g_script.GetLoopFileShortPath(); else aBuf += g_script.GetLoopFileShortPath(aBuf); break;
	case VAR_LOOPFILETIMEMODIFIED: if (!aBuf) return g_script.GetLoopFileTimeModified(); else aBuf += g_script.GetLoopFileTimeModified(aBuf); break;
	case VAR_LOOPFILETIMECREATED: if (!aBuf) return g_script.GetLoopFileTimeCreated(); else aBuf += g_script.GetLoopFileTimeCreated(aBuf); break;
	case VAR_LOOPFILETIMEACCESSED: if (!aBuf) return g_script.GetLoopFileTimeAccessed(); else aBuf += g_script.GetLoopFileTimeAccessed(aBuf); break;
	case VAR_LOOPFILEATTRIB: if (!aBuf) return g_script.GetLoopFileAttrib(); else aBuf += g_script.GetLoopFileAttrib(aBuf); break;
	case VAR_LOOPFILESIZE: if (!aBuf) return g_script.GetLoopFileSize(NULL, 0); else aBuf += g_script.GetLoopFileSize(aBuf, 0); break;
	case VAR_LOOPFILESIZEKB: if (!aBuf) return g_script.GetLoopFileSize(NULL, 1024); else aBuf += g_script.GetLoopFileSize(aBuf, 1024); break;
	case VAR_LOOPFILESIZEMB: if (!aBuf) return g_script.GetLoopFileSize(NULL, 1024*1024); else aBuf += g_script.GetLoopFileSize(aBuf, 1024*1024); break;

	case VAR_LOOPREGTYPE: if (!aBuf) return g_script.GetLoopRegType(); else aBuf += g_script.GetLoopRegType(aBuf); break;
	case VAR_LOOPREGKEY: if (!aBuf) return g_script.GetLoopRegKey(); else aBuf += g_script.GetLoopRegKey(aBuf); break;
	case VAR_LOOPREGSUBKEY: if (!aBuf) return g_script.GetLoopRegSubKey(); else aBuf += g_script.GetLoopRegSubKey(aBuf); break;
	case VAR_LOOPREGNAME: if (!aBuf) return g_script.GetLoopRegName(); else aBuf += g_script.GetLoopRegName(aBuf); break;
	case VAR_LOOPREGTIMEMODIFIED: if (!aBuf) return g_script.GetLoopRegTimeModified(); else aBuf += g_script.GetLoopRegTimeModified(aBuf); break;

	case VAR_LOOPREADLINE: if (!aBuf) return g_script.GetLoopReadLine(); else aBuf += g_script.GetLoopReadLine(aBuf); break;
	case VAR_LOOPFIELD: if (!aBuf) return g_script.GetLoopField(); else aBuf += g_script.GetLoopField(aBuf); break;
	case VAR_INDEX: if (!aBuf) return g_script.GetLoopIndex(); else aBuf += g_script.GetLoopIndex(aBuf); break;

	case VAR_THISMENUITEM: if (!aBuf) return g_script.GetThisMenuItem(); else aBuf += g_script.GetThisMenuItem(aBuf); break;
	case VAR_THISMENUITEMPOS: if (!aBuf) return g_script.GetThisMenuItemPos(); else aBuf += g_script.GetThisMenuItemPos(aBuf); break;
	case VAR_THISMENU: if (!aBuf) return g_script.GetThisMenu(); else aBuf += g_script.GetThisMenu(aBuf); break;
	case VAR_THISHOTKEY: if (!aBuf) return g_script.GetThisHotkey(); else aBuf += g_script.GetThisHotkey(aBuf); break;
	case VAR_PRIORHOTKEY: if (!aBuf) return g_script.GetPriorHotkey(); else aBuf += g_script.GetPriorHotkey(aBuf); break;
	case VAR_TIMESINCETHISHOTKEY: if (!aBuf) return g_script.GetTimeSinceThisHotkey(); else aBuf += g_script.GetTimeSinceThisHotkey(aBuf); break;
	case VAR_TIMESINCEPRIORHOTKEY: if (!aBuf) return g_script.GetTimeSincePriorHotkey(); else aBuf += g_script.GetTimeSincePriorHotkey(aBuf); break;
	case VAR_ENDCHAR: if (!aBuf) return g_script.GetEndChar(); else aBuf += g_script.GetEndChar(aBuf); break;

	case VAR_GUICONTROL: if (!aBuf) return g_script.GetGuiControl(); else aBuf += g_script.GetGuiControl(aBuf); break;
	case VAR_GUICONTROLEVENT: if (!aBuf) return g_script.GetGuiControlEvent(); else aBuf += g_script.GetGuiControlEvent(aBuf); break;
	case VAR_GUI:
	case VAR_GUIWIDTH:
	case VAR_GUIHEIGHT: // All of the above use the below.
		if (!aBuf) return g_script.GetGui(mType); else aBuf += g_script.GetGui(mType, aBuf); break;

	case VAR_TIMEIDLE: if (!aBuf) return g_script.GetTimeIdle(); else aBuf += g_script.GetTimeIdle(aBuf); break;
	case VAR_TIMEIDLEPHYSICAL: if (!aBuf) return g_script.GetTimeIdlePhysical(); else aBuf += g_script.GetTimeIdlePhysical(aBuf); break;

	// This one is done this way, rather than using an escape sequence such as `s, because the escape
	// sequence method doesn't work (probably because `s resolves to a space and is that trimmed at
	// some point in process prior to when it can be used):
	case VAR_TAB:
	case VAR_SPACE: if (!aBuf) return g_script.GetSpace(mType); else aBuf += g_script.GetSpace(mType, aBuf); break;
	case VAR_AHKVERSION: if (!aBuf) return g_script.GetAhkVersion(); else aBuf += g_script.GetAhkVersion(aBuf); break;

	case VAR_MMMM:   if (!aBuf) return g_script.GetMMMM(); else aBuf += g_script.GetMMMM(aBuf); break;
	case VAR_MMM:  if (!aBuf) return g_script.GetMMM(); else aBuf += g_script.GetMMM(aBuf); break;
	case VAR_DDDD:  if (!aBuf) return g_script.GetDDDD(); else aBuf += g_script.GetDDDD(aBuf); break;
	case VAR_DDD: if (!aBuf) return g_script.GetDDD(); else aBuf += g_script.GetDDD(aBuf); break;

	case VAR_TICKCOUNT: if (!aBuf) return g_script.MyGetTickCount(); else aBuf += g_script.MyGetTickCount(aBuf); break;
	case VAR_NOW: if (!aBuf) return g_script.GetNow(); else aBuf += g_script.GetNow(aBuf); break;
	case VAR_NOWUTC: if (!aBuf) return g_script.GetNowUTC(); else aBuf += g_script.GetNowUTC(aBuf); break;

	case VAR_YYYY: if (!aBuf) return 4; // else fall through, which admittedly is somewhat inefficient here.
	case VAR_MM:
	case VAR_DD:
	case VAR_HOUR:
	case VAR_MIN:
	case VAR_SEC:   if (!aBuf) return 2; // length 2 for this and the above.
	case VAR_WDAY:  if (!aBuf) return 1; // else fall through.
	case VAR_YWEEK: if (!aBuf) return 6; // else fall through.
	case VAR_YDAY:  if (!aBuf) return 3; // Always return maximum allowed length as the estimate.
		// The current time is refreshed only if it's been a certain number of milliseconds since
		// the last fetch of one of these built-in time variables.  This keeps the variables in
		// sync with one another when they are used consecutively such as this example:
		// Var = %A_Hour%:%A_Min%:%A_Sec%
		// Using GetTickCount() because it's very low overhead compared to the other time functions:
		now_tick = GetTickCount();
		if (now_tick - timestamp_tick > 50 || !st_static.wYear)
		{
			GetLocalTime(&st_static);
			timestamp_tick = now_tick;
		}
		switch (mType)
		{
		case VAR_YYYY:  aBuf += sprintf(aBuf, "%d", st_static.wYear); break;
		case VAR_MM:    aBuf += sprintf(aBuf, "%02d", st_static.wMonth); break;
		case VAR_DD:    aBuf += sprintf(aBuf, "%02d", st_static.wDay); break;
		case VAR_HOUR:  aBuf += sprintf(aBuf, "%02d", st_static.wHour); break;
		case VAR_MIN:   aBuf += sprintf(aBuf, "%02d", st_static.wMinute); break;
		case VAR_SEC:   aBuf += sprintf(aBuf, "%02d", st_static.wSecond); break;
		case VAR_WDAY:  aBuf += sprintf(aBuf, "%d", st_static.wDayOfWeek + 1); break;
		case VAR_YWEEK:
			aBuf += GetISOWeekNumber(aBuf, st_static.wYear
				, GetYDay(st_static.wMonth, st_static.wDay, IS_LEAP_YEAR(st_static.wYear))
				, st_static.wDayOfWeek);
			break;
		case VAR_YDAY:
			aBuf += sprintf(aBuf, "%d", GetYDay(st_static.wMonth, st_static.wDay, IS_LEAP_YEAR(st_static.wYear)));
			break;
		} // inner switch()
		break; // outer switch()
	}

	*aBuf = '\0'; // Terminate the buffer in case above didn't.  Not counted as part of the length, so don't increment aBuf.
	return (VarSizeType)(aBuf - aBuf_orig);  // This is how many were written, not including the zero terminator.
}



ResultType Var::ValidateName(char *aName, bool aIsRuntime, bool aDisplayError)
// Returns OK or FAIL.
{
	if (!*aName) return FAIL;
	// Seems best to disallow variables that start with numbers so that more advanced
	// parsing (e.g. expressions) can be more easily added in the future.  UPDATE: For
	// backward compatibility with AutoIt2's technique of storing command line args in
	// a numically-named var (e.g. %1% is the first arg), decided not to do this:
	//if (*aName >= '0' && *aName <= '9')
	//	return g_script.ScriptError("This variable name starts with a number, which is not allowed.", aName);
	for (char c, *cp = aName; *cp; ++cp)
	{
		// ispunct(): @ # $ _ [ ] ? ! % & " ' ( ) * + - ^ . / \ : ; , < = > ` ~ | { }
		// Of the above, it seems best to disallow the following:
		// () future: user-defined functions, e.g. var = myfunc2(myfunc1() + 2)
		// ! future "not"
		// % reserved: deref char
		// & future: "and" or "bitwise and"
		// ' seems too iffy
		// " seems too iffy
		// */-+ reserved: math
		// , reserved: delimiter
		// . future: structs
		// : seems too iffy
		// ; reserved: comment
		// \ seems too iffy
		// <=> reserved: comparison
		// ^ future: "exp/power"
		// ` reserved: escape char
		// {} reserved: blocks
		// | future: "or" or "bitwise or"
		// ~ future: "bitwise not"
		c = *cp;  // For performance.
		if (IS_SPACE_OR_TAB(c) || c == g_delimiter || c == g_DerefChar
			|| c == '!' || c == '%' || c == '&' || c == '"' || c == '\'' || c == '(' || c == ')'
			|| c == '*' || c == '+' || c == '-' || c == '^' || c == '.' || c == '/' || c == '\\'
			|| c == ':' || c == ';' || c == ',' || c == '<' || c == '=' || c == '>'
			|| c == '`' || c == '~' || c == '|' || c == '{' || c == '}')
			if (aDisplayError)
				if (aIsRuntime)
					return g_script.ScriptError("This variable name contains an illegal character." ERR_ABORT, aName);
				else
					return g_script.ScriptError("This variable name contains an illegal character.", aName);
			else
				return FAIL;
	}
	// Otherwise:
	return OK;
}
