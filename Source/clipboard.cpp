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
#include "clipboard.h"
#include "globaldata.h"  // for g_script.ScriptError() and g_ClipboardTimeout
#include "application.h" // for MsgSleep()
#include "util.h" // for strlcpy()

size_t Clipboard::Get(char *aBuf)
// If aBuf is NULL, it returns the length of the text on the clipboard and leaves the
// clipboard open.  Otherwise, it copies the clipboard text into aBuf and closes
// the clipboard (UPDATE: But only if the clipboard is still open from a prior call
// to determine the length -- see later comments for details).  In both cases, the
// length of the clipboard text is returned (or the value CLIPBOARD_FAILURE if error).
// If the clipboard is still open when the next MsgSleep() is called -- presumably
// because the caller never followed up with a second call to this function, perhaps
// due to having insufficient memory -- MsgSleep() will close it so that our
// app doesn't keep the clipboard tied up.  Note: In all current cases, the caller
// will use MsgBox to display an error, which in turn calls MsgSleep(), which will
// immediately close the clipboard.
{
	// Seems best to always have done this even if we return early due to failure:
	if (aBuf)
		// It should be safe to do this even at its peak capacity, because caller
		// would have then given us the last char in the buffer, which is already
		// a zero terminator, so this would have no effect:
		*aBuf = '\0';

	UINT i, file_count = 0;
	BOOL clipboard_contains_text = IsClipboardFormatAvailable(CF_TEXT);
	BOOL clipboard_contains_files = IsClipboardFormatAvailable(CF_HDROP);
	if (!(clipboard_contains_text || clipboard_contains_files))
		return 0;

	if (!mIsOpen)
	{
		// As a precaution, don't give the caller anything from the clipboard
		// if the clipboard isn't already open from the caller's previous
		// call to determine the size of what's on the clipboard (no other app
		// can alter its size while we have it open).  The is to prevent a
		// buffer overflow from happening in a scenario such as the following:
		// Caller calls us and we return zero size, either because there's no
		// CF_TEXT on the clipboard orthere was a problem opening the clipboard.
		// In these two cases, the clipboard isn't open, so by the time the
		// caller calls us again, there's a chance (vanishingly small perhaps)
		// that another app (if our thread were preempted long enough, or the
		// platform is multiprocessor) will have changed the contents of the
		// clipboard to something larger than zero.  Thus, if we copy that
		// into the caller's buffer, the buffer might overflow:
		if (aBuf)
			return 0;
		if (!Open())
		{
			// Since this should be very rare, a shorter message is now used.  Formerly, it was
			// "Could not open clipboard for reading after many timed attempts. Another program is probably holding it open."
			Close(CANT_OPEN_CLIPBOARD_READ);
			return CLIPBOARD_FAILURE;
		}
		// GetClipboardData() apparently fails when the text on the clipboard is greater than a certain size
		// (Even though GetLastError() reports "Operation completed successfully").  The data size at which
		// this occurs is somewhere between 20 to 96 MB (perhaps depending on system's memory and CPU speed).
		if (   !(mClipMemNow = GetClipboardData(clipboard_contains_files ? CF_HDROP : CF_TEXT))   )
		{
			// Fix for v1.0.31.02: GetClipboardData() apparently fails when there are files on the clipboard
			// but either: 1) zero of them; 2) the CF_HDROP on the clipboard is somehow misformatted.
			// If you select the parent ".." folder in WinRar then use the following hotkey, the script
			// would previously yield a runtime error:
			//#q::
			//Send, ^c
			//ClipWait, 0.5, 1
			//msgbox %Clipboard%
			//Return
			if (clipboard_contains_files) // Tolerate failure in this case, see above.
			{
				Close();
				if (aBuf)
					*aBuf = '\0';
				return 0;
			}
			Close("GetClipboardData"); // Short error message since so rare.
			return CLIPBOARD_FAILURE;
		}
		// Although GlobalSize(mClipMemNow) can yield zero in some cases -- in which case GlobalLock() should
		// not be attempted -- it probably can't yield zero for CF_HDROP and CF_TEXT because such a thing has
		// never been reported by anyone.  Therefore, GlobalSize() is currently not called.
		if (   !(mClipMemNowLocked = (char *)GlobalLock(mClipMemNow))   )
		{
			Close("GlobalLock");  // Short error message since so rare.
			return CLIPBOARD_FAILURE;
		}
		// Otherwise: Update length after every successful new open&lock:
		// Determine the length (size - 1) of the buffer than would be
		// needed to hold what's on the clipboard:
		if (clipboard_contains_files)
		{
			if (file_count = DragQueryFile((HDROP)mClipMemNowLocked, 0xFFFFFFFF, "", 0))
			{
				mLength = (file_count - 1) * 2;  // Init; -1 if don't want a newline after last file.
				for (i = 0; i < file_count; ++i)
					mLength += DragQueryFile((HDROP)mClipMemNowLocked, i, NULL, 0);
			}
			else
				mLength = 0;
		}
		else // clipboard_contains_text
			mLength = strlen(mClipMemNowLocked);
		if (mLength >= CLIPBOARD_FAILURE) // Can't realistically happen, so just indicate silent failure.
			return CLIPBOARD_FAILURE;
	}
	if (!aBuf)
		return mLength;
		// Above: Just return the length; don't close the clipboard because we expect
		// to be called again soon.  If for some reason we aren't called, MsgSleep()
		// will automatically close the clipboard and clean up things.  It's done this
		// way to avoid the chance that the clipboard contents (and thus its length)
		// will change while we don't have it open, possibly resulting in a buffer
		// overflow.  In addition, this approach performs better because it avoids
		// the overhead of having to close and reopen the clipboard.

	// Otherwise:
	if (clipboard_contains_files)
	{
		if (file_count = DragQueryFile((HDROP)mClipMemNowLocked, 0xFFFFFFFF, "", 0))
			for (i = 0; i < file_count; ++i)
			{
				// Caller has already ensured aBuf is large enough to hold them all:
				aBuf += DragQueryFile((HDROP)mClipMemNowLocked, i, aBuf, 999);
				if (i < file_count - 1) // i.e. don't add newline after the last filename.
				{
					*aBuf++ = '\r';  // These two are the proper newline sequence that the OS prefers.
					*aBuf++ = '\n';
				}
				//else DragQueryFile() has ensured that aBuf is terminated.
			}
		// else aBuf has already been terminated upon entrance to this function.
	}
	else
		strcpy(aBuf, mClipMemNowLocked);  // Caller has already ensured that aBuf is large enough.
	// Fix for v1.0.37: Close() is no longer called here because it prevents the clipboard variable
	// from being referred to more than once in a line.  For example:
	// Msgbox %Clipboard%%Clipboard%
	// ToolTip % StrLen(Clipboard) . Clipboard
	// Instead, the clipboard is later closed in other places (search on CLOSE_CLIPBOARD_IF_OPEN
	// to find them).  The alternative to fixing it this way would be to let it reopen the clipboard
	// by means getting rid of the following lines above:
	//if (aBuf)
	//	return 0;
	// However, that has the risks described in the comments above those two lines.
	return mLength;
}



ResultType Clipboard::Set(char *aBuf, UINT aLength) //, bool aTrimIt)
// Returns OK or FAIL.
{
	// It was already open for writing from a prior call.  Return failure because callers that do this
	// are probably handling things wrong:
	if (IsReadyForWrite()) return FAIL;

	if (!aBuf)
	{
		aBuf = "";
		aLength = 0;
	}
	else
		if (aLength == UINT_MAX) // Caller wants us to determine the length.
			aLength = (UINT)strlen(aBuf);

	if (aLength)
	{
		if (!PrepareForWrite(aLength + 1))
			return FAIL;  // It already displayed the error.
		strlcpy(mClipMemNewLocked, aBuf, aLength + 1);  // Copy only a substring, if aLength specifies such.
		// Only trim when the caller told us to, rather than always if g_script.mIsAutoIt2
		// is true, since AutoIt2 doesn't always trim things (e.g. FileReadLine probably
		// does not trim the line that was read into its output var).  UPDATE: This is
		// no longer needed because I think AutoIt2 only auto-trims when SetEnv is used:
		//if (aTrimIt)
		//	trim(mClipMemNewLocked);
	}
	// else just do the below to empty the clipboard, which is different than setting
	// the clipboard equal to the empty string: it's not truly empty then, as reported
	// by IsClipboardFormatAvailable(CF_TEXT) -- and we want to be able to make it truly
	// empty for use with functions such as ClipWait:
	return Commit();  // It will display any errors.
}



char *Clipboard::PrepareForWrite(size_t aAllocSize)
{
	if (!aAllocSize) return NULL; // Caller should ensure that size is at least 1, i.e. room for the zero terminator.
	if (IsReadyForWrite())
		// It was already prepared due to a prior call.  Currently, the most useful thing to do
		// here is return the memory area that's already been reserved:
		return mClipMemNewLocked;
	// Note: I think GMEM_DDESHARE is recommended in addition to the usual GMEM_MOVEABLE:
	// UPDATE: MSDN: "The following values are obsolete, but are provided for compatibility
	// with 16-bit Windows. They are ignored.": GMEM_DDESHARE
	if (   !(mClipMemNew = GlobalAlloc(GMEM_MOVEABLE, aAllocSize))   )
	{
		g_script.ScriptError("GlobalAlloc");  // Short error message since so rare.
		return NULL;
	}
	if (   !(mClipMemNewLocked = (char *)GlobalLock(mClipMemNew))   )
	{
		mClipMemNew = GlobalFree(mClipMemNew);  // This keeps mClipMemNew in sync with its state.
		g_script.ScriptError("GlobalLock"); // Short error message since so rare.
		return NULL;
	}
	mCapacity = (UINT)aAllocSize; // Keep mCapacity in sync with the state of mClipMemNewLocked.
	*mClipMemNewLocked = '\0'; // Init for caller.
	return mClipMemNewLocked;  // The caller can now write to this mem.
}



ResultType Clipboard::Commit(UINT aFormat)
// If this is called while mClipMemNew is NULL, the clipboard will be set to be truly
// empty, which is different from writing an empty string to it.  Note: If the clipboard
// was already physically open, this function will close it as part of the commit (since
// whoever had it open before can't use the prior contents, since they're invalid).
{
	if (!mIsOpen && !Open())
		// Since this should be very rare, a shorter message is now used.  Formerly, it was
		// "Could not open clipboard for writing after many timed attempts.  Another program is probably holding it open."
		return AbortWrite(CANT_OPEN_CLIPBOARD_WRITE);
	if (!EmptyClipboard())
	{
		Close();
		return AbortWrite("EmptyClipboard"); // Short error message since so rare.
	}
	if (mClipMemNew)
	{
		bool new_is_empty = false;
		// Unlock prior to calling SetClipboardData:
		if (mClipMemNewLocked) // probably always true if we're here.
		{
			// Best to access the memory while it's still locked, which is why this temp var is used:
			// v1.0.40.02: The following was fixed to properly recognize 0x0000 as the Unicode string terminator,
			// which fixes problems with Transform Unicode.
			new_is_empty = !mClipMemNewLocked[0] && (aFormat != CF_UNICODETEXT || !mClipMemNewLocked[1]);
			GlobalUnlock(mClipMemNew); // mClipMemNew not mClipMemNewLocked.
			mClipMemNewLocked = NULL;  // Keep this in sync with the above action.
			mCapacity = 0; // Keep mCapacity in sync with the state of mClipMemNewLocked.
		}
		if (new_is_empty)
			// Leave the clipboard truly empty rather than setting it to be the
			// empty string (i.e. these two conditions are NOT the same).
			// But be sure to free the memory since we're not giving custody
			// of it to the system:
			mClipMemNew = GlobalFree(mClipMemNew);
		else
			if (SetClipboardData(aFormat, mClipMemNew))
				// In any of the failure conditions above, Close() ensures that mClipMemNew is
				// freed if it was allocated.  But now that we're here, the memory should not be
				// freed because it is owned by the clipboard (it will free it at the appropriate time).
				// Thus, we relinquish the memory because we shouldn't be looking at it anymore:
				mClipMemNew = NULL;
			else
			{
				Close();
				return AbortWrite("SetClipboardData"); // Short error message since so rare.
			}
	}
	// else we will close it after having done only the EmptyClipboard(), above.
	// Note: Decided not to update mLength for performance reasons (in case clipboard is huge).
	// Anyway, it seems rather pointless because once the clipboard is closed, our app instantly
	// loses sight of how large it is, so the the value of mLength wouldn't be reliable unless
	// the clipboard were going to be immediately opened again.
	return Close();
}



ResultType Clipboard::AbortWrite(char *aErrorMessage)
// Always returns FAIL.
{
	// Since we were called in conjunction with an aborted attempt to Commit(), always
	// ensure the clipboard is physically closed because even an attempt to Commit()
	// should physically close it:
	Close();
	if (mClipMemNewLocked)
	{
		GlobalUnlock(mClipMemNew); // mClipMemNew not mClipMemNewLocked.
		mClipMemNewLocked = NULL;
		mCapacity = 0; // Keep mCapacity in sync with the state of mClipMemNewLocked.
	}
	// Above: Unlock prior to freeing below.
	if (mClipMemNew)
		mClipMemNew = GlobalFree(mClipMemNew);
	// Caller needs us to always return FAIL:
	return *aErrorMessage ? g_script.ScriptError(aErrorMessage) : FAIL;
}



ResultType Clipboard::Close(char *aErrorMessage)
// Returns OK or FAIL (but it only returns FAIL if caller gave us a non-NULL aErrorMessage).
{
	// Always close it ASAP so that it is free for other apps to use:
	if (mIsOpen)
	{
		if (mClipMemNowLocked)
		{
			GlobalUnlock(mClipMemNow); // mClipMemNow not mClipMemNowLocked.
			mClipMemNowLocked = NULL;  // Keep this in sync with its state, since it's used as an indicator.
		}
		// Above: It's probably best to unlock prior to closing the clipboard.
		CloseClipboard();
		mIsOpen = false;  // Even if above fails (realistically impossible?), seems best to do this.
		// Must do this only after GlobalUnlock():
		mClipMemNow = NULL;
	}
	// Do this cleanup for callers that didn't make it far enough to even open the clipboard.
	// UPDATE: DO *NOT* do this because it is valid to have the clipboard in a "ReadyForWrite"
	// state even after we physically close it.  Some callers rely on that.
	//if (mClipMemNewLocked)
	//{
	//	GlobalUnlock(mClipMemNew); // mClipMemNew not mClipMemNewLocked.
	//	mClipMemNewLocked = NULL;
	//	mCapacity = 0; // Keep mCapacity in sync with the state of mClipMemNewLocked.
	//}
	//// Above: Unlock prior to freeing below.
	//if (mClipMemNew)
	//	// Commit() was never called after a call to PrepareForWrite(), so just free the memory:
	//	mClipMemNew = GlobalFree(mClipMemNew);
	if (aErrorMessage && *aErrorMessage)
		// Caller needs us to always return FAIL if an error was displayed:
		return g_script.ScriptError(aErrorMessage);

	// Seems best not to reset mLength.  But it will quickly become out of date once
	// the clipboard has been closed and other apps can use it.
	return OK;
}



ResultType Clipboard::Open()
{
	if (mIsOpen)
		return OK;
	for (DWORD start_time = GetTickCount();;)
	{
		if (OpenClipboard(g_hWnd))
		{
			mIsOpen = true;
			return OK;
		}
		if (g_ClipboardTimeout != -1) // We were told not to wait indefinitely.
			if (!g_ClipboardTimeout   // We were told to make only one attempt, or ...
				|| (int)(g_ClipboardTimeout - (GetTickCount() - start_time)) <= SLEEP_INTERVAL_HALF) //...it timed out.
				// Above must cast to int or any negative result will be lost due to DWORD type.
				return FAIL;
		// Use SLEEP_WITHOUT_INTERRUPTION to prevent MainWindowProc() from accepting new hotkeys
		// during our operation, since a new hotkey subroutine might interfere with
		// what we're doing here (e.g. if it tries to use the clipboard, or perhaps overwrites
		// the deref buffer if this object's caller gave it any pointers into that memory area):
		SLEEP_WITHOUT_INTERRUPTION(INTERVAL_UNSPECIFIED)
	}
}
