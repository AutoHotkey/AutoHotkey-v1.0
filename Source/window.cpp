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
#include "window.h"
#include "util.h" // for strlcpy()
#include "application.h" // for MsgSleep()


HWND WinActivate(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText
	, bool aFindLastMatch, HWND aAlreadyVisited[], int aAlreadyVisitedCount)
{
	if (!aTitle) aTitle = "";
	if (!aText) aText = "";
	if (!aExcludeTitle) aExcludeTitle = "";
	if (!aExcludeText) aExcludeText = "";

	// If window is already active, be sure to leave it that way rather than activating some
	// other window that may match title & text also.  NOTE: An explicit check is done
	// for this rather than just relying on EnumWindows() to obey the z-order because
	// EnumWindows() is *not* guaranteed to enumerate windows in z-order, thus the currently
	// active window, even if it's an exact match, might become overlapped by another matching
	// window.  Also, use the USE_FOREGROUND_WINDOW vs. IF_USE_FOREGROUND_WINDOW macro for
	// this because the active window can sometimes be NULL (i.e. if it's a hidden window
	// and DetectHiddenWindows is off):
	HWND target_window;
	if (USE_FOREGROUND_WINDOW(aTitle, aText, aExcludeTitle, aExcludeText))
	{
		// User asked us to activate the "active" window, which by definition already is.
		// However, if the active (foreground) window is hidden and DetectHiddenWindows is
		// off, the below will set target_window to be NULL, which seems like the most
		// consistent result to use:
		SET_TARGET_TO_ALLOWABLE_FOREGROUND
		return target_window;
	}

	if (!aFindLastMatch && !*aTitle && !*aText && !*aExcludeTitle && !*aExcludeText)
	{
		// User passed no params, so use the window most recently found by WinExist():
		if (   !(target_window = g_ValidLastUsedWindow)   )
			return NULL;
	}
	else
	{
		/*
		// Might not help avg. perfomance any?
		if (!aFindLastMatch) // Else even if the windows is already active, we want the bottomost one.
			if (hwnd = WinActive(aTitle, aText, aExcludeTitle, aExcludeText)) // Already active.
				return target_window;
		*/
		// Don't activate in this case, because the top-most window might be an
		// always-on-top but not-meant-to-be-activated window such as AutoIt's
		// splash text:
		if (   !(target_window = WinExist(aTitle, aText, aExcludeTitle, aExcludeText, aFindLastMatch
			, false, aAlreadyVisited, aAlreadyVisitedCount))   )
			return NULL;
	}
	// If it's invisible, don't bother unless the user explicitly wants to operate
	// on invisible windows(although AutoIt2, but seemingly not 3, handles this
	// by unhiding the window).  Some apps aren't tolerant of having their hidden windows
	// shown by 3rd parties, and they might get messed up.  UPDATE: Since it's valid
	// for a hidden window to be the foreground window, and since the user might want that
	// in obscure cases, don't show the window if it's hidden.  The user can always do that
	// with ShowWindow, if desired:
	if (!IsWindowVisible(target_window) && !g.DetectHiddenWindows)
		return NULL;
	return SetForegroundWindowEx(target_window);
}



inline HWND AttemptSetForeground(HWND aTargetWnd, HWND aForeWnd, char *aTargetTitle = NULL)
// A small inline to help with SetForegroundWindowEx() below.
// Returns NULL if aTargetWnd or its owned-window couldn't be brought to the foreground.
// Otherwise, on success, it returns either aTargetWnd or an HWND owned by aTargetWnd.
#define LOGF "c:\\AutoHotkey SetForegroundWindowEx.txt"
{
#ifdef _DEBUG
	if (!aTargetTitle) aTargetTitle = "";
#endif
	// Probably best not to trust its return value.  It's been shown to be unreliable at times.
	// Example: I've confirmed that SetForegroundWindow() sometimes (perhaps about 10% of the time)
	// indicates failure even though it succeeds.  So we specifically check to see if it worked,
	// which helps to avoid using the keystroke (2-alts) method, because that may disrupt the
	// desired state of the keys or disturb any menus that the user may have displayed.
	// Also: I think the 2-alts last-resort may fire when the system is lagging a bit
	// (i.e. a drive spinning up) and the window hasn't actually become active yet,
	// even though it will soon become active on its own.  Also, SetForegroundWindow() sometimes
	// indicates failure even though it succeeded, usually because the window didn't become
	// active immediately -- perhaps because the system was under load -- but did soon become
	// active on its own (after, say, 50ms or so).  UPDATE: If SetForegroundWindow() is called
	// on a hung window, at least when AttachThreadInput is in effect and that window has
	// a modal dialog (such as MSIE's find dialog), this call might never return, locking up
	// our thread.  So now we do this fast-check for whether the window is hung first (and
	// this call is indeed very fast: its worst case is at least 30x faster than the worst-case
	// performance of the ABORT-IF-HUNG method used with SendMessageTimeout:
	BOOL result = IsWindowHung(aTargetWnd) ? NULL : SetForegroundWindow(aTargetWnd);
	// Note: Increasing the sleep time below did not help with occurrences of "indicated success
	// even though it failed", at least with metapad.exe being activated while command prompt
	// and/or AutoIt2's InputBox were active or present on the screen:
	SLEEP_WITHOUT_INTERRUPTION(SLEEP_INTERVAL); // Specify param so that it will try to specifically sleep that long.
	HWND new_fore_window = GetForegroundWindow();
	if (new_fore_window == aTargetWnd)
	{
#ifdef _DEBUG
		if (!result)
		{
			FileAppend(LOGF, "SetForegroundWindow() indicated failure even though it succeeded: ", false);
			FileAppend(LOGF, aTargetTitle);
		}
#endif
		return aTargetWnd;
	}
	if (new_fore_window != aForeWnd && aTargetWnd == GetWindow(new_fore_window, GW_OWNER))
		// The window we're trying to get to the foreground is the owner of the new foreground window.
		// This is considered to be a success because a window that owns other windows can never be
		// made the foreground window, at least if the windows it owns are visible.
		return new_fore_window;
	// Otherwise, failure:
#ifdef _DEBUG
	if (result)
	{
		FileAppend(LOGF, "SetForegroundWindow() indicated success even though it failed: ", false);
		FileAppend(LOGF, aTargetTitle);
	}
#endif
	return NULL;
}	



HWND SetForegroundWindowEx(HWND aWnd)
// Caller must have ensured that aWnd is a valid window or NULL, since we
// don't call IsWindow() here.
{
	if (!aWnd) return NULL;  // When called this way (as it is sometimes), do nothing.

#ifdef _DEBUG
	char win_name[64];
	GetWindowText(aWnd, win_name, sizeof(win_name));
#endif

	HWND orig_foreground_wnd = GetForegroundWindow();
	// AutoIt3: If there is not any foreground window, then input focus is on the TaskBar.
	// MY: It is definitely possible for GetForegroundWindow() to return NULL, even on XP.
	if (!orig_foreground_wnd)
		orig_foreground_wnd = FindWindow("Shell_TrayWnd", NULL);

	// AutoIt3: If the target window is currently top - don't bother:
	if (aWnd == orig_foreground_wnd)
		return aWnd;

	if (IsIconic(aWnd))
		// This might never return if aWnd is a hung window.  But it seems better
		// to do it this way than to use the PostMessage() method, which might not work
		// reliably with apps that don't handle such messages in a standard way.
		// A minimized window must be restored or else SetForegroundWindow() always(?)
		// won't work on it.  UPDATE: ShowWindowAsync() would prevent a hang, but
		// probably shouldn't use it because we rely on the fact that the message
		// has been acted on prior to trying to activate the window (and all Async()
		// does is post a message to its queue):
		ShowWindow(aWnd, SW_RESTORE);

	// This causes more trouble than it's worth.  In fact, the AutoIt author said that
	// he didn't think it even helped with the IE 5.5 related issue it was originally
	// intended for, so it seems a good idea to NOT to this, especially since I'm 80%
	// sure it messes up the Z-order in certain circumstances, causing an unexpected
	// window to pop to the foreground immediately after a modal dialog is dismissed:
	//BringWindowToTop(aWnd); // AutoIt3: IE 5.5 related hack.

	HWND new_foreground_wnd;

	if (!g_WinActivateForce)
	// if (g_os.IsWin95() || (!g_os.IsWin9x() && !g_os.IsWin2000orLater())))  // Win95 or NT
		// Try a simple approach first for these two OS's, since they don't have
		// any restrictions on focus stealing:
#ifdef _DEBUG
#define IF_ATTEMPT_SET_FORE if (new_foreground_wnd = AttemptSetForeground(aWnd, orig_foreground_wnd, win_name))
#else
#define IF_ATTEMPT_SET_FORE if (new_foreground_wnd = AttemptSetForeground(aWnd, orig_foreground_wnd, ""))
#endif
		IF_ATTEMPT_SET_FORE
			return new_foreground_wnd;
		// Otherwise continue with the more drastic methods below.

	// MY: The AttachThreadInput method, when used by itself, seems to always
	// work the first time on my XP system, seemingly regardless of whether the
	// "allow focus steal" change has been made via SystemParametersInfo()
	// (but it seems a good idea to keep the SystemParametersInfo() in effect
	// in case Win2k or Win98 needs it, or in case it really does help in rare cases).
	// In many cases, this avoids the two SetForegroundWindow() attempts that
	// would otherwise be needed; and those two attempts cause some windows
	// to flash in the taskbar, such as Metapad and Excel (less frequently) whenever
	// you quickly activate another window after activating it first (e.g. via hotkeys).
	// So for now, it seems best just to use this method by itself.  The
	// "two-alts" case never seems to fire on my system?  Maybe it will
	// on Win98 sometimes.
	// Note: In addition to the "taskbar button flashing" annoyance mentioned above
	// any SetForegroundWindow() attempt made prior to the one below will,
	// as a side-effect, sometimes trigger the need for the "two-alts" case
	// below.  So that's another reason to just keep it simple and do it this way
	// only.

#ifdef _DEBUG
	char buf[1024];
#endif

	bool is_attached_my_to_fore = false, is_attached_fore_to_target = false;
	DWORD fore_thread, my_thread, target_thread;
	if (orig_foreground_wnd) // Might be NULL from above.
	{
		// Based on MSDN docs, these calls should always succeed due to the other
		// checks done above (e.g. that none of the HWND's are NULL):
		// AutoIt3: Get the details of all the input threads involved (myappswin,
		// foreground win, target win):
		fore_thread = GetWindowThreadProcessId(orig_foreground_wnd, NULL);
		my_thread  = GetCurrentThreadId();  // It's probably best not to have this var be static.
		target_thread = GetWindowThreadProcessId(aWnd, NULL);

		// MY: Normally, it's suggested that you only need to attach the thread of the
		// foreground window to our thread.  However, I've confirmed that doing all three
		// attaches below makes the attempt much more likely to succeed.  In fact, it
		// almost always succeeds whereas the one-attach method hardly ever succeeds the first
		// time (resulting in a flashing taskbar button due to having to invoke a second attempt)
		// when one window is quickly activated after another was just activated.
		// AutoIt3: Attach all our input threads, will cause SetForeground to work under 98/ME.
		// MSDN docs: The AttachThreadInput function fails if either of the specified threads
		// does not have a message queue (My: ok here, since any window's thread MUST have a
		// message queue).  [It] also fails if a journal record hook is installed.  ... Note
		// that key state, which can be ascertained by calls to the GetKeyState or
		// GetKeyboardState function, is reset after a call to AttachThreadInput.  You cannot
		// attach a thread to a thread in another desktop.  A thread cannot attach to itself.
		// Therefore, idAttachTo cannot equal idAttach.  Update: It appears that of the three,
		// this first call does not offer any additional benefit, at least on XP, so not
		// using it for now:
		//if (my_thread != target_thread) // Don't attempt the call otherwise.
		//	AttachThreadInput(my_thread, target_thread, TRUE);
		if (fore_thread && my_thread != fore_thread && !IsWindowHung(orig_foreground_wnd))
			is_attached_my_to_fore = AttachThreadInput(my_thread, fore_thread, TRUE) != 0;
		if (fore_thread && target_thread && fore_thread != target_thread && !IsWindowHung(aWnd))
			is_attached_fore_to_target = AttachThreadInput(fore_thread, target_thread, TRUE) != 0;
	}

	// The log showed that it never seemed to need more than two tries.  But there's
	// not much harm in trying a few extra times.  The number of tries needed might
	// vary depending on how fast the CPU is:
	for (int i = 0; i < 5; ++i)
	{
		IF_ATTEMPT_SET_FORE
		{
#ifdef _DEBUG
			if (i > 0) // More than one attempt was needed.
			{
				snprintf(buf, sizeof(buf), "AttachThreadInput attempt #%d indicated success: %s"
					, i + 1, win_name);
				FileAppend(LOGF, buf);
			}
#endif
			break;
		}
	}

	// I decided to avoid the quick minimize + restore method of activation.  It's
	// not that much more effective (if at all), and there are some significant
	// disadvantages:
	// - This call will often hang our thread if aWnd is a hung window: ShowWindow(aWnd, SW_MINIMIZE)
	// - Using SW_FORCEMINIMIZE instead of SW_MINIMIZE has at least one (and probably more)
	// side effect: When the window is restored, at least via SW_RESTORE, it is no longer
	// maximized even if it was before the minmize.  So don't use it.
	if (!new_foreground_wnd) // Not successful yet.
	{
		// Some apps may be intentionally blocking us by having called the API function
		// LockSetForegroundWindow(), for which MSDN says "The system automatically enables
		// calls to SetForegroundWindow if the user presses the ALT key or takes some action
		// that causes the system itself to change the foreground window (for example,
		// clicking a background window)."  Also, it's probably best to avoid doing
		// the 2-alts method except as a last resort, because I think it may mess up
		// the state of menus the user had displayed.  And of course if the foreground
		// app has special handling for alt-key events, it might get confused.
		// My original note: "The 2-alts case seems to mess up on rare occasions,
		// perhaps due to menu weirdness triggered by the alt key."
		// AutoIt3: OK, this is not funny - bring out the extreme measures (usually for 2000/XP).
		// Simulate two single ALT keystrokes.  UPDATE: This hardly ever succeeds.  Usually when
		// it fails, the foreground window is NULL (none).  I'm going to try an Win-tab instead,
		// which selects a task bar button.  This seems less invasive than doing an alt-tab
		// because not only doesn't it activate some other window first, it also doesn't appear
		// to change the Z-order, which is good because we don't want the alt-tab order
		// that the user sees to be affected by this.  UPDATE: Win-tab isn't doing it, so try
		// Alt-tab.  Alt-tab doesn't do it either.  The window itself (metapad.exe is the only
		// culprit window I've found so far) seems to resist being brought to the foreground,
		// but later, after the hotkey is released, it can be.  So perhaps this is being
		// caused by the fact that the user has keys held down (logically or physically?)
		// Releasing those keys with a key-up event might help, so try that sometime:
		KeyEvent(KEYDOWNANDUP, VK_MENU);
		KeyEvent(KEYDOWNANDUP, VK_MENU);
		//KeyEvent(KEYDOWN, VK_LWIN);
		//KeyEvent(KEYDOWN, VK_TAB);
		//KeyEvent(KEYUP, VK_TAB);
		//KeyEvent(KEYUP, VK_LWIN);
		//KeyEvent(KEYDOWN, VK_MENU);
		//KeyEvent(KEYDOWN, VK_TAB);
		//KeyEvent(KEYUP, VK_TAB);
		//KeyEvent(KEYUP, VK_MENU);
		// Also replacing "2-alts" with "alt-tab" below, for now:

		IF_ATTEMPT_SET_FORE
#ifndef _DEBUG
			0; // Do nothing.
#else
			FileAppend(LOGF, "2-alts ok: ", false);
		else
		{
			FileAppend(LOGF, "2-alts (which is the last resort) failed.  ", false);
			HWND h = GetForegroundWindow();
			if (h)
			{
				char fore_name[64];
				GetWindowText(h, fore_name, sizeof(fore_name));
				FileAppend(LOGF, "Foreground: ", false);
				FileAppend(LOGF, fore_name, false);
			}
			FileAppend(LOGF, ".  Was trying to activate: ", false);
		}
		FileAppend(LOGF, win_name);
#endif
	} // if()

	// Very important to detach any threads whose inputs were attached above,
	// prior to returning, otherwise the next attempt to attach thread inputs
	// for these particular windows may result in a hung thread or other
	// undesirable effect:
	if (is_attached_my_to_fore)
		AttachThreadInput(my_thread, fore_thread, FALSE);
	if (is_attached_fore_to_target)
		AttachThreadInput(fore_thread, target_thread, FALSE);

	// Finally.  This one works, solving the problem of the MessageBox window
	// having the input focus and being the foreground window, but not actually
	// being visible (even though IsVisible() and IsIconic() say it is)!  It may
	// help with other conditions under which this function would otherwise fail.
	// Here's the way the repeat the failure to test how the absence of this line
	// affects things, at least on my XP SP1 system:
	// y::MsgBox, test
	// #e::(some hotkey that activates Windows Explorer)
	// Now: Activate explorer with the hotkey, then invoke the MsgBox.  It will
	// usually be activated but invisible.  Also: Whenever this invisible problem
	// is about to occur, with or without this fix, it appears that the OS's z-order
	// is a bit messed up, because when you dismiss the MessageBox, an unexpected
	// window (probably the one two levels down) becomes active rather than the
	// window that's only 1 level down in the z-order:
	if (new_foreground_wnd) // success.
	{
		// Even though this is already done for the IE 5.5 "hack" above, must at
		// a minimum do it here: The above one may be optional, not sure (safest
		// to leave it unless someone can test with IE 5.5).
		// Note: I suspect the two lines below achieve the same thing.  They may
		// even be functionally identical.  UPDATE: This may no longer be needed
		// now that the first BringWindowToTop(), above, has been disabled due to
		// its causing more trouble than it's worth.  But seems safer to leave
		// this one enabled in case it does resolve IE 5.5 related issues and
		// possible other issues:
		BringWindowToTop(aWnd);
		//SetWindowPos(aWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
		return new_foreground_wnd; // Return this rather than aWnd because it's more appropriate.
	}
	else
		return NULL;
}



HWND WinClose(char *aTitle, char *aText, int aTimeToWaitForClose
	, char *aExcludeTitle, char *aExcludeText, bool aKillIfHung)
// Return the HWND of any found-window to the caller so that it has the option of waiting
// for it to become an invalid (closed) window.
{
	// Methods like this avoid any chance of dereferencing a NULL pointer:
	if (!aTitle) aTitle = "";
	if (!aText) aText = "";
	if (!aExcludeTitle) aExcludeTitle = "";
	if (!aExcludeText) aExcludeText = "";
	if (aTimeToWaitForClose < 0) aTimeToWaitForClose = 0;

	HWND target_window;
	IF_USE_FOREGROUND_WINDOW(aTitle, aText, aExcludeTitle, aExcludeText)
		// Close topmost (better than !F4 since that uses the alt key, effectively resetting
		// its status to UP if it was down before.  Use WM_CLOSE rather than WM_EXIT because
		// I think that's what Alt-F4 sends (and otherwise, app may quit without offering
		// a chance to save).
		// DON'T DISPLAY a MsgBox (e.g. debugging) before trying to close foreground window.
		// Otherwise, it may close the owner of the dialog window (this app), perhaps due to
		// split-second timing issues.
	else if (*aTitle || *aText || *aExcludeTitle || *aExcludeText)
	{
		// Since EnumWindows() is *not* guaranteed to start proceed in z-order from topmost to
		// bottomost (though it almost certainly does), do it this way to ensure that the
		// topmost window is closed in preference to any other windows with the same <aTitle>
		// and <aText>:
		if (   !(target_window = WinActive(aTitle, aText, aExcludeTitle, aExcludeText))   )
			if (   !(target_window = WinExist(aTitle, aText, aExcludeTitle, aExcludeText))   )
				return NULL;
	}
	else
		target_window = g_ValidLastUsedWindow;
	if (!target_window)
		return NULL;

	if (aKillIfHung) // This part is based on the AutoIt3 source.
	{
		// AutoIt3 waits for 500ms.  But because this app is much more sensitive to being
		// in a "not-pumping-messages" state, due to the keyboard & mouse hooks, it seems
		// better to wait for less (e.g. in case the user is gaming and there's a script
		// running in the background that uses WinKill, we don't want key and mouse events
		// to freeze for a long time).  Also, always use WM_CLOSE vs. SC_CLOSE in this case
		// since the target window is slightly more likely to respond to that:
		DWORD dwResult;
		if (!SendMessageTimeout(target_window, WM_CLOSE, 0, 0, SMTO_ABORTIFHUNG, 200, &dwResult))
		{
			// Use more force - Mwuahaha
			DWORD pid = GetWindowThreadProcessId(target_window, NULL);
			HANDLE hProcess = pid ? OpenProcess(PROCESS_ALL_ACCESS, FALSE, pid) : NULL;
			if (hProcess)
			{
				TerminateProcess(hProcess, 0);
				CloseHandle(hProcess);
			}
		}
	}
	else // Don't kill.
		// SC_CLOSE is the same as clicking a window's "X"(close) button or using Alt-F4.
		// Although it's a more friendly way to close windows than WM_CLOSE (and thus
		// avoids incompatibilities with apps such as MS Visual C++), apps that
		// have disabled Alt-F4 processing will not be successfully closed.  It seems
		// best not to send both SC_CLOSE and WM_CLOSE because some apps with an 
		// "Unsaved.  Are you sure?" type dialog might close down completely rather than
		// waiting for the user to confirm.  Anyway, it's extrememly rare for a window
		// not to respond to Alt-F4 (though it is possible that it handles Alt-F4 in a
		// non-standard way, i.e. that sending SC_CLOSE equivalent to Alt-F4
		// for windows that handle Alt-F4 manually?)  But on the upside, this is nicer
		// for apps that upon receiving Alt-F4 do some behavior other than closing, such
		// as minimizing to the tray.  Such apps might shut down entirely if they received
		// a true WM_CLOSE, which is probably not what the user would want.
		// Update: Swithced back to using WM_CLOSE so that instances of AutoHotkey
		// can be terminated via another instances use of the WinClose command:
		//PostMessage(target_window, WM_SYSCOMMAND, SC_CLOSE, 0);
		PostMessage(target_window, WM_CLOSE, 0, 0);

	// Slight delay.  Might help avoid user having to modify script to use WinWaitClose()
	// in many cases.  UPDATE: But this does a Sleep(0), which won't yield the remainder
	// of our thread's timeslice unless there's another app trying to use 100% of the CPU?
	// So, in reality it really doesn't accomplish anything because the window we just
	// closed won't get any CPU time (unless, perhaps, it receives the close message in
	// time to ask the OS for us to yield the timeslice).  Perhaps some finer tuning
	// of this can be done in the future.  UPDATE: Testing of WinActivate, which also
	// uses this to do a Sleep(0), reveals that it may in fact help even when the CPU
	// isn't under load.  Perhaps this is because upon Sleep(0), the OS runs the
	// WindowProc's of windows that have messages waiting for them so that appropriate
	// action can be taken (which may often be nearly instantaneous, perhaps under
	// 1ms for a Window to be logically destroyed even if it hasn't physically been
	// removed from the screen?) prior to returning the CPU to our thread:
	DWORD start_time = GetTickCount(); // Before doing any MsgSleeps, set this.
    //MsgSleep(0); // Always do one small one, see above comments.
	// UPDATE: It seems better just to always do one unspecified-interval sleep
	// rather than MsgSleep(0), which often returns immediately, probably having
	// no effect.

	// Remember that once the first call to MsgSleep() is done, a new hotkey subroutine
	// may fire and suspend what we're doing here.  Such a subroutine might also overwrite
	// the values our params, some of which may be in the deref buffer.  So be sure not
	// to refer to those strings once MsgSleep() has been done, below:

	// This is the same basic code used for ACT_WINWAITCLOSE and such:
	for (;;)
	{
		// Seems best to always do the first one regardless of the value 
		// of aTimeToWaitForClose:
		MsgSleep(INTERVAL_UNSPECIFIED);
		if (!IsWindow(target_window)) // It's gone, so we're done.
			return target_window;
		// Must cast to int or any negative result will be lost due to DWORD type:
		if ((int)(aTimeToWaitForClose - (GetTickCount() - start_time)) <= SLEEP_INTERVAL_HALF)
			break;
			// Last param 0 because we don't want it to restore the
			// current active window after the time expires (in case
			// it's suspended).  INTERVAL_UNSPECIFIED performs better.
	}
	return target_window;  // Done waiting.
}



HWND WinActive(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText, bool aUpdateLastUsed)
{
	if (!aTitle) aTitle = "";
	if (!aText) aText = "";
	if (!aExcludeTitle) aExcludeTitle = "";
	if (!aExcludeText) aExcludeText = "";

	HWND target_window;
	if (USE_FOREGROUND_WINDOW(aTitle, aText, aExcludeTitle, aExcludeText))
	{
		// User asked us if the "active" window is active, which is true if it's not a
		// hidden window or DetectHiddenWindows is ON:
		SET_TARGET_TO_ALLOWABLE_FOREGROUND
		#define UPDATE_AND_RETURN_LAST_USED_WINDOW(hwnd) \
		{\
			if (aUpdateLastUsed && hwnd)\
				g.hWndLastUsed = hwnd;\
			return hwnd;\
		}
		UPDATE_AND_RETURN_LAST_USED_WINDOW(target_window)
	}

	HWND fore_win = GetForegroundWindow();
	if (!fore_win)
		return NULL;
	if (!g.DetectHiddenWindows && !IsWindowVisible(fore_win)) // In this case, the caller's window can't be active.
		return NULL;

	if (!*aTitle && !*aText && !*aExcludeTitle && !*aExcludeText)
		// User passed no params, so use the window most recently found by WinExist().
		return (fore_win == g_ValidLastUsedWindow) ? fore_win : NULL;

	char active_win_title[WINDOW_TEXT_SIZE];
	// Don't use GetWindowTextByTitleMatchMode() because Aut3 uses the same fast
	// method as below for window titles:
	if (!GetWindowText(fore_win, active_win_title, sizeof(active_win_title)))
		return NULL;

	if (!IsTextMatch(active_win_title, aTitle, aExcludeTitle))
		// Active window's title doesn't match.
		return NULL;

	// Otherwise confirm the match by ensuring that active window has a child that contains <aText>.
	// (it will return "success" immediately if aText & aExcludeText are both blank):
	if (HasMatchingChild(fore_win, aText, aExcludeText))
		UPDATE_AND_RETURN_LAST_USED_WINDOW(fore_win)
	else
		return NULL;
}



HWND WinExist(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText
	, bool aFindLastMatch, bool aUpdateLastUsed
	, HWND aAlreadyVisited[], int aAlreadyVisitedCount)
{
	// Seems okay to allow both title and text to be NULL or empty.  It would then find the first window
	// of any kind (and there's probably always at least one, even on a blank desktop).
	if (!aTitle) aTitle = "";
	if (!aText) aText = "";
	if (!aExcludeTitle) aExcludeTitle = "";
	if (!aExcludeText) aExcludeText = "";

	HWND target_window;
	if (USE_FOREGROUND_WINDOW(aTitle, aText, aExcludeTitle, aExcludeText))
	{
		// User asked us if the "active" window exists, which is true if it's not a
		// hidden window or DetectHiddenWindows is ON:
		SET_TARGET_TO_ALLOWABLE_FOREGROUND
		// Updating LastUsed to be hwnd even if it's NULL seems best for consistency?
		// UPDATE: No, it's more flexible not to never set it to NULL, because there
		// will be times when the old value is still useful:
		UPDATE_AND_RETURN_LAST_USED_WINDOW(target_window);
	}

	if (!*aTitle && !*aText && !*aExcludeTitle && !*aExcludeText)
		// User passed no params, so use the window most recently found by WinExist().
		// It's correct to do this even in this function because it's called by
		// WINWAITCLOSE and IFWINEXIST specifically to discover if the Last-Used
		// window still exists.
		return g_ValidLastUsedWindow;

	WindowInfoPackage wip;
	wip.find_last_match = aFindLastMatch;
	strlcpy(wip.title, aTitle, sizeof(wip.title));
	strlcpy(wip.text, aText, sizeof(wip.text));
	strlcpy(wip.exclude_title, aExcludeTitle, sizeof(wip.exclude_title));
	strlcpy(wip.exclude_text, aExcludeText, sizeof(wip.exclude_text));
	wip.already_visited = aAlreadyVisited;
	wip.already_visited_count = aAlreadyVisitedCount;

	// Note: It's a little strange, but I think EnumWindows() returns FALSE when the callback stopped
	// the enumeration prematurely by returning false to its caller.  Otherwise (the enumeration went
	// through every window), it returns TRUE:
	EnumWindows(EnumParentFind, (LPARAM)&wip);
	UPDATE_AND_RETURN_LAST_USED_WINDOW(wip.parent_hwnd);
}



BOOL CALLBACK EnumParentFind(HWND aWnd, LPARAM lParam)
// To continue enumeration, the function must return TRUE; to stop enumeration, it must return FALSE. 
{
	// According to MSDN, GetWindowText() will hang only if it's done against
	// one of your own app's windows and that window is hung.  I suspect
	// this might not be true in Win95, and possibly not even Win98, but
	// it's not really an issue because GetWindowText() has to be called
	// eventually, either here or in an EnumWindowsProc.  The only way
	// to prevent hangs (if indeed it does hang on Win9x) would be to
	// call something like IsWindowHung() before every call to
	// GetWindowText(), which might result in a noticeable delay whenever
	// we search for a window via its title (or even worse: by the title
	// of one of its controls or child windows).  UPDATE: Trying GetWindowTextTimeout()
	// now, which might be the best compromise.  UPDATE: It's annoyingly slow,
	// so went back to using the old method.
	if (!g.DetectHiddenWindows && !IsWindowVisible(aWnd)) // Skip hidden windows in this case.
		return TRUE;
	char win_title[WINDOW_TEXT_SIZE];
	// Don't use GetWindowTextByTitleMatchMode() because it's (always?) unnecessary for
	// window titles (AutoIt3 does this same thing):
	if (!GetWindowText(aWnd, win_title, sizeof(win_title)))
		return TRUE;  // Even if can't get the text of some window, for some reason, keep enumerating.
	// strstr() and related std C functions -- as well as the custom stristr(), will always
	// find the empty string in any string, which is what we want in case title is the empty string.
	if (!IsTextMatch(win_title, pWin->title, pWin->exclude_title))
		// Since title doesn't match there's no point in checking the text of this HWND.
		// Just continue finding more top-level (parent) windows:
		return TRUE;

	// Disqualify this window if the caller provided us a list of unwanted windows:
	if (pWin->already_visited_count && pWin->already_visited)
		for (int i = 0; i < pWin->already_visited_count; ++i)
			if (aWnd == pWin->already_visited[i])
				return TRUE; // Not a match, so skip this one and keep searching.

	// Otherwise, the title matches.  If text is specified, the child windows of this parent
	// must be searched to try to find a match for it:
	if (*pWin->text || *pWin->exclude_text)
	{
		// Search for the specfied text among the children of this window.
		// EnumChildWindows() will return FALSE (failure) in at least two common conditions:
		// 1) It's EnumChildProc callback returned false (i.e. it ended the enumeration prematurely).
		// 2) The specified parent has no children.
		// Since in both these cases GetLastError() returns ERROR_SUCCESS, we discard the return
		// value and just check the struct's child_hwnd to determine whether a match has been found:
		pWin->child_hwnd = NULL;  // Init prior to each call, in case find_last_match is true.
		EnumChildWindows(aWnd, EnumChildFind, lParam);
		if (pWin->child_hwnd == NULL)
			// This parent has no matching child, or no children at all, so search for more parents:
			return TRUE;
	}

	// Otherwise, a complete match has been found.  Set this output value for the caller.
	// If find_last_match is true, this value will stay in effect unless overridden
	// by another matching window:
	pWin->parent_hwnd = aWnd;

	// If find_last_match is true, continue searching.  Otherwise, this first match is the one
	// that's desired so stop here:
	return pWin->find_last_match;  // Returning a bool in lieu of BOOL is safe in this case.
}



BOOL CALLBACK EnumChildFind(HWND aWnd, LPARAM lParam)
// Although this function could be rolled into a generalized version of the EnumWindowsProc(),
// it will perform better this way because there's less checking required and no mode/flag indicator
// is needed inside lParam to indicate which struct element should be searched for.  In addition,
// it's more comprehensible this way.  lParam is a pointer to the struct rather than just a
// string because we want to give back the HWND of any matching window.
{
	char win_text[WINDOW_TEXT_SIZE];
	if (!g.DetectHiddenText && !IsWindowVisible(aWnd))
		return TRUE;  // This child/control is hidden and user doesn't want it considered, so skip it.
	if (!GetWindowTextByTitleMatchMode(aWnd, win_text, sizeof(win_text)))
		return TRUE;  // Even if can't get the text of some window, for some reason, keep enumerating.
	// Below: Tell it to find match anywhere in the child-window text, rather than just
	// in the leading part, because this is how AutoIt2 and AutoIt3 operate:
	if (IsTextMatch(win_text, pWin->text, pWin->exclude_text, true))
	{
		// Match found, so stop searching.
		//char class_name[64];
		//GetClassName(aWnd, class_name, sizeof(class_name));
		//MsgBox(class_name);
		pWin->child_hwnd = aWnd;
		return FALSE;
	}
	// Since this child doesn't match, make sure none of its children (recursive)
	// match prior to continuing the original enumeration.  We don't discard the
	// return value from EnumChildWindows() because it's FALSE in two cases:
	// 1) The given HWND has no children.
	// 2) The given EnumChildProc() stopped prematurely rather than enumerating all the windows.
	// and there's no way to distinguish between the two cases without using the
	// struct's hwnd because GetLastError() seems to return ERROR_SUCCESS in both
	// cases.  UPDATE: The MSDN docs state that EnumChildWindows already handles the
	// recursion for us: "If a child window has created child windows of its own,
	// EnumChildWindows() enumerates those windows as well."
	//EnumChildWindows(aWnd, EnumChildFind, lParam);
	// If matching HWND still hasn't been found, return TRUE to keep searching:
	//return pWin->child_hwnd == NULL;
	return TRUE; // Keep searching.
}



ResultType StatusBarUtil(Var *aOutputVar, HWND aControlWindow, int aPartNumber
	, char *aTextToWaitFor, int aWaitTime, int aCheckInterval)
// Most of this courtesy of AutoIt3 source.
// aOutputVar is allowed to be NULL if aTextToWaitFor isn't NULL or blank.
// aControlWindow is allowed to be NULL because we want to set the output var
// to be the empty string in that case.
{
	// Set default ErrorLevel, which is a special value (2 vs. 1) in the case of StatusBarWait:
	g_ErrorLevel->Assign(aOutputVar ? ERRORLEVEL_ERROR : ERRORLEVEL_ERROR2);
	if (aCheckInterval <= 0) aCheckInterval = SB_DEFAULT_CHECK_INTERVAL; // Caller relies on us doing this.
	if (!aTextToWaitFor) aTextToWaitFor = "";  // For consistency.

	// Must have at least one of these.  UPDATE: We want to allow this so
	// that the command can be used to wait for the status bar text to
	// become blank:
	//if (!aOutputVar && !*aTextToWaitFor) return FAIL;

	// Whenever using SendMessageTimeout(), our app will be unresponsive until
	// the call returns, since our message loop isn't running.  In addition,
	// if the keyboard or mouse hook is installed, the input events will lag during
	// this call.  So keep the timeout value fairly short:
	#define SB_TIMEOUT 100

	// AutoIt3: How many parts does this bar have?
	DWORD nParts = 0;
	if (aControlWindow)
		if (!SendMessageTimeout(aControlWindow, SB_GETPARTS, (WPARAM)0, (LPARAM)0
			, SMTO_ABORTIFHUNG, SB_TIMEOUT, &nParts)) // It failed or timed out.
			nParts = 0; // It case it set it to some other value before failing.

	if (aPartNumber < 1)
		aPartNumber = 1;  // Caller relies on us to set default in this case.
	if (aPartNumber > (int)nParts)
		aPartNumber = 0; // Set this as an indicator for below.

	VarSizeType space_needed;
	char buf[WINDOW_TEXT_SIZE + 1] = ""; // +1 is needed in this case.
	if (!aControlWindow || !aPartNumber)
		space_needed = 1; // 1 for terminator.
	else
	{
		DWORD dwResult;
		if (!SendMessageTimeout(aControlWindow, SB_GETTEXTLENGTH, (WPARAM)(aPartNumber - 1), (LPARAM)0
			, SMTO_ABORTIFHUNG, SB_TIMEOUT, &dwResult))
			// It timed out or failed.  Since we can't even find the length, don't bother
			// with anything else:
			return FAIL;
		if (LOWORD(dwResult) > WINDOW_TEXT_SIZE) // extremely unlikely, perhaps impossible.
			return FAIL;
		if (!aWaitTime)
			// Waiting 500ms in place of a "0" seems more useful than a true zero, which
			// doens't need to be supported because it's the same thing as something like
			// "IfWinExist".
			aWaitTime = 500;
		DWORD start_time;
		#define WAIT_INDEFINITELY (aWaitTime < 0)
		if (!WAIT_INDEFINITELY)
			start_time = GetTickCount();

		LPVOID pMem;
		if (g_os.IsWinNT())
		{
			DWORD dwPid;
			GetWindowThreadProcessId(aControlWindow, &dwPid);
			HANDLE hProcess = OpenProcess(PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_VM_WRITE, FALSE, dwPid);
			if (hProcess)
			{
				// AutoIt3: Dynamic functions to retain 95 compatibility
				// My: The above comment seems wrong since this section is only for NT/2k/XP+.
				// Perhaps it meant that only NT and/or 2k require dynamic functions whereas XP doesn't:
				typedef LPVOID (WINAPI *MyVirtualAllocEx)(HANDLE, LPVOID, SIZE_T, DWORD, DWORD);
				// Static for performance, since value should be always the same.
				static MyVirtualAllocEx lpfnAlloc = (MyVirtualAllocEx)GetProcAddress(GetModuleHandle("kernel32.dll")
					, "VirtualAllocEx");
				pMem = lpfnAlloc(hProcess, NULL, sizeof(buf), MEM_RESERVE | MEM_COMMIT, PAGE_READWRITE);

				for (;;)
				{ // Always do the first iteration so that at least one check is done.
					if (!SendMessageTimeout(aControlWindow, SB_GETTEXT, (WPARAM)(aPartNumber - 1), (LPARAM)pMem
						, SMTO_ABORTIFHUNG, SB_TIMEOUT, &dwResult))
						// It failed or timed out; buf stays as it was: initialized to empty string.
						// Also ErrorLevel stays set to 2, the default set above.
						break;
					if (!ReadProcessMemory(hProcess, pMem, buf, WINDOW_TEXT_SIZE, NULL))
					{
						*buf = '\0';  // In case it changed the buf before failing.
						break;
					}
					buf[sizeof(buf) - 1] = '\0';  // Just to be sure.

					// Below: In addition to normal/intuitive matching, a match is also achieved if
					// both are empty string:
					#define BREAK_IF_MATCH_FOUND_OR_IF_NOT_WAITING \
					if ((!*aTextToWaitFor && !*buf) || (aTextToWaitFor && IsTextMatch(buf, aTextToWaitFor))) \
					{\
						g_ErrorLevel->Assign(ERRORLEVEL_NONE);\
						break;\
					}\
					if (aOutputVar)\
						break;  // i.e. If an output variable was given, we're not waiting for a match.
					BREAK_IF_MATCH_FOUND_OR_IF_NOT_WAITING

					// Don't continue to the wait if the target window is destroyed:
					// Must cast to int or any negative result will be lost due to DWORD type.
					// Also: Last param false because we don't want it to restore the
					// current active window after the time expires (in case
					// our subroutine is suspended).  Also, ERRORLEVEL_ERROR is the value that
					// indicates that we timed out rather than having ever found a match:
					#define SB_SLEEP_IF_NEEDED \
					if (!IsWindow(aControlWindow))\
						break;\
					if (WAIT_INDEFINITELY || (int)(aWaitTime - (GetTickCount() - start_time)) > SLEEP_INTERVAL_HALF)\
						MsgSleep(aCheckInterval);\
					else\
					{\
						g_ErrorLevel->Assign(ERRORLEVEL_ERROR);\
						break;\
					}
					SB_SLEEP_IF_NEEDED
				}

				// AutoIt3: Dynamic functions to retain 95 compatibility
				typedef BOOL (WINAPI *MyVirtualFreeEx)(HANDLE, LPVOID, SIZE_T, DWORD);
				// Static for performance, since value should be always the same.
				static MyVirtualFreeEx lpfnFree = (MyVirtualFreeEx)GetProcAddress(GetModuleHandle("kernel32.dll")
					, "VirtualFreeEx");
				lpfnFree(hProcess, pMem, 0, MEM_RELEASE); // Size 0 is used with MEM_RELEASE.
				CloseHandle(hProcess);
			} // if (hProcess)
		} // WinNT
		else // Win9x
		{
			HANDLE hMapping = CreateFileMapping(INVALID_HANDLE_VALUE, 0, PAGE_READWRITE, 0, WINDOW_TEXT_SIZE, NULL);
			if (hMapping)
			{
				pMem = MapViewOfFile(hMapping, FILE_MAP_ALL_ACCESS, 0, 0, 0);
				for (;;)
				{ // Always do the first iteration so that at least one check is done.
					if (SendMessageTimeout(aControlWindow, SB_GETTEXT, (WPARAM)(aPartNumber - 1), (LPARAM)pMem
						, SMTO_ABORTIFHUNG, SB_TIMEOUT, &dwResult))
					{
						// Not sure why AutoIt3 doesn't use use strcpy() for this, but leave it to be safe:
						CopyMemory(buf, pMem, WINDOW_TEXT_SIZE);
						buf[sizeof(buf) - 1] = '\0';  // Just to be sure.
						BREAK_IF_MATCH_FOUND_OR_IF_NOT_WAITING
					}
					else // it failed or timed out; buf stays as it was: initialized to empty string.
						break;

					SB_SLEEP_IF_NEEDED
				}
				UnmapViewOfFile(pMem);
				CloseHandle(hMapping);
			}
		}
		space_needed = (VarSizeType)(strlen(buf) + 1); // +1 for terminator.
	}

	// Otherwise, consider this to be always successful, even if aControlWindow == NULL
	// or the status bar didn't have the part number provided, unless the below fails.
	if (aOutputVar)
		// Note we use a temp buf rather than writing directly to the var contents above, because
		// we don't know how long the text will be until after the above operation finishes.
		return aOutputVar->Assign(buf, space_needed - 1);
	// else caller didn't want the text.
	return OK;
}



HWND ControlExist(HWND aParentWindow, char *aClassNameAndNum)
{
	if (!aParentWindow) return NULL;
	if (!aClassNameAndNum || !*aClassNameAndNum) return GetTopChild(aParentWindow);
	WindowInfoPackage wip;
	bool is_class_name = isdigit(aClassNameAndNum[strlen(aClassNameAndNum) - 1]);
	if (is_class_name)
        strlcpy(wip.title, aClassNameAndNum, sizeof(wip.title));  // Tell it to search by Class+Num.
	else
		strlcpy(wip.text, aClassNameAndNum, sizeof(wip.text)); // Tell it to search the control's text.
	// It's a little strange, but I think EnumWindows() returns FALSE when the callback stopped
	// the enumeration prematurely by returning false to its caller.  Otherwise (the enumeration went
	// through every window), it returns TRUE:
	EnumChildWindows(aParentWindow, EnumControlFind, (LPARAM)&wip);
	if (is_class_name && !wip.child_hwnd)
	{
		// To reduce problems with ambiguity (a class name and number of one control happens
		// to match the title/text of another control), search again only after the search
		// for the ClassNameAndNum didn't turn up anything.
		*wip.title = '\0';
		strlcpy(wip.text, aClassNameAndNum, sizeof(wip.text)); // Tell it to search the control's text.
		EnumChildWindows(aParentWindow, EnumControlFind, (LPARAM)&wip);
	}
	return wip.child_hwnd;
}



BOOL CALLBACK EnumControlFind(HWND aWnd, LPARAM lParam)
// lParam is a pointer to the struct rather than just a string because we want
// to give back the HWND of any matching window.  This is based on the AutoIt3
// source code.
{
	char buf[WINDOW_TEXT_SIZE];
	if (*pWin->title) // Caller told us to search by class name and number.
	{
		GetClassName(aWnd, buf, sizeof(buf));
		// Below: i.e. this control's title (e.g. List) in contained entirely
		// within the leading part of the user specified title (e.g. ListBox).
		// Even though this is incorrect, the appending of the sequence number
		// in the second comparison will weed out any false matches.
		// Note: since some controls end in a number (e.g. SysListView32),
		// it would not be easy to parse out the user's sequence number to
		// simplify/accelerate the search here.  So instead, use a method
		// more certain to work even though it's a little ugly.  It's also
		// necessary to do this in a way functionally identical to the below
		// so that Window Spy's sequence numbers match the ones generated here:
		if (!strnicmp(pWin->title, buf, strlen(buf)))
		{
			// Use this var, initialized to zero by constructor, to accumulate the found-count:
			++pWin->already_visited_count;
			snprintfcat(buf, sizeof(buf), "%u", pWin->already_visited_count); // Append the count to the class.
			if (!stricmp(buf, pWin->title)) // It matches name and number.
			{
				pWin->child_hwnd = aWnd; // save this in here for return to the caller.
				return FALSE; // stop the enumeration.
			}
		}
	}
	else // Caller told us to search by the text of the control (e.g. the text printed on a button)
	{
		// Use GetWindowText() rather than GetWindowTextTimeout() because we don't want to find
		// the name accidentally in the vast amount of text present in some edit controls (e.g.
		// if the script's source code is open for editing in notepad, GetWindowText() would
		// likely find an unwanted match for just about anything).  In addition,
		// GetWindowText() is much faster.  Update: Yes, it seems better not to use
		// GetWindowTextByTitleMatchMode() in this case, since control names tend to be so
		// short (i.e. they would otherwise be very likely to be undesirably found in any large
		// edit controls the target window happens to own).  Update: Changed from strstr()
		// to strncmp() for greater selectivity.  Even with this degree of selectivity, it's
		// still possible to have ambiguous situations where a control can't be found due
		// to its title being entirely contained within that of another (e.g. a button
		// with title "Connect" would be found in the title of a button "Connect All").
		// The only way to address that would be to insist on an entire title match, but
		// that might be tedious if the title of the control is very long.  As alleviation,
		// the class name + seq. number method above can often be used instead in cases
		// of such ambiguity.  Update: Using IsTextMatch() now so that user-specified
		// TitleMatchMode will be in effect for this also.  Also, it's case sensitivity
		// helps increase selectivity, which is helpful due to how common short or ambiguous
		// control names tend to be:
		GetWindowText(aWnd, buf, sizeof(buf));
		if (IsTextMatch(buf, pWin->text))
		{
			pWin->child_hwnd = aWnd; // save this in here for return to the caller.
			return FALSE;
		}
	}
	// Note: The MSDN docs state that EnumChildWindows already handles the
	// recursion for us: "If a child window has created child windows of its own,
	// EnumChildWindows() enumerates those windows as well."
	return TRUE; // Keep searching.
}



int MsgBox(int aValue)
{
	char str[128];
	snprintf(str, sizeof(str), "Value = %d (0x%X)", aValue, aValue);
	return MsgBox(str);
}



int MsgBox(char *aText, UINT uType, char *aTitle, double aTimeout)
// Returns FAIL if the attempt failed because of too many existing MessageBox windows,
// or if MessageBox() itself failed.
{
	// Set the below globals so that any WM_TIMER messages dispatched by this call to
	// MsgBox() (which may result in a recursive call back to us) know not to display
	// any more MsgBoxes:
	if (g_nMessageBoxes > MAX_MSGBOXES + 1)  // +1 for the final warning dialog.  Verified correct.
		return FAIL;
	if (g_nMessageBoxes == MAX_MSGBOXES)
	{
		// Do a recursive call to self so that it will be forced to the foreground.
		// But must increment this so that the recursive call allows the last MsgBox
		// to be displayed:
		++g_nMessageBoxes;
		MsgBox("The maximum number of MsgBoxes has been reached.");
		--g_nMessageBoxes;
		return FAIL;
	}

	// Set these in case the caller explicitly called it with a NULL, overriding the default:
	if (!aText)
		aText = "";
	if (!aTitle || !*aTitle)
		// If available, the script's filename seems a much better title in case the user has
		// more than one script running:
		aTitle = (g_script.mFileName && *g_script.mFileName) ? g_script.mFileName : NAME_PV;

	// It doesn't feel safe to modify the contents of the caller's aText and aTitle,
	// even if the caller were to tell us it is modifiable.  This is because the text
	// might be the actual contents of a variable, which we wouldn't want to truncate,
	// even temporarily, since other hotkeys can fire while this hotkey subroutine is
	// suspended, and those subroutines may refer to the contents of this (now-altered)
	// variable.  In addition, the text may reside in the clipboard's locked memory
	// area, and altering that might result in the clipboard's contents changing
	// when MsgSleep() closes the clipboard for us (after we display our dialog here).
	// Even though testing reveals that the contents aren't altered (somehow), it
	// seems best to have our own local, limited length versions here:
	// Note: 8000 chars is about the max you could ever fit on-screen at 1024x768 on some
	// XP systems, but it will hold much more before refusing to display at all (i.e.
	// MessageBox() returning failure), perhaps about 150K:
	char text[MSGBOX_TEXT_SIZE];
	char title[DIALOG_TITLE_SIZE];
	strlcpy(text, aText, sizeof(text));
	strlcpy(title, aTitle, sizeof(title));

	uType |= MB_SETFOREGROUND;  // Always do these so that caller doesn't have to specify.

	// In the below, make the MsgBox owned by the topmost window rather than our main
	// window, in case there's another modal dialog already displayed.  The forces the
	// user to deal with the modal dialogs starting with the most recent one, which
	// is what we want.  Otherwise, if a middle dialog was dismissed, it probably
	// won't be able to return which button was pressed to its original caller.
	// UPDATE: It looks like these modal dialogs can't own other modal dialogs,
	// so disabling this:
	/*
	HWND topmost = GetTopWindow(g_hWnd);
	if (!topmost) // It has no child windows.
		topmost = g_hWnd;
	*/

	// Unhide the main window, but have it minimized.  This creates a task
	// bar button so that it's easier the user to remember that a dialog
	// is open and waiting (there are probably better ways to handle
	// this whole thing).  UPDATE: This isn't done because it seems
	// best not to have the main window be inaccessible until the
	// dialogs are dismissed (in case ever want to use it to display
	// status info, etc).  It seems that MessageBoxes get their own
	// task bar button when they're not AppModal, which is one of the
	// main things I wanted, so that's good too):
//	if (!IsWindowVisible(g_hWnd) || !IsIconic(g_hWnd))
//		ShowWindowAsync(g_hWnd, SW_SHOWMINIMIZED);

	/*
	If the script contains a line such as "#y::MsgBox, test", and a hotkey is used
	to activate Windows Explorer and another hotkey is then used to invoke a MsgBox,
	that MsgBox will be psuedo-minimized or invisible, even though it does have the
	input focus.  This attempt to fix it didn't work, so something is probably checking
	the physical key state of LWIN/RWIN and seeing that they're down:
	modLR_type modLR_now = GetModifierLRState();
	modLR_type win_keys_down = modLR_now & (MOD_LWIN | MOD_RWIN);
	if (win_keys_down)
		SetModifierLRStateSpecific(win_keys_down, modLR_now, KEYUP);
	*/

	// Note: Even though when multiple messageboxes exist, they might be
	// destroyed via a direct call to their WindowProc from our message pump's
	// DispatchMessage, or that of another MessageBox's message pump, it
	// appears that MessageBox() is designed to be called recursively like
	// this, since it always returns the proper result for the button on the
	// actual MessageBox it originally invoked.  In other words, if a bunch
	// of Messageboxes are displayed, and this user dismisses an older
	// one prior to dealing with a newer one, all the MessageBox()
	// return values will still wind up being correct anyway, at least
	// on XP.  The only downside to the way this is designed is that
	// the keyboard can't be used to navigate the buttons on older
	// messageboxes (only the most recent one).  This is probably because
	// the message pump of MessageBox() isn't designed to properly dispatch
	// keyboard messages to other MessageBox window instances.  I tried
	// to fix that by making our main message pump handle all messages
	// for all dialogs, but that turns out to be pretty complicated, so
	// I abandoned it for now.

	// Note: It appears that MessageBox windows, and perhaps all modal dialogs in general,
	// cannot own other windows.  That's too bad because it would have allowed each new
	// MsgBox window to be owned by any previously existing one, so that the user would
	// be forced to close them in order if they were APPL_MODAL.  But it's not too big
	// an issue since the only disadvantage is that the keyboard can't be use to
	// to navigate in MessageBoxes other than the most recent.  And it's actually better
	// the way it is now in the sense that the user can dismiss the messageboxes out of
	// order, which might (in rare cases) be desireable.

	POST_AHK_DIALOG((int)(aTimeout * 1000))

	++g_nMessageBoxes;  // This value will also be used as the Timer ID if there's a timeout.
	g.MsgBoxResult = MessageBox(NULL, text, title, uType);
	--g_nMessageBoxes;

//	if (!g_nMessageBoxes)
//		ShowWindowAsync(g_hWnd, SW_HIDE);  // Hide the main window if it no longer has any child windows.
//	else

	// This is done so that the next message box of ours will be brought to the foreground,
	// to remind the user that they're still out there waiting, and for convenience.
	// Update: It seems bad to do this in cases where the user intentionally wants the older
	// messageboxes left in the background, to deal with them later.  So, in those cases,
	// we might be doing more harm than good because the user's foreground window would
	// be intrusively changed by this:
	//WinActivateOurTopDialog();

	// Unfortunately, it appears that MessageBox() will return zero rather
	// than AHK_TIMEOUT that was specified in EndDialog() at least under WinXP.
	if (!g.MsgBoxResult && aTimeout > 0)
		// Assume it timed out rather than failed, since failure should be
		// VERY rare.
		g.MsgBoxResult = AHK_TIMEOUT;
	// else let the caller handle the display of the error message because only it knows
	// whether to also tell the user something like "the script will not continue".
	return g.MsgBoxResult;
}



HWND FindOurTopDialog()
// Returns the HWND of our topmost MsgBox or FileOpen dialog (and perhaps other types of modal
// dialogs if they are of class #32770) even if it wasn't successfully brought to
// the foreground here.
// Using Enum() seems to be the only easy way to do this, since these modal MessageBoxes are
// *owned*, not children of the main window.  There doesn't appear to be any easier way to
// find out which windows another window owns.  GetTopWindow(), GetActiveWindow(), and GetWindow()
// do not work for this purpose.  And using FindWindow() discouraged because it can hang
// in certain circumtances (Enum is probably just as fast anyway).
{
	// The return value of EnumWindows() is probably a raw indicator of success or failure,
	// not whether the Enum found something or continued all the way through all windows.
	// So don't bother using it.
	pid_and_hwnd_type pid_and_hwnd;
	pid_and_hwnd.pid = GetCurrentProcessId();
	pid_and_hwnd.hwnd = NULL;  // Init.  Called function will set this for us if it finds a match.
	EnumWindows(EnumDialog, (LPARAM)&pid_and_hwnd);
	return pid_and_hwnd.hwnd;
}



BOOL CALLBACK EnumDialog(HWND aWnd, LPARAM lParam)
// lParam should be a pointer to a ProcessId (ProcessIds are always non-zero?)
// To continue enumeration, the function must return TRUE; to stop enumeration, it must return FALSE. 
#define pThing ((pid_and_hwnd_type *)lParam)
{
	if (!lParam || !pThing->pid) return FALSE;
	DWORD pid;
	GetWindowThreadProcessId(aWnd, &pid);
	if (pid == pThing->pid)
	{
		char buf[32];
		GetClassName(aWnd, buf, sizeof(buf));
		// This is the class name for windows created via MessageBox(), GetOpenFileName(), and probably
		// other things that use modal dialogs:
		if(!strcmp(buf, "#32770"))
		{
			pThing->hwnd = aWnd;  // An output value for the caller.
			return FALSE;  // We're done.
		}
	}
	return TRUE;  // Keep searching.
}



BOOL CALLBACK EnumDialogClose(HWND aWnd, LPARAM lParam)
// lParam should be a pointer to a ProcessId (ProcessIds are always non-zero?)
// To continue enumeration, the function must return TRUE; to stop enumeration, it must return FALSE. 
#define pThing ((pid_and_hwnd_type *)lParam)
{
	if (!lParam || !pThing->pid) return FALSE;
	DWORD pid;
	GetWindowThreadProcessId(aWnd, &pid);
	if (pid == pThing->pid)
	{
		char buf[32];
		GetClassName(aWnd, buf, sizeof(buf));
		// This is the class name for windows created via MessageBox(), GetOpenFileName(), and probably
		// other things that use modal dialogs:
		if(!strcmp(buf, "#32770"))
		{
			// Since it's our window, I think this will effectively use our thread to immediately
			// call the WindowProc() of the target dialog.  Testing reveals that under WinXP at least,
			// this call does not destroy the windows (WM_CLOSE).  However, it seems better to use
			// Send than Post in the hopes that Send has immediately caused the WindowProc to do
			// things which indicate to the OS that the dialogs are marked for destruction.
			// Note: Not supposed to call EndDialog() outside of a DialogProc(), so do this instead.
			// UPDATE: Sending WM_QUIT vs. WM_CLOSE immediately terminates the dialogs, which seems
			// better in light of the fact that our caller is trying to exit the program immediately.
			// UPDATE #2: That is unrepeatable, I don't know how it happened that once.  Still, WM_QUIT
			// might be better than WM_CLOSE in this case:
			SendMessage(aWnd, WM_QUIT, 0, 0);
			pThing->hwnd = aWnd;  // An output value for the caller so that it knows we closed at least one.
		}
	}
	return TRUE;  // Keep searching so that all our dialogs will be sent the message.
}



struct owning_struct {HWND owner_hwnd; HWND first_child;};
HWND WindowOwnsOthers(HWND aWnd)
// Only finds owned windows if they are visible, by design.
{
	owning_struct own = {aWnd, NULL};
	EnumWindows(EnumParentFindOwned, (LPARAM)&own);
	return own.first_child;
}



BOOL CALLBACK EnumParentFindOwned(HWND aWnd, LPARAM lParam)
{
	HWND owner_hwnd = GetWindow(aWnd, GW_OWNER);
	// Note: Many windows seem to own other invisible windows that have blank titles.
	// In our case, require that it be visible because we don't want to return an invisible
	// window to the caller because such windows aren't designed to be activated:
	if (owner_hwnd && owner_hwnd == ((owning_struct *)lParam)->owner_hwnd && IsWindowVisible(aWnd))
	{
		((owning_struct *)lParam)->first_child = aWnd;
		return FALSE; // Match found, we're done.
	}
	return TRUE;  // Continue enumerating.
}



HWND GetTopChild(HWND aParent)
{
	if (!aParent) return aParent;
	HWND hwnd_top, next_top;
	// Get the topmost window of the topmost window of...
	// i.e. Since child windows can also have children, we keep going until
	// we reach the "last topmost" window:
	for (hwnd_top = GetTopWindow(aParent)
		; hwnd_top && (next_top = GetTopWindow(hwnd_top))
		; hwnd_top = next_top);

	//if (!hwnd_top)
	//{
	//	MsgBox("no top");
	//	return FAIL;
	//}
	//else
	//{
	//	//if (GetTopWindow(hwnd_top))
	//	//	hwnd_top = GetTopWindow(hwnd_top);
	//	char class_name[64];
	//	GetClassName(next_top, class_name, sizeof(class_name));
	//	MsgBox(class_name);
	//}

	return hwnd_top ? hwnd_top : aParent;  // Caller relies on us never returning NULL if aParent is non-NULL.
}



bool IsWindowHung(HWND aWnd)
{
	if (!aWnd) return false;

	// OLD, SLOWER METHOD:
	// Don't want to use a long delay because then our messages wouldn't get processed
	// in a timely fashion.  But I'm not entirely sure if the 10ms delay used below
	// is even used by the function in this case?  Also, the docs aren't clear on whether
	// the function returns success or failure if the window is hung (probably failure).
	// If failure, perhaps you have to call GetLastError() to determine whether it failed
	// due to being hung or some other reason?  Does the output param dwResult have any
	// useful info in this case?  I expect what will happen is that in most cases, the OS
	// will already know that the window is hung.  However, if the window just became hung
	// in the last 5 seconds, I think it may take the remainder of the 5 seconds for the OS
	// to notice it.  However, allowing it the option of sleeping up to 5 seconds seems
	// really bad, since keyboard and mouse input would probably be frozen (actually it
	// would just be really laggy because the OS would bypass the hook during that time).
	// So some compromise value seems in order.  500ms seems about right.  UPDATE: Some
	// windows might need longer than 500ms because their threads are engaged in
	// heavy operations.  Since this method is only used as a fallback method now,
	// it seems best to give them the full 5000ms default, which is what (all?) Windows
	// OSes use as a cutoff to determine whether a window is "not responding":
	DWORD dwResult;
	#define Slow_IsWindowHung !SendMessageTimeout(aWnd, WM_NULL, (WPARAM)0, (LPARAM)0\
		, SMTO_ABORTIFHUNG, 5000, &dwResult)

	// NEW, FASTER METHOD:
	// This newer method's worst-case performance is at least 30x faster than the worst-case
	// performance of the old method that  uses SendMessageTimeout().
	// And an even worse case can be envisioned which makes the use of this method
	// even more compelling: If the OS considers a window NOT to be hung, but the
	// window's message pump is sluggish about responding to the SendMessageTimeout() (perhaps
	// taking 2000ms or more to respond due to heavy disk I/O or other activity), the old method
	// will take several seconds to return, causing mouse and keyboard lag if our hook(s)
	// are installed; not to mention making our app's windows, tray menu, and other GUI controls
	// unresponsive during that time).  But I believe in this case the new method will return
	// instantly, since the OS has been keeping track in the background, and can tell us
	// immediately that the window isn't hung.
	// Here are some seemingly contradictory statements uttered by MSDN.  Perhaps they're
	// not contradictory if the first sentence really means "by a different thread of the same
	// process":
	// "If the specified window was created by a different thread, the system switches to that
	// thread and calls the appropriate window procedure.  Messages sent between threads are
	// processed only when the receiving thread executes message retrieval code. The sending
	// thread is blocked until the receiving thread processes the message."
	if (g_os.IsWin9x())
	{
		typedef BOOL (WINAPI *MyIsHungThread)(DWORD);
		static MyIsHungThread IsHungThread = (MyIsHungThread)GetProcAddress(GetModuleHandle("User32.dll")
			, "IsHungThread");
		// When function not available, fall back to the old method:
		return IsHungThread ? IsHungThread(GetWindowThreadProcessId(aWnd, NULL)) : Slow_IsWindowHung;
	}
	else // Otherwise: NT/2k/XP/2003 or some later OS (e.g. 64 bit?), so try to use the newer method.
	{
		typedef BOOL (WINAPI *MyIsHungAppWindow)(HWND);
		static MyIsHungAppWindow IsHungAppWindow = (MyIsHungAppWindow)GetProcAddress(GetModuleHandle("User32.dll")
			, "IsHungAppWindow");
		// When function not available, fall back to the old method:
		return IsHungAppWindow ? IsHungAppWindow(aWnd) : Slow_IsWindowHung;
	}
}



int GetWindowTextTimeout(HWND aWnd, char *aBuf, int aBufSize, UINT aTimeout)
// Returns the length of what would be copied (not including the zero terminator).
// In addition, if aBuf is not NULL, the window text is copied into aBuf (not to exceed aBufSize).
// AutoIt3 author indicates that using WM_GETTEXT vs. GetWindowText() sometimes yields more text.
// Perhaps this is because GetWindowText() has built-in protection against hung windows and
// thus isn't actually sending WM_GETTEXT.  The method here is hopefully the best of both worlds
// (protection against hung windows causing our thread to hang, and getting more text).
// Another tidbit from MSDN about SendMessage() that might be of use here sometime:
// "However, the sending thread will process incoming nonqueued (those sent directly to a window
// procedure) messages while waiting for its message to be processed. To prevent this, use
// SendMessageTimeout with SMTO_BLOCK set."  Currently not using SMTO_BLOCK because it
// doesn't seem necessary.
// Update: GetWindowText() is so much faster than SendMessage() and SendMessageTimeout(), at
// least on XP, so GetWindowTextTimeout() should probably only be used when getting the max amount
// of text is important (e.g. this function can fetch the text in a RichEdit20A control and
// other edit controls, whereas GetWindowText() doesn't).  This function is used to implement
// things like WinGetText and ControlGetText, in which getting the maximum amount and types
// of text is more important than performance.
{
	if (!aWnd) return 0; // Seems better than -1 or some error code.
	if (aBuf && aBufSize < 1) aBuf = NULL; // Just return the length.
	if (aBuf) *aBuf = '\0';  // Init just to get it out of the way in case of early return/error.
	// Override for Win95 because AutoIt3 author says it might crash otherwise:
	if (aBufSize > WINDOW_TEXT_SIZE && g_os.IsWin95()) aBufSize = WINDOW_TEXT_SIZE;
	DWORD result;
	LRESULT lresult;
	if (aBuf)
	{
		// Below demonstrated that GetWindowText() is dramatically faster than either SendMessage()
		// or SendMessageTimeout() (noticeably faster when you have hotkeys that activate
		// windows, or toggle between two windows):
		//return GetWindowText(aWnd, aBuf, aBufSize);
		//return (int)SendMessage(aWnd, WM_GETTEXT, (WPARAM)aBufSize, (LPARAM)aBuf);

		// Don't bother calling IsWindowHung() because the below call will return
		// nearly instantly if the OS already "knows" that the target window has
		// be unresponsive for 5 seconds or so (i.e. it keeps track of such things
		// on an ongoing basis, at least XP seems to).
		lresult = SendMessageTimeout(aWnd, WM_GETTEXT, (WPARAM)aBufSize, (LPARAM)aBuf
			, SMTO_ABORTIFHUNG, aTimeout, &result);
		// Just to make sure because MSDN docs aren't clear that it will always be terminated:
		aBuf[aBufSize - 1] = '\0';
	}
	else
		lresult = SendMessageTimeout(aWnd, WM_GETTEXTLENGTH, (WPARAM)0, (LPARAM)0  // Both must be zero.
			, SMTO_ABORTIFHUNG, aTimeout, &result);
	if (!lresult) // It failed or timed out.
		return 0;
	// <result> contains the length of what was (or would have been) copied, not including the terminator:
	return (int)result;
}



void SetForegroundLockTimeout()
{
	// Even though they may not help in all OSs and situations, this lends peace-of-mind.
	// (it doesn't appear to help on my XP?)
	if (g_os.IsWin98orLater() || g_os.IsWin2000orLater())
	{
		// Don't check for failure since this operation isn't critical, and don't want
		// users continually haunted by startup error if for some reason this doesn't
		// work on their system:
		if (SystemParametersInfo(SPI_GETFOREGROUNDLOCKTIMEOUT, 0, &g_OriginalTimeout, 0))
			if (g_OriginalTimeout) // Anti-focus stealing measure is in effect.
			{
				// Set it to zero instead, disabling the measure:
				SystemParametersInfo(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, (PVOID)0, SPIF_SENDCHANGE);
//				if (!SystemParametersInfo(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, (PVOID)0, SPIF_SENDCHANGE))
//					MsgBox("Enable focus-stealing: set-call to SystemParametersInfo() failed.");
			}
//			else
//				MsgBox("Enable focus-stealing: it was already enabled.");
//		else
//			MsgBox("Enable focus-stealing: get-call to SystemParametersInfo() failed.");
	}
//	else
//		MsgBox("Enable focus-stealing: neither needed nor supported under Win95 and WinNT.");
}
