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
#include "WinGroup.h"
#include "window.h" // for several lower level window functions
#include "globaldata.h" // for DoWinDelay
#include "application.h" // for DoWinDelay's MsgSleep()

// Define static members data:
WinGroup *WinGroup::sGroupLastUsed = NULL;
HWND *WinGroup::sAlreadyVisited = NULL;
int WinGroup::sAlreadyVisitedCount = 0;


ResultType WinGroup::AddWindow(char *aTitle, char *aText, void *aJumpToLine
	, char *aExcludeTitle, char *aExcludeText)
// Caller should ensure that at least one param isn't NULL/blank.
// GroupActivate will tell its caller to jump to aJumpToLine if a WindowSpec isn't found.
{
	if (!aTitle) aTitle = "";
	if (!aText) aText = "";
	if (!aExcludeTitle) aExcludeTitle = "";
	if (!aExcludeText) aExcludeText = "";

	// SimpleHeap::Malloc() will set these new vars to the constant empty string if their
	// corresponding params are blank:
	char *new_title, *new_text, *new_exclude_title, *new_exclude_text;
	if (!(new_title = SimpleHeap::Malloc(aTitle))) return FAIL; // It already displayed the error for us.
	if (!(new_text = SimpleHeap::Malloc(aText)))return FAIL;
	if (!(new_exclude_title = SimpleHeap::Malloc(aExcludeTitle))) return FAIL;
	if (!(new_exclude_text = SimpleHeap::Malloc(aExcludeText)))   return FAIL;

	WindowSpec *the_new_win = new WindowSpec(new_title, new_text, aJumpToLine, new_exclude_title, new_exclude_text);
	if (the_new_win == NULL)
		return g_script.ScriptError("WinGroup::AddWindow(): Out of memory.");
	if (mFirstWindow == NULL)
		mFirstWindow = mLastWindow = the_new_win;
	else
	{
		mLastWindow->mNextWindow = the_new_win; // Formerly it pointed to First, so nothing is lost here.
		// This must be done after the above:
		mLastWindow = the_new_win;
	}
	// Make it circular: Last always points to First.  It's okay if it points to itself:
	mLastWindow->mNextWindow = mFirstWindow;
	++mWindowCount;
	return OK;
}



ResultType WinGroup::CloseAll()
{
	if (IsEmpty())
		return OK;  // OK since this is the expected behavior in this case.
	// Don't need to call Update() in this case.
	// Close all windows that match any WindowSpec in the group:
	WindowInfoPackage wip;
	wip.win_spec = mFirstWindow;
	EnumWindows(EnumParentCloseAny, (LPARAM)&wip);
	if (wip.parent_hwnd) // It closed at least one window.
		DoWinDelay;
	return OK;
}



ResultType WinGroup::CloseAndGoToNext(bool aStartWithMostRecent)
// If the foreground window is a member of this group, close it and activate
// the next member.
{
	if (IsEmpty())
		return OK;  // OK since this is the expected behavior in this case.
	// Otherwise:
	// Don't call Update(), let (De)Activate() do that.
	WindowSpec *win_spec = IsMember(GetForegroundWindow());
	if (   (mIsModeActivate && win_spec) || (!mIsModeActivate && !win_spec)   )
	{
		// If the user is using a GroupActivate hotkey, we don't want to close
		// the foreground window if it's not a member of the group.  Conversely,
		// if the user is using GroupDeactivate, we don't want to close a
		// member of the group.  This precaution helps prevent accidental closing
		// of windows that suddenly pop up to the foreground just as you've
		// realized (too late) that you pressed the "close" hotkey.
		// MS Visual Studio/C++ gets messed up when it is directly sent a WM_CLOSE,
		// probably because the wrong window (it has two mains) is being sent the close.
		// But since that's the only app I've ever found that doesn't work right,
		// it seems best not to change our close method just for it because sending
		// keys is a fairly high overhead operation, and not without some risk due to
		// not knowing exactly what keys the user may have physically held down.
		// Also, we'd have to make this module dependent on the keyboard module,
		// which would be another drawback.
		// Try to wait for it to close, otherwise the same window may be activated
		// again before it has been destroyed, defeating the purpose of the
		// "ActivateNext" part of this function's job:
		// SendKeys("!{F4}");
		WinClose("a", "", 500); // a=active; Use this rather than PostMessage because it will wait-for-close.
		DoWinDelay;
	}
	return mIsModeActivate ? Activate(aStartWithMostRecent, win_spec) : Deactivate(aStartWithMostRecent);
}



ResultType WinGroup::Activate(bool aStartWithMostRecent, WindowSpec *aWinSpec, void **aJumpToLine)
{
	// Be sure initialize this before doing any returns:
	if (aJumpToLine)
		*aJumpToLine = NULL;
	if (IsEmpty())
		return OK;  // OK since this is the expected behavior in this case.
	// Otherwise:
	if (!Update(true)) // Update our private member vars.
		return FAIL;  // It already displayed the error for us.
	WindowSpec *win, *win_to_activate_next = aWinSpec;
	bool group_is_active = false; // Set default.
	HWND activate_win, fore_hwnd = GetForegroundWindow(); // This value is used in more than one place.
	if (win_to_activate_next)
	{
		// The caller told us which WindowSpec to start off trying to activate.
		// If the foreground window matches that WindowSpec, do nothing except
		// marking it as visited, because we want to stay on this window under
		// the assumption that it was newly revealed due to a window on top
		// of it having just been closed:
		if (win_to_activate_next == IsMember(fore_hwnd))
		{
			group_is_active = true;
			MarkAsVisited(fore_hwnd);
			return OK;
		}
		// else don't mark as visited even if it's a member of the group because
		// we're about to attempt to activate a different window: the next
		// unvisited member of this same WindowSpec.  If the below doesn't
		// find any of those, it continue on through the list normally.
	}
	else // Caller didn't tell us which, so determine it.
	{
		if (win_to_activate_next = IsMember(fore_hwnd)) // Foreground window is a member of this group.
		{
			// Set it to activate this same WindowSpec again in case there's
			// more than one that matches (e.g. multiple notepads).  But first,
			// mark the current window as having been visited if it hasn't
			// already by marked by a prior iteration.  Update: This method
			// doesn't work because if a unvisted matching window became the
			// foreground window by means other than using GroupActivate
			// (e.g. launching a new instance of the app: now there's another
			// matching window in the foreground).  So just call it straight
			// out.  It has built-in dupe-checking which should prevent the
			// list from filling up with dupes if there are any special
			// situations in which that might otherwise happen:
			//if (!sAlreadyVisitedCount)
			group_is_active = true;
			MarkAsVisited(fore_hwnd);
		}
		else // It's not a member.
		{
			win_to_activate_next = mFirstWindow;  // We're starting fresh, so start at the first window.
			// Reset the list of visited windows:
			sAlreadyVisitedCount = 0;
		}
	}

	// Activate any unvisited window that matches the win_to_activate_next spec.
	// If none, activate the next window spec in the series that does have an
	// existing window:
	// If the spec we're starting at already has some windows marked as visited,
	// set this variable so that we know to retry the first spec again in case
	// a full circuit is made through the window specs without finding a window
	// to activate.  Note: Using >1 vs. >0 might protect against any infinite-loop
	// conditions that may be lurking:
	bool retry_starting_win_spec = (sAlreadyVisitedCount > 1);
	bool retry_is_in_effect = false;
	for (win = win_to_activate_next;;)
	{
		// Call this in the mode to find the last match, which  makes things nicer
		// because when the sequence wraps around to the beginning, the windows will
		// occur in the same order that they did the first time, rather than going
		// backwards through the sequence (which is counterintuitive for the user):
		if (   activate_win = WinActivate(win->mTitle, win->mText, win->mExcludeTitle, win->mExcludeText
			// This next line is whether to find last or first match.  We always find the oldest
			// (bottommost) match except when the user has specifically asked to start with the
			// most recent.  But it only makes sense to start with the most recent if the
			// group isn't currently active (i.e. we're starting fresh), because otherwise
			// windows would be activated in an order different from what was already shown
			// the first time through the enumeration, which doesn't seem to be ever desirable:
			, !aStartWithMostRecent || group_is_active
			, sAlreadyVisited, sAlreadyVisitedCount)   )
		{
			// We found a window to activate, so we're done.
			// Probably best to do this before WinDelay in case another hotkey fires during the delay:
			MarkAsVisited(activate_win);
			DoWinDelay;
			//MsgBox(win->mText, 0, win->mTitle);
			break;
		}
		// Otherwise, no window was found to activate.
		if (aJumpToLine && win->mJumpToLine && !sAlreadyVisitedCount)
		{
			// Caller asked us to return in this case, so that it can
			// use this value to execute a user-specified Gosub:
			*aJumpToLine = win->mJumpToLine;  // Set output param for the caller.
			return OK;
		}
		if (retry_is_in_effect)
			// This was the final attempt because we've already gone all the
			// way around the circular linked list of WindowSpecs.  This check
			// must be done, otherwise an infinite loop might result if the windows
			// that formed the basis for determining the value of
			// retry_starting_win_spec have since been destroyed:
			break;
		// Otherwise, go onto the next one in the group:
		win = win->mNextWindow;
        // Even if the above didn't change the value of <win> (because there's only
		// one WinSpec in the list), it's still correct to reset this count because
		// we want to start the fresh again after all the windows have been
		// visited.  Note: The only purpose of sAlreadyVisitedCount as used by
		// this function is to indicate which windows in a given WindowSpec have
		// been visited, not which windows altogether (i.e. it's not necessary to
		// remember which windows have been visited once we move on to a new
		// WindowSpec).
		sAlreadyVisitedCount = 0;
		if (win == win_to_activate_next)
		{
			// We've made one full circuit of the circular linked list without
			// finding an existing window to activate. At this point, the user
			// has pressed a hotkey to do a GroupActivate, but nothing has happened
			// yet.  We always want something to happen unless there's absolutely
			// no existing windows to activate, or there's only a single window in
			// the system that matches the group and it's already active.
			if (retry_starting_win_spec)
			{
				// Mark the foreground window as visited so that it won't be
				// mistakenly activated again by the next iteration:
				MarkAsVisited(fore_hwnd);
				retry_is_in_effect = true;
				// Now continue with the next iteration of the loop so that it
				// will activate a different instance of this WindowSpec rather
				// than getting stuck on this one.
			}
			else
				break;
		}
	}
	return OK;
}



ResultType WinGroup::Deactivate(bool aStartWithMostRecent)
{
	if (IsEmpty())
		return OK;  // OK since this is the expected behavior in this case.
	// Otherwise:
	if (!Update(false)) // Update our private member vars.
		return FAIL;  // It already displayed the error for us.

	HWND fore_hwnd = GetForegroundWindow();
	if (IsMember(fore_hwnd))
		sAlreadyVisitedCount = 0;

	// Activate the next unvisited non-member:
	WindowInfoPackage wip;
	wip.already_visited = sAlreadyVisited;
	wip.already_visited_count = sAlreadyVisitedCount;
	wip.win_spec = mFirstWindow;
	wip.find_last_match = !aStartWithMostRecent || sAlreadyVisitedCount;
	EnumWindows(EnumParentFindAnyExcept, (LPARAM)&wip);
	if (wip.parent_hwnd)
	{
		// If the window we're about to activate owns other visble parent windows, it can
		// never truly be activated because it must always be below them in the z-order.
		// Thus, instead of activating it, activate the first (and usually the only?)
		// visible window that it owns.  Doing this makes things nicer for some apps that
		// have a pair of main windows, such as MS Visual Studio (and probably many more),
		// because it avoids activating such apps twice in a row as the user progresses
		// through the sequence:
		HWND first_visible_owned = WindowOwnsOthers(wip.parent_hwnd);
		if (first_visible_owned)
		{
			MarkAsVisited(wip.parent_hwnd);  // Must mark owner as well as the owned window.
			// Activate the owned window instead of the owner because it usually
			// (probably always, given the comments above) is the real main window:
			wip.parent_hwnd = first_visible_owned;
		}
		SetForegroundWindowEx(wip.parent_hwnd);
		// Probably best to do this before WinDelay in case another hotkey fires during the delay:
		MarkAsVisited(wip.parent_hwnd);
		DoWinDelay;
	}
	else // No window was found to activate (they have all been visited).
	{
		if (sAlreadyVisitedCount)
		{
			bool wrap_around = (sAlreadyVisitedCount > 1);
			sAlreadyVisitedCount = 0;
			if (wrap_around)
			{
				// The user pressed a hotkey to do something, yet nothing has happened yet.
				// We want something to happen every time if there's a qualifying
				// "something" that we can do.  And in this case there is: we can start
				// over again through the list, excluding the foreground window (which
				// the user has already had a chance to review):
				MarkAsVisited(fore_hwnd);
				// Make a recursive call to self.  This can't result in an infinite
				// recursion (stack fault) because the called layer will only
				// recurse a second time if sAlreadyVisitedCount > 1, which is
				// impossible with the current logic:
				Deactivate(false); // Seems best to ignore aStartWithMostRecent in this case?
			}
		}
	}
	// Even if a window wasn't found, we've done our job so return OK:
	return OK;
}



inline ResultType WinGroup::Update(bool aIsModeActivate)
{
	mIsModeActivate = aIsModeActivate;
	if (sGroupLastUsed != this)
	{
		sGroupLastUsed = this;
		sAlreadyVisitedCount = 0; // Since it's a new group, reset the array to start fresh.
	}
	if (!sAlreadyVisited) // Allocate the array on first use.
		// Getting it from SimpleHeap reduces overhead for the avg. case (i.e. the first
		// block of SimpleHeap is usually never fully used, and this array won't even
		// be allocated for short scripts that don't even using window groups.
		if (   !(sAlreadyVisited = (HWND *)SimpleHeap::Malloc(MAX_ALREADY_VISITED * sizeof(HWND)))   )
			return FAIL;  // It already displayed the error for us.
	return OK;
}



inline WindowSpec *WinGroup::IsMember(HWND aWnd)
{
	if (!aWnd)
		return NULL;  // Caller relies on us to return "no match" in this case.
	char fore_title[WINDOW_TEXT_SIZE];
	if (GetWindowText(aWnd, fore_title, sizeof(fore_title)))
	{
		for (WindowSpec *win = mFirstWindow;;)
		{
			if (IsTextMatch(fore_title, win->mTitle, win->mExcludeTitle))
				if (HasMatchingChild(aWnd, win->mText, win->mExcludeText))
					return win;
			// Otherwise, no match, so go onto the next one:
			win = win->mNextWindow;
			if (win == mFirstWindow)
				// We've made one full circuit of the circular linked list,
				// discovering that the foreground window isn't a member
				// of the group:
				break;
		}
	}
	return NULL;  // Because it would have returned already if a match was found.
}


/////////////////////////////////////////////////////////////////////////


BOOL CALLBACK EnumParentFindAnyExcept(HWND aWnd, LPARAM lParam)
// Find the first parent window that doesn't match any of the WindowSpecs in
// the linked list, and that hasn't already been visited.
// Caller must have ensured that lParam isn't NULL.
// lParam must contain the address of a WindowSpec object.
{
	if (!IsWindowVisible(aWnd))
		// Skip these because we alwayswant them to stay invisible, regardless
		// of the setting for g.DetectHiddenWindows:
		return TRUE;
	LONG style = GetWindowLong(aWnd, GWL_EXSTYLE);
	if (style & WS_EX_TOPMOST)
		// Skip always-on-top windows, such as AutoIt's SplashText, because they're already
		// visible so the user already knows about them, so there's no need to have them
		// presented for review:
		return TRUE;
	char win_title[WINDOW_TEXT_SIZE];
	if (!GetWindowText(aWnd, win_title, sizeof(win_title)))
		return TRUE;  // Even if can't get the text of some window, for some reason, keep enumerating.
	if (!stricmp(win_title, "Program Manager"))
		// Skip this too because activating it would serve no purpose.  This is probably the
		// same HWND that GetShellWindow() returns, but GetShellWindow() isn't supported on
		// Win9x or WinNT, so don't bother using it.  And GetDeskTopWindow() apparently doesn't
		// return "Program Manager" (something with a blank title I think):
		return TRUE;

	for (WindowSpec *win = pWin->win_spec;;)
	{
		// For each window in the linked list, check if aWnd is a match
		// for it:
		if (IsTextMatch(win_title, win->mTitle, win->mExcludeTitle))
			if (HasMatchingChild(aWnd, win->mText, win->mExcludeText))
				// Match found, so aWnd is a member of the group.
				// But we want to find non-members only, so keep
				// searching:
				return TRUE;
		// Otherwise, no match, keep checking until aWnd has been compared against
		// all the WindowSpecs in the group:
		win = win->mNextWindow;
		if (win == pWin->win_spec)
		{
			// We've made one full circuit of the circular linked list without
			// finding a match.  So aWnd is the one we're looking for unless
			// it's in the list of exceptions:
			if (pWin->already_visited_count && pWin->already_visited)
				for (int i = 0; i < pWin->already_visited_count; ++i)
					if (aWnd == pWin->already_visited[i])
						return TRUE; // It's an exception, so keep searching.
			// Otherwise, this window meets the criteria, so return it to the caller and
			// stop the enumeration.  UPDATE: Rather than stopping the enumeration,
			// continue on through all windows so that the last match is found.
			// That makes things nicer because when the sequence wraps around to the
			// beginning, the windows will occur in the same order that they did
			// the first time, rather than going backwards through the sequence
			// (which is counterintuitive for the user):
			pWin->parent_hwnd = aWnd;
			return pWin->find_last_match; // bool vs. BOOL should be okay in this case.
		}
	}
}



BOOL CALLBACK EnumParentCloseAny(HWND aWnd, LPARAM lParam)
// Caller must have ensured that lParam isn't NULL.
// lParam must contain the address of a WindowSpec object.
{
	if (!IsWindowVisible(aWnd))
		// Skip these because it seems safest to never close invisible windows --
		// regardless of the setting of g.DetectHiddenWindows -- because of the
		// slight risk that some important hidden system window would accidentally
		// match one of the WindowSpecs in the group:
		return TRUE;
	char win_title[WINDOW_TEXT_SIZE];
	if (!GetWindowText(aWnd, win_title, sizeof(win_title)))
		return TRUE;  // Even if can't get the text of some window, for some reason, keep enumerating.
	if (!stricmp(win_title, "Program Manager"))
		// Skip this too because never want to close it as part of a group close.
		return TRUE;
	for (WindowSpec *win = pWin->win_spec;;)
	{
		// For each window in the linked list, check if aWnd is a match
		// for it:
		if (IsTextMatch(win_title, win->mTitle, win->mExcludeTitle))
			if (HasMatchingChild(aWnd, win->mText, win->mExcludeText))
			{
				// Match found, so aWnd is a member of the group.
				pWin->parent_hwnd = aWnd;  // So that the caller knows we closed at least one.
				PostMessage(aWnd, WM_CLOSE, 0, 0);  // Ask it nicely to close.
				return TRUE; // Continue the enumeration.
			}
		// Otherwise, no match, keep checking until aWnd has been compared against
		// all the WindowSpecs in the group:
		win = win->mNextWindow;
		if (win == pWin->win_spec)
			// We've made one full circuit of the circular linked list without
			// finding a match, so aWnd is not a member of the group and
			// should not be closed.
			return TRUE; // Continue the enumeration.
	}
}
