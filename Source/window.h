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

#ifndef window_h
#define window_h

#include "defines.h"
#include "globaldata.h"
#include "util.h" // for strlcpy()


// Note: it is apparently possible for a hidden window to be the foreground
// window (it just looks strange).  If DetectHiddenWindows is off, set
// target_window to NULL if it's hidden.  Doing this prevents, for example,
// WinClose() from closing the hidden foreground if it's some important hidden
// window like the shell or the desktop:
#define IF_USE_FOREGROUND_WINDOW(title, text, exclude_title, exclude_text)\
if ((*title == 'A' || *title == 'a') && !*(title + 1) && !*text && !*exclude_title && !*exclude_text)\
{\
	if (target_window = GetForegroundWindow())\
		if (!g.DetectHiddenWindows && !IsWindowVisible(target_window))\
			target_window = NULL;\
}

#define pWin ((WindowInfoPackage *)lParam)
#define SEARCH_PHRASE_SIZE 1024
// Info from AutoIt3 source: GetWindowText fails under 95 if >65535, WM_GETTEXT randomly fails if > 32767.
// My: And since 32767 is what AutoIt3 passes to the API functions as the size (not the length, i.e.
// it can only store 32766 if room is left for the zero terminator) we'll use that for the size too.
// Note: MSDN says (for functions like GetWindowText): "Specifies the maximum number of characters to
// copy to the buffer, including the NULL character. If the text exceeds this limit, it is truncated."
#define WINDOW_TEXT_SIZE 32767

struct WindowInfoPackage // A simple struct to help with EnumWindows().
{
	char title[SEARCH_PHRASE_SIZE];
	char text[SEARCH_PHRASE_SIZE];
	char exclude_title[SEARCH_PHRASE_SIZE];
	char exclude_text[SEARCH_PHRASE_SIZE];
	// Whether to keep searching even after a match is found, so that last one is found.
	bool find_last_match;
		// Above made into int vs. FindMatchType so it can be used for other purposes.
	HWND parent_hwnd, child_hwnd; // Returned to the caller, but the caller should initialize it to NULL beforehand.
	HWND *already_visited; // Array of HWNDs to exclude from consideration.
	int already_visited_count;
	WindowSpec *win_spec; // Linked list.
	WindowInfoPackage::WindowInfoPackage()
		: find_last_match(false) // default
		, parent_hwnd(NULL), child_hwnd(NULL), already_visited(NULL), already_visited_count(0)
		, win_spec(NULL)
	{
		// Can't use initializer list for these:
		*title = *text = *exclude_title = *exclude_text = '\0';
	}
};

struct pid_and_hwnd_type
{
	DWORD pid;
	HWND hwnd;
};

struct length_and_buf_type
{
	size_t total_length;
	size_t capacity;
	char *buf;
};


HWND WinActivate(char *aTitle, char *aText = "", char *aExcludeTitle = "", char *aExcludeText = ""
	, bool aFindLastMatch = false
	, HWND aAlreadyVisited[] = NULL, int aAlreadyVisitedCount = 0);
HWND SetForegroundWindowEx(HWND aWnd);

// Defaulting to a non-zero wait-time solves a lot of script problems that would otherwise
// require the user to specify the last param (or use WinWaitClose):
#define DEFAULT_WINCLOSE_WAIT 20
HWND WinClose(char *aTitle = "", char *aText = "", int aTimeToWaitForClose = DEFAULT_WINCLOSE_WAIT
	, char *aExcludeTitle = "", char *aExcludeText = "", bool aKillIfHung = false);

HWND WinActive(char *aTitle, char *aText = "", char *aExcludeTitle = "", char *aExcludeText = "");

HWND WinExist(char *aTitle, char *aText = "", char *aExcludeTitle = "", char *aExcludeText = ""
	, bool aFindLastMatch = false, bool aUpdateLastUsed = false
	, HWND aAlreadyVisited[] = NULL, int aAlreadyVisitedCount = 0);

BOOL CALLBACK EnumParentFind(HWND hwnd, LPARAM lParam);
BOOL CALLBACK EnumChildFind(HWND hwnd, LPARAM lParam);


// Use a fairly long default for aCheckInterval since the contents of this function's loops
// might be somewhat high in overhead (especially SendMessageTimeout):
#define SB_DEFAULT_CHECK_INTERVAL 50
ResultType StatusBarUtil(Var *aOutputVar, HWND aControlWindow, int aPartNumber = 1
	, char *aTextToWaitFor = "", int aWaitTime = -1, int aCheckInterval = SB_DEFAULT_CHECK_INTERVAL);
HWND ControlExist(HWND aParentWindow, char *aClassNameAndNum = NULL);
BOOL CALLBACK EnumControlFind(HWND aWnd, LPARAM lParam);

#define MSGBOX_NORMAL (MB_OK | MB_SETFOREGROUND)
#define MSGBOX_TEXT_SIZE (1024 * 8)
#define DIALOG_TITLE_SIZE 1024
int MsgBox(int aValue);
int MsgBox(char *aText = "", UINT uType = MSGBOX_NORMAL, char *aTitle = NULL, UINT aTimeout = 0);
HWND WinActivateOurTopDialog();
BOOL CALLBACK EnumDialog(HWND hwnd, LPARAM lParam);

HWND WindowOwnsOthers(HWND aWnd);
BOOL CALLBACK EnumParentFindOwned(HWND aWnd, LPARAM lParam);
HWND GetTopChild(HWND aParent);
bool IsWindowHung(HWND aWnd);

// Defaults to a low timeout because a window may have hundreds of controls, and if the window
// is hung, each control might result in a delay of size aTimeout during an EnumWindows.
// It shouldn't need much time anyway since the moment the call to SendMessageTimeout()
// is made, our thread is suspended and the target thread's WindowProc called directly.
// In addition:
// Whenever using SendMessageTimeout(), our app will be unresponsive until
// the call returns, since our message loop isn't running.  In addition,
// if the keyboard or mouse hook is installed, the events will lag during
// this call.  So keep the timeout value fairly short.  UPDATE: Need a longer
// timeout because otherwise searching will be inconsistent / unreliable for the
// slow Title Match method, since some apps are lazy about keeping their
// message pumps running, such as during long disk I/O operations, and thus
// may sometimes (randomly) take a long time to respond to the WM_GETTEXT message.
// 5000 seems about the largest value that should ever be needed since this is what
// Windows uses as the cutoff for determining if a window has become "unresponsive":
int GetWindowTextTimeout(HWND aWnd, char *aBuf = NULL, int aBufSize = 0, UINT aTimeout = 5000);
void SetForegroundLockTimeout();


inline int GetWindowTextByTitleMatchMode(HWND aWnd, char *aBuf = NULL, int aBufSize = 0)
{
	// Due to potential key and mouse lag caused by GetWindowTextTimeout() preventing us
	// from pumping messages for up to several seconds at a time (only if the user has specified
	// that the hook(s) be installed), it might be best to always attempt GetWindowText() prior
	// to the GetWindowTextTimeout().  Only if GetWindowText() gets 0 length would we try the other
	// method (and of course, don't bother using GetWindowTextTimeout() at all if "fast" mode is in
	// effect).  The problem with this is that many controls always return 0 length regardless of
	// which method is used, so this would slow things down a little (but not too badly since
	// GetWindowText() is so much faster than GetWindowTextTimeout()).  Another potential problem
	// is that some controls may return less text, or different text, when used with the fast mode
	// vs. the slow mode (unverified).  So it seems best NOT to do this and stick with the simple
	// approach below.
	if (g.TitleFindFast)
		return GetWindowText(aWnd, aBuf, aBufSize);
	else
		// We're using the slower method that is able to get text from more types of
		// controls (e.g. large edit controls).
		return GetWindowTextTimeout(aWnd, aBuf, aBufSize);
}



inline HWND HasMatchingChild(HWND aWnd, char *aText, char *aExcludeText)
// Caller must have verified that all params are not NULL.
{
	if (!*aText && !*aExcludeText)
		// This condition is defined as always being a match, so return the parent window
		// itself to indicate success:
		return aWnd;
	WindowInfoPackage wip;
	strlcpy(wip.text, aText, sizeof(wip.text));
	strlcpy(wip.exclude_text, aExcludeText, sizeof(wip.exclude_text));
	EnumChildWindows(aWnd, EnumChildFind, (LPARAM)&wip);
	return wip.child_hwnd; // Returns non-NULL on success.
}

inline bool IsTextMatch(char *aHaystack, char *aNeedle, char *aExcludeText = ""
	, bool aFindAnywhere = g.TitleFindAnywhere)
// To help performance, it's the caller's responsibility to ensure that all params are not NULL.
// Use the AutoIt2 convention (same in AutoIt3?) of making searches for window titles
// and text case sensitive: "N.B. Windows titles and text are CASE SENSITIVE!"
{
	if (aFindAnywhere)
		return (!*aNeedle || strstr(aHaystack, aNeedle)) // Either one of these makes half a match.
			&& (!*aExcludeText || !strstr(aHaystack, aExcludeText));  // And this is the other half.
	// Otherwise, search the leading part only:
	return (!*aNeedle || !strncmp(aHaystack, aNeedle, strlen(aNeedle)))
		&& (!*aExcludeText || strncmp(aHaystack, aExcludeText, strlen(aExcludeText)));
}

#endif
