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
#define USE_FOREGROUND_WINDOW(title, text, exclude_title, exclude_text)\
	((*title == 'A' || *title == 'a') && !*(title + 1) && !*text && !*exclude_title && !*exclude_text)
#define SET_TARGET_TO_ALLOWABLE_FOREGROUND \
{\
	if (target_window = GetForegroundWindow())\
		if (!g.DetectHiddenWindows && !IsWindowVisible(target_window))\
			target_window = NULL;\
}
#define IF_USE_FOREGROUND_WINDOW(title, text, exclude_title, exclude_text)\
if (USE_FOREGROUND_WINDOW(title, text, exclude_title, exclude_text))\
{\
	SET_TARGET_TO_ALLOWABLE_FOREGROUND\
}

#define SEARCH_PHRASE_SIZE 1024
// Info from AutoIt3 source: GetWindowText fails under 95 if >65535, WM_GETTEXT randomly fails if > 32767.
// My: And since 32767 is what AutoIt3 passes to the API functions as the size (not the length, i.e.
// it can only store 32766 if room is left for the zero terminator) we'll use that for the size too.
// Note: MSDN says (for functions like GetWindowText): "Specifies the maximum number of characters to
// copy to the buffer, including the NULL character. If the text exceeds this limit, it is truncated."
#define WINDOW_TEXT_SIZE 32767
#define WINDOW_CLASS_SIZE 1024  // Haven't found anything that documents how long one can be, so use this.
#define AHK_CLASS_FLAG "ahk_class"
#define AHK_CLASS_FLAG_LENGTH 9  // The length of the above string.
#define AHK_ID_FLAG "ahk_id"
#define AHK_ID_FLAG_LENGTH 6  // The length of the above string.

struct WindowInfoPackage // A simple struct to help with EnumWindows().
{
	char title[SEARCH_PHRASE_SIZE];
	char text[SEARCH_PHRASE_SIZE];
	char exclude_title[SEARCH_PHRASE_SIZE];
	char exclude_text[SEARCH_PHRASE_SIZE];
	// Whether to keep searching even after a match is found, so that last one is found.
	bool find_last_match;
	int match_count;
	HWND parent_hwnd, child_hwnd; // Returned to the caller, but the caller should initialize it to NULL beforehand.
	HWND *already_visited; // Array of HWNDs to exclude from consideration.
	int already_visited_count;
	WindowSpec *win_spec; // Linked list.
	Var *array_start; // Used by WinGet() to fetch an array of matching HWNDs.
	WindowInfoPackage::WindowInfoPackage()
		: find_last_match(false) // default
		, match_count(0)
		, parent_hwnd(NULL), child_hwnd(NULL), already_visited(NULL), already_visited_count(0)
		, win_spec(NULL), array_start(NULL)
	{
		// Can't use initializer list for these:
		*title = *text = *exclude_title = *exclude_text = '\0';
	}
};

struct control_list_type
{
	// For something this simple, a macro is probably a lot less overhead that making this struct
	// non-POD and giving it a constructor:
	#define CL_INIT_CONTROL_LIST(cl) \
		cl.is_first_iteration = true;\
		cl.total_classes = 0;\
		cl.total_length = 0;\
		cl.buf_free_spot = cl.class_buf; // Points to the next available/writable place in the buf.
	bool is_first_iteration;  // Must be initialized to true by caller.
	int total_classes;        // Must be initialized to 0.
	VarSizeType total_length; // Must be initialized to 0.
	VarSizeType capacity;     // Must be initialized to size of the below buffer.
	char *target_buf;         // Caller sets it to NULL if only the total_length is to be retrieved.
	#define CL_CLASS_BUF_SIZE (32 * 1024) // Even if class names average 50 chars long, this supports 655 of them.
	char class_buf[CL_CLASS_BUF_SIZE];
	char *buf_free_spot;      // Must be initialized to point to the beginning of class_buf.
	#define CL_MAX_CLASSES 500  // The number of distinct class names that can be supported in a single window.
	char *class_name[CL_MAX_CLASSES]; // Array of distinct class names, stored consecutively in class_buf.
	int class_count[CL_MAX_CLASSES];  // The quantity found for each of the above classes.
};

struct MonitorInfoPackage // A simple struct to help with EnumDisplayMonitors().
{
	int count;
	#define COUNT_ALL_MONITORS INT_MIN  // A special value that can be assigned to the below.
	int monitor_number_to_find;  // If this is left as zero, it will find the primary monitor by default.
	MONITORINFOEX monitor_info_ex;
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

struct class_and_hwnd_type
{
	char *class_name;
	bool is_found;
	int class_count;
	HWND hwnd;
};

struct point_and_hwnd_type
{
	POINT pt;
	RECT rect_found;
	HWND hwnd_found;
	double distance;
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

HWND WinActive(char *aTitle, char *aText = "", char *aExcludeTitle = "", char *aExcludeText = ""
	, bool aUpdateLastUsed = false);

HWND WinExist(char *aTitle, char *aText = "", char *aExcludeTitle = "", char *aExcludeText = ""
	, bool aFindLastMatch = false, bool aUpdateLastUsed = false
	, HWND aAlreadyVisited[] = NULL, int aAlreadyVisitedCount = 0, bool aReturnTheCount = false);

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
int MsgBox(char *aText = "", UINT uType = MSGBOX_NORMAL, char *aTitle = NULL, double aTimeout = 0);
HWND FindOurTopDialog();
BOOL CALLBACK EnumDialog(HWND hwnd, LPARAM lParam);

HWND WindowOwnsOthers(HWND aWnd);
BOOL CALLBACK EnumParentFindOwned(HWND aWnd, LPARAM lParam);
HWND GetNonChildParent(HWND aWnd);
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

inline bool IsTextMatch(char *aHaystack, char *aNeedle, TitleMatchModes aTitleMatchMode = g.TitleMatchMode)
// To help performance, it's the caller's responsibility to ensure that all params are not NULL.
// Use the AutoIt2 convention (same in AutoIt3?) of making searches for window titles
// and text case sensitive: "N.B. Windows titles and text are CASE SENSITIVE!"
{
	if (aTitleMatchMode == FIND_ANYWHERE)
		return !*aNeedle || strstr(aHaystack, aNeedle);
	else if (aTitleMatchMode == FIND_IN_LEADING_PART)
		return !*aNeedle || !strncmp(aHaystack, aNeedle, strlen(aNeedle));
	else // Exact match.
		return !*aNeedle || !strcmp(aHaystack, aNeedle);
}

inline bool IsTitleMatch(HWND aWnd, char *aHaystack, char *aNeedle, char *aExcludeTitle)
// To help performance, it's the caller's responsibility to ensure that all params are not NULL.
// Use the AutoIt2 convention (same in AutoIt3?) of making searches for window titles
// and text case sensitive.
{
	if (strnicmp(aNeedle, AHK_CLASS_FLAG, AHK_CLASS_FLAG_LENGTH)) // aNeedle doesn't specify a class name.
	{
		if (g.TitleMatchMode == FIND_ANYWHERE)
			return (!*aNeedle || strstr(aHaystack, aNeedle)) // Either one of these makes half a match.
				&& (!*aExcludeTitle || !strstr(aHaystack, aExcludeTitle));  // And this is the other half.
		else if (g.TitleMatchMode == FIND_IN_LEADING_PART)
			return (!*aNeedle || !strncmp(aHaystack, aNeedle, strlen(aNeedle)))
				&& (!*aExcludeTitle || strncmp(aHaystack, aExcludeTitle, strlen(aExcludeTitle)));
		else // Exact match.
			return (!*aNeedle || !strcmp(aHaystack, aNeedle))
				&& (!*aExcludeTitle || strcmp(aHaystack, aExcludeTitle));
	}
	// Otherwise, aNeedle specifies a class name rather than a window title.
	aNeedle = omit_leading_whitespace(aNeedle + AHK_CLASS_FLAG_LENGTH);
	char fore_class[WINDOW_CLASS_SIZE];
	if (!GetClassName(aWnd, fore_class, WINDOW_CLASS_SIZE - 1)) // Assume its not a match.
		return false;
	// To be a match, the class names must match exactly (case sensitive).  This seems best to
	// avoid problems with ambiguity, since some apps might use very short class names that
	// overlap with more "official" classnames, or vice versa.  User can always define a Window
	// Group to operate upon more than one class simultaneously.
	if (strcmp(fore_class, aNeedle))
		return false;
	// The other requirement for a match is that ExcludeTitle not be found in aHaystack.
	if (!*aExcludeTitle)
		return true;
	if (g.TitleMatchMode == FIND_ANYWHERE)
		return !strstr(aHaystack, aExcludeTitle);
	else if (g.TitleMatchMode == FIND_IN_LEADING_PART)
		return strncmp(aHaystack, aExcludeTitle, strlen(aExcludeTitle));
	else // Exact match.
		return strcmp(aHaystack, aExcludeTitle);
}



////////////////////
// PROCESS ROUTINES
////////////////////

DWORD ProcessExist9x2000(char *aProcess, char *aProcessName);
DWORD ProcessExistNT4(char *aProcess, char *aProcessName);

inline DWORD ProcessExist(char *aProcess, char *aProcessName = NULL)
{
	return g_os.IsWinNT4() ? ProcessExistNT4(aProcess, aProcessName)
		: ProcessExist9x2000(aProcess, aProcessName);
}

#endif
