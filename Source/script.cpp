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
#include <winsock.h>  // for WSADATA.  This also requires wsock32.lib to be linked in.
#include "script.h"
#include "globaldata.h" // for a lot of things
#include "util.h" // for strlcpy() etc.
#include "mt19937ar-cok.h" // for random number generator
#include "window.h" // for a lot of things
#include "application.h" // for MsgSleep()
#include "qmath.h" // For ExpandExpression()

// Globals that are for only this module:
#define MAX_COMMENT_FLAG_LENGTH 15
static char g_CommentFlag[MAX_COMMENT_FLAG_LENGTH + 1] = ";";
static size_t g_CommentFlagLength = strlen(g_CommentFlag); // pre-calculated for performance

// General note about the methods in here:
// Want to be able to support multiple simultaneous points of execution
// because more than one subroutine can be executing simultaneously
// (well, more precisely, there can be more than one script subroutine
// that's in a "currently running" state, even though all such subroutines,
// except for the most recent one, are suspended.  So keep this in mind when
// using things such as static data members or static local variables.


Script::Script()
	: mFirstLine(NULL), mLastLine(NULL), mCurrLine(NULL)
	, mLoopFile(NULL), mLoopRegItem(NULL), mLoopReadFile(NULL), mLoopField(NULL), mLoopIteration(0)
	, mThisHotkeyName(""), mPriorHotkeyName(""), mThisHotkeyStartTime(0), mPriorHotkeyStartTime(0)
	, mEndChar(0), mThisHotkeyModifiersLR(0)
	, mOnExitLabel(NULL), mExitReason(EXIT_NONE)
	, mFirstLabel(NULL), mLastLabel(NULL)
	, mFirstTimer(NULL), mLastTimer(NULL), mTimerEnabledCount(0), mTimerCount(0)
	, mFirstMenu(NULL), mLastMenu(NULL), mMenuCount(0)
	, mFirstVar(NULL), mLastVar(NULL)
	, mLineCount(0), mLabelCount(0), mVarCount(0), mGroupCount(0)
#ifdef AUTOHOTKEYSC
	, mCompiledHasCustomIcon(false)
#endif;
	, mCurrFileNumber(0), mCurrLineNumber(0), mNoHotkeyLabels(true), mMenuUseErrorLevel(false)
	, mFileSpec(""), mFileDir(""), mFileName(""), mOurEXE(""), mOurEXEDir(""), mMainWindowTitle("")
	, mIsReadyToExecute(false), AutoExecSectionIsRunning(false)
	, mIsRestart(false), mIsAutoIt2(false), mErrorStdOut(false)
	, mLinesExecutedThisCycle(0), mUninterruptedLineCountMax(1000), mUninterruptibleTime(15)
	, mRunAsUser(NULL), mRunAsPass(NULL), mRunAsDomain(NULL)
	, mCustomIcon(NULL) // Normally NULL unless there's a custom tray icon loaded dynamically.
	, mCustomIconFile(NULL), mTrayIconTip(NULL) // Allocated on first use.
	, mCustomIconNumber(0)
{
	// v1.0.25: mLastScriptRest and mLastPeekTime are now initialized right before the auto-exec
	// section of the script is launched, which avoids an initial Sleep(10) in ExecUntil
	// that would otherwise occur.
	*mThisMenuItemName = *mThisMenuName = '\0';
	ZeroMemory(&mNIC, sizeof(mNIC));  // Constructor initializes this, to be safe.
	mNIC.hWnd = NULL;  // Set this as an indicator that it tray icon is not installed.

	// Lastly (after the above have been initialized), anything that can fail:
	if (   !(mTrayMenu = AddMenu("Tray"))   ) // realistically never happens
	{
		ScriptError("No tray mem");
		ExitApp(EXIT_CRITICAL);
	}
	else
		mTrayMenu->mIncludeStandardItems = true;

#ifdef _DEBUG
	if (ID_FILE_EXIT < ID_MAIN_FIRST) // Not a very thorough check.
		ScriptError("DEBUG: ID_FILE_EXIT is too large (conflicts with IDs reserved via ID_USER_FIRST).");
	if (MAX_CONTROLS_PER_GUI > 1000)
		ScriptError("DEBUG: MAX_CONTROLS_PER_GUI is too large (conflicts with IDs reserved via ID_USER_FIRST).");
	int LargestMaxParams, i, j;
	ActionTypeType *np;
	// Find the Largest value of MaxParams used by any command and make sure it
	// isn't something larger than expected by the parsing routines:
	for (LargestMaxParams = i = 0; i < g_ActionCount; ++i)
	{
		if (g_act[i].MaxParams > LargestMaxParams)
			LargestMaxParams = g_act[i].MaxParams;
		// This next part has been tested and it does work, but only if one of the arrays
		// contains exactly MAX_NUMERIC_PARAMS number of elements and isn't zero terminated.
		// Relies on short-circuit boolean order:
		for (np = g_act[i].NumericParams, j = 0; j < MAX_NUMERIC_PARAMS && *np; ++j, ++np);
		if (j >= MAX_NUMERIC_PARAMS)
		{
			ScriptError("DEBUG: At least one command has a NumericParams array that isn't zero-terminated."
				"  This would result in reading beyond the bounds of the array.");
			return;
		}
	}
	if (LargestMaxParams > MAX_ARGS)
		ScriptError("DEBUG: At least one command supports more arguments than allowed.");
	if (sizeof(ActionTypeType) == 1 && g_ActionCount > 256)
		ScriptError("DEBUG: Since there are now more than 256 Action Types, the ActionTypeType"
			" typedef must be changed.");
#endif
}



Script::~Script()
{
	// MSDN: "Before terminating, an application must call the UnhookWindowsHookEx function to free
	// system resources associated with the hook."
	RemoveAllHooks();
	if (mNIC.hWnd) // Tray icon is installed.
		Shell_NotifyIcon(NIM_DELETE, &mNIC); // Remove it.
	// Only after the above do we get rid of the icon:
	if (mCustomIcon) // A custom icon was installed
		DestroyIcon(mCustomIcon);
	// Destroy any Progress/SplashImage windows that haven't already been destroyed.  This is necessary
	// because sometimes these windows aren't owned by the main window:
	int i;
	for (i = 0; i < MAX_PROGRESS_WINDOWS; ++i)
	{
		if (g_Progress[i].hwnd && IsWindow(g_Progress[i].hwnd))
			DestroyWindow(g_Progress[i].hwnd);
		if (g_Progress[i].hfont1) // Destroy font only after destroying the window that uses it.
			DeleteObject(g_Progress[i].hfont1);
		if (g_Progress[i].hfont2) // Destroy font only after destroying the window that uses it.
			DeleteObject(g_Progress[i].hfont2);
		if (g_Progress[i].hbrush)
			DeleteObject(g_Progress[i].hbrush);
	}
	for (i = 0; i < MAX_SPLASHIMAGE_WINDOWS; ++i)
	{
		if (g_SplashImage[i].pic)
			g_SplashImage[i].pic->Release();
		if (g_SplashImage[i].hwnd && IsWindow(g_SplashImage[i].hwnd))
			DestroyWindow(g_SplashImage[i].hwnd);
		if (g_SplashImage[i].hfont1) // Destroy font only after destroying the window that uses it.
			DeleteObject(g_SplashImage[i].hfont1);
		if (g_SplashImage[i].hfont2) // Destroy font only after destroying the window that uses it.
			DeleteObject(g_SplashImage[i].hfont2);
		if (g_SplashImage[i].hbrush)
			DeleteObject(g_SplashImage[i].hbrush);
	}

	// It is safer/easier to destroy the GUI windows prior to the menus (especially the menu bars).
	// This is because one GUI window might get destroyed and take with it a menu bar that is still
	// in use by an existing GUI window.  GuiType::Destroy() adheres to this philosophy by detaching
	// its menu bar prior to destroying its window:
	for (i = 0; i < MAX_GUI_WINDOWS; ++i)
		GuiType::Destroy(i); // Static method to avoid problems with object destroying itself.
	for (i = 0; i < GuiType::sFontCount; ++i) // Now that GUI windows are gone, delete all GUI fonts.
		if (GuiType::sFont[i].hfont)
			DeleteObject(GuiType::sFont[i].hfont);
	// The above might attempt to delete an HFONT from GetStockObject(DEFAULT_GUI_FONT), etc.
	// But that should be harmless:
	// MSDN: "It is not necessary (but it is not harmful) to delete stock objects by calling DeleteObject."

	// Since they're not associated with a window, we must free the resources for all popup menus.
	// Update: Even if a menu is being used as a GUI window's menu bar, see note above for why menu
	// destruction is done AFTER the GUI windows are destroyed:
	UserMenu *menu_to_delete;
	for (UserMenu *m = mFirstMenu; m;)
	{
		menu_to_delete = m;
		m = m->mNextMenu;
		ScriptDeleteMenu(menu_to_delete);
		// Above call should not return FAIL, since the only way FAIL can realistically happen is
		// when a GUI window is still using the menu as its menu bar.  But all GUI windows are gone now.
	}

	// Since tooltip windows are unowned, they should be destroyed to avoid resource leak:
	for (i = 0; i < MAX_TOOLTIPS; ++i)
		if (g_hWndToolTip[i] && IsWindow(g_hWndToolTip[i]))
			DestroyWindow(g_hWndToolTip[i]);

	if (g_hFontSplash) // The splash window itself should auto-destroyed, since it's owned by main.
		DeleteObject(g_hFontSplash);

	// Close any open sound item to prevent hang-on-exit in certain operating systems or conditions.
	// If there's any chance that a sound was played and not closed out, or that it is still playing,
	// this check is done.  Otherwise, the check is avoided since it might be a high overhead call,
	// especially if the sound subsystem part of the OS is currently swapped out or something:
	if (g_SoundWasPlayed)
	{
		char buf[MAX_PATH * 2];
		mciSendString("status " SOUNDPLAY_ALIAS " mode", buf, sizeof(buf), NULL);
		if (*buf) // "playing" or "stopped"
			mciSendString("close " SOUNDPLAY_ALIAS, NULL, 0, NULL);
	}
#ifdef ENABLE_KEY_HISTORY_FILE
	KeyHistoryToFile();  // Close the KeyHistory file if it's open.
#endif
}



ResultType Script::Init(char *aScriptFilename, bool aIsRestart)
// Returns OK or FAIL.
{
	mIsRestart = aIsRestart;
	if (!aScriptFilename || !*aScriptFilename) return FAIL;
	char buf[2048]; // Just make sure we have plenty of room to do things with.
	char *filename_marker;
	// In case the script is a relative filespec (relative to current working dir):
	if (!GetFullPathName(aScriptFilename, sizeof(buf), buf, &filename_marker))
	{
		MsgBox("GetFullPathName"); // Short msg since so rare.
		return FAIL;
	}
	// Using the correct case not only makes it look better in title bar & tray tool tip,
	// it also helps with the detection of "this script already running" since otherwise
	// it might not find the dupe if the same script name is launched with different
	// lowercase/uppercase letters:
	ConvertFilespecToCorrectCase(buf);
	// In case the above changed the length, e.g. due to expansion of 8.3 filename:
	if (   !(filename_marker = strrchr(buf, '\\'))   )
		filename_marker = buf;
	else
		++filename_marker;
	if (   !(mFileSpec = SimpleHeap::Malloc(buf))   )  // The full spec is stored for convenience.
		return FAIL;  // It already displayed the error for us.
	filename_marker[-1] = '\0'; // Terminate buf in this position to divide the string.
	size_t filename_length = strlen(filename_marker);
	if (   mIsAutoIt2 = (filename_length >= 4 && !stricmp(filename_marker + filename_length - 4, EXT_AUTOIT2))   )
	{
		// Set the old/AutoIt2 defaults for maximum safety and compatibilility.
		// Standalone EXEs (compiled scripts) are always considered to be non-AutoIt2 (otherwise,
		// the user should probably be using the AutoIt2 compiler).
		g_AllowSameLineComments = false;
		g_EscapeChar = '\\';
		g.TitleFindFast = true; // In case the normal default is false.
		g.DetectHiddenText = false;
		// Make the mouse fast like AutoIt2, but not quite insta-move.  2 is expected to be more
		// reliable than 1 since the AutoIt author said that values less than 2 might cause the
		// drag to fail (perhaps just for specific apps, such as games):
		g.DefaultMouseSpeed = 2;
		g.KeyDelay = 20;
		g.WinDelay = 500;
		g.LinesPerCycle = 1;
		g.IntervalBeforeRest = -1;  // i.e. this method is disabled by default for AutoIt2 scripts.
		// Reduce max params so that any non escaped delimiters the user may be using literally
		// in "window text" will still be considered literal, rather than as delimiters for
		// args that are not supported by AutoIt2, such as exclude-title, exclude-text, MsgBox
		// timeout, etc.  Note: Don't need to change IfWinExist and such because those already
		// have special handling to recognize whether exclude-title is really a valid command
		// instead (e.g. IfWinExist, title, text, Gosub, something).

		// NOTE: DO NOT ADD the IfWin command series to this section, since there is special handling
		// for parsing those commands to figure out whether they're being used in the old AutoIt2
		// style or the new Exclude Title/Text mode.

		g_act[ACT_FILESELECTFILE].MaxParams -= 2;
		g_act[ACT_FILEREMOVEDIR].MaxParams -= 1;
		g_act[ACT_MSGBOX].MaxParams -= 1;
		g_act[ACT_INIREAD].MaxParams -= 1;
		g_act[ACT_STRINGREPLACE].MaxParams -= 1;
		g_act[ACT_STRINGGETPOS].MaxParams -= 1;
		g_act[ACT_WINCLOSE].MaxParams -= 3;  // -3 for these two, -2 for the others.
		g_act[ACT_WINKILL].MaxParams -= 3;
		g_act[ACT_WINACTIVATE].MaxParams -= 2;
		g_act[ACT_WINMINIMIZE].MaxParams -= 2;
		g_act[ACT_WINMAXIMIZE].MaxParams -= 2;
		g_act[ACT_WINRESTORE].MaxParams -= 2;
		g_act[ACT_WINHIDE].MaxParams -= 2;
		g_act[ACT_WINSHOW].MaxParams -= 2;
		g_act[ACT_WINSETTITLE].MaxParams -= 2;
		g_act[ACT_WINGETTITLE].MaxParams -= 2;
	}
	if (   !(mFileDir = SimpleHeap::Malloc(buf))   )
		return FAIL;  // It already displayed the error for us.
	if (   !(mFileName = SimpleHeap::Malloc(filename_marker))   )
		return FAIL;  // It already displayed the error for us.
#ifdef AUTOHOTKEYSC
	// Omit AutoHotkey from the window title, like AutoIt3 does for its compiled scripts.
	// One reason for this is to reduce backlash if evil-doers create viruses and such
	// with the program:
	snprintf(buf, sizeof(buf), "%s\\%s", mFileDir, mFileName);
#else
	snprintf(buf, sizeof(buf), "%s\\%s - %s", mFileDir, mFileName, NAME_PV);
#endif
	if (   !(mMainWindowTitle = SimpleHeap::Malloc(buf))   )
		return FAIL;  // It already displayed the error for us.

	// It may be better to get the module name this way rather than reading it from the registry
	// (though it might be more proper to parse it out of the command line args or something),
	// in case the user has moved it to a folder other than the install folder, hasn't installed it,
	// or has renamed the EXE file itself.  Also, enclose the full filespec of the module in double
	// quotes since that's how callers usually want it because ActionExec() currently needs it that way:
	*buf = '"';
	if (GetModuleFileName(NULL, buf + 1, sizeof(buf) - 2)) // -2 to leave room for the enclosing double quotes.
	{
		size_t buf_length = strlen(buf);
		buf[buf_length++] = '"';
		buf[buf_length] = '\0';
		if (   !(mOurEXE = SimpleHeap::Malloc(buf))   )
			return FAIL;  // It already displayed the error for us.
		else
		{
			char *last_backslash = strrchr(buf, '\\');
			if (!last_backslash) // probably can't happen due to the nature of GetModuleFileName().
				mOurEXEDir = "";
			last_backslash[1] = '\0'; // i.e. keep the trailing backslash for convenience.
			if (   !(mOurEXEDir = SimpleHeap::Malloc(buf + 1))   ) // +1 to omit the leading double-quote.
				return FAIL;  // It already displayed the error for us.
		}
	}
	return OK;
}

	

ResultType Script::CreateWindows()
// Returns OK or FAIL.
{
	if (!mMainWindowTitle || !*mMainWindowTitle) return FAIL;  // Init() must be called before this function.
	// Register a window class for the main window:
	HICON hIcon = LoadIcon(g_hInstance, MAKEINTRESOURCE(IDI_MAIN)); // LoadIcon((HINSTANCE) NULL, IDI_APPLICATION)
	WNDCLASSEX wc = {0};
	wc.cbSize = sizeof(wc);
	wc.lpszClassName = WINDOW_CLASS_MAIN;
	wc.hInstance = g_hInstance;
	wc.lpfnWndProc = MainWindowProc;
	// Provided from some example code:
	wc.style = 0;  // CS_HREDRAW | CS_VREDRAW
	wc.cbClsExtra = 0;
	wc.cbWndExtra = 0;
	wc.hIcon = hIcon;
	wc.hIconSm = hIcon;
	wc.hCursor = LoadCursor((HINSTANCE) NULL, IDC_ARROW);
	wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);  // Needed for ProgressBar. Old: (HBRUSH)GetStockObject(WHITE_BRUSH);
	wc.lpszMenuName = MAKEINTRESOURCE(IDR_MENU_MAIN); // NULL; // "MainMenu";
	if (!RegisterClassEx(&wc))
	{
		MsgBox("RegisterClass() #1 failed.");
		return FAIL;
	}

	// Register a second class for the splash window.  The only difference is that
	// it doesn't have the menu bar:
	wc.lpszClassName = WINDOW_CLASS_SPLASH;
	wc.lpszMenuName = NULL;
	if (!RegisterClassEx(&wc))
	{
		MsgBox("RegisterClass() #2 failed.");
		return FAIL;
	}

	char class_name[64];
	HWND fore_win = GetForegroundWindow();
	bool do_minimize = !fore_win || (GetClassName(fore_win, class_name, sizeof(class_name))
		&& !stricmp(class_name, "Shell_TrayWnd")); // Shell_TrayWnd is the taskbar's class on Win98/XP and probably the others too.

	// Note: the title below must be constructed the same was as is done by our
	// WinMain() (so that we can detect whether this script is already running)
	// which is why it's standardized in g_script.mMainWindowTitle.
	// Create the main window.  Prevent momentary disruption of Start Menu, which
	// some users understandably don't like, by omitting the taskbar button temporarily.
	// This is done because testing shows that minimizing the window further below, even
	// though the window is hidden, would otherwise briefly show the taskbar button (or
	// at least redraw the taskbar).  Sometimes this isn't noticeable, but other times
	// (such as when the system is under heavy load) a user reported that it is quite
	// noticeable:
	if (   !(g_hWnd = CreateWindowEx(do_minimize ? WS_EX_TOOLWINDOW : 0
		, WINDOW_CLASS_MAIN
		, mMainWindowTitle
		, WS_OVERLAPPEDWINDOW // Style.  Alt: WS_POPUP or maybe 0.
		, CW_USEDEFAULT // xpos
		, CW_USEDEFAULT // ypos
		, CW_USEDEFAULT // width
		, CW_USEDEFAULT // height
		, NULL // parent window
		, NULL // Identifies a menu, or specifies a child-window identifier depending on the window style
		, g_hInstance // passed into WinMain
		, NULL))   ) // lpParam
	{
		MsgBox("CreateWindow"); // Short msg since so rare.
		return FAIL;
	}
#ifdef AUTOHOTKEYSC
	HMENU menu = GetMenu(g_hWnd);
	// Disable the Edit menu item, since it does nothing for a compiled script:
	EnableMenuItem(menu, ID_FILE_EDITSCRIPT, MF_DISABLED | MF_GRAYED);
	if (!g_AllowMainWindow)
	{
		EnableMenuItem(menu, ID_VIEW_KEYHISTORY, MF_DISABLED | MF_GRAYED);
		EnableMenuItem(menu, ID_VIEW_LINES, MF_DISABLED | MF_GRAYED);
		EnableMenuItem(menu, ID_VIEW_VARIABLES, MF_DISABLED | MF_GRAYED);
		EnableMenuItem(menu, ID_VIEW_HOTKEYS, MF_DISABLED | MF_GRAYED);
		// But leave the ID_VIEW_REFRESH menu item enabled because if the script contains a
		// command such as ListLines in it, Refresh can be validly used.
	}
#endif

	if (    !(g_hWndEdit = CreateWindow("edit", NULL, WS_CHILD | WS_VISIBLE | WS_BORDER
		| ES_LEFT | ES_MULTILINE | ES_READONLY | WS_VSCROLL // | WS_HSCROLL (saves space)
		, 0, 0, 0, 0, g_hWnd, (HMENU)1, g_hInstance, NULL))  )
	{
		MsgBox("CreateWindow"); // Short msg since so rare.
		return FAIL;
	}

	// Some of the MSDN docs mention that an app's very first call to ShowWindow() makes that
	// function operate in a special mode. Therefore, it seems best to get that first call out
	// of the way to avoid the possibility that the first-call behavior will cause problems with
	// our normal use of ShowWindow() below and other places.  Also, decided to ignore nCmdShow,
    // to avoid any momentary visual effects on startup:
	ShowWindow(g_hWnd, SW_HIDE);

	// Now that the first call to ShowWindow() is out of the way, minimize the main window so that
	// if the script is launched from the Start Menu (and perhaps other places such as the
	// Quick-launch toolbar), the window that was active before the Start Menu was displayed will
	// become active again.  But as of v1.0.25.09, this minimize is done more selectively to prevent
	// the launch of a script from knocking the user out of a full-screen game or other application
	// that would be disrupted by an SW_MINIMIZE:
	if (do_minimize)
	{
		ShowWindow(g_hWnd, SW_MINIMIZE);
		SetWindowLong(g_hWnd, GWL_EXSTYLE, 0); // Give the main window back its taskbar button.
	}

	g_hAccelTable = LoadAccelerators(g_hInstance, MAKEINTRESOURCE(IDR_ACCELERATOR1));

	if (g_NoTrayIcon)
		mNIC.hWnd = NULL;  // Set this as an indicator that tray icon is not installed.
	else
		// Even if the below fails, don't return FAIL in case the user is using a different shell
		// or something.  In other words, it is expected to fail under certain circumstances and
		// we want to tolerate that:
		CreateTrayIcon();
	return OK;
}



void Script::CreateTrayIcon()
// It is the caller's responsibility to ensure that the previous icon is first freed/destroyed
// before calling us to install a new one.  However, that is probably not needed if the Explorer
// crashed, since the memory used by the tray icon was probably destroyed along with it.
{
	ZeroMemory(&mNIC, sizeof(mNIC));  // To be safe.
	// Using NOTIFYICONDATA_V2_SIZE vs. sizeof(NOTIFYICONDATA) improves compatibility with Win9x maybe.
	// MSDN: "Using [NOTIFYICONDATA_V2_SIZE] for cbSize will allow your application to use NOTIFYICONDATA
	// with earlier Shell32.dll versions, although without the version 6.0 enhancements."
	// Update: Using V2 gives an compile error so trying V1.  Update: Trying sizeof(NOTIFYICONDATA)
	// for compatibility with VC++ 6.x.  This is also what AutoIt3 uses:
	mNIC.cbSize = sizeof(NOTIFYICONDATA);  // NOTIFYICONDATA_V1_SIZE
	mNIC.hWnd = g_hWnd;
	mNIC.uID = AHK_NOTIFYICON; // This is also used for the ID, see TRANSLATE_AHK_MSG for details.
	mNIC.uFlags = NIF_MESSAGE | NIF_TIP | NIF_ICON;
	mNIC.uCallbackMessage = AHK_NOTIFYICON;
#ifdef AUTOHOTKEYSC
	// i.e. don't override the user's custom icon:
	mNIC.hIcon = mCustomIcon ? mCustomIcon : LoadIcon(g_hInstance, MAKEINTRESOURCE(mCompiledHasCustomIcon ? IDI_MAIN : g_IconTray));
#else
	mNIC.hIcon = mCustomIcon ? mCustomIcon : LoadIcon(g_hInstance, MAKEINTRESOURCE(g_IconTray));
#endif
	UPDATE_TIP_FIELD
	// If we were called due to an Explorer crash, I don't think it's necessary to call
	// Shell_NotifyIcon() to remove the old tray icon because it was likely destroyed
	// along with Explorer.  So just add it unconditionally:
	if (!Shell_NotifyIcon(NIM_ADD, &mNIC))
		mNIC.hWnd = NULL;  // Set this as an indicator that tray icon is not installed.
}



void Script::UpdateTrayIcon(bool aForceUpdate)
{
	if (!mNIC.hWnd) // tray icon is not installed
		return;
	static bool icon_shows_paused = false;
	static bool icon_shows_suspended = false;
	if (!aForceUpdate && g.IsPaused == icon_shows_paused && g_IsSuspended == icon_shows_suspended)
		return; // it's already in the right state
	int icon;
	if (g.IsPaused && g_IsSuspended)
		icon = IDI_PAUSE_SUSPEND;
	else if (g.IsPaused)
		icon = IDI_PAUSE;
	else if (g_IsSuspended)
		icon = g_IconTraySuspend;
	else
#ifdef AUTOHOTKEYSC
		icon = mCompiledHasCustomIcon ? IDI_MAIN : g_IconTray;  // i.e. don't override the user's custom icon.
#else
		icon = g_IconTray;
#endif
	// Use the custom tray icon if the icon is normal (non-paused & non-suspended):
	mNIC.hIcon = (mCustomIcon && !g.IsPaused && !g_IsSuspended) ? mCustomIcon
		: LoadIcon(g_hInstance, MAKEINTRESOURCE(icon));
	if (Shell_NotifyIcon(NIM_MODIFY, &mNIC))
	{
		icon_shows_paused = g.IsPaused;
		icon_shows_suspended = g_IsSuspended;
	}
	// else do nothing, just leave it in the same state.
}



ResultType Script::AutoExecSection()
{
	if (!mIsReadyToExecute)
		return FAIL;
	if (mFirstLine != NULL)
	{
		// Choose a timeout that's a reasonable compromise between the following competing priorities:
		// 1) That we want hotkeys to be responsive as soon as possible after the program launches
		//    in case the user launches by pressing ENTER on a script, for example, and then immediately
		//    tries to use a hotkey.  In addition, we want any timed subroutines to start running ASAP
		//    because in rare cases the user might rely upon that happening.
		// 2) To support the case when the auto-execute section never finishes (such as when it contains
		//    an infinite loop to do background processing), yet we still want to allow the script
		//    to put custom defaults into effect globally (for things such as KeyDelay).
		// Obviously, the above approach has its flaws; there are ways to construct a script that would
		// result in unexpected behavior.  However, the combination of this approach with the fact that
		// the global defaults are updated *again* when/if the auto-execute section finally completes
		// raises the expectation of proper behavior to a very high level.  In any case, I'm not sure there
		// is any better approach that wouldn't break existing scripts or require a redesign of some kind.
		// If this method proves unreliable due to disk activity slowing the program down to a crawl during
		// the critical milliseconds after launch, one thing that might fix that is to have ExecUntil()
		// be forced to run a minimum of, say, 100 lines (if there are that many) before allowing the
		// timer expiration to have its effect.  But that's getting complicated and I'd rather not do it
		// unless someone actually reports that such a thing ever happens.  Still, to reduce the chance
		// of such a thing ever happening, it seems best to boost the timeout from 50 up to 100:
		SET_AUTOEXEC_TIMER(100);
		AutoExecSectionIsRunning = true;

		// v1.0.25: This is no done here, closer to the actual execution of the first line in the script,
		// to avoid an unnecessary Sleep(10) that would otherwise occur in ExecUntil:
		mLastScriptRest = mLastPeekTime = GetTickCount();

		++g_nThreads;
		ResultType result = mFirstLine->ExecUntil(UNTIL_RETURN);
		--g_nThreads;

		KILL_AUTOEXEC_TIMER  // This also does "g.AllowThisThreadToBeInterrupted = true"
		AutoExecSectionIsRunning = false;

		return result;
	}
	return OK;
}



ResultType Script::Edit()
{
#ifdef AUTOHOTKEYSC
	return OK; // Do nothing.
#else
	// This is here in case a compiled script ever uses the Edit command.  Since the "Edit This
	// Script" menu item is not available for compiled scripts, it can't be called from there.
	TitleMatchModes old_mode = g.TitleMatchMode;
	g.TitleMatchMode = FIND_ANYWHERE;
	HWND hwnd = WinExist(mFileName, "", mMainWindowTitle); // Exclude our own main window.
	g.TitleMatchMode = old_mode;
	if (hwnd)
	{
		char class_name[32];
		GetClassName(hwnd, class_name, sizeof(class_name));
		if (!strcmp(class_name, "#32770"))  // MessageBox(), InputBox(), or FileSelectFile() window.
			hwnd = NULL;  // Exclude it from consideration.
	}
	if (hwnd)  // File appears to already be open for editing, so use the current window.
		SetForegroundWindowEx(hwnd);
	else
	{
		char buf[MAX_PATH * 2];
		// Enclose in double quotes anything that might contain spaces since the CreateProcess()
		// method, which is attempted first, is more likely to succeed.  This is because it uses
		// the command line method of creating the process, with everything all lumped together:
		snprintf(buf, sizeof(buf), "\"%s\"", mFileSpec);
		if (!ActionExec("edit", buf, mFileDir, false))  // Since this didn't work, try notepad.
		{
			// Even though notepad properly handles filenames with spaces in them under WinXP,
			// even without double quotes around them, it seems safer and more correct to always
			// enclose the filename in double quotes for maximum compatibility with all OSes:
			if (!ActionExec("notepad.exe", buf, mFileDir, false))
				MsgBox("Could not open file for editing using the associated \"edit\" action or Notepad.");
		}
	}
	return OK;
#endif
}



ResultType Script::Reload(bool aDisplayErrors)
{
	// The new instance we're about to start will tell our process to stop, or it will display
	// a syntax error or some other error, in which case our process will still be running:
#ifdef AUTOHOTKEYSC
	// This is here in case a compiled script ever uses the Reload command.  Since the "Reload This
	// Script" menu item is not available for compiled scripts, it can't be called from there.
	return g_script.ActionExec(mOurEXE, "/restart", g_WorkingDirOrig, aDisplayErrors);
#else
	char arg_string[MAX_PATH + 512];
	snprintf(arg_string, sizeof(arg_string), "/restart \"%s\"", mFileSpec);
	return g_script.ActionExec(mOurEXE, arg_string, g_WorkingDirOrig, aDisplayErrors);
#endif
}



ResultType Script::ExitApp(ExitReasons aExitReason, char *aBuf, int aExitCode)
// Normal exit (if aBuf is NULL), or a way to exit immediately on error (which is mostly
// for times when it would be unsafe to call MsgBox() due to the possibility that it would
// make the situation even worse).
{
	mExitReason = aExitReason;
	bool terminate_afterward = aBuf && !*aBuf;
	if (aBuf && *aBuf)
	{
		char buf[1024];
		// No more than size-1 chars will be written and string will be terminated:
		snprintf(buf, sizeof(buf), "Critical Error: %s\n\n" WILL_EXIT, aBuf);
		// To avoid chance of more errors, don't use MsgBox():
		MessageBox(g_hWnd, buf, g_script.mFileSpec, MB_OK | MB_SETFOREGROUND | MB_APPLMODAL);
		TerminateApp(CRITICAL_ERROR); // Only after the above.
	}

	// Otherwise, it's not a critical error.  Note that currently, mOnExitLabel can only be
	// non-NULL if the script is in a runnable state (since registering an OnExit label requires
	// that a script command has executed to do it).  If this ever changes, the !mIsReadyToExecute
	// condition should be added to the below if statement:
	static bool sExitLabelIsRunning = false;
	if (!mOnExitLabel || sExitLabelIsRunning)  // || !mIsReadyToExecute
		// In the case of sExitLabelIsRunning == true:
		// There is another instance of this function beneath us on the stack.  Since we have
		// been called, this is a true exit condition and we exit immediately:
		TerminateApp(aExitCode);

	// Otherwise, the script contains the special RunOnExit label that we will run here instead
	// of exiting.  And since it does, we know that the script is in a ready-to-execute state
	// because that is the only way an OnExit label could have been defined in the first place.
	// Usually, the RunOnExit subroutine will contain an Exit or ExitApp statement
	// which results in a recursive call to this function, but this is not required (e.g. the
	// Exit subroutine could display an "Are you sure?" prompt, and if the user chooses "No",
	// the Exit sequence can be aborted by simply not calling ExitApp and letting the thread
	// we create below end normally).

	// Next, save the current state of the globals so that they can be restored just prior
	// to returning to our caller:
	strlcpy(g.ErrorLevel, g_ErrorLevel->Contents(), sizeof(g.ErrorLevel)); // Save caller's errorlevel.
	global_struct global_saved;
	CopyMemory(&global_saved, &g, sizeof(global_struct));
	CopyMemory(&g, &g_default, sizeof(global_struct)); // Mirrors the behavior of INIT_NEW_THREAD

	// Make it start fresh to avoid unnecessary delays due to SetBatchLines:
	g_script.mLinesExecutedThisCycle = 0;

	if (g_nFileDialogs) // See MsgSleep() for comments on this.
		SetCurrentDirectory(g_WorkingDir);

	// Use g.AllowThisThreadToBeInterrupted to forbid any hotkeys, timers, or user defined menu items
	// to interrupt.  This is mainly done for peace-of-mind (since possible interactions due to
	// interruptions have not been studied) and the fact that this most users would not want this
	// subroutine to be interruptible (it usually runs quickly anyway).  Another reason to make
	// it non-interruptible is that some OnExit subroutines might destruct things used by the
	// script's hotkeys/timers/menu items, and activating these items during the deconstruction
	// would not be safe.  Finally, if a logoff or shutdown is occurring, it seems best to prevent
	// timed subroutines from running -- which might take too much time and prevent the exit from
	// occurring in a timely fashion.  An option can be added via the FutureUse param to make it
	// interruptible if there is ever a demand for that.  UPDATE: g_AllowInterruption is now used
	// instead of g.AllowThisThreadToBeInterrupted for two reasons:
	// 1) It avoids the need to do "int mUninterruptedLineCountMax_prev = g_script.mUninterruptedLineCountMax;"
	//    (Disable this item so that ExecUntil() won't automatically make our new thread uninterruptible
	//    after it has executed a certain number of lines).
	// 2) If the thread we're interrupting is uninterruptible, the uinterruptible timer might be
	//    currently pending.  When it fires, it would make the OnExit subroutine interruptible
	//    rather than the underlying subroutine.  The above fixes the first part of that problem.
	//    The 2nd part is fixed by reinstating the timer when the uninterruptible thread is resumed.
	//    This special handling is only necessary here -- not in other places where new threads are
	//    created -- because OnExit is the only type of thread that can interrupt an uninterruptible
	//    thread.
	bool g_AllowInterruption_prev = g_AllowInterruption;  // Save current setting.
	g_AllowInterruption = false; // Mark the thread just created above as permanently uninterruptible (i.e. until it finishes and is destroyed).

	// This addresses the 2nd part of the problem described in the above large comment:
	bool uninterruptible_timer_was_pending = g_UninterruptibleTimerExists;

	// If the current quasi-thread is paused, the thread we're about to launch
	// will not be, so the icon needs to be checked:
	g_script.UpdateTrayIcon();

	sExitLabelIsRunning = true;
	++g_nThreads;  // i.e. This special thread does not obey g_MaxThreadsTotal.
	if (mOnExitLabel->mJumpToLine->ExecUntil(UNTIL_RETURN) == FAIL)
		// If the subroutine encounters a failure condition such as a runtime error, exit immediately.
		// Otherwise, there will be no way to exit the script if the subroutine fails on each attempt.
		TerminateApp(aExitCode);
	--g_nThreads;  // Since this instance of this function only had one thread in use at a time.
	sExitLabelIsRunning = false;  // In case the user wanted the thread to end normally (see above).

	if (terminate_afterward)
		TerminateApp(aExitCode);

	// Otherwise:
	RESUME_UNDERLYING_THREAD // For explanation, see comments where the macro is defined.
	g_AllowInterruption = g_AllowInterruption_prev;  // Restore original setting.
	if (uninterruptible_timer_was_pending)
		// Update: An alternative to the below would be to make the current thread interruptible
		// right before the OnExit thread interrupts it, and keep it that way.
		// Below macro recreates the timer if it doesn't already exists (i.e. if it fired during
		// the running of the OnExit subroutine).  Although such a firing would have had
		// no negative impact on the OnExit subroutine (since it's kept always-uninterruptible
		// via g_AllowInterruption), reinstate the timer so that it will make the thread
		// we're resuming interruptible.  The interval might not be exactly right -- since we
		// don't know when the thread started -- but it seems relatively unimportant to
		// worry about such accuracy given how rare and usually-inconsequential this whole
		// scenario is:
		SET_UNINTERRUPTIBLE_TIMER

	return OK;  // for caller convenience.
}



void Script::TerminateApp(int aExitCode)
// Note that g_script's destructor takes care of most other cleanup work, such as destroying
// tray icons, menus, and unowned windows such as ToolTip.
{
	// We call DestroyWindow() because MainWindowProc() has left that up to us.
	// DestroyWindow() will cause MainWindowProc() to immediately receive and process the
	// WM_DESTROY msg, which should in turn result in any child windows being destroyed
	// and other cleanup being done:
	if (IsWindow(g_hWnd)) // Adds peace of mind in case WM_DESTROY was already received in some unusual way.
	{
		g_DestroyWindowCalled = true;
		DestroyWindow(g_hWnd);
	}
	Hotkey::AllDestructAndExit(aExitCode);
}



LineNumberType Script::LoadFromFile()
// Returns the number of non-comment lines that were loaded, or LOADING_FAILED on error.
{
	mNoHotkeyLabels = true;  // Indicate that there are no hotkey labels, since we're (re)loading the entire file.
	mIsReadyToExecute = AutoExecSectionIsRunning = false;
	if (!mFileSpec || !*mFileSpec) return LOADING_FAILED;

#ifndef AUTOHOTKEYSC  // When not in stand-alone mode, read an external script file.
	DWORD attr = GetFileAttributes(mFileSpec);
	if (attr == MAXDWORD) // File does not exist or lacking the authorization to get its attributes.
	{
		char buf[MAX_PATH + 256];
		snprintf(buf, sizeof(buf), "The script file \"%s\" does not exist.  Create it now?", mFileSpec);
		int response = MsgBox(buf, MB_YESNO);
		if (response != IDYES)
			return 0;
		FILE *fp2 = fopen(mFileSpec, "a");
		if (!fp2)
		{
			MsgBox("Could not create file, perhaps because the current directory is read-only"
				" or has insufficient permissions.");
			return LOADING_FAILED;
		}
		fputs(
"; IMPORTANT INFO ABOUT GETTING STARTED: Lines that start with a\n"
"; semicolon, such as this one, are comments.  They are not executed.\n"
"\n"
"; This script is a .INI file because it is a special script that is\n"
"; automatically launched when you run the program directly. By contrast,\n"
"; text files that end in .ahk are associated with the program, which\n"
"; means that they can be launched simply by double-clicking them.\n"
"; You can have as many .ahk files as you want, located in any folder.\n"
"; You can also run more than one .ahk file simultaneously and each will\n"
"; get its own tray icon.\n"
"\n"
"; Please read the QUICK-START TUTORIAL near the top of the help file.\n"
"; It explains how to perform common automation tasks such as sending\n"
"; keystrokes and mouse clicks.  It also explains how to use hotkeys.\n"
"\n"
"; SAMPLE HOTKEYS: Below are two sample hotkeys.  The first is Win+Z and it\n"
"; launches a web site in the default browser.  The second is Control+Alt+N\n"
"; and it launches a new Notepad window (or activates an existing one).  To\n"
"; try out these hotkeys, run AutoHotkey again, which will load this file.\n"
"\n"
"#z::Run, www.autohotkey.com\n"
"\n"
"^!n::\n"
"IfWinExist, Untitled - Notepad\n"
"\tWinActivate\n"
"else\n"
"\tRun, Notepad\n"
"return\n"
"\n"
"\n"
"; Note: From now on whenever you run AutoHotkey directly, this script\n"
"; will be loaded.  So feel free to customize it to suit your needs.\n"
, fp2);
		fclose(fp2);
		// One or both of the below would probably fail -- at least on Win95 -- if mFileSpec ever
		// has spaces in it (since it's passed as the entire param string).  So enclose the filename
		// in double quotes.  I don't believe the directory needs to be in double quotes since it's
		// a separate field within the CreateProcess() and ShellExecute() structures:
		snprintf(buf, sizeof(buf), "\"%s\"", mFileSpec);
		if (!ActionExec("edit", buf, mFileDir, false))
			if (!ActionExec("Notepad.exe", buf, mFileDir, false))
			{
				MsgBox("The new script file was created, but could not be opened with the default editor or with Notepad.");
				return LOADING_FAILED;
			}
		// future: have it wait for the process to close, then try to open the script again:
		return 0;
	}
#endif

	if (LoadIncludedFile(mFileSpec, false) != OK)
		return LOADING_FAILED;

	// Rather than do this, which seems kinda nasty if ever someday support same-line
	// else actions such as "else return", just add two EXITs to the end of every script.
	// That way, if the first EXIT added accidentally "corrects" an actionless ELSE
	// or IF, the second one will serve as the anchoring end-point (mRelatedLine) for that
	// IF or ELSE.  In other words, since we never want mRelatedLine to be NULL, this should
	// make absolutely sure of that:
	//if (mLastLine->mActionType == ACT_ELSE ||
	//	ACT_IS_IF(mLastLine->mActionType)
	//	...
	++mCurrLineNumber;
	if (AddLine(ACT_EXIT) != OK) // First exit.
		return LOADING_FAILED;

	// Even if the last line of the script is already ACT_EXIT, always add another
	// one in case the script ends in a label.  That way, every label will have
	// a non-NULL target, which simplifies other aspects of script execution.
	// Making sure that all scripts end with an EXIT ensures that if the script
	// file ends with ELSEless IF or an ELSE, that IF's or ELSE's mRelatedLine
	// will be non-NULL, which further simplifies script execution:
	++mCurrLineNumber;
	if (AddLine(ACT_EXIT) != OK) // Second exit to guaranty non-NULL mRelatedLine(s).
		return LOADING_FAILED;

	// Always preparse the blocks before the If/Else's because If/Else may rely on blocks:
	if (PreparseBlocks(mFirstLine) && PreparseIfElse(mFirstLine))
	{
		// Use FindOrAdd, not Add, because the user may already have added it simply by
		// referring to it in the script:
		if (   !(g_ErrorLevel = FindOrAddVar("ErrorLevel"))   )
			return LOADING_FAILED; // Error.  Above already displayed it for us.
		// Initialize the var state to zero right before running anything in the script:
		g_ErrorLevel->Assign(ERRORLEVEL_NONE);

		// Initialize the random number generator:
		// Note: On 32-bit hardware, the generator module uses only 2506 bytes of static
		// data, so it doesn't seem worthwhile to put it in a class (so that the mem is
		// only allocated on first use of the generator).  For v1.0.24, _ftime() is not
		// used since it could be as large as 0.5 KB of non-compressed code.  A simple call to
		// GetSystemTimeAsFileTime() seems just as good or better, since it produces
		// a FILETIME, which is "the number of 100-nanosecond intervals since January 1, 1601."
		FILETIME ft;
		GetSystemTimeAsFileTime(&ft);
		init_genrand(ft.dwLowDateTime); // Use the low-order DWORD since the high-order one rarely changes.
		return mLineCount; // The count of runnable lines that were loaded, which might be zero.
	}
	else
		return LOADING_FAILED; // Error was already displayed by the above calls.
}



ResultType Script::LoadIncludedFile(char *aFileSpec, bool aAllowDuplicateInclude)
// Returns the number of non-comment lines that were loaded, or LOADING_FAILED on error.
// Below: Use double-colon as delimiter to set these apart from normal labels.
// The main reason for this is that otherwise the user would have to worry
// about a normal label being unintentionally valid as a hotkey, e.g.
// "Shift:" might be a legitimate label that the user forgot is also
// a valid hotkey:
#define HOTKEY_FLAG "::"
#define HOTKEY_FLAG_LENGTH 2
{
	if (!aFileSpec || !*aFileSpec) return FAIL;

	if (Line::nSourceFiles >= MAX_SCRIPT_FILES)
	{
		// Only 255 because the main file uses up one slot:
		MsgBox("The number of included files cannot exceed 255.");
		return FAIL;
	}

	// Keep this var on the stack due to recursion, which allows newly created lines to be given the
	// correct file number even when some #include's have been encountered in the middle of the script:
	UCHAR source_file_number = Line::nSourceFiles;

	if (!source_file_number)
		// Since this is the first source file, it must be the main script file.  Just point it to the
		// location of the filespec already dynamically allocated:
		Line::sSourceFile[source_file_number] = mFileSpec;
	else
	{
		// Get the full path in case aFileSpec has a relative path.  This is done so that duplicates
		// can be reliably detected (we only want to avoid including a given file more than once):
		char full_path[MAX_PATH], *filename_marker;
		GetFullPathName(aFileSpec, sizeof(full_path), full_path, &filename_marker);
		// Check if this file was already included.  If so, it's not an error because we want
		// to support automatic "include once" behavior.  So just ignore repeats:
		if (!aAllowDuplicateInclude)
			for (int f = 0; f < source_file_number; ++f)
				if (!stricmp(Line::sSourceFile[f], full_path)) // Case insensitive like the file system.
					return OK;
		Line::sSourceFile[source_file_number] = SimpleHeap::Malloc(full_path);
	}
	++Line::nSourceFiles;

	UCHAR *script_buf = NULL;  // Init for the case when the buffer isn't used (non-standalone mode).
	ULONG nDataSize = 0;

#ifdef AUTOHOTKEYSC
	HS_EXEArc_Read oRead;
	// AutoIt3: Open the archive in this compiled exe.
	// Jon gave me some details about why a password isn't needed: "The code in those libararies will
	// only allow files to be extracted from the exe is is bound to (i.e the script that it was
	// compiled with).  There are various checks and CRCs to make sure that it can't be used to read
	// the files from any other exe that is passed."
	if (oRead.Open(aFileSpec, "") != HS_EXEARC_E_OK)
	{
		MsgBox("Could not open the script inside the EXE.", 0, aFileSpec);
		return FAIL;
	}
	// AutoIt3: Read the script (the func allocates the memory for the buffer :) )
	if (oRead.FileExtractToMem(">AUTOHOTKEY SCRIPT<", &script_buf, &nDataSize) == HS_EXEARC_E_OK)
		mCompiledHasCustomIcon = false;
	else if (oRead.FileExtractToMem(">AHK WITH ICON<", &script_buf, &nDataSize) == HS_EXEARC_E_OK)
		mCompiledHasCustomIcon = true;
	else
	{
		oRead.Close();							// Close the archive
		MsgBox("Could not extract script from EXE.", 0, aFileSpec);
		return FAIL;
	}
	UCHAR *script_buf_marker = script_buf;  // "marker" will track where we are in the mem. file as we read from it.

	// Must cast to int to avoid loss of negative values:
	#define SCRIPT_BUF_SPACE_REMAINING ((int)(nDataSize - (script_buf_marker - script_buf)))
	int script_buf_space_remaining, max_chars_to_read;

	// AutoIt3: We have the data in RAW BINARY FORM, the script is a text file, so
	// this means that instead of a newline character, there may also be carridge
	// returns 0x0d 0x0a (\r\n)
	HS_EXEArc_Read *fp = &oRead;  // To help consolidate the code below.

#else  // Not in stand-alone mode, so read an external script file.
	// Future: might be best to put a stat() or GetFileAttributes() in here for better handling.
	FILE *fp = fopen(aFileSpec, "r");
	if (!fp)
	{
		char msg_text[MAX_PATH + 256];
		snprintf(msg_text, sizeof(msg_text), "%s file \"%s\" cannot be opened."
			, Line::nSourceFiles > 1 ? "#include" : "Script", aFileSpec);
		MsgBox(msg_text);
		return FAIL;
	}
#endif


	// File is now open, read lines from it.

	// <buf> should be no larger than LINE_SIZE because some later functions rely upon that:
	char buf[LINE_SIZE], *hotkey_flag, *cp, *cp1, *hotstring_start, *hotstring_options;
	HookActionType hook_action;
	size_t buf_length;
	bool is_label, section_comment = false;

	// Init both for main file and any included files loaded by this function:
	mCurrFileNumber = source_file_number;  // source_file_number is kept on the stack due to recursion.
	// Keep a copy of mCurrLineNumber on the stack to help with recursion:
#ifdef AUTOHOTKEYSC
	// -1 (MAX_UINT in this case) to compensate for the fact that there is a comment containing
	// the version number added to the top of each compiled script:
	LineNumberType line_number = mCurrLineNumber = -1;
#else
	LineNumberType line_number = mCurrLineNumber = 0;
#endif

	for (;;)
	{
		mCurrLine = NULL;  // To signify that we're in transition, trying to load a new one.
#ifdef AUTOHOTKEYSC
		script_buf_space_remaining = SCRIPT_BUF_SPACE_REMAINING;  // Temporary storage for performance.
		max_chars_to_read = (sizeof(buf) - 1 < script_buf_space_remaining) ? sizeof(buf) - 1
			: script_buf_space_remaining;
		if (   -1 == (buf_length = GetLine(buf, max_chars_to_read, script_buf_marker))   )
#else
		if (   -1 == (buf_length = GetLine(buf, (int)(sizeof(buf) - 1), fp))   )
#endif
			break;

		++mCurrLineNumber; // Keep track of the physical line number in the file for debugging purposes.
		++line_number; // A local copy on the stack to help with recursion.

		if (!buf_length)
			continue;

		if (section_comment) // Look for the uncomment-flag.
		{
			if (!strncmp(buf, "*/", 2))
			{
				section_comment = false;
				memmove(buf, buf + 2, buf_length - 2 + 1);  // +1 to include the string terminator.
				ltrim(buf); // Get rid of any whitespace that was between the comment-end and remaining text.
				if (!*buf) // The rest of the line is empty, so it was just a naked comment-end.
					continue;
				buf_length = strlen(buf);
			}
			else
				continue;
		}
		else if (!strncmp(buf, "/*", 2))
		{
			section_comment = true;
			continue; // It's now commented out, so the rest of this line is ignored.
		}


		// "::" alone isn't a hotstring, it's a label whose name is colon.
		// Below relies on the fact that no valid hotkey can start with a colon, since
		// ": & somekey" is not valid (since colon is a shifted key) and colon itself
		// should instead be defined as "+;::".  It also relies on short-circuit boolean:
		hotstring_start = hotstring_options = hotkey_flag = NULL;
		if (buf[0] == ':' && buf[1])
		{
			if (buf[1] != ':')
			{
				hotstring_options = buf + 1; // Point it to the hotstring's option letters.
				// Relies on the fact that options should never contain a literal colon:
				if (   !(hotstring_start = strchr(hotstring_options, ':'))   )
					hotstring_start = NULL; // Indicate that this isn't a hotstring after all.
				else
					++hotstring_start; // Points to the hotstring itself.
			}
			else // Double-colon, so it's a hotstring if there's more after this (but no options are present).
				if (buf[2])
					hotstring_start = buf + 2;
				//else it's just a naked "::", which is considered to be a mundane label whose name is colon.
		}
		if (hotstring_start)
		{
			// Find the hotstring's final double-colon by consider escape sequences from left to right.
			// This is necessary for to handles cases such as the following:
			// ::abc```::::Replacement String
			// The above hotstring translates literally into "abc`::".
			char *escaped_double_colon = NULL;
			for (cp = hotstring_start; ; ++cp)  // Increment to skip over the symbol just found by the inner for().
			{
				for (; *cp && *cp != g_EscapeChar && *cp != ':'; ++cp);  // Find the next escape char or colon.
				if (!*cp) // end of string.
					break;
				cp1 = cp + 1;
				if (*cp == ':')
				{
					if (*cp1 == ':') // Found a non-escaped double-colon, so this is the right one.
						hotkey_flag = cp++;  // Increment to have loop skip over both colons.
						// and the continue with the loop so that escape sequences in the replacement
						// text (if there is replacement text) are also translated.
					// else just a single colon, or the second colon of an escaped pair (`::), so continue.
					continue;
				}
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
				// Since single colons normally do not need to be escaped, this increments one extra
				// for double-colons to skip over the entire pair so that its second colon
				// is not seen as part of the hotstring's final double-colon.  Example:
				// ::ahc```::::Replacement String
				if (*cp == ':' && *cp1 == ':')
					++cp;
			} // for()
			if (!hotkey_flag)
				hotstring_start = NULL;  // Indicate that this isn't a hotstring after all.
		}
		else // Not a hotstring
			// Note that there may be an action following the HOTKEY_FLAG (on the same line).
			hotkey_flag = strstr(buf, HOTKEY_FLAG); // Find the first one from the left, in case there's more than 1.

		// The below considers all of the following to be non-hotkey labels:
		// A naked "::", which is treated as a normal label whose label name is colon.
		// A command containing an escaped literal double-colon, e.g. Run, Something.exe `:: /b
		//    In the above case, the escape sequence just tell us here not to process it as a label,
		//    and later the escape sequence `: resolves to a single colon.
		// But the below also ensures that `:: is a valid hotkey label whose hotkey is accent.
		// And it relies on short-circuit boolean:
		is_label = (hotkey_flag && hotkey_flag > buf);
		if (is_label && !hotstring_start && *(hotkey_flag - 1) == g_EscapeChar && hotkey_flag - buf > 2)
		{
			// Greater-than 2 is used above because "x=`::" is currently the smallest possible command
			// that might be confused as a hotkey label.
			// Since it appears to be a hotkey label and it's known not to be a hotstring label (hotstrings
			// can't suffer from this type of ambiguity because a leading colon or pair of colons makes them
			// easier to detect), ensure it really is a hotkey label rather than an escaped double-colon by
			// eliminating ambiguity as much as possible.  Examples of ambiguity:
			//run Enter & `::
			//vs.
			//Enter & `::   (i.e. accent is the suffix of a composite hotkey)
			//
			//x = Enter & `::
			//vs.
			//Enter & `::
			//
			//x=^`::
			//vs.
			//^`::  (i.e. accent is the suffix key)
			// First check if it has the composite delimiter.  If it does, it's 99.9% certain that it's
			// a hotkey due to the extreme odds against " & `::" ever appearing literally in a command.
			size_t available_length = hotkey_flag - buf;
			if (available_length <= COMPOSITE_DELIMITER_LENGTH + 1  // i.e. it must long enough to contain "x & `::"
                || strnicmp(hotkey_flag - 4, COMPOSITE_DELIMITER, COMPOSITE_DELIMITER_LENGTH)) // it's not it.
			{
				// Assume it's not a label unless it's short, contains no spaces, and contains only
				// hotkey modifiers.
				is_label = false;
				if (available_length < 10)  // Longest possible non-composite hotkey: "$*~<^+!#`::"
				{
					char *bcp;
					for (bcp = buf; bcp < hotkey_flag - 1; ++bcp)
						if (!strchr("><*~$!^+#", *bcp))
							break;
					if (bcp == hotkey_flag - 1)
						// It appears to be a hotkey after all, since only modifiers exist to the left of "`::"
						is_label = true;
				}
				//else it's too long to be a hotkey.
			}
		}

		if (is_label) // It's a label and a hotkey.
		{
			*hotkey_flag = '\0'; // Terminate so that buf is now the label itself.
			hotkey_flag += HOTKEY_FLAG_LENGTH;  // Now hotkey_flag is the hotkey's action, if any.
			if (!hotstring_start)
			{
				ltrim(hotkey_flag); // Has already been rtrimmed by GetLine().
				rtrim(buf); // Trim the new substring inside of buf (due to temp termination). It has already been ltrimmed.
			}
			// else don't trim hotstrings since literal spaces in both substrings are significant.

			// If this is the first hotkey label encountered, Add a return before
			// adding the label, so that the auto-exectute section is terminated.
			// Only do this if the label is a hotkey because, for example,
			// the user may want to fully execute a normal script that contains
			// no hotkeys but does contain normal labels to which the execution
			// should fall through, if specified, rather than returning.
			// But this might result in weirdness?  Example:
			//testlabel:
			// Sleep, 1
			// return
			// ^a::
			// return
			// It would put the hard return in between, which is wrong.  But in the case above,
			// the first sub shouldn't have a return unless there's a part up top that ends in Exit.
			// So if Exit is encountered before the first hotkey, don't add the return?
			// Even though wrong, the return is harmless because it's never executed?  Except when
			// falling through from above into a hotkey (which probably isn't very valid anyway)?
			// Update: Below must check if there are any true hotkey labels, not just regular labels.
			// Otherwise, a normal (non-hotkey) label in the autoexecute section would count and
			// thus the RETURN would never be added here, even though it should be:
			if (mNoHotkeyLabels)
			{
				mNoHotkeyLabels = false;
				if (AddLine(ACT_RETURN) != OK)
					return CloseAndReturn(fp, script_buf, FAIL);
				mCurrLine = NULL;  // To signify that we're in transition, trying to load a new one.
			}
			// For hotstrings, the below makes the label include leading colon(s) and the full option
			// string (if any) so that the uniqueness of labels is preserved.  For example, we want
			// the following two hotstring labels to be unique rather than considered duplicates:
			// ::abc::
			// :c:abc::
			if (AddLabel(buf) != OK) // Always add a label before adding the first line of its section.
				return CloseAndReturn(fp, script_buf, FAIL);
			if (*hotkey_flag) // This hotkey's action is on the same line as its label.
			{
				if (!hotstring_start)
					// Don't add the alt-tabs as a line, since it has no meaning as a script command.
					// But do put in the Return regardless, in case this label is ever jumped to
					// via Goto/Gosub:
					if (   !(hook_action = Hotkey::ConvertAltTab(hotkey_flag, false))   )
						if (ParseAndAddLine(hotkey_flag) != OK)
							return CloseAndReturn(fp, script_buf, FAIL);
				// Also add a Return that's implicit for a single-line hotkey.  This is also
				// done for auto-replace hotstrings in case gosub/goto is ever used to jump
				// to their labels:
				if (AddLine(ACT_RETURN) != OK)
					return CloseAndReturn(fp, script_buf, FAIL);
			}
			else
				hook_action = 0;

			if (hotstring_start)
			{
				if (!*hotstring_start)
				{
					// The following error message won't indicate the correct line number because
					// the hotstring (as a label) does not actually exist as a line.  But it seems
					// best to report it this way in case the hotstring is inside a #Include file,
					// so that the correct file name and approximate line number are shown:
					ScriptError("This hotstring is missing its abbreviation.", hotkey_flag);
					return CloseAndReturn(fp, script_buf, FAIL);
				}
				// In the case of hotstrings, hotstring_start is the beginning of the hotstring itself,
				// i.e. the character after the second colon.  hotstring_options is NULL if no options,
				// otherwise it's the first character in the options list (option string is not terminated,
				// but instead ends in a colon).  hotkey_flag is blank if it's not an auto-replace
				// hotstring, otherwise it contains the auto-replace text.
				if (!Hotstring::AddHotstring(mLastLabel, hotstring_options ? hotstring_options : ""
					, hotstring_start, hotkey_flag))
					return CloseAndReturn(fp, script_buf, FAIL);
			}
			else
				if (Hotkey::AddHotkey(mLastLabel, hook_action) != OK) // Set hotkey to jump to this label.
					return CloseAndReturn(fp, script_buf, FAIL);
			continue;
		}

		// Otherwise, not a hotkey.  Check if it's a generic, non-hotkey label:
		if (buf[buf_length - 1] == ':') // Labels must end in a colon (buf was previously rtrimmed).
			// Labels (except hotkeys) must contain no whitespace, delimiters, or escape-chars.
			// This is to avoid problems where a legitimate action-line ends in a colon,
			// such as WinActivate, SomeTitle:
			// We allow hotkeys to violate this since they may contain commas, and since a normal
			// script line (i.e. just a plain command) is unlikely to ever end in a double-colon:
			for (cp = buf, is_label = true; *cp; ++cp)
				if (IS_SPACE_OR_TAB(*cp) || *cp == g_delimiter || *cp == g_EscapeChar)
				{
					is_label = false;
					break;
				}
		if (is_label)
		{
			buf[buf_length - 1] = '\0';  // Remove the trailing colon.
			rtrim(buf); // Has already been ltrimmed.
			if (AddLabel(buf) != OK)
				return CloseAndReturn(fp, script_buf, FAIL);
			continue;
		}
		// It's not a label.
		if (*buf == '#')
		{
			switch(IsPreprocessorDirective(buf))
			{
			case CONDITION_TRUE:
				// Since the directive may have been a #include which called us recursively,
				// restore the class's values for these two, which are maintained separately
				// like this to avoid having to specify them in various calls, especially the
				// hundreds of calls to ScriptError() and LineError():
				mCurrFileNumber = source_file_number;
				mCurrLineNumber = line_number;
				continue;
			case FAIL:
				return CloseAndReturn(fp, script_buf, FAIL); // It already reported the error.
			// Otherwise it's CONDITION_FALSE.  Do nothing.
			}
		}
		// Otherwise it's just a normal script line.
		// First do a little special handling to support actions on the same line as their
		// ELSE, e.g.:
		// else if x = 1
		// This is done here rather than in ParseAndAddLine() because it's fairly
		// complicated to do there (already tried it) mostly due to the fact that
		// literal_map has to be properly passed in a recursive call to itself, as well
		// as properly detecting special commands that don't have keywords such as
		// IF comparisons, ACT_ASSIGN, +=, -=, etc.
		char *action_start = omit_leading_whitespace(buf);
		char *action_end = *action_start ? StrChrAny(action_start, "\t ") : NULL;
		if (!action_end)
			action_end = action_start + strlen(action_start);
		// Now action_end is the position of the terminator, or the tab/space following the command name.
		if (strlicmp(action_start, g_act[ACT_ELSE].Name, (UINT)(action_end - action_start)))
		{
			if (ParseAndAddLine(buf) != OK)
				return CloseAndReturn(fp, script_buf, FAIL);
		}
		else // This line is an ELSE.
		{
			// Add the ELSE directly rather than calling ParseAndAddLine() because that function
			// would resolve escape sequences throughout the entire length of <buf>, which we
			// don't want because we wouldn't have access to the corresponding literal-map to
			// figure out the proper use of escaped characters:
			if (AddLine(ACT_ELSE) != OK)
				return CloseAndReturn(fp, script_buf, FAIL);
			mCurrLine = NULL;  // To signify that we're in transition, trying to load a new one.
			action_end = omit_leading_whitespace(action_end); // Now action_end is the word after the ELSE.
			if (*action_end == g_delimiter) // Allow "else, action"
				action_end = omit_leading_whitespace(action_end + 1);
			if (*action_end && ParseAndAddLine(action_end) != OK)
				return CloseAndReturn(fp, script_buf, FAIL);
			// Otherwise, there was either no same-line action or the same-line action was successfully added,
			// so do nothing.
		}
	}

#ifdef AUTOHOTKEYSC
	// AutoIt3: Close the archive and free the file in memory
	free(script_buf);
	oRead.Close();
#else
	fclose(fp);
#endif
	return OK;
}



// Small inline to make LoadIncludedFile() code cleaner.
#ifdef AUTOHOTKEYSC
inline ResultType Script::CloseAndReturn(HS_EXEArc_Read *fp, UCHAR *aBuf, ResultType aReturnValue)
{
	free(aBuf);
	fp->Close();
	return aReturnValue;
}
#else
inline ResultType Script::CloseAndReturn(FILE *fp, UCHAR *aBuf, ResultType aReturnValue)
{
	// aBuf is unused in this case.
	fclose(fp);
	return aReturnValue;
}
#endif



#ifdef AUTOHOTKEYSC
size_t Script::GetLine(char *aBuf, int aMaxCharsToRead, UCHAR *&aMemFile) // last param = reference to pointer
#else
size_t Script::GetLine(char *aBuf, int aMaxCharsToRead, FILE *fp)
#endif
{
	size_t aBuf_length = 0;
#ifdef AUTOHOTKEYSC
	if (!aBuf || !aMemFile) return -1;
	if (aMaxCharsToRead <= 0) return -1; // We're signaling to caller that the end of the memory file has been reached.
	// Otherwise, continue reading characters from the memory file until either a newline is
	// reached or aMaxCharsToRead have been read:
	// Track "i" separately from aBuf_length because we want to read beyond the bounds of the memory file.
	int i;
	for (i = 0; i < aMaxCharsToRead; ++i)
	{
		if (aMemFile[i] == '\n')
		{
			// The end of this line has been reached.  Don't copy this char into the target buffer.
			// In addition, if the previous char was '\r', remove it from the target buffer:
			if (aBuf_length > 0 && aBuf[aBuf_length - 1] == '\r')
				aBuf[--aBuf_length] = '\0';
			++i; // i.e. so that aMemFile will be adjusted to omit this newline char.
			break;
		}
		else
			aBuf[aBuf_length++] = aMemFile[i];
	}
	// We either read aMaxCharsToRead or reached the end of the line (as indicated by the newline char).
	// In the former case, aMemFile might now be changed to be a position outside the bounds of the
	// memory area, which the caller will reflect back to us during the next call as a 0 value for
	// aMaxCharsToRead, which we then signal to the caller (above) as the end of the file):
	aMemFile += i; // Update this value for use by the caller.
	// Terminate the buffer (the caller has already ensured that there's room for the terminator
	// via its value of aMaxCharsToRead):
	aBuf[aBuf_length] = '\0';
#else
	if (!aBuf || !fp) return -1;
	if (aMaxCharsToRead <= 0) return 0;
	if (feof(fp)) return -1; // Previous call to this function probably already read the last line.
	if (fgets(aBuf, aMaxCharsToRead, fp) == NULL) // end-of-file or error
	{
		*aBuf = '\0';  // Reset since on error, contents added by fgets() are indeterminate.
		return -1;
	}
	aBuf_length = strlen(aBuf);
	if (!aBuf_length)
		return 0;
	if (aBuf[aBuf_length-1] == '\n')
		aBuf[--aBuf_length] = '\0';
	if (aBuf[aBuf_length-1] == '\r')  // In case there are any, e.g. a Macintosh or Unix file?
		aBuf[--aBuf_length] = '\0';
#endif

	// ltrim to support semicolons after tab keys or other whitespace.  Seems best to rtrim also:
	trim(aBuf);
	if (!strncmp(aBuf, g_CommentFlag, g_CommentFlagLength)) // Case sensitive.
	{
		*aBuf = '\0';
		return 0;
	}

	// Handle comment-flags that appear to the right of a valid line.  But don't
	// allow these types of comments if the script is considers to be the AutoIt2
	// style, to improve compatibility with old scripts that may use non-escaped
	// comment-flags as literal characters rather than comments:
	if (g_AllowSameLineComments)
	{
		char *cp, *prevp;
		for (cp = strstr(aBuf, g_CommentFlag); cp; cp = strstr(cp + g_CommentFlagLength, g_CommentFlag))
		{
			// If no whitespace to its left, it's not a valid comment.
			// We insist on this so that a semi-colon (for example) immediately after
			// a word (as semi-colons are often used) will not be considered a comment.
			prevp = cp - 1;
			if (prevp < aBuf) // should never happen because we already checked above.
			{
				*aBuf = '\0';
				return 0;
			}
			if (IS_SPACE_OR_TAB_OR_NBSP(*prevp)) // consider it to be a valid comment flag
			{
				*prevp = '\0';
				rtrim_with_nbsp(aBuf); // Since it's our responsibility to return a fully trimmed string.
				break; // Once the first valid comment-flag is found, nothing after it can matter.
			}
			else // No whitespace to the left.
				if (*prevp == g_EscapeChar) // Remove the escape char.
					memmove(prevp, prevp + 1, strlen(prevp + 1) + 1);  // +1 for the terminator.
					// Then continue looking for others.
				// else there wasn't any whitespace to its left, so keep looking in case there's
				// another further on in the line.
		}
	}

	return strlen(aBuf);  // Return an updated length due to trim().
}



inline ResultType Script::IsPreprocessorDirective(char *aBuf)
// Returns CONDITION_TRUE, CONDITION_FALSE, or FAIL.
// Note: Don't assume that every line in the script that starts with '#' is a directive
// because hotkeys can legitimately start with that as well.  i.e., the following line should
// not be unconditionally ignored, just because it starts with '#', since it is a valid hotkey:
// #y::run, notepad
{
	char end_flags[] = {' ', '\t', g_delimiter, '\0'}; // '\0' must be last.
	char *cp, *directive_end;
	int value; // Helps detect values that are too large, since some of the target globals are UCHAR.

	// Use strnicmp() so that a match is found as long as aBuf starts with the string in question.
	// e.g. so that "#SingleInstance, on" will still work too, but
	// "#a::run, something, "#SingleInstance" (i.e. a hotkey) will not be falsely detected
	// due to using a more lenient function such as strcasestr().  UPDATE: Using strlicmp() now so
	// that overlapping names, such as #MaxThreads and #MaxThreadsPerHotkey won't get mixed up:
	if (   !(directive_end = StrChrAny(aBuf, end_flags))   )
		directive_end = aBuf + strlen(aBuf); // Point it to the zero terminator.
	#define IS_DIRECTIVE_MATCH(directive) (!strlicmp(aBuf, directive, (UINT)(directive_end - aBuf)))

	if (IS_DIRECTIVE_MATCH("#WinActivateForce"))
	{
		g_WinActivateForce = true;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#Persistent"))
	{
		g_persistent = true;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#ErrorStdOut"))
	{
		mErrorStdOut = true;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#SingleInstance"))
	{
		g_AllowOnlyOneInstance = SINGLE_INSTANCE_PROMPT; // Set default.
		if (*directive_end && *(cp = omit_leading_whitespace(directive_end)))
		{
			if (*cp == g_delimiter)
				cp = omit_leading_whitespace(cp + 1);
			if (!stricmp(cp, "force"))
				g_AllowOnlyOneInstance = SINGLE_INSTANCE_REPLACE;
			else if (!stricmp(cp, "ignore"))
				g_AllowOnlyOneInstance = SINGLE_INSTANCE_IGNORE;
			else if (!stricmp(cp, "off"))
				g_AllowOnlyOneInstance = SINGLE_INSTANCE_OFF;
		}
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#Hotstring"))
	{
		if (*directive_end && *(cp = omit_leading_whitespace(directive_end)))
		{
			if (*cp == g_delimiter)
				cp = omit_leading_whitespace(cp + 1);
			if (!*cp)
				return CONDITION_TRUE;
			char *suboption = strcasestr(cp, "EndChars");
			if (suboption)
			{
				// Since it's not realistic to have only a couple, spaces and literal tabs
				// must be included in between other chars, e.g. `n `t has a space in between.
				// Also, EndChar  \t  will have a space and a tab since there are two spaces
				// after the word EndChar.
				if (    !(cp = StrChrAny(suboption, "\t "))   )
					return CONDITION_TRUE;
				strlcpy(g_EndChars, ++cp, sizeof(g_EndChars));
				ConvertEscapeSequences(g_EndChars, g_EscapeChar);
				return CONDITION_TRUE;
			}
		}
		// Otherwise assume it's a list of options.  Note that for compatibility with its
		// other caller, it will stop at end-of-string or ':', whichever comes first.
		Hotstring::ParseOptions(cp, g_HSPriority, g_HSKeyDelay, g_HSCaseSensitive, g_HSConformToCase
			, g_HSDoBackspace, g_HSOmitEndChar, g_HSSendRaw, g_HSEndCharRequired, g_HSDetectWhenInsideWord);
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#NoTrayIcon"))
	{
		g_NoTrayIcon = true;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#AllowSameLineComments"))  // i.e. There's no way to turn it off, only on.
	{
		g_AllowSameLineComments = true;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#UseHook"))
	{
		// Set the default mode that will be used if there's no parameter at all:
		g_ForceKeybdHook = true;
		#define RETURN_IF_NO_CHAR \
			if (!*directive_end)\
				return CONDITION_TRUE;\
			if (   !*(cp = omit_leading_whitespace(directive_end))   )\
				return CONDITION_TRUE;\
			if (*cp == g_delimiter && !*(cp = omit_leading_whitespace(cp + 1))   )\
				return CONDITION_TRUE;
		RETURN_IF_NO_CHAR
		if (Line::ConvertOnOff(cp) == TOGGLED_OFF)
			g_ForceKeybdHook = false;
		// else leave the default to "true" as set above.
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#InstallKeybdHook"))
	{
		// It seems best not to report this warning because a user may want to use partial functionality
		// of a script on Win9x:
		//MsgBox("#InstallKeybdHook is not supported on Windows 95/98/Me.  This line will be ignored.");
		if (!g_os.IsWin9x())
		{
			Hotkey::RequireHook(HOOK_KEYBD);
#ifdef HOOK_WARNING
			if (*directive_end && *(cp = omit_leading_whitespace(directive_end)))
			{
				if (*cp == g_delimiter)
					cp = omit_leading_whitespace(cp + 1);
				if (!stricmp(cp, "force"))
					sWhichHookSkipWarning |= HOOK_KEYBD;
			}
#endif
		}
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#InstallMouseHook"))
	{
		// It seems best not to report this warning because a user may want to use partial functionality
		// of a script on Win9x:
		//MsgBox("#InstallMouseHook is not supported on Windows 95/98/Me.  This line will be ignored.");
		if (!g_os.IsWin9x())
		{
			Hotkey::RequireHook(HOOK_MOUSE);
#ifdef HOOK_WARNING
			if (*directive_end && *(cp = omit_leading_whitespace(directive_end)))
			{
				if (!stricmp(cp, "force"))
					sWhichHookSkipWarning |= HOOK_MOUSE;
			}
#endif
		}
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#MaxThreadsBuffer"))
	{
		// Set the default mode that will be used if there's no parameter at all:
		g_MaxThreadsBuffer = true;
		RETURN_IF_NO_CHAR
		if (Line::ConvertOnOff(cp) == TOGGLED_OFF)
			g_MaxThreadsBuffer = false;
		// else leave the default to "true" as set above.
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#HotkeyModifierTimeout"))
	{
		RETURN_IF_NO_CHAR
		g_HotkeyModifierTimeout = ATOI(cp);  // cp was set to the right position by the above macro
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#MaxMem"))
	{
		RETURN_IF_NO_CHAR
		double valuef = ATOF(cp);  // cp was set to the right position by the above macro
		if (valuef > 4095)  // Don't exceed capacity of VarSizeType, which is currently a DWORD (4 gig).
			valuef = 4095;  // Don't use 4096 since that might be a special/reserved value for some functions.
		else if (valuef  < 1)
			valuef = 1;
		g_MaxVarCapacity = (VarSizeType)(valuef * 1024 * 1024);
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#MaxThreads"))
	{
		RETURN_IF_NO_CHAR
		value = ATOI(cp);  // cp was set to the right position by the above macro
		if (value > MAX_THREADS_LIMIT) // For now, keep this limited to prevent stack overflow due to too many pseudo-threads.
			value = MAX_THREADS_LIMIT;
		else if (value < 1)
			value = 1;
		g_MaxThreadsTotal = value;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#MaxThreadsPerHotkey"))
	{
		RETURN_IF_NO_CHAR
		// Use value as a temp holder since it's int vs. UCHAR and can thus detect very large or negative values:
		value = ATOI(cp);  // cp was set to the right position by the above macro
		if (value > MAX_THREADS_LIMIT) // For now, keep this limited to prevent stack overflow due to too many pseudo-threads.
			value = MAX_THREADS_LIMIT;
		else if (value < 1)
			value = 1;
		g_MaxThreadsPerHotkey = value; // Note: g_MaxThreadsPerHotkey is UCHAR.
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#HotkeyInterval"))
	{
		RETURN_IF_NO_CHAR
		g_HotkeyThrottleInterval = ATOI(cp);  // cp was set to the right position by the above macro
		if (g_HotkeyThrottleInterval < 10) // values under 10 wouldn't be useful due to timer granularity.
			g_HotkeyThrottleInterval = 10;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#MaxHotkeysPerInterval"))
	{
		RETURN_IF_NO_CHAR
		g_MaxHotkeysPerInterval = ATOI(cp);  // cp was set to the right position by the above macro
		if (g_MaxHotkeysPerInterval <= 0) // sanity check
			g_MaxHotkeysPerInterval = 1;
		return CONDITION_TRUE;
	}

	// For the below series, it seems okay to allow the comment flag to contain other reserved chars,
	// such as DerefChar, since comments are evaluated, and then taken out of the game at an earlier
	// stage than DerefChar and the other special chars.
	if (IS_DIRECTIVE_MATCH("#CommentFlag"))
	{
		RETURN_IF_NO_CHAR
		if (!*(cp + 1))  // i.e. the length is 1
		{
			// Don't allow '#' since it's the preprocessor directive symbol being used here.
			// Seems ok to allow "." to be the comment flag, since other constraints mandate
			// that at least one space or tab occur to its left for it to be considered a
			// comment marker.
			if (*cp == '#' || *cp == g_DerefChar || *cp == g_EscapeChar || *cp == g_delimiter)
				return ScriptError(ERR_DEFINE_CHAR);
			// Exclude hotkey definition chars, such as ^ and !, because otherwise
			// the following example wouldn't work:
			// User defines ! as the comment flag.
			// The following hotkey would never be in effect since it's considered to
			// be commented out:
			// !^a::run,notepad
			if (*cp == '!' || *cp == '^' || *cp == '+' || *cp == '$' || *cp == '~' || *cp == '*'
				|| *cp == '<' || *cp == '>')
				// Note that '#' is already covered by the other stmt. above.
				return ScriptError(ERR_DEFINE_COMMENT);
		}
		strlcpy(g_CommentFlag, cp, MAX_COMMENT_FLAG_LENGTH + 1);
		g_CommentFlagLength = strlen(g_CommentFlag);  // Keep this in sync with above.
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#EscapeChar"))
	{
		RETURN_IF_NO_CHAR
		// Don't allow '.' since that can be part of literal floating point numbers:
		if (   *cp == '#' || *cp == g_DerefChar || *cp == g_delimiter || *cp == '.'
			|| (g_CommentFlagLength == 1 && *cp == *g_CommentFlag)   )
			return ScriptError(ERR_DEFINE_CHAR);
		g_EscapeChar = *cp;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#DerefChar"))
	{
		RETURN_IF_NO_CHAR
		if (   *cp == '#' || *cp == g_EscapeChar || *cp == g_delimiter || *cp == '.'
			|| (g_CommentFlagLength == 1 && *cp == *g_CommentFlag)   )
			return ScriptError(ERR_DEFINE_CHAR);
		g_DerefChar = *cp;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#Delimiter"))
	{
		// This macro will skip over any leading delimiter than may be present, e.g. #Delimiter, ^
		// This should be okay since the user shouldn't be attempting to change the delimiter
		// to what it already is, and even if this is attempted, it would just be ignored.
		// For example, "#Delimiter ," isn't meaningful if the delimiter already is a comma,
		// which is good because the RETURN_IF_NO_CHAR macro would assume that the comma
		// is accidental (not a symbol) and try to find the symbol after it.
		RETURN_IF_NO_CHAR
		if (   *cp == '#' || *cp == g_EscapeChar || *cp == g_DerefChar || *cp == '.'
			|| (g_CommentFlagLength == 1 && *cp == *g_CommentFlag)   )
			return ScriptError(ERR_DEFINE_CHAR);
		g_delimiter = *cp;
		return CONDITION_TRUE;
	}

	bool include_again = false; // Set default in case of short-circuit boolean.
	if (IS_DIRECTIVE_MATCH("#Include") || (include_again = IS_DIRECTIVE_MATCH("#IncludeAgain")))
	{
		// Standalone EXEs ignore this directive since the included files were already merged in
		// with the main file when the script was compiled.  These should have been removed
		// or commented out by Ahk2Exe, but just in case, it's safest to ignore them:
#ifdef AUTOHOTKEYSC
		return CONDITION_TRUE;
#else
		if (   !*(cp = omit_leading_whitespace(directive_end))   )
			return ScriptError(ERR_INCLUDE_FILE);
		if (*cp == g_delimiter)
		{
			++cp;
			if (   !*(cp = omit_leading_whitespace(cp))   )
				return ScriptError(ERR_INCLUDE_FILE);
		}
		return (LoadIncludedFile(cp, include_again) == FAIL) ? FAIL : CONDITION_TRUE;  // It will have already displayed any error.
#endif
	}

	// Otherwise:
	return CONDITION_FALSE;
}



ResultType Script::UpdateOrCreateTimer(Label *aLabel, char *aPeriod, char *aPriority, bool aEnable
	, bool aUpdatePriorityOnly)
// Caller should specific a blank aPeriod to prevent the timer's period from being changed
// (i.e. if caller just wants to turn on or off an existing timer).  But if it does this
// for a non-existent timer, that timer will be created with the default period as specfied in
// the constructor.
{
	ScriptTimer *timer;
	for (timer = mFirstTimer; timer != NULL; timer = timer->mNextTimer)
		if (timer->mLabel == aLabel) // Match found.
			break;
	bool timer_existed = (timer != NULL);
	if (!timer_existed)  // Create it.
	{
		if (   !(timer = new ScriptTimer(aLabel))   )
			return ScriptError(ERR_OUTOFMEM);
		if (!mFirstTimer)
			mFirstTimer = mLastTimer = timer;
		else
		{
			mLastTimer->mNextTimer = timer;
			// This must be done after the above:
			mLastTimer = timer;
		}
		++mTimerCount;
	}
	// Update its members:
	if (aEnable && !timer->mEnabled) // Must check both or the below count will be wrong.
	{
		// The exception is if the timer already existed but the caller only wanted its priority changed:
		if (!(timer_existed && aUpdatePriorityOnly))
		{
			timer->mEnabled = true;
			++mTimerEnabledCount;
			SET_MAIN_TIMER  // Ensure the timer is always running when there is at least one enabled timed subroutine.
		}
		//else do nothing, leave it disabled.
	}
	else if (!aEnable && timer->mEnabled) // Must check both or the below count will be wrong.
	{
		timer->mEnabled = false;
		--mTimerEnabledCount;
		// If there are now no enabled timed subroutines, kill the main timer since there's no other
		// reason for it to exist if we're here.   This is because or direct or indirect caller is
		// currently always ExecUntil(), which doesn't need the timer while its running except to
		// support timed subroutines.  UPDATE: The above is faulty; Must also check g_nLayersNeedingTimer
		// because our caller can be one that still needs a timer as proven by this script that
		// hangs otherwise:
		//SetTimer, Test, on 
		//Sleep, 1000 
		//msgbox, done
		//return
		//Test: 
		//SetTimer, Test, off 
		//return
		if (!mTimerEnabledCount && !g_nLayersNeedingTimer && !Hotkey::sJoyHotkeyCount)
			KILL_MAIN_TIMER
	}

	if (*aPeriod) // Caller wanted us to update this member.
	{
		int period = ATOI(aPeriod);  // Always use this method & check to retain compatibility with existing scripts.
		// Allow a period of zero even though CheckScriptTimers() is usually called no more often than every 10ms:
		if (period >= 0)
			timer->mPeriod = period; // Read any float in a runtime variable reference as an int.
	}

	if (*aPriority) // Caller wants this member to be changed from its current or default value.
		timer->mPriority = ATOI(aPriority); // Read any float in a runtime variable reference as an int.

	if (!(timer_existed && aUpdatePriorityOnly))
		// Caller relies on us updating mTimeLastRun in this case.  This is done because it's more
		// flexible, e.g. a user might want to create a timer that is triggered 5 seconds from now.
		// In such a case, we don't want the timer's first triggering to occur immediately.
		// Instead, we want it to occur only when the full 5 seconds have elapsed:
		timer->mTimeLastRun = GetTickCount();

    // Below is obsolete, see above for why:
	// We don't have to kill or set the main timer because the only way this function is called
	// is directly from the execution of a script line inside ExecUntil(), in which case:
	// 1) KILL_MAIN_TIMER is never needed because the timer shouldn't exist while in ExecUntil().
	// 2) SET_MAIN_TIMER is never needed because it will be set automatically the next time ExecUntil()
	//    calls MsgSleep().
	return OK;
}



Label *Script::FindLabel(char *aLabelName)
// Returns the label whose name matches aLabelName, or NULL if not found.
{
	if (!aLabelName || !*aLabelName) return NULL;
	for (Label *label = mFirstLabel; label != NULL; label = label->mNextLabel)
		if (!stricmp(label->mName, aLabelName)) // Match found.
			return label;
	return NULL; // No match found.
}



ResultType Script::AddLabel(char *aLabelName)
// Returns OK or FAIL.
{
	if (!aLabelName || !*aLabelName) return FAIL;
	Label *duplicate_label = FindLabel(aLabelName);
	if (duplicate_label)
		// Don't attempt to dereference "duplicate_label->mJumpToLine because it might not
		// exist yet.  Example:
		// label1:
		// label1:  <-- This would be a dupe-error but it doesn't yet have an mJumpToLine.
		// return
		return ScriptError("This label has been defined more than once.", aLabelName);
	char *new_name = SimpleHeap::Malloc(aLabelName);
	if (!new_name)
		return FAIL;  // It already displayed the error for us.
	Label *the_new_label = new Label(new_name); // Pass it the dynamic memory area we created.
	if (the_new_label == NULL)
		return ScriptError(ERR_OUTOFMEM);
	the_new_label->mPrevLabel = mLastLabel;  // Whether NULL or not.
	if (mFirstLabel == NULL)
		mFirstLabel = mLastLabel = the_new_label;
	else
	{
		mLastLabel->mNextLabel = the_new_label;
		// This must be done after the above:
		mLastLabel = the_new_label;
	}
	++mLabelCount;
	return OK;
}



ResultType Script::ParseAndAddLine(char *aLineText, char *aActionName, char *aEndMarker
	, char *aLiteralMap, size_t aLiteralMapLength
	, ActionTypeType aActionType, ActionTypeType aOldActionType)
// Returns OK or FAIL.
// aLineText needs to be a string whose contents are modifiable (this
// helps performance by allowing the string to be split into sections
// without having to make temporary copies).
{
#ifdef _DEBUG
	if (!aLineText || !*aLineText)
		return ScriptError("ParseAndAddLine() called incorrectly.");
#endif

	// The characters below are ordered with most-often used ones first, for performance:
	#define DEFINE_END_FLAGS \
		char end_flags[] = {' ', g_delimiter, '(', '\t', '<', '>', ':', '=', '+', '-', '*', '/', '!', '~', '&', '|', '^', '\0'}; // '\0' must be last.
	DEFINE_END_FLAGS

	char action_name[MAX_VAR_NAME_LENGTH + 1], *end_marker;
	if (aActionName) // i.e. this function was called recursively with explicit values for the optional params.
	{
		strcpy(action_name, aActionName);
		end_marker = aEndMarker;
	}
	else
		if (   !(end_marker = ParseActionType(action_name, aLineText, true))   )
			return FAIL; // It already displayed the error.

	// Find the arguments (not to be confused with exec_params) of this action, if it has any:
	char *action_args = omit_leading_whitespace(end_marker + 1);
	// Now action_args is either the first delimiter or the first parameter (if it optional first
	// delimiter was omitted):
	bool is_var_and_operator;
	if (*action_args == g_delimiter)
	{
		is_var_and_operator = false;  // Since there's a comma, we know it's not, e.g. something, += 4
		// Find the start of the next token (or its ending delimiter if the token is blank such as ", ,"):
		for (++action_args; IS_SPACE_OR_TAB(*action_args); ++action_args);
	}
	else
	{
		// The next line is used to help avoid ambiguity of a line such as the following:
		// Input = test  ; Would otherwise be confused with the Input command.
		// But there may be times when a line like this would be used:
		// MsgBox =  ; i.e. the equals is intended to be the first parameter, not an operator.
		// In the above case, the user can provide the optional comma to avoid the ambiguity:
		// MsgBox, =
		switch(*action_args)
		{
		case '=':
		case ':':  // i.e. :=
		case '(':  // i.e. "if (expr)"
			is_var_and_operator = true;
			break;
		case '*':
		case '/':
		case '-':
		case '+':
			// Insist that the next symbol be equals to form a complete operator.  This allows
			// a line such as the following, which omits the first optional comma, to still
			// be recognized as a command rather than a variable-with-operator:
			// SetBatchLines -1
			is_var_and_operator = *(action_args + 1) == '=';
			break;
		default:
			is_var_and_operator = false;
		}
	}
	// Now the above has ensured that action_args is the first parameter itself, or empty-string if none.
	// If action_args now starts with a delimiter, it means that the first param is blank/empty.

	///////////////////////////////////////////////
	// Check if this line contains a valid command.
	///////////////////////////////////////////////
	// Check for macro commands first because these commands can *include* normal executable programs
	// and documents.  In other words, don't look for .EXE and such yet, because might find them
	// somewhere later in the string where they belong to some sub-action that we're not supposed
	// to handle:

	// It might perform a little worse on avg. to check for old cmds prior to checking for
	// the "special handling for ACT_ASSIGN" etc., but it makes things more understandable.
	// OLDER comment:
	// Check if it's an old command.  It's not necessary to do this before checking for the
	// special actions, above, because all the old commands shouldn't allow special chars
	// such as < > = (used above) to be used at the beginning of their first params.
	// e.g. IfEqual, varname, value ... SetEnv, varname, value ... EnvAdd, varname, value
	// And it helps avg. performance to check for old commands only after all checks for
	// new commands have been completed.
	ActionTypeType action_type = aActionType;
	ActionTypeType old_action_type = aOldActionType;
	if (action_type == ACT_INVALID && old_action_type == OLD_INVALID && !is_var_and_operator)
		if (   (action_type = ConvertActionType(action_name)) == ACT_INVALID   )
			old_action_type = ConvertOldActionType(action_name);

	/////////////////////////////////////////////////////////////////////////////
	// Special handling for ACT_ASSIGN/ADD/SUB/MULT/DIV and IFEQUAL/GREATER/LESS.
	/////////////////////////////////////////////////////////////////////////////
	if (action_type == ACT_INVALID && old_action_type == OLD_INVALID)
	{
		// No match found, but is it a special type of action?
		// Support for ++i and --i.  In these cases, action_name must be either "+" or "-"
		// and the first character of action_args must match it.
		if (   !*(action_name + 1) && ((*action_name == '+' && *action_args == '+')
			|| (*action_name == '-' && *action_args == '-'))   )
		{
			action_type = *action_name == '+' ? ACT_ADD : ACT_SUB;
			// Set action_args to be the word that occurs after the ++ or --:
			++action_args;
			action_args = omit_leading_whitespace(action_args); // Though there really shouldn't be any.
			// Set up aLineText and action_args to be parsed later on as a list of two parameters:
			// The variable name followed by the amount to be added or subtracted (e.g. "ScriptVar, 1").
			// We're not changing the length of aLineText by doing this, so it should be large enough:
			size_t new_length = strlen(action_args);
			// Since action_args is just a pointer into the aLineText buffer (which caller has ensured
			// is modifiable), use memmove() so that overlapping source & dest are properly handled:
			memmove(aLineText, action_args, new_length + 1); // +1 to include the zero terminator.
			// Append the second param, which is just "1" since the ++ and -- only inc/dec by 1:
			aLineText[new_length++] = g_delimiter;
			aLineText[new_length++] = '1';
			aLineText[new_length] = '\0';
			action_args = aLineText;
		}
		else if (!stricmp(action_name, "IF"))
		{
			if (*action_args == '(') // This is "if (expr)"
			{
				action_type = ACT_IFEXPR;
				// To support things like the following, the outermost enclosing parentheses are not removed:
				// if (x < 3) or (x > 6)
				// Also note that although the expr. must normally start with an open-paren to be
				// recognized as ACT_IFEXPR, it need not end in a close-paren, e.g. if (x = 1) or !done.
				// If these or any other parentheses are unbalanced, it will caught further below.
			}
			else 
			{
				// Skip over the variable name so that the "is" and "is not" operators are properly supported:
				char *operation = StrChrAny(action_args, end_flags);
				if (!operation)
					operation = action_args + strlen(action_args); // Point it to the NULL terminator instead.
				else
					operation = omit_leading_whitespace(operation);

				char *next_word;
				switch (*operation)
				{
				case '=': // But don't allow == to be "Equals" since the 2nd '=' might be literal.
					action_type = ACT_IFEQUAL;
					break;
				case '<':
					// Note: User can use whitespace to differentiate a literal symbol from
					// part of an operator, e.g. if var1 < =  <--- char is literal
					switch(*(operation + 1))
					{
					case '=': action_type = ACT_IFLESSOREQUAL; *(operation + 1) = ' '; break;
					case '>': action_type = ACT_IFNOTEQUAL; *(operation + 1) = ' '; break;
					default: action_type = ACT_IFLESS;  // i.e. some other symbol follows '<'
					}
					break;
				case '>': // Don't allow >< to be NotEqual since the '<' might be literal.
					if (*(operation + 1) == '=')
					{
						action_type = ACT_IFGREATEROREQUAL;
						*(operation + 1) = ' '; // Remove it from so that it won't be considered by later parsing.
					}
					else
						action_type = ACT_IFGREATER;
					break;
				case '!':
					if (*(operation + 1) == '=')
					{
						action_type = ACT_IFNOTEQUAL;
						*(operation + 1) = ' '; // Remove it from so that it won't be considered by later parsing.
					}
					else
						// To minimize the times where expressions must have an outer set of parens,
						// assume all unknown operators are expressions, e.g. "if !var"
						action_type = ACT_IFEXPR;
					break;
				case 'b': // "Between"
				case 'B':
					if (strnicmp(operation, "between", 7)) // Seems too rare a thing to warrant falling back to ACT_IFEXPR for this.
						return ScriptError("The word BETWEEN was expected but not found.", aLineText);
					action_type = ACT_IFBETWEEN;
					// Set things up to be parsed as args further down.  A delimiter is inserted later below:
					memset(operation, ' ', 7);
					break;
				case 'c': // "Contains"
				case 'C':
					if (strnicmp(operation, "contains", 8)) // Seems too rare a thing to warrant falling back to ACT_IFEXPR for this.
						return ScriptError("The word CONTAINS was expected but not found.", aLineText);
					action_type = ACT_IFCONTAINS;
					// Set things up to be parsed as args further down.  A delimiter is inserted later below:
					memset(operation, ' ', 8);
					break;
				case 'i':  // "is" or "is not"
				case 'I':
					switch (toupper(*(operation + 1))) // Seems too rare a thing to warrant falling back to ACT_IFEXPR for this.
					{
					case 's':  // "IS"
					case 'S':
						next_word = omit_leading_whitespace(operation + 2);
						if (strnicmp(next_word, "not", 3))
							action_type = ACT_IFIS;
						else
						{
							action_type = ACT_IFISNOT;
							// Remove the word "not" to set things up to be parsed as args further down.
							memset(next_word, ' ', 3);
						}
						*(operation + 1) = ' '; // Remove the 'S' in "IS".  'I' is replaced with ',' later below.
						break;
					case 'n':  // "IN"
					case 'N':
						action_type = ACT_IFIN;
						*(operation + 1) = ' '; // Remove the 'N' in "IN".  'I' is replaced with ',' later below.
						break;
					default:
						return ScriptError("The word IS or IN was expected but not found.", aLineText);
					} // switch()
					break;
				case 'n':  // It's either "not in", "not between", or "not contains"
				case 'N':
					if (strnicmp(operation, "not", 3)) // Seems too rare a thing to warrant falling back to ACT_IFEXPR for this.
						return ScriptError("The word NOT was expected but not found.", aLineText);
					// Remove the "NOT" separately in case there is more than one space or tab between
					// it and the following word, e.g. "not   between":
					memset(operation, ' ', 3);
					next_word = omit_leading_whitespace(operation + 3);
					if (!strnicmp(next_word, "in", 2))
					{
						action_type = ACT_IFNOTIN;
						memset(next_word, ' ', 2);
					}
					else if (!strnicmp(next_word, "between", 7))
					{
						action_type = ACT_IFNOTBETWEEN;
						memset(next_word, ' ', 7);
					}
					else if (!strnicmp(next_word, "contains", 8))
					{
						action_type = ACT_IFNOTCONTAINS;
						memset(next_word, ' ', 8);
					}
					break;

				default: // To minimize the times where expressions must have an outer set of parens, assume all unknown operators are expressions.
					action_type = ACT_IFEXPR;
				} // switch()

				// Set things up to be parsed as args later on:
				if (action_type != ACT_IFEXPR)
				{
					*operation = g_delimiter;
					if (action_type == ACT_IFBETWEEN || action_type == ACT_IFNOTBETWEEN)
					{
						// I decided against the syntax "if var between 3,8" because the gain in simplicity
						// and the small avoidance of ambiguity didn't seem worth the cost in terms of readability.
						for (next_word = operation;;)
						{
							if (   !(next_word = strcasestr(next_word, "and"))   )
								return ScriptError("BETWEEN requires the word AND.", aLineText); // Seems too rare a thing to warrant falling back to ACT_IFEXPR for this.
							if (strchr(" \t", *(next_word - 1)) && strchr(" \t", *(next_word + 3)))
							{
								// Since there's a space or tab on both sides, we know this is the correct "and",
								// i.e. not one contained within one of the parameters.  Examples:
								// if var between band and cat  ; Don't falsely detect "band"
								// if var betwwen Andy and David  ; Don't falsely detect "Andy".
								// Replace the word AND with a delimiter so that it will be parsed correctly later:
								*next_word = g_delimiter;
								*(next_word + 1) = ' ';
								*(next_word + 2) = ' ';
								break;
							}
							else
								next_word += 3;  // Skip over this false "and".
						}
					} // ACT_IFBETWEEN
				} // action_type != ACT_IFEXPR
			} // operation isn't "if (expr)"
		}
		else // The action type is something other than an IF.
		{
			if (*action_args == '=')
				action_type = ACT_ASSIGN;
			else if (*action_args == ':' && *(action_args + 1) == '=')  // :=
				action_type = ACT_ASSIGNEXPR;
			else if (*action_args == '+' && (*(action_args + 1) == '=' || *(action_args + 1) == '+')) // += or ++
				action_type = ACT_ADD;
			else if (*action_args == '-' && (*(action_args + 1) == '=' || *(action_args + 1) == '-')) // -= or --
				action_type = ACT_SUB;
			else if (*action_args == '*' && *(action_args + 1) == '=') // *=
				action_type = ACT_MULT;
			else if (*action_args == '/' && *(action_args + 1) == '=') // /=
				action_type = ACT_DIV;
			if (action_type != ACT_INVALID)
			{
				// Set things up to be parsed as args later on:
				*action_args = g_delimiter; // Replace the +,-,:,*,/ with a delimiter for later parsing.
				if (action_type != ACT_ASSIGN)
				{
					if (*(action_args + 1) == '=')
						*(action_args + 1) = ' ';  // Remove the "=" from consideration.
					else
						*(action_args + 1) = '1';  // Turn ++ and -- into ",1"
				}
				action_args = aLineText;
			}
		}
		if (action_type == ACT_INVALID)
			return ScriptError(ERR_UNRECOGNIZED_ACTION, aLineText);
	} // If no matching command found.

	Action *this_action = (action_type == ACT_INVALID) ? &g_old_act[old_action_type] : &g_act[action_type];

	//////////////////////////////////////////////////////////////////////////////////////////////
	// Handle escaped-sequences (escaped delimiters and all others except variable deref symbols).
	// This section must occur after all other changes to the pointer value action_args have
	// occurred above.
	//////////////////////////////////////////////////////////////////////////////////////////////
	// The size of this relies on the fact that caller made sure that aLineText isn't
	// longer than LINE_SIZE.  Also, it seems safer to use char rather than bool, even
	// though on most compilers they're the same size.  Char is always of size 1, but bool
	// can be bigger depending on platform/compiler:
	char literal_map[LINE_SIZE];
	ZeroMemory(literal_map, sizeof(literal_map));  // Must be fully zeroed for this purpose.
	if (aLiteralMap)
	{
		// Since literal map is NOT a string, just an array of char values, be sure to
		// use memcpy() vs. strcpy() on it.  Also, caller's aLiteralMap starts at aEndMarker,
		// so adjust it so that it starts at the newly found position of action_args instead:
		int map_offset = (int)(action_args - end_marker);
		int map_length = (int)(aLiteralMapLength - map_offset);
		if (map_length > 0)
			memcpy(literal_map, aLiteralMap + map_offset, map_length);
	}
	else
	{
		// Resolve escaped sequences and make a map of which characters in the string should
		// be interpreted literally rather than as their native function.  In other words,
		// convert any escape sequences in order from left to right (this order is important,
		// e.g. ``% should evaluate to `g_DerefChar not `LITERAL_PERCENT.  This part must be
		// done *after* checking for comment-flags that appear to the right of a valid line, above.
		// How literal comment-flags (e.g. semicolons) work:
		//string1; string2 <-- not a problem since string2 won't be considered a comment by the above.
		//string1 ; string2  <-- this would be a user mistake if string2 wasn't supposed to be a comment.
		//string1 `; string 2  <-- since esc seq. is resolved *after* checking for comments, this behaves as intended.
		// Current limitation: a comment-flag longer than 1 can't be escaped, so if "//" were used,
		// as a comment flag, it could never have whitespace to the left of it if it were meant to be literal.
		// Note: This section resolves all escape sequences except those involving g_DerefChar, which
		// are handled by a later section.
		char c;
		int i;
		for (i = 0; ; ++i)  // Increment to skip over the symbol just found by the inner for().
		{
			for (; action_args[i] && action_args[i] != g_EscapeChar; ++i);  // Find the next escape char.
			if (!action_args[i]) // end of string.
				break;
			c = action_args[i + 1];
			switch (c)
			{
				// Only lowercase is recognized for these:
				case 'a': action_args[i + 1] = '\a'; break;  // alert (bell) character
				case 'b': action_args[i + 1] = '\b'; break;  // backspace
				case 'f': action_args[i + 1] = '\f'; break;  // formfeed
				case 'n': action_args[i + 1] = '\n'; break;  // newline
				case 'r': action_args[i + 1] = '\r'; break;  // carriage return
				case 't': action_args[i + 1] = '\t'; break;  // horizontal tab
				case 'v': action_args[i + 1] = '\v'; break;  // vertical tab
			}
			// Replace escape-sequence with its single-char value.  This is done event if the pair isn't
			// a recognizable escape sequence (e.g. `? becomes ?), which is the Microsoft approach
			// and might not be a bad way of handing things.  There are some exceptions, however.
			// The first of these exceptions (g_DerefChar) is mandatory because that char must be
			// handled at a later stage or escaped g_DerefChars won't work right.  The others are
			// questionable, and might be worth further consideration.  UPDATE: g_DerefChar is now
			// done here because otherwise, examples such as this fail:
			// - The escape char is backslash.
			// - any instances of \\%, such as c:\\%var% , will not work because the first escape
			// sequence (\\) is resolved to a single literal backslash.  But then when \% is encountered
			// by the section that resolves escape sequences for g_DerefChar, the backslash is seen
			// as an escape char rather than a literal backslash, which is not correct.  Thus, we
			// resolve all escapes sequences HERE in one go, from left to right.

			// AutoIt2 definitely treats an escape char that occurs at the very end of
			// a line as literal.  It seems best to also do it for these other cases too.
			// UPDATE: I cannot reproduce the above behavior in AutoIt2.  Maybe it only
			// does it for some commands or maybe I was mistaken.  So for now, this part
			// is disabled:
			//if (c == '\0' || c == ' ' || c == '\t')
			//	literal_map[i] = 1;  // In the map, mark this char as literal.
			//else
			{
				// So these are also done as well, and don't need an explicit check:
				// g_EscapeChar , g_delimiter , (when g_CommentFlagLength > 1 ??): *g_CommentFlag
				// Below has a final +1 to include the terminator:
				MoveMemory(action_args + i, action_args + i + 1, strlen(action_args + i + 1) + 1);
				literal_map[i] = 1;  // In the map, mark this char as literal.
			}
			// else: Do nothing, even if the value is zero (the string's terminator).
		}
	}

	////////////////////////////////////////////////////////////////////////////////////////
	// Do some special preparsing of the MsgBox command, since it is so frequently used and
	// it is also the source of problem areas going from AutoIt2 to 3 and also due to the
	// new numeric parameter at the end.  Whenever possible, we want to avoid the need for
	// the user to have to escape commas that are intended to be literal.
	///////////////////////////////////////////////////////////////////////////////////////
	int mark, max_params_override = 0; // Set default.
	if (action_type == ACT_MSGBOX)
	{
		// First find out how many non-literal (non-escaped) delimiters are present.
		// Use a high maximum so that we can almost always find and analyze the command's
		// last apparent parameter.  This helps error-checking be more informative in a
		// case where the command specifies a timeout as its last param but it's next-to-last
		// param contains delimiters that the user forgot to escape.  In other words, this
		// helps detect more often when the user is trying to use the timeout feature.
		// If this weren't done, the command would more often forgive improper syntax
		// and not report a load-time error, even though it's pretty obvious that a load-time
		// error should have been reported:
		#define MAX_MSGBOX_DELIMITERS 20
		char *delimiter[MAX_MSGBOX_DELIMITERS];
		int delimiter_count;
		for (mark = delimiter_count = 0; action_args[mark] && delimiter_count < MAX_MSGBOX_DELIMITERS;)
		{
			for (; action_args[mark]; ++mark)
				if (action_args[mark] == g_delimiter && !literal_map[mark]) // Match found: a non-literal delimiter.
				{
					delimiter[delimiter_count++] = action_args + mark;
					++mark; // Skip over this delimiter for the next iteration of the outer loop.
					break;
				}
		}
		// If it has only 1 arg (i.e. 0 delimiters within the arg list) no override is needed.
		// Otherwise do more checking:
		if (delimiter_count)
		{
			// If the first apparent arg is not a non-blank pure number or there are apparently
			// only 2 args present (i.e. 1 delimiter in the arg list), assume the command is being
			// used in its 1-parameter mode:
			if (delimiter_count <= 1) // 2 parameters or less.
				// Force it to be 1-param mode.  In other words, we want to make MsgBox a very forgiving
				// command and have it rarely if ever report syntax errors:
				max_params_override = 1;
			else // It has more than 3 apparent params, but is the first param even numeric?
			{
				*delimiter[0] = '\0'; // Temporarily terminate action_args at the first delimiter.
				// Note: If it's a number inside a variable reference, it's still considered 1-parameter
				// mode to avoid ambiguity (unlike the new deref checking for param #4 mentioned below,
				// there seems to be too much ambiguity in this case to justify trying to figure out
				// if the first parameter is a pure deref, and thus that the command should use
				// 3-param or 4-param mode instead).
				if (!IsPureNumeric(action_args)) // No floats allowed.  Allow all-whitespace for aut2 compatibility.
					max_params_override = 1;
				*delimiter[0] = g_delimiter; // Restore the string.
				if (!max_params_override)
				{
					// IMPORATANT: The MsgBox cmd effectively has 3 parameter modes:
					// 1-parameter (where all commas in the 1st parameter are automatically literal)
					// 3-parameter (where all commas in the 3rd parameter are automatically literal)
					// 4-parameter (whether the 4th parameter is the timeout value)
					// Thus, the below must be done in a way that recognizes & supports all 3 modes.
					// The above has determined that the cmd isn't in 1-parameter mode.
					// If at this point it has exactly 3 apparent params, allow the command to be
					// processed normally without an override.  Otherwise, do more checking:
					if (delimiter_count == 3) // i.e. 3 delimiters, which means 4 params.
					{
						// If the 4th parameter isn't blank or pure numeric (i.e. even if it's a pure
						// deref, since trying to figure out what's a pure deref is somewhat complicated
						// at this early stage of parsing), assume the user didn't intend it to be the
						// MsgBox timeout (since that feature is rarely used), instead intending it
						// to be part of parameter #3.
						if (!IsPureNumeric(delimiter[2] + 1, false, true, true))
						{
							// Not blank and not a int or float.  Update for v1.0.20: Check if it's a
							// single deref.  If so, assume that deref contains the timeout and thus
							// 4-param mode is in effect.  This allows the timeout to be contained in
							// a variable, which was requested by one user:
							char *cp = omit_leading_whitespace(delimiter[2] + 1);
							// Relies on short-circuit boolean order:
							if (*cp != g_DerefChar || literal_map[cp - action_args]) // not a proper deref char.
								max_params_override = 3;
							// else since it does start with a real deref symbol, it must end with one otherwise
							// that will be caught later on as a syntax error anyway.  Therefore, don't override
							// max_params, just let it be parsed as 4 parameters.
						}
						// If it has more than 4 params or it has exactly 4 but the 4th isn't blank,
						// pure numeric, or a deref: assume it's being used in 3-parameter mode and
						// that all the other delimiters were intended to be literal.
					}
					else if (delimiter_count > 3) // i.e. 4 or more delimiters, which means 5 or more params.
						// Since it has too many delimiters to be 4-param mode, Assume it's 3-param mode
						// so that non-escaped commas in parameters 4 and beyond will be all treated as
						// strings that are part of parameter #3.
						max_params_override = 3;
					//else if 3 params or less: Don't override via max_params_override, just parse it normally.
				}
			}
		}
	}


	////////////////////////////////////////////////////////////
	// Parse the parmeter string into a list of separate params.
	////////////////////////////////////////////////////////////
	// MaxParams has already been verified as being <= MAX_ARGS.
	// Any g_delimiter-delimited items beyond MaxParams will be included in a lump inside the last param:
	int nArgs, nArgs_plus_one;
	char *arg[MAX_ARGS], *arg_map[MAX_ARGS];
	ActionTypeType subaction_type = ACT_INVALID; // Must init these.
	ActionTypeType suboldaction_type = OLD_INVALID;
	char subaction_name[MAX_VAR_NAME_LENGTH + 1], *subaction_end_marker = NULL, *subaction_start = NULL;
	int max_params = max_params_override ? max_params_override : this_action->MaxParams;
	int max_params_minus_one = max_params - 1;
	bool in_quotes;
	ActionTypeType *np;
	for (nArgs = mark = 0; action_args[mark] && nArgs < max_params; ++nArgs)
	{
		if (nArgs == 2) // i.e. the 3rd arg is about to be added.
		{
			switch (action_type) // will be ACT_INVALID if this_action is an old-style command.
			{
			case ACT_IFWINEXIST:
			case ACT_IFWINNOTEXIST:
			case ACT_IFWINACTIVE:
			case ACT_IFWINNOTACTIVE:
				subaction_start = action_args + mark;
				if (subaction_end_marker = ParseActionType(subaction_name, subaction_start, false))
					if (   (subaction_type = ConvertActionType(subaction_name)) == ACT_INVALID   )
						suboldaction_type = ConvertOldActionType(subaction_name);
				break;
			}
			if (subaction_type || suboldaction_type)
				// A valid command was found (i.e. AutoIt2-style) in place of this commands Exclude Title
				// parameter, so don't add this item as a param to the command.
				break;
		}
		arg[nArgs] = action_args + mark;
		arg_map[nArgs] = literal_map + mark;
		if (nArgs == max_params_minus_one)
		{
			// Don't terminate the last param, just put all the rest of the line
			// into it.  This avoids the need for the user to escape any commas
			// that may appear in the last param.  i.e. any commas beyond this
			// point can't be delimiters because we've already reached MaxArgs
			// for this command:
			++nArgs;
			break;
		}
		// Find the end of the above arg:
		for (in_quotes = false; action_args[mark]; ++mark)
		{
			// The simple method below is sufficient for our purpose even if a quoted string contains
			// pairs of double-quotes to represent a single literal quote, e.g. "quoted ""word""".
			// In other words, it relies on the fact that there must be an even number of quotes
			// inside any mandatory-numeric arg that is an expressions such as x=="red,blue"
			if (action_args[mark] == '"')
				in_quotes = !in_quotes;
			if (action_args[mark] == g_delimiter && !literal_map[mark])  // A non-literal delimiter (unless its within double-quotes of a mandatory-numeric arg) is a match.
			{
				// If we're inside a pair of quotes and this arg is mandatory-numeric, this delimiter
				// is part of the quoted string and thus not a delimiter:
				if (in_quotes && (np = g_act[action_type].NumericParams)) // This command has at least one numeric parameter.
				{
					// As of v1.0.25, pure numeric parameters can optionally be numeric expressions, so check for that:
					nArgs_plus_one = nArgs + 1;
					for (; *np; ++np)
						if (*np == nArgs_plus_one) // This arg is enforced to be purely numeric.
							break;
					if (*np) // Match found, so this is a purely numeric arg.
						continue; // This delimiter is disqualified, so look for the next one.
				}
				// Since above didn't "continue", this is a real delimiter.
				action_args[mark] = '\0';  // Terminate the previous arg.
				// Trim any whitespace from the previous arg.  This operation
				// will not alter the contents of anything beyond action_args[i],
				// so it should be safe.  In addition, even though it changes
				// the contents of the arg[nArgs] substring, we don't have to
				// update literal_map because the map is still accurate due
				// to the nature of rtrim).  UPDATE: Note that this version
				// of rtrim() specifically avoids trimming newline characters,
				// since the user may have included literal newlines at the end
				// of the string by using an escape sequence:
				rtrim(arg[nArgs]);
				// Omit the leading whitespace from the next arg:
				for (++mark; IS_SPACE_OR_TAB(action_args[mark]); ++mark);
				// Now <mark> marks the end of the string, the start of the next arg,
				// or a delimiter-char (if the next arg is blank).
				break;  // Arg was found, so let the outer loop handle it.
			}
		}
	}

	///////////////////////////////////////////////////////////////////////////////
	// Ensure there are sufficient parameters for this command.  Note: If MinParams
	// is greater than 0, the param numbers 1 through MinParams are required to be
	// non-blank.
	///////////////////////////////////////////////////////////////////////////////
	char error_msg[1024];
	if (nArgs < this_action->MinParams)
	{
		snprintf(error_msg, sizeof(error_msg), "\"%s\" requires at least %d parameter%s."
			, this_action->Name, this_action->MinParams
			, this_action->MinParams > 1 ? "s" : "");
		return ScriptError(error_msg, aLineText);
	}
	for (int i = 0; i < this_action->MinParams; ++i) // It's only safe to do this after the above.
		if (!*arg[i])
		{
			snprintf(error_msg, sizeof(error_msg), "\"%s\" requires that parameter #%u be non-blank."
				, this_action->Name, i + 1);
			return ScriptError(error_msg, aLineText);
		}

	////////////////
	//
	////////////////
	if (old_action_type != OLD_INVALID)
	{
		switch(old_action_type)
		{
		case OLD_LEFTCLICK:
		case OLD_RIGHTCLICK:
			// Insert an arg at the beginning of the list to indicate the mouse button.
			arg[2] = arg[1];  arg_map[2] = arg_map[1];
			arg[1] = arg[0];  arg_map[1] = arg_map[0];
			arg[0] = old_action_type == OLD_LEFTCLICK ? "Left" : "Right";  arg_map[0] = NULL;
			return AddLine(ACT_MOUSECLICK, arg, ++nArgs, arg_map);
		case OLD_LEFTCLICKDRAG:
		case OLD_RIGHTCLICKDRAG:
			// Insert an arg at the beginning of the list to indicate the mouse button.
			arg[4] = arg[3];  arg_map[4] = arg_map[3]; // Set the 5th arg to be the 4th, etc.
			arg[3] = arg[2];  arg_map[3] = arg_map[2];
			arg[2] = arg[1];  arg_map[2] = arg_map[1];
			arg[1] = arg[0];  arg_map[1] = arg_map[0];
			arg[0] = (old_action_type == OLD_LEFTCLICKDRAG) ? "Left" : "Right";  arg_map[0] = NULL;
			return AddLine(ACT_MOUSECLICKDRAG, arg, ++nArgs, arg_map);
		case OLD_HIDEAUTOITWIN:
			// This isn't a perfect mapping because the word "on" or "off" might be contained
			// in a variable reference, in which case this conversion will be incorrect.
			// However, variable ref. is exceedingly rare.
			arg[1] = stricmp(arg[0], "On") ? "Icon" : "NoIcon";
			arg[0] = "Tray"; // Assign only after we're done using the old arg[0] value above.
			return AddLine(ACT_MENU, arg, 2, arg_map);
		case OLD_REPEAT:
			if (AddLine(ACT_REPEAT, arg, nArgs, arg_map) != OK)
				return FAIL;
			// For simplicity, always enclose repeat-loop's contents in in a block rather
			// than trying to detect if it has only one line:
			return AddLine(ACT_BLOCK_BEGIN);
		case OLD_ENDREPEAT:
			return AddLine(ACT_BLOCK_END);
		case OLD_WINGETACTIVETITLE:
			arg[nArgs] = "A";  arg_map[nArgs] = NULL; // "A" signifies the active window.
			++nArgs;
			return AddLine(ACT_WINGETTITLE, arg, nArgs, arg_map);
		case OLD_WINGETACTIVESTATS:
		{
			// Convert OLD_WINGETACTIVESTATS into *two* new commands:
			// Command #1: WinGetTitle, OutputVar, A
			char *width = arg[1];  // Temporary placeholder.
			arg[1] = "A";  arg_map[1] = NULL;  // Signifies the active window.
			if (AddLine(ACT_WINGETTITLE, arg, 2, arg_map) != OK)
				return FAIL;
			// Command #2: WinGetPos, XPos, YPos, Width, Height, A
			// Reassign args in the new command's ordering.  These lines must occur
			// in this exact order for the copy to work properly:
			arg[0] = arg[3];  arg_map[0] = arg_map[3];  // xpos
			arg[3] = arg[2];  arg_map[3] = arg_map[2];  // height
			arg[2] = width;   arg_map[2] = arg_map[1];  // width
			arg[1] = arg[4];  arg_map[1] = arg_map[4];  // ypos
			arg[4] = "A";  arg_map[4] = NULL;  // "A" signifies the active window.
			return AddLine(ACT_WINGETPOS, arg, 5, arg_map);
		}

		case OLD_SETENV:
			return AddLine(ACT_ASSIGN, arg, nArgs, arg_map);
		case OLD_ENVADD:
			return AddLine(ACT_ADD, arg, nArgs, arg_map);
		case OLD_ENVSUB:
			return AddLine(ACT_SUB, arg, nArgs, arg_map);
		case OLD_ENVMULT:
			return AddLine(ACT_MULT, arg, nArgs, arg_map);
		case OLD_ENVDIV:
			return AddLine(ACT_DIV, arg, nArgs, arg_map);

		// For these, break rather than return so that further processing can be done:
		case OLD_IFEQUAL:
			action_type = ACT_IFEQUAL;
			break;
		case OLD_IFNOTEQUAL:
			action_type = ACT_IFNOTEQUAL;
			break;
		case OLD_IFGREATER:
			action_type = ACT_IFGREATER;
			break;
		case OLD_IFGREATEROREQUAL:
			action_type = ACT_IFGREATEROREQUAL;
			break;
		case OLD_IFLESS:
			action_type = ACT_IFLESS;
			break;
		case OLD_IFLESSOREQUAL:
			action_type = ACT_IFLESSOREQUAL;
			break;
#ifdef _DEBUG
		default:
			return ScriptError("Unhandled Old-Command.", action_name);
#endif
		} // switch()
	}

	//////////////////////////////////////////////////////////////////////////////////////////////////
	// Handle AutoIt2-style IF-statements (i.e. the IF's action is on the same line as the condition).
	//////////////////////////////////////////////////////////////////////////////////////////////////
	// The check below: Don't bother if this IF (e.g. IfWinActive) has zero params or if the
	// subaction was already found above:
	if (nArgs && !subaction_type && !suboldaction_type && ACT_IS_IF_OLD(action_type, old_action_type))
	{
		char *delimiter;
		char *last_arg = arg[nArgs - 1];
		for (mark = (int)(last_arg - action_args); action_args[mark]; ++mark)
		{
			if (action_args[mark] == g_delimiter && !literal_map[mark])  // Match found: a non-literal delimiter.
			{
				delimiter = action_args + mark; // save the location of this delimiter
				// Omit the leading whitespace from the next arg:
				for (++mark; IS_SPACE_OR_TAB(action_args[mark]); ++mark);
				// Now <mark> marks the end of the string, the start of the next arg,
				// or a delimiter-char (if the next arg is blank).
				subaction_start = action_args + mark;
				if (subaction_end_marker = ParseActionType(subaction_name, subaction_start, false))
				{
					if (   (subaction_type = ConvertActionType(subaction_name)) == ACT_INVALID   )
						suboldaction_type = ConvertOldActionType(subaction_name);
					if (subaction_type || suboldaction_type) // A valid sub-action (command) was found.
					{
						// Remove this subaction from its parent line; we want it separate:
						*delimiter = '\0';
						rtrim(last_arg);
					}
					// else leave it as-is, i.e. as part of the last param, because the delimiter
					// found above is probably being used as a literal char even though it isn't
					// escaped, e.g. "ifequal, var1, string with embedded, but non-escaped, commas"
				}
				// else, do nothing; reasoning perhaps similar to above comment.
				break;
			}
		}
	}

	if (AddLine(action_type, arg, nArgs, arg_map) != OK)
		return FAIL;
	if (!subaction_type && !suboldaction_type) // There is no subaction in this case.
		return OK;
	// Otherwise, recursively add the subaction, and any subactions it might have, beneath
	// the line just added.  The following example:
	// IfWinExist, x, y, IfWinNotExist, a, b, Gosub, Sub1
	// would break down into these lines:
	// IfWinExist, x, y
	//    IfWinNotExist, a, b
	//       Gosub, Sub1
	return ParseAndAddLine(subaction_start, subaction_name, subaction_end_marker
		, literal_map + (subaction_end_marker - action_args) // Pass only the relevant substring of literal_map.
		, strlen(subaction_end_marker), subaction_type, suboldaction_type);
}



inline char *Script::ParseActionType(char *aBufTarget, char *aBufSource, bool aDisplayErrors)
// inline since it's called so often.
// aBufTarget should be at least MAX_VAR_NAME_LENGTH + 1 in size.
// Returns NULL if a failure condition occurs; otherwise, the address of the last
// character of the action name in aBufSource.
{
	////////////////////////////////////////////////////////
	// Find the action name and the start of the param list.
	////////////////////////////////////////////////////////
	// Allows the delimiter between action-type-name and the first param to be optional by
	// relying on the fact that action-type-names can't contain spaces. Find first char in
	// aLineText that is a space, a delimiter, or a tab. Also search for operator symbols
	// so that assignments and IFs without whitespace are supported, e.g. var1=5,
	// if var2<%var3%.  Not static in case g_delimiter is allowed to vary:
	DEFINE_END_FLAGS
	char *end_marker = StrChrAny(aBufSource, end_flags);
	if (end_marker) // Found a delimiter.
	{
		if (end_marker > aBufSource) // The delimiter isn't very first char in aBufSource.
			--end_marker;
		// else we allow it to be the first char to support "++i" etc.
	}
	else // No delimiter found, so set end_marker to the location of the last char in string.
		end_marker = aBufSource + strlen(aBufSource) - 1;
	// Now end_marker is the character just prior to the first delimiter or whitespace,
	// or (in the case of ++ and --) the first delimiter itself.  Find the end of
	// the action-type name by omitting trailing whitespace:
	end_marker = omit_trailing_whitespace(aBufSource, end_marker);
	// If first char in aBufSource is a delimiter, action_name will consist of just that first char:
	size_t action_name_length = end_marker - aBufSource + 1;
	if (action_name_length > MAX_VAR_NAME_LENGTH)
	{
		if (aDisplayErrors)
			ScriptError("The first word in this line is too long to be any valid command or variable name."
				, aBufSource);
		return NULL;
	}
	strlcpy(aBufTarget, aBufSource, action_name_length + 1);
	return end_marker;
}



inline ActionTypeType Script::ConvertActionType(char *aActionTypeString)
// inline since it's called so often, but don't keep it in the .h due to #include issues.
{
	// For the loop's index:
	// Use an int rather than ActionTypeType since it's sure to be large enough to go beyond
	// 256 if there happen to be exactly 256 actions in the array:
 	for (int action_type = ACT_FIRST_COMMAND; action_type < g_ActionCount; ++action_type)
		if (!stricmp(aActionTypeString, g_act[action_type].Name)) // Match found.
			return action_type;
	return ACT_INVALID;  // On failure to find a match.
}



inline ActionTypeType Script::ConvertOldActionType(char *aActionTypeString)
// inline since it's called so often, but don't keep it in the .h due to #include issues.
{
 	for (int action_type = OLD_INVALID + 1; action_type < g_OldActionCount; ++action_type)
		if (!stricmp(aActionTypeString, g_old_act[action_type].Name)) // Match found.
			return action_type;
	return OLD_INVALID;  // On failure to find a match.
}



ResultType Script::AddLine(ActionTypeType aActionType, char *aArg[], ArgCountType aArgc, char *aArgMap[])
// aArg must be a collection of pointers to memory areas that are modifiable, and there
// must be at least MAX_ARGS number of pointers in the aArg array.
// Returns OK or FAIL.
{
	if (aActionType == ACT_INVALID)
		return ScriptError("BAD AddLine", aArgc > 0 ? aArg[0] : "");

	//////////////////////////////////////////////////////////
	// Build the new arg list in dynamic memory.
	// The allocated structs will be attached to the new line.
	//////////////////////////////////////////////////////////
	Var *target_var;
	DerefType deref[MAX_DEREFS_PER_ARG];  // Will be used to temporarily store the var-deref locations in each arg.
	int deref_count;  // How many items are in deref array.
	ArgStruct *new_arg;  // We will allocate some dynamic memory for this, then hang it onto the new line.
	size_t operand_length;
	char *op_begin, *op_end, orig_char;
	char *this_aArgMap, *this_aArg, *cp;
	int open_parens;
	ActionTypeType *np;

	if (!aArgc)
		new_arg = NULL;  // Just need an empty array in this case.
	else
	{
		if (   !(new_arg = (ArgStruct *)SimpleHeap::Malloc(aArgc * sizeof(ArgStruct)))   )
			return ScriptError(ERR_OUTOFMEM);
		int i, j, i_plus_one;
		for (i = 0; i < aArgc; ++i)
		{
			////////////////
			// FOR EACH ARG:
			////////////////
			this_aArg = aArg[i];                        // For performance and convenience.
			this_aArgMap = aArgMap ? aArgMap[i] : NULL; // Same.
			ArgStruct &this_new_arg = new_arg[i];       // Same.
			this_new_arg.is_expression = false;         // Set default early, for maintainability.

			// Before allocating memory for this Arg's text, first check if it's a pure
			// variable.  If it is, we store it differently (and there's no need to resolve
			// escape sequences in these cases, since var names can't contain them):
			if (aActionType == ACT_LOOP && i == 1 && aArg[0] && !stricmp(aArg[0], "Parse")) // Verified.
				// i==1 --> 2nd arg's type is based on 1st arg's text.
				this_new_arg.type = ARG_TYPE_INPUT_VAR;
			else
				this_new_arg.type = Line::ArgIsVar(aActionType, i);
			// Since some vars are optional, the below allows them all to be blank or
			// not present in the arg list.  If a mandatory var is blank at this stage,
			// it's okay because all mandatory args are validated to be non-blank elsewhere:
			if (this_new_arg.type != ARG_TYPE_NORMAL)
			{
				if (!*this_aArg)
					// An optional input or output variable has been omitted, so indicate
					// that this arg is not a variable, just a normal empty arg.  Functions
					// such as ListLines() rely on this having been done because they assume,
					// for performance reasons, that args marked as variables really are
					// variables.  In addition, ExpandArgs() relies on this having been done
					// as does the load-time validation for ACT_DRIVEGET:
					this_new_arg.type = ARG_TYPE_NORMAL;
				else
				{
					// Does this input or output variable contain a dereference?  If so, it must
					// be resolved at runtime (to support old-style AutoIt2 arrays, etc.).
					// Find the first non-escaped dereference symbol:
					for (j = 0; this_aArg[j] && (this_aArg[j] != g_DerefChar || (this_aArgMap && this_aArgMap[j])); ++j);
					if (!this_aArg[j])
					{
						// A non-escaped deref symbol wasn't found, therefore this variable does not
						// appear to be something that must be resolved dynamically at runtime.
						if (   !(target_var = FindOrAddVar(this_aArg))   )
							return FAIL;  // The above already displayed the error.
						// If this action type is something that modifies the contents of the var, ensure the var
						// isn't a special/reserved one:
						if (this_new_arg.type == ARG_TYPE_OUTPUT_VAR && VAR_IS_RESERVED(target_var))
							return ScriptError(ERR_VAR_IS_RESERVED, this_aArg);
						// Rather than removing this arg from the list altogether -- which would distrub
						// the ordering and hurt the maintainability of the code -- the next best thing
						// in terms of saving memory is to store an empty string in place of the arg's
						// text if that arg is a pure variable (i.e. since the name of the variable is already
						// stored in the Var object, we don't need to store it twice):
						this_new_arg.text = "";
						this_new_arg.deref = (DerefType *)target_var;
						continue;
					}
					// else continue on to the below so that this input or output variable name's dynamic part
					// (e.g. array%i%) can be partially resolved.
				}
			}

			// Below will set the new var to be the constant empty string if the
			// source var is NULL or blank.
			// e.g. If WinTitle is unspecified (blank), but WinText is non-blank.
			// Using empty string is much safer than NULL because these args
			// will be frequently accessed by various functions that might
			// not be equipped to handle NULLs.  Rather than having to remember
			// to check for NULL in every such case, setting it to a constant
			// empty string here should make things a lot more maintainable
			// and less bug-prone.  If there's ever a need for the contents
			// of this_new_arg to be modifiable (perhaps some obscure API calls require
			// modifiable strings?) can malloc a single-char to contain the empty string:
			if (   !(this_new_arg.text = SimpleHeap::Malloc(this_aArg))   )
				return FAIL;  // It already displayed the error for us.

			////////////////////////////////////////////////////
			// Build the list of dereferenced vars for this arg.
			////////////////////////////////////////////////////
			// Now that any escaped g_DerefChars have been marked, scan new_arg.text to
			// determine where the variable dereferences are (if any).  In addition to helping
			// runtime performance, this also serves to validate the script at load-time
			// so that some errors can be caught early.  Note: this_new_arg.text is scanned rather
			// than this_aArg because we want to establish pointers to the correct area of
			// memory:
			deref_count = 0;  // Init for each arg.

			if (np = g_act[aActionType].NumericParams) // This command has at least one numeric parameter.
			{
				// As of v1.0.25, pure numeric parameters can optionally be numeric expressions, so check for that:
				i_plus_one = i + 1;
				for (; *np; ++np)
				{
					if (*np == i_plus_one) // This arg is enforced to be purely numeric.
					{
						if (aActionType == ACT_WINMOVE)
						{
							if (i > 1)
							{
								// i indicates this is Arg #3 or beyond, which is one of the args that is
								// either the word "default" or a number/expression.
								if (!stricmp(this_new_arg.text, "default")) // It's not an expression.
									break; // The loop is over because this arg was found in the list.
							}
							else // This is the first or second arg, which are title/text vs. X/Y when aArgc > 2.
								if (aArgc > 2) // Title/text are not numeric/expressions.
									break; // The loop is over because this arg was found in the list.
						}
						// To help runtime performance, the below leaves an ACT_ASSIGNEXPR marked as a
						// non-expression if its expression consists of only a single literal number.
						// At runtime, such args are expanded normally rather than having to run them
						// through the expression evaluator:
						if (IsPureNumeric(this_new_arg.text, true, true, true))
							break; // The loop is over because this arg was found in the list.
						// Otherwise, it might be an expression so do the final checks.
						// Override the original false default of is_expression unless an exception applies.
						// Since ACT_ASSIGNEXPR is not a legacy command, none of the legacy exceptions need
						// to be applied to it.  For other commands, if any telltale character is present
						// it's definitely an expression and the complex check after this one isn't needed:
						if (aActionType == ACT_ASSIGNEXPR || StrChrAny(this_new_arg.text, EXP_TELLTALES))
							this_new_arg.is_expression = true;
						else
						{
							// The section below is here in light of rare legacy cases such as the below:
							// -%y%   ; i.e. make it negative.
							// +%y%   ; might happen with up/down adjustments on SoundSet, GuiControl progress/slider, etc?
							// Although the above are detected as non-expressions and thus non-double-derefs,
							// the following are not because they're too rare or would sacrifice too much flexibility:
							// 1%y%.0 ; i.e. at a tens/hundreds place and make it a floating point.  In addition,
							//          1%y% could be an array, so best not to tag that as non-expression.
							//          For that matter, %y%.0 could be an obscure kind of reverse-notation array itself.
							// 0x%y%  ; i.e. make it hex (too rare to check for, plus it could be an array).
							// %y%%z% ; i.e. concatenate two numbers to make a larger number (too rare to check for)
							cp = this_new_arg.text + (*this_new_arg.text == '-' || *this_new_arg.text == '+'); // i.e. +1 if second term evaluates to true.
							if (*cp != g_DerefChar // If no deref, for simplicity assume it's an expression since any such non-numeric item would be extremely rare in pre-expression era.
								|| !this_aArgMap || *(this_aArgMap + (cp != this_new_arg.text)) // There's no literal-map or this deref char is not really a deref char because it's marked as a literal.
								|| !(cp = strchr(cp + 1, g_DerefChar)) // There is no next deref char.
								|| *(cp + 1)) // But that next deref char is not the last deref char, which means this is not a single isolated deref.
								// Above does not need to check whether last deref char is marked literal in the
								// arg map because if it is, it would mean the first deref char lacks a matching
								// close-symbol, which will be caught as a syntax error below regardless of whether
								// this is an expression.
								this_new_arg.is_expression = true;
						}
						break; // The loop is over if this arg is found in the list of mandatory-numeric args.
					} // i is a mandatory-numeric arg
				} // for each mandatory-numeric arg of this command, see if this arg matches its number.
			} // this command has a list of mandatory numeric-args.

			if (this_new_arg.is_expression)
			{
				// Ensure parentheses are balanced:
				for (cp = this_new_arg.text, open_parens = 0; *cp; ++cp)
				{
					if (*cp == '(')
						++open_parens;
					else if (*cp == ')')
					{
						if (!open_parens)
							return ScriptError("Close-paren with no open-paren.", cp); // And indicate cp as the exact spot.
						--open_parens;
					}
				}
				if (open_parens) // At least one open-paren is never closed.
					return ScriptError("Missing \")\"", this_new_arg.text);

				// ParseDerefs() won't consider escaped percent signs to be illegal, but in this case
				// they should be since they have no meaning in expressions:
				#define ERR_EXP_ILLEGAL_CHAR "The first character above is illegal in an expression." // "above" refers to the layout of the error dialog.
				if (this_aArgMap) // This arg has an arg map indicating which chars are escaped/literal vs. normal.
					for (j = 0; this_new_arg.text[j]; ++j)
						if (this_aArgMap[j] && this_new_arg.text[j] == g_DerefChar)
							return ScriptError(ERR_EXP_ILLEGAL_CHAR, this_new_arg.text + j);

				// Resolve all operands that aren't numbers into variable references.  Doing this here at
				// load-time greatly improves runtime performance, especially for scripts that have a lot
				// of variables.
				for (op_begin = this_new_arg.text; *op_begin; op_begin = op_end)
				{
					for (; *op_begin && strchr(EXP_OPERAND_TERMINATORS, *op_begin); ++op_begin); // Skip over whitespace, operators, and parens.
					if (!*op_begin) // The above loop reached the end of the string: No operands remaining.
						break;

					// Now op_begin is the start of an operand, which might be a variable reference, a numeric
					// literal, or a string literal.  If it's a string literal, it is left as-is:
					if (*op_begin == '"')
					{
						// Find the end of this string literal, noting that a pair of double quotes is
						// a literal double quote inside the string:
						for (op_end = op_begin + 1;; ++op_end)
						{
							if (!*op_end)
								return ScriptError("Missing close-quote.", op_begin);
							if (*op_end == '"') // If not followed immediately by another, this is the end of it.
							{
								++op_end;
								if (*op_end != '"') // String terminator or some non-quote character.
									break;  // The previous char is the ending quote.
								//else a pair of quotes, which resolves to a single literal quote.
								// This pair is skipped over and the loop continues until the real end-quote is found.
							}
						}
						// op_end is now set correctly to allow the outer loop to continue.
						continue; // Ignore this literal string, letting the runtime expression parser recognize it.
					}
					
					// Find the end of this operand (if *op_end is '\0', strchr() will find that too):
					for (op_end = op_begin + 1; !strchr(EXP_OPERAND_TERMINATORS, *op_end); ++op_end); // Find first whitespace, operator, or paren.
					// Now op_end marks the end of this operand.  The end might be the zero terminator, an operator, etc.
					operand_length = op_end - op_begin;

					// Check if it's AND/OR/NOT:
					if (operand_length < 4 && operand_length > 1) // Ordered for short-circuit performance.
					{
						if (operand_length == 2)
						{
							if ((*op_begin == 'o' || *op_begin == 'O') && (op_begin[1] == 'r' || op_begin[1] == 'R'))
								continue; // "OR" was found.
						}
						else // operand_length must be 3
						{
							switch (*op_begin)
							{
							case 'a':
							case 'A':
								if (   (op_begin[1] == 'n' || op_begin[1] == 'N') // Relies on short-circuit boolean order.
									&& (op_begin[2] == 'd' || op_begin[2] == 'D')   )
									continue; // "AND" was found.
								break;

							case 'n':
							case 'N':
								if (   (op_begin[1] == 'o' || op_begin[1] == 'O') // Relies on short-circuit boolean order.
									&& (op_begin[2] == 't' || op_begin[2] == 'T')   )
									continue; // "NOT" was found.
								break;
							}
						}
					}
					// Temporarily terminate, which avoids at least the below issue:
					// Two or more extremely long var names together could exceed MAX_VAR_NAME_LENGTH
					// e.g. LongVar%LongVar2% would be too long to store in a buffer of size MAX_VAR_NAME_LENGTH.
					// This seems pretty darn unlikely, but perhaps doubling it would be okay.
					// UPDATE: Above is now not an issue since caller's string is temporarily terminated rather
					// than making a copy of it.
					orig_char = *op_end;
					*op_end = '\0';

					// Illegal characters are legal when enclosed in double quotes.  So the following is
					// done only after the above has ensured this operand is one enclosed entirely in
					// double quotes.
					// The following characters are either illegal in expressions or reserved for future use.
					// Rather than forbidding g_delimiter and g_DerefChar, it seems best to assume they are at
					// their default values for this purpose.  Otherwise, if g_delimiter is an operator, that
					// operator would then become impossible inside the expression.
					if (cp = StrChrAny(op_begin, EXP_ILLEGAL_CHARS))
						return ScriptError(ERR_EXP_ILLEGAL_CHAR, cp);

					// Below takes care of recognizing hexadecimal integers, which avoids the 'x' character
					// inside of something like 0xFF from being detected as the name of a variable:
					if (!IsPureNumeric(op_begin, true, false, true))
					{
						// This operand must be a variable reference or string literal, otherwise it's a syntax error.
						// Check explicitly for derefs since the vast majority don't have any, and this
						// avoids the function call in those cases:
						if (strchr(op_begin, g_DerefChar)) // This operand contains at least one variable reference.
						{
							// The derefs are parsed and added to the deref array at this stage (on a
							// per-operand basis) rather than all at once for the entire arg because
							// the deref array must be ordered according to the physical position of
							// derefs inside the arg.  In the following example, the order of derefs
							// must be x,i,y: if (x = Array%i% and y = 3)
							if (!ParseDerefs(op_begin, this_aArgMap + (op_begin - this_new_arg.text), deref, deref_count))
								return FAIL; // It already displayed the error.  No need to undo temp. termination.
							// And now leave this operand "raw" so that it will later be dereferenced again.
							// In the following example, i made into a deref but the result (Array33) must be
							// dereferenced during a second stage at runtime: if (x = Array%i%).
						}
						else // This operand is a variable name.
						{
							#define TOO_MANY_VAR_REFS "Too many var refs." // Short msg since so rare.
							if (deref_count >= MAX_DEREFS_PER_ARG)
								return ScriptError(TOO_MANY_VAR_REFS, op_begin); // Indicate which operand it ran out of space at.
							deref[deref_count].marker = op_begin;  // Store the deref's starting location.
							deref[deref_count].length = (DerefLengthType)operand_length;
							if (   !(deref[deref_count].var = FindOrAddVar(op_begin, operand_length))   )
								return FAIL; // The called function already displayed the error.
							++deref_count;
						}
					}
					//else purely numeric.  Do nothing since pure numbers don't need any processing at this stage.
					*op_end = orig_char; // Undo the temporary termination.
				} // expression pre-parsing loop.

				// Now that the derefs have all been recognized above, simplify any special cases --
				// such as single isolated derefs -- to enhance runtime performance.
				// Make args that consist only of a quoted string-literal into non-expressions also:
				if (!deref_count && *this_new_arg.text == '"')
				{
					// It has no derefs (e.g. x:="string" or x:=1024*1024), but since it's a single
					// string literal, convert into a non-expression.  This is mainly for use by
					// ACT_ASSIGNEXPR, but it seems slightly beneficial for other things in case
					// they ever use quoted numeric ARGS such as "13", etc.  It's also simpler
					// to do it unconditionally.
					// Find the end of this string literal, noting that a pair of double quotes is
					// a literal double quote inside the string:
					for (cp = this_new_arg.text + 1;; ++cp)
					{
						if (!*cp) // No matching end-quote. Probably impossible due to validation further above.
							return FAIL; // Force a silent failure so that the below can continue with confidence.
						if (*cp == '"') // If not followed immediately by another, this is the end of it.
						{
							++cp;
							if (*cp != '"') // String terminator or some non-quote character.
								break;  // The previous char is the ending quote.
							//else a pair of quotes, which resolves to a single literal quote.
							// This pair is skipped over and the loop continues until the real end-quote is found.
						}
					}
					// cp is now the character after the first literal string's ending quote.
					// If that char is the terminator, that first string is the only string and this
					// is a simple assignment of a string literal to be converted here.
					if (!*cp)
					{
						this_new_arg.is_expression = false;
						// Bugfix for 1.0.25.06: The below has been disabled because:
						// 1) It yields inconsistent results due to AutoTrim.  For example, the assignment
						//    x := "  string" should retain the leading spaces unconditionally, since any
						//    more complex expression would.  But if := were converted to = in this case,
						//    AutoTrim would be in effect for it, which is undesirable.
						// 2) It's not necessary in since ASSIGNEXPR handles both expressions and non-expressions.
						//if (aActionType == ACT_ASSIGNEXPR)
						//	aActionType = ACT_ASSIGN; // Convert to simple assignment.
						*(--cp) = '\0'; // Remove the ending quote.
						memmove(this_new_arg.text, this_new_arg.text + 1, cp - this_new_arg.text); // Remove the starting quote.
						// Convert all pairs of quotes into single literal quotes:
						StrReplaceAll(this_new_arg.text, "\"\"", "\"", true);
						// Above relies on the fact that StrReplaceAll() does not do cascading replacements,
						// meaning that a series of characters such as """" would be correctly converted into
						// two double quotes rather than just collapsing into one.
					}
				}
				// Make things like "Sleep Var" and "Var := X" into non-expressions.  At runtime,
				// such args are expanded normally rather than having to run them through the
				// expression evaluator.  A simple test script shows that this one change can
				// double the runtime performance of certain commands such as EnvAdd:
				// Below is somewhat obsolete but kept for reference:
				// This policy is basically saying that expressions are allowed to evaluate to strings
				// everywhere appropriate, but that at the moment the only appropriate place is x := y
				// because all other expressions should resolve to a numeric value by virtue of the fact
				// that they *are* numeric parameters.
				else if (deref_count == 1 && Var::ValidateName(this_new_arg.text, false, false)) // Single isolated deref.
					this_new_arg.is_expression = false;
				else if (deref_count && !StrChrAny(this_new_arg.text, EXP_OPERAND_TERMINATORS))
				{
					// Adjust if any of the following special cases apply:
					// x := y  -> Mark as non-expression (after expression-parsing set up parsed derefs above)
					//            so that the y deref will be only a single-deref to be directly stored in x.
					//            This is done in case y contains a string.  Since an expression normally
					//            evaluates to a number, without this workaround, x := y would be useless for
					//            a simple assignment of a string.  This case is handled above.
					// x := %y% -> Mark the right-side arg as an input variable so that it will be doubly
					//             dereferenced, similar to StringTrimRight, Out, %y%, 0.  This seems best
					//             because there would be little or no point to having it behave identically
					//             to x := y.  It might even be confusing in light of the next case below.
					// CASE #3:
					// x := Literal%y%Literal%z%Literal -> Same as above.  This is done mostly to support
					// retrieving array elements whose contents are *non-numeric* without having to use
					// something like StringTrimRight.
					
					// Now we know it has at least one deref.  But if any operators or other characters disallowed
					// in variables are present, it all three cases are disqualified and kept as expressions.
					// This check is necessary for all three cases:

					// No operators of any kind anywhere.  Not even +/- prefix, since those imply a numeric
					// expression.  No chars illegal in var names except the percent signs themselves,
					// e.g. *no* whitespace.
					// Also, the first deref (indeed, all of them) should point to a percent sign, since
					// there should not be any way for non-percent derefs to get mixed in with cases
					// 2 or 3.
					if (*deref[0].marker == g_DerefChar) // This appears to be case #2 or #3.
					{
						// Set it up so that x:=Array%i% behaves the same as StringTrimRight, Out, Array%i%, 0.
						this_new_arg.is_expression = false;
						this_new_arg.type = ARG_TYPE_INPUT_VAR;
					}
				}
			}
			else // this arg does not contain an expression.
				if (!ParseDerefs(this_new_arg.text, this_aArgMap, deref, deref_count))
					return FAIL; // It already displayed the error.

			//////////////////////////////////////////////////////////////
			// Allocate mem for this arg's list of dereferenced variables.
			//////////////////////////////////////////////////////////////
			if (deref_count)
			{
				// +1 for the "NULL-item" terminator:
				this_new_arg.deref = (DerefType *)SimpleHeap::Malloc((deref_count + 1) * sizeof(DerefType));
				if (!this_new_arg.deref)
					return ScriptError(ERR_OUTOFMEM);
				for (j = 0; j < deref_count; ++j)
					this_new_arg.deref[j] = deref[j]; // Fortunately, struct copying does not generate a large amount of code.
				// Terminate the list of derefs with a deref that has a NULL marker,
				// but only if the last one added isn't NULL (which it would be if it's
				// the fake-deref used to store an output-parameter's variable):
				if (this_new_arg.deref[deref_count - 1].marker)
					this_new_arg.deref[deref_count].marker = NULL;
					// No need to increment deref_count since it's never used beyond this point.
			}
			else
				this_new_arg.deref = NULL;
		} // for each arg.
	} // else there are more than zero args.

	//////////////////////////////////////////////////////////////////////////////////////
	// Now the above has allocated some dynamic memory, the pointers to which we turn over
	// to Line's constructor so that they can be anchored to the new line.
	//////////////////////////////////////////////////////////////////////////////////////
	Line *line = new Line(mCurrFileNumber, mCurrLineNumber, aActionType, new_arg, aArgc);
	if (line == NULL)
		return ScriptError(ERR_OUTOFMEM);
	line->mPrevLine = mLastLine;  // Whether NULL or not.
	if (mFirstLine == NULL)
		mFirstLine = mLastLine = line;
	else
	{
		mLastLine->mNextLine = line;
		// This must be done after the above:
		mLastLine = line;
	}
	mCurrLine = line;  // To help error reporting.

	///////////////////////////////////////////////////////////////////
	// Do any post-add validation & handling for specific action types.
	///////////////////////////////////////////////////////////////////
	int value;    // For temp use during validation.
	double value_float;
	SYSTEMTIME st;  // same.

	switch(aActionType)
	{
	case ACT_AUTOTRIM:
	case ACT_STRINGCASESENSE:
	case ACT_DETECTHIDDENWINDOWS:
	case ACT_DETECTHIDDENTEXT:
	case ACT_SETSTORECAPSLOCKMODE:
		if (aArgc > 0 && !line->ArgHasDeref(1) && !line->ConvertOnOff(NEW_RAW_ARG1))
			return ScriptError(ERR_ON_OFF, NEW_RAW_ARG1);
		break;

	case ACT_SETBATCHLINES:
		if (aArgc > 0 && !line->ArgHasDeref(1))
		{
			if (!strcasestr(NEW_RAW_ARG1, "ms") && !IsPureNumeric(NEW_RAW_ARG1, true, false))
				return ScriptError("Parameter #1 must be an integer, an interval string such as 15ms,"
					" or a variable reference.", NEW_RAW_ARG1);
		}
		break;

	case ACT_SUSPEND:
		if (aArgc > 0 && !line->ArgHasDeref(1) && !line->ConvertOnOffTogglePermit(NEW_RAW_ARG1))
			return ScriptError(ERR_ON_OFF_TOGGLE_PERMIT, NEW_RAW_ARG1);
		break;

	case ACT_BLOCKINPUT:
		if (aArgc > 0 && !line->ArgHasDeref(1) && !line->ConvertBlockInput(NEW_RAW_ARG1))
			return ScriptError(ERR_BLOCKINPUT, NEW_RAW_ARG1);
		break;

	case ACT_PAUSE:
	case ACT_KEYHISTORY:
		if (aArgc > 0 && !line->ArgHasDeref(1) && !line->ConvertOnOffToggle(NEW_RAW_ARG1))
			return ScriptError(ERR_ON_OFF_TOGGLE, NEW_RAW_ARG1);
		break;

	case ACT_SETNUMLOCKSTATE:
	case ACT_SETSCROLLLOCKSTATE:
	case ACT_SETCAPSLOCKSTATE:
		if (aArgc > 0 && !line->ArgHasDeref(1) && !line->ConvertOnOffAlways(NEW_RAW_ARG1))
			return ScriptError(ERR_ON_OFF_ALWAYS, NEW_RAW_ARG1);
		break;

	case ACT_STRINGMID:
	case ACT_FILEREADLINE:
		if (aArgc > 2 && !line->ArgHasDeref(3))
		{
			if (!IsPureNumeric(NEW_RAW_ARG3, false, false, false) || !ATOI64(NEW_RAW_ARG3))
				// This error is caught at load-time, but at runtime it's not considered
				// an error (i.e. if a variable resolves to zero or less, StringMid will
				// automatically consider it to be 1, though FileReadLine would consider
				// it an error):
				return ScriptError("Parameter #3 must be a number greater than zero or a variable reference."
					, NEW_RAW_ARG3);
		}
		if (aArgc > 4 && !line->ArgHasDeref(5) && stricmp(NEW_RAW_ARG5, "L"))
			return ScriptError("Parameter #5 must be blank, the letter L, or a variable reference.", NEW_RAW_ARG5);
		break;

	case ACT_STRINGGETPOS:
		if (*NEW_RAW_ARG4 && !line->ArgHasDeref(4) && !strchr("LR1", toupper(*NEW_RAW_ARG4)))
			return ScriptError("If not blank or a variable reference, parameter #4 must be 1 or start with the letter L or R.", NEW_RAW_ARG4);
		break;

	case ACT_STRINGSPLIT:
		if (*NEW_RAW_ARG1 && !line->ArgHasDeref(1)) // The output array must be a legal name.
			if (!Var::ValidateName(NEW_RAW_ARG1)) // It already displayed the error.
				return FAIL;
		break;

	case ACT_REGREAD:
		// The below has two checks in case the user is using the 5-param method with the 5th parameter
		// being blank to indicate that the key's "default" value should be read.  For example:
		// RegRead, OutVar, REG_SZ, HKEY_CURRENT_USER, Software\Winamp,
		if (aArgc > 4 || line->RegConvertValueType(NEW_RAW_ARG2))
		{
			// The obsolete 5-param method is being used, wherein ValueType is the 2nd param.
			if (*NEW_RAW_ARG3 && !line->ArgHasDeref(3) && !line->RegConvertRootKey(NEW_RAW_ARG3))
				return ScriptError(ERR_REG_KEY, NEW_RAW_ARG3);
		}
		else // 4-param method.
			if (*NEW_RAW_ARG2 && !line->ArgHasDeref(2) && !line->RegConvertRootKey(NEW_RAW_ARG2))
				return ScriptError(ERR_REG_KEY, NEW_RAW_ARG2);
		break;

	case ACT_REGWRITE:
		// Both of these checks require that at least two parameters be present.  Otherwise, the command
		// is being used in its registry-loop mode and is validated elsewhere:
		if (aArgc > 1)
		{
			if (*NEW_RAW_ARG1 && !line->ArgHasDeref(1) && !line->RegConvertValueType(NEW_RAW_ARG1))
				return ScriptError(ERR_REG_VALUE_TYPE, NEW_RAW_ARG1);
			if (*NEW_RAW_ARG2 && !line->ArgHasDeref(2) && !line->RegConvertRootKey(NEW_RAW_ARG2))
				return ScriptError(ERR_REG_KEY, NEW_RAW_ARG2);
		}
		break;

	case ACT_REGDELETE:
		if (*NEW_RAW_ARG1 && !line->ArgHasDeref(1) && !line->RegConvertRootKey(NEW_RAW_ARG1))
			return ScriptError(ERR_REG_KEY, NEW_RAW_ARG1);
		break;

	case ACT_SOUNDGET:
	case ACT_SOUNDSET:
		if (aActionType == ACT_SOUNDSET && aArgc > 0 && !line->ArgHasDeref(1))
		{
			value_float = ATOF(NEW_RAW_ARG1);
			if (value_float < -100 || value_float > 100)
				return ScriptError(ERR_PERCENT, NEW_RAW_ARG1);
		}
		if (*NEW_RAW_ARG2 && !line->ArgHasDeref(2) && !line->SoundConvertComponentType(NEW_RAW_ARG2))
			return ScriptError("Parameter #2 specifies an invalid component type.", NEW_RAW_ARG2);
		if (*NEW_RAW_ARG3 && !line->ArgHasDeref(3) && !line->SoundConvertControlType(NEW_RAW_ARG3))
			return ScriptError("Parameter #3 specifies an invalid control type.", NEW_RAW_ARG3);
		break;

	case ACT_SOUNDSETWAVEVOLUME:
		if (aArgc > 0 && !line->ArgHasDeref(1))
		{
			value_float = ATOF(NEW_RAW_ARG1);
			if (value_float < -100 || value_float > 100)
				return ScriptError(ERR_PERCENT, NEW_RAW_ARG1);
		}
		break;

	case ACT_SOUNDPLAY:
		if (*NEW_RAW_ARG2 && !line->ArgHasDeref(2) && stricmp(NEW_RAW_ARG2, "wait") && stricmp(NEW_RAW_ARG2, "1"))
			return ScriptError("If not blank, parameter #2 must be 1, WAIT, or a variable reference.", NEW_RAW_ARG2);
		break;

	case ACT_PIXELGETCOLOR:
		if (*NEW_RAW_ARG4 && !line->ArgHasDeref(4) && stricmp(NEW_RAW_ARG4, "RGB"))
			return ScriptError("Parameter #4 must be blank or the word RGB.", NEW_RAW_ARG4);
		break;

	case ACT_PIXELSEARCH:
	//case ACT_IMAGESEARCH:
		if (!*NEW_RAW_ARG3 || !*NEW_RAW_ARG4 || !*NEW_RAW_ARG5 || !*NEW_RAW_ARG6 || !*NEW_RAW_ARG7)
			return ScriptError("Parameters 3 through 7 must not be blank.");
		//if (mActionType != ACT_IMAGESEARCH)
		//{
			if (*NEW_RAW_ARG8 && !line->ArgHasDeref(8))
			{
				value = ATOI(NEW_RAW_ARG8);
				if (value < 0 || value > 255)
					return ScriptError("Parameter #8 must be number between 0 and 255, blank, or a variable reference."
						, NEW_RAW_ARG8);
			}
			if (*NEW_RAW_ARG9 && !line->ArgHasDeref(9) && stricmp(NEW_RAW_ARG9, "RGB"))
				return ScriptError("Parameter #9 must be blank or the word RGB.", NEW_RAW_ARG9);
		//}
		break;

	case ACT_COORDMODE:
		if (*NEW_RAW_ARG1 && !line->ArgHasDeref(1) && !line->ConvertCoordModeAttrib(NEW_RAW_ARG1))
			return ScriptError(ERR_COORDMODE, NEW_RAW_ARG1);
		break;

	case ACT_SETDEFAULTMOUSESPEED:
		if (*NEW_RAW_ARG1 && !line->ArgHasDeref(1))
		{
			value = ATOI(NEW_RAW_ARG1);
			if (value < 0 || value > MAX_MOUSE_SPEED)
				return ScriptError(ERR_MOUSE_SPEED, NEW_RAW_ARG1);
		}
		break;

	case ACT_MOUSEMOVE:
		if (*NEW_RAW_ARG3 && !line->ArgHasDeref(3))
		{
			value = ATOI(NEW_RAW_ARG3);
			if (value < 0 || value > MAX_MOUSE_SPEED)
				return ScriptError(ERR_MOUSE_SPEED, NEW_RAW_ARG3);
		}
		if (*NEW_RAW_ARG4 && !line->ArgHasDeref(4) && toupper(*NEW_RAW_ARG4) != 'R')
			return ScriptError("Parameter #4 (if present) must be the letter R.", NEW_RAW_ARG4);
		if (!line->ValidateMouseCoords(NEW_RAW_ARG1, NEW_RAW_ARG2))
			return ScriptError(ERR_MOUSE_COORD, NEW_RAW_ARG1);
		break;

	case ACT_MOUSECLICK:
		if (*NEW_RAW_ARG5 && !line->ArgHasDeref(5))
		{
			value = ATOI(NEW_RAW_ARG5);
			if (value < 0 || value > MAX_MOUSE_SPEED)
				return ScriptError(ERR_MOUSE_SPEED, NEW_RAW_ARG5);
		}
		if (*NEW_RAW_ARG6 && !line->ArgHasDeref(6))
			if (strlen(NEW_RAW_ARG6) > 1 || !strchr("UD", toupper(*NEW_RAW_ARG6)))  // Up / Down
				return ScriptError(ERR_MOUSE_UPDOWN, NEW_RAW_ARG6);
		if (*NEW_RAW_ARG7 && !line->ArgHasDeref(7) && toupper(*NEW_RAW_ARG7) != 'R')
			return ScriptError("Parameter #7 (if present) must be the letter R.", NEW_RAW_ARG7);
		// Check that the button is valid (e.g. left/right/middle):
		if (!line->ArgHasDeref(1))
			if (!line->ConvertMouseButton(NEW_RAW_ARG1))
				return ScriptError(ERR_MOUSE_BUTTON, NEW_RAW_ARG1);
		if (!line->ValidateMouseCoords(NEW_RAW_ARG2, NEW_RAW_ARG3))
			return ScriptError(ERR_MOUSE_COORD, NEW_RAW_ARG2);
		break;

	case ACT_MOUSECLICKDRAG:
		// Even though we check for blanks here at load-time, we don't bother to do so at runtime
		// (i.e. if a dereferenced var resolved to blank, it will be treated as a zero):
		if (!*NEW_RAW_ARG4 || !*NEW_RAW_ARG5)
			return ScriptError("Parameter #4 and #5 must not be blank.");
		if (*NEW_RAW_ARG6 && !line->ArgHasDeref(6))
		{
			value = ATOI(NEW_RAW_ARG6);
			if (value < 0 || value > MAX_MOUSE_SPEED)
				return ScriptError(ERR_MOUSE_SPEED, NEW_RAW_ARG6);
		}
		if (*NEW_RAW_ARG7 && !line->ArgHasDeref(7) && toupper(*NEW_RAW_ARG7) != 'R')
			return ScriptError("Parameter #7 (if present) must be the letter R.", NEW_RAW_ARG7);
		if (!line->ArgHasDeref(1))
			if (!line->ConvertMouseButton(NEW_RAW_ARG1, false))
				return ScriptError(ERR_MOUSE_BUTTON, NEW_RAW_ARG1);
		if (!line->ValidateMouseCoords(NEW_RAW_ARG2, NEW_RAW_ARG3))
			return ScriptError(ERR_MOUSE_COORD, NEW_RAW_ARG2);
		if (!line->ValidateMouseCoords(NEW_RAW_ARG4, NEW_RAW_ARG5))
			return ScriptError(ERR_MOUSE_COORD, NEW_RAW_ARG4);
		break;

	case ACT_CONTROLSEND:
	case ACT_CONTROLSENDRAW:
		// Window params can all be blank in this case, but characters to send should
		// be non-blank (but it's ok if its a dereferenced var that resolves to blank
		// at runtime):
		if (!*NEW_RAW_ARG2)
			return ScriptError("Parameter #2 must not be blank.");
		break;

	case ACT_CONTROLCLICK:
		// Check that the button is valid (e.g. left/right/middle):
		if (*NEW_RAW_ARG4 && !line->ArgHasDeref(4)) // i.e. it's allowed to be blank (defaults to left).
			if (!line->ConvertMouseButton(NEW_RAW_ARG4))
				return ScriptError(ERR_MOUSE_BUTTON, NEW_RAW_ARG4);
		break;

	case ACT_ADD:
	case ACT_SUB:
		if (aArgc > 2)
		{
			if (*NEW_RAW_ARG3 && !line->ArgHasDeref(3))
				if (!strchr("SMHD", toupper(*NEW_RAW_ARG3)))  // (S)econds, (M)inutes, (H)ours, or (D)ays
					return ScriptError(ERR_COMPARE_TIMES, NEW_RAW_ARG3);
			if (aActionType == ACT_SUB && *NEW_RAW_ARG2 && !line->ArgHasDeref(2))
				if (!YYYYMMDDToSystemTime(NEW_RAW_ARG2, st, true))
					return ScriptError(ERR_INVALID_DATETIME, NEW_RAW_ARG2);
		}
		break;

	case ACT_FILEINSTALL:
	case ACT_FILECOPY:
	case ACT_FILEMOVE:
	case ACT_FILECOPYDIR:
	case ACT_FILEMOVEDIR:
	case ACT_FILESELECTFOLDER:
		if (*NEW_RAW_ARG3 && !line->ArgHasDeref(3))
		{
			if (strlen(NEW_RAW_ARG3) > 1 || *NEW_RAW_ARG3 < '0' || *NEW_RAW_ARG3 > '3')
				if (aActionType != ACT_FILEMOVEDIR || toupper(*NEW_RAW_ARG3) != 'R')
					return ScriptError("Parameter #3 is not valid.", NEW_RAW_ARG3);
		}
		if (aActionType == ACT_FILEINSTALL)
		{
			if (aArgc > 0 && line->ArgHasDeref(1))
				return ScriptError("Parameter #1 must not contain references to variables."
					, NEW_RAW_ARG1);
		}
		break;

	case ACT_FILEREMOVEDIR:
		if (*NEW_RAW_ARG2 && !line->ArgHasDeref(2))
		{
			if (strlen(NEW_RAW_ARG2) > 1 || (*NEW_RAW_ARG2 != '0' && *NEW_RAW_ARG2 != '1'))
				return ScriptError("Parameter #2 must be either blank, 0, 1, or a variable reference."
					, NEW_RAW_ARG2);
		}
		break;

	case ACT_FILESETATTRIB:
		if (*NEW_RAW_ARG1 && !line->ArgHasDeref(1))
		{
			for (char *cp = NEW_RAW_ARG1; *cp; ++cp)
				if (!strchr("+-^RASHNOT", toupper(*cp)))
					return ScriptError("Parameter #1 contains unsupported file-attribute letters or symbols."
						, NEW_RAW_ARG1);
		}
		if (aArgc > 2 && !line->ArgHasDeref(3) && line->ConvertLoopMode(NEW_RAW_ARG3) == FILE_LOOP_INVALID)
			return ScriptError("If not blank, parameter #3 must be either 0, 1, 2, or a variable reference."
				, NEW_RAW_ARG3);
		if (*NEW_RAW_ARG4 && !line->ArgHasDeref(4))
			if (strlen(NEW_RAW_ARG4) > 1 || (*NEW_RAW_ARG4 != '0' && *NEW_RAW_ARG4 != '1'))
				return ScriptError("Parameter #4 must be either blank, 0, 1, or a variable reference."
					, NEW_RAW_ARG4);
		break;

	case ACT_FILEGETTIME:
		if (*NEW_RAW_ARG3 && !line->ArgHasDeref(3))
			if (strlen(NEW_RAW_ARG3) > 1 || !strchr("MCA", toupper(*NEW_RAW_ARG3)))
				return ScriptError(ERR_FILE_TIME, NEW_RAW_ARG3);
		break;

	case ACT_FILESETTIME:
		if (*NEW_RAW_ARG1 && !line->ArgHasDeref(1))
			if (!YYYYMMDDToSystemTime(NEW_RAW_ARG1, st, true))
				return ScriptError(ERR_INVALID_DATETIME, NEW_RAW_ARG1);
		if (*NEW_RAW_ARG3 && !line->ArgHasDeref(3))
			if (strlen(NEW_RAW_ARG3) > 1 || !strchr("MCA", toupper(*NEW_RAW_ARG3)))
				return ScriptError(ERR_FILE_TIME, NEW_RAW_ARG3);
		if (aArgc > 3 && !line->ArgHasDeref(4) && line->ConvertLoopMode(NEW_RAW_ARG4) == FILE_LOOP_INVALID)
			return ScriptError("If not blank, parameter #4 must be either 0, 1, 2, or a variable reference."
				, NEW_RAW_ARG4);
		if (*NEW_RAW_ARG5 && !line->ArgHasDeref(5))
			if (strlen(NEW_RAW_ARG5) > 1 || (*NEW_RAW_ARG5 != '0' && *NEW_RAW_ARG5 != '1'))
				return ScriptError("Parameter #5 must be either blank, 0, 1, or a variable reference."
					, NEW_RAW_ARG5);
		break;

	case ACT_FILEGETSIZE:
		if (*NEW_RAW_ARG3 && !line->ArgHasDeref(3))
			if (strlen(NEW_RAW_ARG3) > 1 || !strchr("BKM", toupper(*NEW_RAW_ARG3))) // Allow B=Bytes as undocumented.
				return ScriptError("Parameter #3 must be either blank, K, M, or a variable reference."
					, NEW_RAW_ARG3);
		break;

	case ACT_FILESELECTFILE:
		if (*NEW_RAW_ARG2 && !line->ArgHasDeref(2))
		{
			if (toupper(*NEW_RAW_ARG2 == 'S'))
                value = ATOI(NEW_RAW_ARG2 + 1);
			else
                value = ATOI(NEW_RAW_ARG2);
			if (value < 0 || value > 31)
				return ScriptError("Paremeter #2 must be either blank, a variable reference,"
					" or contain a number between 0 and 31 inclusive.", NEW_RAW_ARG2);
		}
		break;

	case ACT_SETTITLEMATCHMODE:
		if (aArgc > 0 && !line->ArgHasDeref(1) && !line->ConvertTitleMatchMode(NEW_RAW_ARG1))
			return ScriptError(ERR_TITLEMATCHMODE, NEW_RAW_ARG1);
		break;

	case ACT_SETFORMAT:
		if (aArgc > 0 && !line->ArgHasDeref(1))
		{
            if (!stricmp(NEW_RAW_ARG1, "float"))
			{
				if (aArgc > 1 && !line->ArgHasDeref(2))
				{
					if (!IsPureNumeric(NEW_RAW_ARG2, true, false, true))
						return ScriptError("Parameter #2 must be numeric when used with FLOAT.", NEW_RAW_ARG2);
					if (strlen(NEW_RAW_ARG2) >= sizeof(g.FormatFloat) - 2)
						return ScriptError("Parameter #2 is too long for use with FLOAT.", NEW_RAW_ARG2);
				}
			}
			else if (!stricmp(NEW_RAW_ARG1, "integer"))
			{
				if (aArgc > 1 && !line->ArgHasDeref(2) && toupper(*NEW_RAW_ARG2) != 'H' && toupper(*NEW_RAW_ARG2) != 'D')
					return ScriptError("Parameter #2 must be either H or D when used with INTEGER.", NEW_RAW_ARG2);
			}
			else
				return ScriptError("Parameter #1 must be a variable reference or the word INTEGER or FLOAT.", NEW_RAW_ARG1);
		}
		// Size must be less than sizeof() minus 2 because need room to prepend the '%' and append
		// the 'f' to make it a valid format specifier string:
		break;

	case ACT_TRANSFORM:
		if (aArgc > 1 && !line->ArgHasDeref(2))
		{
			TransformCmds trans_cmd = line->ConvertTransformCmd(NEW_RAW_ARG2);
			if (trans_cmd == TRANS_CMD_INVALID)
				return ScriptError(ERR_TRANSFORMCOMMAND, NEW_RAW_ARG2);
			if (trans_cmd == TRANS_CMD_UNICODE && !*line->mArg[0].text) // blank text means output-var is not a dynamically built one.
			{
				// If the output var isn't the clipboard, the mode is "retrieve clipboard text as UTF-8".
				// Therefore, Param#3 should be blank in that case to avoid unnecessary fetching of the
				// entire clipboard contents as plain text when in fact the command itself will be
				// directly accessing the clipboard rather than relying on the automatic parameter and
				// deref handling.
				if (VAR(line->mArg[0])->mType == VAR_CLIPBOARD)
				{
					if (aArgc < 3)
						return ScriptError("Parameter #3 must not be blank in this case.");
				}
				else
					if (aArgc > 2)
						return ScriptError("Parameter #3 must be blank in this case.", NEW_RAW_ARG3);
				break; // This type has been fully checked above.
			}
			if (!line->ArgHasDeref(3))
			{
				switch(trans_cmd)
				{
				case TRANS_CMD_CHR:
				case TRANS_CMD_BITNOT:
				case TRANS_CMD_BITSHIFTLEFT:
				case TRANS_CMD_BITSHIFTRIGHT:
				case TRANS_CMD_BITAND:
				case TRANS_CMD_BITOR:
				case TRANS_CMD_BITXOR:
					if (!IsPureNumeric(NEW_RAW_ARG3, true, false))
						return ScriptError("Parameter #3 must be an integer in this case.", NEW_RAW_ARG3);
					break;

				case TRANS_CMD_MOD:
				case TRANS_CMD_EXP:
				case TRANS_CMD_ROUND:
				case TRANS_CMD_CEIL:
				case TRANS_CMD_FLOOR:
				case TRANS_CMD_ABS:
				case TRANS_CMD_SIN:
				case TRANS_CMD_COS:
				case TRANS_CMD_TAN:
				case TRANS_CMD_ASIN:
				case TRANS_CMD_ACOS:
				case TRANS_CMD_ATAN:
					if (!IsPureNumeric(NEW_RAW_ARG3, true, false, true))
						return ScriptError("Parameter #3 must be a number in this case.", NEW_RAW_ARG3);
					break;

				case TRANS_CMD_POW:
				case TRANS_CMD_SQRT:
				case TRANS_CMD_LOG:
				case TRANS_CMD_LN:
					if (!IsPureNumeric(NEW_RAW_ARG3, false, false, true))
						return ScriptError("Parameter #3 must be a positive integer in this case.", NEW_RAW_ARG3);
					break;

				// The following are not listed above because no validation of Paramter #3 is needed at this stage:
				// TRANS_CMD_ASC
				// TRANS_CMD_UNICODE
				// TRANS_CMD_HTML
				// TRANS_CMD_DEREF
				}
			}

			switch(trans_cmd)
			{
			case TRANS_CMD_ASC:
			case TRANS_CMD_CHR:
			case TRANS_CMD_DEREF:
			case TRANS_CMD_UNICODE:
			case TRANS_CMD_HTML:
			case TRANS_CMD_EXP:
			case TRANS_CMD_SQRT:
			case TRANS_CMD_LOG:
			case TRANS_CMD_LN:
			case TRANS_CMD_CEIL:
			case TRANS_CMD_FLOOR:
			case TRANS_CMD_ABS:
			case TRANS_CMD_SIN:
			case TRANS_CMD_COS:
			case TRANS_CMD_TAN:
			case TRANS_CMD_ASIN:
			case TRANS_CMD_ACOS:
			case TRANS_CMD_ATAN:
			case TRANS_CMD_BITNOT:
				if (*NEW_RAW_ARG4)
					return ScriptError("Parameter #4 should be omitted in this case.", NEW_RAW_ARG4);
				break;

			case TRANS_CMD_BITAND:
			case TRANS_CMD_BITOR:
			case TRANS_CMD_BITXOR:
				if (!line->ArgHasDeref(4) && !IsPureNumeric(NEW_RAW_ARG4, true, false))
					return ScriptError("Parameter #4 must be an integer in this case.", NEW_RAW_ARG4);
				break;

			case TRANS_CMD_BITSHIFTLEFT:
			case TRANS_CMD_BITSHIFTRIGHT:
				if (!line->ArgHasDeref(4) && !IsPureNumeric(NEW_RAW_ARG4, false, false))
					return ScriptError("Parameter #4 must be a positive integer in this case.", NEW_RAW_ARG4);
				break;

			case TRANS_CMD_ROUND:
				if (*NEW_RAW_ARG4 && !line->ArgHasDeref(4) && !IsPureNumeric(NEW_RAW_ARG4, true, false))
					return ScriptError("Parameter #4 must be blank or an integer in this case.", NEW_RAW_ARG4);
				break;

			case TRANS_CMD_MOD:
			case TRANS_CMD_POW:
				if (!line->ArgHasDeref(4) && !IsPureNumeric(NEW_RAW_ARG4, true, false, true))
					return ScriptError("Parameter #4 must be a number in this case.", NEW_RAW_ARG4);
				break;
#ifdef _DEBUG
			default:
				return ScriptError("Unhandled", NEW_RAW_ARG2);  // To improve maintainability.
#endif
			}

			switch(trans_cmd)
			{
			case TRANS_CMD_CHR:
				if (!line->ArgHasDeref(3))
				{
					value = ATOI(NEW_RAW_ARG3);
					if (!IsPureNumeric(NEW_RAW_ARG3, false, false) || value < 0 || value > 255)
						return ScriptError("Parameter #3 must be between 0 and 255 inclusive.", NEW_RAW_ARG3);
				}
				break;
			case TRANS_CMD_DEREF:
				// Seems best to issue a warning for this, since it's hard to imagine anyone ever intentionally
				// doing it under any circumstances:
				if (!line->ArgHasDeref(3))
					return ScriptError("Because parameter #3 contains no variable reference(s), "
						"this line serves no purpose.", NEW_RAW_ARG3);
				break;
			case TRANS_CMD_MOD:
				if (!line->ArgHasDeref(4) && !ATOI64(NEW_RAW_ARG4)) // Parameter is omitted or something that resolves to zero.
					return ScriptError(ERR_DIVIDEBYZERO, NEW_RAW_ARG4);
				break;
			}
		}
		break;

	case ACT_MENU:
		if (aArgc > 1 && !line->ArgHasDeref(2))
		{
			MenuCommands menu_cmd = line->ConvertMenuCommand(NEW_RAW_ARG2);

			switch(menu_cmd)
			{
			case MENU_CMD_TIP:
			case MENU_CMD_ICON:
			case MENU_CMD_NOICON:
			case MENU_CMD_MAINWINDOW:
			case MENU_CMD_NOMAINWINDOW:
			case MENU_CMD_CLICK:
			{
				bool is_tray = true;  // Assume true if unknown.
				if (aArgc > 0 && !line->ArgHasDeref(1))
					if (stricmp(NEW_RAW_ARG1, "tray"))
						is_tray = false;
				if (!is_tray)
					return ScriptError(ERR_MENUTRAY, NEW_RAW_ARG1);
				break;
			}
			}

			switch (menu_cmd)
			{
			case MENU_CMD_INVALID:
				return ScriptError(ERR_MENUCOMMAND, NEW_RAW_ARG2);

			case MENU_CMD_NODEFAULT:
			case MENU_CMD_STANDARD:
			case MENU_CMD_NOSTANDARD:
			case MENU_CMD_DELETEALL:
			case MENU_CMD_NOICON:
			case MENU_CMD_MAINWINDOW:
			case MENU_CMD_NOMAINWINDOW:
				if (*NEW_RAW_ARG3 || *NEW_RAW_ARG4 || *NEW_RAW_ARG5 || *NEW_RAW_ARG6)
					return ScriptError("Parameter #3 and beyond should be omitted in this case.", NEW_RAW_ARG3);
				break;

			case MENU_CMD_RENAME:
			case MENU_CMD_USEERRORLEVEL:
			case MENU_CMD_CHECK:
			case MENU_CMD_UNCHECK:
			case MENU_CMD_TOGGLECHECK:
			case MENU_CMD_ENABLE:
			case MENU_CMD_DISABLE:
			case MENU_CMD_TOGGLEENABLE:
			case MENU_CMD_DEFAULT:
			case MENU_CMD_DELETE:
			case MENU_CMD_TIP:
			case MENU_CMD_CLICK:
				if (   menu_cmd != MENU_CMD_RENAME && (*NEW_RAW_ARG4 || *NEW_RAW_ARG5 || *NEW_RAW_ARG6)   )
					return ScriptError("Parameter #4 and beyond should be omitted in this case.", NEW_RAW_ARG4);
				switch(menu_cmd)
				{
				case MENU_CMD_USEERRORLEVEL:
				case MENU_CMD_TIP:
				case MENU_CMD_DEFAULT:
				case MENU_CMD_DELETE:
					break;  // i.e. for commands other than the above, do the default below.
				default:
					if (!*NEW_RAW_ARG3)
						return ScriptError("Parameter #3 must not be blank in this case.");
				}
				break;

			// These have a highly variable number of parameters, or are too rarely used
			// to warrant detailed load-time checking, so are not validated here:
			//case MENU_CMD_SHOW:
			//case MENU_CMD_ADD:
			//case MENU_CMD_COLOR:
			//case MENU_CMD_ICON:
			}
		}
		break;

	case ACT_GUI:
		// By design, scripts that use the GUI cmd anywhere are persistent.  Doing this here
		// also allows WinMain() to later detect whether this script should become #SingleInstance.
		// Note: Don't directly change g_AllowOnlyOneInstance here in case the  remainder of the
		// script-loading process comes across any explicit uses of #SinngleInstance,
		// in which case it would take precedence.
		g_persistent = true;
		if (aArgc > 0 && !line->ArgHasDeref(1))
		{
			GuiCommands gui_cmd = line->ConvertGuiCommand(NEW_RAW_ARG1);

			switch (gui_cmd)
			{
			case GUI_CMD_INVALID:
				return ScriptError(ERR_GUICOMMAND, NEW_RAW_ARG1);
			case GUI_CMD_ADD:
				if (aArgc > 1 && !line->ArgHasDeref(2) && !line->ConvertGuiControl(NEW_RAW_ARG2))
					return ScriptError(ERR_GUICONTROL, NEW_RAW_ARG2);
				break;
			case GUI_CMD_CANCEL:
			case GUI_CMD_DESTROY:
			case GUI_CMD_DEFAULT:
				if (aArgc > 1)
					return ScriptError("Parameter #2 and beyond should be omitted in this case.", NEW_RAW_ARG2);
				break;
			case GUI_CMD_SUBMIT:
			case GUI_CMD_MENU:
			case GUI_CMD_FLASH:
				if (aArgc > 2)
					return ScriptError("Parameter #3 and beyond should be omitted in this case.", NEW_RAW_ARG3);
				break;
			// No action for these since they have a varying number of optional params:
			//case GUI_CMD_SHOW:
			//case GUI_CMD_FONT:
			//case GUI_CMD_TAB:
			//case GUI_CMD_COLOR: No load-time param validation to avoid larger EXE size.
			}
		}
		break;

	case ACT_THREAD:
		if (aArgc > 0 && !line->ArgHasDeref(1) && !line->ConvertThreadCommand(NEW_RAW_ARG1))
			return ScriptError(ERR_THREADCOMMAND, NEW_RAW_ARG1);
		break;

	case ACT_CONTROL:
		if (aArgc > 0 && !line->ArgHasDeref(1))
		{
			ControlCmds control_cmd = line->ConvertControlCmd(NEW_RAW_ARG1);
			switch (control_cmd)
			{
			case CONTROL_CMD_INVALID:
				return ScriptError(ERR_CONTROLCOMMAND, NEW_RAW_ARG1);
			case CONTROL_CMD_TABLEFT:
			case CONTROL_CMD_TABRIGHT:
			case CONTROL_CMD_ADD:
			case CONTROL_CMD_DELETE:
			case CONTROL_CMD_CHOOSE:
			case CONTROL_CMD_CHOOSESTRING:
			case CONTROL_CMD_EDITPASTE:
				if (control_cmd != CONTROL_CMD_TABLEFT && control_cmd != CONTROL_CMD_TABRIGHT && !*NEW_RAW_ARG2)
					return ScriptError("Parameter #2 must not be blank in this case.");
				break;
			default: // All commands except the above should have a blank Value parameter.
				if (*NEW_RAW_ARG2)
					return ScriptError("Parameter #2 must be blank in this case.", NEW_RAW_ARG2);
			}
		}
		break;

	case ACT_CONTROLGET:
		if (aArgc > 1 && !line->ArgHasDeref(2))
		{
			ControlGetCmds control_get_cmd = line->ConvertControlGetCmd(NEW_RAW_ARG2);
			switch (control_get_cmd)
			{
			case CONTROLGET_CMD_INVALID:
				return ScriptError(ERR_CONTROLGETCOMMAND, NEW_RAW_ARG2);
			case CONTROLGET_CMD_FINDSTRING:
			case CONTROLGET_CMD_LINE:
				if (!*NEW_RAW_ARG3)
					return ScriptError("Parameter #3 must not be blank in this case.");
				break;
			default: // All commands except the above should have a blank Value parameter.
				if (*NEW_RAW_ARG3)
					return ScriptError("Parameter #3 must be blank in this case.", NEW_RAW_ARG3);
			}
		}
		break;

	case ACT_GUICONTROL:
		if (!*NEW_RAW_ARG2) // ControlID
			return ScriptError("Parameter #2 must not be blank.");
		if (aArgc > 0 && !line->ArgHasDeref(1))
		{
			GuiControlCmds guicontrol_cmd = line->ConvertGuiControlCmd(NEW_RAW_ARG1);
			switch (guicontrol_cmd)
			{
			case GUICONTROL_CMD_INVALID:
				return ScriptError(ERR_GUICONTROLCOMMAND, NEW_RAW_ARG1);
			case GUICONTROL_CMD_CONTENTS:
			case GUICONTROL_CMD_TEXT:
				break; // Do nothing for the above commands since Param3 is optional.
			case GUICONTROL_CMD_MOVE:
			case GUICONTROL_CMD_CHOOSE:
			case GUICONTROL_CMD_CHOOSESTRING:
				if (!*NEW_RAW_ARG3)
					return ScriptError("Parameter #3 must not be blank in this case.");
				break;
			default: // All commands except the above should have a blank Text parameter.
				if (*NEW_RAW_ARG3)
					return ScriptError("Parameter #3 must be blank in this case.", NEW_RAW_ARG3);
			}
		}
		break;

	case ACT_GUICONTROLGET:
		if (aArgc > 1 && !line->ArgHasDeref(2))
		{
			GuiControlGetCmds guicontrolget_cmd = line->ConvertGuiControlGetCmd(NEW_RAW_ARG2);
			// This first check's error messages take precedence over the next check's:
			switch (guicontrolget_cmd)
			{
			case GUICONTROLGET_CMD_INVALID:
				return ScriptError(ERR_GUICONTROLGETCOMMAND, NEW_RAW_ARG2);
			case GUICONTROLGET_CMD_CONTENTS:
				break; // Do nothing, since Param4 is optional in this case.
			default: // All commands except the above should have a blank parameter here.
				if (*NEW_RAW_ARG4) // Currently true for all, since it's a FutureUse param.
					return ScriptError("Parameter #4 must be blank in this case.", NEW_RAW_ARG4);
			}
			if (guicontrolget_cmd == GUICONTROLGET_CMD_FOCUS)
			{
				if (*NEW_RAW_ARG3)
					return ScriptError("Parameter #3 must be blank in this case.", NEW_RAW_ARG3);
			}
			// else it can be optionally blank, in which case the output variable is used as the
			// ControlID also.
		}
		break;

	case ACT_DRIVE:
		if (aArgc > 0 && !line->ArgHasDeref(1))
		{
			DriveCmds drive_cmd = line->ConvertDriveCmd(NEW_RAW_ARG1);
			if (!drive_cmd)
				return ScriptError(ERR_DRIVECOMMAND, NEW_RAW_ARG1);
			if (drive_cmd != DRIVE_CMD_EJECT && !*NEW_RAW_ARG2)
				return ScriptError("Parameter #2 must not be blank in this case.");
			// For DRIVE_CMD_LABEL: Note that is is possible and allowed for the new label to be blank.
			// Not currently done since all sub-commands take a mandatory or optional ARG3:
			//if (drive_cmd != ... && *NEW_RAW_ARG3)
			//	return ScriptError("Parameter #3 must be blank in this case.", NEW_RAW_ARG3);
		}
		break;

	case ACT_DRIVEGET:
		if (!line->ArgHasDeref(2))  // Don't check "aArgc > 1" because of DRIVEGET_CMD_SETLABEL's param format.
		{
			DriveGetCmds drive_get_cmd = line->ConvertDriveGetCmd(NEW_RAW_ARG2);
			if (!drive_get_cmd)
				return ScriptError(ERR_DRIVEGETCOMMAND, NEW_RAW_ARG2);
			if (drive_get_cmd != DRIVEGET_CMD_LIST && drive_get_cmd != DRIVEGET_CMD_STATUSCD && !*NEW_RAW_ARG3)
				return ScriptError("Parameter #3 must not be blank in this case.");
			if (drive_get_cmd != DRIVEGET_CMD_SETLABEL && (aArgc < 1 || line->mArg[0].type == ARG_TYPE_NORMAL))
				// The output variable has been omitted.
				return ScriptError("Parameter #1 must not be blank in this case.");
		}
		break;

	case ACT_PROCESS:
		if (aArgc > 0 && !line->ArgHasDeref(1))
		{
			ProcessCmds process_cmd = line->ConvertProcessCmd(NEW_RAW_ARG1);
			if (process_cmd != PROCESS_CMD_PRIORITY && !*NEW_RAW_ARG2)
				return ScriptError("Parameter #2 must not be blank in this case.");
			switch (process_cmd)
			{
			case PROCESS_CMD_INVALID:
				return ScriptError(ERR_PROCESSCOMMAND, NEW_RAW_ARG1);
			case PROCESS_CMD_EXIST:
			case PROCESS_CMD_CLOSE:
				if (*NEW_RAW_ARG3)
					return ScriptError("Parameter #3 must be blank in this case.", NEW_RAW_ARG3);
				break;
			case PROCESS_CMD_PRIORITY:
				if (!*NEW_RAW_ARG3 || (!line->ArgHasDeref(3) && !strchr(PROCESS_PRIORITY_LETTERS, toupper(*NEW_RAW_ARG3))))
					return ScriptError("Parameter #3 must be one of the following letters: "
						PROCESS_PRIORITY_LETTERS ".", NEW_RAW_ARG3);
				break;
			case PROCESS_CMD_WAIT:
			case PROCESS_CMD_WAITCLOSE:
				if (*NEW_RAW_ARG3 && !line->ArgHasDeref(3) && !IsPureNumeric(NEW_RAW_ARG3, false, true, true))
					return ScriptError("If present, parameter #3 must be a positive number in this case.", NEW_RAW_ARG3);
				break;
			}
		}
		break;

	// For ACT_WINMOVE, don't validate anything for mandatory args so that its two modes of
	// operation can be supported: 2-param mode and normal-param mode.
	// For these, although we validate that at least one is non-blank here, it's okay at
	// runtime for them all to resolve to be blank, without an error being reported.
	// It's probably more flexible that way since the commands are equipped to handle
	// all-blank params.
	// Not these because they can be used with the "last-used window" mode:
	//case ACT_IFWINEXIST:
	//case ACT_IFWINNOTEXIST:
	// Not these because they can have their window params all-blank to work in "last-used window" mode:
	//case ACT_IFWINACTIVE:
	//case ACT_IFWINNOTACTIVE:
	//case ACT_WINACTIVATE:
	//case ACT_WINWAITCLOSE:
	//case ACT_WINWAITACTIVE:
	//case ACT_WINWAITNOTACTIVE:
	case ACT_WINACTIVATEBOTTOM:
		if (!*NEW_RAW_ARG1 && !*NEW_RAW_ARG2 && !*NEW_RAW_ARG3 && !*NEW_RAW_ARG4)
			return ScriptError(ERR_WINDOW_PARAM);
		break;

	case ACT_WINWAIT:
		if (!*NEW_RAW_ARG1 && !*NEW_RAW_ARG2 && !*NEW_RAW_ARG4 && !*NEW_RAW_ARG5) // ARG3 is omitted because it's the timeout.
			return ScriptError(ERR_WINDOW_PARAM);
		break;

	case ACT_WINMENUSELECTITEM:
		// Window params can all be blank in this case, but the first menu param should
		// be non-blank (but it's ok if its a dereferenced var that resolves to blank
		// at runtime):
		if (!*NEW_RAW_ARG3)
			return ScriptError("Parameter #3 must not be blank.");
		break;

	case ACT_WINSET:
		if (aArgc > 0 && !line->ArgHasDeref(1))
		{
			switch(line->ConvertWinSetAttribute(NEW_RAW_ARG1))
			{
			case WINSET_TRANSPARENT:
				if (aArgc > 1 && !line->ArgHasDeref(2))
				{
					value = ATOI(NEW_RAW_ARG2);
					if (value < 0 || value > 255)
						return ScriptError("Parameter #2 must be between 0 and 255 when used with TRANSPARENT."
							, NEW_RAW_ARG2);
				}
				break;
			case WINSET_TRANSCOLOR:
				if (!*NEW_RAW_ARG2)
					return ScriptError("Parameter #2 must not be blank in this case.");
				break;
			case WINSET_ALWAYSONTOP:
				if (aArgc > 1 && !line->ArgHasDeref(2) && !line->ConvertOnOffToggle(NEW_RAW_ARG2))
					return ScriptError(ERR_ON_OFF_TOGGLE, NEW_RAW_ARG2);
				break;
			case WINSET_BOTTOM:
				if (*NEW_RAW_ARG2)
					return ScriptError("Parameter #2 must be blank in this case.");
				break;
			case WINSET_INVALID:
				return ScriptError(ERR_WINSET, NEW_RAW_ARG1);
			}
		}
		break;

	case ACT_WINGET:
		if (!line->ArgHasDeref(2) && !line->ConvertWinGetCmd(NEW_RAW_ARG2)) // It's okay if ARG2 is blank.
			return ScriptError(ERR_WINGET, NEW_RAW_ARG2);
		break;

	case ACT_SYSGET:
		if (!line->ArgHasDeref(2) && !line->ConvertSysGetCmd(NEW_RAW_ARG2))
			return ScriptError(ERR_SYSGET, NEW_RAW_ARG2);
		break;

	case ACT_INPUTBOX:
		if (*NEW_RAW_ARG9)  // && !line->ArgHasDeref(9)
			return ScriptError("Parameter #9 is for future use and should be left blank.", NEW_RAW_ARG9);
		break;

	case ACT_MSGBOX:
		if (aArgc > 1) // i.e. this MsgBox is using the 3-param or 4-param style.
			if (!line->ArgHasDeref(1)) // i.e. if it's a deref, we won't try to validate it now.
				if (!IsPureNumeric(NEW_RAW_ARG1)) // Allow it to be entirely whitespace to indicate 0, like Aut2.
					return ScriptError("When used with more than one parameter, MsgBox requires that"
						" the 1st parameter be numeric or a variable reference.", NEW_RAW_ARG1);
		if (aArgc > 3)
			if (!line->ArgHasDeref(4)) // i.e. if it's a deref, we won't try to validate it now.
				if (!IsPureNumeric(NEW_RAW_ARG4, false, true, true))
					return ScriptError("MsgBox requires that the 4th parameter, if present, be numeric & positive,"
						" or a variable reference.", NEW_RAW_ARG4);
		break;

	case ACT_IFMSGBOX:
		if (aArgc > 0 && !line->ArgHasDeref(1) && !line->ConvertMsgBoxResult(NEW_RAW_ARG1))
			return ScriptError(ERR_IFMSGBOX, NEW_RAW_ARG1);
		break;

	case ACT_IFIS:
	case ACT_IFISNOT:
		if (aArgc > 1 && !line->ArgHasDeref(2) && !line->ConvertVariableTypeName(NEW_RAW_ARG2))
			// Don't refer to it as "Parameter #2" because this command isn't formatted/displayed that way:
			return ScriptError("The type name must be either NUMBER, INTEGER, FLOAT, TIME, DATE, DIGIT, ALPHA, ALNUM, SPACE, or a variable reference."
				, NEW_RAW_ARG2);
		break;

	case ACT_GETKEYSTATE:
		if (aArgc > 1 && !line->ArgHasDeref(2) && !TextToVK(NEW_RAW_ARG2) && !Line::ConvertJoy(NEW_RAW_ARG2))
			return ScriptError("This is not a valid key, mouse, or joystick button.", NEW_RAW_ARG2);
		break;

	case ACT_KEYWAIT:
		if (aArgc > 0 && !line->ArgHasDeref(1) && !TextToVK(NEW_RAW_ARG1) && !Line::ConvertJoy(NEW_RAW_ARG1))
			return ScriptError("This is not a valid key, mouse, or joystick button.", NEW_RAW_ARG1);
		break;

	case ACT_DIV:
		if (!line->ArgHasDeref(2)) // i.e. if it's a deref, we won't try to validate it now.
			if (!ATOI64(NEW_RAW_ARG2))
				return ScriptError(ERR_DIVIDEBYZERO, NEW_RAW_ARG2);
		break;

	case ACT_GROUPADD:
	case ACT_GROUPACTIVATE:
	case ACT_GROUPDEACTIVATE:
	case ACT_GROUPCLOSE:
		// For all these, store a pointer to the group to help performance.
		// We create a non-existent group even for ACT_GROUPACTIVATE, ACT_GROUPDEACTIVATE
		// and ACT_GROUPCLOSE because we can't rely on the ACT_GROUPADD commands having
		// been parsed prior to them (e.g. something like "Gosub, DefineGroups" may appear
		// in the auto-execute portion of the script).
		if (!line->ArgHasDeref(1))
			if (   !(line->mAttribute = FindOrAddGroup(NEW_RAW_ARG1))   )
				return FAIL;  // The above already displayed the error.
		if (aActionType == ACT_GROUPADD && !*NEW_RAW_ARG2 && !*NEW_RAW_ARG3 && !*NEW_RAW_ARG5 && !*NEW_RAW_ARG6) // ARG4 is the JumpToLine
			return ScriptError(ERR_WINDOW_PARAM);
		if (aActionType == ACT_GROUPACTIVATE || aActionType == ACT_GROUPDEACTIVATE)
		{
			if (*NEW_RAW_ARG2 && !line->ArgHasDeref(2))
				if (strlen(NEW_RAW_ARG2) > 1 || toupper(*NEW_RAW_ARG2) != 'R')
					return ScriptError("Parameter #2 must be either blank, R, or a variable reference."
						, NEW_RAW_ARG2);
		}
		else if (aActionType == ACT_GROUPCLOSE)
			if (*NEW_RAW_ARG2 && !line->ArgHasDeref(2))
				if (strlen(NEW_RAW_ARG2) > 1 || !strchr("RA", toupper(*NEW_RAW_ARG2)))
					return ScriptError("Parameter #2 must be either blank, R, A, or a variable reference."
						, NEW_RAW_ARG2);
		break;

	case ACT_REPEAT: // These types of loops are always "NORMAL".
		line->mAttribute = ATTR_LOOP_NORMAL;
		break;

	case ACT_LOOP:
		// If possible, determine the type of loop so that the preparser can better
		// validate some things:
		switch (aArgc)
		{
		case 0:
			line->mAttribute = ATTR_LOOP_NORMAL;
			break;
		case 1:
			if (line->ArgHasDeref(1)) // Impossible to know now what type of loop (only at runtime).
				line->mAttribute = ATTR_LOOP_UNKNOWN;
			else
			{
				if (IsPureNumeric(NEW_RAW_ARG1, false))
					line->mAttribute = ATTR_LOOP_NORMAL;
				else
					line->mAttribute = line->RegConvertRootKey(NEW_RAW_ARG1) ? ATTR_LOOP_REG : ATTR_LOOP_FILE;
			}
			break;
		default:  // has 2 or more args.
			if (line->ArgHasDeref(1)) // Impossible to know now what type of loop (only at runtime).
				line->mAttribute = ATTR_LOOP_UNKNOWN;
			else if (!stricmp(NEW_RAW_ARG1, "Read"))
				line->mAttribute = ATTR_LOOP_READ_FILE;
			else if (!stricmp(NEW_RAW_ARG1, "Parse"))
				line->mAttribute = ATTR_LOOP_PARSE;
			else // the 1st arg can either be a Root Key or a File Pattern, depending on the type of loop.
			{
				line->mAttribute = line->RegConvertRootKey(NEW_RAW_ARG1) ? ATTR_LOOP_REG : ATTR_LOOP_FILE;
				if (line->mAttribute == ATTR_LOOP_FILE)
				{
					// Validate whatever we can rather than waiting for runtime validation:
					if (!line->ArgHasDeref(2) && Line::ConvertLoopMode(NEW_RAW_ARG2) == FILE_LOOP_INVALID)
						return ScriptError(ERR_LOOP_FILE_MODE, NEW_RAW_ARG2);
					if (*NEW_RAW_ARG3 && !line->ArgHasDeref(3))
						if (strlen(NEW_RAW_ARG3) > 1 || (*NEW_RAW_ARG3 != '0' && *NEW_RAW_ARG3 != '1'))
							return ScriptError("Parameter #3 must be either blank, 0, 1, or a variable reference."
								, NEW_RAW_ARG3);
				}
				else // Registry loop.
				{
					if (aArgc > 2 && !line->ArgHasDeref(3) && Line::ConvertLoopMode(NEW_RAW_ARG3) == FILE_LOOP_INVALID)
						return ScriptError(ERR_LOOP_REG_MODE, NEW_RAW_ARG3);
					if (*NEW_RAW_ARG4 && !line->ArgHasDeref(4))
						if (strlen(NEW_RAW_ARG4) > 1 || (*NEW_RAW_ARG4 != '0' && *NEW_RAW_ARG4 != '1'))
							return ScriptError("Parameter #4 must be either blank, 0, 1, or a variable reference."
								, NEW_RAW_ARG4);
				}
			}
		}
		break; // Outer switch().
	}


	///////////////////////////////////////////////////////////////
	// Update any labels that should refer to the newly added line.
	///////////////////////////////////////////////////////////////
	// If the label most recently added doesn't yet have an anchor in the script, provide it.
	// UPDATE: In addition, keep searching backward through the labels until a non-NULL
	// mJumpToLine is found.  All the ones with a NULL should point to this new line to
	// support cases where one label immediately follows another in the script.
	// Example:
	// #a::  <-- don't leave this label with a NULL jumppoint.
	// LaunchA:
	// ...
	// return
	for (Label *label = mLastLabel; label != NULL && label->mJumpToLine == NULL; label = label->mPrevLabel)
	{
		if (line->mActionType == ACT_ELSE)
			return ScriptError("A label mustn't point to an ELSE.");
		// Don't allow this because it may cause problems in a case such as this because
		// label1 points to the end-block which is at the same level (and thus normally
		// an allowable jumppoint) as the goto.  But we don't want to allow jumping into
		// a block that belongs to a control structure.  In this case, it would probably
		// result in a runtime error when the execution unexpectedly encounters the ELSE
		// after performing the goto:
		// goto, label1
		// if x
		// {
		//    ...
		//    label1:
		// }
		// else
		//    ...
		//
		// An alternate way to deal with the above would be to make each block-end be owned
		// by its block-begin rather than the block that encloses them both.
		if (line->mActionType == ACT_BLOCK_END)
			return ScriptError("A label mustn't point to the end of a block."
				"  If this block is a loop, you can use the \"continue\" command to jump"
				" to the end of the block.");
		label->mJumpToLine = line;
	}

	++mLineCount;  // Right before returning "success", increment our count.
	return OK;
}



ResultType Script::ParseDerefs(char *aArgText, char *aArgMap, DerefType *aDeref, int &aDerefCount)
// Caller provides modifiable aDerefCount, which might be non-zero to indicate that there are already
// some items in the aDeref array.
// Returns FAIL or OK.
{
	size_t deref_string_length; // So that overflow can be detected, this is not of type DerefLengthType.

	// For each dereference found in aArgText:
	for (int j = 0;; ++j)  // Increment to skip over the symbol just found by the inner for().
	{
		// Find next non-literal g_DerefChar:
		for (; aArgText[j] && (aArgText[j] != g_DerefChar || (aArgMap && aArgMap[j])); ++j);
		if (!aArgText[j])
			break;
		// else: Match was found; this is the deref's open-symbol.
		if (aDerefCount >= MAX_DEREFS_PER_ARG)
			return ScriptError(TOO_MANY_VAR_REFS, aArgText); // Short msg since so rare.
		DerefType &this_deref = aDeref[aDerefCount];  // For performance.
		this_deref.marker = aArgText + j;  // Store the deref's starting location.
		// Find next g_DerefChar, even if it's a literal.
		for (++j; aArgText[j] && aArgText[j] != g_DerefChar; ++j);
		if (!aArgText[j])
			return ScriptError("This parameter contains a variable name missing its ending percent sign.", aArgText);
		// Otherwise: Match was found; this should be the deref's close-symbol.
		if (aArgMap && aArgMap[j])  // But it's mapped as literal g_DerefChar.
			return ScriptError("This parameter contains a variable name with an invalid escaped percent sign.", aArgText);
		deref_string_length = aArgText + j - this_deref.marker + 1;
		if (deref_string_length == 2) // The percent signs were empty, e.g. %%
			return ScriptError("This parmeter contains an empty variable reference (%%).", aArgText);
		if (deref_string_length - 2 > MAX_VAR_NAME_LENGTH) // -2 for the opening & closing g_DerefChars
			return ScriptError("This parmeter contains a variable name that is too long.", aArgText);
		this_deref.length = (DerefLengthType)deref_string_length;
		if (   !(this_deref.var = FindOrAddVar(this_deref.marker + 1, this_deref.length - 2))   )
			return FAIL;  // The called function already displayed the error.
		++aDerefCount;
	} // for each dereference.

	return OK;
}



Var *Line::ResolveVarOfArg(int aArgIndex, bool aCreateIfNecessary)
// Returns NULL on failure.
// Args that are input or output variables are normally resolved at load-time, so that
// they contain a pointer to their Var object.  This is done for performance.  However,
// in order to support dynamically resolved variables names like AutoIt2 (e.g. arrays),
// we have to do some extra work here at runtime.
// Callers specify false for aCreateIfNecessary whenever the contents of the variable
// they're trying to find is unimportant.  For example, dynamically built input variables,
// such as "StringLen, length, array%i%", do not need to be created if they weren't
// previously assigned to (i.e. they weren't previously used as an output variable).
// In the above example, the array element would never be created here.  But if the output
// variable were dynamic, our call would have told us to create it.
{
	// The requested ARG isn't even present, so it can't have a variable.  Currently, this should
	// never happen because the loading procedure ensures that input/output args are not marked
	// as variables if they are blank (and our caller should check for this and not call in that case):
	if (aArgIndex >= mArgc)
		return NULL;
	// Since this function isn't inline (since it's called so frequently), there isn't that much more
	// overhead to doing this check, even though it shouldn't be needed since it's the caller's
	// responsibility:
	if (mArg[aArgIndex].type == ARG_TYPE_NORMAL) // Arg isn't an input or output variable.
		return NULL;
	if (!*mArg[aArgIndex].text) // The arg's variable is not one that needs to be dynamically resolved.
		return VAR(mArg[aArgIndex]); // Return the var's address that was already determined at load-time.
	// The above might return NULL in the case where the arg is optional (i.e. the command allows
	// the var name to be omitted).  But in that case, the caller should either never have called this
	// function or should check for NULL upon return.  UPDATE: This actually never happens, see
	// comment above the "if (aArgIndex >= mArgc)" line.

	// Static to correspond to the static empty_var further below.  It needs the memory area
	// to support resolving dynamic environment variables.  In the following example,
	// the result will be blank unless the first line is present (without this fix here):
	//null = %SystemRoot%  ; bogus line as a required workaround in versions prior to v1.0.16
	//thing = SystemRoot
	//StringTrimLeft, output, %thing%, 0
	//msgbox %output%

	static char var_name[MAX_VAR_NAME_LENGTH + 1];  // Will hold the dynamically built name.

	// At this point, we know the requested arg is a variable that must be dynamically resolved.
	// This section is similar to that in ExpandArg(), so they should be maintained together:
	char *pText;
	DerefType *deref;
	int vni;

	for (vni = 0, pText = mArg[aArgIndex].text  // Start at the begining of this arg's text.
		, deref = mArg[aArgIndex].deref  // Start off by looking for the first deref.
		; deref && deref->marker; ++deref)  // A deref with a NULL marker terminates the list.
	{
		// FOR EACH DEREF IN AN ARG (if we're here, there's at least one):
		// Copy the chars that occur prior to deref->marker into the buffer:
		for (; pText < deref->marker && vni < MAX_VAR_NAME_LENGTH; var_name[vni++] = *pText++);
		if (vni >= MAX_VAR_NAME_LENGTH && pText < deref->marker) // The variable name would be too long!
		{
			// This type of error is just a warning because this function isn't set up to cause a true
			// failure.  This is because the use of dynamically named variables is rare, and only for
			// people who should know what they're doing.  In any case, when the caller of this
			// function called it to resolve an output variable, it will see tha the result is
			// NULL and terminate the current subroutine.
			#define DYNAMIC_TOO_LONG "This dynamically built variable name is too long." \
				"  If this variable was not intended to be dynamic, remove the % symbols from it."
			LineError(DYNAMIC_TOO_LONG, FAIL, mArg[aArgIndex].text);
			return NULL;
		}
		// Now copy the contents of the dereferenced var.  For all cases, aBuf has already
		// been verified to be large enough, assuming the value hasn't changed between the
		// time we were called and the time the caller calculated the space needed.
		if (deref->var->Get() > (VarSizeType)(MAX_VAR_NAME_LENGTH - vni)) // The variable name would be too long!
		{
			LineError(DYNAMIC_TOO_LONG, FAIL, mArg[aArgIndex].text);
			return NULL;
		}
		vni += deref->var->Get(var_name + vni);
		// Finally, jump over the dereference text. Note that in the case of an expression, there might not
		// be any percent signs within the text of the dereference, e.g. x + y, not %x% + %y%.
		pText += deref->length;
	}
	// Copy any chars that occur after the final deref into the buffer:
	for (; *pText && vni < MAX_VAR_NAME_LENGTH; var_name[vni++] = *pText++);
	if (vni >= MAX_VAR_NAME_LENGTH && *pText) // The variable name would be too long!
	{
		LineError(DYNAMIC_TOO_LONG, FAIL, mArg[aArgIndex].text);
		return NULL;
	}
	
	if (!vni)
	{
		LineError("This dynamic variable is blank.  If this variable was not intended to be dynamic,"
			" remove the % symbols from it.", FAIL, mArg[aArgIndex].text);
		return NULL;
	}

	// Terminate the buffer, even if nothing was written into it:
	var_name[vni] = '\0';

	static Var empty_var(var_name);  // Must use var_name here.  See comment above for why.

	Var *found_var;
	if (!aCreateIfNecessary)
	{
		// Now we've dynamically build the variable name.  It's possible that the name is illegal,
		// so check that (the name is automatically checked by FindOrAddVar(), so we only need to
		// check it if we're not calling that):
		if (!Var::ValidateName(var_name, g_script.mIsReadyToExecute))
			return NULL; // Above already displayed error for us.
		if (found_var = g_script.FindVar(var_name)) // Assign.
			return found_var;
		// At this point, this is either a non-existent variable or a reserved/built-in variable
		// that was never statically referenced in the script (only dynamically), e.g. A_IPAddress%A_Index%
		if (Script::GetVarType(var_name) == VAR_NORMAL)
			// If not found: for performance reasons, don't create it because caller just wants an empty variable.
			return &empty_var;
		//else it's the clipboard or some other built-in variable, so continue onward so that the
		// variable gets created in the variable list, which is necessary to allow it to be properly
		// dereferenced, e.g. in a script consisting of only the following:
		// Loop, 4
		//     StringTrimRight, IP, A_IPAddress%A_Index%, 0
	}
	// Otherwise, aCreateIfNecessary is true or we want to create this variable unconditionally for the
	// reason described above.
	if (   !(found_var = g_script.FindOrAddVar(var_name))   )
		return NULL;  // Above will already have displayed the error.
	if (mArg[aArgIndex].type == ARG_TYPE_OUTPUT_VAR && VAR_IS_RESERVED(found_var))
	{
		LineError(ERR_VAR_IS_RESERVED, FAIL, var_name);
		return NULL;  // Don't return the var, preventing the caller from assigning to it.
	}
	else
		return found_var;
}



Var *Script::FindOrAddVar(char *aVarName, size_t aVarNameLength, Var *aSearchStart)
// Returns the Var whose name matches aVarName.  If it doesn't exist, it is created.
{
	if (!aVarName || !*aVarName) return NULL;
	Var *var_prev;
	Var *var = FindVar(aVarName, aVarNameLength, &var_prev, aSearchStart);  // 3rd param is a handle.
	if (var)
		return var;
	// Otherwise, no match found, so create a new var.  This will return NULL if there was a problem,
	// in which case AddVar() will already have displayed the error:
	return AddVar(aVarName, aVarNameLength, var_prev);
}



Var *Script::FindVar(char *aVarName, size_t aVarNameLength, Var **apVarPrev, Var *aSearchStart)
// Returns the Var whose name matches aVarName.  If it doesn't exist, NULL is returned.
// If caller provided a non-NULL apVarPrev, it will be given a pointer to the variable
// in the linked list after which the new variable would need to be inserted to maintain
// alphabetical order, which helps performance and allows the ListVars command display
// the variables in alphabetical order.
{
	if (!aVarName || !*aVarName) return NULL;
	if (!aVarNameLength) // Caller didn't specify, so use the entire string.
		aVarNameLength = strlen(aVarName);

	int result;
	Var *var, *var_prev;

	// If possible, start at aSearchStart, which is a performance hint from the caller:
	for (var_prev = NULL, var = aSearchStart ? aSearchStart : mFirstVar
		; var
		; var_prev = var, var = var->mNextVar)
	{
		result = strlicmp(aVarName, var->mName, (UINT)aVarNameLength);
		if (!result) // Match found.
		{
			if (apVarPrev) // Caller should ignore the value in this case, but set to NULL for consistency.
				*apVarPrev = NULL;
			return var;
		}
		if (result < 0) // Because the list is sorted, we now know that aVarName does not exist in the list.
		{
			// Set the output parameter, if present:
			if (apVarPrev)
				*apVarPrev = var_prev; // This is the var after which the new var should be insertted.
			return NULL;
		}
	}
	// Otherwise, no match found after going through the entire list, which means that aVarName
	// should be insertted at the end of the list by our caller (if it wishes to do that). So set the
	// output parameter, if present, to the last item (which will be NULL if the linked list is empty).
	if (apVarPrev)
		*apVarPrev = mLastVar;
	return NULL;
}



Var *Script::AddVar(char *aVarName, size_t aVarNameLength, Var *aVarPrev)
// This function should probably not be called by anyone except FindOrAddVar, which has already done
// the dupe-checking and is also providing us with the insertion point so that the list stays sorted.
// Returns the address of the new variable or NULL on failure.
// The caller must already have verified that this isn't a duplicate var.
{
	if (!aVarName || !*aVarName) return NULL;  // Should never happen, so just silently indicate failure.
	if (!aVarNameLength) // Caller didn't specify, so use the entire string.
		aVarNameLength = strlen(aVarName);

	if (aVarNameLength > MAX_VAR_NAME_LENGTH)
	{
		// Load-time callers should check for this.  But at runtime, it's possible for a dynamically
		// resolved variable name to be too long.  Note that aVarName should be the exact variable
		// name and does not need to be truncated to aVarNameLength whenever this error occurs
		// (i.e. at runtime):
		if (mIsReadyToExecute) // Runtime error.
			ScriptError("Variable name too long." ERR_ABORT, aVarName);
		else
			ScriptError("Variable name too long.", aVarName);
		return NULL;
	}

	// Make a temporary copy that includes only the first aVarNameLength characters from aVarName:
	char var_name[MAX_VAR_NAME_LENGTH + 1];
	strlcpy(var_name, aVarName, aVarNameLength + 1);  // +1 to convert length to size.

	if (!Var::ValidateName(var_name, mIsReadyToExecute))
		// Above already displayed error for us.  This can happen at loadtime or runtime (e.g. StringSplit).
		return NULL;

	// Allocate some dynamic memory to pass to the constructor:
	char *new_name = SimpleHeap::Malloc(var_name);
	if (!new_name)
		// It already displayed the error for us.  These mem errors are so unusual that we're not going
		// to bother varying the error message to include ERR_ABORT if this occurs during runtime.
		return NULL;

	VarTypeType var_type = GetVarType(new_name);

	Var *the_new_var = new Var(new_name, var_type);
	if (the_new_var == NULL)
	{
		ScriptError(ERR_OUTOFMEM);
		return NULL;
	}
	if (mFirstVar == NULL) // The list is empty, so this will be the first and last item.
		mFirstVar = mLastVar = the_new_var;
	else
	{
		if (!aVarPrev) // Caller is telling us to insert it at the beginning of the list.
		{
			// These two statements must be done in this order:
			the_new_var->mNextVar = mFirstVar;
			mFirstVar = the_new_var;
		}
		else if (aVarPrev == mLastVar) // Caller is telling us to insert it at the end of the list.
		{
			// These two statements must be done in this order:
			mLastVar->mNextVar = the_new_var;
			mLastVar = the_new_var;
		}
		else // Caller is telling us to insert it somewhere in the middle of the list.
		{
			// These two statements must be done in this order:
			the_new_var->mNextVar = aVarPrev->mNextVar;
			aVarPrev->mNextVar = the_new_var;
		}
	}

	++mVarCount;
	return the_new_var;
}



VarTypes Script::GetVarType(char *aVarName)
{
	if (toupper(*aVarName) != 'A' || *(aVarName + 1) != '_')  // This check helps average performance.
	{
		if (!stricmp(aVarName, "true")) return VAR_TRUE;
		if (!stricmp(aVarName, "false")) return VAR_FALSE;
		if (!stricmp(aVarName, "clipboard")) return VAR_CLIPBOARD;
		// Otherwise:
		return VAR_NORMAL;
	}

	// Otherwise, aVarName begins with "A_", so it's probably one of the built-in variables.
	// Keeping the most common ones near the top helps performance a little:

	if (!stricmp(aVarName, "A_YYYY") || !stricmp(aVarName, "A_Year")) return VAR_YYYY;
	if (!stricmp(aVarName, "A_MMMM")) return VAR_MMMM; // Long name of month.
	if (!stricmp(aVarName, "A_MMM")) return VAR_MMM;   // 3-char abbrev. month name.
	if (!stricmp(aVarName, "A_MM") || !stricmp(aVarName, "A_Mon")) return VAR_MM;  // 01 thru 12
	if (!stricmp(aVarName, "A_DDDD")) return VAR_DDDD; // Name of weekday, e.g. Sunday
	if (!stricmp(aVarName, "A_DDD")) return VAR_DDD;   // Abbrev., e.g. Sun
	if (!stricmp(aVarName, "A_DD") || !stricmp(aVarName, "A_Mday")) return VAR_DD; // 01 thru 31
	if (!stricmp(aVarName, "A_Wday")) return VAR_WDAY;
	if (!stricmp(aVarName, "A_Yday")) return VAR_YDAY;
	if (!stricmp(aVarName, "A_Yweek")) return VAR_YWEEK;

	if (!stricmp(aVarName, "A_Hour")) return VAR_HOUR;
	if (!stricmp(aVarName, "A_Min")) return VAR_MIN;
	if (!stricmp(aVarName, "A_Sec")) return VAR_SEC;
	if (!stricmp(aVarName, "A_TickCount")) return VAR_TICKCOUNT;
	if (!stricmp(aVarName, "A_Now")) return VAR_NOW;
	if (!stricmp(aVarName, "A_NowUTC")) return VAR_NOWUTC;

	if (!stricmp(aVarName, "A_WorkingDir")) return VAR_WORKINGDIR;
	if (!stricmp(aVarName, "A_ScriptName")) return VAR_SCRIPTNAME;
	if (!stricmp(aVarName, "A_ScriptDir")) return VAR_SCRIPTDIR;
	if (!stricmp(aVarName, "A_ScriptFullPath")) return VAR_SCRIPTFULLPATH;

	if (!stricmp(aVarName, "A_BatchLines") || !stricmp(aVarName, "A_NumBatchLines")) return VAR_BATCHLINES;
	if (!stricmp(aVarName, "A_TitleMatchMode")) return VAR_TITLEMATCHMODE;
	if (!stricmp(aVarName, "A_TitleMatchModeSpeed")) return VAR_TITLEMATCHMODESPEED;
	if (!stricmp(aVarName, "A_DetectHiddenWindows")) return VAR_DETECTHIDDENWINDOWS;
	if (!stricmp(aVarName, "A_DetectHiddenText")) return VAR_DETECTHIDDENTEXT;
	if (!stricmp(aVarName, "A_AutoTrim")) return VAR_AUTOTRIM;
	if (!stricmp(aVarName, "A_StringCaseSense")) return VAR_STRINGCASESENSE;
	if (!stricmp(aVarName, "A_FormatInteger")) return VAR_FORMATINTEGER;
	if (!stricmp(aVarName, "A_FormatFloat")) return VAR_FORMATFLOAT;
	if (!stricmp(aVarName, "A_KeyDelay")) return VAR_KEYDELAY;
	if (!stricmp(aVarName, "A_WinDelay")) return VAR_WINDELAY;
	if (!stricmp(aVarName, "A_ControlDelay")) return VAR_CONTROLDELAY;
	if (!stricmp(aVarName, "A_MouseDelay")) return VAR_MOUSEDELAY;
	if (!stricmp(aVarName, "A_DefaultMouseSpeed")) return VAR_DEFAULTMOUSESPEED;

	if (!stricmp(aVarName, "A_IconHidden")) return VAR_ICONHIDDEN;
	if (!stricmp(aVarName, "A_IconTip")) return VAR_ICONTIP;
	if (!stricmp(aVarName, "A_IconFile")) return VAR_ICONFILE;
	if (!stricmp(aVarName, "A_IconNumber")) return VAR_ICONNUMBER;

	if (!stricmp(aVarName, "A_ExitReason")) return VAR_EXITREASON;

	if (!stricmp(aVarName, "A_OStype")) return VAR_OSTYPE;
	if (!stricmp(aVarName, "A_OSversion")) return VAR_OSVERSION;
	if (!stricmp(aVarName, "A_Language")) return VAR_LANGUAGE;
	if (!stricmp(aVarName, "A_ComputerName")) return VAR_COMPUTERNAME;
	if (!stricmp(aVarName, "A_UserName")) return VAR_USERNAME;

	if (!stricmp(aVarName, "A_WinDir")) return VAR_WINDIR;
	if (!stricmp(aVarName, "A_ProgramFiles")) return VAR_PROGRAMFILES;
	if (!stricmp(aVarName, "A_Desktop")) return VAR_DESKTOP;
	if (!stricmp(aVarName, "A_DesktopCommon")) return VAR_DESKTOPCOMMON;
	if (!stricmp(aVarName, "A_StartMenu")) return VAR_STARTMENU;
	if (!stricmp(aVarName, "A_StartMenuCommon")) return VAR_STARTMENUCOMMON;
	if (!stricmp(aVarName, "A_Programs")) return VAR_PROGRAMS;
	if (!stricmp(aVarName, "A_ProgramsCommon")) return VAR_PROGRAMSCOMMON;
	if (!stricmp(aVarName, "A_Startup")) return VAR_STARTUP;
	if (!stricmp(aVarName, "A_StartupCommon")) return VAR_STARTUPCOMMON;
	if (!stricmp(aVarName, "A_MyDocuments")) return VAR_MYDOCUMENTS;

	if (!stricmp(aVarName, "A_IsAdmin")) return VAR_ISADMIN;
	if (!stricmp(aVarName, "A_Cursor")) return VAR_CURSOR;
	if (!stricmp(aVarName, "A_CaretX")) return VAR_CARETX;
	if (!stricmp(aVarName, "A_CaretY")) return VAR_CARETY;
	if (!stricmp(aVarName, "A_ScreenWidth")) return VAR_SCREENWIDTH;
	if (!stricmp(aVarName, "A_ScreenHeight")) return VAR_SCREENHEIGHT;
	if (!stricmp(aVarName, "A_IPAddress1")) return VAR_IPADDRESS1;
	if (!stricmp(aVarName, "A_IPAddress2")) return VAR_IPADDRESS2;
	if (!stricmp(aVarName, "A_IPAddress3")) return VAR_IPADDRESS3;
	if (!stricmp(aVarName, "A_IPAddress4")) return VAR_IPADDRESS4;

	if (!stricmp(aVarName, "A_LoopFileName")) return VAR_LOOPFILENAME;
	if (!stricmp(aVarName, "A_LoopFileShortName")) return VAR_LOOPFILESHORTNAME;
	if (!stricmp(aVarName, "A_LoopFileDir")) return VAR_LOOPFILEDIR;
	if (!stricmp(aVarName, "A_LoopFileFullPath")) return VAR_LOOPFILEFULLPATH;
	if (!stricmp(aVarName, "A_LoopFileShortPath")) return VAR_LOOPFILESHORTPATH;
	if (!stricmp(aVarName, "A_LoopFileTimeModified")) return VAR_LOOPFILETIMEMODIFIED;
	if (!stricmp(aVarName, "A_LoopFileTimeCreated")) return VAR_LOOPFILETIMECREATED;
	if (!stricmp(aVarName, "A_LoopFileTimeAccessed")) return VAR_LOOPFILETIMEACCESSED;
	if (!stricmp(aVarName, "A_LoopFileAttrib")) return VAR_LOOPFILEATTRIB;
	if (!stricmp(aVarName, "A_LoopFileSize")) return VAR_LOOPFILESIZE;
	if (!stricmp(aVarName, "A_LoopFileSizeKB")) return VAR_LOOPFILESIZEKB;
	if (!stricmp(aVarName, "A_LoopFileSizeMB")) return VAR_LOOPFILESIZEMB;

	if (!stricmp(aVarName, "A_LoopRegType")) return VAR_LOOPREGTYPE;
	if (!stricmp(aVarName, "A_LoopRegKey")) return VAR_LOOPREGKEY;
	if (!stricmp(aVarName, "A_LoopRegSubKey")) return VAR_LOOPREGSUBKEY;
	if (!stricmp(aVarName, "A_LoopRegName")) return VAR_LOOPREGNAME;
	if (!stricmp(aVarName, "A_LoopRegTimeModified")) return VAR_LOOPREGTIMEMODIFIED;

	if (!stricmp(aVarName, "A_LoopReadLine")) return VAR_LOOPREADLINE;
	if (!stricmp(aVarName, "A_LoopField")) return VAR_LOOPFIELD;
	if (!stricmp(aVarName, "A_Index")) return VAR_INDEX;  // A short name since it maybe be typed so often.

	if (!stricmp(aVarName, "A_ThisMenuItem")) return VAR_THISMENUITEM;
	if (!stricmp(aVarName, "A_ThisMenuItemPos")) return VAR_THISMENUITEMPOS;
	if (!stricmp(aVarName, "A_ThisMenu")) return VAR_THISMENU;
	if (!stricmp(aVarName, "A_ThisHotkey")) return VAR_THISHOTKEY;
	if (!stricmp(aVarName, "A_PriorHotkey")) return VAR_PRIORHOTKEY;
	if (!stricmp(aVarName, "A_TimeSinceThisHotkey")) return VAR_TIMESINCETHISHOTKEY;
	if (!stricmp(aVarName, "A_TimeSincePriorHotkey")) return VAR_TIMESINCEPRIORHOTKEY;
	if (!stricmp(aVarName, "A_EndChar")) return VAR_ENDCHAR;
	if (!stricmp(aVarName, "A_Gui")) return VAR_GUI;
	if (!stricmp(aVarName, "A_GuiControl")) return VAR_GUICONTROL;
	if (!stricmp(aVarName, "A_GuiControlEvent")) return VAR_GUICONTROLEVENT;
	if (!stricmp(aVarName, "A_GuiWidth")) return VAR_GUIWIDTH;
	if (!stricmp(aVarName, "A_GuiHeight")) return VAR_GUIHEIGHT;

	if (!stricmp(aVarName, "A_TimeIdle")) return VAR_TIMEIDLE;
	if (!stricmp(aVarName, "A_TimeIdlePhysical")) return VAR_TIMEIDLEPHYSICAL;
	if (!stricmp(aVarName, "A_Space")) return VAR_SPACE;
	if (!stricmp(aVarName, "A_Tab")) return VAR_TAB;
	if (!stricmp(aVarName, "A_AhkVersion")) return VAR_AHKVERSION;

	// Since above didn't return:
	return VAR_NORMAL;
}



WinGroup *Script::FindOrAddGroup(char *aGroupName)
// Caller must ensure that aGroupName isn't NULL.  But if it's the empty string, NULL is returned.
// Returns the Group whose name matches aGroupName.  If it doesn't exist, it is created.
{
	if (!*aGroupName)
		return NULL;
	for (WinGroup *group = mFirstGroup; group != NULL; group = group->mNextGroup)
		if (!stricmp(group->mName, aGroupName)) // Match found.
			return group;
	// Otherwise, no match found, so create a new group.
	if (AddGroup(aGroupName) != OK)
		return NULL;
	return mLastGroup;
}



ResultType Script::AddGroup(char *aGroupName)
// Returns OK or FAIL.
// The caller must already have verfied that this isn't a duplicate group.
{
	if (strlen(aGroupName) > MAX_VAR_NAME_LENGTH)
		return ScriptError("Group name too long.");
	if (!Var::ValidateName(aGroupName)) // Seems best to use same validation as var names.
		return FAIL; // Above already displayed error for us.

	char *new_name = SimpleHeap::Malloc(aGroupName);
	if (!new_name)
		return FAIL;  // It already displayed the error for us.

	WinGroup *the_new_group = new WinGroup(new_name);
	if (the_new_group == NULL)
		return ScriptError(ERR_OUTOFMEM);
	if (mFirstGroup == NULL)
		mFirstGroup = mLastGroup = the_new_group;
	else
	{
		mLastGroup->mNextGroup = the_new_group;
		// This must be done after the above:
		mLastGroup = the_new_group;
	}
	++mGroupCount;
	return OK;
}



Line *Script::PreparseBlocks(Line *aStartingLine, bool aFindBlockEnd, Line *aParentLine)
// aFindBlockEnd should be true, only when this function is called
// by itself.  The end of this function relies upon this definition.
// Will return NULL to the top-level caller if there's an error, or if
// mLastLine is NULL (i.e. the script is empty).
{
	// Not thread-safe, so this can only parse one script at a time.
	// Not a problem for the foreseeable future:
	static nest_level; // Level zero is the outermost one: outside all blocks.
	static bool abort;
	if (aParentLine == NULL)
	{
		// We were called from outside, not recursively, so init these.  This is
		// very important if this function is ever to be called from outside
		// more than once, even though it isn't currently:
		nest_level = 0;
		abort = false;
	}

	// Don't check aStartingLine here at top: only do it at the bottom
	// for its differing return values.
	for (Line *line = aStartingLine; line != NULL;)
	{
		// All lines in our recursion layer are assigned to the block that the caller specified:
		if (line->mParentLine == NULL) // i.e. don't do it if it's already "owned" by an IF or ELSE.
			line->mParentLine = aParentLine; // Can be NULL.

		if (ACT_IS_IF(line->mActionType) || line->mActionType == ACT_ELSE
			|| line->mActionType == ACT_LOOP || line->mActionType == ACT_REPEAT)
		{
			// Make the line immediately following each ELSE, IF or LOOP be enclosed by that stmt.
			// This is done to make it illegal for a Goto or Gosub to jump into a deeper layer,
			// such as in this example:
			// #y::
			// ifwinexist, pad
			// {
			//    goto, label1
			//    ifwinexist, pad
			//    label1:
			//    ; With or without the enclosing block, the goto would still go to an illegal place
			//    ; in the below, resulting in an "unexpected else" error:
			//    {
			//	     msgbox, ifaction
			//    } ; not necessary to make this line enclosed by the if because labels can't point to it?
			// else
			//    msgbox, elseaction
			// }
			// return

			// In this case, the loader should have already ensured that line->mNextLine is not NULL:
			line->mNextLine->mParentLine = line;
			// Go onto the IF's or ELSE's action in case it too is an IF, rather than skipping over it:
			line = line->mNextLine;
			continue;
		}

		switch (line->mActionType)
		{
		case ACT_BLOCK_BEGIN:
			// Some insane limit too large to ever likely be exceeded, yet small enough not
			// to be a risk of stack overflow when recursing in ExecUntil().  Mostly, this is
			// here to reduce the chance of a program crash if a binary file, a corrupted file,
			// or something unexpected has been loaded as a script when it shouldn't have been.
			// Update: Increased the limit from 100 to 1000 so that large "else if" ladders
			// can be constructed.  Going much larger than 1000 seems unwise since ExecUntil()
			// will have to recurse for each nest-level, possibly resulting in stack overflow
			// if things get too deep (though I think most/all(?) versions of Windows will
			// dynamically grow the stack to try to keep up):
			if (nest_level > 1000)
			{
				abort = true; // So that the caller doesn't also report an error.
				return line->PreparseError("Nesting this deep might cause a stack overflow so is not allowed.");
			}
			// Since the current convention is to store the line *after* the
			// BLOCK_END as the BLOCK_BEGIN's related line, that line can
			// be legitimately NULL if this block's BLOCK_END is the last
			// line in the script.  So it's up to the called function
			// to report an error if it never finds a BLOCK_END for us.
			// UPDATE: The design requires that we do it here instead:
			++nest_level;
			if (NULL == (line->mRelatedLine = PreparseBlocks(line->mNextLine, 1, line)))
				if (abort) // the above call already reported the error.
					return NULL;
				else
					return line->PreparseError("This open block is never closed."
						"  If this block is for a REPEAT command, its ENDREPEAT may be missing.");
			--nest_level;
			// The convention is to have the BLOCK_BEGIN's related_line
			// point to the line *after* the BLOCK_END.
			line->mRelatedLine = line->mRelatedLine->mNextLine;  // Might be NULL now.
			// Otherwise, since any blocks contained inside this one would already
			// have been handled by the recursion in the above call, continue searching
			// from the end of this block:
			line = line->mRelatedLine; // If NULL, the loop-condition will catch it.
			break;
		case ACT_BLOCK_END:
			// Return NULL (failure) if the end was found but we weren't looking for one
			// (i.e. it's an orphan).  Otherwise return the line after the block_end line,
			// which will become the caller's mRelatedLine.  UPDATE: Return the
			// END_BLOCK line itself so that the caller can differentiate between
			// a NULL due to end-of-script and a NULL caused by an error:
			return aFindBlockEnd ? line
				: line->PreparseError("Attempt to close a non-existent block.");
		default: // Continue line-by-line.
			line = line->mNextLine;
		} // switch()
	} // for()
	// End of script has been reached.  <line> is now NULL so don't attempt to dereference it.
	// If we were still looking for an EndBlock to match up with a begin, that's an error.
	// Don't report the error here because we don't know which begin-block is waiting
	// for an end (the caller knows and must report the error).  UPDATE: Must report
	// the error here (see comments further above for explanation).   UPDATE #2: Changed
	// it again: Now we let the caller handle it again:
	if (aFindBlockEnd)
		//return mLastLine->PreparseError("The script ends while a block is still open (missing }).");
		return NULL;
	// If no error, return something non-NULL to indicate success to the top-level caller.
	// We know we're returning to the top-level caller because aFindBlockEnd is only true
	// when we're recursed, and in that case the above would have returned.  Thus,
	// we're not recursed upon reaching this line:
	return mLastLine;
}



Line *Script::PreparseIfElse(Line *aStartingLine, ExecUntilMode aMode, AttributeType aLoopTypeFile
	, AttributeType aLoopTypeReg, AttributeType aLoopTypeRead, AttributeType aLoopTypeParse)
// Zero is the default for aMode, otherwise:
// Will return NULL to the top-level caller if there's an error, or if
// mLastLine is NULL (i.e. the script is empty).
// Note: This function should be called with aMode == ONLY_ONE_LINE
// only when aStartingLine's ActionType is something recursable such
// as IF and BEGIN_BLOCK.  Otherwise, it won't return after only one line.
{
	// Don't check aStartingLine here at top: only do it at the bottom
	// for it's differing return values.
	Line *line_temp;
	// Although rare, a statement can be enclosed in more than one type of special loop,
	// e.g. both a file-loop and a reg-loop:
	AttributeType loop_type_file, loop_type_reg, loop_type_read, loop_type_parse;
	for (Line *line = aStartingLine; line != NULL;)
	{
		if (ACT_IS_IF(line->mActionType) || line->mActionType == ACT_LOOP || line->mActionType == ACT_REPEAT)
		{
			// ActionType is an IF or a LOOP.
			line_temp = line->mNextLine;  // line_temp is now this IF's or LOOP's action-line.
			if (line_temp == NULL) // This is an orphan IF/LOOP (has no action-line) at the end of the script.
				// Update: this is now impossible because all scripts end in ACT_EXIT.
				return line->PreparseError("This if-statement or loop has no action.");
			if (line_temp->mActionType == ACT_ELSE || line_temp->mActionType == ACT_BLOCK_END)
				return line->PreparseError("The line beneath this IF or LOOP is an invalid action.");

			// We're checking for ATTR_LOOP_FILE here to detect whether qualified commands enclosed
			// in a true file loop are allowed to omit their filename paremeter:
			loop_type_file = ATTR_NONE;
			if (aLoopTypeFile == ATTR_LOOP_FILE || line->mAttribute == ATTR_LOOP_FILE)
				// i.e. if either one is a file-loop, that's enough to establish
				// the fact that we're in a file loop.
				loop_type_file = ATTR_LOOP_FILE;
			else if (aLoopTypeFile == ATTR_LOOP_UNKNOWN || line->mAttribute == ATTR_LOOP_UNKNOWN)
				// ATTR_LOOP_UNKNOWN takes precedence over ATTR_LOOP_NORMAL because
				// we can't be sure if we're in a file loop, but it's correct to
				// assume that we are (otherwise, unwarranted syntax errors may be reported
				// later on in here).
				loop_type_file = ATTR_LOOP_UNKNOWN;
			else if (aLoopTypeFile == ATTR_LOOP_NORMAL || line->mAttribute == ATTR_LOOP_NORMAL)
				loop_type_file = ATTR_LOOP_NORMAL;

			// The section is the same as above except for registry vs. file loops:
			loop_type_reg = ATTR_NONE;
			if (aLoopTypeReg == ATTR_LOOP_REG || line->mAttribute == ATTR_LOOP_REG)
				loop_type_reg = ATTR_LOOP_REG;
			else if (aLoopTypeReg == ATTR_LOOP_UNKNOWN || line->mAttribute == ATTR_LOOP_UNKNOWN)
				loop_type_reg = ATTR_LOOP_UNKNOWN;
			else if (aLoopTypeReg == ATTR_LOOP_NORMAL || line->mAttribute == ATTR_LOOP_NORMAL)
				loop_type_reg = ATTR_LOOP_NORMAL;

			// Same as above except for READ-FILE loops:
			loop_type_read = ATTR_NONE;
			if (aLoopTypeRead == ATTR_LOOP_READ_FILE || line->mAttribute == ATTR_LOOP_READ_FILE)
				loop_type_read = ATTR_LOOP_READ_FILE;
			else if (aLoopTypeRead == ATTR_LOOP_UNKNOWN || line->mAttribute == ATTR_LOOP_UNKNOWN)
				loop_type_read = ATTR_LOOP_UNKNOWN;
			else if (aLoopTypeRead == ATTR_LOOP_NORMAL || line->mAttribute == ATTR_LOOP_NORMAL)
				loop_type_read = ATTR_LOOP_NORMAL;

			// Same as above except for PARSING loops:
			loop_type_parse = ATTR_NONE;
			if (aLoopTypeParse == ATTR_LOOP_PARSE || line->mAttribute == ATTR_LOOP_PARSE)
				loop_type_parse = ATTR_LOOP_PARSE;
			else if (aLoopTypeParse == ATTR_LOOP_UNKNOWN || line->mAttribute == ATTR_LOOP_UNKNOWN)
				loop_type_parse = ATTR_LOOP_UNKNOWN;
			else if (aLoopTypeParse == ATTR_LOOP_NORMAL || line->mAttribute == ATTR_LOOP_NORMAL)
				loop_type_parse = ATTR_LOOP_NORMAL;

			// Check if the IF's action-line is something we want to recurse.  UPDATE: Always
			// recurse because other line types, such as Goto and Gosub, need to be preparsed
			// by this function even if they are the single-line actions of an IF or an ELSE:
			// Recurse this line rather than the next because we want
			// the called function to recurse again if this line is a ACT_BLOCK_BEGIN
			// or is itself an IF:
			line_temp = PreparseIfElse(line_temp, ONLY_ONE_LINE, loop_type_file, loop_type_reg, loop_type_read
				, loop_type_parse);
			// If not end-of-script or error, line_temp is now either:
			// 1) If this if's/loop's action was a BEGIN_BLOCK: The line after the end of the block.
			// 2) If this if's/loop's action was another IF or LOOP:
			//    a) the line after that if's else's action; or (if it doesn't have one):
			//    b) the line after that if's/loop's action
			// 3) If this if's/loop's action was some single-line action: the line after that action.
			// In all of the above cases, line_temp is now the line where we
			// would expect to find an ELSE for this IF, if it has one.

			// Now the above has ensured that line_temp is this line's else, if it has one.
			// Note: line_temp will be NULL if the end of the script has been reached.
			// UPDATE: That can't happen now because all scripts end in ACT_EXIT:
			if (line_temp == NULL) // Error or end-of-script was reached.
				return NULL;

			// Temporary check for something that should be impossible if design is correct:
			if (line->mRelatedLine != NULL)
				return line->PreparseError("This if-statement or LOOP unexpectedly already had an ELSE or end-point.");
			// Set it to the else's action, rather than the else itself, since the else itself
			// is never needed during execution.  UPDATE: No, instead set it to the ELSE itself
			// (if it has one) since we jump here at runtime when the IF is finished (whether
			// it's condition was true or false), thus skipping over any nested IF's that
			// aren't in blocks beneath it.  If there's no ELSE, the below value serves as
			// the jumppoint we go to when the if-statement is finished.  Example:
			// if x
			//   if y
			//     if z
			//       action1
			//     else
			//       action2
			// action3
			// x's jumppoint should be action3 so that all the nested if's
			// under the first one can be skipped after the "if x" line is recursively
			// evaluated.  Because of this behavior, all IFs will have a related line
			// with the possibly exception of the very last if-statement in the script
			// (which is possible only if the script doesn't end in a Return or Exit).
			line->mRelatedLine = line_temp;  // Even if <line> is a LOOP and line_temp and else?

			// Even if aMode == ONLY_ONE_LINE, an IF and its ELSE count as a single
			// statement (one line) due to its very nature (at least for this purpose),
			// so always continue on to evaluate the IF's ELSE, if present:
			if (line_temp->mActionType == ACT_ELSE)
			{
				if (line->mActionType == ACT_LOOP || line->mActionType == ACT_REPEAT)
				{
					 // this can't be our else, so let the caller handle it.
					if (aMode != ONLY_ONE_LINE)
						// This ELSE was encountered while sequentially scanning the contents
						// of a block or at the otuermost nesting layer.  More thought is required
						// to verify this is correct:
						return line_temp->PreparseError(ERR_ELSE_WITH_NO_IF);
					// Let the caller handle this else, since it can't be ours:
					return line_temp;
				}
				// Now use line vs. line_temp to hold the new values, so that line_temp
				// stays as a marker to the ELSE line itself:
				line = line_temp->mNextLine;  // Set it to the else's action line.
				if (line == NULL) // An else with no action.
					// Update: this is now impossible because all scripts end in ACT_EXIT.
					return line_temp->PreparseError("This ELSE has no action.");
				if (line->mActionType == ACT_ELSE || line->mActionType == ACT_BLOCK_END)
					return line_temp->PreparseError("The line beneath this ELSE is an invalid action.");
				// Assign to line rather than line_temp:
				line = PreparseIfElse(line, ONLY_ONE_LINE, aLoopTypeFile, aLoopTypeReg, aLoopTypeRead
					, aLoopTypeParse);
				if (line == NULL)
					return NULL; // Error or end-of-script.
				// Set this ELSE's jumppoint.  This is similar to the jumppoint set for
				// an ELSEless IF, so see related comments above:
				line_temp->mRelatedLine = line;
			}
			else // line doesn't have an else, so just continue processing from line_temp's position
				line = line_temp;

			// Both cases above have ensured that line is now the first line beyond the
			// scope of the if-statement and that of any ELSE it may have.

			if (aMode == ONLY_ONE_LINE) // Return the next unprocessed line to the caller.
				return line;
			// Otherwise, continue processing at line's new location:
			continue;
		} // ActionType is "IF".

		// Since above didn't continue, do the switch:
		switch (line->mActionType)
		{
		case ACT_BLOCK_BEGIN:
			line = PreparseIfElse(line->mNextLine, UNTIL_BLOCK_END, aLoopTypeFile, aLoopTypeReg, aLoopTypeRead
				, aLoopTypeParse);
			// line is now either NULL due to an error, or the location
			// of the END_BLOCK itself.
			if (line == NULL)
				return NULL; // Error.
			break;
		case ACT_BLOCK_END:
			if (aMode == ONLY_ONE_LINE)
				 // Syntax error.  The caller would never expect this single-line to be an
				 // end-block.  UPDATE: I think this is impossible because callers only use
				 // aMode == ONLY_ONE_LINE when aStartingLine's ActionType is already
				 // known to be an IF or a BLOCK_BEGIN:
				 return line->PreparseError("Unexpected end-of-block (single).");
			if (UNTIL_BLOCK_END)
				// Return line rather than line->mNextLine because, if we're at the end of
				// the script, it's up to the caller to differentiate between that condition
				// and the condition where NULL is an error indicator.
				return line;
			// Otherwise, we found an end-block we weren't looking for.  This should be
			// impossible since the block pre-parsing already balanced all the blocks?
			return line->PreparseError("Unexpected end-of-block (multi).");
		case ACT_BREAK:
		case ACT_CONTINUE:
			if (!aLoopTypeFile && !aLoopTypeReg && !aLoopTypeRead && !aLoopTypeParse)
				return line->PreparseError("This break or continue statement is not enclosed by a loop.");
			break;

		case ACT_GOTO:  // These two must be done here (i.e. *after* all the script lines have been added),
		case ACT_GOSUB: // so that labels both above and below each Gosub/Goto can be resolved.
			if (line->ArgHasDeref(1))
				// Since the jump-point contains a deref, it must be resolved at runtime:
				line->mRelatedLine = NULL;
			else
				if (!line->GetJumpTarget(false))
					return NULL; // Error was already displayed by called function.
			break;

		// These next 4 must also be done here (i.e. *after* all the script lines have been added),
		// so that labels both above and below this line can be resolved:
		case ACT_ONEXIT:
			if (*LINE_RAW_ARG1 && !line->ArgHasDeref(1))
				if (   !(line->mAttribute = FindLabel(LINE_RAW_ARG1))   )
					return line->PreparseError(ERR_ONEXIT_LABEL);
			break;

		case ACT_HOTKEY:
			if (*LINE_RAW_ARG2 && !line->ArgHasDeref(2))
				if (   !(line->mAttribute = FindLabel(LINE_RAW_ARG2))   )
					if (!Hotkey::ConvertAltTab(LINE_RAW_ARG2, true))
						return line->PreparseError(ERR_HOTKEY_LABEL);
			break;

		case ACT_SETTIMER:
			if (!line->ArgHasDeref(1))
				if (   !(line->mAttribute = FindLabel(LINE_RAW_ARG1))   )
					return line->PreparseError(ERR_SETTIMER);
			if (*LINE_RAW_ARG2 && !line->ArgHasDeref(2))
				if (!Line::ConvertOnOff(LINE_RAW_ARG2) && !IsPureNumeric(LINE_RAW_ARG2))
					return line->PreparseError("If not blank or a variable reference, parameter #2"
						" must be either a positive integer or the word ON or OFF.");
			break;

		case ACT_GROUPADD: // This must be done here because it relies on all other lines already having been added.
			if (*LINE_RAW_ARG4 && !line->ArgHasDeref(4))
			{
				// If the label name was contained in a variable, that label is now resolved and cannot
				// be changed.  This is in contrast to something like "Gosub, %MyLabel%" where a change in
				// the value of MyLabel will change the behavior of the Gosub at runtime:
				Label *label = FindLabel(LINE_RAW_ARG4);
				if (!label)
					return line->PreparseError(ERR_GROUPADD_LABEL);
				line->mRelatedLine = label->mJumpToLine; // The script loader has ensured that this can't be NULL.
				// Can't do this because the current line won't be the launching point for the
				// Gosub.  Instead, the launching point will be the GroupActivate rather than the
				// GroupAdd, so it will be checked by the GroupActivate or not at all (since it's
				// not that important in the case of a Gosub -- it's mostly for Goto's):
				//return IsJumpValid(label->mJumpToLine);
			}
			break;

		case ACT_ELSE:
			// Should never happen because the part that handles the if's, above, should find
			// all the elses and handle them.  UPDATE: This happens if there's
			// an extra ELSE in this scope level that has no IF:
			return line->PreparseError(ERR_ELSE_WITH_NO_IF);
		} // switch()

		line = line->mNextLine; // If NULL due to physical end-of-script, the for-loop's condition will catch it.
		if (aMode == ONLY_ONE_LINE) // Return the next unprocessed line to the caller.
			// In this case, line shouldn't be (and probably can't be?) NULL because the line after
			// a single-line action shouldn't be the physical end of the script.  That's because
			// the loader has ensured that all scripts now end in ACT_EXIT.  And that final
			// ACT_EXIT should never be parsed here in ONLY_ONE_LINE mode because the only time
			// that mode is used is for the action of an IF, an ELSE, or possibly a LOOP.
			// In all of those cases, the final ACT_EXIT line in the script (which is explicitly
			// insertted by the loader) cannot be the line that was just processed by the
			// switch().  Therefore, the above assignment should not have set line to NULL
			// (which is good because NULL would probably be construed as "failure" by our
			// caller in this case):
			return line;
		// else just continue the for-loop at the new value of line.
	} // for()

	// End of script has been reached.  line is now NULL so don't dereference it.

	// If we were still looking for an EndBlock to match up with a begin, that's an error.
	// This indicates that the at least one BLOCK_BEGIN is missing a BLOCK_END.
	// However, since the blocks were already balanced by the block pre-parsing function,
	// this should be impossible unless the design of this function is flawed.
	if (aMode == UNTIL_BLOCK_END)
#ifdef _DEBUG
		return mLastLine->PreparseError("The script ended while a block was still open."); // This is a bug because the preparser already verified all blocks are balanced.
#else
		return NULL; // Shouldn't happen, so just return failure.
#endif

	// If we were told to process a single line, we were recursed and it should have returned above,
	// so it's an error here (can happen if we were called with aStartingLine == NULL?):
	if (aMode == ONLY_ONE_LINE)
		return mLastLine->PreparseError("The script ended while an action was still expected.");

	// Otherwise, return something non-NULL to indicate success to the top-level caller:
	return mLastLine;
}


//-------------------------------------------------------------------------------------

// Init static vars:
Line *Line::sLog[] = {NULL};  // Initialize all the array elements.
DWORD Line::sLogTick[]; // No initialization needed.
int Line::sLogNext = 0;  // Start at the first element.

char *Line::sSourceFile[MAX_SCRIPT_FILES]; // No init needed.
int Line::nSourceFiles = 0;  // Zero source files initially.  The main script will be the first.

char *Line::sDerefBuf = NULL;  // Buffer to hold the values of any args that need to be dereferenced.
char *Line::sDerefBufMarker = NULL;
size_t Line::sDerefBufSize = 0;
char *Line::sArgDeref[MAX_ARGS]; // No init needed.


ResultType Line::ExecUntil(ExecUntilMode aMode, Line **apJumpToLine, WIN32_FIND_DATA *aCurrentFile
	, RegItemStruct *aCurrentRegItem, LoopReadFileStruct *aCurrentReadFile, char *aCurrentField
	, __int64 aCurrentLoopIteration)
// Start executing at "this" line, stop when aMode indicates.
// RECURSIVE: Handles all lines that involve flow-control.
// aMode can be UNTIL_RETURN, UNTIL_BLOCK_END, ONLY_ONE_LINE.
// Returns FAIL, OK, EARLY_RETURN, or EARLY_EXIT.
// apJumpToLine is a pointer to Line-ptr (handle), which is an output param.  If NULL,
// the caller is indicating it doesn't need this value, so it won't (and can't) be set by
// the called recursion layer.
{
	if (apJumpToLine != NULL)
		// Important to init, since most of the time it will keep this value.
		// Tells caller that no jump is required (default):
		*apJumpToLine = NULL;

	Line *jump_to_line; // Don't use *apJumpToLine because it might not exist.
	Line *jump_target;  // For use with Gosub & Goto
	ResultType if_condition, result;
	LONG_OPERATION_INIT

	for (Line *line = this; line != NULL;)
	{
		// If a previous command (line) had the clipboard open, perhaps because it directly accessed
		// the clipboard via Var::Contents(), we close it here for performance reasons (see notes
		// in Clipboard::Open() for details):
		CLOSE_CLIPBOARD_IF_OPEN;

		// The below must be done at least when Hotkey::HookIsActive() is true, but is currently
		// always done since it's a very low overhead call, and has the side-benefit of making
		// the app maximally responsive when the script is busy during high BatchLines.
		// This low-overhead call achieves at least two purposes optimally:
		// 1) Keyboard and mouse lag is minimized when the hook(s) are installed, since this single
		//    Peek() is apparently enough to route all pending input to the hooks (though it's inexplicable
		//    why calling MsgSleep(-1) does not achieve this goal, since it too does a Peek().
		//    Nevertheless, that is the testing result that was obtained: the mouse cursor lagged
		//    in tight script loops even when MsgSleep(-1) or (0) was called ever 10ms or so.
		// 2) The app is maximally responsive while executing with a high or infinite BatchLines.
		// 3) Hotkeys are maximally responsive.  For example, if a user has game hotkeys, using
		//    a GetTickCount() method (which very slightly improves performance by cutting back on
		//    the number of Peek() calls) would introduce up to 10ms of delay before the hotkey
		//    finally takes effect.  10ms can be significant in games, where ping (latency) itself
		//    can sometimes be only 10 or 20ms.
		// 4) Timed subroutines are run as consistently as possible (to help with this, a check
		//    similar to the below is also done for single commmands that take a long time, such
		//    as URLDownloadToFile, FileSetAttrib, etc.
		// UPDATE: It looks like PeekMessage() yields CPU time automatically, similar to a
		// Sleep(0).  Since this would make scripts slow to a crawl, only do the Peek() every 5ms
		// or so (though the timer grandurity is 10ms on mosts OSes, so that's the true interval):
		LONG_OPERATION_UPDATE

		// If interruptions are currently forbidden, it's our responsibility to check if the number
		// of lines that have been run since this quasi-thread started now indicate that
		// interruptibility should be reenabled.  But if UninterruptedLineCountMax is negative, don't
		// bother checking because this quasi-thread will stay non-interruptible until it finishes:
		if (!g.AllowThisThreadToBeInterrupted && g_script.mUninterruptedLineCountMax >= 0)
		{
			// Note that there is a timer that handles the UninterruptibleTime setting, so we don't
			// have handle that setting here.  But that timer is killed by the DISABLE_UNINTERRUPTIBLE
			// macro we call below.  This is because we don't want the timer to "fire" after we've
			// already met the conditions which allow interruptibility to be restored, because if
			// it did, it might interfere with the fact that some other code might already be using
			// g.AllowThisThreadToBeInterrupted again for its own purpose:
			if (g.UninterruptedLineCount > g_script.mUninterruptedLineCountMax)
				MAKE_THREAD_INTERRUPTIBLE
			else
				// Incrementing this unconditionally makes it a cruder measure than g.LinesPerCycle,
				// but it seems okay to be less accurate for this purpose:
				++g.UninterruptedLineCount;
		}

		// The below handles the message-loop checking regardless of whether
		// aMode is ONLY_ONE_LINE (i.e. recursed) or not (i.e. we're using
		// the for-loop to execute the script linearly):
		if ((g.LinesPerCycle >= 0 && g_script.mLinesExecutedThisCycle >= g.LinesPerCycle)
			|| (g.IntervalBeforeRest >= 0 && tick_now - g_script.mLastScriptRest >= (DWORD)g.IntervalBeforeRest))
			// Sleep in between batches of lines, like AutoIt, to reduce the chance that
			// a maxed CPU will interfere with time-critical apps such as games,
			// video capture, or video playback.  Note: MsgSleep() will reset
			// mLinesExecutedThisCycle for us:
			MsgSleep(10);  // Don't use INTERVAL_UNSPECIFIED, which wouldn't sleep at all if there's a msg waiting.

		// At this point, a pause may have been triggered either by the above MsgSleep()
		// or due to the action of a command (e.g. Pause, or perhaps tray menu "pause" was selected during Sleep):
		for (;;)
		{
			if (g.IsPaused)
				MsgSleep(INTERVAL_UNSPECIFIED);  // Must check often to periodically run timed subroutines.
			else
				break;
		}

		// Do these only after the above has had its opportunity to spend a significant amount
		// of time doing what it needed to do.  i.e. do these immediately before the line will actually
		// be run so that the time it takes to run will be reflected in the ListLines log.
        g_script.mCurrLine = line;  // Simplifies error reporting when we get deep into function calls.

		// Maintain a circular queue of the lines most recently executed:
		sLog[sLogNext] = line; // The code actually runs faster this way than if this were combined with the above.
		// Get a fresh tick in case tick_now is out of date.  Strangely, it takes benchmarks 3% faster
		// on my system with this line than without it, but that's probably just a quirk of the build
		// or the CPU's caching.  It was already shown previously that the released version of 1.0.09
		// was almost 2% faster than an early version of this version (yet even now, that prior version
		// benchmarks slower than this one, which I can't explain).
		sLogTick[sLogNext++] = GetTickCount();  // Incrementing here vs. separately benches a little faster.
		if (sLogNext >= LINE_LOG_SIZE)
			sLogNext = 0;

		// Do this only after the opportunity to Sleep (above) has passed, but before
		// calling ExpandArgs() below, because any of the MsgSleep's above or any of
		// those called the command of a prior loop iteration (e.g. WinWait or Sleep)
		// may have changed the value of the global.  We now want the global to reflect
		// the memory address that our caller gave to us (by definition, we must have
		// been called recursively if we're in a file-loop, so aCurrentFile will be
		// non-NULL if our caller, or its caller, was in a file-loop).  If aCurrentFile
		// is NULL, set the variable to be our own layer's current_file, so that the
		// value of a variable such as %A_LoopFileName% can still be fetched even if
		// we're currently outside a file-loop (in which case, the most recently found
		// file by any loop invoked by our layer will be used?  UPDATE: It seems better,
		// both to simplify the code and enforce the fact that loop variables should not
		// be valid outside the context of a loop, to use NULL for the current file
		// whenever aCurrentFile is NULL.  Note: The memory for this variable resides
		// in the stack of an instance of PerformLoop(), which is our caller or our
		// caller's caller, etc.  In other words, it shouldn't be possible for
		// aCurrentFile to be non-NULL if there isn't a PerformLoop() beneath us
		// in the stack:
		g_script.mLoopFile = aCurrentFile;
		g_script.mLoopRegItem = aCurrentRegItem; // Similar in function to the above.
		g_script.mLoopReadFile = aCurrentReadFile; // Similar in function to the above.
		g_script.mLoopField = aCurrentField; // Similar in function to the above.
		g_script.mLoopIteration = aCurrentLoopIteration; // Similar in function to the above.

		// Do this only after the opportunity to Sleep (above) has passed, because during
		// that sleep, a new subroutine might be launched which would likely overwrite the
		// deref buffer used for arg expansion, below:
		// Expand any dereferences contained in this line's args.
		// Note: Only one line at a time be expanded via the above function.  So be sure
		// to store any parts of a line that are needed prior to moving on to the next
		// line (e.g. control stmts such as IF and LOOP).  Also, don't expand
		// ACT_ASSIGN because a more efficient way of dereferencing may be possible
		// in that case:
		if (line->mActionType != ACT_ASSIGN && line->ExpandArgs() != OK)
			return FAIL;  // Abort the current subroutine, but don't terminate the app.

		if (ACT_IS_IF(line->mActionType))
		{
			++g_script.mLinesExecutedThisCycle;  // If and its else count as one line for this purpose.
			if_condition = line->EvaluateCondition();
			if (if_condition == FAIL)
				return FAIL;
			if (if_condition == CONDITION_TRUE)
			{
				// line->mNextLine has already been verified non-NULL by the pre-parser, so
				// this dereference is safe:
				result = line->mNextLine->ExecUntil(ONLY_ONE_LINE, &jump_to_line, aCurrentFile
					, aCurrentRegItem, aCurrentReadFile, aCurrentField, aCurrentLoopIteration);
				if (jump_to_line == line)
					// Since this IF's ExecUntil() encountered a Goto whose target is the IF
					// itself, continue with the for-loop without moving to a different
					// line.  Also: stay in this recursion layer even if aMode == ONLY_ONE_LINE
					// because we don't want the caller handling it because then it's cleanup
					// to jump to its end-point (beyond its own and any unowned elses) won't work.
					// Example:
					// if x  <-- If this layer were to do it, its own else would be unexpectedly encountered.
					//    label1:
					//    if y  <-- We want this statement's layer to handle the goto.
					//       goto, label1
					//    else
					//       ...
					// else
					//   ...
					continue;
				if (aMode == ONLY_ONE_LINE && jump_to_line != NULL && apJumpToLine != NULL)
					// The above call to ExecUntil() told us to jump somewhere.  But since we're in
					// ONLY_ONE_LINE mode, our caller must handle it because only it knows how
					// to extricate itself from whatever it's doing:
					*apJumpToLine = jump_to_line; // Tell the caller to handle this jump.
				if (result == FAIL || result == EARLY_RETURN || result == EARLY_EXIT
					|| result == LOOP_BREAK || result == LOOP_CONTINUE
					|| aMode == ONLY_ONE_LINE)
					// EARLY_RETURN can occur if this if's action was a block, and that block
					// contained a RETURN, or if this if's only action is RETURN.  It can't
					// occur if we just executed a Gosub, because that Gosub would have been
					// done from a deeper recursion layer (and executing a Gosub in
					// ONLY_ONE_LINE mode can never return EARLY_RETURN).
					return result;
				// Now this if-statement, including any nested if's and their else's,
				// has been fully evaluated by the recusion above.  We must jump to
				// the end of this if-statement to get to the right place for
				// execution to resume.  UPDATE: Or jump to the goto target if the
				// call to ExecUntil told us to do that instead:
				if (jump_to_line != NULL && jump_to_line->mParentLine != line->mParentLine)
				{
					if (apJumpToLine != NULL) // In this case, it should always be non-NULL?
						*apJumpToLine = jump_to_line; // Tell the caller to handle this jump.
					return OK;
				}
				if (jump_to_line != NULL) // jump to where the caller told us to go, rather than the end of IF.
					line = jump_to_line;
				else // Do the normal clean-up for an IF statement:
				{
					// Set line to be either the IF's else or the end of the if-stmt:
					if (   !(line = line->mRelatedLine)   )
						// The preparser has ensured that the only time this can happen is when
						// the end of the script has been reached (i.e. this if-statement
						// has no else and it's the last statement in the script):
						return OK;
					if (line->mActionType == ACT_ELSE)
						line = line->mRelatedLine;
						// Now line is the ELSE's "I'm finished" jump-point, which is where
						// we want to be.  If line is now NULL, it will be caught when this
						// loop iteration is ended by the "continue" stmt below.  UPDATE:
						// it can't be NULL since all scripts now end in ACT_EXIT.
					// else the IF had NO else, so we're already at the IF's "I'm finished" jump-point.
				}
			}
			else // if_condition == CONDITION_FALSE
			{
				if (   !(line = line->mRelatedLine)   )  // Set to IF's related line.
					// The preparser has ensured that this can only happen if the end of the script
					// has been reached.  UPDATE: Probably can't happen anymore since all scripts
					// are now provided with a terminating ACT_EXIT:
					return OK;
				if (line->mActionType != ACT_ELSE && aMode == ONLY_ONE_LINE)
					// Since this IF statement has no ELSE, and since it was executed
					// in ONLY_ONE_LINE mode, the IF-ELSE statement, which counts as
					// one line for the purpose of ONLY_ONE_LINE mode, has finished:
					return OK;
				if (line->mActionType == ACT_ELSE) // This IF has an else.
				{
					// Preparser has ensured that every ELSE has a non-NULL next line:
					result = line->mNextLine->ExecUntil(ONLY_ONE_LINE, &jump_to_line, aCurrentFile
						, aCurrentRegItem, aCurrentReadFile, aCurrentField, aCurrentLoopIteration);
					if (aMode == ONLY_ONE_LINE && jump_to_line != NULL && apJumpToLine != NULL)
						// The above call to ExecUntil() told us to jump somewhere.  But since we're in
						// ONLY_ONE_LINE mode, our caller must handle it because only it knows how
						// to extricate itself from whatever it's doing:
						*apJumpToLine = jump_to_line; // Tell the caller to handle this jump.
					if (result == FAIL || result == EARLY_RETURN || result == EARLY_EXIT
						|| result == LOOP_BREAK || result == LOOP_CONTINUE
						|| aMode == ONLY_ONE_LINE)
						return result;
					if (jump_to_line != NULL && jump_to_line->mParentLine != line->mParentLine)
					{
						if (apJumpToLine != NULL) // In this case, it should always be non-NULL?
							*apJumpToLine = jump_to_line; // Tell the caller to handle this jump.
						return OK;
					}
					if (jump_to_line != NULL)
						// jump to where the called function told us to go, rather than the end of our ELSE.
						line = jump_to_line;
					else // Do the normal clean-up for an ELSE statement.
						line = line->mRelatedLine;
						// Now line is the ELSE's "I'm finished" jump-point, which is where
						// we want to be.  If line is now NULL, it will be caught when this
						// loop iteration is ended by the "continue" stmt below.  UPDATE:
						// it can't be NULL since all scripts now end in ACT_EXIT.
					// else the IF had NO else, so we're already at the IF's "I'm finished" jump-point.
				}
				// else the IF had NO else, so we're already at the IF's "I'm finished" jump-point.
			}
			continue; // Let the for-loop process the new location specified by <line>.
		}

		// If above didn't continue, it's not an IF, so handle the other
		// flow-control types:
		switch (line->mActionType)
		{
		case ACT_GOTO:
			// A single goto can cause an infinite loop if misused, so be sure to do this to
			// prevent the program from hanging:
			++g_script.mLinesExecutedThisCycle;
			if (   !(jump_target = line->mRelatedLine)   )
				// The label is a dereference, otherwise it would have been resolved at load-time.
				// So send true because we don't want to update its mRelatedLine.  This is because
				// we want to resolve the label every time through the loop in case the variable
				// that contains the label changes, e.g. Gosub, %MyLabel%
				if (   !(jump_target = line->GetJumpTarget(true))   )
					return FAIL; // Error was already displayed by called function.
			// One or both of these lines can be NULL.  But the preparser should have
			// ensured that all we need to do is a simple compare to determine
			// whether this Goto should be handled by this layer or its caller
			// (i.e. if this Goto's target is not in our nesting level, it MUST be the
			// caller's responsibility to either handle it or pass it on to its
			// caller).
			if (aMode == ONLY_ONE_LINE || line->mParentLine != jump_target->mParentLine)
			{
				if (apJumpToLine != NULL)
					*apJumpToLine = jump_target; // Tell the caller to handle this jump.
				return OK;
			}
			// Otherwise, we will handle this Goto since it's in our nesting layer:
			line = jump_target;
			break;  // Resume looping starting at the above line.
		case ACT_GOSUB:
			// A single gosub can cause an infinite loop if misused (i.e. recusive gosubs),
			// so be sure to do this to prevent the program from hanging:
			++g_script.mLinesExecutedThisCycle;
			if (   !(jump_target = line->mRelatedLine)   )
				// The label is a dereference, otherwise it would have been resolved at load-time.
				// So send true because we don't want to update its mRelatedLine.  This is because
				// we want to resolve the label every time through the loop in case the variable
				// that contains the label changes, e.g. Gosub, %MyLabel%
				if (   !(jump_target = line->GetJumpTarget(true))   )
					return FAIL; // Error was already displayed by called function.
			// I'm pretty sure it's not valid for this call to ExecUntil() to tell us to jump
			// somewhere, because the called function, or a layer even deeper, should handle
			// the goto prior to returning to us?  So the last param is omitted:
			result = jump_target->ExecUntil(UNTIL_RETURN, NULL, aCurrentFile, aCurrentRegItem
				, aCurrentReadFile, aCurrentField, aCurrentLoopIteration);
			// Must do these return conditions in this specific order:
			if (result == FAIL || result == EARLY_EXIT)
				return result;
			if (aMode == ONLY_ONE_LINE)
				// This Gosub doesn't want its caller to know that the gosub's
				// subroutine returned early:
				return (result == EARLY_RETURN) ? OK : result;
			// If the above didn't return, the subroutine finished successfully and
			// we should now continue on with the line after the Gosub:
			line = line->mNextLine;
			break;  // Resume looping starting at the above line.
		case ACT_GROUPACTIVATE:
		{
			// This section is here rather than in Perform() because GroupActivate can
			// sometimes execute a Gosub.
			++g_script.mLinesExecutedThisCycle; // Always increment for GroupActivate.
			WinGroup *group;
			if (   !(group = (WinGroup *)mAttribute)   )
				if (   !(group = g_script.FindOrAddGroup(LINE_ARG1))   )
					return FAIL;  // It already displayed the error for us.
			Line *jump_to_line;
			// Note: This will take care of DoWinDelay if needed:
			group->Activate(*LINE_ARG2 && !stricmp(LINE_ARG2, "R"), NULL, (void **)&jump_to_line);
			if (jump_to_line)
			{
				if (!line->IsJumpValid(jump_to_line))
					// This check probably isn't necessary since IsJumpValid() is mostly
					// for Goto's.  But just in case the gosub's target label is some
					// crazy place:
					return FAIL;
				// This section is just like the Gosub code above, so maintain them together.
				result = jump_to_line->ExecUntil(UNTIL_RETURN, NULL, aCurrentFile, aCurrentRegItem
					, aCurrentReadFile, aCurrentField, aCurrentLoopIteration);
				if (result == FAIL || result == EARLY_EXIT)
					return result;
				if (aMode == ONLY_ONE_LINE)
					return (result == EARLY_RETURN) ? OK : result;
			}
			line = line->mNextLine;
			break;
		}
		case ACT_RETURN:
			// Don't count returns against the total since they should be nearly instantaneous:
			//++g_script.mLinesExecutedThisCycle;
			// Note about RETURNs:
			// Although a return is really just a kind of block-end, keep it separate
			// because when a return is encountered inside a block, it has a double function:
			// to first break out of all enclosing blocks and then return from the gosub.
			if (aMode != UNTIL_RETURN)
				// Tells the caller to return early if it's not the Gosub that directly
				// brought us into this subroutine.  i.e. it allows us to escape from
				// any number of nested blocks in order to get back out of their
				// recursive layers and back to the place this RETURN has meaning
				// to someone (at the right recursion layer):
				return EARLY_RETURN;
			// For now, I don't think there's any way to generate a "Return with
			// no matching Gosub" error, because RETURN effectively exits the
			// script, going back into the message event loop in WinMain()?
			return OK;

		case ACT_REPEAT:
		case ACT_LOOP:
		{
			AttributeType attr = line->mAttribute;
			HKEY root_key_type = NULL; // This will hold the type of root key, independent of whether it is local or remote.
			if (attr == ATTR_LOOP_REG)
				root_key_type = RegConvertRootKey(LINE_ARG1);
			else if (attr == ATTR_LOOP_UNKNOWN || attr == ATTR_NONE)
			{
				// Since it couldn't be determined at load-time (probably due to derefs),
				// determine whether it's a file-loop, registry-loop or a normal/counter loop.
				// But don't change the value of line->mAttribute because that's our
				// indicator of whether this needs to be evaluated every time for
				// this particular loop (since the nature of the loop can change if the
				// contents of the variables dereferenced for this line change during runtime):
				switch (line->mArgc)
				{
				case 0:
					attr = ATTR_LOOP_NORMAL;
					break;
				case 1:
					// Unlike at loadtime, allow it to be negative at runtime in case it was a variable
					// reference that resolved to a negative number, to indicate that 0 iterations
					// should be performed.  UPDATE: Also allow floating point numbers at runtime
					// but not at load-time (since it doesn't make sense to have a literal floating
					// point number as the iteration count, but a variable containing a pure float
					// should be allowed):
					if (IsPureNumeric(LINE_ARG1, true, true, true))
						attr = ATTR_LOOP_NORMAL;
					else
					{
						root_key_type = RegConvertRootKey(LINE_ARG1);
						attr = root_key_type ? ATTR_LOOP_REG : ATTR_LOOP_FILE;
					}
					break;
				default: // 2 or more args.
					if (!stricmp(LINE_ARG1, "Read"))
						attr = ATTR_LOOP_READ_FILE;
					// Note that a "Parse" loop is not allowed to have it's first param be a variable reference
					// that resolves to the word "Parse" at runtime.  This is because the input variable would not
					// have been resolved in this case (since the type of loop was unknown at load-time),
					// and it would be complicated to have to add code for that, especially since there's
					// virtually no conceivable use for allowing it be a variable reference.
					else
					{
						root_key_type = RegConvertRootKey(LINE_ARG1);
						attr = root_key_type ? ATTR_LOOP_REG : ATTR_LOOP_FILE;
					}
				}
			}

			bool recurse_subfolders = (attr == ATTR_LOOP_FILE && *LINE_ARG3 == '1' && !*(LINE_ARG3 + 1))
				|| (attr == ATTR_LOOP_REG && *LINE_ARG4 == '1' && !*(LINE_ARG4 + 1));

			__int64 iteration_limit = 0;
			bool is_infinite = line->mArgc < 1;
			if (!is_infinite)
				// Must be set to zero at least for ATTR_LOOP_FILE:
				iteration_limit = (attr == ATTR_LOOP_FILE || attr == ATTR_LOOP_REG || attr == ATTR_LOOP_READ_FILE
					||  attr == ATTR_LOOP_PARSE) ? 0 : ATOI64(LINE_ARG1);

			if (line->mActionType == ACT_REPEAT && !iteration_limit)
				is_infinite = true;  // Because a 0 means infinite in AutoIt2 for the REPEAT command.
			// else if it's negative, zero iterations will be performed automatically.

			FileLoopModeType file_loop_mode;
			if (attr == ATTR_LOOP_FILE)
			{
				file_loop_mode = (line->mArgc <= 1) ? FILE_LOOP_FILES_ONLY : ConvertLoopMode(LINE_ARG2);
				if (file_loop_mode == FILE_LOOP_INVALID)
					return line->LineError(ERR_LOOP_FILE_MODE ERR_ABORT, FAIL, LINE_ARG2);
			}
			else if (attr == ATTR_LOOP_REG)
			{
				file_loop_mode = (line->mArgc <= 2) ? FILE_LOOP_FILES_ONLY : ConvertLoopMode(LINE_ARG3);
				if (file_loop_mode == FILE_LOOP_INVALID)
					return line->LineError(ERR_LOOP_REG_MODE ERR_ABORT, FAIL, LINE_ARG3);
			}
			else
				file_loop_mode = FILE_LOOP_INVALID;

			bool continue_main_loop = false; // Init prior to below call.
			jump_to_line = NULL; // Init prior to below call.

			// This must be a variable in this function's stack because registry loops and file-pattern
			// loops can be intrinsically recursive, in which case not only does it need to pass on the
			// correct/current iteration to the instance it calls, it must also receive back the updated
			// count when it is returned to.  This is also related to the loop-recursion bugfix documented
			// for v1.0.20: fixes A_Index so that it doesn't wrongly reset to 0 inside recursive file-loops
			// and registry loops:
			__int64 script_iteration = 0;

			if (attr == ATTR_LOOP_PARSE)
			{
				// The phrase "csv" is unique enough since user can always rearrange the letters
				// to do a literal parse using C, S, and V as delimiters:
				if (stricmp(LINE_ARG3, "CSV"))
					result = line->PerformLoopParse(aCurrentFile, aCurrentRegItem, aCurrentReadFile
						, continue_main_loop, jump_to_line, script_iteration);
				else
					result = line->PerformLoopParseCSV(aCurrentFile, aCurrentRegItem, aCurrentReadFile
						, continue_main_loop, jump_to_line, script_iteration);
			}
			else if (attr == ATTR_LOOP_READ_FILE)
			{
				// Open the input file:
				FILE *read_file = fopen(LINE_ARG2, "r");
				if (read_file)
				{
					result = line->PerformLoopReadFile(aCurrentFile, aCurrentRegItem, aCurrentField
						, continue_main_loop, jump_to_line, read_file, LINE_ARG3, script_iteration);
					fclose(read_file);
				}
				else
					// The open of a the input file failed.  So just set result to OK since no ErrorLevel
					// setting is supported with loops (since that seems like it would be an overuse
					// of ErrorLevel, perhaps changing its value too often when the user would want
					// it saved -- in any case, changing that now might break existing scripts).
					result = OK;
			}
			else if (attr == ATTR_LOOP_REG)
			{
				// This isn't the most efficient way to do things (e.g. the repeated calls to
				// RegConvertRootKey()), but it the simplest way for now.  Optimization can
				// be done at a later time:
				bool is_remote_registry;
				// This will open the key if it's remote:
				HKEY root_key = RegConvertRootKey(LINE_ARG1, &is_remote_registry);
				if (root_key)
				{
					// root_key_type needs to be passed in order to support GetLoopRegKey():
					result = line->PerformLoopReg(aCurrentFile, aCurrentReadFile, aCurrentField
						, continue_main_loop, jump_to_line, file_loop_mode, recurse_subfolders
						, root_key_type, root_key, LINE_ARG2, script_iteration);
					if (is_remote_registry)
						RegCloseKey(root_key);
				}
				else
					// The open of a remote key failed (we know it's remote otherwise it should have
					// failed earlier rather than here).  So just set result to OK since no ErrorLevel
					// setting is supported with loops (since that seems like it would be an overuse
					// of ErrorLevel, perhaps changing its value too often when the user would want
					// it saved.  But in any case, changing that now might break existing scripts).
					result = OK;
			}
			else // All other loops types are handled this way:
				result = line->PerformLoop(aCurrentFile, aCurrentRegItem, aCurrentReadFile
					, aCurrentField, continue_main_loop, jump_to_line, attr, file_loop_mode, recurse_subfolders
					, LINE_ARG1, iteration_limit, is_infinite, script_iteration);

			if (result == FAIL || result == EARLY_RETURN || result == EARLY_EXIT)
				return result;
			// else result can be LOOP_BREAK or OK, but not LOOP_CONTINUE.
			if (continue_main_loop) // It signaled us to do this:
				continue;

			if (aMode == ONLY_ONE_LINE)
			{
				if (jump_to_line != NULL && apJumpToLine != NULL)
					// The above call to ExecUntil() told us to jump somewhere.  But since we're in
					// ONLY_ONE_LINE mode, our caller must handle it because only it knows how
					// to extricate itself from whatever it's doing:
					*apJumpToLine = jump_to_line; // Tell the caller to handle this jump.
				// Return OK even if our result was LOOP_CONTINUE because we already handled the continue:
				return OK;
			}
			if (jump_to_line)
			{
				if (jump_to_line->mParentLine != line->mParentLine)
				{
					// Our caller must handle the jump if it doesn't share the same parent as the
					// current line (i.e. it's not at the same nesting level) because that means
					// the jump target is at a more shallow nesting level than where we are now:
					if (apJumpToLine != NULL) // i.e. caller gave us a place to store the jump target.
						*apJumpToLine = jump_to_line; // Tell the caller to handle this jump.
					return OK;
				}
				// Since above didn't return, we're supposed to handle this jump.  So jump and then
				// continue execution from there:
				line = jump_to_line;
				break; // end this case of the switch().
			}
			// Since the above didn't return or break, either the loop has completed the specified
			// number of iterations or it was broken via the break command.  In either case, we jump
			// to the line after our loop's structure and continue there:
			line = line->mRelatedLine;
			break;
		} // case ACT_LOOP.
		case ACT_BREAK:
			return LOOP_BREAK;
		case ACT_CONTINUE:
			return LOOP_CONTINUE;
		case ACT_EXIT:
			// If this script has no hotkeys and hasn't activated one of the hooks, EXIT will cause the
			// the program itself to terminate.  Otherwise, it causes us to return from all blocks
			// and Gosubs (i.e. all the way out of the current subroutine, which was usually triggered
			// by a hotkey):
			if (IS_PERSISTENT)
				return EARLY_EXIT;  // It's "early" because only the very end of the script is the "normal" exit.
			else
				// This has been tested and it does yield to the OS the error code indicated in ARG1,
				// if present (otherwise it returns 0, naturally) as expected:
				return g_script.ExitApp(EXIT_EXIT, NULL, ATOI(LINE_ARG1));
		case ACT_EXITAPP: // Unconditional exit.
			return g_script.ExitApp(EXIT_EXIT, NULL, ATOI(LINE_ARG1));
		case ACT_BLOCK_BEGIN:
			// Don't count block-begin/end against the total since they should be nearly instantaneous:
			//++g_script.mLinesExecutedThisCycle;
			// In this case, line->mNextLine is already verified non-NULL by the pre-parser:
			result = line->mNextLine->ExecUntil(UNTIL_BLOCK_END, &jump_to_line, aCurrentFile, aCurrentRegItem
				, aCurrentReadFile, aCurrentField, aCurrentLoopIteration);
			if (jump_to_line == line)
				// Since this Block-begin's ExecUntil() encountered a Goto whose target is the
				// block-begin itself, continue with the for-loop without moving to a different
				// line.  Also: stay in this recursion layer even if aMode == ONLY_ONE_LINE
				// because we don't want the caller handling it because then it's cleanup
				// to jump to its end-point (beyond its own and any unowned elses) won't work.
				// Example:
				// if x  <-- If this layer were to do it, its own else would be unexpectedly encountered.
				// label1:
				// { <-- We want this statement's layer to handle the goto.
				//    if y
				//       goto, label1
				//    else
				//       ...
				// }
				// else
				//   ...
				continue;
			if (aMode == ONLY_ONE_LINE && jump_to_line != NULL && apJumpToLine != NULL)
				// The above call to ExecUntil() told us to jump somewhere.  But since we're in
				// ONLY_ONE_LINE mode, our caller must handle it because only it knows how
				// to extricate itself from whatever it's doing:
				*apJumpToLine = jump_to_line; // Tell the caller to handle this jump.
			if (result == FAIL || result == EARLY_RETURN || result == EARLY_EXIT
				|| result == LOOP_BREAK || result == LOOP_CONTINUE
				|| aMode == ONLY_ONE_LINE)
				return result;
			// Currently, all blocks are normally executed in ONLY_ONE_LINE mode because
			// they are the direct actions of an IF, an ELSE, or a LOOP.  So the
			// above will already have returned except when the user has created a
			// generic, standalone block with no assciated control statement.
			// Check to see if we need to jump somewhere:
			if (jump_to_line != NULL && line->mParentLine != jump_to_line->mParentLine)
			{
				if (apJumpToLine != NULL) // In this case, it should always be non-NULL?
					*apJumpToLine = jump_to_line; // Tell the caller to handle this jump.
				return OK;
			}
			if (jump_to_line != NULL) // jump to where the caller told us to go, rather than the end of our block.
				line = jump_to_line;
			else // Just go to the end of our block and continue from there.
				line = line->mRelatedLine;
				// Now line is the line after the end of this block.  Can be NULL (end of script).
				// UPDATE: It can't be NULL (not that it matters in this case) since the loader
				// has ensured that all scripts now end in an ACT_EXIT.
			break;  // Resume looping starting at the above line.
		case ACT_BLOCK_END:
			// Don't count block-begin/end against the total since they should be nearly instantaneous:
			//++g_script.mLinesExecutedThisCycle;
			if (aMode != UNTIL_BLOCK_END)
				// Shouldn't happen if the pre-parser and this function are designed properly?
				// Update: Rajat found a way for this to happen that basically amounts to this:
				// If within a loop you gosub a label that is also inside of the block, and
				// that label sometimes doesn't return (i.e. due to a missing "return" somewhere
				// in its flow of control), the loop(s)'s block-end symbols will be encountered
				// by the subroutine, and these symbols don't have meaning to it.  In other words,
				// the subroutine has put us into a waiting-for-return state rather than a
				// waiting-for-block-end state, so when block-end's are encountered, that is
				// considered a runtime error:
				return line->LineError("Unexpected end-of-block (probably a Gosub without Return)." ERR_ABORT);
			return OK; // It's the caller's responsibility to resume execution at the next line, if appropriate.
		case ACT_ELSE:
			// Shouldn't happen if the pre-parser and this function are designed properly?
			return line->LineError("Unexpected ELSE." ERR_ABORT);
		default:
			++g_script.mLinesExecutedThisCycle;
			result = line->Perform(aCurrentFile, aCurrentRegItem, aCurrentReadFile);
			if (result == FAIL || aMode == ONLY_ONE_LINE)
				// Thus, Perform() should be designed to only return FAIL if it's an
				// error that would make it unsafe to proceed the subroutine
				// we're executing now:
				return result;
			line = line->mNextLine;
		}
	}

	// Above loop ended because the end of the script was reached.
	// At this point, it should be impossible for aMode to be
	// UNTIL_BLOCK_END because that would mean that the blocks
	// aren't all balanced (or there is a design flaw in this
	// function), but they are balanced because the preparser
	// verified that.  It should also be impossible for the
	// aMode to be ONLY_ONE_LINE because the function is only
	// called in that mode to execute the first action-line
	// beneath an IF or an ELSE, and the preparser has already
	// verified that every such IF and ELSE has a non-NULL
	// line after it.  Finally, aMode can be UNTIL_RETURN, but
	// that is normal mode of operation at the top level,
	// so probably shouldn't be considered an error.  For example,
	// if the script has no hotkeys, it is executed from its
	// first line all the way to the end.  For it not to have
	// a RETURN or EXIT is not an error.  UPDATE: The loader
	// now ensures that all scripts end in ACT_EXIT, so
	// this line should never be reached:
	return OK;
}



inline ResultType Line::EvaluateCondition()
// Returns FAIL, CONDITION_TRUE, or CONDITION_FALSE.
{
#ifdef _DEBUG
	if (!ACT_IS_IF(mActionType))
		return LineError("EvaluateCondition() was called with a line that isn't a condition."
			ERR_ABORT);
#endif

	SymbolType var_is_pure_numeric, value_is_pure_numeric, value2_is_pure_numeric;
	int if_condition;
	char *cp;

	switch (mActionType)
	{
	case ACT_IFEXPR:
		// Use ATOF to support hex, float, and integer formats.  Also, explicitly compare to 0.0
		// to avoid truncation of double, which would result in a value such as 0.1 being seen
		// as false rather than true.  Fixed in v1.0.25.12 so that only the following are false:
		// 0
		// 0.0
		// 0x0
		// (variants of the above)
		// blank string
		// ... in other words, "if var" should be true if it contains a non-numeric string.
		cp = ARG1;  // It should help performance to resolve the ARG1 macro only once.
		if (!*cp)
			if_condition = false;
		else if (!IsPureNumeric(cp, true, false, true)) // i.e. a var containing all whitespace would be considered "true", since it's a non-blank string that isn't equal to 0.0.
			if_condition = true;
		else // It's purely numeric, not blank, and not all whitespace.
			if_condition = (ATOF(cp) != 0.0);
		break;

	// For ACT_IFWINEXIST and ACT_IFWINNOTEXIST, although we validate that at least one
	// of their window params is non-blank during load, it's okay at runtime for them
	// all to resolve to be blank (due to derefs), without an error being reported.
	// It's probably more flexible that way, and in any event WinExist() is equipped to
	// handle all-blank params:
	case ACT_IFWINEXIST:
		// NULL-check this way avoids compiler warnings:
		if_condition = (WinExist(FOUR_ARGS, false, true) != NULL);
		break;
	case ACT_IFWINNOTEXIST:
		if_condition = !WinExist(FOUR_ARGS, false, true); // Seems best to update last-used even here.
		break;
	case ACT_IFWINACTIVE:
		if_condition = (WinActive(FOUR_ARGS, true) != NULL);
		break;
	case ACT_IFWINNOTACTIVE:
		if_condition = !WinActive(FOUR_ARGS, true);
		break;

	case ACT_IFEXIST:
		if_condition = DoesFilePatternExist(ARG1);
		break;
	case ACT_IFNOTEXIST:
		if_condition = !DoesFilePatternExist(ARG1);
		break;

	case ACT_IFINSTRING:
		#define STRING_SEARCH (g.StringCaseSense ? strstr(ARG1, ARG2) : strcasestr(ARG1, ARG2))
		if_condition = STRING_SEARCH != NULL;
		break;
	case ACT_IFNOTINSTRING:
		if_condition = STRING_SEARCH == NULL;
		break;

	case ACT_IFEQUAL:
		// For now, these seem to be the best rules to follow:
		// 1) If either one is non-empty and non-numeric, they're compared as strings.
		// 2) Otherwise, they're compared as numbers (with empty vars treated as zero).
		// In light of the above, two empty values compared to each other is the same as
		// "0 compared to 0".  e.g. if the clipboard is blank, the line "if clipboard ="
		// would be true.  However, the following are side-effects (are there any more?):
		// if var1 =    ; statement is true if var1 contains a literal zero (possibly harmful)
		// if var1 = 0  ; statement is true if var1 is blank (mostly harmless?)
		// if var1 !=   ; statement is false if var1 contains a literal zero (possibly harmful)
		// if var1 != 0 ; statement is false if var1 is blank (mostly harmless?)
		// In light of the above, the BOTH_ARE_NUMERIC macro has been altered to return
		// false if one of the items is a literal zero and the other is blank, so that
		// the two items will be compared as strings.  UPDATE: Altered it again because it
		// seems best to consider blanks to always be non-numeric (i.e. if either var is blank,
		// they will be compared as strings rather than as numbers):
		#undef STRING_COMPARE
		#define STRING_COMPARE (g.StringCaseSense ? strcmp(ARG1, ARG2) : stricmp(ARG1, ARG2))
		#undef DETERMINE_NUMERIC_TYPES
		#define DETERMINE_NUMERIC_TYPES \
			value_is_pure_numeric = IsPureNumeric(ARG2, true, false, true);\
			var_is_pure_numeric = IsPureNumeric(ARG1, true, false, true);
		#define DETERMINE_NUMERIC_TYPES2 \
			DETERMINE_NUMERIC_TYPES \
			value2_is_pure_numeric = IsPureNumeric(ARG3, true, false, true);
		#define IF_EITHER_IS_NON_NUMERIC if (!value_is_pure_numeric || !var_is_pure_numeric)
		#define IF_EITHER_IS_NON_NUMERIC2 if (!value_is_pure_numeric || !value2_is_pure_numeric || !var_is_pure_numeric)
		#undef IF_EITHER_IS_FLOAT
		#define IF_EITHER_IS_FLOAT if (value_is_pure_numeric == PURE_FLOAT || var_is_pure_numeric == PURE_FLOAT)

		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_NON_NUMERIC
			if_condition = !STRING_COMPARE;
		else IF_EITHER_IS_FLOAT  // It might perform better to only do float conversions & math when necessary.
			if_condition = ATOF(ARG1) == ATOF(ARG2);
		else
			if_condition = ATOI64(ARG1) == ATOI64(ARG2);

		break;
	case ACT_IFNOTEQUAL:
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_NON_NUMERIC
			if_condition = STRING_COMPARE;
		else IF_EITHER_IS_FLOAT  // It might perform better to only do float conversions & math when necessary.
			if_condition = ATOF(ARG1) != ATOF(ARG2);
		else
			if_condition = ATOI64(ARG1) != ATOI64(ARG2);
		break;
	case ACT_IFLESS:
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_NON_NUMERIC
			if_condition = STRING_COMPARE < 0;
		else IF_EITHER_IS_FLOAT  // It might perform better to only do float conversions & math when necessary.
			if_condition = ATOF(ARG1) < ATOF(ARG2);
		else
			if_condition = ATOI64(ARG1) < ATOI64(ARG2);
		break;
	case ACT_IFLESSOREQUAL:
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_NON_NUMERIC
			if_condition = STRING_COMPARE <= 0;
		else IF_EITHER_IS_FLOAT  // It might perform better to only do float conversions & math when necessary.
			if_condition = ATOF(ARG1) <= ATOF(ARG2);
		else
			if_condition = ATOI64(ARG1) <= ATOI64(ARG2);
		break;
	case ACT_IFGREATER:
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_NON_NUMERIC
			if_condition = STRING_COMPARE > 0;
		else IF_EITHER_IS_FLOAT  // It might perform better to only do float conversions & math when necessary.
			if_condition = ATOF(ARG1) > ATOF(ARG2);
		else
			if_condition = ATOI64(ARG1) > ATOI64(ARG2);
		break;
	case ACT_IFGREATEROREQUAL:
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_NON_NUMERIC
			if_condition = STRING_COMPARE >= 0;
		else IF_EITHER_IS_FLOAT  // It might perform better to only do float conversions & math when necessary.
			if_condition = ATOF(ARG1) >= ATOF(ARG2);
		else
			if_condition = ATOI64(ARG1) >= ATOI64(ARG2);
		break;

	case ACT_IFBETWEEN:
	case ACT_IFNOTBETWEEN:
		DETERMINE_NUMERIC_TYPES2
		IF_EITHER_IS_NON_NUMERIC2
		{
			if (g.StringCaseSense)
				if_condition = !(strcmp(ARG1, ARG2) < 0 || strcmp(ARG1, ARG3) > 0);
			else  // case insensitive
				if_condition = !(stricmp(ARG1, ARG2) < 0 || stricmp(ARG1, ARG3) > 0);
		}
		else IF_EITHER_IS_FLOAT
		{
			double arg1_as_float = ATOF(ARG1);
			if_condition = arg1_as_float >= ATOF(ARG2) && arg1_as_float <= ATOF(ARG3);
		}
		else
		{
			__int64 arg1_as_int64 = ATOI64(ARG1);
			if_condition = arg1_as_int64 >= ATOI64(ARG2) && arg1_as_int64 <= ATOI64(ARG3);
		}
		if (mActionType == ACT_IFNOTBETWEEN)
			if_condition = !if_condition;
		break;

	case ACT_IFIN:
	case ACT_IFNOTIN:
		if_condition = IsStringInList(ARG1, ARG2, true, g.StringCaseSense);
		if (mActionType == ACT_IFNOTIN)
			if_condition = !if_condition;
		break;

	case ACT_IFCONTAINS:
	case ACT_IFNOTCONTAINS:
		if_condition = IsStringInList(ARG1, ARG2, false, g.StringCaseSense);
		if (mActionType == ACT_IFNOTCONTAINS)
			if_condition = !if_condition;
		break;

	case ACT_IFIS:
	case ACT_IFISNOT:
	{
		char *cp;
		VariableTypeType variable_type = ConvertVariableTypeName(ARG2);
		if (variable_type == VAR_TYPE_INVALID)
		{
			// Type is probably a dereferenced variable that resolves to an invalid type name.
			// It seems best to make the condition false in these cases, rather than pop up
			// a runtime error dialog:
			if_condition = false;
			break;
		}
		switch(variable_type)
		{
		case VAR_TYPE_NUMBER:
			if_condition = IsPureNumeric(ARG1, true, false, true);  // Floats are defined as being numeric.
			break;
		case VAR_TYPE_INTEGER:
			if_condition = IsPureNumeric(ARG1, true, false, false);
			break;
		case VAR_TYPE_FLOAT:
			if_condition = IsPureNumeric(ARG1, true, false, true) == PURE_FLOAT;
			break;
		case VAR_TYPE_TIME:
		{
			SYSTEMTIME st;
			// Also insist on numeric, because even though YYYYMMDDToFileTime() will properly convert a
			// non-conformant string such as "2004.4", for future compatibility, we don't want to
			// report that such strings are valid times:
			if_condition = IsPureNumeric(ARG1, false, false, false) && YYYYMMDDToSystemTime(ARG1, st, true);
			break;
		}
		case VAR_TYPE_DIGIT:
			if_condition = true;
			for (cp = ARG1; *cp; ++cp)
				if (!isdigit((UCHAR)*cp))
				{
					if_condition = false;
					break;
				}
			break;
		case VAR_TYPE_XDIGIT:
			if_condition = true;
			for (cp = ARG1; *cp; ++cp)
				if (!isxdigit((UCHAR)*cp))
				{
					if_condition = false;
					break;
				}
			break;
		case VAR_TYPE_ALNUM:
			// Like AutoIt3, the empty string is considered to be alphabetic, which is only slightly debatable.
			if_condition = true;
			for (cp = ARG1; *cp; ++cp)
				if (!IsCharAlphaNumeric(*cp)) // Use this to better support chars from non-English languages.
				{
					if_condition = false;
					break;
				}
			break;
		case VAR_TYPE_ALPHA:
			// Like AutoIt3, the empty string is considered to be alphabetic, which is only slightly debatable.
			if_condition = true;
			for (cp = ARG1; *cp; ++cp)
				if (!IsCharAlpha(*cp)) // Use this to better support chars from non-English languages.
				{
					if_condition = false;
					break;
				}
			break;
		case VAR_TYPE_UPPER:
			if_condition = true;
			for (cp = ARG1; *cp; ++cp)
				if (!IsCharUpper(*cp)) // Use this to better support chars from non-English languages.
				{
					if_condition = false;
					break;
				}
			break;
		case VAR_TYPE_LOWER:
			if_condition = true;
			for (cp = ARG1; *cp; ++cp)
				if (!IsCharLower(*cp)) // Use this to better support chars from non-English languages.
				{
					if_condition = false;
					break;
				}
			break;
		case VAR_TYPE_SPACE:
			if_condition = true;
			for (cp = ARG1; *cp; ++cp)
				if (!isspace(*cp))
				{
					if_condition = false;
					break;
				}
			break;
		}
		if (mActionType == ACT_IFISNOT)
			if_condition = !if_condition;
		break;
	}

	case ACT_IFMSGBOX:
	{
		int mb_result = ConvertMsgBoxResult(ARG1);
		if (!mb_result)
			return LineError(ERR_IFMSGBOX ERR_ABORT, FAIL, ARG1);
		if_condition = (g.MsgBoxResult == mb_result);
		break;
	}
	default: // Should never happen, but return an error if it does.
#ifdef _DEBUG
		return LineError("EvaluateCondition(): Unhandled windowing action type." ERR_ABORT);
#else
		return FAIL;
#endif
	}
	return if_condition ? CONDITION_TRUE : CONDITION_FALSE;
}



ResultType Line::PerformLoop(WIN32_FIND_DATA *apCurrentFile, RegItemStruct *apCurrentRegItem
	, LoopReadFileStruct *apCurrentReadFile, char *aCurrentField, bool &aContinueMainLoop
	, Line *&aJumpToLine, AttributeType aAttr, FileLoopModeType aFileLoopMode
	, bool aRecurseSubfolders, char *aFilePattern, __int64 aIterationLimit, bool aIsInfinite
	, __int64 &aIndex)
// Note: Even if aFilePattern is just a directory (i.e. with not wildcard pattern), it seems best
// not to append "\\*.*" to it because the pattern might be a script variable that the user wants
// to conditionally resolve to various things at runtime.  In other words, it's valid to have
// only a single directory be the target of the loop.
{
	BOOL file_found = FALSE;
	HANDLE file_search = INVALID_HANDLE_VALUE;
	char file_path[MAX_PATH] = "";
	char naked_filename_or_pattern[MAX_PATH] = "";
	// apCurrentFile is the file of the file-loop that encloses this file-loop, if any.
	// The below is our own current_file, which will take precedence over apCurrentFile
	// if this loop is a file-loop:
	WIN32_FIND_DATA new_current_file = {0};
	if (aAttr == ATTR_LOOP_FILE)
	{
		file_search = FindFirstFile(aFilePattern, &new_current_file);
		file_found = (file_search != INVALID_HANDLE_VALUE);
		// Make a local copy of the path given in aFilePattern because as the lines of
		// the loop are executed, the deref buffer (which is what aFilePattern might
		// point to if we were called from ExecUntil()) may be overwritten --
		// and we will need the path string for every loop iteration.  We also need
		// to determine naked_filename_or_pattern:
		strlcpy(file_path, aFilePattern, sizeof(file_path));
		char *last_backslash = strrchr(file_path, '\\');
		if (last_backslash)
		{
			strlcpy(naked_filename_or_pattern, last_backslash + 1, sizeof(naked_filename_or_pattern));
			*(last_backslash + 1) = '\0';  // i.e. retain the final backslash on the string.
		}
		else
		{
			strlcpy(naked_filename_or_pattern, file_path, sizeof(naked_filename_or_pattern));
			*file_path = '\0'; // There is no path, so use current working directory.
		}
		for (; file_found && FileIsFilteredOut(new_current_file, aFileLoopMode, file_path)
			; file_found = FindNextFile(file_search, &new_current_file));
	}

	// Note: It seems best NOT to report warning if the loop iterates zero times
	// (e.g if no files are found by FindFirstFile() above), since that could
	// easily be an expected outcome.

	ResultType result;
	Line *jump_to_line = NULL;
	for (; aIsInfinite || file_found || aIndex < aIterationLimit; ++aIndex)
	{
		// Execute once the body of the loop (either just one statement or a block of statements).
		// Preparser has ensured that every LOOP has a non-NULL next line.
		// apCurrentFile is sent as an arg so that more than one nested/recursive
		// file-loop can be supported:
		result = mNextLine->ExecUntil(ONLY_ONE_LINE, &jump_to_line
			, file_found ? &new_current_file : apCurrentFile // inner loop's file takes precedence over outer's.
			, apCurrentRegItem, apCurrentReadFile, aCurrentField
			, aIndex + 1);  // i+1, since 1 is the first iteration as reported to the script.
		if (result == LOOP_BREAK || result == EARLY_RETURN || result == EARLY_EXIT || result == FAIL)
		{
			#define CLOSE_FILE_SEARCH \
			if (file_search != INVALID_HANDLE_VALUE)\
			{\
				FindClose(file_search);\
				file_search = INVALID_HANDLE_VALUE;\
			}
			CLOSE_FILE_SEARCH
			// Although ExecUntil() will treat the LOOP_BREAK result identically to OK, we
			// need to return LOOP_BREAK in case our caller is another instance of this
			// same function (i.e. due to recursing into subfolders):
			return result;
		}
		if (jump_to_line == this)
		{
			// Since this LOOP's ExecUntil() encountered a Goto whose target is the LOOP
			// itself, continue with the for-loop without moving to a different
			// line.  Also: stay in this recursion layer even if aMode == ONLY_ONE_LINE
			// because we don't want the caller handling it because then it's cleanup
			// to jump to its end-point (beyond its own and any unowned elses) won't work.
			// Example:
			// if x  <-- If this layer were to do it, its own else would be unexpectedly encountered.
			//    label1:
			//    loop  <-- We want this statement's layer to handle the goto.
			//       goto, label1
			// else
			//   ...
			// Also, signal all our callers to return until they get back to the original
			// ExecUntil() instance that started the loop:
			aContinueMainLoop = true;
			break;
		}
		if (jump_to_line)
		{
			// Since the above didn't break, jump_to_line must be a line that's at the
			// same level or higher as our Exec_Until's LOOP statement itself.
			// Our ExecUntilCaller must handle the jump in either case:
			aJumpToLine = jump_to_line; // Signal the caller to handle this jump.
			break;
		}
		// Otherwise, the result of executing the body of the loop, above, was either OK
		// (the current iteration completed normally) or LOOP_CONTINUE (the current loop
		// iteration was cut short).  In both cases, just continue on through the loop.
		// But first do any end-of-iteration stuff:
		if (file_search != INVALID_HANDLE_VALUE)
		{
			for (;;)
			{
				if (   !(file_found = FindNextFile(file_search, &new_current_file))   )
					break;
				if (FileIsFilteredOut(new_current_file, aFileLoopMode, file_path))
					continue; // Ignore this one, get another one.
				else
					break;
			}
		}
	}
	// The script's loop is now over.
	CLOSE_FILE_SEARCH

	// If it's a file_loop and aRecurseSubfolders is true, we now need to perform the loop's body for
	// every subfolder to search for more files and folders inside that match aFilePattern.  We can't
	// do this in the first loop, above, because it may have a restricted file-pattern such as *.txt
	// and we want to find and recurse into ALL folders:
	if (aAttr != ATTR_LOOP_FILE || !aRecurseSubfolders)
		return OK;

	// Append *.* to file_path so that we can retrieve all files and folders in the aFilePattern
	// main folder.  We're only interested in the folders, but we have to use *.* to ensure
	// that the search will find all folder names:
	char *append_location = file_path + strlen(file_path);
	strlcpy(append_location, "*.*", sizeof(file_path) - (append_location - file_path));
	file_search = FindFirstFile(file_path, &new_current_file);
	file_found = (file_search != INVALID_HANDLE_VALUE);
	*append_location = '\0'; // Restore file_path to be just the path (i.e. remove the wildcard pattern) for use below.

	for (; file_found; file_found = FindNextFile(file_search, &new_current_file))
	{
		if (!(new_current_file.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) // We only want directories.
			continue;
		if (!strcmp(new_current_file.cFileName, "..") || !strcmp(new_current_file.cFileName, "."))
			continue; // Never recurse into these.
		// Build the new search pattern, which consists of the original file_path + the subfolder name
		// we just discovered + the original pattern:
		snprintf(append_location, sizeof(file_path) - (append_location - file_path), "%s\\%s"
			, new_current_file.cFileName, naked_filename_or_pattern);
		// Pass NULL for the 2nd param because it will determine its own current-file when it does
		// its first loop iteration.  This is because this directory is being recursed into, not
		// processed itself as a file-loop item (since this was already done in the first loop,
		// above, if its name matches the original search pattern):
		result = PerformLoop(NULL, apCurrentRegItem, apCurrentReadFile, aCurrentField
			, aContinueMainLoop, aJumpToLine, aAttr, aFileLoopMode, aRecurseSubfolders, file_path
			, aIterationLimit, aIsInfinite, aIndex);
		// result should never be LOOP_CONTINUE because the above call to PerformLoop() should have
		// handled that case.  However, it can be LOOP_BREAK if it encoutered the break command.
		if (result == LOOP_BREAK || result == EARLY_RETURN || result == EARLY_EXIT || result == FAIL)
		{
			CLOSE_FILE_SEARCH
			return result;  // Return even LOOP_BREAK, since our caller can be either ExecUntil() or ourself.
		}
		if (aContinueMainLoop) // The call to PerformLoop() above signaled us to break & return.
			break;
		// There's no need to check "aJumpToLine == this" because PerformLoop() would already have handled it.
		// But if it set aJumpToLine to be non-NULL, it means we have to return and let our caller handle
		// the jump:
		if (aJumpToLine)
			break;
	}
	CLOSE_FILE_SEARCH
	return OK;
}



ResultType Line::PerformLoopReg(WIN32_FIND_DATA *apCurrentFile, LoopReadFileStruct *apCurrentReadFile
	, char *aCurrentField, bool &aContinueMainLoop, Line *&aJumpToLine, FileLoopModeType aFileLoopMode
	, bool aRecurseSubfolders, HKEY aRootKeyType, HKEY aRootKey, char *aRegSubkey, __int64 &aIndex)
// aRootKeyType is the type of root key, independent of whether it's local or remote.
// This is used because there's no easy way to determine which root key a remote HKEY
// refers to.
{
	RegItemStruct reg_item(aRootKeyType, aRootKey, aRegSubkey);
	HKEY hRegKey;

	// Open the specified subkey.  Be sure to only open with the minimum permission level so that
	// the keys & values can be deleted or written to (though I'm not sure this would be an issue
	// in most cases):
	if (RegOpenKeyEx(reg_item.root_key, reg_item.subkey, 0, KEY_QUERY_VALUE | KEY_ENUMERATE_SUB_KEYS, &hRegKey) != ERROR_SUCCESS)
		return OK;

	// Get the count of how many values and subkeys are contained in this parent key:
	DWORD count_subkeys;
	DWORD count_values;
	if (RegQueryInfoKey(hRegKey, NULL, NULL, NULL, &count_subkeys, NULL, NULL
		, &count_values, NULL, NULL, NULL, NULL) != ERROR_SUCCESS)
	{
		RegCloseKey(hRegKey);
		return OK;
	}

	ResultType result;
	Line *jump_to_line;
	DWORD i;

	// See comments in PerformLoop() for details about this section.
	// Note that &reg_item is passed to ExecUntil() rather than 
	#define MAKE_SCRIPT_LOOP_PROCESS_THIS_ITEM \
	{\
		result = mNextLine->ExecUntil(ONLY_ONE_LINE, &jump_to_line, apCurrentFile\
			, &reg_item, apCurrentReadFile, aCurrentField, ++aIndex);\
		if (result == LOOP_BREAK || result == EARLY_RETURN || result == EARLY_EXIT || result == FAIL)\
		{\
			RegCloseKey(hRegKey);\
			return result;\
		}\
		if (jump_to_line == this)\
		{\
			aContinueMainLoop = true;\
			break;\
		}\
		if (jump_to_line)\
		{\
			aJumpToLine = jump_to_line;\
			break;\
		}\
	}

	DWORD name_size;

	// First enumerate the values, which are analogous to files in the file system.
	// Later, the subkeys ("subfolders") will be done:
	if (count_values > 0 && aFileLoopMode != FILE_LOOP_FOLDERS_ONLY) // The caller doesn't want "files" (values) excluded.
	{
		reg_item.InitForValues();
		// Going in reverse order allows values to be deleted without disrupting the enumeration,
		// at least in some cases:
		for (i = count_values - 1, jump_to_line = NULL;; --i) 
		{ 
			// Don't use CONTINUE in loops such as this due to the loop-ending condition being explicitly
			// checked at the bottom.
			name_size = sizeof(reg_item.name);  // Must reset this every time through the loop.
			*reg_item.name = '\0';
			if (RegEnumValue(hRegKey, i, reg_item.name, &name_size, NULL, &reg_item.type, NULL, NULL) == ERROR_SUCCESS)
				MAKE_SCRIPT_LOOP_PROCESS_THIS_ITEM
			// else continue the loop in case some of the lower indexes can still be retrieved successfully.
			if (i == 0)  // Check this here due to it being an unsigned value that we don't want to go negative.
				break;
		}
	}

	// If the loop is neither processing subfolders nor recursing into them, don't waste the performance
	// doing the next loop:
	if (!count_subkeys || (aFileLoopMode == FILE_LOOP_FILES_ONLY && !aRecurseSubfolders))
	{
		RegCloseKey(hRegKey);
		return OK;
	}

	// Enumerate the subkeys, which are analogous to subfolders in the files system:
	// Going in reverse order allows keys to be deleted without disrupting the enumeration,
	// at least in some cases:
	reg_item.InitForSubkeys();
	char subkey_full_path[MAX_KEY_LENGTH + 1]; // But doesn't include the root key name.
	for (i = count_subkeys - 1, jump_to_line = NULL;; --i) // Will have zero iterations if there are no subkeys.
	{
		// Don't use CONTINUE in loops such as this due to the loop-ending condition being explicitly
		// checked at the bottom.
		name_size = sizeof(reg_item.name); // Must be reset for every iteration.
		if (RegEnumKeyEx(hRegKey, i, reg_item.name, &name_size, NULL, NULL, NULL, &reg_item.ftLastWriteTime) == ERROR_SUCCESS)
		{
			if (aFileLoopMode != FILE_LOOP_FILES_ONLY) // have the script's loop process this subkey.
				MAKE_SCRIPT_LOOP_PROCESS_THIS_ITEM
			if (aRecurseSubfolders) // Now recurse into the subkey, regardless of whether it was processed above.
			{
				// Build the new subkey name using the an area of memory on the stack that we won't need
				// after the recusive call returns to us.  Omit the leading backslash if subkey is blank,
				// which supports recursively searching the contents of keys contained within a root key
				// (fixed for v1.0.17):
				snprintf(subkey_full_path, sizeof(subkey_full_path), "%s%s%s", reg_item.subkey
					, *reg_item.subkey ? "\\" : "", reg_item.name);
				// This section is very similar to the one in PerformLoop(), so see it for comments:
				result = PerformLoopReg(apCurrentFile, apCurrentReadFile, aCurrentField, aContinueMainLoop
					, aJumpToLine, aFileLoopMode, aRecurseSubfolders, aRootKeyType, aRootKey, subkey_full_path, aIndex);
				if (result == LOOP_BREAK || result == EARLY_RETURN || result == EARLY_EXIT || result == FAIL)
				{
					RegCloseKey(hRegKey);
					return result;
				}
				if (aContinueMainLoop || aJumpToLine)
					break;
			}
		}
		// else continue the loop in case some of the lower indexes can still be retrieved successfully.
		if (i == 0)  // Check this here due to it being an unsigned value that we don't want to go negative.
			break;
	}
	RegCloseKey(hRegKey);
	return OK;
}



ResultType Line::PerformLoopParse(WIN32_FIND_DATA *apCurrentFile, RegItemStruct *apCurrentRegItem
	, LoopReadFileStruct *apCurrentReadFile, bool &aContinueMainLoop, Line *&aJumpToLine, __int64 &aIndex)
{
	if (!*ARG2) // Since the input variable's contents are blank, the loop will execute zero times.
		return OK;

	// This will be used to hold the parsed items.  It needs to have its own storage because
	// even though ARG2 might always be a writable memory area, we can't rely upon it being
	// persistent because it might reside in the deref buffer, in which case the other commands
	// in the loop's body would probably overwrite it.  Even if the ARG2's contents aren't in
	// the deref buffer, we still can't modify it (i.e. to temporarily terminate it and thus
	// bypass the need for malloc() below) because that might modify the variable contents, and
	// that variable may be referenced elsewhere in the body of the loop (which would result
	// in unexpected side-effects).  So, rather than have a limit of 64K or something (which
	// would limit this feature's usefulness for parsing a large list of filenames, for example),
	// it seems best to dynamically allocate a temporary buffer large enough to hold the
	// contents of ARG2 (the input variable).  Update: Since these loops tend to be enclosed
	// by file-read loops, and thus may be called thousands of times in a short period,
	// it should help average performance to use the stack for small vars rather than
	// constantly doing malloc() and free(), which are much higher overhead and probably
	// cause memory fragmentation (especially with thousands of calls):
	char stack_buf[16384], *buf;
	size_t space_needed = strlen(ARG2) + 1;  // +1 for the zero terminator.
	if (space_needed <= sizeof(stack_buf))
		buf = stack_buf;
	else
	{
		if (   !(buf = (char *)malloc(space_needed))   )
			// Probably best to consider this a critical error, since on the rare times it does happen, the user
			// would probably want to know about it immediately.
			return LineError("Failed to make a temp copy of the input variable.", FAIL, ARG2);
	}
	strcpy(buf, ARG2); // Make the copy.

	#define FREE_PARSE_MEMORY if (buf != stack_buf) free(buf)

	// Make a copy of ARG3 and ARG4 in case either one's contents are in the deref buffer, which would
	// probably be overwritten by the commands in the script loop's body:
	char delimiters[512], omit_list[512];
	strlcpy(delimiters, ARG3, sizeof(delimiters));
	strlcpy(omit_list, ARG4, sizeof(omit_list));

	ResultType result;
	Line *jump_to_line;
	char *field, *field_end, saved_char;
	size_t field_length;

	for (field = buf;;)
	{ 
		if (*delimiters)
		{
			if (   !(field_end = StrChrAny(field, delimiters))   ) // No more delimiters found.
				field_end = field + strlen(field);  // Set it to the position of the zero terminator instead.
		}
		else // Since no delimiters, every char in the input string is treated as a separate field.
		{
			// But exclude this char if it's in the omit_list:
			if (*omit_list && strchr(omit_list, *field))
			{
				++field; // Move on to the next char.
				if (!*field) // The end of the string has been reached.
					break;
				continue;
			}
			field_end = field + 1;
		}

		saved_char = *field_end;  // In case it's a non-delimited list of single chars.
		*field_end = '\0';  // Temporarily terminate so that GetLoopField() will see the correct substring.

		if (*omit_list && *field && *delimiters)  // If no delimiters, the omit_list has already been handled above.
		{
			// Process the omit list.
			field = omit_leading_any(field, omit_list, field_end - field);
			if (*field) // i.e. the above didn't remove all the chars due to them all being in the omit-list.
			{
				field_length = omit_trailing_any(field, omit_list, field_end - 1);
				field[field_length] = '\0';  // Terminate here, but don't update field_end, since saved_char needs it.
			}
		}

		// See comments in PerformLoop() for details about this section.
		result = mNextLine->ExecUntil(ONLY_ONE_LINE, &jump_to_line, apCurrentFile, apCurrentRegItem
			, apCurrentReadFile, field, ++aIndex);

		if (result == LOOP_BREAK || result == EARLY_RETURN || result == EARLY_EXIT || result == FAIL)
		{
			FREE_PARSE_MEMORY;
			return result;
		}
		if (jump_to_line == this)
		{
			aContinueMainLoop = true;
			break;
		}
		if (jump_to_line)
		{
			aJumpToLine = jump_to_line;
			break;
		}
		if (!saved_char) // The last item in the list has just been processed, so the loop is done.
			break;
		*field_end = saved_char;  // Undo the temporary termination, in case the list of delimiters is blank.
		field = *delimiters ? field_end + 1 : field_end;  // Move on to the next field.
	}
	FREE_PARSE_MEMORY;
	return OK;
}



ResultType Line::PerformLoopParseCSV(WIN32_FIND_DATA *apCurrentFile, RegItemStruct *apCurrentRegItem
	, LoopReadFileStruct *apCurrentReadFile, bool &aContinueMainLoop, Line *&aJumpToLine, __int64 &aIndex)
// This function is similar to PerformLoopParse() so the two should be maintained together.
// See PerformLoopParse() for comments about the below (comments have been mostly stripped
// from this function).
{
	if (!*ARG2) // Since the input variable's contents are blank, the loop will execute zero times.
		return OK;

	char stack_buf[16384], *buf;
	size_t space_needed = strlen(ARG2) + 1;  // +1 for the zero terminator.
	if (space_needed <= sizeof(stack_buf))
		buf = stack_buf;
	else
	{
		if (   !(buf = (char *)malloc(space_needed))   )
			return LineError("Failed to make a temp copy of the input variable.", FAIL, ARG2);
	}
	strcpy(buf, ARG2); // Make the copy.

	#define FREE_PARSE_MEMORY if (buf != stack_buf) free(buf)

	char omit_list[512];
	strlcpy(omit_list, ARG4, sizeof(omit_list));

	ResultType result;
	Line *jump_to_line;
	char *field, *field_end, saved_char;
	size_t field_length;
	bool field_is_enclosed_in_quotes;

	for (field = buf;;)
	{
		if (*field == '"')
		{
			// For each field, check if the optional leading double-quote is present.  If it is,
			// skip over it since we always know it's the one that marks the beginning of
			// the that field.  This assumes that a field containing escaped double-quote is
			// always contained in double quotes, which is how Excel does it.  For example:
			// """string with escaped quotes""" resolves to a literal quoted string:
			field_is_enclosed_in_quotes = true;
			++field;
		}
		else
			field_is_enclosed_in_quotes = false;

		for (field_end = field;;)
		{
			if (   !(field_end = strchr(field_end, field_is_enclosed_in_quotes ? '"' : ','))   )
			{
				// This is the last field in the string, so set field_end to the position of
				// the zero terminator instead:
				field_end = field + strlen(field);
				break;
			}
			if (field_is_enclosed_in_quotes)
			{
				// The quote discovered above marks the end of the string if it isn't followed
				// by another quote.  But if it is a pair of quotes, replace it with a single
				// literal double-quote and then keep searching for the real ending quote:
				if (field_end[1] == '"')  // A pair of quotes was encountered.
				{
					memmove(field_end, field_end + 1, strlen(field_end + 1) + 1); // +1 to include terminator.
					++field_end; // Skip over the literal double quote that we just produced.
					continue; // Keep looking for the "real" ending quote.
				}
				// Otherwise, this quote marks the end of the field, so just fall through and break.
			}
			// else field is not enclosed in quotes, so the comma discovered above must be a delimiter.
			break;
		}

		saved_char = *field_end; // This can be the terminator, a comma, or a double-quote.
		*field_end = '\0';  // Terminate here so that GetLoopField() will see the correct substring.

		if (*omit_list && *field)
		{
			// Process the omit list.
			field = omit_leading_any(field, omit_list, field_end - field);
			if (*field) // i.e. the above didn't remove all the chars due to them all being in the omit-list.
			{
				field_length = omit_trailing_any(field, omit_list, field_end - 1);
				field[field_length] = '\0';  // Terminate here, but don't update field_end, since we need its pos.
			}
		}

		// See comments in PerformLoop() for details about this section.
		result = mNextLine->ExecUntil(ONLY_ONE_LINE, &jump_to_line, apCurrentFile, apCurrentRegItem
			, apCurrentReadFile, field, ++aIndex);

		if (result == LOOP_BREAK || result == EARLY_RETURN || result == EARLY_EXIT || result == FAIL)
		{
			FREE_PARSE_MEMORY;
			return result;
		}
		if (jump_to_line == this)
		{
			aContinueMainLoop = true;
			break;
		}
		if (jump_to_line)
		{
			aJumpToLine = jump_to_line;
			break;
		}

		if (!saved_char) // The last item in the list has just been processed, so the loop is done.
			break;
		if (saved_char == ',') // Set "field" to be the position of the next field.
			field = field_end + 1;
		else // saved_char must be a double-quote char.
		{
			field = field_end + 1;
			if (!*field) // No more fields occur after this one.
				break;
			// Find the next comma, which must be a real delimiter since we're in between fields:
			if (   !(field = strchr(field, ','))   ) // No more fields.
				break;
			// Set it to be the first character of the next field, which might be a double-quote
			// or another comma (if the field is empty).
			++field;
		}
	}
	FREE_PARSE_MEMORY;
	return OK;
}



ResultType Line::PerformLoopReadFile(WIN32_FIND_DATA *apCurrentFile, RegItemStruct *apCurrentRegItem
	, char *aCurrentField, bool &aContinueMainLoop, Line *&aJumpToLine, FILE *aReadFile, char *aWriteFileName
	, __int64 &aIndex)
{
	LoopReadFileStruct loop_info(aReadFile, aWriteFileName);
	size_t line_length;
	ResultType result;
	Line *jump_to_line;

	for (; fgets(loop_info.mCurrentLine, sizeof(loop_info.mCurrentLine), loop_info.mReadFile);)
	{ 
		line_length = strlen(loop_info.mCurrentLine);
		if (line_length && loop_info.mCurrentLine[line_length - 1] == '\n') // Remove newlines like FileReadLine does.
			loop_info.mCurrentLine[--line_length] = '\0';
		// See comments in PerformLoop() for details about this section.
		result = mNextLine->ExecUntil(ONLY_ONE_LINE, &jump_to_line, apCurrentFile, apCurrentRegItem
			, &loop_info, aCurrentField, ++aIndex);
		if (result == LOOP_BREAK || result == EARLY_RETURN || result == EARLY_EXIT || result == FAIL)
		{
			if (loop_info.mWriteFile)
				fclose(loop_info.mWriteFile);
			return result;
		}
		if (jump_to_line == this)
		{
			aContinueMainLoop = true;
			break;
		}
		if (jump_to_line)
		{
			aJumpToLine = jump_to_line;
			break;
		}
	}

	if (loop_info.mWriteFile)
		fclose(loop_info.mWriteFile);

	// Don't return result because we want to always return OK unless it was one of the values
	// already explicitly checked and returned above.  In other words, there might be values other
	// than OK that aren't explicitly checked for, above.
	return OK;
}



inline ResultType Line::Perform(WIN32_FIND_DATA *aCurrentFile, RegItemStruct *aCurrentRegItem
	, LoopReadFileStruct *aCurrentReadFile)
// Performs only this line's action.
// Returns OK or FAIL.
// The function should not be called to perform any flow-control actions such as
// Goto, Gosub, Return, Block-Begin, Block-End, If, Else, etc.
{
	// Since most compilers create functions that, upon entry, reserve stack space
	// for all of their automatic/local vars (even if those vars are inside
	// nested blocks whose code is never executed): Rather than having
	// a dozen or more buffers, one for each case in the switch() stmt that needs
	// one, just have one or two for general purpose use (helps conserve stack space):
	char buf_temp[LINE_SIZE]; // Work area for use in various places.  Some things rely on it being large.
	WinGroup *group; // For the group commands.
	Var *output_var = NULL; // Init to help catch bugs.
	VarSizeType space_needed; // For the commands that assign directly to an output var.
	ToggleValueType toggle;  // For commands that use on/off/neutral.
	int x, y;   // For mouse commands.
	// Use signed values for these in case they're really given an explicit negative value:
	int start_char_num, chars_to_extract; // For String commands.
	size_t source_length; // For String commands.
	SymbolType var_is_pure_numeric, value_is_pure_numeric; // For math operations.
	vk_type vk; // For mouse commands and GetKeyState.
	Label *target_label;  // For ACT_SETTIMER and ACT_HOTKEY
	int instance_number;  // For sound commands.
	DWORD component_type; // For sound commands.
	__int64 device_id;  // For sound commands.  __int64 helps avoid compiler warning for some conversions.
	bool is_remote_registry; // For Registry commands.
	HKEY root_key; // For Registry commands.
	ResultType result;  // General purpose.
	HANDLE running_process; // For RUNWAIT
	DWORD exit_code; // For RUNWAIT
	bool do_selective_blockinput, blockinput_prev;  // For the mouse commands.

	// Even though the loading-parser already checked, check again, for now,
	// at least until testing raises confidence.  UPDATE: Don't this because
	// sometimes (e.g. ACT_ASSIGN/ADD/SUB/MULT/DIV) the number of parameters
	// required at load-time is different from that at runtime, because params
	// are taken out or added to the param list:
	//if (nArgs < g_act[mActionType].MinParams) ...

	switch (mActionType)
	{
	case ACT_WINACTIVATE:
	case ACT_WINACTIVATEBOTTOM:
		if (WinActivate(FOUR_ARGS, mActionType == ACT_WINACTIVATEBOTTOM))
			// It seems best to do these sleeps here rather than in the windowing
			// functions themselves because that way, the program can use the
			// windowing functions without being subject to the script's delay
			// setting (i.e. there are probably cases when we don't need
			// to wait, such as bringing a message box to the foreground,
			// since no other actions will be dependent on it actually
			// having happened:
			DoWinDelay;
		return OK;  // Always successful, like AutoIt.
	case ACT_WINCLOSE:
	case ACT_WINKILL:
	{
		int wait_time = *ARG3 ? (int)(1000 * ATOF(ARG3)) : DEFAULT_WINCLOSE_WAIT;
		if (!wait_time) // 0 is defined as 500ms, which seems more useful than a true zero.
			wait_time = 500;
		if (WinClose(ARG1, ARG2, wait_time, ARG4, ARG5, mActionType == ACT_WINKILL))
			DoWinDelay;
		return OK;  // Always successful, like AutoIt.
	}

	case ACT_INIREAD:
		return IniRead(ARG2, ARG3, ARG4, ARG5);
	case ACT_INIWRITE:
		return IniWrite(FOUR_ARGS);
	case ACT_INIDELETE:
		// To preserve maximum compatibility with existing scripts, only send NULL if ARG3
		// was explicitly omitted.  This is because some older scripts might rely on the
		// fact that a blank ARG3 does not delete the entire section, but rather does
		// nothing (that fact is untested):
		return IniDelete(ARG1, ARG2, mArgc < 3 ? NULL : ARG3);

	case ACT_REGREAD:
		if (mArgc < 2 && aCurrentRegItem) // Uses the registry loop's current item.
			// If aCurrentRegItem->name specifies a subkey rather than a value name, do this anyway
			// so that it will set ErrorLevel to ERROR and set the output variable to be blank.
			// Also, do not use RegCloseKey() on this, even if it's a remote key, since our caller handles that:
			return RegRead(aCurrentRegItem->root_key, aCurrentRegItem->subkey, aCurrentRegItem->name);
		// Otherwise:
		if (mArgc > 4 || RegConvertValueType(ARG2)) // The obsolete 5-param method (ARG2 is unused).
			result = RegRead(root_key = RegConvertRootKey(ARG3, &is_remote_registry), ARG4, ARG5);
		else
			result = RegRead(root_key = RegConvertRootKey(ARG2, &is_remote_registry), ARG3, ARG4);
		if (is_remote_registry && root_key) // Never try to close local root keys, which the OS keeps always-open.
			RegCloseKey(root_key);
		return result;
	case ACT_REGWRITE:
		if (mArgc < 2 && aCurrentRegItem) // Uses the registry loop's current item.
			// If aCurrentRegItem->name specifies a subkey rather than a value name, do this anyway
			// so that it will set ErrorLevel to ERROR.  An error will also be indicated if
			// aCurrentRegItem->type is an unsupported type:
			return RegWrite(aCurrentRegItem->type, aCurrentRegItem->root_key, aCurrentRegItem->subkey, aCurrentRegItem->name, ARG1);
		// Otherwise:
		result = RegWrite(RegConvertValueType(ARG1), root_key = RegConvertRootKey(ARG2, &is_remote_registry)
			, ARG3, ARG4, ARG5); // If RegConvertValueType(ARG1) yields REG_NONE, RegWrite() will set ErrorLevel rather than displaying a runtime error.
		if (is_remote_registry && root_key) // Never try to close local root keys, which the OS keeps always-open.
			RegCloseKey(root_key);
		return result;
	case ACT_REGDELETE:
		if (mArgc < 1 && aCurrentRegItem) // Uses the registry loop's current item.
		{
			// In this case, if the CurrentRegItem is a value, just delete it normally.
			// But if it's a subkey, append it to the dir name so that the proper subkey
			// will be deleted as the user intended:
			if (aCurrentRegItem->type == REG_SUBKEY)
			{
				snprintf(buf_temp, sizeof(buf_temp), "%s\\%s", aCurrentRegItem->subkey, aCurrentRegItem->name);
				return RegDelete(aCurrentRegItem->root_key, buf_temp, "");
			}
			else
				return RegDelete(aCurrentRegItem->root_key, aCurrentRegItem->subkey, aCurrentRegItem->name);
		}
		// Otherwise:
		result = RegDelete(root_key = RegConvertRootKey(ARG1, &is_remote_registry), ARG2, ARG3);
		if (is_remote_registry && root_key) // Never try to close local root keys, which the OS always keeps open.
			RegCloseKey(root_key);
		return result;

	case ACT_SHUTDOWN:
		return Util_Shutdown(ATOI(ARG1)) ? OK : FAIL; // Range of ARG1 is not validated in case other values are supported in the future.

	case ACT_SLEEP:
	{
		// Only support 32-bit values for this command, since it seems unlikely anyone would to have
		// it sleep more than 24.8 days or so.  It also helps performance on 32-bit hardware because
		// MsgSleep() is so heavily called and checks the value of the first parameter frequently:
		int sleep_time = ATOI(ARG1);

		// Do a true sleep on Win9x because the MsgSleep() method is very inaccurate on Win9x
		// for some reason (a MsgSleep(1) causes a sleep between 10 and 55ms, for example).
		// But only do so for short sleeps, for which the user has a greater expectation of
		// accuracy:
		if (sleep_time < 25 && g_os.IsWin9x())
			Sleep(sleep_time);
		else
			MsgSleep(sleep_time);
		return OK;
	}
	case ACT_ENVSET:
		// MSDN: "If [the 2nd] parameter is NULL, the variable is deleted from the current process’s environment."
		// My: Though it seems okay, for now, just to set it to be blank if the user omitted the 2nd param or
		// left it blank (AutoIt3 does this too).  Also, no checking is currently done to ensure that ARG2
		// isn't longer than 32K, since future OSes may support longer env. vars.  SetEnvironmentVariable()
		// might return 0(fail) in that case anyway.  Also, ARG1 may be a dereferenced variable that resolves
		// to the name of an Env. Variable.  In any case, this name need not correspond to any existing
		// variable name within the script (i.e. script variables and env. variables aren't tied to each other
		// in any way).  This seems to be the most flexible approach, but are there any shortcomings?
		// The only one I can think of is that if the script tries to fetch the value of an env. var (perhaps
		// one that some other spawned program changed), and that var's name corresponds to the name of a
		// script var, the script var's value (if non-blank) will be fetched instead.
		// Note: It seems, at least under WinXP, that env variable names can contain spaces.  So it's best
		// not to validate ARG1 the same way we validate script variables (i.e. just let\
		// SetEnvironmentVariable()'s return value determine whether there's an error).  However, I just
		// realized that it's impossible to "retrieve" the value of an env var that has spaces since
		// there is no EnvGet() command (EnvGet() is implicit whenever an undefined or blank script
		// variable is dereferenced).  For now, this is documented here as a known limitation.
		return g_ErrorLevel->Assign(SetEnvironmentVariable(ARG1, ARG2) ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);

	case ACT_ENVUPDATE:
	{
		// From the AutoIt3 source:
		// AutoIt3 uses SMTO_BLOCK (which prevents our thread from doing anything during the call)
		// vs. SMTO_NORMAL.  Since I'm not sure why, I'm leaving it that way for now:
		ULONG nResult;
		if (SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0, (LPARAM)"Environment", SMTO_BLOCK, 15000, &nResult))
			return g_ErrorLevel->Assign(ERRORLEVEL_NONE);
		else
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	}

	case ACT_URLDOWNLOADTOFILE:
		return URLDownloadToFile(TWO_ARGS);

	case ACT_RUNAS:
		if (!g_os.IsWin2000orLater()) // Do nothing if the OS doesn't support it.
			return OK;
		if (mArgc < 1)
		{
			if (!g_script.mRunAsUser) // memory not yet allocated so nothing needs to be done.
				return OK;
			*g_script.mRunAsUser = *g_script.mRunAsPass = *g_script.mRunAsDomain = 0; // wide-character terminator.
			return OK;
		}
		// Otherwise, the credentials are being set or updated:
		if (!g_script.mRunAsUser) // allocate memory (only needed the first time this is done).
		{
			// It's more memory efficient to allocate a single block and divide it up.
			// This memory is freed automatically by the OS upon program termination.
			if (   !(g_script.mRunAsUser = (wchar_t *)malloc(3 * RUNAS_ITEM_SIZE))   )
				return LineError(ERR_OUTOFMEM ERR_ABORT);
			g_script.mRunAsPass = g_script.mRunAsUser + RUNAS_ITEM_SIZE;
			g_script.mRunAsDomain = g_script.mRunAsPass + RUNAS_ITEM_SIZE;
		}
		mbstowcs(g_script.mRunAsUser, ARG1, RUNAS_ITEM_SIZE);
		mbstowcs(g_script.mRunAsPass, ARG2, RUNAS_ITEM_SIZE);
		mbstowcs(g_script.mRunAsDomain, ARG3, RUNAS_ITEM_SIZE);
		return OK;

	case ACT_RUN: // Be sure to pass NULL for 2nd param.
		if (strcasestr(ARG3, "UseErrorLevel"))
			return g_ErrorLevel->Assign(g_script.ActionExec(ARG1, NULL, ARG2, false, ARG3, NULL, true
				, ResolveVarOfArg(3)) ? ERRORLEVEL_NONE : "ERROR");
			// The special string ERROR is used, rather than a number like 1, because currently
			// RunWait might in the future be able to return any value, including 259 (STATUS_PENDING).
		else // If launch fails, display warning dialog and terminate current thread.
			return g_script.ActionExec(ARG1, NULL, ARG2, true, ARG3, NULL, true, ResolveVarOfArg(3));

	case ACT_RUNWAIT:
		if (strcasestr(ARG3, "UseErrorLevel"))
		{
			if (!g_script.ActionExec(ARG1, NULL, ARG2, false, ARG3, &running_process, true, NULL))
				return g_ErrorLevel->Assign("ERROR"); // See above comment for explanation.
			//else fall through to the waiting-phase of the operation.
			// Above: The special string ERROR is used, rather than a number like 1, because currently
			// RunWait might in the future be able to return any value, including 259 (STATUS_PENDING).
		}
		else // If launch fails, display warning dialog and terminate current thread.
			if (!g_script.ActionExec(ARG1, NULL, ARG2, true, ARG3, &running_process, true, NULL))
				return FAIL;
			//else fall through to the waiting-phase of the operation.

	case ACT_CLIPWAIT:
	case ACT_KEYWAIT:
	case ACT_WINWAIT:
	case ACT_WINWAITCLOSE:
	case ACT_WINWAITACTIVE:
	case ACT_WINWAITNOTACTIVE:
	{
		bool wait_indefinitely;
		int sleep_duration;
		DWORD start_time;
		// For ACT_KEYWAIT:
		bool wait_for_keydown;
		KeyStateTypes key_state_type;
		JoyControls joy;
		int joystick_id;

		if (mActionType == ACT_KEYWAIT)
		{
			if (   !(vk = TextToVK(ARG1))   )
			{
				if (   !(joy = (JoyControls)ConvertJoy(ARG1, &joystick_id))   ) // Not a valid key name.
					// Indicate immediate timeout (if timeout was specified) or error.
					return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
				if (!IS_JOYSTICK_BUTTON(joy)) // Currently, only buttons are supported.
					return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
			}
			// Set defaults:
			wait_for_keydown = false;  // The default is to wait for the key to be released.
			key_state_type = KEYSTATE_PHYSICAL;  // Since physical is more often used.
			wait_indefinitely = true;
			sleep_duration = 0;
			for (char *cp = ARG2; *cp; ++cp)
			{
				switch(toupper(*cp))
				{
				case 'D':
					wait_for_keydown = true;
					break;
				case 'L':
					key_state_type = KEYSTATE_LOGICAL;
					break;
				case 'T':
					// Although ATOF() supports hex, it's been documented in the help file that hex should
					// not be used (see comment above) so if someone does it anyway, some option letters
					// might be misinterpreted:
					wait_indefinitely = false;
					sleep_duration = (int)(ATOF(cp + 1) * 1000);
					start_time = GetTickCount();
					break;
				}
			}
		}
		else if (   (mActionType != ACT_RUNWAIT && mActionType != ACT_CLIPWAIT && *ARG3)
			|| (mActionType == ACT_CLIPWAIT && *ARG1)   )
		{
			// Since the param containing the timeout value isn't blank, it must be numeric,
			// otherwise, the loading validation would have prevented the script from loading.
			wait_indefinitely = false;
			sleep_duration = (int)(ATOF(mActionType == ACT_CLIPWAIT ? ARG1 : ARG3) * 1000); // Can be zero.
			if (sleep_duration <= 0)
				// Waiting 500ms in place of a "0" seems more useful than a true zero, which
				// doens't need to be supported because it's the same thing as something like
				// "IfWinExist".  A true zero for clipboard would be the same as
				// "IfEqual, clipboard, , xxx" (though admittedly it's higher overhead to
				// actually fetch the contents of the clipboard).
				sleep_duration = 500;
			start_time = GetTickCount();
		}
		else
		{
			wait_indefinitely = true;
			sleep_duration = 0; // Just to catch any bugs.
		}

		if (mActionType != ACT_RUNWAIT) // Set default error-level unless overridden below.
			g_ErrorLevel->Assign(ERRORLEVEL_NONE);

		// Right before starting the wait-loop, make a copy of our args using the stack
		// space in our recursion layer.  This is done in case other hotkey subroutine(s)
		// are launched while we're waiting here, which might cause our args to be overwritten
		// if any of them happen to be in the Deref buffer:
		char *arg[MAX_ARGS], *marker;
		int i, space_remaining;
		for (i = 0, space_remaining = (int)sizeof(buf_temp), marker = buf_temp; i < mArgc; ++i)
		{
			if (!space_remaining) // Realistically, should never happen.
				arg[i] = "";
			else
			{
				arg[i] = marker;  // Point it to its place in the buffer.
				strlcpy(marker, sArgDeref[i], space_remaining); // Make the copy.
				marker += strlen(marker) + 1;  // +1 for the zero terminator of each arg.
				space_remaining = (int)(sizeof(buf_temp) - (marker - buf_temp));
			}
		}

		for (;;)
		{ // Always do the first iteration so that at least one check is done.
			switch(mActionType)
			{
			case ACT_WINWAIT:
				#define SAVED_WIN_ARGS SAVED_ARG1, SAVED_ARG2, SAVED_ARG4, SAVED_ARG5
				if (WinExist(SAVED_WIN_ARGS, false, true))
				{
					DoWinDelay;
					return OK;
				}
				break;
			case ACT_WINWAITCLOSE:
				if (!WinExist(SAVED_WIN_ARGS))
				{
					DoWinDelay;
					return OK;
				}
				break;
			case ACT_WINWAITACTIVE:
				if (WinActive(SAVED_WIN_ARGS, true))
				{
					DoWinDelay;
					return OK;
				}
				break;
			case ACT_WINWAITNOTACTIVE:
				if (!WinActive(SAVED_WIN_ARGS, true))
				{
					DoWinDelay;
					return OK;
				}
				break;
			case ACT_CLIPWAIT:
				// Seems best to consider CF_HDROP to be a non-empty clipboard, since we
				// support the implicit conversion of that format to text:
				if (IsClipboardFormatAvailable(CF_TEXT) || IsClipboardFormatAvailable(CF_HDROP))
					return OK;
				break;
			case ACT_KEYWAIT:
				if (vk) // Waiting for key or mouse button, not joystick.
				{
					if (ScriptGetKeyState(vk, key_state_type) == wait_for_keydown)
						return OK;
				}
				else // Waiting for joystick button
				{
					if ((bool)ScriptGetJoyState(joy, joystick_id, NULL) == wait_for_keydown)
						return OK;
				}
				break;
			case ACT_RUNWAIT:
				// Pretty nasty, but for now, nothing is done to prevent an infinite loop.
				// In the future, maybe OpenProcess() can be used to detect if a process still
				// exists (is there any other way?):
				// MSDN: "Warning: If a process happens to return STILL_ACTIVE (259) as an error code,
				// applications that test for this value could end up in an infinite loop."
				if (running_process)
					GetExitCodeProcess(running_process, &exit_code);
				else // it can be NULL in the case of launching things like "find D:\" or "www.yahoo.com"
					exit_code = 0;
				if (exit_code != STATUS_PENDING) // STATUS_PENDING == STILL_ACTIVE
				{
					if (running_process)
						CloseHandle(running_process);
					// Use signed vs. unsigned, since that is more typical?  No, it seems better
					// to use unsigned now that script variables store 64-bit ints.  This is because
					// GetExitCodeProcess() yields a DWORD, implying that the value should be unsigned.
					// Unsigned also is more useful in cases where an app returns a (potentially large)
					// count of something as its result.  However, if this is done, it won't be easy
					// to check against a return value of -1, for example, which I suspect many apps
					// return.  AutoIt3 (and probably 2) use a signed int as well, so that is another
					// reason to keep it this way:
					return g_ErrorLevel->Assign((int)exit_code);
				}
				break;
			}

			// Must cast to int or any negative result will be lost due to DWORD type:
			if (wait_indefinitely || (int)(sleep_duration - (GetTickCount() - start_time)) > SLEEP_INTERVAL_HALF)
				MsgSleep(INTERVAL_UNSPECIFIED); // INTERVAL_UNSPECIFIED performs better.
			else // Done waiting.
				return g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Since it timed out, we override the default with this.
		} // for()
		break; // Never executed, just for safety.
	}

	case ACT_WINMOVE:
		return mArgc > 2 ? WinMove(EIGHT_ARGS) : WinMove("", "", ARG1, ARG2);

	case ACT_WINMENUSELECTITEM:
		return WinMenuSelectItem(ELEVEN_ARGS);

	case ACT_CONTROLSEND:
	case ACT_CONTROLSENDRAW:
		return ControlSend(SIX_ARGS, mActionType == ACT_CONTROLSENDRAW);

	case ACT_CONTROLCLICK:
		if (*ARG4)
		{
			if (   !(vk = ConvertMouseButton(ARG4))   )
				return LineError(ERR_MOUSE_BUTTON ERR_ABORT, FAIL, ARG4);
		}
		else // Default button when the param is blank or an reference to an empty var.
			vk = VK_LBUTTON;
		return ControlClick(vk, *ARG5 ? ATOI(ARG5) : 1, ARG6, ARG1, ARG2, ARG3, ARG7, ARG8);

	case ACT_CONTROLMOVE:
		return ControlMove(NINE_ARGS);
	case ACT_CONTROLGETPOS:
		return ControlGetPos(ARG5, ARG6, ARG7, ARG8, ARG9);
	case ACT_CONTROLGETFOCUS:
		return ControlGetFocus(ARG2, ARG3, ARG4, ARG5);
	case ACT_CONTROLFOCUS:
		return ControlFocus(FIVE_ARGS);
	case ACT_CONTROLSETTEXT:
		return ControlSetText(SIX_ARGS);
	case ACT_CONTROLGETTEXT:
		return ControlGetText(ARG2, ARG3, ARG4, ARG5, ARG6);
	case ACT_CONTROL:
		return Control(SEVEN_ARGS);
	case ACT_CONTROLGET:
		return ControlGet(ARG2, ARG3, ARG4, ARG5, ARG6, ARG7, ARG8);
	case ACT_STATUSBARGETTEXT:
		return StatusBarGetText(ARG2, ARG3, ARG4, ARG5, ARG6);
	case ACT_STATUSBARWAIT:
		return StatusBarWait(EIGHT_ARGS);
	case ACT_POSTMESSAGE:
		return ScriptPostMessage(EIGHT_ARGS);
	case ACT_SENDMESSAGE:
		return ScriptSendMessage(EIGHT_ARGS);
	case ACT_PROCESS:
		return ScriptProcess(THREE_ARGS);
	case ACT_WINSET:
		return WinSet(SIX_ARGS);
	case ACT_WINSETTITLE:
		return mArgc > 1 ? WinSetTitle(FIVE_ARGS) : WinSetTitle("", "", ARG1);
	case ACT_WINGETTITLE:
		return WinGetTitle(ARG2, ARG3, ARG4, ARG5);
	case ACT_WINGETCLASS:
		return WinGetClass(ARG2, ARG3, ARG4, ARG5);
	case ACT_WINGET:
		return WinGet(ARG2, ARG3, ARG4, ARG5, ARG6);
	case ACT_WINGETTEXT:
		return WinGetText(ARG2, ARG3, ARG4, ARG5);
	case ACT_WINGETPOS:
		return WinGetPos(ARG5, ARG6, ARG7, ARG8);

	case ACT_SYSGET:
		return SysGet(ARG2, ARG3);

	case ACT_PIXELSEARCH:
		return PixelSearch(ATOI(ARG3), ATOI(ARG4), ATOI(ARG5), ATOI(ARG6), ATOI(ARG7), ATOI(ARG8), !stricmp(ARG9, "RGB"));
	//case ACT_IMAGESEARCH:
	//	return ImageSearch(ATOI(ARG3), ATOI(ARG4), ATOI(ARG5), ATOI(ARG6), ARG7);
	case ACT_PIXELGETCOLOR:
		return PixelGetColor(ATOI(ARG2), ATOI(ARG3), !stricmp(ARG4, "RGB"));
	case ACT_WINMINIMIZEALL: // From AutoIt3 source:
		PostMessage(FindWindow("Shell_TrayWnd", NULL), WM_COMMAND, 419, 0);
		DoWinDelay;
		return OK;
	case ACT_WINMINIMIZEALLUNDO: // From AutoIt3 source:
		PostMessage(FindWindow("Shell_TrayWnd", NULL), WM_COMMAND, 416, 0);
		DoWinDelay;
		return OK;

	case ACT_WINMINIMIZE:
	case ACT_WINMAXIMIZE:
	case ACT_WINHIDE:
	case ACT_WINSHOW:
	case ACT_WINRESTORE:
		return PerformShowWindow(mActionType, FOUR_ARGS);

	case ACT_ONEXIT:
		if (!*ARG1) // Reset to normal Exit behavior.
		{
			g_script.mOnExitLabel = NULL;
			return OK;
		}
		// If it wasn't resolved at load-time, it must be a variable reference:
		if (   !(target_label = (Label *)mAttribute)   )
			if (   !(target_label = g_script.FindLabel(ARG1))   )
				return LineError(ERR_ONEXIT_LABEL ERR_ABORT, FAIL, ARG1);
		g_script.mOnExitLabel = target_label;
		return OK;

	case ACT_HOTKEY:
	{
		HookActionType hook_action = 0; // Set default.
		// If it wasn't resolved at load-time, it must be a variable reference or a special value:
		if (   !(target_label = (Label *)mAttribute)   )
			if (   !(hook_action = Hotkey::ConvertAltTab(ARG2, true))   )
				if (   *ARG2 && !(target_label = g_script.FindLabel(ARG2))   )  // Allow ARG2 to be blank.
					return LineError(ERR_HOTKEY_LABEL ERR_ABORT, FAIL, ARG2);
		return Hotkey::Dynamic(ARG1, target_label, hook_action, ARG3);
	}

	case ACT_SETTIMER: // A timer is being created, changed, or enabled/disabled.
		// Note that only one timer per label is allowed because the label is the unique identifier
		// that allows us to figure out whether to "update or create" when searching the list of timers.
		if (   !(target_label = (Label *)mAttribute)   ) // Since it wasn't resolved at load-time, it must be a variable reference.
			if (   !(target_label = g_script.FindLabel(ARG1))   )
				return LineError(ERR_SETTIMER ERR_ABORT, FAIL, ARG1);
		// And don't update mAttribute (leave it NULL) because we want ARG1 to be dynamically resolved
		// every time the command is executed (in case the contents of the referenced variable change).
		// In the data structure that holds the timers, we store the target label rather than the target
		// line so that a label can be registered independently as a timers even if there another label
		// that points to the same line such as in this example:
		// Label1::
		// Label2::
		// ...
		// return
		if (*ARG2)
		{
			toggle = Line::ConvertOnOff(ARG2);
			if (!toggle && !IsPureNumeric(ARG2, true, true, true)) // Allow it to be neg. or floating point at runtime.
				return LineError("Parameter #2 must be either a pure number or the word ON or OFF.", FAIL, ARG2);
		}
		else
			toggle = TOGGLE_INVALID;
		// Below relies on distinguishing a true empty string from one that is sent to the function
		// as empty as a signal.  Don't change it without a full understanding because it's likely
		// to break compatibility or something else:
		switch(toggle)
		{
		case TOGGLED_ON:  g_script.UpdateOrCreateTimer(target_label, "", ARG3, true, false); break;
		case TOGGLED_OFF: g_script.UpdateOrCreateTimer(target_label, "", ARG3, false, false); break;
		// Timer is always (re)enabled when ARG2 specifies a numeric period or is blank + there's no ARG3.
		// If ARG2 is blank but ARG3 (priority) isn't, tell it to update only the priority and nothing else:
		default: g_script.UpdateOrCreateTimer(target_label, ARG2, ARG3, true, !*ARG2 && *ARG3);
		}
		return OK;

	case ACT_THREAD:
		switch (ConvertThreadCommand(ARG1))
		{
		case THREAD_CMD_PRIORITY:
			g.Priority = ATOI(ARG2);
			break;
		case THREAD_CMD_INTERRUPT:
			// If either one is blank, leave that setting as it was before.
			if (*ARG1)
				g_script.mUninterruptibleTime = ATOI(ARG2);  // 32-bit (for compatibility with DWORDs returned by GetTickCount).
			if (*ARG2)
				g_script.mUninterruptedLineCountMax = ATOI(ARG3);  // 32-bit also, to help performance (since huge values seem unnecessary).
			break;
		// If invalid command, do nothing since that is always caught at load-time unless the command
		// is in a variable reference (very rare in this case).
		}
		return OK;

	case ACT_GROUPADD: // Adding a WindowSpec *to* a group, not adding a group.
	{
		if (!*ARG2 && !*ARG3 && !*ARG5 && !*ARG6) // Arg4 is the jump-to label.
			// Unlike commands such as IfWinExist, we DO validate that the
			// expanded (dereferenced) window params have at least one non-blank
			// string among them, since it seems likely to be an error the user
			// would want to know about.  It's also likely the error would
			// occur immediately since ACT_GROUPADD tends to be used only in the
			// auto-execute part of the script:
			return LineError(ERR_WINDOW_PARAM, WARN);
		if (   !(group = (WinGroup *)mAttribute)   )
			if (   !(group = g_script.FindOrAddGroup(ARG1))   )
				return FAIL;  // It already displayed the error for us.
		Line *jump_to_line = NULL;
		if (*ARG4)
		{
			jump_to_line = mRelatedLine;
			if (!jump_to_line) // Jump target hasn't been resolved yet, probably due to it being a deref.
			{
				Label *label = g_script.FindLabel(ARG4);
				if (!label)
					return LineError(ERR_GROUPADD_LABEL ERR_ABORT, FAIL, ARG4);
				jump_to_line = label->mJumpToLine; // The script loader has ensured that this can't be NULL.
			}
			// Can't do this because the current line won't be the launching point for the
			// Gosub.  Instead, the launching point will be the GroupActivate rather than the
			// GroupAdd, so it will be checked by the GroupActivate or not at all (since it's
			// not that important in the case of a Gosub -- it's mostly for Goto's):
			//return IsJumpValid(label->mJumpToLine);
		}
		return group->AddWindow(ARG2, ARG3, jump_to_line, ARG5, ARG6);
	}

	// Note ACT_GROUPACTIVATE is handled by ExecUntil(), since it's better suited to do the Gosub.
	case ACT_GROUPDEACTIVATE:
		if (   !(group = (WinGroup *)mAttribute)   )
			if (   !(group = g_script.FindOrAddGroup(ARG1))   )
				return FAIL;  // It already displayed the error for us.
		group->Deactivate(*ARG2 && !stricmp(ARG2, "R"));  // Note: It will take care of DoWinDelay if needed.
		return OK;

	case ACT_GROUPCLOSE:
		if (   !(group = (WinGroup *)mAttribute)   )
			if (   !(group = g_script.FindOrAddGroup(ARG1))   )
				return FAIL;  // It already displayed the error for us.
		if (*ARG2 && !stricmp(ARG2, "A"))
			group->CloseAll();  // Note: It will take care of DoWinDelay if needed.
		else
			group->CloseAndGoToNext(*ARG2 && !stricmp(ARG2, "R"));  // Note: It will take care of DoWinDelay if needed.
		return OK;

	case ACT_TRANSFORM:
		return Transform(ARG2, ARG3, ARG4);

	case ACT_STRINGLEFT:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		chars_to_extract = ATOI(ARG3); // Use 32-bit signed to detect negatives and fit it VarSizeType.
		if (chars_to_extract < 0)
			// For these we don't report an error, since it might be intentional for
			// it to be called this way, in which case it will do nothing other than
			// set the output var to be blank.
			chars_to_extract = 0;
		// It will display any error that occurs.  Also, tell it to trim if AutoIt2 mode is in effect,
		// because I think all of AutoIt2's string commands trim the output variable as part of the
		// process.  But don't change Assign() to unconditionally trim because AutoIt2 does not trim
		// for some of its commands, such as FileReadLine.  UPDATE: AutoIt2 apprarently doesn't trim
		// when one of the STRING commands does an assignment.
		return output_var->Assign(ARG2, (VarSizeType)strnlen(ARG2, chars_to_extract));

	case ACT_STRINGRIGHT:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		chars_to_extract = ATOI(ARG3); // Use 32-bit signed to detect negatives and fit it VarSizeType.
		if (chars_to_extract < 0)
			chars_to_extract = 0;
		source_length = strlen(ARG2);
		if (chars_to_extract > (int)source_length)
			chars_to_extract = (int)source_length;
		// It will display any error that occurs:
		return output_var->Assign(ARG2 + source_length - chars_to_extract, chars_to_extract);

	case ACT_STRINGMID:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		chars_to_extract = ATOI(ARG4); // Use 32-bit signed to detect negatives and fit it VarSizeType.
		if (chars_to_extract < 0)
			return output_var->Assign();  // Set it to be blank in this case.
		start_char_num = ATOI(ARG3);
		if (toupper(*ARG5) == 'L')  // Chars to the left of start_char_num will be extracted.
		{
			if (start_char_num < 1)
				return output_var->Assign();  // Blank seems most appropriate for the L option in this case.
			start_char_num -= (chars_to_extract - 1);
			if (start_char_num < 1)
				// Reduce chars_to_extract to reflect the fact that there aren't enough chars
				// to the left of start_char_num, so we'll extract only them:
				chars_to_extract -= (1 - start_char_num);
		}
		// UPDATE: The below is now also needed for the L option to work correctly.  Older:
		// It's somewhat debatable, but it seems best not to report an error in this and
		// other cases.  The result here is probably enough to speak for itself, for script
		// debugging purposes:
		if (start_char_num < 1)
			start_char_num = 1; // 1 is the position of the first char, unlike StringGetPos.
		// Assign() is capable of doing what we want in this case.
		// It will display any error that occurs:
		if ((int)strlen(ARG2) < start_char_num)
			return output_var->Assign();  // Set it to be blank in this case.
		else
			return output_var->Assign(ARG2 + start_char_num - 1, chars_to_extract);

	case ACT_STRINGTRIMLEFT:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		chars_to_extract = ATOI(ARG3); // Use 32-bit signed to detect negatives and fit it VarSizeType.
		if (chars_to_extract < 0)
			chars_to_extract = 0;
		source_length = strlen(ARG2);
		if (chars_to_extract > (int)source_length) // This could be intentional, so don't display an error.
			chars_to_extract = (int)source_length;
		return output_var->Assign(ARG2 + chars_to_extract, (VarSizeType)(source_length - chars_to_extract));

	case ACT_STRINGTRIMRIGHT:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		chars_to_extract = ATOI(ARG3); // Use 32-bit signed to detect negatives and fit it VarSizeType.
		if (chars_to_extract < 0)
			chars_to_extract = 0;
		source_length = strlen(ARG2);
		if (chars_to_extract > (int)source_length) // This could be intentional, so don't display an error.
			chars_to_extract = (int)source_length;
		return output_var->Assign(ARG2, (VarSizeType)(source_length - chars_to_extract)); // It already displayed any error.

	case ACT_STRINGLOWER:
	case ACT_STRINGUPPER:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		space_needed = (VarSizeType)(strlen(ARG2) + 1);
		// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
		// this call will set up the clipboard for writing:
		if (output_var->Assign(NULL, space_needed - 1) != OK)
			return FAIL;
		// Copy the input variable's text directly into the output variable:
		strlcpy(output_var->Contents(), ARG2, space_needed);
		if (*ARG3 && toupper(*ARG3) == 'T' && !*(ARG3 + 1)) // Convert to title case
			StrToTitleCase(output_var->Contents());
		else if (mActionType == ACT_STRINGLOWER)
			CharLower(output_var->Contents());
		else
			CharUpper(output_var->Contents());
		return output_var->Close();  // In case it's the clipboard.

	case ACT_STRINGLEN:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		// Since ARG2 can be a reserved variable whose contents does not have an accurate
		// Var::mLength member, strlen() is always called explicitly rather than just
		// returning mLength.  Another reason to call strlen explicitly is for a case such
		// as StringLen, OutputVar, %InputVar%, where InputVar contains the name of the
		// variable to be checked.  ExpandArgs() has already taken care of resolving that,
		// so the only way to avoid strlen() would be to avoid the call to ExpandArgs()
		// in ExecUntil(), and instead resolve the variable name here directly.
		return output_var->Assign((__int64)strlen(ARG2)); // It already displayed any error.

	case ACT_STRINGGETPOS:
	{
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		bool search_from_the_right = toupper(*ARG4) == 'R' || *ARG4 == '1';
		int i, pos = -1, occurrence_number = 1; // Set defaults.
		if (strchr("LR", toupper(*ARG4)))
			occurrence_number = *(ARG4 + 1) ? ATOI(ARG4 + 1) : 1;
		// Intentionally allow occurrence_number to resolve to a negative, for scripting flexibililty:
		if (occurrence_number > 0)
		{
			if (!*ARG3) // It might be intentional, in obscure cases, to search for the empty string.
				pos = 0;
				// Above: empty string is always found immediately (first char from left) regardless
				// of whether search_from_the_right is true.  This is because it's too rare to
				// worry about giving it any more explicit handling based on search_from_the_right.
			else
			{
				char *found, *search_str = ARG3;
				size_t search_str_length = strlen(search_str);
				if (search_from_the_right)
					// Want it to behave like in this example: If searching for the 2nd occurrence of
					// FF in the string FFFF, it should find the first two F's, not the middle two:
					found = strrstr(ARG2, search_str, g.StringCaseSense, occurrence_number);
				else
				{
					// Want it to behave like in this example: If searching for the 2nd occurrence of
					// FF in the string FFFF, it should find position 3 (the 2nd pair), not position 2:
					for (i = 1, found = ARG2; ; ++i, found += search_str_length)
					{
						if (   !(found = g.StringCaseSense ? strstr(found, search_str) : strcasestr(found, search_str))   )
							break;
						if (i == occurrence_number)
							break;
					}
				}
				if (found)
					pos = (int)(found - ARG2);
				// else leave pos set to its default value, -1.
			}
		}
		g_ErrorLevel->Assign(pos < 0 ? ERRORLEVEL_ERROR : ERRORLEVEL_NONE);
		return output_var->Assign(pos); // Assign() already displayed any error that may have occurred.
	}

	case ACT_STRINGREPLACE:
	{
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		source_length = strlen(ARG2);
		space_needed = (VarSizeType)source_length + 1;  // Set default, or starting value for accumulation.
		VarSizeType final_space_needed = space_needed;
		bool do_replace = *ARG2 && *ARG3; // i.e. don't allow replacement of the empty string.
		bool always_use_slow_mode = strcasestr(ARG5, "AllSlow");
		bool alternate_error_level = strcasestr(ARG5, "UseErrorLevel");
		// Both AllSlow and UseErrorLevel imply "replace all":
		bool replace_all = always_use_slow_mode || alternate_error_level || StrChrAny(ARG5, "1aA");
		DWORD found_count = 0; // Set default.

		if (do_replace) 
		{
			// Note: It's okay if Search String is a subset of Replace String.
			// Example: Replacing all occurrences of "a" with "abc" would be
			// safe the way this StrReplaceAll() works (other implementations
			// might cause on infinite loop).
			size_t search_str_len = strlen(ARG3);
			size_t replace_str_len = strlen(ARG4);
			char *found_pos;
			for (found_pos = ARG2;;)
			{
				if (   !(found_pos = g.StringCaseSense ? strstr(found_pos, ARG3) : strcasestr(found_pos, ARG3))   )
					break;
				++found_count;
				// Jump to the end of the string that was just found, in preparation
				// for the next iteration:
				found_pos += search_str_len;
				if (!replace_all) // Replacing only one, so we're done.
					break;
			}
			final_space_needed += (int)found_count * (int)(replace_str_len - search_str_len); // Must cast to int in case value is negative.
			// Use the greater of the two because temporarily need more space in the output
			// var, because that is where the replacement will be conducted:
			if (final_space_needed > space_needed)
				space_needed = final_space_needed;
		}

		// AutoIt2: "If the search string cannot be found, the contents of <Output Variable>
		// will be the same as <Input Variable>."
		if (alternate_error_level)
			g_ErrorLevel->Assign(found_count);
		else
			g_ErrorLevel->Assign(found_count ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR); // So that it behaves like AutoIt2.
	
		// Handle the output parameter.  This section is similar to that in PerformAssign().
		// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
		// this call will set up the clipboard for writing:
		if (output_var->Assign(NULL, space_needed - 1) != OK)
			return FAIL;
		// Fetch the text directly into the var:
		if (space_needed == 1)
			*output_var->Contents() = '\0';
		else
			strlcpy(output_var->Contents(), ARG2, space_needed);
		output_var->Length() = final_space_needed - 1;  // This will be the length after replacement is done.

		// Now that we've put a copy of the Input Variable into the Output Variable,
		// and ensured that the output var is large enough to holder either
		// the original or the new version (whichever is greatest), we can actually
		// do the replacement.
		if (do_replace)
			if (replace_all)
				// Note: The current implementation of StrReplaceAll() should be
				// able to handle any conceivable inputs without causing an infinite
				// loop due to empty string (since we already checked for the empty
				// string above) and without going infinite due to finding the
				// search string inside of newly-inserted replace strings (e.g.
				// replacing all occurrences of b with bcd would not keep finding
				// b in the newly inserted bcd, infinitely):
				StrReplaceAll(output_var->Contents(), ARG3, ARG4, always_use_slow_mode, g.StringCaseSense, found_count);
			else
				StrReplace(output_var->Contents(), ARG3, ARG4, g.StringCaseSense); // Don't pass output_var->Length() because it's not up-to-date yet.

		// UPDATE: This is NOT how AutoIt2 behaves, so don't do it:
		//if (g_script.mIsAutoIt2)
		//{
		//	trim(output_var->Contents());  // Since this is how AutoIt2 behaves.
		//	output_var->Length() = (VarSizeType)strlen(output_var->Contents());
		//}

		// Consider the above to have been always successful unless the below returns an error:
		return output_var->Close();  // In case it's the clipboard.
	}

	case ACT_STRINGSPLIT:
		return StringSplit(ARG1, ARG2, ARG3, ARG4);

	case ACT_SPLITPATH:
		return SplitPath(ARG1);

	case ACT_SORT:
		return PerformSort(ARG1, ARG2);

	case ACT_GETKEYSTATE:
		return GetKeyJoyState(ARG2, ARG3);

	case ACT_RANDOM:
	{
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		bool use_float = IsPureNumeric(ARG2, true, false, true) == PURE_FLOAT
			|| IsPureNumeric(ARG3, true, false, true) == PURE_FLOAT;
		if (use_float)
		{
			double rand_min = *ARG2 ? ATOF(ARG2) : 0;
			double rand_max = *ARG3 ? ATOF(ARG3) : INT_MAX;
			// Seems best not to use ErrorLevel for this command at all, since silly cases
			// such as Max > Min are too rare.  Swap the two values instead.
			if (rand_min > rand_max)
			{
				double rand_swap = rand_min;
				rand_min = rand_max;
				rand_max = rand_swap;
			}
			return output_var->Assign((genrand_real1() * (rand_max - rand_min)) + rand_min);
		}
		else // Avoid using floating point, where possible, which may improve speed a lot more than expected.
		{
			int rand_min = *ARG2 ? ATOI(ARG2) : 0;
			int rand_max = *ARG3 ? ATOI(ARG3) : INT_MAX;
			// Seems best not to use ErrorLevel for this command at all, since silly cases
			// such as Max > Min are too rare.  Swap the two values instead.
			if (rand_min > rand_max)
			{
				int rand_swap = rand_min;
				rand_min = rand_max;
				rand_max = rand_swap;
			}
			// Do NOT use genrand_real1() to generate random integers because of cases like
			// min=0 and max=1: we want an even distribution of 1's and 0's in that case, not
			// something skewed that might result due to rounding/truncation issues caused by
			// the float method used above:
			// AutoIt3: __int64 is needed here to do the proper conversion from unsigned long to signed long:
			return output_var->Assign(   (int)(__int64(genrand_int32()
				% ((__int64)rand_max - rand_min + 1)) + rand_min)   );
		}
	}

	case ACT_ASSIGN:
		// Note: This line's args have not yet been dereferenced in this case.  The below
		// function will handle that if it is needed.
		return PerformAssign();  // It will report any errors for us.

	case ACT_ASSIGNEXPR:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		return output_var->Assign(ARG2); // ARG2 now contains the evaluated result of the expression.

	case ACT_DRIVESPACEFREE:
		return DriveSpace(ARG2, true);

	case ACT_DRIVE:
		return Drive(THREE_ARGS);

	case ACT_DRIVEGET:
		return DriveGet(ARG2, ARG3);

	case ACT_SOUNDGET:
	case ACT_SOUNDSET:
		device_id = *ARG4 ? ATOI(ARG4) - 1 : 0;
		if (device_id < 0)
			device_id = 0;
		instance_number = 1;  // Set default.
		component_type = *ARG2 ? SoundConvertComponentType(ARG2, &instance_number) : MIXERLINE_COMPONENTTYPE_DST_SPEAKERS;
		return SoundSetGet(mActionType == ACT_SOUNDGET ? NULL : ARG1
			, component_type, instance_number  // Which instance of this component, 1 = first
			, *ARG3 ? SoundConvertControlType(ARG3) : MIXERCONTROL_CONTROLTYPE_VOLUME  // Default
			, (UINT)device_id);

	case ACT_SOUNDGETWAVEVOLUME:
	case ACT_SOUNDSETWAVEVOLUME:
		device_id = *ARG2 ? ATOI(ARG2) - 1 : 0;
		if (device_id < 0)
			device_id = 0;
		return (mActionType == ACT_SOUNDGETWAVEVOLUME) ? SoundGetWaveVolume((HWAVEOUT)device_id)
			: SoundSetWaveVolume(ARG1, (HWAVEOUT)device_id);

	case ACT_SOUNDPLAY:
		return SoundPlay(ARG1, *ARG2 && !stricmp(ARG2, "wait") || !stricmp(ARG2, "1"));

	case ACT_FILEAPPEND:
		// Uses the read-file loop's current item filename was explicitly leave blank (i.e. not just
		// a reference to a variable that's blank):
		return FileAppend(ARG2, ARG1, (mArgc < 2) ? aCurrentReadFile : NULL);

	case ACT_FILEREAD:
		return FileRead(ARG2);

	case ACT_FILEREADLINE:
		return FileReadLine(ARG2, ARG3);

	case ACT_FILEDELETE:
		return FileDelete(ARG1);

	case ACT_FILERECYCLE:
		return FileRecycle(ARG1);

	case ACT_FILERECYCLEEMPTY:
		return FileRecycleEmpty(ARG1);

	case ACT_FILEINSTALL:
		return FileInstall(THREE_ARGS);

	case ACT_FILECOPY:
	{
		int error_count = Util_CopyFile(ARG1, ARG2, *ARG3 == '1' && !*(ARG3 + 1), false);
		if (!error_count)
			return g_ErrorLevel->Assign(ERRORLEVEL_NONE);
		if (g_script.mIsAutoIt2)
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // For backward compatibility with v2.
		return g_ErrorLevel->Assign(error_count);
	}
	case ACT_FILEMOVE:
		return g_ErrorLevel->Assign(Util_CopyFile(ARG1, ARG2, *ARG3 == '1' && !*(ARG3 + 1), true));
	case ACT_FILECOPYDIR:
		return g_ErrorLevel->Assign(Util_CopyDir(ARG1, ARG2, *ARG3 == '1' && !*(ARG3 + 1)) ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);
	case ACT_FILEMOVEDIR:
		if (toupper(*ARG3) == 'R')
		{
			// Perform a simple rename instead, which prevents the operation from being only partially
			// complete if the source directory is in use (due to being a working dir for a currently
			// running process, or containing a file that is being written to).  In other words,
			// the operation will be "all or none":
			g_ErrorLevel->Assign(MoveFile(ARG1, ARG2) ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);
			return OK;
		}
		// Otherwise:
		return g_ErrorLevel->Assign(Util_MoveDir(ARG1, ARG2, *ARG3 == '1' && !*(ARG3 + 1)) ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);

	case ACT_FILECREATEDIR:
		return FileCreateDir(ARG1);
	case ACT_FILEREMOVEDIR:
		if (!*ARG1) // Consider an attempt to create or remove a blank dir to be an error.
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		return g_ErrorLevel->Assign(Util_RemoveDir(ARG1, *ARG2 == '1' && !*(ARG2 + 1)) ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);

	case ACT_FILEGETATTRIB:
		// The specified ARG, if non-blank, takes precedence over the file-loop's file (if any):
		#define USE_FILE_LOOP_FILE_IF_ARG_BLANK(arg) (*arg ? arg : (aCurrentFile ? aCurrentFile->cFileName : ""))
		return FileGetAttrib(USE_FILE_LOOP_FILE_IF_ARG_BLANK(ARG2));
	case ACT_FILESETATTRIB:
		FileSetAttrib(ARG1, USE_FILE_LOOP_FILE_IF_ARG_BLANK(ARG2), ConvertLoopMode(ARG3), !strcmp(ARG4, "1"));
		return OK; // Always return OK since ErrorLevel will indicate if there was a problem.
	case ACT_FILEGETTIME:
		return FileGetTime(USE_FILE_LOOP_FILE_IF_ARG_BLANK(ARG2), *ARG3);
	case ACT_FILESETTIME:
		FileSetTime(ARG1, USE_FILE_LOOP_FILE_IF_ARG_BLANK(ARG2), *ARG3, ConvertLoopMode(ARG4), !strcmp(ARG5, "1"));
		return OK; // Always return OK since ErrorLevel will indicate if there was a problem.
	case ACT_FILEGETSIZE:
		return FileGetSize(USE_FILE_LOOP_FILE_IF_ARG_BLANK(ARG2), ARG3);
	case ACT_FILEGETVERSION:
		return FileGetVersion(USE_FILE_LOOP_FILE_IF_ARG_BLANK(ARG2));

	case ACT_SETWORKINGDIR:
		if (SetCurrentDirectory(ARG1))
		{
			// Other than during program startup, this should be the only place where the official
			// working dir can change.  The exception is FileSelectFile(), which changes the working
			// dir as the user navigates from folder to folder.  However, the whole purpose of
			// maintaining g_WorkingDir is to workaround that very issue.
			// NOTE: GetCurrentDirectory() is called explicitly in case ARG1 is a relative path.
			// We want to store the absolute path:
			if (!GetCurrentDirectory(sizeof(g_WorkingDir), g_WorkingDir))
				strlcpy(g_WorkingDir, ARG1, sizeof(g_WorkingDir));
			return g_ErrorLevel->Assign(ERRORLEVEL_NONE);
		}
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);

	case ACT_FILESELECTFILE:
		return FileSelectFile(ARG2, ARG3, ARG4, ARG5);

	case ACT_FILESELECTFOLDER:
		return FileSelectFolder(ARG2, (DWORD)(*ARG3 ? ATOI(ARG3) : FSF_ALLOW_CREATE), ARG4);

	case ACT_FILEGETSHORTCUT:
		return FileGetShortcut(ARG1);
	case ACT_FILECREATESHORTCUT:
		return FileCreateShortcut(NINE_ARGS);

	// Like AutoIt2, if either output_var or ARG1 aren't purely numeric, they
	// will be considered to be zero for all of the below math functions:
	case ACT_ADD:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;

		#undef DETERMINE_NUMERIC_TYPES
		#define DETERMINE_NUMERIC_TYPES \
			value_is_pure_numeric = IsPureNumeric(ARG2, true, false, true, true);\
			var_is_pure_numeric = IsPureNumeric(output_var->Contents(), true, false, true, true);

		// Some performance can be gained by relying on the fact that short-circuit boolean
		// can skip the "var_is_pure_numeric" check whenever value_is_pure_numeric == PURE_FLOAT.
		// This is because var_is_pure_numeric is never directly needed here (unlike EvaluateCondition()).
		// However, benchmarks show that this makes such a small difference that it's not worth the
		// loss of maintainability and the slightly larger code size due to macro expansion:
		//#undef IF_EITHER_IS_FLOAT
		//#define IF_EITHER_IS_FLOAT if (value_is_pure_numeric == PURE_FLOAT \
		//	|| IsPureNumeric(output_var->Contents(), true, false, true, true) == PURE_FLOAT)

		DETERMINE_NUMERIC_TYPES

		if (*ARG3 && strchr("SMHD", toupper(*ARG3))) // the command is being used to add a value to a date-time.
		{
			if (!value_is_pure_numeric) // It's considered to be zero, so the output_var is left unchanged:
				return OK;
			else
			{
				// Use double to support a floating point value for days, hours, minutes, etc:
				double nUnits = ATOF(ARG2);  // ATOF() returns a double, at least on MSVC++ 7.x
				FILETIME ft, ftNowUTC;
				if (*output_var->Contents())
				{
					if (!YYYYMMDDToFileTime(output_var->Contents(), ft))
						return output_var->Assign(""); // Set to blank to indicate the problem.
				}
				else // The output variable is currently blank, so substitute the current time for it.
				{
					GetSystemTimeAsFileTime(&ftNowUTC);
					FileTimeToLocalFileTime(&ftNowUTC, &ft);  // Convert UTC to local time.
				}
				// Convert to 10ths of a microsecond (the units of the FILETIME struct):
				switch (toupper(*ARG3))
				{
				case 'S': // Seconds
					nUnits *= (double)10000000;
					break;
				case 'M': // Minutes
					nUnits *= ((double)10000000 * 60);
					break;
				case 'H': // Hours
					nUnits *= ((double)10000000 * 60 * 60);
					break;
				case 'D': // Days
					nUnits *= ((double)10000000 * 60 * 60 * 24);
					break;
				}
				// Convert ft struct to a 64-bit variable (maybe there's some way to avoid these conversions):
				ULARGE_INTEGER ul;
				ul.LowPart = ft.dwLowDateTime;
				ul.HighPart = ft.dwHighDateTime;
				// Add the specified amount of time to the result value:
				ul.QuadPart += (__int64)nUnits;  // Seems ok to cast/truncate in light of the *=10000000 above.
				// Convert back into ft struct:
				ft.dwLowDateTime = ul.LowPart;
				ft.dwHighDateTime = ul.HighPart;
				FileTimeToYYYYMMDD(buf_temp, ft, false);
				return output_var->Assign(buf_temp);
			}
		}
		else // The command is being used to do normal math (not date-time).
		{
			IF_EITHER_IS_FLOAT
				return output_var->Assign(ATOF(output_var->Contents()) + ATOF(ARG2));  // Overload: Assigns a double.
			else // Non-numeric variables or values are considered to be zero for the purpose of the calculation.
				return output_var->Assign(ATOI64(output_var->Contents()) + ATOI64(ARG2));  // Overload: Assigns an int.
		}
		return OK;  // Never executed.
	case ACT_SUB:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;

		if (*ARG3 && strchr("SMHD", toupper(*ARG3))) // the command is being used to subtract date-time values.
		{
			bool failed;
			// If either ARG2 or output_var->Contents() is blank, it will default to the current time:
			__int64 time_until = YYYYMMDDSecondsUntil(ARG2, output_var->Contents(), failed);
			if (failed) // Usually caused by an invalid component in the date-time string.
				return output_var->Assign("");
			switch (toupper(*ARG3))
			{
			// Do nothing in the case of 'S' (seconds).  Otherwise:
			case 'M': time_until /= 60; break; // Minutes
			case 'H': time_until /= 60 * 60; break; // Hours
			case 'D': time_until /= 60 * 60 * 24; break; // Days
			}
			// Only now that any division has been performed (to reduce the magnitude of
			// time_until) do we cast down into an int, which is the standard size
			// used for non-float results (the result is always non-float for subtraction
			// of two date-times):
			return output_var->Assign(time_until); // Assign as signed 64-bit.
		}
		else
		{
			DETERMINE_NUMERIC_TYPES
			IF_EITHER_IS_FLOAT
				return output_var->Assign(ATOF(output_var->Contents()) - ATOF(ARG2));  // Overload: Assigns a double.
			else // Non-numeric variables or values are considered to be zero for the purpose of the calculation.
				return output_var->Assign(ATOI64(output_var->Contents()) - ATOI64(ARG2));  // Overload: Assigns an INT.
			break;
		}

		// If above didn't return, buf_temp now has the value to store:
		return output_var->Assign(buf_temp);

	case ACT_MULT:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_FLOAT
			return output_var->Assign(ATOF(output_var->Contents()) * ATOF(ARG2));  // Overload: Assigns a double.
		else // Non-numeric variables or values are considered to be zero for the purpose of the calculation.
			return output_var->Assign(ATOI64(output_var->Contents()) * ATOI64(ARG2));  // Overload: Assigns an INT.

	case ACT_DIV:
	{
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_FLOAT
		{
			double ARG2_as_float = ATOF(ARG2);  // Since ATOF() returns double, at least on MSVC++ 7.x
			// It's a little iffy to compare floats this way, but what's the alternative?:
			if (ARG2_as_float == (double)0.0)
				return LineError("This line would attempt to divide by zero." ERR_ABORT, FAIL, ARG2);
			return output_var->Assign(ATOF(output_var->Contents()) / ARG2_as_float);  // Overload: Assigns a double.
		}
		else // Non-numeric variables or values are considered to be zero for the purpose of the calculation.
		{
			__int64 ARG2_as_int = ATOI64(ARG2);
			if (!ARG2_as_int)
				return LineError("This line would attempt to divide by zero." ERR_ABORT, FAIL, ARG2);
			return output_var->Assign(ATOI64(output_var->Contents()) / ARG2_as_int);  // Overload: Assigns an INT.
		}
	}

	case ACT_KEYHISTORY:
#ifdef ENABLE_KEY_HISTORY_FILE
		if (*ARG1 || *ARG2)
		{
			switch (ConvertOnOffToggle(ARG1))
			{
			case NEUTRAL:
			case TOGGLE:
				g_KeyHistoryToFile = !g_KeyHistoryToFile;
				if (!g_KeyHistoryToFile)
					KeyHistoryToFile();  // Signal it to close the file, if it's open.
				break;
			case TOGGLED_ON:
				g_KeyHistoryToFile = true;
				break;
			case TOGGLED_OFF:
				g_KeyHistoryToFile = false;
				KeyHistoryToFile();  // Signal it to close the file, if it's open.
				break;
			// We know it's a variable because otherwise the loading validation would have caught it earlier:
			case TOGGLE_INVALID:
				return LineError("The variable in param #1 does not resolve to an allowed value.", FAIL, ARG1);
			}
			if (*ARG2) // The user also specified a filename, so update the target filename.
				KeyHistoryToFile(ARG2);
			return OK;
		}
#endif
		// Otherwise:
		return ShowMainWindow(MAIN_MODE_KEYHISTORY, false); // Pass "unrestricted" when the command is explicitly used in the script.
	case ACT_LISTLINES:
		return ShowMainWindow(MAIN_MODE_LINES, false); // Pass "unrestricted" when the command is explicitly used in the script.
	case ACT_LISTVARS:
		return ShowMainWindow(MAIN_MODE_VARS, false); // Pass "unrestricted" when the command is explicitly used in the script.
	case ACT_LISTHOTKEYS:
		return ShowMainWindow(MAIN_MODE_HOTKEYS, false); // Pass "unrestricted" when the command is explicitly used in the script.

	case ACT_MSGBOX:
	{
		int result;
		// If the MsgBox window can't be displayed for any reason, always return FAIL to
		// the caller because it would be unsafe to proceed with the execution of the
		// current script subroutine.  For example, if the script contains an IfMsgBox after,
		// this line, it's result would be unpredictable and might cause the subroutine to perform
		// the opposite action from what was intended (e.g. Delete vs. don't delete a file).
		if (!mArgc) // When called explicitly with zero params, it displays this default msg.
			result = MsgBox("Press OK to continue.");
		else if (mArgc == 1) // In the special 1-parameter mode, the first param is the prompt.
			result = MsgBox(ARG1);
		else
			result = MsgBox(ARG3, ATOI(ARG1), ARG2, ATOF(ARG4));
		// Above allows backward compatibility with AutoIt2's param ordering while still
		// permitting the new method of allowing only a single param.
		if (!result)
			// It will fail if the text is too large (say, over 150K or so on XP), but that
			// has since been fixed by limiting how much it tries to display.
			// If there were too many message boxes displayed, it will already have notified
			// the user of this via a final MessageBox dialog, so our call here will
			// not have any effect.  The below only takes effect if MsgBox()'s call to
			// MessageBox() failed in some unexpected way:
			LineError("The MsgBox dialog could not be displayed." ERR_ABORT);
		return result ? OK : FAIL;
	}

	case ACT_INPUTBOX:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		return InputBox(output_var, ARG2, ARG3, toupper(*ARG4) == 'H' // 4th is whether to hide input.
			, *ARG5 ? ATOI(ARG5) : INPUTBOX_DEFAULT  // Width
			, *ARG6 ? ATOI(ARG6) : INPUTBOX_DEFAULT  // Height
			, *ARG7 ? ATOI(ARG7) : INPUTBOX_DEFAULT  // Xpos
			, *ARG8 ? ATOI(ARG8) : INPUTBOX_DEFAULT  // Ypos
			// ARG9: future use for Font name & size, e.g. "Courier:8"
			, ATOF(ARG10)  // Timeout
			, ARG11  // Initial default string for the edit field.
			);

	case ACT_SPLASHTEXTON:
	{
		// The SplashText command has been adapted from the AutoIt3 source.
		int W = *ARG1 ? ATOI(ARG1) : 200;
		int H = *ARG2 ? ATOI(ARG2) : 0;

		// Add some caption and frame size to window:
		W += GetSystemMetrics(SM_CXFIXEDFRAME) * 2;
		int min_height = GetSystemMetrics(SM_CYCAPTION) + (GetSystemMetrics(SM_CXFIXEDFRAME) * 2);
		if (g_script.mIsAutoIt2)
		{
			// I think this is probably something like how AutoIt2 does things:
			if (H < min_height)
				H = min_height;
		}
		else // Use the new, AutoIt3 method.
			H += min_height;

		POINT pt = CenterWindow(W, H);  // Determine how to center the window in the region that excludes the task bar.

		// My: Probably not to much overhead to do this, though it probably would
		// perform better to resize and "re-text" the existing window rather than
		// recreating it:
		#define DESTROY_SPLASH \
		{\
			if (g_hWndSplash && IsWindow(g_hWndSplash))\
				DestroyWindow(g_hWndSplash);\
			g_hWndSplash = NULL;\
		}
		DESTROY_SPLASH

		// Doesn't seem necessary to have it owned by the main window, but neither
		// does doing so seem to cause any harm.  Feels safer to have it be
		// an independent window.  Update: Must make it owned by the parent window
		// otherwise it will get its own task-bar icon, which is usually undesirable.
		// In addition, making it an owned window should automatically cause it to be
		// destroyed when it's parent window is destroyed:
		g_hWndSplash = CreateWindowEx(WS_EX_TOPMOST, WINDOW_CLASS_SPLASH, ARG3  // ARG3 is the window title
			, WS_DISABLED|WS_POPUP|WS_CAPTION, pt.x, pt.y, W, H, g_hWnd, (HMENU)NULL, g_hInstance, NULL);

		RECT rect;
		GetClientRect(g_hWndSplash, &rect);	// get the client size

		// CREATE static label full size of client area
		HWND static_win = CreateWindowEx(0, "static", ARG4 // ARG4 is the window's text
			, WS_CHILD|WS_VISIBLE|SS_CENTER
			, 0, 0, rect.right - rect.left, rect.bottom - rect.top
			, g_hWndSplash, (HMENU)NULL, g_hInstance, NULL);

		if (!g_hFontSplash)
		{
			char default_font_name[65];
			int CyPixels, nSize = 12, nWeight = FW_NORMAL;
			HDC hdc = CreateDC("DISPLAY", NULL, NULL, NULL);
			SelectObject(hdc, (HFONT)GetStockObject(DEFAULT_GUI_FONT));		// Get Default Font Name
			GetTextFace(hdc, sizeof(default_font_name) - 1, default_font_name); // -1 just in case, like AutoIt3.
			CyPixels = GetDeviceCaps(hdc, LOGPIXELSY);			// For Some Font Size Math
			DeleteDC(hdc);
			//strcpy(default_font_name,vParams[7].szValue());	// Font Name
			//nSize = vParams[8].nValue();		// Font Size
			//if ( vParams[9].nValue() >= 0 && vParams[9].nValue() <= 1000 )
			//	nWeight = vParams[9].nValue();			// Font Weight
			g_hFontSplash = CreateFont(0-(nSize*CyPixels)/72,0,0,0,nWeight,0,0,0,DEFAULT_CHARSET,
				OUT_TT_PRECIS,CLIP_DEFAULT_PRECIS,PROOF_QUALITY,FF_DONTCARE,default_font_name);	// Create Font
			// The font is deleted when by g_script's destructor.
		}

		SendMessage(static_win, WM_SETFONT, (WPARAM)g_hFontSplash, MAKELPARAM(TRUE, 0));	// Do Font
		ShowWindow(g_hWndSplash, SW_SHOWNOACTIVATE);				// Show the Splash
		// Doesn't help with the brief delay in updating the window that happens when
		// something like URLDownloadToFile is used immediately after SplashTextOn:
		//InvalidateRect(g_hWndSplash, NULL, TRUE);
		// But this does, but for now it seems unnecessary since the user can always do
		// a manual sleep in the extremely rare cases this ever happens (even when it does
		// happen, the window updates eventually, after the download starts, at least on
		// my system.  Update: Might as well do it since it's a little nicer this way
		// (the text appears more quickly when the command after the splash is something
		// that might keep our thread tied up and unable to check messages).
		SLEEP_WITHOUT_INTERRUPTION(-1)
		return OK;
	}
	case ACT_SPLASHTEXTOFF:
		DESTROY_SPLASH
		return OK;

	case ACT_PROGRESS:
		return Splash(FIVE_ARGS, "", false);  // ARG6 is for future use and currently not passed.

	case ACT_SPLASHIMAGE:
		return Splash(ARG2, ARG3, ARG4, ARG5, ARG6, ARG1, true);  // ARG7 is for future use and currently not passed.

	case ACT_TOOLTIP:
		return ToolTip(FOUR_ARGS);

	case ACT_TRAYTIP:
		return TrayTip(FOUR_ARGS);

	case ACT_INPUT:
		return Input(ARG2, ARG3, ARG4);

	case ACT_SEND:
	case ACT_SENDRAW:
		SendKeys(ARG1, mActionType == ACT_SENDRAW);
		return OK;


	// Macros used by the Mouse commands.  These macros are executed here rather than inside the various
	// MouseXXX() functions because some of those functions call the others.  Notes:
	// Turn it on unconditionally even if it was on, since Ctrl-Alt-Del might have disabled it.
	// Turn it back off only if it wasn't ON before we started.
	#define MOUSE_BLOCKINPUT_ON \
		if (do_selective_blockinput = (g_BlockInputMode == TOGGLE_MOUSE || g_BlockInputMode == TOGGLE_SENDANDMOUSE) \
			&& g_os.IsWinNT4orLater())\
		{\
			blockinput_prev = g_BlockInput;\
			ScriptBlockInput(true);\
		}
	#define MOUSE_BLOCKINPUT_OFF \
	if (do_selective_blockinput && !blockinput_prev)\
		ScriptBlockInput(false);


	case ACT_MOUSECLICKDRAG:
		if (   !(vk = ConvertMouseButton(ARG1, false))   )
			return LineError(ERR_MOUSE_BUTTON ERR_ABORT, FAIL, ARG1);
		if (!ValidateMouseCoords(ARG2, ARG3))
			return LineError(ERR_MOUSE_COORD ERR_ABORT, FAIL, ARG2);
		if (!ValidateMouseCoords(ARG4, ARG5))
			return LineError(ERR_MOUSE_COORD ERR_ABORT, FAIL, ARG4);
		// If no starting coords are specified, we tell the function to start at the
		// current mouse position:
		x = *ARG2 ? ATOI(ARG2) : COORD_UNSPECIFIED;
		y = *ARG3 ? ATOI(ARG3) : COORD_UNSPECIFIED;
		MOUSE_BLOCKINPUT_ON
		MouseClickDrag(vk, x, y, ATOI(ARG4), ATOI(ARG5), *ARG6 ? ATOI(ARG6) : g.DefaultMouseSpeed
			, toupper(*ARG7) == 'R');
		MOUSE_BLOCKINPUT_OFF
		return OK;

	case ACT_MOUSECLICK:
		if (   !(vk = ConvertMouseButton(ARG1))   )
			return LineError(ERR_MOUSE_BUTTON ERR_ABORT, FAIL, ARG1);
		if (!ValidateMouseCoords(ARG2, ARG3))
			return LineError(ERR_MOUSE_COORD ERR_ABORT, FAIL, ARG2);
		x = *ARG2 ? ATOI(ARG2) : COORD_UNSPECIFIED;
		y = *ARG3 ? ATOI(ARG3) : COORD_UNSPECIFIED;
		KeyEventTypes event_type;
		switch(*ARG6)
		{
		case 'u':
		case 'U':
			event_type = KEYUP;
			break;
		case 'd':
		case 'D':
			event_type = KEYDOWN;
			break;
		default:
			event_type = KEYDOWNANDUP;
		}
		MOUSE_BLOCKINPUT_ON
		MouseClick(vk, x, y, *ARG4 ? ATOI(ARG4) : 1, *ARG5 ? ATOI(ARG5) : g.DefaultMouseSpeed, event_type
			, toupper(*ARG7) == 'R');
		MOUSE_BLOCKINPUT_OFF
		return OK;

	case ACT_MOUSEMOVE:
		if (!ValidateMouseCoords(ARG1, ARG2))
			return LineError(ERR_MOUSE_COORD ERR_ABORT, FAIL, ARG1);
		x = *ARG1 ? ATOI(ARG1) : COORD_UNSPECIFIED;
		y = *ARG2 ? ATOI(ARG2) : COORD_UNSPECIFIED;
		MOUSE_BLOCKINPUT_ON
		MouseMove(x, y, *ARG3 ? ATOI(ARG3) : g.DefaultMouseSpeed, toupper(*ARG4) == 'R');
		MOUSE_BLOCKINPUT_OFF
		return OK;

	case ACT_MOUSEGETPOS:
		return MouseGetPos(ATOI(ARG5) == 1);

//////////////////////////////////////////////////////////////////////////

	case ACT_COORDMODE:
	{
		bool screen_mode;
		if (!*ARG2 || !stricmp(ARG2, "Screen"))
			screen_mode = true;
		else if (!stricmp(ARG2, "Relative"))
			screen_mode = false;
		else  // Since validated at load-time, too rare to return FAIL for.
			return OK;
		CoordModeAttribType attrib = ConvertCoordModeAttrib(ARG1);
		if (attrib)
		{
			if (screen_mode)
				g.CoordMode |= attrib;
			else
				g.CoordMode &= ~attrib;
		}
		//else too rare to report an error, since load-time validation normally catches it.
		return OK;
	}

	case ACT_SETDEFAULTMOUSESPEED:
		g.DefaultMouseSpeed = (UCHAR)ATOI(ARG1);
		// In case it was a deref, force it to be some default value if it's out of range:
		if (g.DefaultMouseSpeed < 0 || g.DefaultMouseSpeed > MAX_MOUSE_SPEED)
			g.DefaultMouseSpeed = DEFAULT_MOUSE_SPEED;
		return OK;

	case ACT_SETTITLEMATCHMODE:
		switch (ConvertTitleMatchMode(ARG1))
		{
		case FIND_IN_LEADING_PART: g.TitleMatchMode = FIND_IN_LEADING_PART; return OK;
		case FIND_ANYWHERE: g.TitleMatchMode = FIND_ANYWHERE; return OK;
		case FIND_EXACT: g.TitleMatchMode = FIND_EXACT; return OK;
		case FIND_FAST: g.TitleFindFast = true; return OK;
		case FIND_SLOW: g.TitleFindFast = false; return OK;
		}
		return LineError(ERR_TITLEMATCHMODE2, FAIL, ARG1);

	case ACT_SETFORMAT:
		// For now, it doesn't seem necessary to have runtime validation of the first parameter.
		// Just ignore the command if it's not valid:
		if (!stricmp(ARG1, "float"))
		{
			// -2 to allow room for the letter 'f' and the '%' that will be added:
			if (strlen(ARG2) >= sizeof(g.FormatFloat) - 2) // A variable that resolved to something too long.
				return OK; // Seems best not to bother with a runtime error for something so rare.
			// Make sure the formatted string wouldn't exceed the buffer size:
			__int64 width = ATOI64(ARG2);
			char *dot_pos = strchr(ARG2, '.');
			__int64 precision = dot_pos ? ATOI64(dot_pos + 1) : 0;
			if (width + precision + 2 > MAX_FORMATTED_NUMBER_LENGTH) // +2 to allow room for decimal point itself and leading minus sign.
				return OK; // Don't change it.
			// Create as "%ARG2f".  Add a dot if none was specified so that "0" is the same as "0.", which
			// seems like the most user-friendly approach; it's also easier to document in the help file.
			// Note that %f can handle doubles in MSVC++:
			sprintf(g.FormatFloat, "%%%s%sf", ARG2, dot_pos ? "" : ".");
		}
		else if (!stricmp(ARG1, "integer"))
		{
			switch(*ARG2)
			{
			case 'd':
			case 'D':
				g.FormatIntAsHex = false;
				break;
			case 'h':
			case 'H':
				g.FormatIntAsHex = true;
				break;
			// Otherwise, since the first letter isn't recongized, do nothing since 99% of the time such a
			// probably would be caught at load-time.
			}
		}
		// Otherwise, ignore invalid type at runtime since 99% of the time it would be caught at load-time:
		return OK;

	case ACT_FORMATTIME:
		return FormatTime(ARG2, ARG3);

	case ACT_MENU:
		return g_script.PerformMenu(FIVE_ARGS);

	case ACT_GUI:
		return g_script.PerformGui(FOUR_ARGS);

	case ACT_GUICONTROL:
		return GuiControl(THREE_ARGS);

	case ACT_GUICONTROLGET:
		return GuiControlGet(ARG2, ARG3, ARG4);

	case ACT_SETCONTROLDELAY: g.ControlDelay = ATOI(ARG1); return OK;
	case ACT_SETWINDELAY: g.WinDelay = ATOI(ARG1); return OK;
	case ACT_SETMOUSEDELAY: g.MouseDelay = ATOI(ARG1); return OK;
	case ACT_SETKEYDELAY:
		if (*ARG1)
			g.KeyDelay = ATOI(ARG1);
		if (*ARG2)
			g.PressDuration = ATOI(ARG2);
		return OK;

	case ACT_SETBATCHLINES:
		// This below ensures that IntervalBeforeRest and LinesPerCycle aren't both in effect simultaneously
		// (i.e. that both aren't greater than -1), even though ExecUntil() has code to prevent a double-sleep
		// even if that were to happen.
		if (strcasestr(ARG1, "ms")) // This detection isn't perfect, but it doesn't seem necessary to be too demanding.
		{
			g.LinesPerCycle = -1;  // Disable the old BatchLines method in favor of the new one below.
			g.IntervalBeforeRest = ATOI(ARG1);  // If negative, script never rests.  If 0, it rests after every line.
		}
		else
		{
			g.IntervalBeforeRest = -1;  // Disable the new method in favor of the old one below:
			// This value is signed 64-bits to support variable reference (i.e. containing a large int)
			// the user might throw at it:
			if (   !(g.LinesPerCycle = ATOI64(ARG1))   )
				// Don't interpret zero as "infinite" because zero can accidentally
				// occur if the dereferenced var was blank:
				g.LinesPerCycle = 10;  // The old default, which is retained for compatbility with existing scripts.
		}
		return OK;

	////////////////////////////////////////////////////////////////////////////////////////
	// For these, it seems best not to report an error during runtime if there's
	// an invalid value (e.g. something other than On/Off/Blank) in a param containing
	// a dereferenced variable, since settings are global and affect all subroutines,
	// not just the one that we would otherwise report failure for:
	case ACT_SETSTORECAPSLOCKMODE:
		if (   (toggle = ConvertOnOff(ARG1, NEUTRAL)) != NEUTRAL   )
			g.StoreCapslockMode = (toggle == TOGGLED_ON);
		return OK;
	case ACT_SUSPEND:
		switch (ConvertOnOffTogglePermit(ARG1))
		{
		case NEUTRAL:
		case TOGGLE:
			ToggleSuspendState();
			break;
		case TOGGLED_ON:
			if (!g_IsSuspended)
				ToggleSuspendState();
			break;
		case TOGGLED_OFF:
			if (g_IsSuspended)
				ToggleSuspendState();
			break;
		case TOGGLE_PERMIT:
			// In this case do nothing.  The user is just using this command as a flag to indicate that
			// this subroutine should not be suspended.
			break;
		// We know it's a variable because otherwise the loading validation would have caught it earlier:
		case TOGGLE_INVALID:
			return LineError("The variable in param #1 does not resolve to an allowed value.", FAIL, ARG1);
		}
		return OK;
	case ACT_PAUSE:
		return ChangePauseState(ConvertOnOffToggle(ARG1));
	case ACT_AUTOTRIM:
		if (   (toggle = ConvertOnOff(ARG1, NEUTRAL)) != NEUTRAL   )
			g.AutoTrim = (toggle == TOGGLED_ON);
		return OK;
	case ACT_STRINGCASESENSE:
		if (   (toggle = ConvertOnOff(ARG1, NEUTRAL)) != NEUTRAL   )
			g.StringCaseSense = (toggle == TOGGLED_ON);
		return OK;
	case ACT_DETECTHIDDENWINDOWS:
		if (   (toggle = ConvertOnOff(ARG1, NEUTRAL)) != NEUTRAL   )
			g.DetectHiddenWindows = (toggle == TOGGLED_ON);
		return OK;
	case ACT_DETECTHIDDENTEXT:
		if (   (toggle = ConvertOnOff(ARG1, NEUTRAL)) != NEUTRAL   )
			g.DetectHiddenText = (toggle == TOGGLED_ON);
		return OK;
	case ACT_BLOCKINPUT:
		switch (toggle = ConvertBlockInput(ARG1))
		{
		case TOGGLED_ON:
			ScriptBlockInput(true);
			break;
		case TOGGLED_OFF:
			ScriptBlockInput(false);
			break;
		case TOGGLE_SEND:
		case TOGGLE_MOUSE:
		case TOGGLE_SENDANDMOUSE:
		case TOGGLE_DEFAULT:
			g_BlockInputMode = toggle;
			break;
		// default (NEUTRAL or TOGGLE_INVALID): do nothing.
		}
		return OK;

	////////////////////////////////////////////////////////////////////////////////////////
	// For these, it seems best not to report an error during runtime if there's
	// an invalid value (e.g. something other than On/Off/Blank) in a param containing
	// a dereferenced variable, since settings are global and affect all subroutines,
	// not just the one that we would otherwise report failure for:
	case ACT_SETNUMLOCKSTATE:
		return SetToggleState(VK_NUMLOCK, g_ForceNumLock, ARG1);
	case ACT_SETCAPSLOCKSTATE:
		return SetToggleState(VK_CAPITAL, g_ForceCapsLock, ARG1);
	case ACT_SETSCROLLLOCKSTATE:
		return SetToggleState(VK_SCROLL, g_ForceScrollLock, ARG1);

	case ACT_EDIT:
		g_script.Edit();
		return OK;
	case ACT_RELOAD:
		g_script.Reload(true);
		// Even if the reload failed, it seems best to return OK anyway.  That way,
		// the script can take some follow-on action, e.g. it can sleep for 1000
		// after issuing the reload command and then take action under the assumption
		// that the reload didn't work (since obviously if the process and thread
		// in which the Sleep is running still exist, it didn't work):
		return OK;
	}

	// Since above didn't return, this line's mActionType isn't handled here,
	// so caller called it wrong.  ACT_INVALID should be impossible because
	// Script::AddLine() forbids it.

#ifdef _DEBUG
	return LineError("Perform(): Unhandled action type." ERR_ABORT);
#else
	return FAIL;
#endif
}



ResultType Line::ExpandArgs()
{
	// Make two passes through this line's arg list.  This is done because the performance of
	// realloc() is worse than doing a free() and malloc() because the former does a memcpy()
	// in addition to the latter's steps.  In addition, realloc() as much as doubles the memory
	// load on the system during the brief time that both the old and the new blocks of memory exist.
	// First pass: determine how much space will be needed to do all the args and allocate
	// more memory if needed.  Second pass: dereference the args into the buffer.

	size_t space_needed = GetExpandedArgSize(true);  // First pass. It takes into account the same things as 2nd pass.
	if (space_needed == VARSIZE_ERROR)
		return FAIL;  // It will have already displayed the error.

	if (space_needed > g_MaxVarCapacity)
		// Dereferencing the variables in this line's parameters would exceed the allowed size of the temp buffer:
		return LineError(ERR_MEM_LIMIT_REACHED);

	// Only allocate the buf at the last possible moment,
	// when it's sure the buffer will be used (improves performance when only a short
	// script with no derefs is being run):
	if (space_needed > sDerefBufSize)
	{
		// K&R: Integer division truncates the fractional part:
		size_t increments_needed = space_needed / DEREF_BUF_EXPAND_INCREMENT;
		if (space_needed % DEREF_BUF_EXPAND_INCREMENT)  // Need one more if above division truncated it.
			++increments_needed;
		sDerefBufSize = increments_needed * DEREF_BUF_EXPAND_INCREMENT;
		if (sDerefBuf)
			// Do a free() and malloc(), which should be far more efficient than realloc(),
			// especially if there is a large amount of memory involved here:
			free(sDerefBuf);
		if (!(sDerefBufMarker = sDerefBuf = (char *)malloc(sDerefBufSize)))
		{
			// Error msg was formerly: "Ran out of memory while attempting to dereference this line's parameters."
			sDerefBufSize = 0;  // Reset so that it can attempt a smaller amount next time.
			return LineError(ERR_OUTOFMEM ERR_ABORT); // Short msg since so rare.
		}
	}
	else
		sDerefBufMarker = sDerefBuf;  // Prior contents of buffer will be overwritten in any case.
	// Always init sDerefBufMarker even if zero iterations, because we want to enforce
	// the fact that its prior contents become invalid once this function is called.
	// It's also necessary due to the fact that all the old memory is discarded by
	// the above if more space was needed to accommodate this line:

	// (Re)set the timer unconditionally so that it starts counting again from time zero.
	// In other words, we only want the timer to fire when the large deref buffer has been
	// unused/idle for a straight 10 seconds.
	if (sDerefBufSize > 4*1024*1024)
		SET_DEREF_TIMER(10000)

	Var *the_only_var_of_this_arg;
	for (int iArg = 0; iArg < mArgc && iArg < MAX_ARGS; ++iArg) // Second pass.  For each arg:
	{
		ArgStruct &this_arg = mArg[iArg]; // For performance and convenience.

		// Load-time routines have already ensured that an arg can be an expression only if
		// it's not an input or output var.
		if (this_arg.is_expression)
		{
			sArgDeref[iArg] = sDerefBufMarker; // Point it to its location in the buffer.
			sDerefBufMarker = ExpandExpression(sDerefBufMarker, iArg); // Expand the expression into that location.
			continue;
		}

		if (this_arg.type == ARG_TYPE_OUTPUT_VAR)  // Don't bother wasting the mem to deref output var.
		{
			// In case its "dereferenced" contents are ever directly examined, set it to be
			// the empty string.  This also allows the ARG to be passed a dummy param, which
			// makes things more convenient and maintainable in other places:
			sArgDeref[iArg] = "";
			continue;
		}

		the_only_var_of_this_arg = NULL;
		if (this_arg.type == ARG_TYPE_INPUT_VAR)
			if (   !(the_only_var_of_this_arg = ResolveVarOfArg(iArg, false))   )
				return FAIL;  // The above will have already displayed the error.

		if (!the_only_var_of_this_arg) // Arg isn't an input var.
		{
			#define NO_DEREF (!ArgHasDeref(iArg + 1))
			if (NO_DEREF)
			{
				sArgDeref[iArg] = this_arg.text;  // Point the dereferenced arg to the arg text itself.
				continue;  // Don't need to use the deref buffer in this case.
			}
			// Now we know it has at least one deref.  If the second deref's marker is NULL,
			// the first is the only deref in this arg:
			#define SINGLE_ISOLATED_DEREF (!this_arg.deref[1].marker\
				&& this_arg.deref[0].length == strlen(this_arg.text)) // and the arg contains no literal text
			if (SINGLE_ISOLATED_DEREF)
				the_only_var_of_this_arg = this_arg.deref[0].var;
		}

		if (the_only_var_of_this_arg) // This arg resolves to only a single, naked var.
		{
			switch(ArgMustBeDereferenced(the_only_var_of_this_arg, iArg))
			{
			case CONDITION_FALSE:
				// This arg contains only a single dereference variable, and no
				// other text at all.  So rather than copy the contents into the
				// temp buffer, it's much better for performance (especially for
				// potentially huge variables like %clipboard%) to simply set
				// the pointer to be the variable itself.  However, this can only
				// be done if the var is the clipboard or a normal var of non-zero
				// length (since zero-length normal vars need to be fetched via
				// GetEnvironment()).  Update: Changed it so that it will deref
				// the clipboard if it contains only files and no text, so that
				// the files will be transcribed into the deref buffer.  This is
				// because the clipboard object needs a memory area into which
				// to write the filespecs it translated:
				sArgDeref[iArg] = the_only_var_of_this_arg->Contents();
				break;
			case CONDITION_TRUE:
				// the_only_var_of_this_arg is either a reserved var or a normal var of
				// zero length (for which GetEnvironment() is called for), or is used
				// again in this line as an output variable.  In all these cases, it must
				// be expanded into the buffer rather than accessed directly:
				sArgDeref[iArg] = sDerefBufMarker; // Point it to its location in the buffer.
				sDerefBufMarker += the_only_var_of_this_arg->Get(sDerefBufMarker) + 1; // +1 for terminator.
				break;
			default: // FAIL should be the only other possibility.
				return FAIL;  // ArgMustBeDereferenced() will already have displayed the error.
			}
		}
		else // The arg must be expanded in the normal, lower-performance way.
		{
			sArgDeref[iArg] = sDerefBufMarker; // Point it to its location in the buffer.
			sDerefBufMarker = ExpandArg(sDerefBufMarker, iArg); // Expand the arg into that location.
			if (!sDerefBufMarker)
				return FAIL;  // ExpandArg() will have already displayed the error.
		}
	}

	return OK;
}

	

inline VarSizeType Line::GetExpandedArgSize(bool aCalcDerefBufSize)
// Args that are expressions are only calculated correctly if aCalcDerefBufSize is true,
// which is okay for the moment since the only caller that can have expressions does call
// it that way.
// Returns the size, or VARSIZE_ERROR if there was a problem.
// WARNING: This function can return a size larger than what winds up actually being needed
// (e.g. caused by ScriptGetCursor()), so our callers should be aware that that can happen.
{
	int iArg;
	VarSizeType space_needed, space;
	DerefType *deref;
	Var *the_only_var_of_this_arg;
	bool include_this_arg;
	ResultType result;

	// Note: the below loop is similar to the one in ExpandArgs(), so the two should be
	// maintained together:
	for (iArg = 0, space_needed = 0; iArg < mArgc && iArg < MAX_ARGS; ++iArg) // For each arg:
	{
		ArgStruct &this_arg = mArg[iArg]; // For performance and convenience.

		// If this_arg.is_expression is true, the space is still calculated as though the
		// expression itself will be inside the arg.  This is done so that an expression
		// such as if(Array%i% = LargeString) can be expanded temporarily into the deref
		// buffer so that it can be evaluated more easily.

		// Accumulate the total of how much space we will need.
		if (this_arg.type == ARG_TYPE_OUTPUT_VAR)  // These should never be included in the space calculation.
			continue;

		// Always do this check before attempting to traverse the list of dereferences, since
		// such an attempt would be invalid in this case:
		the_only_var_of_this_arg = NULL;
		if (this_arg.type == ARG_TYPE_INPUT_VAR)
			if (   !(the_only_var_of_this_arg = ResolveVarOfArg(iArg, false))   )
				return VARSIZE_ERROR;  // The above will have already displayed the error.

		if (!the_only_var_of_this_arg) // It's not an input var.
		{
			if (NO_DEREF)
			{
				// Below relies on the fact that caller has ensure no args are expressions
				// when !aCalcDerefBufSize.
				if (!aCalcDerefBufSize || this_arg.is_expression) // i.e. we want the total size of what the args resolve to.
					space_needed += (VarSizeType)strlen(this_arg.text) + 1;  // +1 for the zero terminator.
				// else don't increase space_needed, even by 1 for the zero terminator, because
				// the terminator isn't needed if the arg won't exist in the buffer at all.
				continue;
			}
			if (SINGLE_ISOLATED_DEREF)
				the_only_var_of_this_arg = this_arg.deref[0].var;
		}
		if (the_only_var_of_this_arg)
		{
			include_this_arg = !aCalcDerefBufSize || this_arg.is_expression;  // i.e. caller wanted its size unconditionally included
			if (!include_this_arg)
			{
				result = ArgMustBeDereferenced(the_only_var_of_this_arg, iArg);
				if (result == FAIL)
					return VARSIZE_ERROR;
				if (result == CONDITION_TRUE) // The size of these types of args is always included.
					include_this_arg = true;
				// else leave it as false
			}
			if (!include_this_arg) // No extra space is needed in the buffer for this arg.
				continue;
			space = the_only_var_of_this_arg->Get() + 1;  // +1 for the zero terminator.
			// NOTE: Get() (with no params) can retrieve a size larger that what winds up actually
			// being needed, so our callers should be aware that that can happen.
			if (this_arg.is_expression) // Space is needed for the result of the expression or the expanded expression itself, whichever is greater.
				space_needed += (space > MAX_FORMATTED_NUMBER_LENGTH) ? space : MAX_FORMATTED_NUMBER_LENGTH + 1;
			else
				space_needed += space;
			continue;
		}
		// Otherwise: This arg has more than one deref, or a single deref with some literal text around it.
		space = (VarSizeType)strlen(this_arg.text);
		for (deref = this_arg.deref; deref && deref->marker; ++deref)
		{
			space -= deref->length;                              // Substitute the length of the variable contents
			space += deref->var->Get() + this_arg.is_expression; // for that of the literal text of the deref.
			// Above adds 1 if this arg is an expression to the insertion of an extra space after
			// every single deref.  This is done for parsing reasons described in ExpandExpression().
			// NOTE: Get() (with no params) can retrieve a size larger that what winds up actually
			// being needed, so our callers should be aware that that can happen.
		}
		// +1 for this arg's zero terminator in the buffer:
		++space;
		if (this_arg.is_expression) // Space is needed for the result of the expression or the expanded expression itself, whichever is greater.
			space_needed += (space > MAX_FORMATTED_NUMBER_LENGTH) ? space : MAX_FORMATTED_NUMBER_LENGTH + 1;
		else
			space_needed += space;
	}
	return space_needed;
}



ResultType Line::ArgMustBeDereferenced(Var *aVar, int aArgIndexToExclude)
// Returns CONDITION_TRUE, CONDITION_FALSE, or FAIL.
{
	if (mActionType == ACT_SORT) // See PerformSort() for why it's always dereferenced.
		return CONDITION_TRUE;
	if (aVar->mType == VAR_CLIPBOARD)
		// Even if the clipboard is both an input and an output var, it still
		// doesn't need to be dereferenced into the temp buffer because the
		// clipboard has two buffers of its own.  The only exception is when
		// the clipboard has only files on it, in which case those files need
		// to be converted into plain text:
		return CLIPBOARD_CONTAINS_ONLY_FILES ? CONDITION_TRUE : CONDITION_FALSE;
	if (aVar->mType != VAR_NORMAL || !aVar->Length() || aVar == g_ErrorLevel)
		// Reserved vars must always be dereferenced due to their volatile nature.
		// Normal vars of length zero are dereferenced because they might exist
		// as system environment variables, whose contents are also potentially
		// volatile (i.e. they are sometimes changed by outside forces).
		// As of v1.0.25.12, g_ErrorLevel is always dereferenced also so that
		// a command that sets ErrorLevel can itself use ErrorLevel as in
		// this example: StringReplace, EndKey, ErrorLevel, EndKey:
		return CONDITION_TRUE;
	// Since the above didn't return, we know that this is a NORMAL input var of
	// non-zero length.  Such input vars only need to be dereferenced if they are
	// also used as an output var by the current script line:
	Var *output_var;
	for (int iArg = 0; iArg < mArgc; ++iArg)
		if (iArg != aArgIndexToExclude && mArg[iArg].type == ARG_TYPE_OUTPUT_VAR)
		{
			if (   !(output_var = ResolveVarOfArg(iArg, false))   )
				return FAIL;  // It will have already displayed the error.
			if (output_var == aVar)
				return CONDITION_TRUE;
		}
	// Otherwise:
	return CONDITION_FALSE;
}



inline char *Line::ExpandArg(char *aBuf, int aArgIndex)
// Caller must be sure not to call this for an arg that's marked as an expression, since
// expressions are handled by a different function.
// Caller must ensure that aBuf is large enough to accommodate the translation
// of the Arg.  No validation of above params is done, caller must do that.
// Returns a pointer to the char in aBuf that occurs after the zero terminator
// (because that's the position where the caller would normally resume writing
// if there are more args, since the zero terminator must normally be retained
// between args).
{
	// This should never be called if the given arg is an output var, but we check just in case:
	if (mArg[aArgIndex].type == ARG_TYPE_OUTPUT_VAR)
	{
#ifdef _DEBUG
		LineError("ExpandArg() was called to expand an arg that contains only an output variable.");
#endif
		return NULL;
	}

	// Always do this part before attempting to traverse the list of dereferences, since
	// such an attempt would be invalid in this case:
	Var *the_only_var_of_this_arg = NULL;
	if (mArg[aArgIndex].type == ARG_TYPE_INPUT_VAR)
		if (   !(the_only_var_of_this_arg = ResolveVarOfArg(aArgIndex, false))   )
			return NULL;  // The above will have already displayed the error.

	if (the_only_var_of_this_arg)
		// +1 so that we return the position after the terminator, as required.
		return aBuf += the_only_var_of_this_arg->Get(aBuf) + 1;

	char *pText, *this_marker;
	DerefType *deref;
	for (pText = mArg[aArgIndex].text  // Start at the begining of this arg's text.
		, deref = mArg[aArgIndex].deref  // Start off by looking for the first deref.
		; deref && deref->marker; ++deref)  // A deref with a NULL marker terminates the list.
	{
		// FOR EACH DEREF IN AN ARG (if we're here, there's at least one):
		// Copy the chars that occur prior to deref->marker into the buffer:
		for (this_marker = deref->marker; pText < this_marker; *aBuf++ = *pText++); // this_marker is used to help performance.

		// Now copy the contents of the dereferenced var.  For all cases, aBuf has already
		// been verified to be large enough, assuming the value hasn't changed between the
		// time we were called and the time the caller calculated the space needed.
		aBuf += deref->var->Get(aBuf);
		// Finally, jump over the dereference text. Note that in the case of an expression, there might not
		// be any percent signs within the text of the dereference, e.g. x + y, not %x% + %y%.
		pText += deref->length;
	}
	// Copy any chars that occur after the final deref into the buffer:
	for (; *pText; *aBuf++ = *pText++);
	// Terminate the buffer, even if nothing was written into it:
	*aBuf++ = '\0';
	return aBuf; // Returns the position after the terminator.
}



char *Line::ExpandExpression(char *aBuf, int aArgIndex)
// Returns a pointer to the char in aBuf that occurs after the zero terminator
// (because that's the position where the caller would normally resume writing
// if there are more args, since the zero terminator must normally be retained
// between args).
// Caller must ensure that aBuf is large enough to accommodate the translation
// of the Arg.  No validation of above params is done, caller must do that.
// Thanks to Joost Mulders for providing the expression evaluation code upon
// which this function is based.
{
	char *aBuf_orig = aBuf;

	map_item map[MAX_DEREFS_PER_ARG*2 + 1];
	int map_count = 0;
	// Above sizes the map to "times 2 plus 1" to handle worst case, which is -y + 1 (raw+deref+raw).
	// Thus, if this particular arg has the maximum number of derefs, the number of map markers
	// needed would be twice that, plus one for the last raw text's marker.

	///////////////////////////////////////////////////////////////////////////////////////
	// EXPAND DEREFS AND MAKE A MAP THAT INDICATES THE POSITIONS IN THE BUFFER WHERE DEREFS
	// VS. RAW TEXT BEGIN AND END.
	///////////////////////////////////////////////////////////////////////////////////////
	char *pText, *this_marker;
	DerefType *deref;
	for (pText = mArg[aArgIndex].text  // Start at the begining of this arg's text.
		, deref = mArg[aArgIndex].deref  // Start off by looking for the first deref.
		; deref && deref->marker; ++deref)  // A deref with a NULL marker terminates the list.
	{
		// FOR EACH DEREF IN AN ARG (if we're here, there's at least one):
		if (pText < deref->marker)
		{
			map[map_count].type = EXP_RAW;
			map[map_count].marker = aBuf;  // Indicate its position in the buffer.
			++map_count;
			// Copy the chars that occur prior to deref->marker into the buffer:
            for (this_marker = deref->marker; pText < this_marker; *aBuf++ = *pText++); // this_marker is used to help performance.
		}

		// Now copy the contents of the dereferenced var.  For all cases, aBuf has already
		// been verified to be large enough, assuming the value hasn't changed between the
		// time we were called and the time the caller calculated the space needed.
		map[map_count].type = (*deref->marker == g_DerefChar) ? EXP_DEREF_DOUBLE : EXP_DEREF_SINGLE;
		map[map_count].marker = aBuf;  // Indicate its position in the buffer.
		aBuf += deref->var->Get(aBuf); // Might retrieve a blank, in which case this item in the map will be the empty string.
		// For performance reasons, the expression parser relies on an extra space to the right of each
		// single deref.  For example, (x=str), which is seen as (x_contents=str_contents) during
		// evaluation, would instead me seen as (x_contents =str_contents ), which allows string
		// terminators to be put in place of those two spaces in case either or both contents-items
		// are strings rather than numbers (such termination also simplifies number recognition).
		// GetExpandedArgSize() has already ensured there is enough room in the deref buffer for these:
		if (map[map_count].type == EXP_DEREF_SINGLE)
			*aBuf++ = ' ';
		++map_count; // i.e. don't increment until after we're done using the old value.
		// Finally, jump over the dereference text. Note that in the case of an expression, there might not
		// be any percent signs within the text of the dereference, e.g. x + y, not %x% + %y%.
		pText += deref->length;
	}
	// Copy any chars that occur after the final deref into the buffer:
	if (*pText)
	{
		map[map_count].type = EXP_RAW;
		map[map_count].marker = aBuf;  // Indicate its position in the buffer.
		++map_count;
		for (; *pText; *aBuf++ = *pText++);
	}

	// Terminate the buffer, even if nothing was written into it:
	*aBuf = '\0';

/////////////////////////////////////////

	// Although having a precedence array is probably not strictly required (since the order of evaluation
	// of operators of equal precedence doesn't seem to matter for any of the operators in the list),
	// it probably helps performance by avoiding unnecessary pushing and popping of operators to the stack.
	// This array must be kept in sync with "enum SymbolType".  Also, dimensioning explicitly by SYM_COUNT
	// helps enforce that at compile-time:
	static int sPrecedence[SYM_COUNT] =
	{
		  0, 0, 0, 0, 0 // SYM_STRING, SYM_INTEGER, SYM_FLOAT, SYM_OPERAND, SYM_BEGIN (SYM_BEGIN must be lowest precedence).
		, 1, 1          // SYM_OPAREN, SYM_CPAREN (to simplify the code, these must be lower than all operators in precedence).
		, 2             // SYM_OR
		, 3             // SYM_AND
		, 4             // SYM_LOWNOT  // The low precedence version of logical-not.
		, 5, 5, 5       // SYM_EQUAL, SYM_EQUALCASE, SYM_NOTEQUAL (lower prec. than the below so that "x < 5 = var" means "result of comparison is the boolean value in var".
		, 6, 6, 6, 6    // SYM_GT, SYM_LT, SYM_GTOE, SYM_LTOE
		, 7             // SYM_BITOR  // Seems more intuitive to have these higher in prec. than the above, unlike C and Perl, but like Python.
		, 8             // SYM_BITXOR
		, 9             // SYM_BITAND
		, 10, 10        // SYM_BITSHIFTLEFT, SYM_BITSHIFTRIGHT
		, 11, 11        // SYM_PLUS, SYM_MINUS
		, 12, 12        // SYM_TIMES, SYM_DIVIDE
		, 13, 13, 13    // SYM_NEGATIVE (unary minus), SYM_HIGHNOT (the high precedence "not" operator), SYM_BITNOT
		, 14            // SYM_POWER (see note below).
	};
	// Most programming languages give exponentiation a higher precedence than unary minus and !/not.
	// For example, -2**2 is evaluated as -(2**2), not (-2)**2 (the latter is unsupported by qmathPow anyway).
	// However, this rule requires a small workaround in the postfix-builder to allow 2**-2 to be
	// evaluated as 2**(-2) rather than being seen as an error.
	// On a related note, the right-to-left tradition of of something like 2**3**4 is not implemented.
	// Instead, the expression is evaluated in left-to-right (like other operators) to simplify the code.

	#define MAX_TOKENS 512 // Max number of operators/operands.  Seems enough to handle anything realistic, while conserving call-stack space.
	ExprTokenType infix[MAX_TOKENS], postfix[MAX_TOKENS], stack[MAX_TOKENS];
	int infix_count = 0, postfix_count = 0, stack_count = 0;
	// Above dimensions the stack to be as large as the infix/postfix arrays to cover worst-case
	// scenarios and avoid having to check for overflow.  For the infix-to-postfix conversion, the
	// stack must be large enough to hold a malformed expression consisting entirely of operators
	// (though other checks might prevent this).  It must also be large enough for use by the final
	// expression evaluation phase, the worst-case of which is unknown but certainly not larger
	// than MAX_TOKENS.

	////////////////////////////////////////////////////////////////////////////////////////////
	// TOKENIZE THE INFIX EXPRESSION INTO AN ARRAY: Avoids the performance overhead of having to
	// re-detect whether each symbol is an operand vs. operator at multiple stages.
	////////////////////////////////////////////////////////////////////////////////////////////
	SymbolType sym_prev;
	char *item_end, *op_end, *cp, *temp, *terminate_string_here, buf_temp[MAX_FORMATTED_NUMBER_LENGTH + 1];
	size_t op_length;
	Var *found_var;
	for (int map_index = 0; map_index < map_count; ++map_index) // For each deref and raw item in map.
	{
		// Because neither the postfix array nor the stack can ever wind up with more tokens than were
		// contained in the original infix array, only the infix array need be checked for overflow:
		if (infix_count >= MAX_TOKENS) // No room for this operator or operand to be added.
			goto fail;

		map_item &this_item = map[map_index];
		item_end = (map_index < map_count - 1) ? map[map_index + 1].marker : this_item.marker + strlen(this_item.marker);

		// At this stage, an EXP_DEREF_SINGLE item is seen as a numeric literal or a string-literal
		// (without enclosing double quotes, since those are only needed for raw string literals).
		// An EXP_DEREF_SINGLE item cannot extend beyond into the map item to its right, since such
		// a condition can never occur due to load-time preparsing (e.g. the x and y in x+y are two
		// separate items because there's an operator between them). Even a concat expression such as
		// (x y) would still have x and y separate because the space between them counts as a raw
		// map item, which keeps them separate.
		if (this_item.type == EXP_DEREF_SINGLE)
		{
			// Single derefs must always be terminated exactly here because that is where the extent
			// of their contents ends, which is important for strings to avoid having extra whitespace
			// incorrectly added to the end.  The previous stage has ensured there is exactly one
			// whitespace to be "sacrificed" at the end of each single-deref to be used for this
			// purpose, which is necessary after each of the operands in an example such as (5+3),
			// which is seen here as (5 +3 ) due to earlier processing.
			*(item_end - 1) = '\0';
			if (infix_count)
			{
				ExprTokenType &prev_infix_item = infix[infix_count - 1];
				if (IS_OPERAND(prev_infix_item.symbol)) // Since it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
				{
					// Under these conditions, this single-deref item is concatenated onto the previous
					// item, forming a single operand.  The entire operand is shifted to the left a little,
					// but the remaining part of the expression beyond this operand's former position is
					// left unchanged because:
					// 1) It avoids the need to shift-left the positions of all the map items that lie
					//    to the right of this item.
					// 2) Better performance since the memmove() is doing a smaller area.
					if (prev_infix_item.marker >= aBuf_orig && prev_infix_item.marker < this_item.marker) // It's not a double deref.
						memmove(prev_infix_item.marker + strlen(prev_infix_item.marker), this_item.marker
							, item_end - this_item.marker);
					else
						prev_infix_item.marker = this_item.marker; // i.e. consider the double to be an empty string for this concat.
						// And prev_infix_item.symbol should already be set to SYM_OPERAND at this stage.
					continue;
					// Currently, attempts to concat something onto a double-deref, or a double onto
					// something, treat the double as an empty string, as documented.
				}
			}
			// Since above didn't "continue":
			infix[infix_count].symbol = SYM_OPERAND;
			infix[infix_count].marker = this_item.marker; // This operand has already been terminated above.
			++infix_count;
			// This map item has been fully processed.  A new loop iteration will be started to move onto
			// the next, if any:
			continue;
		}

		// An EXP_DEREF_DOUBLE item must be a isolated double-reference or one that extends to the right
		// into other map item(s).  If not, a previous iteration would have merged it in with a previous
		// EXP_RAW item and we could never reach this point.
		// At this stage, an EXP_DEREF_DOUBLE looks like one of the following: abc, 33, abcArray (via
		// extending into an item to its right), or 33Array (overlap).  It can also consist more than
		// two adjacent items as in this example: %ArrayName%[%i%][%j%].  That example would appear as
		// MyArray[33][44] here because the first dereferences have already been done.  MyArray[33][44]
		// (and all the other examples here) are not yet operands because the need a second dereference
		// to resolve them into a number or string.
		if (this_item.type == EXP_DEREF_DOUBLE)
		{
			// Find the end of this operand.  StrChrAny() is not used because if *op_end is '\0'
			// (i.e. this_item is the last operand), the strchr() below will find that too:
			for (op_end = this_item.marker; !strchr(EXP_OPERAND_TERMINATORS, *op_end); ++op_end);
			// Note that above has deteremined op_end correctly because any expression, even those not
			// properly formatted, will have an operator or whitespace between each operand and the next.
			// In the following example, let's say var contains the string -3:
			// %Index%Array var
			// The whitespace-char between the two operands above is a member of EXP_OPERAND_TERMINATORS,
			// so it (and not the minus inside "var") marks the end of the first operand. If there were no
			// space, the entire thing would be one operand so it wouldn't matter (though in this case, it
			// would form an invalid var-name since dashes can't exist in them, which is caught later).
			cp = this_item.marker; // Set for use by the label below.
			goto double_deref;
		}

		// RAW is of lower precedence than the above, so is checked last.  For example, if a single
		// or double deref contains double quotes, those quotes do not delimit a string literal.
		// Instead, the quotes themselves are part of the string.  Similarly, a single or double
		// deref containing as string such as 5+3 is a string, not a subexpression to be evaluated.
		// Since the above didn't "goto", this map item is EXP_RAW, which is the only type that
		// can contain operators and raw literal numbers and strings (which are double-quoted when raw).
		for (cp = this_item.marker;; ++infix_count) // For each token inside this map item.
		{
			// Because neither the postfix array nor the stack can ever wind up with more tokens than were
			// contained in the original infix array, only the infix array need be checked for overflow:
			if (infix_count >= MAX_TOKENS) // No room for this operator or operand to be added.
				goto fail;

			// Only spaces and tabs are considered whitespace, leaving newlines and other whitespace characters
			// for possible future use:
			cp = omit_leading_whitespace(cp);
			if (cp >= item_end)
				break; // End of map item (or entire expression if this is the last map item) has been reached.

			terminate_string_here = cp; // See comments near other uses of terminate_string_here below.

			ExprTokenType &this_infix_item = infix[infix_count]; // Might help reduce code size since it's referenced many places below.

			// Check if it's an operator.
			switch(*cp)
			{
			// The most common cases are kept up top to enhance performance if switch() is implemented as if-else ladder.
			case '+':
				sym_prev = infix_count ? infix[infix_count - 1].symbol : SYM_OPAREN; // OPARAN is just a placeholder.
				if (IS_OPERAND(sym_prev) || sym_prev == SYM_CPAREN)
					this_infix_item.symbol = SYM_PLUS;
				else // Remove unary pluses from consideration since they do not change the calculation.
					--infix_count; // Counteract the loop's increment.
				break;
			case '-':
				sym_prev = infix_count ? infix[infix_count - 1].symbol : SYM_OPAREN; // OPARAN is just a placeholder.
				// Must allow consecutive unary minuses because otherwise, the following example
				// would not work correctly when y contains a negative value: var := 3 * -y
				if (sym_prev == SYM_NEGATIVE) // Have this negative cancel out the previous negative.
					infix_count -= 2;  // Subtracts 1 for the loop's increment, and 1 to remove the previous item.
				else // Differentiate between unary minus and the "subtract" operator:
					this_infix_item.symbol = (IS_OPERAND(sym_prev) || sym_prev == SYM_CPAREN) ? SYM_MINUS : SYM_NEGATIVE;
				break;
			case '/':
				this_infix_item.symbol = SYM_DIVIDE;
				break;
			case '*':
				if (*(cp + 1) == '*') // Seems best to reserve ^ for bitwise-xor.  Python, Perl, and other languages also use ** for power.
				{
					++cp; // An additional increment to have loop skip over the second '*' too.
					this_infix_item.symbol = SYM_POWER;
				}
				else
					this_infix_item.symbol = SYM_TIMES;
				break;
			case '!':
				if (*(cp + 1) == '=') // i.e. != is synonymous with <>, which is also already supported by legacy.
				{
					++cp; // An additional increment to have loop skip over the '=' too.
					this_infix_item.symbol = SYM_NOTEQUAL;
				}
				else
					this_infix_item.symbol = SYM_HIGHNOT; // High-precedence counterpart of the word "not".
				break;
			case '(':
				this_infix_item.symbol = SYM_OPAREN;
				break;
			case ')':
				this_infix_item.symbol = SYM_CPAREN;
				break;
			case '=':
				if (*(cp + 1) == '=')
				{
					// In this case, it's not necessary to check cp >= item_end prior to ++cp,
					// since symbols such as > and = can't appear in a double-deref, which at
					// this stage must be a legal variable name:
					++cp; // An additional increment to have loop skip over the other '=' too.
					this_infix_item.symbol = SYM_EQUALCASE;
				}
				else
					this_infix_item.symbol = SYM_EQUAL;
				break;
			case '>':
				switch (*(cp + 1))
				{
				case '=':
					++cp; // An additional increment to have loop skip over the '=' too.
					this_infix_item.symbol = SYM_GTOE;
					break;
				case '>':
					++cp; // An additional increment to have loop skip over the second '>' too.
					this_infix_item.symbol = SYM_BITSHIFTRIGHT;
					break;
				default:
					this_infix_item.symbol = SYM_GT;
				}
				break;
			case '<':
				switch (*(cp + 1))
				{
				case '=':
					++cp; // An additional increment to have loop skip over the '=' too.
					this_infix_item.symbol = SYM_LTOE;
					break;
				case '>':
					++cp; // An additional increment to have loop skip over the '>' too.
					this_infix_item.symbol = SYM_NOTEQUAL;
					break;
				case '<':
					++cp; // An additional increment to have loop skip over the second '<' too.
					this_infix_item.symbol = SYM_BITSHIFTLEFT;
					break;
				default:
					this_infix_item.symbol = SYM_LT;
				}
				break;
			case '&':
				if (*(cp + 1) == '&')
				{
					++cp; // An additional increment to have loop skip over the second '&' too.
					this_infix_item.symbol = SYM_AND;
				}
				else
					this_infix_item.symbol = SYM_BITAND;
				break;
			case '|':
				if (*(cp + 1) == '|')
				{
					++cp; // An additional increment to have loop skip over the second '|' too.
					this_infix_item.symbol = SYM_OR;
				}
				else
					this_infix_item.symbol = SYM_BITOR;
				break;
			case '^':
				this_infix_item.symbol = SYM_BITXOR;
				break;
			case '~':
				this_infix_item.symbol = SYM_BITNOT;
				break;

			case '"': // Raw string literal.
				// Note that single and double-derefs are impossible inside string-literals
				// because the load-time deref parser would never detect anything inside
				// of quotes -- even non-escaped percent signs -- as derefs.
				// Find the end of this string literal, noting that a pair of double quotes is
				// a literal double quote inside the string:
				++cp; // Omit the starting-quote from consideration, and from the operand's eventual contents.
				for (op_end = cp;; ++op_end)
				{
					if (!*op_end) // No matching end-quote. Probably impossible due to load-time validation.
						goto fail;
					if (*op_end == '"') // If not followed immediately by another, this is the end of it.
					{
						++op_end;
						if (*op_end != '"') // String terminator or some non-quote character.
							break;  // The previous char is the ending quote.
						//else a pair of quotes, which resolves to a single literal quote.
						// This pair is skipped over and the loop continues until the real end-quote is found.
					}
				}
				// op_end is now the character after the first literal string's ending quote, which might be the terminator.
				*(--op_end) = '\0'; // Remove the ending quote.
				// Convert all pairs of quotes inside into single literal quotes:
				StrReplaceAll(cp, "\"\"", "\"", true);
				// Above relies on the fact that StrReplaceAll() does not do cascading replacements,
				// meaning that a series of characters such as """" would be correctly converted into
				// two double quotes rather than just collapsing into one.
				// For the below, see comments at (this_item.type == EXP_DEREF_SINGLE) above.
				if (infix_count)
				{
					ExprTokenType &prev_infix_item = infix[infix_count - 1];
					if (IS_OPERAND(prev_infix_item.symbol)) // Since it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
					{
						if (prev_infix_item.marker >= aBuf_orig && prev_infix_item.marker < cp) // It's not a double deref.
							// Below recomputes strlen(cp) in case StrReplaceAll, above, changed it:
							memmove(prev_infix_item.marker + strlen(prev_infix_item.marker), cp, strlen(cp) + 1);
						else
						{
							prev_infix_item.marker = cp; // i.e. consider the double deref to be the empty string for this concat.
							prev_infix_item.symbol = SYM_STRING; // Indicate that it's a literal string.
						}
						--infix_count; // Counteract the loops ++infix_count.
						cp = op_end + 1; // Set it up for the next iteration (terminate_string_here is not needed).
						continue;
					}
				}
				// Since above didn't "continue", this operand is not being combined with the previous.
				this_infix_item.symbol = SYM_STRING;
				this_infix_item.marker = cp; // This string-operand has already been terminated above.
				cp = op_end + 1;  // Set it up for the next iteration (terminate_string_here is not needed in this case).
				continue;

			default: // Numeric-literal, relational operator such as and/or/not, or unrecognized symbol.
				// Unrecognized symbols should be impossible at this stage because load-time validation
				// would have caught them.  Also, a non-pure-numeric operand should also be impossible
				// because string-literals were handled above, and the load-time validator would not
				// have let any raw non-numeric operands get this far (such operands would have been
				// converted to single or double derefs at load-time, in which case they wouldn't be
				// raw and would never reach this point in the code).
				// To conform to the way the load-time pre-parser recognizes and/or/not, and to support
				// things like (x=3)and(5=4) or even "x and!y", the and/or/not operators are processed
				// here with the numeric literals since we want to find op_end the same way.

				// Find the end of this operand or keyword, even if that end is beyond item_end.
				// StrChrAny() is not used because if *op_end is '\0', the strchr() below will find it too:
				for (op_end = cp + 1; !strchr(EXP_OPERAND_TERMINATORS, *op_end); ++op_end);
				// Now op_end marks the end of this operand or keyword.  That end might be the zero terminator
				// or the next operator in the expression, or just a whitespace.
				if (*item_end && op_end >= item_end) // Where item_end is the first character of the next item.
					// This operand extends into one or more map items to the right of this_item.  The first
					// of those map items must be a EXP_DEREF_DOUBLE, since for various reasons, the other types
					// are impossible at this stage.  Example:
					// Array[%i%][%j%] (which at this stage appears to us as Array[33][44]).
					goto double_deref; // This also serves to break out of this for(), equivalent to a break.
				// Otherwise, this operand is a normal raw numeric-literal or a word-operator (and/or/not).
				// The section below is very similar to the one used at load-time to recognize and/or/not,
				// so it should be maintained with that section:
				op_length = op_end - cp;
				if (op_length < 4 && op_length > 1) // Ordered for short-circuit performance.
				{
					// Since this item is of an appropriate length, check if it's AND/OR/NOT:
					if (op_length == 2)
					{
						if ((*cp == 'o' || *cp == 'O') && (cp[1] == 'r' || cp[1] == 'R')) // "OR" was found.
						{
							this_infix_item.symbol = SYM_OR;
							*cp = '\0';  // Terminate any previous raw numeric-literal such as "1 or (x < 3)"
							cp = op_end; // Have the loop process whatever lies at op_end and beyond.
							continue;
						}
					}
					else // op_length must be 3
					{
						switch (*cp)
						{
						case 'a':
						case 'A':
							if ((cp[1] == 'n' || cp[1] == 'N') && (cp[2] == 'd' || cp[2] == 'D')) // "AND" was found.
							{
								this_infix_item.symbol = SYM_AND;
								*cp = '\0';  // Terminate any previous raw numeric-literal such as "1 and (x < 3)"
								cp = op_end; // Have the loop process whatever lies at op_end and beyond.
								continue;
							}
							break;

						case 'n':
						case 'N':
							if ((cp[1] == 'o' || cp[1] == 'O') && (cp[2] == 't' || cp[2] == 'T')) // "NOT" was found.
							{
								this_infix_item.symbol = SYM_LOWNOT;
								*cp = '\0';  // Terminate any previous raw numeric-literal such as "1 not (x < 3)" (even though "not" would be invalid if used this way)
								cp = op_end; // Have the loop process whatever lies at op_end and beyond.
								continue;
							}
							break;
						}
					}
				}
				// Since above didn't "continue", this item is a raw numeric literal, either SYM_FLOAT or
				// SYM_INTEGER (to be differentiated later).
				// For the below, see comments at (this_item.type == EXP_DEREF_SINGLE) above.
				if (infix_count)
				{
					ExprTokenType &prev_infix_item = infix[infix_count - 1];
					if (IS_OPERAND(prev_infix_item.symbol)) // Since it's an operand, at this stage it can only be SYM_OPERAND or SYM_STRING.
					{
						if (prev_infix_item.marker >= aBuf_orig && prev_infix_item.marker < cp) // It's not a double deref.
						{
							// Below moves only the number itself, not any zero terminator (since there isn't one yet):
							temp = prev_infix_item.marker + strlen(prev_infix_item.marker);
							memmove(temp, cp, op_length);
							// The memmove() should have shifted the above to the left by at least one
							// position in each of the following cases:
							// "string literal" 123 --> The position of the final double-quote is available.
							// Var 123 --> There must be at least one space between them for load-time to have parsed it.
							// %Var% 123 --> Not supported (it would never reach here, instead failing at a later stage).
							// Terminate this literal in the right place.  Otherwise, a future iteration's
							// provided terminator will be too far to the right and thus append garbage to
							// the end of this newly concatenated string:
							*(temp + op_length) = '\0';
						}
						else // It's a concat onto a double deref, so treat the double as an empty string.
							prev_infix_item.marker = cp;
							// And prev_infix_item.symbol should already be set to SYM_OPERAND at this stage.
						--infix_count; // Counteract the loops ++infix_count.
						cp = op_end; // Have the loop process whatever lies at op_end and beyond.
						// Not necessary to do terminate_string_here because the following is currently
						// impossible: x:=0"" (load-time error because raw string-literal can't be up
						// against the right side of another operand).
						continue;
					}
				}
				this_infix_item.symbol = SYM_OPERAND;
				this_infix_item.marker = cp; // This numeric operand will be terminated later via terminate_string_here.
				cp = op_end; // Have the loop process whatever lies at op_end and beyond.
				// The below is necessary to support an expression such as (1 "" 0), which
				// would otherwise result in 1"0 instead of 10 because the 1 was lazily
				// terminated by the next iteration rather than our iteration at it's
				// precise viewed-as-string ending point.  It might also be needed for
				// the same reason for concatenating things like (1 var).
				if (IS_SPACE_OR_TAB(*cp))
					*cp++ = '\0';
				continue; // i.e. don't do the terminate_string_here and ++cp steps below.
			} // switch() for type of symbol/operand.

			// If the above didn't "continue", it just processed a non-operand symbol.  So terminate
			// the string at the first character of that symbol (e.g. the first character of <=).
			// This sets up raw operands to be always-terminated, such as the ones in 5+10+20.  Note
			// that this is not done for operator-words (and/or/not) since it's not valid to write
			// something like 1and3 (such a thing would be considered a variable and converted into
			// a single-deref by the load-time pre-parser).  It's done this way because we don't
			// want to convert these raw operands into numbers yet because their original strings
			// might be needed in the case where this operand will be involved in an operation with
			// another string operand, in which case both are treated as strings:
			if (terminate_string_here)
				*terminate_string_here = '\0';
			++cp; // i.e. increment only if a "continue" wasn't encountered somewhere above.
		} // for each token
		continue;  // To avoid falling into the label below.

double_deref:
		if (*item_end && op_end >= item_end) // This operand includes this map item and one or more to its right, so adjust map_index accordingly.
		{
			// Find the map item containing the end of this operand.  If the loop ends when
			// map_index==map_count, it ran out of map items, thus the final map item contains the
			// end of this operand. The other termination condition below occurs when this operand
			// is ended in the map item to the left of map_index:
			for (;;)
			{
				++map_index; // Due to the check above, there must be at least one map item beyond starting map item of this operand.
				if (op_end <= map[map_index].marker // The map item to the right of the one containg the end of this operand has been found.
					// ... except if that item is of zero length, in which case the below consumes it and
					// any other consecutive empty map items to the right of it.  If this weren't done,
					// a double-deref such as 100/(3 + Array%BlankVar%) would not work properly.
					&& ((map_index == map_count - 1) || map[map_index + 1].marker != map[map_index].marker))
				{
					item_end = map[map_index].marker; // End of previous map item is beginning of this one.
					--map_index;  // Set it to the final map item of this operand.
					break;
				}
				if (map_index == map_count - 1) // No more map items.  This final item must be the one that contains the end of this operand.
				{
					item_end = map[map_index].marker + strlen(map[map_index].marker);
					break;
				}
			}
			// If this map item wasn't fully consumed by this operand, alter the map item to contain
			// only the part left to be processed.  For example, in Array[%i%]/3, the final map item
			// is ]/3, of which only the ] is consumed by the Array[%i%] operand:
			if (op_end < item_end)
			{
				map[map_index].marker = op_end;
				--map_index;  // Compensate for the loop's ++map_index.
			}
			//else do nothing since map_index is now set to the final map item of this operand, and that
			// map item is fully consumed by this operand and needs no further processing.
		} // This map item stretches into others to its right.

		// Check if this double is being concatenated onto a previous operand.  If so, it is not
		// currently supported so this double-deref will be treated as an empty string, as documented.
		// Example 1: Var := "abc" %naked_double_ref%
		// Example 2: Var := "abc" Array%Index%
		if (!infix_count || !IS_OPERAND(infix[infix_count - 1].symbol)) // Relies on short-circuit boolean order.
		{
			// Callers of this label have set cp to the start of the variable name and op_end to the
			// position of the character after the last one in the name.  The below does not validate
			// that the variable name is legal.  Instead, it just tries to look up an existing variable.
			// If there is no such variable, it is considered to be the empty string (which seems better
			// than causing the entire expression to evaluate to blank) to indicate the problem.
			// The same thing happens if this is the clipboard because temporary memory would be
			// needed somewhere if the clipboard contains files that need to be expanded to text.
			// Finally, all other non-normal variables (such as reserved variables) are also failure
			// conditions because they would need to be fetched into some kind of temp. memory somewhere.
			// Next condition:
			// An environment variable might not have an entry in the script's variable list if it is
			// only ever referred to dynamically, never directly by name.  In that case, FindVar() below
			// will not find it in the list.  But if it is found in the list, it seems best to make it
			// blank so that the behavior for env. vars is always the same.  Related: all other aspects
			// of script behavior treat variables of length zero as the contents of the corresponding
			// environment variable (if one exists).
			op_length = op_end - cp;
			infix[infix_count].marker = (!op_length || !(found_var = g_script.FindVar(cp, op_length)) // i.e. don't call FindVar with zero for length, since that's a special mode.
				|| found_var->mType != VAR_NORMAL // Relies on short-circuit boolean order.
				|| !found_var->Length()) // i.e. it's either an environment variable or a zero-length variable.  But either way it's treated as a blank string.
					? "" // Var is not found, not a normal var, or it *is* an environment variable.
					: found_var->Contents(); // This operand becomes the variable's contents.
					// "" is used above, rather than a real empty string residing within limits of
					// the expression's text, so that attempts to concat this item will fail in the
					// same way that other concats involving double derefs fail.
			// If this item is the empty string or consists entirely of whitespace, it will be treated
			// as a string-operand rather than a number:
			infix[infix_count].symbol = SYM_OPERAND;
			++infix_count;
		}
	} // for each map item

	////////////////////////////
	// CONVERT INFIX TO POSTFIX.
	////////////////////////////
	int i;
	SymbolType stack_symbol, infix_symbol;
	stack[stack_count++].symbol = SYM_BEGIN; // A flag to indicate that conversion to postfix has begun.
	// Unlike the below, the above isn't a complete PUSH since it only sets the symbol member.
	#define STACK_PUSH(token) stack[stack_count++] = token
	#define STACK_POP stack[--stack_count]  // To be used as the r-value for an assignment.
	for (i = 0; stack_count > 0;) // While SYM_BEGIN is still on the stack, continue iterating.
	{
		stack_symbol = stack[stack_count - 1].symbol; // Frequently used, so resolve only once to help performance.

		// "i" will be out of bounds if infix expression is complete but stack is not empty.
		// So the very first check must be for that.
		if (i == infix_count) // End of infix expression, but loop's check says stack still has items on it.
		{
			if (stack_symbol == SYM_BEGIN) // Stack is basically empty, so stop the loop.
				// Remove SYM_BEGIN from the stack, leaving the stack empty for use in the next stage.
				// This also signals our loop to stop.
				--stack_count;
			else if (stack_symbol == SYM_OPAREN) // Open paren is never closed (currently impossible due to load-time balancing, but kept for completeness).
				goto fail;
			else // Pop item of the stack, and continue iterating to hit this line until stack is empty.
				postfix[postfix_count++] = STACK_POP;
			continue;
		}

		// Only after the above is it safe to use "i" as an index.
		infix_symbol = infix[i].symbol; // Frequently used, so resolve only once to help performance.

		// Put operands into the postfix array immediately, then move on to the next infix item:
		if (IS_OPERAND(infix_symbol)) // At this stage, operands consist of only SYM_OPERAND and SYM_STRING.
		{
			postfix[postfix_count++] = infix[i++];
			continue;
		}

		// Since above didn't "continue", the current infix symbol is not an operand.
		switch(infix_symbol)
		{
		// CPAREN is listed first for performance.  It occurs frequently while emptying the stack to search
		// for the matching open-paren:
		case SYM_CPAREN:
			if (stack_symbol == SYM_OPAREN) // The first open-paren on the stack must be the one that goes with this close-paren.
			{
				--stack_count; // Remove this open-paren from the stack, since it is now complete.
				++i;           // Since this pair of parens is done, move on to the next token in the infix expression.
			}
			else if (stack_symbol == SYM_BEGIN) // Paren is closed without having been opened (currently impossible due to load-time balancing, but kept for completeness).
				goto fail; 
			else
				postfix[postfix_count++] = STACK_POP;
				// By not incrementing i, the loop will continue to encounter SYM_CPAREN and thus
				// continue to pop things off the stack until the corresponding OPAREN is reached.
			break;

		// Open-parens always go on the stack to await their matching close-parens:
		case SYM_OPAREN:
			STACK_PUSH(infix[i++]);
			break;

		default: // Symbol is an operator, so act according to its precedence.
			// If the symbol waiting on the stack has a lower precedence than the current symbol, push the
			// current symbol onto the stack so that it will be processed sooner than the waiting one.
			// Otherwise, pop waiting items off the stack (by means of i not being incremented) until their
			// precedence falls below the current item's precedence, or the stack is emptied.
			// Note: BEGIN and OPAREN are the lowest precedence items ever to appear on the stack (CPAREN
			// never goes on the stack, so can't be encountered there).
			if (   sPrecedence[stack_symbol] < sPrecedence[infix_symbol]
				|| (stack_symbol == SYM_POWER && infix_symbol == SYM_NEGATIVE)   )
				// The line above is a workaround to allow 2^-2 to be evaluated as 2^(-2) rather
				// than being seen as an error.  However, for simplicity of code, consecutive
				// unary operators are not supported:
				// !-3  ; Not supported (seems of little use anyway; can be written as !(-3) to make it work).
				// -!3  ; Not supported (seems useless anyway, can be written as -(!3) to make it work).
				// !x   ; Supported even if X contains a negative number, since x is recognized as an isolated operand and not something containing unary minus.
				STACK_PUSH(infix[i++]);
			else // Stack item has equal or greater precedence (if equal, left-to-right evaluation order is in effect).
				postfix[postfix_count++] = STACK_POP;
		} // switch()
	} // End of loop that builds postfix array.

	///////////////////////////////////////////////////
	// EVALUATE POSTFIX EXPRESSION (constructed above).
	///////////////////////////////////////////////////
	double right_double, left_double;
	__int64 right_int64, left_int64;
	char *right_string, *left_string;
	SymbolType right_is_number, left_is_number;
	for (i = 0; i < postfix_count; ++i) // For each item in the postfix array: push operands onto stack, evaluate operators and push their results onto stack.
	{
		ExprTokenType &this_token = postfix[i];  // For performance and convenience.

		// At this stage, operands in the postfix array should be either SYM_OPERAND or SYM_STRING.
		// But all are checked since that operation is just as fast:
		if (IS_OPERAND(this_token.symbol)) // If it's an operand, just push it onto stack for use by an operator in a future iteration.
		{
			STACK_PUSH(this_token);
			continue;
		}

		// Otherwise, this token must be a unary or binary operator.
		// GET THE FIRST OPERAND FOR THIS OPERATOR (for non-unary operators, this is the right-side operand):
		if (!stack_count) // Prevent stack underflow.  An expression such as -*3 causes this.
			goto fail;
		ExprTokenType &right = STACK_POP;
		// Below uses IS_OPERAND rather than checking for only SYM_OPERAND because the stack can contain
		// both generic and specific operands.  Specific operands were evaluated by a previous iteration
		// of this section.  Generic ones were pushed as-is onto the stack by a previous iteration.
		if (!IS_OPERAND(right.symbol)) // Haven't found a way to produce this situation yet, but safe to assume it's possible.
			goto fail;
		// If the operand is still generic/undetermined, find out whether it is a string, integer, or float:
		right_is_number = (right.symbol == SYM_OPERAND) ? IsPureNumeric(right.marker, true, false, true) : right.symbol;

		// IF THIS IS A UNARY OPERATOR, we now have the single operand needed to perform the operation.
		// The cases in the switch() below are all unary operators.  The other operators are handled
		// in the switch()'s default section:
		switch (this_token.symbol)
		{
		case SYM_NEGATIVE:  // Unary-minus.
			if (right_is_number == PURE_FLOAT)
				// Overwrite this_token's union with a float. No need to have the overhead of ATOF() since it can't be hex.
				this_token.value_double = -(right.symbol == SYM_OPERAND ? atof(right.marker) : right.value_double);
			else if (right_is_number == PURE_INTEGER)
				this_token.value_int64 = -(right.symbol == SYM_OPERAND ? ATOI64(right.marker) : right.value_int64);
			else // String.
			{
				// Seems best to consider the application of unary minus to a string, even a quoted string
				// literal such as "15", to be a failure.  UPDATE: For v1.0.25.06, invalid operations like
				// this instead treat the operand as an empty string.  This avoids aborting a long, complex
				// expression entirely just because on of its operands is invalid.  However, the net effect
				// in most cases might be the same, since the empty string is a non-numeric result and thus
				// will cause any operator it is involved with to treat its other operand as a string too.
				// And the result of a math operation on two strings is typically an empty string.
				this_token.marker = "";
				this_token.symbol = SYM_STRING;
				break;
			}
			// Since above didn't "break":
			this_token.symbol = right_is_number; // Convert generic SYM_OPERAND into a specific type: float or int.
			break;

		// Both nots are equivalent at this stage because precedence was already acted upon by infix-to-postfix:
		case SYM_LOWNOT:  // The operator-word "not".
		case SYM_HIGHNOT: // The symbol !
			if (right_is_number == PURE_FLOAT) // Convert to float, not int, so that a number between 0.0001 and 0.9999 is considered "true".
				// Using ! vs. comparing explicitly to 0.0 might generate faster code, and K&R implies it's okay:
				this_token.value_int64 = !(right.symbol == SYM_OPERAND ? atof(right.marker) : right.value_double);
			else if (right_is_number == PURE_INTEGER)
				this_token.value_int64 = !(right.symbol == SYM_OPERAND ? ATOI64(right.marker) : right.value_int64);
			else // This is either a non-numeric string or a numeric raw literal string such as "123".
				// All non-numeric strings are considered TRUE here.  In addition, any raw literal string,
				// even "0", is considered to be TRUE.  This relies on the fact that right.symbol will be
				// SYM_OPERAND/generic (and thus handled higher above) for all pure-numeric strings except
				// explicit raw literal strings.  Thus, if something like !"0" ever appears in an expression,
				// it evaluates to !true.  EXCEPTION: Because "if x" evaluates to false when X is blank,
				// it seems best to have "if !x" evaluate to TRUE.
				this_token.value_int64 = !*right.marker; // i.e. result is false except for empty string because !"string" is false.
			this_token.symbol = SYM_INTEGER; // Result of above is always a boolean integer (one or zero).
			break;

		case SYM_BITNOT: // The tilde (~) operator.
			if (right_is_number == PURE_FLOAT)
				// Overwrite this_token's union with a float. No need to have the overhead of ATOI64() since PURE_FLOAT can't be hex.
				this_token.value_int64 = right.symbol == SYM_OPERAND ? _atoi64(right.marker) : (__int64)right.value_double;
			else if (right_is_number == PURE_INTEGER) // But in this case, it can be hex, so use ATOI64().
				this_token.value_int64 = right.symbol == SYM_OPERAND ? ATOI64(right.marker) : right.value_int64;
			else // String.  Seems best to consider the application of unary minus to a string, even a quoted string literal such as "15", to be a failure.
			{
				this_token.marker = "";
				this_token.symbol = SYM_STRING;
				break;
			}
			// Since above didn't "break":
			// Note that it is not legal to perform ~, &, |, or ^ on doubles.  Because of this, and also to
			// conform to the behavior of the Transform command, any floating point operand is truncated to
			// an integer above.
			if (this_token.value_int64 < 0 || this_token.value_int64 > UINT_MAX)
				this_token.value_int64 = ~this_token.value_int64;
			else // See comments at TRANS_CMD_BITNOT for why it's done this way.
				this_token.value_int64 = ~((DWORD)this_token.value_int64);
			this_token.symbol = right_is_number; // Convert generic SYM_OPERAND into a specific type: float or int.
			break;

		default: // Non-unary operator.
			// GET THE SECOND (LEFT-SIDE) OPERAND FOR THIS OPERATOR:
			if (!stack_count) // Prevent stack underflow.
				goto fail;
			ExprTokenType &left  = STACK_POP; // i.e. the right operand always comes off the stack before the left.
			if (!IS_OPERAND(left.symbol)) // Haven't found a way to produce this situation yet, but safe to assume it's possible.
				goto fail;
			// If the operand is still generic/undetermined, find out whether it is a string, integer, or float:
			left_is_number = (left.symbol == SYM_OPERAND) ? IsPureNumeric(left.marker, true, false, true) : left.symbol;

			if (this_token.symbol == SYM_AND || this_token.symbol == SYM_OR)
			{
				// See comments in SYM_LOWNOT about the below:
				if (left_is_number == PURE_FLOAT)
					this_token.value_int64 = (left.symbol == SYM_OPERAND ? atof(left.marker) : left.value_double) != 0.0;
				else if (left_is_number == PURE_INTEGER)
					this_token.value_int64 = (left.symbol == SYM_OPERAND ? ATOI64(left.marker) : left.value_int64) != 0;
				else // string.
					// Since "if x" evaluates to false when x is blank, it seems best to also have blank
					// strings resolve to false when used in more complex ways. In other words "if x or y"
					// should be false if both x and y are blank.  Logical-not also follows this convention.
					this_token.value_int64 = *left.marker != '\0';

				// The check below is a very limited form of short-circuit boolean.  From the script's
				// point-of-view, it isn't short-circuit boolean at all since all it does is help
				// performance a little.  That's because by this stage, the right-side operand has
				// already been fully evaluated if it was itself a sub-expression.
				if (   this_token.symbol == SYM_AND && this_token.value_int64     // Left side is true so right-side will determine the final result.
					|| this_token.symbol == SYM_OR && !this_token.value_int64   ) // Left side is false so right-side will determine the final result.
				{
					// See comments in SYM_LOWNOT about the below:
					if (right_is_number == PURE_FLOAT)
						this_token.value_int64 = (right.symbol == SYM_OPERAND ? atof(right.marker) : right.value_double) != 0.0;
					else if (right_is_number == PURE_INTEGER)
						this_token.value_int64 = (right.symbol == SYM_OPERAND ? ATOI64(right.marker) : right.value_int64) != 0;
					else
						this_token.value_int64 = *right.marker != '\0'; // See above for explanation.
				}
				//else the left-side alone was enough to determine and store the outcome.

				this_token.symbol = SYM_INTEGER; // Result of AND or OR is always a boolean integer (one or zero).
			}

			else if (!right_is_number || !left_is_number) // At least one is a string, so this is a string operation.
			{
				// Above check has ensured that at least one of them is a string.  But the other
				// one might be a number such as in 5+10="15", in which 5+10 would be a numerical
				// result being compared to the raw string literal "15".
				switch (right.symbol)
				{
				// Seems best to obey SetFormat for these two, though it's debatable:
				case SYM_INTEGER: ITOA64(right.value_int64, buf_temp); right_string = buf_temp; break;
				case SYM_FLOAT: snprintf(buf_temp, sizeof(buf_temp), g.FormatFloat, right.value_double); right_string = buf_temp; break;
				default: right_string = right.marker; // SYM_STRING or SYM_OPERAND, which is already in the right format.
				}

				switch (left.symbol)
				{
				// Seems best to obey SetFormat for these two, though it's debatable:
				case SYM_INTEGER: ITOA64(left.value_int64, buf_temp); left_string = buf_temp; break;
				case SYM_FLOAT: snprintf(buf_temp, sizeof(buf_temp), g.FormatFloat, left.value_double); left_string = buf_temp; break;
				default: left_string = left.marker; // SYM_STRING or SYM_OPERAND, which is already in the right format.
				}
				
				#undef STRING_COMPARE
				#define STRING_COMPARE (g.StringCaseSense ? strcmp(left_string, right_string) : stricmp(left_string, right_string))

				switch(this_token.symbol)
				{
				case SYM_EQUAL:     this_token.value_int64 = !stricmp(left_string, right_string); break;
				case SYM_EQUALCASE: this_token.value_int64 = !strcmp(left_string, right_string); break;
				// The rest all obey g.StringCaseSense since they have no case sensitive counterparts:
				case SYM_NOTEQUAL:  this_token.value_int64 = STRING_COMPARE ? 1 : 0; break;
				case SYM_GT:        this_token.value_int64 = STRING_COMPARE > 0; break;
				case SYM_LT:        this_token.value_int64 = STRING_COMPARE < 0; break;
				case SYM_GTOE:      this_token.value_int64 = STRING_COMPARE >= 0; break;
				case SYM_LTOE:      this_token.value_int64 = STRING_COMPARE <= 0; break;
				default:
					// Other operators do not support string operands, so the result is an empty string.
					this_token.marker = "";
					this_token.symbol = SYM_STRING;
					goto push_this_result;
				}
				// Since above didn't "goto":
				this_token.symbol = SYM_INTEGER; // Boolean result is treated as an integer.  Must be done only after the switch() above.
			}

			else if (right_is_number == PURE_INTEGER && left_is_number == PURE_INTEGER && this_token.symbol != SYM_DIVIDE
				|| (this_token.symbol == SYM_BITAND || this_token.symbol == SYM_BITOR
				|| this_token.symbol == SYM_BITXOR || this_token.symbol == SYM_BITSHIFTLEFT
				|| this_token.symbol == SYM_BITSHIFTRIGHT)) // The bitwise operators convert any floating points to integer prior to the calculation.
			{
				// Because both are integers and the operation isn't division, the result is integer.
				// The result is also an integer for the bitwise operations listed in the if-statement
				// above.  This is because it is not legal to perform ~, &, |, or ^ on doubles, and also
				// because this behavior conforms to that of the Transform command.  Any floating point
				// operands are truncated to integers prior to doing the bitwise operation.

				switch (right.symbol)
				{
				case SYM_OPERAND: right_int64 = ATOI64(right.marker); break;
				case SYM_INTEGER: right_int64 = right.value_int64; break;
				default: right_int64 = (__int64)right.value_double; // SYM_FLOAT (the least common).
				// It can't be SYM_STRING because in here, both right and left are known to be numbers.
				}

				switch (left.symbol)
				{
				case SYM_OPERAND: left_int64 = ATOI64(left.marker); break;
				case SYM_INTEGER: left_int64 = left.value_int64; break;
				default: left_int64 = (__int64)left.value_double; // SYM_FLOAT (the least common).
				// It can't be SYM_STRING because in here, both right and left are known to be numbers.
				}

				switch(this_token.symbol)
				{
				// The most common cases are kept up top to enhance performance if switch() is implemented as if-else ladder.
				case SYM_PLUS:     this_token.value_int64 = left_int64 + right_int64; break;
				case SYM_MINUS:	   this_token.value_int64 = left_int64 - right_int64; break;
				case SYM_TIMES:    this_token.value_int64 = left_int64 * right_int64; break;
				// A look at K&R confirms that relational/comparison operations and logical-AND/OR/NOT
				// always yield a one or a zero rather than arbitrary non-zero values:
				case SYM_EQUALCASE: // Same behavior as SYM_EQUAL for numeric operands.
				case SYM_EQUAL:    this_token.value_int64 = left_int64 == right_int64; break;
				case SYM_NOTEQUAL: this_token.value_int64 = left_int64 != right_int64; break;
				case SYM_GT:       this_token.value_int64 = left_int64 > right_int64; break;
				case SYM_LT:       this_token.value_int64 = left_int64 < right_int64; break;
				case SYM_GTOE:     this_token.value_int64 = left_int64 >= right_int64; break;
				case SYM_LTOE:     this_token.value_int64 = left_int64 <= right_int64; break;
				case SYM_BITAND:   this_token.value_int64 = left_int64 & right_int64; break;
				case SYM_BITOR:    this_token.value_int64 = left_int64 | right_int64; break;
				case SYM_BITXOR:   this_token.value_int64 = left_int64 ^ right_int64; break;
				case SYM_BITSHIFTLEFT:  this_token.value_int64 = left_int64 << right_int64; break;
				case SYM_BITSHIFTRIGHT: this_token.value_int64 = left_int64 >> right_int64; break;
				case SYM_POWER:
					// The following comment is from TRANS_CMD_POW.  For consistency, the same policy is applied here:
					// Currently, a negative aValue1 isn't supported.
					// The reason for this is that since fractional exponents are supported (e.g. 0.5, which
					// results in the square root), there would have to be some extra detection to ensure
					// that a negative aValue1 is never used with fractional exponent (since the sqrt of
					// a negative is undefined).  In addition, qmathPow() doesn't support negatives, returning
					// an unexpectedly large value or -1.#IND00 instead.  Also note that zero raised to
					// a negative power is undefined, similar to division-by-zero, and thus treated as a failure.
					if (left_int64 < 0 || (!left_int64 && right_int64 < 0)) // See comments at TRANS_CMD_POW about this.
					{
						// Return a consistent result rather than something that varies:
						this_token.marker = "";
						this_token.symbol = SYM_STRING;
						goto push_this_result;
					}
					if (right_int64 < 0)
					{
						this_token.value_double = qmathPow((double)left_int64, (double)right_int64);
						this_token.symbol = SYM_FLOAT;  // Due to negative exponent, override to float like TRANS_CMD_POW.
					}
					else
						this_token.value_int64 = (__int64)qmathPow((double)left_int64, (double)right_int64);
					break;
				}
				if (this_token.symbol != SYM_FLOAT)  // It wasn't overridden by SYM_POWER.
					this_token.symbol = SYM_INTEGER; // Must be done only after the switch() above.
			}

			else // Since one or both operands is floating point (or this is the division of two integers), the result will be floating point.
			{
				// For these two, use ATOF vs. atof so that if one of them is an integer to be converted
				// to a float for the purpose of this calculation, hex will be supported:
				switch (right.symbol)
				{
				case SYM_OPERAND: right_double = ATOF(right.marker); break;
				case SYM_INTEGER: right_double = (double)right.value_int64; break;
				default: right_double = right.value_double; break; // SYM_FLOAT (the least common).
				// It can't be SYM_STRING because in here, both right and left are known to be numbers.
				}

				switch (left.symbol)
				{
				case SYM_OPERAND: left_double = ATOF(left.marker); break;
				case SYM_INTEGER: left_double = (double)left.value_int64; break;
				default: left_double = left.value_double; break; // SYM_FLOAT (the least common).
				// It can't be SYM_STRING because in here, both right and left are known to be numbers.
				}

				switch(this_token.symbol)
				{
				case SYM_PLUS:     this_token.value_double = left_double + right_double; break;
				case SYM_MINUS:	   this_token.value_double = left_double - right_double; break;
				case SYM_TIMES:    this_token.value_double = left_double * right_double; break;
				case SYM_DIVIDE:
					if (right_double == 0.0)
					{
						this_token.marker = "";
						this_token.symbol = SYM_STRING;
						goto push_this_result;
					}
					// Otherwise:
					this_token.value_double = left_double / right_double;
					break;
				case SYM_EQUALCASE: // Same behavior as SYM_EQUAL for numeric operands.
				case SYM_EQUAL:    this_token.value_double = left_double == right_double; break;
				case SYM_NOTEQUAL: this_token.value_double = left_double != right_double; break;
				case SYM_GT:       this_token.value_double = left_double > right_double; break;
				case SYM_LT:       this_token.value_double = left_double < right_double; break;
				case SYM_GTOE:     this_token.value_double = left_double >= right_double; break;
				case SYM_LTOE:     this_token.value_double = left_double <= right_double; break;
				case SYM_POWER: // See the other SYM_POWER higher above for explanation of the below:
					if (left_double < 0 || (left_double == 0.0 && right_double < 0))
					{
						this_token.marker = "";
						this_token.symbol = SYM_STRING;
						goto push_this_result;
					}
					// Otherwise:
					this_token.value_double = qmathPow(left_double, right_double);
					break;
				}
				this_token.symbol = SYM_FLOAT; // Must be done only after the switch() above.
			} // Result is floating point.
		} // switch() operator type

push_this_result:
		STACK_PUSH(this_token); // Push the result onto the stack for use as an operand by a future operator.
	} // For each item in the postfix array.

	if (stack_count != 1) // Stack should have only one item left on it: the result. If not, it's a syntax error.
		goto fail; // This can happen with an expression consisting of only (), and definitely other ways too.

	ExprTokenType &result = stack[0];  // For performance and convenience.

	if (result.symbol == SYM_OPERAND) // The operand on the stack is generic (undetermined type).
	{
		// For this to happen, this operand must not have been involved with any operator, since
		// all operators resolve generic operands into specific operands.  If it wasn't involved
		// with any operator, that implies this expression contains no operators.  Although this
		// generally doesn't happen since the load-time routine marks expressions consisting of
		// only a single operand as non-expressions, it certainly does happen in these cases:
		// var:=+3, since the := operator does not have the legacy load-time logic used by other
		// expressions.  When the +3 arrives here, the unary plus is discarded as a non-operator,
		// leaving a raw 3 behind for evaluation here.
		// "Sleep Array%i%": This is still an exprssion even though there's only one operand
		// and no operators.
		// "Sleep VarContainingNumber": Load-time routine isn't quite smart enough to catch
		// this as a single-isolated deref (though it does catch "Sleep %VarContainingNumber%").
		switch (result.symbol = IsPureNumeric(result.marker, true, false, true)) // Assign.
		{
		case PURE_INTEGER:  result.value_int64 = ATOI64(result.marker); break;
		case PURE_FLOAT:    result.value_double = atof(result.marker); break;
		// Otherwise, symbol has been converted to SYM_STRING by the assignment in the switch expression,
		// and the other members of "result" are left unchanged.
		}
	}

	// Store the result of the expression in the deref buffer for the caller.  It is stored in the current
	// format in effect via SetFormat because:
	// 1) The := operator then doesn't have to convert to int/double then back to string to put the right format into effect.
	// 2) It might add a little bit of flexibility in places parameters where floating point values are expected
	//    (i.e. it allows a way to do automatic rounding), without giving up too much.  Changing floating point
	//    precision from the default of 6 decimal places is rare anyway, so as long as this behavior is documented,
	//    it seems okay for the moment.
	size_t result_size;
	switch (result.symbol)
	{
	case SYM_FLOAT:
		// In case of float formats that are too long to be supported, use snprint() to restrict the length.
		snprintf(aBuf_orig, MAX_FORMATTED_NUMBER_LENGTH + 1, g.FormatFloat, result.value_double); // %f probably defaults to %0.6f.  %f can handle doubles in MSVC++.
		break;
	case SYM_INTEGER:
		ITOA64(result.value_int64, aBuf_orig); // Store in hex or decimal format, as appropriate.
		break;
	case SYM_STRING: // And note that it can't be SYM_OPERAND because step right above would have remarked it as SYM_STRING.
		result_size = strlen(result.marker) + 1;
		memmove(aBuf_orig, result.marker, result_size);
		return aBuf_orig + result_size;
	default: // Result contains a non-operand symbol such as an operator.
		goto fail;
	}

	// It would have returned above if SYM_STRING, so this is SYM_FLOAT/SYM_INTEGER.  Calculate the length:
	return aBuf_orig + strlen(aBuf_orig) + 1;  // +1 because that's what callers want; i.e. the position after the terminator.

fail:
	*aBuf_orig++ = '\0'; // Increment because that's what callers want; i.e. the position after the terminator.
	return aBuf_orig;
}



ResultType Line::Deref(Var *aOutputVar, char *aBuf)
// Similar to ExpandArg(), except it parses and expands all variable references contained in aBuf.
{
	// This transient variable is used resolving environment variables that don't already exist
	// in the script's variable list (due to the fact that they aren't directly referenced elsewhere
	// in the script):
	char var_name[MAX_VAR_NAME_LENGTH + 1] = "";
	Var temp_var(var_name);
	Var *var;
	VarSizeType expanded_length;
	size_t var_name_length;
	char *cp, *cp1, *dest;

	// Do two passes:
	// #1: Calculate the space needed so that aOutputVar can be given more capacity if necessary.
	// #2: Expand the contents of aBuf into aOutputVar.

	for (int which_pass = 0; which_pass < 2; ++which_pass)
	{
		if (which_pass) // Start of second pass.
		{
			// Set up aOutputVar, enlarging it if necessary.  If it is of type VAR_CLIPBOARD,
			// this call will set up the clipboard for writing:
			if (aOutputVar->Assign(NULL, expanded_length) != OK)
				return FAIL;
			dest = aOutputVar->Contents();  // Init, and for performance.
		}
		else // First pass.
			expanded_length = 0; // Init prior to accumulation.

		for (cp = aBuf; ; ++cp)  // Increment to skip over the deref/escape just found by the inner for().
		{
			// Find the next escape char or deref symbol:
			for (; *cp && *cp != g_EscapeChar && *cp != g_DerefChar; ++cp)
			{
				if (which_pass) // 2nd pass
					*dest++ = *cp;  // Copy all non-variable-ref characters literally.
				else // just accumulate the length
					++expanded_length;
			}
			if (!*cp) // End of string while scanning/copying.  The current pass is now complete.
				break;
			if (*cp == g_EscapeChar)
			{
				if (which_pass) // 2nd pass
				{
					cp1 = cp + 1;
					switch (*cp1) // See ConvertEscapeSequences() for more details.
					{
						// Only lowercase is recognized for these:
						case 'a': *dest = '\a'; break;  // alert (bell) character
						case 'b': *dest = '\b'; break;  // backspace
						case 'f': *dest = '\f'; break;  // formfeed
						case 'n': *dest = '\n'; break;  // newline
						case 'r': *dest = '\r'; break;  // carriage return
						case 't': *dest = '\t'; break;  // horizontal tab
						case 'v': *dest = '\v'; break;  // vertical tab
						default:  *dest = *cp1; // Other characters are resolved just as they are.
					}
					++dest;
				}
				else
					++expanded_length;
				// Increment cp here and it will be incremented again by the outer loop, i.e. +2.
				// In other words, skip over the escape character, treating it and its target character
				// as a single character.
				++cp;
				continue;
			}
			// Otherwise, it's a dereference symbol, so calculate the size of that variable's contents
			// and add that to expanded_length (or copy the contents into aOutputVar if this is the
			// second pass).
			// Find the reference's ending symbol (don't bother with catching escaped deref chars here
			// -- e.g. %MyVar`% --  since it seems too troublesome to justify given how extremely rarely
			// it would be an issue):
			for (cp1 = cp + 1; *cp1 && *cp1 != g_DerefChar; ++cp1);
			if (!*cp1)    // Since end of string was found, this deref is not correctly terminated.
				continue; // For consistency, omit it entirely.
			var_name_length = cp1 - cp - 1;
			if (var_name_length && var_name_length <= MAX_VAR_NAME_LENGTH)
			{
				strlcpy(var_name, cp + 1, var_name_length + 1);  // +1 to convert var_name_length to size.
				if (   !(var = g_script.FindVar(var_name, var_name_length))   )
					// Variable doesn't exist, but since it might be an environment variable never referenced
					// directly elsewhere in the script, do special handling:
					var = &temp_var;  // Relies on the fact that var_temp.mName *is* the var_name pointer.
				// Don't allow the output variable to be read into itself this way because its contents
				if (var != aOutputVar)
				{
					if (which_pass) // 2nd pass
						dest += var->Get(dest);
					else // just accumulate the length
						expanded_length += var->Get(); // Add in the length of the variable's contents.
				}
			}
			// else since the variable name between the deref symbols is blank or too long: for consistency in behavior,
			// it seems best to omit the dereference entirely (don't put it into aOutputVar).
			cp = cp1; // For the next loop iteration, continue at the char after this reference's final deref symbol.
		} // for()
	} // for() (first and second passes)

	*dest = '\0';  // Terminate the output variable.
	aOutputVar->Length() = (VarSizeType)strlen(aOutputVar->Contents());  // Update to actual in case estimate was too large.
	return aOutputVar->Close();  // In case it's the clipboard.
}



VarSizeType Script::GetBatchLines(char *aBuf)
{
	// The BatchLine value can be either a numerical string or a string that ends in "ms".
	char buf[256];
	char *target_buf = aBuf ? aBuf : buf;
	if (g.IntervalBeforeRest >= 0) // Have this new method take precedence, if it's in use by the script.
		sprintf(target_buf, "%dms", g.IntervalBeforeRest); // Not snprintf().
	else
		ITOA64(g.LinesPerCycle, target_buf);
	return (VarSizeType)strlen(target_buf);
}

VarSizeType Script::GetTitleMatchMode(char *aBuf)
{
	if (!aBuf)
		return MAX_NUMBER_LENGTH;  // Just in case it's ever allowed to go beyond 3
	_itoa(g.TitleMatchMode, aBuf, 10);  // Always output as decimal vs. hex in this case.
	return (VarSizeType)strlen(aBuf);
}

VarSizeType Script::GetTitleMatchModeSpeed(char *aBuf)
{
	if (!aBuf)
		return 4;
	// For backward compatibility (due to StringCaseSense), never change the case used here:
	strcpy(aBuf, g.TitleFindFast ? "Fast" : "Slow");
	return 4;  // Always length 4
}

VarSizeType Script::GetDetectHiddenWindows(char *aBuf)
{
	if (!aBuf)
		return 3;  // Room for either On or Off
	// For backward compatibility (due to StringCaseSense), never change the case used here:
	strcpy(aBuf, g.DetectHiddenWindows ? "On" : "Off");
	return 3;
}

VarSizeType Script::GetDetectHiddenText(char *aBuf)
{
	if (!aBuf)
		return 3;  // Room for either On or Off
	// For backward compatibility (due to StringCaseSense), never change the case used here:
	strcpy(aBuf, g.DetectHiddenText ? "On" : "Off");
	return 3;
}

VarSizeType Script::GetAutoTrim(char *aBuf)
{
	if (!aBuf)
		return 3;  // Room for either On or Off
	// For backward compatibility (due to StringCaseSense), never change the case used here:
	strcpy(aBuf, g.AutoTrim ? "On" : "Off");
	return 3;
}

VarSizeType Script::GetStringCaseSense(char *aBuf)
{
	if (!aBuf)
		return 3;  // Room for either On or Off
	// For backward compatibility (due to StringCaseSense), never change the case used here:
	strcpy(aBuf, g.StringCaseSense ? "On" : "Off");
	return 3;
}

VarSizeType Script::GetFormatInteger(char *aBuf)
{
	if (!aBuf)
		return 1;
	// For backward compatibility (due to StringCaseSense), never change the case used here:
	*aBuf = g.FormatIntAsHex ? 'H' : 'D';
	*(aBuf + 1) = '\0';
	return 1;
}

VarSizeType Script::GetFormatFloat(char *aBuf)
{
	if (!aBuf)
		return (VarSizeType)strlen(g.FormatFloat);  // Include the extra chars since this is just an estimate.
	strlcpy(aBuf, g.FormatFloat + 1, strlen(g.FormatFloat + 1));   // Omit the leading % and the trailing 'f'.
	return (VarSizeType)strlen(aBuf);
}

VarSizeType Script::GetKeyDelay(char *aBuf)
{
	if (!aBuf)
		return MAX_NUMBER_LENGTH;
	_itoa(g.KeyDelay, aBuf, 10);  // Always output as decimal vs. hex in this case.
	return (VarSizeType)strlen(aBuf);
}

VarSizeType Script::GetWinDelay(char *aBuf)
{
	if (!aBuf)
		return MAX_NUMBER_LENGTH;
	_itoa(g.WinDelay, aBuf, 10);  // Always output as decimal vs. hex in this case.
	return (VarSizeType)strlen(aBuf);
}

VarSizeType Script::GetControlDelay(char *aBuf)
{
	if (!aBuf)
		return MAX_NUMBER_LENGTH;
	_itoa(g.ControlDelay, aBuf, 10);  // Always output as decimal vs. hex in this case.
	return (VarSizeType)strlen(aBuf);
}

VarSizeType Script::GetMouseDelay(char *aBuf)
{
	if (!aBuf)
		return MAX_NUMBER_LENGTH;
	_itoa(g.MouseDelay, aBuf, 10);  // Always output as decimal vs. hex in this case.
	return (VarSizeType)strlen(aBuf);
}

VarSizeType Script::GetDefaultMouseSpeed(char *aBuf)
{
	if (!aBuf)
		return MAX_NUMBER_LENGTH;  // Just in case it's ever allowed to go beyond 100.
	_itoa(g.DefaultMouseSpeed, aBuf, 10);  // Always output as decimal vs. hex in this case.
	return (VarSizeType)strlen(aBuf);
}



VarSizeType Script::GetIconHidden(char *aBuf)
{
	if (aBuf)
	{
		*aBuf = g_NoTrayIcon ? '1' : '0';
		*(aBuf + 1) = '\0';
	}
	return 1;  // Length is always 1.
}

VarSizeType Script::GetIconTip(char *aBuf)
{
	if (!aBuf)
		return g_script.mTrayIconTip ? (VarSizeType)strlen(g_script.mTrayIconTip) : 0;
	if (g_script.mTrayIconTip)
	{
		strcpy(aBuf, g_script.mTrayIconTip);
		return (VarSizeType)strlen(aBuf);
	}
	else
	{
		*aBuf = '\0';
		return 0;
	}
}

VarSizeType Script::GetIconFile(char *aBuf)
{
	if (!aBuf)
		return g_script.mCustomIconFile ? (VarSizeType)strlen(g_script.mCustomIconFile) : 0;
	if (g_script.mCustomIconFile)
	{
		strcpy(aBuf, g_script.mCustomIconFile);
		return (VarSizeType)strlen(aBuf);
	}
	else
	{
		*aBuf = '\0';
		return 0;
	}
}

VarSizeType Script::GetIconNumber(char *aBuf)
{
	if (!aBuf)
		return MAX_NUMBER_LENGTH;
	if (!g_script.mCustomIconNumber)
	{
		*aBuf = '\0';
		return 0;
	}
	UTOA(g_script.mCustomIconNumber, aBuf);
	return (VarSizeType)strlen(aBuf);
}



VarSizeType Script::GetExitReason(char *aBuf)
{
	char *str;
	switch(mExitReason)
	{
	case EXIT_LOGOFF: str = "Logoff"; break;
	case EXIT_SHUTDOWN: str = "Shutdown"; break;
	// Since the below are all relatively rare, except WM_CLOSE perhaps, they are all included
	// as one word to cut down on the number of possible words (it's easier to write OnExit
	// routines to cover all possibilities if there are fewer of them).
	case EXIT_WM_QUIT:
	case EXIT_CRITICAL:
	case EXIT_DESTROY:
	case EXIT_WM_CLOSE: str = "Close"; break;
	case EXIT_ERROR: str = "Error"; break;
	case EXIT_MENU: str = "Menu"; break;  // Standard menu, not a user-defined menu.
	case EXIT_EXIT: str = "Exit"; break;  // ExitApp or Exit command.
	case EXIT_RELOAD: str = "Reload"; break;
	case EXIT_SINGLEINSTANCE: str = "Single"; break;
	default:  // EXIT_NONE or unknown value (unknown would be considered a bug if it ever happened).
		str = "";
	}
	if (aBuf)
		strcpy(aBuf, str);
	return (VarSizeType)strlen(str);
}



VarSizeType Script::GetTimeIdlePhysical(char *aBuf)
// This is here rather than in script.h with the others because it depends on
// hotkey.h and globaldata.h, which can't be easily included in script.h due to
// mutual dependency issues.
{
	// If neither hook is active, default this to the same as the regular idle time:
	if (!Hotkey::HookIsActive())
		return GetTimeIdle(aBuf);
	if (!aBuf)
		return MAX_NUMBER_LENGTH;
	ITOA64(GetTickCount() - g_TimeLastInputPhysical, aBuf);
	return (VarSizeType)strlen(aBuf);
}



VarSizeType Script::GetSpace(VarTypeType aType, char *aBuf)
{
	if (!aBuf) return 1;  // i.e. the length of a single space char.
	*(aBuf++) = aType == VAR_SPACE ? ' ' : '\t';
	*aBuf = '\0';
	return 1;
}

VarSizeType Script::GetAhkVersion(char *aBuf)
{
	if (aBuf)
		strcpy(aBuf, NAME_VERSION);
	return (VarSizeType)strlen(NAME_VERSION);
}



// Confirmed: The below will all automatically use the local time (not UTC) when 3rd param is NULL.
VarSizeType Script::GetMMMM(char *aBuf)
{
	return (VarSizeType)(GetDateFormat(LOCALE_USER_DEFAULT, 0, NULL, "MMMM", aBuf, aBuf ? 999 : 0) - 1);
}

VarSizeType Script::GetMMM(char *aBuf)
{
	return (VarSizeType)(GetDateFormat(LOCALE_USER_DEFAULT, 0, NULL, "MMM", aBuf, aBuf ? 999 : 0) - 1);
}

VarSizeType Script::GetDDDD(char *aBuf)
{
	return (VarSizeType)(GetDateFormat(LOCALE_USER_DEFAULT, 0, NULL, "dddd", aBuf, aBuf ? 999 : 0) - 1);
}

VarSizeType Script::GetDDD(char *aBuf)
{
	return (VarSizeType)(GetDateFormat(LOCALE_USER_DEFAULT, 0, NULL, "ddd", aBuf, aBuf ? 999 : 0) - 1);
}



VarSizeType Script::MyGetTickCount(char *aBuf)
{
	// UPDATE: The below comments are now obsolete in light of having switched over to
	// using 64-bit integers (which aren't that much slower than 32-bit on 32-bit hardware):
	// Known limitation:
	// Although TickCount is an unsigned value, I'm not sure that our EnvSub command
	// will properly be able to compare two tick-counts if either value is larger than
	// INT_MAX.  So if the system has been up for more than about 25 days, there might be
	// problems if the user tries compare two tick-counts in the script using EnvSub.
	// UPDATE: It seems better to store all unsigned values as signed within script
	// variables.  Otherwise, when the var's value is next accessed and converted using
	// ATOI(), the outcome won't be as useful.  In other words, since the negative value
	// will be properly converted by ATOI(), comparing two negative tickcounts works
	// correctly (confirmed).  Even if one of them is negative and the other positive,
	// it will probably work correctly due to the nature of implicit unsigned math.
	// Thus, we use %d vs. %u in the snprintf() call below.
	if (!aBuf)
		return MAX_NUMBER_LENGTH; // Especially in this case, since tick might change between 1st & 2nd calls.
	ITOA64(GetTickCount(), aBuf);
	return (VarSizeType)strlen(aBuf);
}



VarSizeType Script::GetNow(char *aBuf)
{
	if (!aBuf)
		return DATE_FORMAT_LENGTH;
	SYSTEMTIME st;
	GetLocalTime(&st);
	SystemTimeToYYYYMMDD(aBuf, st);
	return (VarSizeType)strlen(aBuf);
}

VarSizeType Script::GetNowUTC(char *aBuf)
{
	if (!aBuf)
		return DATE_FORMAT_LENGTH;
	SYSTEMTIME st;
	GetSystemTime(&st);
	SystemTimeToYYYYMMDD(aBuf, st);
	return (VarSizeType)strlen(aBuf);
}



VarSizeType Script::GetOSType(char *aBuf)
{
	char *type = g_os.IsWinNT() ? "WIN32_NT" : "WIN32_WINDOWS";
	if (aBuf)
		strcpy(aBuf, type);
	return (VarSizeType)strlen(type); // Return length of type, not aBuf.
}

VarSizeType Script::GetOSVersion(char *aBuf)
// Adapted from AutoIt3 source.
{
	char *version;
	if (g_os.IsWinNT())
	{
		if (g_os.IsWinXP())
			version = "WIN_XP";
		else
		{
			if (g_os.IsWin2000())
				version = "WIN_2000";
			else
				version = "WIN_NT4";
		}
	}
	else
	{
		if (g_os.IsWin95())
			version = "WIN_95";
		else
		{
			if (g_os.IsWin98())
				version = "WIN_98";
			else
				version = "WIN_ME";
		}
	}
	if (aBuf)
		strcpy(aBuf, version);
	return (VarSizeType)strlen(version); // Always return length of version, not aBuf.
}

VarSizeType Script::GetLanguage(char *aBuf)
// Registry locations from J-Paul Mesnage.
{
	char buf[MAX_PATH];
	if (g_os.IsWinNT())  // NT/2k/XP+
	{
		if (g_os.IsWin2000orLater())
			RegReadString(HKEY_LOCAL_MACHINE, "SYSTEM\\CurrentControlSet\\Control\\Nls\\Language", "InstallLanguage", buf, MAX_PATH);
		else // NT4
			RegReadString(HKEY_LOCAL_MACHINE, "SYSTEM\\CurrentControlSet\\Control\\Nls\\Language", "Default", buf, MAX_PATH);
	}
	else // Win9x
	{
		RegReadString(HKEY_USERS, ".DEFAULT\\Control Panel\\Desktop\\ResourceLocale", "", buf, MAX_PATH);
		memmove(buf, buf + 4, strlen(buf + 4) + 1); // +1 to include the zero terminator.
	}
	if (aBuf)
		strcpy(aBuf, buf);
	return (VarSizeType)strlen(buf);
}

VarSizeType Script::GetUserOrComputer(bool aGetUser, char *aBuf)
{
	char buf[MAX_PATH];  // Doesn't use MAX_COMPUTERNAME_LENGTH + 1 in case longer names are allowed in the future.
	DWORD buf_size = MAX_PATH;
	if (   !(aGetUser ? GetUserName(buf, &buf_size) : GetComputerName(buf, &buf_size))   )
		*buf = '\0';
	if (aBuf)
		strcpy(aBuf, buf);
	return (VarSizeType)strlen(buf);
}



VarSizeType Script::GetProgramFiles(char *aBuf)
{
	char buf[MAX_PATH];
	RegReadString(HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion", "ProgramFilesDir", buf, MAX_PATH);
	if (aBuf)
		strcpy(aBuf, buf);
	return (VarSizeType)strlen(buf);
}

VarSizeType Script::GetDesktop(bool aGetCommon, char *aBuf)
{
	char buf[MAX_PATH] = "";
	if (aGetCommon)
		RegReadString(HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders", "Common Desktop", buf, MAX_PATH);
	if (!*buf) // Either the above failed or we were told to get the user/private dir instead.
		RegReadString(HKEY_CURRENT_USER, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders", "Desktop", buf, MAX_PATH);
	if (aBuf)
		strcpy(aBuf, buf);
	return (VarSizeType)strlen(buf);
}

VarSizeType Script::GetStartMenu(bool aGetCommon, char *aBuf)
{
	char buf[MAX_PATH] = "";
	if (aGetCommon)
		RegReadString(HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders", "Common Start Menu", buf, MAX_PATH);
	if (!*buf) // Either the above failed or we were told to get the user/private dir instead.
		RegReadString(HKEY_CURRENT_USER, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders", "Start Menu", buf, MAX_PATH);
	if (aBuf)
		strcpy(aBuf, buf);
	return (VarSizeType)strlen(buf);
}

VarSizeType Script::GetPrograms(bool aGetCommon, char *aBuf)
{
	char buf[MAX_PATH] = "";
	if (aGetCommon)
		RegReadString(HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders", "Common Programs", buf, MAX_PATH);
	if (!*buf) // Either the above failed or we were told to get the user/private dir instead.
		RegReadString(HKEY_CURRENT_USER, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders", "Programs", buf, MAX_PATH);
	if (aBuf)
		strcpy(aBuf, buf);
	return (VarSizeType)strlen(buf);
}

VarSizeType Script::GetStartup(bool aGetCommon, char *aBuf)
{
	char buf[MAX_PATH] = "";
	if (aGetCommon)
		RegReadString(HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders", "Common Startup", buf, MAX_PATH);
	if (!*buf) // Either the above failed or we were told to get the user/private dir instead.
		RegReadString(HKEY_CURRENT_USER, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders", "Startup", buf, MAX_PATH);
	if (aBuf)
		strcpy(aBuf, buf);
	return (VarSizeType)strlen(buf);
}

VarSizeType Script::GetMyDocuments(char *aBuf)
{
	char buf[MAX_PATH];
	RegReadString(HKEY_CURRENT_USER, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders", "Personal", buf, MAX_PATH);
	// Since it is common (such as in networked environments) to have My Documents on the root of a drive
	// (such as a mapped drive letter), remove the backslash from something like M:\ because M: is more
	// appropriate for most uses:
	Line::Util_StripTrailingDir(buf);
	if (aBuf)
		strcpy(aBuf, buf);
	return (VarSizeType)strlen(buf);
}



VarSizeType Script::GetIsAdmin(char *aBuf)
// Adapted from AutoIt3 source.
{
	if (!aBuf)
		return 1;  // The length of the string "1" or "0".
	char result = '0';  // Default.
	if (g_os.IsWin9x())
		result = '1';
	else
	{
		SC_HANDLE h = OpenSCManager(NULL, NULL, SC_MANAGER_LOCK);
		if (h)
		{
			SC_LOCK lock = LockServiceDatabase(h);
			if (lock)
			{
				UnlockServiceDatabase(lock);
				result = '1';
			}
			else
			{
				DWORD lastErr = GetLastError();
				if (lastErr == ERROR_SERVICE_DATABASE_LOCKED)
					result = '1';
			}
			CloseServiceHandle(h);
		}
	}
	aBuf[0] = result;
	aBuf[1] = '\0';
	return 1; // Length of aBuf.
}



VarSizeType Script::ScriptGetCursor(char *aBuf)
// Adapted from the AutoIt3 source.
{
	if (!aBuf)
		return SMALL_STRING_LENGTH;  // we're returning the length of the var's contents, not the size.

	POINT point;
	GetCursorPos(&point);
	HWND target_window = WindowFromPoint(point);

	// MSDN docs imply that threads must be attached for GetCursor() to work.
	// A side-effect of attaching threads or of GetCursor() itself is that mouse double-clicks
	// are interfered with, at least if this function is called repeatedly at a high frequency.
	ATTACH_THREAD_INPUT
	HCURSOR current_cursor = GetCursor();
	DETACH_THREAD_INPUT

	if (!current_cursor)
	{
		#define CURSOR_UNKNOWN "Unknown"
		strlcpy(aBuf, CURSOR_UNKNOWN, SMALL_STRING_LENGTH + 1);
		return (VarSizeType)strlen(aBuf);
	}

	// Static so that it's initialized on first use (should help performance after the first time):
	static HCURSOR cursor[] = {LoadCursor(0,IDC_APPSTARTING), LoadCursor(0,IDC_ARROW), LoadCursor(0,IDC_CROSS)
		, LoadCursor(0,IDC_HELP), LoadCursor(0,IDC_IBEAM), LoadCursor(0,IDC_ICON), LoadCursor(0,IDC_NO)
		, LoadCursor(0,IDC_SIZE), LoadCursor(0,IDC_SIZEALL), LoadCursor(0,IDC_SIZENESW), LoadCursor(0,IDC_SIZENS)
		, LoadCursor(0,IDC_SIZENWSE), LoadCursor(0,IDC_SIZEWE), LoadCursor(0,IDC_UPARROW), LoadCursor(0,IDC_WAIT)};
	// The order in the below array must correspond to the order in the above array:
	static char *cursor_name[] = {"AppStarting", "Arrow", "Cross"
		, "Help", "IBeam", "Icon", "No"
		, "Size", "SizeAll", "SizeNESW", "SizeNS"  // NESW = NorthEast or SouthWest
		, "SizeNWSE", "SizeWE", "UpArrow", "Wait", CURSOR_UNKNOWN};  // The last item is used to mark end-of-array.
	static int cursor_count = sizeof(cursor) / sizeof(HCURSOR);

	int a;
	for (a = 0; a < cursor_count; ++a)
		if (cursor[a] == current_cursor)
			break;

	strlcpy(aBuf, cursor_name[a], SMALL_STRING_LENGTH + 1);  // If a is out-of-bounds, "Unknown" will be used.
	return (VarSizeType)strlen(aBuf);
}



VarSizeType Script::ScriptGetCaret(VarTypeType aVarType, char *aBuf)
{
	if (!aBuf)
		return MAX_NUMBER_LENGTH;

	// These static variables are used to keep the X and Y coordinates in sync with each other, as a snapshot
	// of where the caret was at one precise instant in time.  This is because the X and Y vars are resolved
	// separately by the script, and due to split second timing, they might otherwise not be accurate with
	// respect to each other.  This method also helps performance since it avoids unnecessary calls to
	// ATTACH_THREAD_INPUT.
	static HWND sForeWinPrev = NULL;
	static DWORD sTimestamp = GetTickCount();
	static POINT sPoint;
	static BOOL sResult;

	// I believe only the foreground window can have a caret position due to relationship with focused control.
	HWND target_window = GetForegroundWindow(); // Variable must be named target_window for ATTACH_THREAD_INPUT.
	if (!target_window) // No window is in the foreground, report blank coordinate.
	{
		*aBuf = '\0';
		return 0;
	}

	DWORD now_tick = GetTickCount();

	if (target_window != sForeWinPrev || now_tick - sTimestamp > 5) // Different window or too much time has passed.
	{
		// Otherwise:
		ATTACH_THREAD_INPUT  // au3: Doesn't work without attaching.
		sResult = GetCaretPos(&sPoint);
		HWND focused_control = GetFocus();  // Also relies on threads being attached.
		DETACH_THREAD_INPUT
		if (!sResult)
		{
			*aBuf = '\0';
			return 0;
		}
		ClientToScreen(focused_control ? focused_control : target_window, &sPoint);
		if (!(g.CoordMode & COORD_MODE_CARET))  // Using the default, which is coordinates relative to window.
		{
			// Convert screen coordinates to window coordinates:
			RECT rect;
			GetWindowRect(target_window, &rect);
			sPoint.x -= rect.left;
			sPoint.y -= rect.top;
		}
		// Now that all failure conditions have been checked, update static variables for the next caller:
		sForeWinPrev = target_window;
		sTimestamp = now_tick;
	}
	else // Same window and recent enough, but did prior call fail?  If so, provide a blank result like the prior.
	{
		if (!sResult)
		{
			*aBuf = '\0';
			return 0;
		}
	}
	// Now the above has ensured that sPoint contains valid coordinates that are up-to-date enough to be used.
	_itoa(aVarType == VAR_CARETX ? sPoint.x : sPoint.y, aBuf, 10);  // Always output as decimal vs. hex in this case.
	return (VarSizeType)strlen(aBuf);
}



VarSizeType Script::GetScreenWidth(char *aBuf)
{
	if (!aBuf)
		return MAX_NUMBER_LENGTH;
	ITOA(GetSystemMetrics(SM_CXSCREEN), aBuf);
	return (VarSizeType)strlen(aBuf);
}

VarSizeType Script::GetScreenHeight(char *aBuf)
{
	if (!aBuf)
		return MAX_NUMBER_LENGTH;
	ITOA(GetSystemMetrics(SM_CYSCREEN), aBuf);
	return (VarSizeType)strlen(aBuf);
}



VarSizeType Script::GetIP(int aAdapterIndex, char *aBuf)
// Adapted from the AutoIt3 source.
{
	// aaa.bbb.ccc.ddd = 15, but allow room for larger IP's in the future.
	#define IP_ADDRESS_SIZE 32 // The maximum size of any of the strings we return, including terminator.
	if (!aBuf)
		return IP_ADDRESS_SIZE - 1;  // -1 since we're returning the length of the var's contents, not the size.

	WSADATA wsadata;
	if (WSAStartup(MAKEWORD(1, 1), &wsadata)) // Failed (it returns 0 on success).
	{
		*aBuf = '\0';
		return 0;
	}

	char host_name[256];
	gethostname(host_name, sizeof(host_name));
	HOSTENT *lpHost = gethostbyname(host_name);

	// au3: How many adapters have we?
	int adapter_count = 0;
	while (lpHost->h_addr_list[adapter_count])
		++adapter_count;

	if (aAdapterIndex >= adapter_count)
		strlcpy(aBuf, "0.0.0.0", IP_ADDRESS_SIZE);
	else
	{
		IN_ADDR inaddr;
		memcpy(&inaddr, lpHost->h_addr_list[aAdapterIndex], 4);
		strlcpy(aBuf, (char *)inet_ntoa(inaddr), IP_ADDRESS_SIZE);
	}

	WSACleanup();
	return (VarSizeType)strlen(aBuf);
}



VarSizeType Script::GetFilename(char *aBuf)
{
	if (aBuf)
		strcpy(aBuf, mFileName);
	return (VarSizeType)strlen(mFileName);
}

VarSizeType Script::GetFileDir(char *aBuf)
{
	char str[MAX_PATH + 1] = "";  // Set default.  Uses +1 to append final backslash for AutoIt2 (.aut) scripts.
	strlcpy(str, mFileDir, sizeof(str));
	size_t length = strlen(str); // Needed not just for AutoIt2.
	// If it doesn't already have a final backslash, namely due to it being a root directory,
	// provide one so that it is backward compatible with AutoIt v2:
	if (mIsAutoIt2 && length && str[length - 1] != '\\')
	{
		str[length++] = '\\';
		str[length] = '\0';
	}
	if (aBuf)
		strcpy(aBuf, str);
	return (VarSizeType)length;
}

VarSizeType Script::GetFilespec(char *aBuf)
{
	if (aBuf)
		sprintf(aBuf, "%s\\%s", mFileDir, mFileName);
	return (VarSizeType)(strlen(mFileDir) + strlen(mFileName) + 1);
}



VarSizeType Script::GetLoopFileName(char *aBuf)
{
	char *str = "";  // Set default.
	if (mLoopFile)
	{
		// The loop handler already prepended the script's directory in here for us:
		if (str = strrchr(mLoopFile->cFileName, '\\'))
			++str;
		else // No backslash, so just make it the entire file name.
			str = mLoopFile->cFileName;
	}
	if (aBuf)
		strcpy(aBuf, str);
	return (VarSizeType)strlen(str);
}

VarSizeType Script::GetLoopFileShortName(char *aBuf)
{
	char *str = "";  // Set default.
	if (mLoopFile)
	{
		if (   !*(str = mLoopFile->cAlternateFileName)   )
			// Files whose long name is shorter than the 8.3 usually don't have value stored here,
			// so use the long name whenever a short name is unavailable for any reason (could
			// also happen if NTFS has short-name generation disabled?)
			return GetLoopFileName(aBuf);
	}
	if (aBuf)
		strcpy(aBuf, str);
	return (VarSizeType)strlen(str);
}

VarSizeType Script::GetLoopFileDir(char *aBuf)
{
	char *str = "";  // Set default.
	char *last_backslash = NULL;
	if (mLoopFile)
	{
		// The loop handler already prepended the script's directory in here for us.
		// But if the loop had a relative path in its FilePattern, there might be
		// only a relative directory here, or no directory at all if the current
		// file is in the origin/root dir of the search:
		if (last_backslash = strrchr(mLoopFile->cFileName, '\\'))
		{
			*last_backslash = '\0'; // Temporarily terminate.
			str = mLoopFile->cFileName;
		}
		else // No backslash, so there is no directory in this case.
			str = "";
	}
	VarSizeType length = (VarSizeType)strlen(str);
	if (!aBuf)
	{
		if (last_backslash)
			*last_backslash = '\\';  // Restore the orginal value.
		return length;
	}
	strcpy(aBuf, str);
	if (last_backslash)
		*last_backslash = '\\';  // Restore the orginal value.
	return length;
}

VarSizeType Script::GetLoopFileFullPath(char *aBuf)
{
	char *str = "";  // Set default.
	if (mLoopFile)
		// The loop handler already prepended the script's directory in here for us:
		str = mLoopFile->cFileName;
	if (aBuf)
		strcpy(aBuf, str);
	return (VarSizeType)strlen(str);
}

VarSizeType Script::GetLoopFileShortPath(char *aBuf)
// Unlike GetLoopFileShortName(), this function returns blank when there is no short path.
// This is done so that there's a way for the script to more easily tell the difference between
// an 8.3 name not being available (due to the being disabled in the registry) and the short
// name simply being the same as the long name.  For example, if short name creation is disabled
// in the registry, A_LoopFileShortName would contain the long name instead, as documented.
// But to detect if that short name is really a long name, A_LoopFileShortPath could be checked
// and if it's blank, there is no short name available.
{
	char buf[MAX_PATH] = "";  // Set default.
	DWORD length = 0;         // Set default.
	if (mLoopFile)
		// The loop handler already prepended the script's directory in cFileName for us:
		if (   !(length = GetShortPathName(mLoopFile->cFileName, buf, sizeof(buf)))   )
			*buf = '\0'; // It might fail if NtfsDisable8dot3NameCreation is turned on in the registry, and possibly for other reasons.
	if (aBuf)
		strcpy(aBuf, buf);
	return (VarSizeType)length;
}

VarSizeType Script::GetLoopFileTimeModified(char *aBuf)
{
	char str[64] = "";  // Set default.
	if (mLoopFile)
		FileTimeToYYYYMMDD(str, mLoopFile->ftLastWriteTime, true);
	if (aBuf)
		strcpy(aBuf, str);
	return (VarSizeType)strlen(str);
}

VarSizeType Script::GetLoopFileTimeCreated(char *aBuf)
{
	char str[64] = "";  // Set default.
	if (mLoopFile)
		FileTimeToYYYYMMDD(str, mLoopFile->ftCreationTime, true);
	if (aBuf)
		strcpy(aBuf, str);
	return (VarSizeType)strlen(str);
}

VarSizeType Script::GetLoopFileTimeAccessed(char *aBuf)
{
	char str[64] = "";  // Set default.
	if (mLoopFile)
		FileTimeToYYYYMMDD(str, mLoopFile->ftLastAccessTime, true);
	if (aBuf)
		strcpy(aBuf, str);
	return (VarSizeType)strlen(str);
}

VarSizeType Script::GetLoopFileAttrib(char *aBuf)
{
	char str[64] = "";  // Set default.
	if (mLoopFile)
		FileAttribToStr(str, mLoopFile->dwFileAttributes);
	if (aBuf)
		strcpy(aBuf, str);
	return (VarSizeType)strlen(str);
}

VarSizeType Script::GetLoopFileSize(char *aBuf, int aDivider)
{
	// Don't use MAX_NUMBER_LENGTH in case user has selected a very long float format via SetFormat.
	char str[128];
	char *target_buf = aBuf ? aBuf : str;
	*target_buf = '\0';  // Set default.
	if (mLoopFile)
	{
		// It's a documented limitation that the size will show as negative if
		// greater than 2 gig, and will be wrong if greater than 4 gig.  For files
		// that large, scripts should use the KB version of this function instead.
		// If a file is over 4gig, set the value to be the maximum size (-1 when
		// expressed as a signed integer, since script variables are based entirely
		// on 32-bit signed integers due to the use of ATOI(), etc.).  UPDATE: 64-bit
		// ints are now standard, so the above is unnecessary:
		//sprintf(str, "%d%", mLoopFile->nFileSizeHigh ? -1 : (int)mLoopFile->nFileSizeLow);
		ULARGE_INTEGER ul;
		ul.HighPart = mLoopFile->nFileSizeHigh;
		ul.LowPart = mLoopFile->nFileSizeLow;
		ITOA64((__int64)(aDivider ? ((unsigned __int64)ul.QuadPart / aDivider) : ul.QuadPart), target_buf);
	}
	return (VarSizeType)strlen(target_buf);
}

VarSizeType Script::GetLoopRegType(char *aBuf)
{
	char str[256] = "";  // Set default.
	if (mLoopRegItem)
		Line::RegConvertValueType(str, sizeof(str), mLoopRegItem->type);
	if (aBuf)
		strcpy(aBuf, str);
	return (VarSizeType)strlen(str);
}

VarSizeType Script::GetLoopRegKey(char *aBuf)
{
	char str[256] = "";  // Set default.
	if (mLoopRegItem)
		// Use root_key_type, not root_key (which might be a remote vs. local HKEY):
		Line::RegConvertRootKey(str, sizeof(str), mLoopRegItem->root_key_type);
	if (aBuf)
		strcpy(aBuf, str);
	return (VarSizeType)strlen(str);
}

VarSizeType Script::GetLoopRegSubKey(char *aBuf)
{
	char *str = "";  // Set default.
	if (mLoopRegItem)
		str = mLoopRegItem->subkey;
	if (aBuf)
		strcpy(aBuf, str);
	return (VarSizeType)strlen(str);
}

VarSizeType Script::GetLoopRegName(char *aBuf)
{
	char *str = "";  // Set default.
	if (mLoopRegItem)
		str = mLoopRegItem->name; // This can be either the name of a subkey or the name of a value.
	if (aBuf)
		strcpy(aBuf, str);
	return (VarSizeType)strlen(str);
}

VarSizeType Script::GetLoopRegTimeModified(char *aBuf)
{
	char str[64] = "";  // Set default.
	// Only subkeys (not values) have a time.  In addition, Win9x doesn't support retrieval
	// of the time (nor does it store it), so make the var blank in that case:
	if (mLoopRegItem && mLoopRegItem->type == REG_SUBKEY && !g_os.IsWin9x())
		FileTimeToYYYYMMDD(str, mLoopRegItem->ftLastWriteTime, true);
	if (aBuf)
		strcpy(aBuf, str);
	return (VarSizeType)strlen(str);
}

VarSizeType Script::GetLoopReadLine(char *aBuf)
{
	char *str = "";  // Set default.
	if (mLoopReadFile)
		str = mLoopReadFile->mCurrentLine;
	if (aBuf)
		strcpy(aBuf, str);
	return (VarSizeType)strlen(str);
}

VarSizeType Script::GetLoopField(char *aBuf)
{
	char *str = "";  // Set default.
	if (mLoopField)
		str = mLoopField;
	if (aBuf)
		strcpy(aBuf, str);
	return (VarSizeType)strlen(str);
}

VarSizeType Script::GetLoopIndex(char *aBuf)
{
	if (!aBuf)
		return MAX_NUMBER_LENGTH;
	ITOA64(mLoopIteration, aBuf);
	return (VarSizeType)strlen(aBuf);
}



VarSizeType Script::GetThisMenuItem(char *aBuf)
{
	if (aBuf)
		strcpy(aBuf, mThisMenuItemName);
	return (VarSizeType)strlen(mThisMenuItemName);
}

VarSizeType Script::GetThisMenuItemPos(char *aBuf)
{
	if (!aBuf)
		return MAX_NUMBER_LENGTH;
	// The menu item's position is discovered through this process -- rather than doing
	// something higher performance such as storing the menu handle or pointer to menu/item
	// object in g_script -- because those things tend to be volatile.  For example, a menu
	// or menu item object might be destroyed between the time the user selects it and the
	// time this variable is referenced in the script.  Thus, by definition, this variable
	// contains the CURRENT position of the most recently selected menu item within its
	// CURRENT menu.
	if (*mThisMenuName && *mThisMenuItemName)
	{
		UserMenu *menu = FindMenu(mThisMenuName);
		if (menu)
		{
			// If the menu does not physically exist yet (perhaps due to being destroyed as a result
			// of DeleteAll, Delete, or some other operation), create it so that the position of the
			// item can be determined.  This is done for consistency in behavior.
			if (!menu->mMenu)
				menu->Create();
			UINT menu_item_pos = menu->GetItemPos(mThisMenuItemName);
			if (menu_item_pos < UINT_MAX) // Success
			{
				UTOA(menu_item_pos + 1, aBuf);  // Add one to convert from zero-based to 1-based.
				return (VarSizeType)strlen(aBuf);
			}
		}
	}
	// Otherwise:
	*aBuf = '\0';
	return 0;
}

VarSizeType Script::GetThisMenu(char *aBuf)
{
	if (aBuf)
		strcpy(aBuf, mThisMenuName);
	return (VarSizeType)strlen(mThisMenuName);
}

VarSizeType Script::GetThisHotkey(char *aBuf)
{
	if (aBuf)
		strcpy(aBuf, mThisHotkeyName);
	return (VarSizeType)strlen(mThisHotkeyName);
}

VarSizeType Script::GetPriorHotkey(char *aBuf)
{
	if (aBuf)
		strcpy(aBuf, mPriorHotkeyName);
	return (VarSizeType)strlen(mPriorHotkeyName);
}

VarSizeType Script::GetTimeSinceThisHotkey(char *aBuf)
{
	if (!aBuf)
		return MAX_NUMBER_LENGTH;
	// It must be the type of hotkey that has a label because we want the TimeSinceThisHotkey
	// value to be "in sync" with the value of ThisHotkey itself (i.e. use the same method
	// to determine which hotkey is the "this" hotkey):
	if (*mThisHotkeyName)
		// Even if GetTickCount()'s TickCount has wrapped around to zero and the timestamp hasn't,
		// DWORD math still gives the right answer as long as the number of days between
		// isn't greater than about 49.  See MyGetTickCount() for explanation of %d vs. %u.
		// Update: Using 64-bit ints now, so above is obsolete:
		//snprintf(str, sizeof(str), "%d", (DWORD)(GetTickCount() - mThisHotkeyStartTime));
		ITOA64((__int64)(GetTickCount() - mThisHotkeyStartTime), aBuf)  // No semicolon
	else
		strcpy(aBuf, "-1");
	return (VarSizeType)strlen(aBuf);
}

VarSizeType Script::GetTimeSincePriorHotkey(char *aBuf)
{
	if (!aBuf)
		return MAX_NUMBER_LENGTH;
	if (*mPriorHotkeyName)
		// See MyGetTickCount() for explanation for explanation:
		//snprintf(str, sizeof(str), "%d", (DWORD)(GetTickCount() - mPriorHotkeyStartTime));
		ITOA64((__int64)(GetTickCount() - mPriorHotkeyStartTime), aBuf)
	else
		strcpy(aBuf, "-1");
	return (VarSizeType)strlen(aBuf);
}

VarSizeType Script::GetEndChar(char *aBuf)
{
	if (!aBuf)
		return 1;
	*aBuf = mEndChar;
	*(aBuf + 1) = '\0';
	return 1;
}



VarSizeType Script::GetGui(VarTypeType aVarType, char *aBuf)
// We're returning the length of the var's contents, not the size.
{
	if (g.GuiWindowIndex >= MAX_GUI_WINDOWS) // The current thread was not launched as a result of GUI action.
	{
		if (aBuf)
			*aBuf = '\0';
		return 0;
	}

	GuiType *pgui;
	char buf[MAX_NUMBER_LENGTH + 1];

	switch (aVarType)
	{
	case VAR_GUI:
		_itoa(g.GuiWindowIndex + 1, buf, 10);  // Always stored as decimal vs. hex, regardless of script settings.
		break;
	case VAR_GUIWIDTH:
	case VAR_GUIHEIGHT:
		if (   !(pgui = g_gui[g.GuiWindowIndex])   ) // Gui window doesn't currently exist, so return a blank.
		{
			if (aBuf)
				*aBuf = '\0';
			return 0;
		}
		// Otherwise:
		_itoa(aVarType == VAR_GUIWIDTH ? LOWORD(pgui->mSizeWidthHeight) : HIWORD(pgui->mSizeWidthHeight), buf, 10);
		// Above is always stored as decimal vs. hex, regardless of script settings.
		break;
	}
	if (aBuf)
		strcpy(aBuf, buf);
	return (VarSizeType)strlen(buf);
}



VarSizeType Script::GetGuiControl(char *aBuf)
// We're returning the length of the var's contents, not the size.
{
	GuiType *pgui;
	// Note that other logic ensures that g.GuiControlIndex is out-of-bounds whenever g.GuiWindowIndex is.
	// That is why g.GuiWindowIndex is not checked to make sure it's less than MAX_GUI_WINDOWS.
	// Relies on short-circuit boolean order:
	if (g.GuiControlIndex >= MAX_CONTROLS_PER_GUI // Must check this first due to short-circuit boolean.  A non-GUI thread or one triggered by GuiClose/Escape or Gui menu bar.
		|| !(pgui = g_gui[g.GuiWindowIndex]) // Gui Window no longer exists.
		|| g.GuiControlIndex >= pgui->mControlCount) // Gui control no longer exists, perhaps because window was destroyed and recreated with fewer controls.
	{
		if (aBuf)
			*aBuf = '\0';
		return 0;
	}
	GuiControlType &control = pgui->mControl[g.GuiControlIndex]; // For performance and convenience.
    if (aBuf)
	{
		// Caller has already ensured aBuf is large enough.
		if (control.output_var)
			return (VarSizeType)strlen(strcpy(aBuf, control.output_var->mName));
		else // Fall back to getting the leading characters of its caption (most often used for buttons).
			#define A_GUICONTROL_TEXT_LENGTH (MAX_ALLOC_SIMPLE - 1)
			return GetWindowText(control.hwnd, aBuf, A_GUICONTROL_TEXT_LENGTH + 1); // +1 is verified correct.
	}
	// Otherwise, just return the length:
	if (control.output_var)
		return (VarSizeType)strlen(control.output_var->mName);
	// Otherwise: Fall back to getting the leading characters of its caption (most often used for buttons)
	VarSizeType length = GetWindowTextLength(control.hwnd);
	return (length > A_GUICONTROL_TEXT_LENGTH) ? A_GUICONTROL_TEXT_LENGTH : length;
}



VarSizeType Script::GetGuiControlEvent(char *aBuf)
// We're returning the length of the var's contents, not the size.
{
	if (g.GuiEvent == GUI_EVENT_DROPFILES)
	{
		GuiType *pgui;
		UINT u, file_count;
		// GUI_EVENT_DROPFILES should mean that g.GuiWindowIndex < MAX_GUI_WINDOWS, but the below will double check
		// that in case g.GuiEvent can ever be set to that value as a result of receiving a bogus message in the queue.
		if (g.GuiWindowIndex >= MAX_GUI_WINDOWS  // The current thread was not launched as a result of GUI action or this is a bogus msg.
			|| !(pgui = g_gui[g.GuiWindowIndex]) // Gui window no longer exists.  Relies on short-circuit boolean.
			|| !pgui->mHdrop // No HDROP (probably impossible unless g.GuiEvent was given a bogus value somehow).
			|| !(file_count = DragQueryFile(pgui->mHdrop, 0xFFFFFFFF, NULL, 0))) // No files in the drop (not sure if this is possible).
			// All of the above rely on short-circuit boolean order.
		{
			// Make the dropped-files list blank since there is no HDROP to query (or no files in it).
			if (aBuf)
				*aBuf = '\0';
			return 0;
		}
		// Above has ensured that file_count > 0
		if (aBuf)
		{
			char *cp = aBuf;
			for (u = 0; u < file_count; ++u)
			{
				cp += DragQueryFile(pgui->mHdrop, u, cp, MAX_PATH); // MAX_PATH is arbitrary since aBuf is already known to be large enough.
				if (u < file_count - 1) // i.e omit the LF after the last file to make parsing via "Loop, Parse" easier.
					*cp++ = '\n';
				// Although the transcription of files on the clipboard into their text filenames is done
				// with \r\n (so that they're in the right format to be pasted to other apps as a plain text
				// list), it seems best to use a plain linefeed for dropped files since they won't be going
				// onto the clipboard nearly as often, and `n is easier to parse.  Also, a script array isn't
				// used because large file lists would then consume a lot more of memory because arrays
				// are permanent once created, and also there would be wasted space due to the part of each
				// variable's capacity not used by the filename.
			}
			// No need for final termination of string because the last item lacks a newline.
			return (VarSizeType)(cp - aBuf); // This is the length of what's in the buffer.
		}
		else
		{
			VarSizeType total_length = 0;
			for (u = 0; u < file_count; ++u)
				total_length += DragQueryFile(pgui->mHdrop, u, NULL, 0);
				// Above: MSDN: "If the lpszFile buffer address is NULL, the return value is the required size,
				// in characters, of the buffer, not including the terminating null character."
			return total_length + file_count - 1; // Include space for a linefeed after each filename except the last.
		}
		// Don't call DragFinish() because this variable might be referred to again before this thread
		// is done.  DragFinish() is called by MsgSleep() when the current thread finishes.
	}

	// Otherwise, this event is not GUI_EVENT_DROPFILES, so use standard modes of operation.
	static char *names[] = GUI_EVENT_NAMES;
	if (!aBuf)
		return (g.GuiEvent < GUI_EVENT_ILLEGAL) ? (VarSizeType)strlen(names[g.GuiEvent]) : 1;
	// Otherwise:
	if (g.GuiEvent < GUI_EVENT_ILLEGAL)
	{
		strcpy(aBuf, names[g.GuiEvent]);
		return (VarSizeType)strlen(aBuf);
	}
	else // g.GuiEvent is assumed to be an ASCII value, such as a digit.  This supports Slider controls.
	{
		*aBuf++ = (char)(UCHAR)g.GuiEvent;
		*aBuf = '\0';
		return 1;
	}
}



VarSizeType Script::GetTimeIdle(char *aBuf)
{
	if (!aBuf)
		return MAX_NUMBER_LENGTH;
	*aBuf = '\0';  // Set default.
	if (g_os.IsWin2000orLater()) // Checked in case the function is present but "not implemented".
	{
		// Must fetch it at runtime, otherwise the program can't even be launched on Win9x/NT:
		typedef BOOL (WINAPI *MyGetLastInputInfoType)(PLASTINPUTINFO);
		static MyGetLastInputInfoType MyGetLastInputInfo = (MyGetLastInputInfoType)
			GetProcAddress(GetModuleHandle("User32.dll"), "GetLastInputInfo");
		if (MyGetLastInputInfo)
		{
			LASTINPUTINFO lii;
			lii.cbSize = sizeof(lii);
			if (MyGetLastInputInfo(&lii))
				ITOA64(GetTickCount() - lii.dwTime, aBuf);
		}
	}
	return (VarSizeType)strlen(aBuf);
}



char *Line::LogToText(char *aBuf, size_t aBufSize)
// Static method:
// Translates sLog into its text equivalent, putting the result into aBuf and
// returning the position in aBuf of its new string terminator.
{
	if (!aBuf || aBufSize < 256) return NULL;
	char *aBuf_orig = aBuf;
	//snprintf(aBuf, BUF_SPACE_REMAINING, "Deref buffer size: %u\n\n", sDerefBufSize);
	//aBuf += strlen(aBuf);
	snprintf(aBuf, BUF_SPACE_REMAINING, "Script lines most recently executed (oldest first).  Press [F5] to refresh.\r\n"
		"The seconds elapsed between a line and the one after it is in parens to the right (if not 0).\r\n"
		"The bottommost line's elapsed time is the number of seconds since it executed.\r\n\r\n");
	aBuf += strlen(aBuf);
	// Start at the oldest and continue up through the newest:
	int i, line_index;
	DWORD elapsed;
	size_t space_remaining;
	for (i = 0, line_index = sLogNext; i < LINE_LOG_SIZE; ++i, ++line_index)
	{
		if (line_index >= LINE_LOG_SIZE) // wrap around
			line_index = 0;
		if (sLog[line_index] == NULL)
			continue;
		if (i + 1 < LINE_LOG_SIZE)  // There are still more lines to be processed
			elapsed = sLogTick[line_index + 1 >= LINE_LOG_SIZE ? 0 : line_index + 1] - sLogTick[line_index];
		else // This is the line, so compare it's time against the current time instead.
			elapsed = GetTickCount() - sLogTick[line_index];
		space_remaining = BUF_SPACE_REMAINING;  // Resolve macro only once for performance.
		// Truncate really huge lines so that the Edit control's size is less likely to be exceeded:
		aBuf = sLog[line_index]->ToText(aBuf, space_remaining < 500 ? space_remaining : 500, true, elapsed);
	}
	snprintf(aBuf, BUF_SPACE_REMAINING, "\r\nPress [F5] to refresh.");
	aBuf += strlen(aBuf);
	return aBuf;
}



char *Line::VicinityToText(char *aBuf, size_t aBufSize, int aMaxLines)
// Translates the current line and the lines above and below it into their text equivalent
// putting the result into aBuf and returning the position in aBuf of its new string terminator.
{
	if (!aBuf || aBufSize < 256) return NULL;

	char *aBuf_orig = aBuf;

	if (aMaxLines < 5)
		aMaxLines = 5;
	--aMaxLines; // -1 to leave room for the current line itself.

	int lines_following = (int)(aMaxLines / 2);
	int lines_preceding = aMaxLines - lines_following;

	// Determine the correct value for line_start and line_end:
	int i;
	Line *line_start, *line_end;
	for (i = 0, line_start = this
		; i < lines_preceding && line_start->mPrevLine != NULL
		; ++i, line_start = line_start->mPrevLine);

	for (i = 0, line_end = this
		; i < lines_following && line_end->mNextLine != NULL
		; ++i, line_end = line_end->mNextLine);

	// Now line_start and line_end are the first and last lines of the range
	// we want to convert to text, and they're non-NULL.
	snprintf(aBuf, BUF_SPACE_REMAINING, "\tLine#\n");
	aBuf += strlen(aBuf);

	size_t space_remaining;

	// Start at the oldest and continue up through the newest:
	for (Line *line = line_start;;)
	{
		if (line == this)
			strlcpy(aBuf, "--->\t", BUF_SPACE_REMAINING);
		else
			strlcpy(aBuf, "\t", BUF_SPACE_REMAINING);
		aBuf += strlen(aBuf);
		space_remaining = BUF_SPACE_REMAINING;  // Resolve macro only once for performance.
		// Truncate large lines so that the dialog is more readable:
		aBuf = line->ToText(aBuf, space_remaining < 500 ? space_remaining : 500, false);
		if (line == line_end)
			break;
		line = line->mNextLine;
	}
	return aBuf;
}



char *Line::ToText(char *aBuf, size_t aBufSize, bool aCRLF, DWORD aElapsed)
// Translates this line into its text equivalent, putting the result into aBuf and
// returning the position in aBuf of its new string terminator.
{
	if (!aBuf) return NULL;
	if (aBufSize < 3)
		return aBuf;
	else
		aBufSize -= (1 + aCRLF);  // Reserve one char for LF/CRLF after each line (so that it always get added).
	char *aBuf_orig = aBuf;
	snprintf(aBuf, BUF_SPACE_REMAINING, "%03u: ", mLineNumber);
	aBuf += strlen(aBuf);
	if (mActionType == ACT_IFBETWEEN || mActionType == ACT_IFNOTBETWEEN)
	{
		snprintf(aBuf, BUF_SPACE_REMAINING, "if %s %s %s and %s"
			, *mArg[0].text ? mArg[0].text : VAR(mArg[0])->mName  // i.e. don't resolve dynamic variable names.
			, g_act[mActionType].Name, RAW_ARG2, RAW_ARG3);
		aBuf += strlen(aBuf);
	}
	else if (ACT_IS_ASSIGN(mActionType) || (ACT_IS_IF(mActionType) && mActionType < ACT_FIRST_COMMAND))
	{
		// Only these other commands need custom conversion.
		snprintf(aBuf, BUF_SPACE_REMAINING, "%s%s %s %s"
			, ACT_IS_IF(mActionType) ? "if " : ""
			, *mArg[0].text ? mArg[0].text : VAR(mArg[0])->mName  // i.e. don't resolve dynamic variable names.
			, g_act[mActionType].Name, RAW_ARG2);
		aBuf += strlen(aBuf);
	}
	else
	{
		snprintf(aBuf, BUF_SPACE_REMAINING, "%s", g_act[mActionType].Name);
		aBuf += strlen(aBuf);
		for (int i = 0; i < mArgc; ++i)
		{
			// This method a little more efficient than using snprintfcat().
			// Also, always use the arg's text for input and output args whose variables haven't
			// been been resolved at load-time, since the text has everything in it we want to display
			// and thus there's no need to "resolve" dynamic variables here (e.g. array%i%).
			snprintf(aBuf, BUF_SPACE_REMAINING, ",%s", (mArg[i].type != ARG_TYPE_NORMAL && !*mArg[i].text)
				? VAR(mArg[i])->mName : mArg[i].text);
			aBuf += strlen(aBuf);
		}
	}
	if (aElapsed)
	{
		snprintf(aBuf, BUF_SPACE_REMAINING, " (%0.2f)", (float)aElapsed / 1000.0);
		aBuf += strlen(aBuf);
	}
	// UPDATE for v1.0.25: It seems that MessageBox(), which is the only way these lines are currently
	// displayed, prefers \n over \r\n because otherwise, Ctrl-C on the MsgBox copies the lines all
	// onto one line rather than formatted nicely as separate lines.
	// Room for LF or CRLF was reserved at the top of this function:
	if (aCRLF)
		*aBuf++ = '\r';
	*aBuf++ = '\n';
	*aBuf = '\0';
	return aBuf;
}



void Line::ToggleSuspendState()
{
	g_IsSuspended = !g_IsSuspended;
	Hotstring::SuspendAll(g_IsSuspended);  // Must do this prior to AllActivate() to avoid incorrect removal of hook.
	if (g_IsSuspended)
		Hotkey::AllDeactivate(true); // This will also reset the RunAgainAfterFinished flags for all those deactivated.
		// It seems unnecessary, and possibly undesirable, to purge any pending hotkey msgs from the msg queue.
		// Even if there are some, it's possible that they are exempt from suspension so we wouldn't want to
		// globally purge all messages anyway.
	else
		// For now, it seems best to call it with "true" in case the any hotkeys were dynamically added
		// while the suspension was in effect, in which case the nature of some old hotkeys may have changed
		// from registered to hook due to an interdependency with a newly added hotkey.  There are comments
		// in AllActivate() that describe these interdependencies:
		Hotkey::AllActivate();
	g_script.UpdateTrayIcon();
	CheckMenuItem(GetMenu(g_hWnd), ID_FILE_SUSPEND, g_IsSuspended ? MF_CHECKED : MF_UNCHECKED);
}



ResultType Line::ChangePauseState(ToggleValueType aChangeTo)
// Returns OK or FAIL.
// Note: g_Idle must be false since we're always called from a script subroutine, not from
// the tray menu.  Therefore, the value of g_Idle need never be checked here.
{
	switch (aChangeTo)
	{
	case TOGGLED_ON:
		// Pause the current subroutine (which by definition isn't paused since it had to be 
		// active to call us).  It seems best not to attempt to change the Hotkey mRunAgainAfterFinished
		// attribute for the current hotkey (assuming it's even a hotkey that got us here) or
		// for them all.  This is because it's conceivable that this Pause command occurred
		// in a background thread, such as a timed subroutine, in which case we wouldn't want the
		// pausing of that thread to affect anything else the user might be doing with hotkeys.
		// UPDATE: The above is flawed because by definition the script's quasi-thread that got
		// us here is now active.  Since it is active, the script will immediately become dormant
		// when this is executed, waiting for the user to press a hotkey to launch a new
		// quasi-thread.  Thus, it seems best to reset all the mRunAgainAfterFinished flags
		// in case we are in a hotkey subroutine and in case this hotkey has a buffered repeat-again
		// action pending, which the user probably wouldn't want to happen after the script is unpaused:
		#define PAUSE_CURRENT_THREAD \
		{\
			Hotkey::ResetRunAgainAfterFinished();\
			g.IsPaused = true;\
			++g_nPausedThreads;\
			g_script.UpdateTrayIcon();\
			CheckMenuItem(GetMenu(g_hWnd), ID_FILE_PAUSE, MF_CHECKED);\
		}
		PAUSE_CURRENT_THREAD
		return OK;
	case TOGGLED_OFF:
		// Unpause the uppermost underlying paused subroutine.  If none of the interrupted subroutines
		// are paused, do nothing. This relies on the fact that the current script subroutine,
		// which got us here, cannot itself be currently paused by definition:
		if (g_nPausedThreads > 0)
			g_UnpauseWhenResumed = true;
		return OK;
	case NEUTRAL: // the user omitted the parameter entirely, which is considered the same as "toggle"
	case TOGGLE:
		// See TOGGLED_OFF comment above:
		if (g_nPausedThreads > 0)
			g_UnpauseWhenResumed = true;
		else
			// Since there are no underlying paused subroutines, pause the current subroutine.
			// There must be a current subroutine (i.e. the script can't be idle) because
			// it's the one that called us directly, and we know it's not paused:
			PAUSE_CURRENT_THREAD
		return OK;
	default: // TOGGLE_INVALID or some other disallowed value.
		// We know it's a variable because otherwise the loading validation would have caught it earlier:
		return LineError("The variable in param #1 does not resolve to an allowed value.", FAIL, ARG1);
	}
}



ResultType Line::ScriptBlockInput(bool aEnable)
// Adapted from the AutoIt3 source.
// Always returns OK for caller convenience.
{
	// Must be running Win98/2000+ for this function to be successful.
	// We must dynamically load the function to retain compatibility with Win95 (program won't launch
	// at all otherwise).
	typedef void (CALLBACK *BlockInput)(BOOL);
	static BlockInput lpfnDLLProc = (BlockInput)GetProcAddress(GetModuleHandle("User32.dll"), "BlockInput");
	// Always turn input ON/OFF even if g_BlockInput says its already in the right state.  This is because
	// BlockInput can be externally and undetectibly disabled, e.g. if the user presses Ctrl-Alt-Del:
	if (lpfnDLLProc)
		(*lpfnDLLProc)(aEnable ? TRUE : FALSE);
	g_BlockInput = aEnable;
	return OK;  // By design, it never returns FAIL.
}



inline Line *Line::PreparseError(char *aErrorText)
// Returns a different type of result for use with the Pre-parsing methods.
{
	// Make all preparsing errors critical because the runtime reliability
	// of the program relies upon the fact that the aren't any kind of
	// problems in the script (otherwise, unexpected behavior may result).
	// Update: It's okay to return FAIL in this case.  CRITICAL_ERROR should
	// be avoided whenever OK and FAIL are sufficient by themselves, because
	// otherwise, callers can't use the NOT operator to detect if a function
	// failed (since FAIL is value zero, but CRITICAL_ERROR is non-zero):
	LineError(aErrorText, FAIL);
	return NULL; // Always return NULL because the callers use it as their return value.
}



ResultType Line::LineError(char *aErrorText, ResultType aErrorType, char *aExtraInfo)
{
	if (!aErrorText)
		aErrorText = "Unknown Error";
	if (!aExtraInfo)
		aExtraInfo = "";

	if (g_script.mErrorStdOut && !g_script.mIsReadyToExecute) // i.e. runtime errors are always displayed via dialog.
	{
		// JdeB said:
		// Just tested it in Textpad, Crimson and Scite. they all recognise the output and jump
		// to the Line containing the error when you double click the error line in the output
		// window (like it works in C++).  Had to change the format of the line to: 
		// printf("%s (%d) : ==> %s: \n%s \n%s\n",szInclude, nAutScriptLine, szText, szScriptLine, szOutput2 );
		// MY: Full filename is required, even if it's the main file, because some editors (EditPlus)
		// seem to rely on that to determine which file and line number to jump to when the user double-clicks
		// the error message in the output window:
		printf("%s (%d): ==> %s\n", sSourceFile[mFileNumber], mLineNumber, aErrorText);
		if (*aExtraInfo)
			printf("     Specifically: %s\n", aExtraInfo);
	}
	else
	{
		char source_file[MAX_PATH * 2];
		if (mFileNumber)
			snprintf(source_file, sizeof(source_file), " in #include file \"%s\"", sSourceFile[mFileNumber]);
		else
			*source_file = '\0'; // Don't bother cluttering the display if it's the main script file.

		char buf[MSGBOX_TEXT_SIZE];
		snprintf(buf, sizeof(buf), "%s%s: %-1.500s\n\n"  // Keep it to a sane size in case it's huge.
			, aErrorType == WARN ? "Warning" : (aErrorType == CRITICAL_ERROR ? "Critical Error" : "Error")
			, source_file, aErrorText);
		if (*aExtraInfo)
			// Use format specifier to make sure really huge strings that get passed our
			// way, such as a var containing clipboard text, are kept to a reasonable size:
			snprintfcat(buf, sizeof(buf), "Specifically: %-1.100s%s\n\n"
			, aExtraInfo, strlen(aExtraInfo) > 100 ? "..." : "");
		char *buf_marker = buf + strlen(buf);
		buf_marker = VicinityToText(buf_marker, sizeof(buf) - (buf_marker - buf));
		if (aErrorType == CRITICAL_ERROR || (aErrorType == FAIL && !g_script.mIsReadyToExecute))
			strlcpy(buf_marker, g_script.mIsRestart ? ("\n" OLD_STILL_IN_EFFECT) : ("\n" WILL_EXIT)
				, sizeof(buf) - (buf_marker - buf));
		//buf_marker += strlen(buf_marker);
		g_script.mCurrLine = this;  // This needs to be set in some cases where the caller didn't.
		//g_script.ShowInEditor();
		MsgBox(buf);
	}

	if (aErrorType == CRITICAL_ERROR && g_script.mIsReadyToExecute)
		// Also ask the main message loop function to quit and announce to the system that
		// we expect it to quit.  In most cases, this is unnecessary because all functions
		// called to get to this point will take note of the CRITICAL_ERROR and thus keep
		// return immediately, all the way back to main.  However, there may cases
		// when this isn't true:
		// Note: Must do this only after MsgBox, since it appears that new dialogs can't
		// be created once it's done.  Update: Using ExitApp() now, since it's known to be
		// more reliable:
		//PostQuitMessage(CRITICAL_ERROR);
		// This will attempt to run the OnExit subroutine, which should be okay since that subroutine
		// will terminate the script if it encounters another runtime error:
		g_script.ExitApp(EXIT_ERROR);

	return aErrorType; // The caller told us whether it should be a critical error or not.
}



ResultType Script::ScriptError(char *aErrorText, char *aExtraInfo) //, ResultType aErrorType)
// Even though this is a Script method, including it here since it shares
// a common theme with the other error-displaying functions:
{
	if (mCurrLine)
		// If a line is available, do LineError instead since it's more specific.
		// If an error occurs before the script is ready to run, assume it's always critical
		// in the sense that the program will exit rather than run the script.
		// Update: It's okay to return FAIL in this case.  CRITICAL_ERROR should
		// be avoided whenever OK and FAIL are sufficient by themselves, because
		// otherwise, callers can't use the NOT operator to detect if a function
		// failed (since FAIL is value zero, but CRITICAL_ERROR is non-zero):
		return mCurrLine->LineError(aErrorText, FAIL, aExtraInfo);
	// Otherwise: The fact that mCurrLine is NULL means that the line currently being loaded
	// has not yet been successfully added to the linked list.  Such errors will always result
	// in the program exiting.
	if (!aErrorText)
		aErrorText = "Unknown Error";
	if (!aExtraInfo) // In case the caller explicitly called it with NULL.
		aExtraInfo = "";

	if (g_script.mErrorStdOut && !g_script.mIsReadyToExecute) // i.e. runtime errors are always displayed via dialog.
	{
		// See LineError() for details.
		printf("%s (%d): ==> %s\n", Line::sSourceFile[mCurrFileNumber], mCurrLineNumber, aErrorText);
		if (*aExtraInfo)
			printf("     Specifically: %s\n", aExtraInfo);
	}
	else
	{
		char source_file[MAX_PATH * 2];
		if (mCurrFileNumber)
			snprintf(source_file, sizeof(source_file), " in #include file \"%s\"", Line::sSourceFile[mCurrFileNumber]);
		else
			*source_file = '\0'; // Don't bother cluttering the display if it's the main script file.

		char buf[MSGBOX_TEXT_SIZE];
		snprintf(buf, sizeof(buf), "Error at line %u%s%s." // Don't call it "critical" because it's usually a syntax error.
			"\n\nLine Text: %-1.100s%s"
			"\nError: %-1.500s"
			"\n\n%s"
			, mCurrLineNumber, source_file, mCurrLineNumber ? "" : " (unknown)"
			, aExtraInfo // aExtraInfo defaults to "" so this is safe.
			, strlen(aExtraInfo) > 100 ? "..." : ""
			, aErrorText
			, mIsRestart ? OLD_STILL_IN_EFFECT : WILL_EXIT
			);
		//ShowInEditor();
		MsgBox(buf);
	}
	return FAIL; // See above for why it's better to return FAIL than CRITICAL_ERROR.
}



char *Script::ListVars(char *aBuf, size_t aBufSize)
// Translates this script's list of variables into text equivalent, putting the result
// into aBuf and returning the position in aBuf of its new string terminator.
{
	if (!aBuf || aBufSize < 256) return NULL;
	char *aBuf_orig = aBuf;
	snprintf(aBuf, BUF_SPACE_REMAINING, "Variables (alphabetical) & their current contents:\r\n\r\n");
	aBuf += strlen(aBuf);
	// Start at the oldest and continue up through the newest:
	for (Var *var = mFirstVar; var; var = var->mNextVar)
		if (var->mType == VAR_NORMAL) // Don't bother showing clipboard and other built-in vars.
			aBuf = var->ToText(aBuf, BUF_SPACE_REMAINING, true);
	return aBuf;
}



char *Script::ListKeyHistory(char *aBuf, size_t aBufSize)
// Translates this key history into text equivalent, putting the result
// into aBuf and returning the position in aBuf of its new string terminator.
{
	if (!aBuf || aBufSize < 256) return NULL;
	char *aBuf_orig = aBuf; // Needed for the BUF_SPACE_REMAINING macro.
	// I was initially concerned that GetWindowText() can hang if the target window is
	// hung.  But at least on newer OS's, this doesn't seem to be a problem: MSDN says
	// "If the window does not have a caption, the return value is a null string. This
	// behavior is by design. It allows applications to call GetWindowText without hanging
	// if the process that owns the target window is hung. However, if the target window
	// is hung and it belongs to the calling application, GetWindowText will hang the
	// calling application."
	HWND target_window = GetForegroundWindow();
	char win_title[100];
	if (target_window)
		GetWindowText(target_window, win_title, sizeof(win_title));
	else
		*win_title = '\0';

	char timer_list[128] = "";
	for (ScriptTimer *timer = mFirstTimer; timer != NULL; timer = timer->mNextTimer)
		if (timer->mEnabled)
			snprintfcat(timer_list, sizeof(timer_list) - 3, "%s ", timer->mLabel->mName); // Allow room for "..."
	if (*timer_list)
	{
		size_t length = strlen(timer_list);
		if (length > (sizeof(timer_list) - 5))
			strlcpy(timer_list + length, "...", sizeof(timer_list) - length);
		else if (timer_list[length - 1] == ' ')
			timer_list[--length] = '\0';  // Remove the last space if there was room enough for it to have been added.
	}

	char LRtext[256];
	snprintf(aBuf, aBufSize,
		"Window: %s"
		//"\r\nBlocks: %u"
		"\r\nKeybd hook: %s"
		"\r\nMouse hook: %s"
		"\r\nLast hotkey type: %s"
		"\r\nEnabled Timers: %u of %u (%s)"
		//"\r\nInterruptible?: %s"
		"\r\nInterrupted threads: %d%s"
		"\r\nPaused threads: %d of %d (%d layers)"
		"\r\nModifiers (GetKeyState() now) = %s"
		"\r\n"
		, win_title
		//, SimpleHeap::GetBlockCount()
		, g_KeybdHook == NULL ? "no" : "yes"
		, g_MouseHook == NULL ? "no" : "yes"
		, g_LastPerformedHotkeyType == HK_KEYBD_HOOK ? "keybd hook" : "not keybd hook"
		, mTimerEnabledCount, mTimerCount, timer_list
		//, INTERRUPTIBLE ? "yes" : "no"
		, g_nThreads > 1 ? g_nThreads - 1 : 0
		, g_nThreads > 1 ? " (preempted: they will resume when the current thread finishes)" : ""
		, g_nPausedThreads, g_nThreads, g_nLayersNeedingTimer
		, ModifiersLRToText(GetModifierLRState(true), LRtext));
	aBuf += strlen(aBuf);
	GetHookStatus(aBuf, BUF_SPACE_REMAINING);
	aBuf += strlen(aBuf);
	snprintf(aBuf, BUF_SPACE_REMAINING, "\r\nPress [F5] to refresh.");
	aBuf += strlen(aBuf);
	return aBuf;
}



ResultType Script::ActionExec(char *aAction, char *aParams, char *aWorkingDir, bool aDisplayErrors
	, char *aRunShowMode, HANDLE *aProcess, bool aUseRunAs, Var *aOutputVar)
// Caller should specify NULL for aParams if it wants us to attempt to parse out params from
// within aAction.  Caller may specify empty string ("") instead to specify no params at all.
// Remember that aAction and aParams can both be NULL, so don't dereference without checking first.
// Note: For the Run & RunWait commands, aParams should always be NULL.  Params are parsed out of
// the aActionString at runtime, here, rather than at load-time because Run & RunWait might contain
// deferenced variable(s), which can only be resolved at runtime.
{
	if (aProcess) // Init output param if the caller gave us memory to store it.
		*aProcess = NULL;
	if (aOutputVar) // Same
		aOutputVar->Assign();

	// Launching nothing is always a success:
	if (!aAction || !*aAction) return OK;

	if (strlen(aAction) >= LINE_SIZE) // This can happen if user runs the contents of a very large variable.
	{
        if (aDisplayErrors)
			ScriptError("The string to be run is too long." ERR_ABORT);
		return FAIL;
	}

	// Make sure this is set to NULL because CreateProcess() won't work if it's the empty string:
	if (aWorkingDir && !*aWorkingDir)
		aWorkingDir = NULL;

	#define IS_VERB(str) (   !stricmp(str, "find") || !stricmp(str, "explore") || !stricmp(str, "open")\
		|| !stricmp(str, "edit") || !stricmp(str, "print") || !stricmp(str, "properties")   )

	// Declare this buf here to ensure it's in scope for the entire function, since its
	// contents may be referred to indirectly:
	char parse_buf[LINE_SIZE];

	// Set default items to be run by ShellExecute().  These are also used by the error
	// reporting at the end, which is why they're initialized even if CreateProcess() works
	// and there's no need to use ShellExecute():
	char *shell_action = aAction;
	char *shell_params = aParams ? aParams : "";
	bool shell_action_is_system_verb = false;

	///////////////////////////////////////////////////////////////////////////////////
	// This next section is done prior to CreateProcess() because when aParams is NULL,
	// we need to find out whether aAction contains a system verb.
	///////////////////////////////////////////////////////////////////////////////////
	if (aParams) // Caller specified the params (even an empty string counts, for this purpose).
		shell_action_is_system_verb = IS_VERB(shell_action);
	else // Caller wants us to try to parse params out of aAction.
	{
		// Make a copy so that we can modify it (i.e. split it into action & params):
		strlcpy(parse_buf, aAction, sizeof(parse_buf));

		// Find out the "first phrase" in the string.  This is done to support the special "find" and "explore"
		// operations as well as minmize the chance that executable names intended by the user to be parameters
		// will not be considered to be the program to run (e.g. for use with a compiler, perhaps).
		char *first_phrase, *first_phrase_end, *second_phrase;
		if (*parse_buf == '\"')
		{
			first_phrase = parse_buf + 1;  // Omit the double-quotes, for use with CreateProcess() and such.
			first_phrase_end = strchr(first_phrase, '\"');
		}
		else
		{
			first_phrase = parse_buf;
			// Set first_phrase_end to be the location of the first whitespace char, if
			// one exists:
			first_phrase_end = StrChrAny(first_phrase, " \t"); // Find space or tab.
		}
		// Now first_phrase_end is either NULL, the position of the last double-quote in first-phrase,
		// or the position of the first whitespace char to the right of first_phrase.
		if (first_phrase_end)
		{
			// Split into two phrases:
			*first_phrase_end = '\0';
			second_phrase = first_phrase_end + 1;
		}
		else // the entire string is considered to be the first_phrase, and there's no second:
			second_phrase = NULL;
		if (shell_action_is_system_verb = IS_VERB(first_phrase))
		{
			shell_action = first_phrase;
			shell_params = second_phrase ? second_phrase : "";
		}
		else
		{
// Rather than just consider the first phrase to be the executable and the rest to be the param, we check it
// for a proper extension so that the user can launch a document name containing spaces, without having to
// enclose it in double quotes.  UPDATE: Want to be able to support executable filespecs without requiring them
// to be enclosed in double quotes.  Therefore, search the entire string, rather than just first_phrase, for
// the left-most occurrence of a valid executable extension.  This should be fine since the user can still
// pass in EXEs and such as params as long as the first executable is fully qualified with its real extension
// so that we can tell that it's the action and not one of the params.

// This method is rather crude because is doesn't handle an extensionless executable such as "notepad test.txt"
// It's important that it finds the first occurrence of an executable extension in case there are other
// occurrences in the parameters.  Also, .pif and .lnk are currently not considered executables for this purpose
// since they probably don't accept parameters:
			strlcpy(parse_buf, aAction, sizeof(parse_buf));  // Restore the original value in case it was changed.
			char *action_extension;
			if (   !(action_extension = strcasestr(parse_buf, ".exe "))   )
				if (   !(action_extension = strcasestr(parse_buf, ".exe\""))   )
					if (   !(action_extension = strcasestr(parse_buf, ".bat "))   )
						if (   !(action_extension = strcasestr(parse_buf, ".bat\""))   )
							if (   !(action_extension = strcasestr(parse_buf, ".com "))   )
								if (   !(action_extension = strcasestr(parse_buf, ".com\""))   )
									// Not 100% sure that .cmd and .hta are genuine executables in every sense:
									if (   !(action_extension = strcasestr(parse_buf, ".cmd "))   )
										if (   !(action_extension = strcasestr(parse_buf, ".cmd\""))   )
											if (   !(action_extension = strcasestr(parse_buf, ".hta "))   )
												action_extension = strcasestr(parse_buf, ".hta\"");

			if (action_extension)
			{
				shell_action = parse_buf;
				// +4 for the 3-char extension with the period:
				shell_params = action_extension + 4;  // exec_params is now the start of params, or empty-string.
				if (*shell_params == '\"')
					// Exclude from shell_params since it's probably belongs to the action, not the params
					// (i.e. it's paired with another double-quote at the start):
					++shell_params;
				if (*shell_params)
				{
					// Terminate the <aAction> string in the right place.  For this to work correctly,
					// at least one space must exist between action & params (shortcoming?):
					*shell_params = '\0';
					++shell_params;
					ltrim(shell_params); // Might be empty string after this, which is ok.
				}
				// else there doesn't appear to be any params, so just leave shell_params set to empty string.
			}
			// else there's no extension: so assume the whole <aAction> is a document name to be opened by
			// the shell.  So leave shell_action and shell_params set their original defaults.
		}
	}

	// This is distinct from new_process being non-NULL because the two aren't always the
	// same.  For example, if the user does "Run, find D:\" or "RunWait, www.yahoo.com",
	// no new process handle will be available even though the launch was successful:
	bool success = false;
	HANDLE new_process = NULL;  // This will hold the handle to the newly created process.
	char system_error_text[512] = "";

	bool use_runas = aUseRunAs && mRunAsUser && (*mRunAsUser || *mRunAsPass || *mRunAsDomain);
	if (use_runas && shell_action_is_system_verb)
	{
		if (aDisplayErrors)
			ScriptError("System verbs aren't supported when RunAs is in effect." ERR_ABORT);
		return FAIL;
	}

	// If the caller originally gave us NULL for aParams, always try CreateProcess() before
	// trying ShellExecute().  This is because ShellExecute() is usually a lot slower.
	// The only exception is if the action appears to be a verb such as open, edit, or find.
	// In that case, we'll also skip the CreateProcess() attempt and do only the ShellExecute().
	// If the user really meant to launch find.bat or find.exe, for example, he should add
	// the extension (e.g. .exe) to differentiate "find" from "find.exe":
	if (!shell_action_is_system_verb)
	{
		STARTUPINFO si = {0};  // Zero fill to be safer.
		si.cb = sizeof(si);
		si.lpReserved = si.lpDesktop = si.lpTitle = NULL;
		si.lpReserved2 = NULL;
		si.dwFlags = STARTF_USESHOWWINDOW;  // This tells it to use the value of wShowWindow below.
		si.wShowWindow = (aRunShowMode && *aRunShowMode) ? Line::ConvertRunMode(aRunShowMode) : SW_SHOWNORMAL;
		PROCESS_INFORMATION pi = {0};

		// Since CreateProcess() requires that the 2nd param be modifiable, ensure that it is
		// (even if this is ANSI and not Unicode; it's just safer):
		char command_line[LINE_SIZE];
		if (aParams && *aParams)
			snprintf(command_line, sizeof(command_line), "%s %s", aAction, aParams);
		else
        	strlcpy(command_line, aAction, sizeof(command_line)); // i.e. we're running the original action from caller.

		if (use_runas)
		{
			// This block has been adapted from the AutoIt3 source.
			typedef BOOL (WINAPI *MyCreateProcessWithLogonW)(
				LPCWSTR lpUsername,                 // user's name
				LPCWSTR lpDomain,                   // user's domain
				LPCWSTR lpPassword,                 // user's password
				DWORD dwLogonFlags,                 // logon option
				LPCWSTR lpApplicationName,          // executable module name
				LPWSTR lpCommandLine,               // command-line string
				DWORD dwCreationFlags,              // creation flags
				LPVOID lpEnvironment,               // new environment block
				LPCWSTR lpCurrentDirectory,         // current directory name
				LPSTARTUPINFOW lpStartupInfo,       // startup information
				LPPROCESS_INFORMATION lpProcessInfo // process information
				);
			// Get a handle to the DLL module that contains CreateProcessWithLogonW
			HINSTANCE hinstLib = LoadLibrary("advapi32.dll");
			if (!hinstLib)
			{
				if (aDisplayErrors)
					ScriptError("RunAs: Could not load advapi32.dll." ERR_ABORT);
				return FAIL;
			}
			MyCreateProcessWithLogonW lpfnDLLProc = (MyCreateProcessWithLogonW)GetProcAddress(hinstLib, "CreateProcessWithLogonW");
			if (!lpfnDLLProc)
			{
				FreeLibrary(hinstLib);
				if (aDisplayErrors)
					ScriptError("CreateProcessWithLogonW." ERR_ABORT); // Short msg since it probably never happens.
				return FAIL;
			}
			// Set up wide char version that we need for CreateProcessWithLogon
			// init structure for running programs (wide char version)
			STARTUPINFOW wsi;
			wsi.cb			= sizeof(STARTUPINFOW);
			wsi.lpReserved	= NULL;
			wsi.lpDesktop	= NULL;
			wsi.lpTitle		= NULL;
			wsi.dwFlags		= STARTF_USESHOWWINDOW;
			wsi.cbReserved2	= 0;
			wsi.lpReserved2	= NULL;
			wsi.wShowWindow = si.wShowWindow;				// Default is same as si

			// Convert to wide character format:
			wchar_t command_line_wide[LINE_SIZE], working_dir_wide[MAX_PATH];
			mbstowcs(command_line_wide, command_line, sizeof(command_line_wide));
			if (aWorkingDir && *aWorkingDir)
				mbstowcs(working_dir_wide, aWorkingDir, sizeof(working_dir_wide));
			else
				*working_dir_wide = 0;  // wide-char terminator.

			if (lpfnDLLProc(mRunAsUser, mRunAsDomain, mRunAsPass, LOGON_WITH_PROFILE, 0
				, command_line_wide, 0, 0, *working_dir_wide ? working_dir_wide : NULL, &wsi, &pi))
			{
				success = true;
				if (pi.hThread)
					CloseHandle(pi.hThread); // Required to avoid memory leak.
				new_process = pi.hProcess;
				if (aOutputVar)
					aOutputVar->Assign(pi.dwProcessId);
			}
			else
				GetLastErrorText(system_error_text, sizeof(system_error_text));
			FreeLibrary(hinstLib);
		}
		else
		{
			// MSDN: "If [lpCurrentDirectory] is NULL, the new process is created with the same
			// current drive and directory as the calling process." (i.e. since caller may have
			// specified a NULL aWorkingDir).  Also, we pass NULL in for the first param so that
			// it will behave the following way (hopefully under all OSes): "the first white-space – delimited
			// token of the command line specifies the module name. If you are using a long file name that
			// contains a space, use quoted strings to indicate where the file name ends and the arguments
			// begin (see the explanation for the lpApplicationName parameter). If the file name does not
			// contain an extension, .exe is appended. Therefore, if the file name extension is .com,
			// this parameter must include the .com extension. If the file name ends in a period (.) with
			// no extension, or if the file name contains a path, .exe is not appended. If the file name does
			// not contain a directory path, the system searches for the executable file in the following
			// sequence...".
			// Provide the app name (first param) if possible, for greater expected reliability.
			// UPDATE: Don't provide the module name because if it's enclosed in double quotes,
			// CreateProcess() will fail, at least under XP:
			//if (CreateProcess(aParams && *aParams ? aAction : NULL
			if (CreateProcess(NULL, command_line, NULL, NULL, FALSE, 0, NULL, aWorkingDir, &si, &pi))
			{
				success = true;
				if (pi.hThread)
					CloseHandle(pi.hThread); // Required to avoid memory leak.
				new_process = pi.hProcess;
				if (aOutputVar)
					aOutputVar->Assign(pi.dwProcessId);
			}
			else
				GetLastErrorText(system_error_text, sizeof(system_error_text));
		}
	}

	if (!success) // Either the above wasn't attempted, or the attempt failed.  So try ShellExecute().
	{
		if (use_runas)
		{
			// Since CreateProcessWithLogonW() was either not attempted or did not work, it's probably
			// best to display an error rather than trying to run it without the RunAs settings.
			// This policy encourages users to have RunAs in effect only when necessary:
			if (aDisplayErrors)
				ScriptError("Launch Error (possibly related to RunAs)." ERR_ABORT, system_error_text);
			return FAIL;
		}
		SHELLEXECUTEINFO sei = {0};
		// sei.hwnd is left NULL to avoid potential side-effects with having a hidden window be the parent.
		// However, doing so may result in the launched app appearing on a different monitor than the
		// script's main window appears on (for multimonitor systems).  This seems fairly inconsequential
		// since scripted workarounds are possible.
		sei.cbSize = sizeof(sei);
		// Below: "indicate that the hProcess member receives the process handle" and not to display error dialog:
		sei.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_FLAG_NO_UI;
		sei.lpDirectory = aWorkingDir; // OK if NULL or blank; that will cause current dir to be used.
		sei.nShow = (aRunShowMode && *aRunShowMode) ? Line::ConvertRunMode(aRunShowMode) : SW_SHOWNORMAL;
		if (shell_action_is_system_verb)
		{
			sei.lpVerb = shell_action;
			if (!stricmp(shell_action, "properties"))
				sei.fMask |= SEE_MASK_INVOKEIDLIST;  // Need to use this for the "properties" verb to work reliably.
			sei.lpFile = shell_params;
			sei.lpParameters = NULL;
		}
		else
		{
			sei.lpVerb = NULL;  // A better choice than "open" because NULL causes default verb to be used.
			sei.lpFile = shell_action;
			sei.lpParameters = shell_params;
		}
		if (ShellExecuteEx(&sei)) // Relies on short-circuit boolean order.
		{
			new_process = sei.hProcess;
			// aOutputVar is left blank because:
			// ProcessID is not available when launched this way, and since GetProcessID() is only
			// available in WinXP SP1, no effort is currently made to dynamically load it from
			// kernel32.dll (to retain compatibility with older OSes).
			success = true;
		}
		else
			GetLastErrorText(system_error_text, sizeof(system_error_text));
	}

	if (!success) // The above attempt(s) to launch failed.
	{
		if (aDisplayErrors)
		{
			char error_text[2048], verb_text[128];
			if (shell_action_is_system_verb)
				snprintf(verb_text, sizeof(verb_text), "\nVerb: <%s>", shell_action);
			else // Don't bother showing it if it's just "open".
				*verb_text = '\0';
			// Use format specifier to make sure it doesn't get too big for the error
			// function to display:
			snprintf(error_text, sizeof(error_text)
				, "Failed attempt to launch program or document:"
				"\nAction: <%-0.400s%s>"
				"%s"
				"\nParams: <%-0.400s%s>\n\n" ERR_ABORT_NO_SPACES
				, shell_action, strlen(shell_action) > 400 ? "..." : ""
				, verb_text
				, shell_params, strlen(shell_params) > 400 ? "..." : ""
				);
			ScriptError(error_text, system_error_text);
		}
		return FAIL;
	}

	if (aProcess) // The caller wanted the process handle and it must eventually call CloseHandle().
		*aProcess = new_process;
	else if (new_process) // It can be NULL in the case of launching things like "find D:\" or "www.yahoo.com"
		CloseHandle(new_process); // Required to avoid memory leak.
	return OK;
}
