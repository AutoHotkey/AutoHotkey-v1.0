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

#include "script.h"
#include "globaldata.h" // for a lot of things
#include "util.h" // for strlcpy() etc.
#include "mt19937ar-cok.h" // for random number generator
#include <time.h> // for time()
#include <sys/timeb.h> // for _timeb struct
#include "window.h" // for a lot of things
#include "application.h" // for MsgSleep()
#include <mmsystem.h> // for waveOutSetVolume()

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
	, mThisHotkeyLabel(NULL), mPriorHotkeyLabel(NULL), mPriorHotkeyStartTime(0)
	, mFirstLabel(NULL), mLastLabel(NULL)
	, mFirstVar(NULL), mLastVar(NULL)
	, mLineCount(0), mLabelCount(0), mVarCount(0), mGroupCount(0)
	, mFileLineCount(0)
	, mFileSpec(""), mFileDir(""), mFileName(""), mOurEXE(""), mMainWindowTitle("")
	, mIsReadyToExecute(false)
	, mIsRestart(false)
	, mIsAutoIt2(false)
	, mLinesExecutedThisCycle(0)
{
	ZeroMemory(&mNIC, sizeof(mNIC));  // To be safe.
	mNIC.hWnd = NULL;  // Set this as an indicator that it tray icon is not installed.
#ifdef _DEBUG
	int LargestMaxParams, i, j;
	ActionTypeType *np;
	// Find the Larged value of MaxParams used by any command and make sure it
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



ResultType Script::Init(char *aScriptFilename, bool aIsRestart)
// Returns OK or FAIL.
{
	mIsRestart = aIsRestart;
	if (!aScriptFilename || !*aScriptFilename) return FAIL;
	char buf[2048]; // Just make sure we have plenty of room to do things with.
	char *filename_marker;
	// In case the config file is a relative filespec (relative to current working dir):
	if (!GetFullPathName(aScriptFilename, sizeof(buf), buf, &filename_marker))
	{
		MsgBox("Script::Init(): GetFullPathName() failed.");
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
		// Set the old/AutoIt2 defaults for maximum safety and compatibilility:
		g_AllowSameLineComments = false;
		g_EscapeChar = '\\';
		g.TitleFindFast = true; // In case the normal default is false.
		g.DetectHiddenText = false;
		g.DefaultMouseSpeed = 1;  // Make the mouse fast like AutoIt2, but not quite insta-move.
		g.KeyDelay = 20;
		g.WinDelay = 500;
		g.LinesPerCycle = 1;
		// Reduce max params so that any non escaped delimiters the user may be using literally
		// in "window text" will still be considered literal, rather than as delimiters for
		// args that are not supported by AutoIt2, such as exclude-title, exclude-text, MsgBox
		// timeout, etc.  Note: Don't need to change IfWinExist and such because those already
		// have special handling to recognize whether exclude-title is really a valid command
		// instead (e.g. IfWinExist, title, text, Gosub, something).

		// NOTE: DO NOT ADD the IfWin command series to this section, since there is special handling
		// for parsing those commands to figure out whether they're being used in the old AutoIt2
		// style or the new Exclude Title/Text mode.

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
	snprintf(buf, sizeof(buf), "%s\\%s - %s", mFileDir, mFileName, NAME_PV);
	if (   !(mMainWindowTitle = SimpleHeap::Malloc(buf))   )
		return FAIL;  // It already displayed the error for us.
	if (GetModuleFileName(NULL, buf, sizeof(buf))) // realistically, probably can't fail.
		if (   !(mOurEXE = SimpleHeap::Malloc(buf))   )
			return FAIL;  // It already displayed the error for us.
	return OK;
}

	

ResultType Script::CreateWindows(HINSTANCE aInstance)
// Returns OK or FAIL.
{
	if (!mMainWindowTitle || !*mMainWindowTitle) return FAIL;  // Init() must be called before this function.
	// Register a window class for the main window:
	HICON hIcon = LoadIcon(aInstance, MAKEINTRESOURCE(IDI_ICON_MAIN)); // LoadIcon((HINSTANCE) NULL, IDI_APPLICATION)
	WNDCLASSEX wc;
	ZeroMemory(&wc, sizeof(wc));
	wc.cbSize = sizeof(wc);
	wc.lpszClassName = WINDOW_CLASS_NAME;
	wc.hInstance = aInstance;
	wc.lpfnWndProc = MainWindowProc;
	// Provided from some example code:
	wc.style = 0;  // Aut3: CS_HREDRAW | CS_VREDRAW
	wc.cbClsExtra = 0;
	wc.cbWndExtra = 0;
	wc.hIcon = hIcon;
	wc.hIconSm = hIcon;
	wc.hCursor = LoadCursor((HINSTANCE) NULL, IDC_ARROW);
	wc.hbrBackground = (HBRUSH)GetStockObject(WHITE_BRUSH); // Aut3: GetSysColorBrush(COLOR_BTNFACE)
	wc.lpszMenuName = NULL; // "MainMenu";
	ATOM wa = RegisterClassEx(&wc);
	if (!wa)
	{
		MsgBox("RegisterClass() failed.");
		return FAIL;
	}

	// Note: the title below must be constructed the same was as is done by our
	// WinMain() (so that we can detect whether this script is already running)
	// which is why it's standardized in g_script.mMainWindowTitle.
	// Create the main window:
	if (   !(g_hWnd = CreateWindow(
		  WINDOW_CLASS_NAME
		, mMainWindowTitle
		, WS_OVERLAPPEDWINDOW // Style.  Alt: WS_POPUP or maybe 0.
		, CW_USEDEFAULT // xpos
		, CW_USEDEFAULT // ypos
		, CW_USEDEFAULT // width
		, CW_USEDEFAULT // height
		, NULL // parent window
		, NULL // Identifies a menu, or specifies a child-window identifier depending on the window style
		, aInstance // passed into WinMain
		, NULL))   ) // lpParam
	{
		MsgBox("CreateWindow() failed.");
		return FAIL;
	}

	// AutoIt3: Add read-only edit control to our main window:
	if (    !(g_hWndEdit = CreateWindow("edit", NULL, WS_CHILD | WS_VISIBLE | WS_BORDER
		| ES_LEFT | ES_MULTILINE | ES_READONLY | WS_VSCROLL // | WS_HSCROLL (saves space)
		, 0, 0, 0, 0, g_hWnd, (HMENU)1, aInstance, NULL))  )
	{
		MsgBox("CreateWindow() for the edit-window child failed.");
		return FAIL;
	}

	// To be compliant, we're supposed to do this.  Also, some of the MSDN docs mention that
	// an app's very first call to ShowWindow() makes that function operate in a special mode.
	// Therefore, it seems best to get that first call out of the way to avoid the possibility
	// that the first-call behavior will cause problems with our normal use of ShowWindow()
	// elsewhere.  UPDATE: Decided to do only the SW_HIDE one, ingoring default / nCmdShow.
	// That should avoid any momentary visual effects on startup:
	//ShowWindow(g_hWnd, SW_SHOWDEFAULT);  // The docs conflict, sometimes suggesting nCmdShow vs. this.
	//UpdateWindow(g_hWnd);  // Not necessary because it's empty.
	// Should do at least one call.  But sometimes SW_HIDE will be ignored the first time
	// (see MSDN docs), so do two calls to be sure the window is really hidden:
	ShowWindow(g_hWnd, SW_HIDE);
	ShowWindow(g_hWnd, SW_HIDE); // 2nd call to be safe.


	////////////////////
	// Set up tray icon.
	////////////////////
	ZeroMemory(&mNIC, sizeof(mNIC));  // To be safe.
	// Using NOTIFYICONDATA_V2_SIZE vs. sizeof(NOTIFYICONDATA) improves compatibility with Win9x maybe.
	// MSDN: "Using [NOTIFYICONDATA_V2_SIZE] for cbSize will allow your application to use NOTIFYICONDATA
	// with earlier Shell32.dll versions, although without the version 6.0 enhancements."
	// Update: Using V2 gives an compile error so trying V1:
	mNIC.cbSize				= NOTIFYICONDATA_V1_SIZE;
	mNIC.hWnd				= g_hWnd;
	mNIC.uID				= 0;  // Icon ID (can be anything, like Timer IDs?)
	mNIC.uFlags				= NIF_MESSAGE | NIF_TIP | NIF_ICON;
	mNIC.uCallbackMessage	= AHK_NOTIFYICON;
	mNIC.hIcon				= LoadIcon(aInstance, MAKEINTRESOURCE(IDI_ICON_MAIN));
	strlcpy(mNIC.szTip, mFileName ? mFileName : NAME_P, sizeof(mNIC.szTip));
	if (!Shell_NotifyIcon(NIM_ADD, &mNIC))
	{
		mNIC.hWnd = NULL;  // Set this as an indicator that tray icon is not installed.
		return FAIL;
	}
	return OK;
}



void Script::UpdateTrayIcon()
{
	if (!mNIC.hWnd) // tray icon is not installed
		return;
	static bool icon_shows_paused = false;
	static bool icon_shows_suspended = false;
	if (g.IsPaused == icon_shows_paused && g_IsSuspended == icon_shows_suspended) // it's already in the right state
		return;
	int icon;
	if (g.IsPaused && g_IsSuspended)
		icon = IDI_ICON_PAUSE_SUSPEND;
	else if (g.IsPaused)
		icon = IDI_ICON_PAUSE;
	else if (g_IsSuspended)
		icon = IDI_ICON_SUSPEND;
	else
		icon = IDI_ICON_MAIN;
	mNIC.hIcon = LoadIcon(g_hInstance, MAKEINTRESOURCE(icon));
	if (Shell_NotifyIcon(NIM_MODIFY, &mNIC))
	{
		icon_shows_paused = g.IsPaused;
		icon_shows_suspended = g_IsSuspended;
	}
	// else do nothing, just leave it in the same state.
}



ResultType Script::Edit()
{
	bool old_mode = g.TitleFindAnywhere;
	g.TitleFindAnywhere = true;
	HWND hwnd = WinExist(mFileName, "", mMainWindowTitle); // Exclude our own main.
	g.TitleFindAnywhere = old_mode;
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
		if (!ActionExec("edit", mFileSpec, mFileDir, false))  // Since this didn't work, try notepad.
		{
			// Even though notepad properly handles filenames with spaces in them under WinXP,
			// even without double quotes around them, it seems safer and more correct to always
			// enclose the filename in double quotes for maximum compatibility with all OSes:
			char buf[MAX_PATH * 2];
			snprintf(buf, sizeof(buf), "\"%s\"", mFileSpec);
			if (!ActionExec("notepad.exe", buf, mFileDir, false))
				MsgBox("Could not open the file for editing using the associated \"edit\" action or Notepad.");
		}
	}
	return OK;
}



ResultType Script::Reload()
{
	char arg_string[MAX_PATH + 512], current_dir[MAX_PATH];
	GetCurrentDirectory(sizeof(current_dir), current_dir);  // In case the user launched it in a non-default dir.
	snprintf(arg_string, sizeof(arg_string), "/restart \"%s\"", mFileSpec);
	g_script.ActionExec(mOurEXE, arg_string, current_dir); // It will tell our process to stop.
	return OK;
}



void Script::ExitApp(char *aBuf, int ExitCode)
// Normal exit (if aBuf is NULL), or a way to exit immediately on error.  This is mostly
// for times when it would be unsafe to call MsgBox() due to the possibility that it would
// make the situation even worse.
{
	if (!aBuf) aBuf = "";
	if (mNIC.hWnd) // Tray icon is installed.
		Shell_NotifyIcon(NIM_DELETE, &mNIC); // Remove it.
	if (*aBuf)
	{
		char buf[1024];
		// No more than size-1 chars will be written and string will be terminated:
		snprintf(buf, sizeof(buf), "Critical Error: %s\n\n" WILL_EXIT, aBuf);
		// To avoid chance of more errors, don't use MsgBox():
		MessageBox(g_hWnd, buf, NAME_PV, MB_OK | MB_SETFOREGROUND | MB_APPLMODAL);
	}
	Hotkey::AllDestructAndExit(*aBuf ? CRITICAL_ERROR : ExitCode); // Terminate the application.
	// Not as reliable: PostQuitMessage(CRITICAL_ERROR);
}



int Script::LoadFromFile()
// Returns the number of non-comment lines that were loaded, or -1 on error.
// Use double-colon as delimiter to set these apart from normal labels.
// The main reason for this is that otherwise the user would have to worry
// about a normal label being unintentionally valid as a hotkey, e.g.
// "Shift:" might be a legitimate label that the user forgot is also
// a valid hotkey:
#define HOTKEY_FLAG "::"
{
	if (!mFileSpec || !*mFileSpec) return -1;

	// Future: might be best to put a stat() in here for better handling.
	FILE *fp = fopen(mFileSpec, "r");
	if (!fp)
	{
		int response = MsgBox("Default script file can't be opened.  Create it now?", MB_YESNO);
		if (response != IDYES)
			return 0;
		FILE *fp2 = fopen(mFileSpec, "a");
		if (!fp2)
		{
			MsgBox("Could not create file, perhaps because the current directory is read-only"
				" or has insufficient permissions.");
			return -1;
		}
		fprintf(fp2, "; " NAME_P " script file\n"
			"\n"
			//"; Uncomment out the below line to try out a sample of an Alt-tab\n"
			//"; substitute (currently not supported in Win9x):\n"
			//";RControl & RShift::AltTab\n"
			"; Sample hotkey:\n"
			"#z::  ; This hotkey is Win-Z (hold down Windows key and press Z).\n"
			"MsgBox, Hotkey was pressed.`n`nNote: MsgBox has a new single-parameter mode now."
				"  The title of this window defaults to the script's filename.\n"
			"return\n"
			"\n"
			"; After you finish editing this file, save it and run the EXE again\n"
			"; (it will open files of this name by default).\n"
			);
		fclose(fp2);
		// One or both of the below would probably fail if mFileSpec ever has spaces in it
		// (since it's passed as the entire param string).  If that ever happens, enclosing
		// the filename in double quotes should do the trick:
		if (!ActionExec("edit", mFileSpec, mFileDir, false))
			if (!ActionExec("Notepad.exe", mFileSpec, mFileDir, false))
			{
				MsgBox("The new config file was created, but could not be opened with the default editor or with Notepad.");
				return -1;
			}
		// future: have it wait for the process to close, then try to open the config file again:
		return 0;
	}

	// File is now open, read lines from it.

	// <buf> should be no larger than LINE_SIZE because some later functions rely upon that:
	char buf[LINE_SIZE], *hotkey_flag, *cp;
	HookActionType hook_action;
	size_t buf_length;
	bool is_label, section_comment = false;
	for (mFileLineCount = 0, mIsReadyToExecute = false;;) // Init in case this func ever called more than once.
	{
		mCurrLine = NULL;  // To signify that we're in transition, trying to load a new one.
		if (   -1 == (buf_length = GetLine(buf, (int)(sizeof(buf) - 1), fp))   )
			break;
		++mFileLineCount; // Keep track of the phyiscal line number in the file for debugging purposes.
		if (!buf_length)
			continue;

		if (section_comment) // Look for the uncomment-flag.
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
		else
			if (!strncmp(buf, "/*", 2))
			{
				section_comment = true;
				continue; // It's now commented out, so the rest of this line is ignored.
			}

		// Note that there may be an action following the HOTKEY_FLAG (on the same line).
		hotkey_flag = strstr(buf, HOTKEY_FLAG);
		is_label = (hotkey_flag != NULL);
		if (is_label) // It's a label and a hotkey.
		{
			*hotkey_flag = '\0'; // Terminate so that buf is now the label itself.
			hotkey_flag += strlen(HOTKEY_FLAG);  // Now hotkey_flag is the hotkey's action, if any.
			ltrim(hotkey_flag); // Has already been rtrimmed by GetLine().
			rtrim(buf); // Has already been ltrimmed.
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
			if (mFirstLabel == NULL)
				if (AddLine(ACT_RETURN) != OK)
					return CloseAndReturn(fp, -1);
			if (AddLabel(buf) != OK) // Always add a label before adding the first line of its section.
				return CloseAndReturn(fp, -1);
			if (*hotkey_flag) // This hotkey's action is on the same line as its label.
			{
				if (!stricmp(hotkey_flag, "AltTab"))
					hook_action = HOTKEY_ID_ALT_TAB;
				else if (!stricmp(hotkey_flag, "ShiftAltTab"))
					hook_action = HOTKEY_ID_ALT_TAB_SHIFT;
				else if (!stricmp(hotkey_flag, "AltTabMenu"))
					hook_action = HOTKEY_ID_ALT_TAB_MENU;
				else if (!stricmp(hotkey_flag, "AltTabAndMenu"))
					hook_action = HOTKEY_ID_ALT_TAB_AND_MENU;
				else if (!stricmp(hotkey_flag, "AltTabMenuDismiss"))
					hook_action = HOTKEY_ID_ALT_TAB_MENU_DISMISS;
				else
					hook_action = 0;
				// Don't add the alt-tabs as a line, since it has no meaning as a script command.
				// But do put in the Return regardless, in case this label is ever jumped to
				// via Goto/Gosub:
				if (!hook_action)
					if (ParseAndAddLine(hotkey_flag) != OK)
						return CloseAndReturn(fp, -1);
				// Also add a Return that's implicit for a single-line hotkey:
				if (AddLine(ACT_RETURN) != OK)
					return CloseAndReturn(fp, -1);
			}
			else
				hook_action = 0;
			// Set the new hotkey will jump to this label to begin execution:
			if (Hotkey::AddHotkey(mLastLabel, hook_action) != OK)
				return CloseAndReturn(fp, -1);
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
				return CloseAndReturn(fp, -1);
			continue;
		}
		// It's not a label.
		if (*buf == '#')
		{
			switch(IsPreprocessorDirective(buf))
			{
			case CONDITION_TRUE:
				continue;
			case FAIL:
				return CloseAndReturn(fp, -1); // It already reported the error.
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
				return CloseAndReturn(fp, -1);
		}
		else // This line is an ELSE.
		{
			// Add the ELSE directly rather than calling ParseAndAddLine() because that function
			// would resolve escape sequences throughout the entire length of <buf>, which we
			// don't want because we wouldn't have access to the corresponding literal-map to
			// figure out the proper use of escaped characters:
			if (AddLine(ACT_ELSE) != OK)
				return CloseAndReturn(fp, -1);
			action_end = omit_leading_whitespace(action_end); // Now action_end is the word after the ELSE.
			if (*action_end && ParseAndAddLine(action_end) != OK)
				return CloseAndReturn(fp, -1);
			// Otherwise, there was either no same-line action or the same-line action was successfully added,
			// so do nothing.
		}
	}
	fclose(fp);

	if (!mLineCount)
		return mLineCount;

	// Rather than do this, which seems kinda nasty if ever someday support same-line
	// else actions such as "else return", just add two EXITs to the end of every script.
	// That way, if the first EXIT added accidentally "corrects" an actionless ELSE
	// or IF, the second one will serve as the anchoring end-point (mRelatedLine) for that
	// IF or ELSE.  In other words, since we never want mRelatedLine to be NULL, the should
	// make absolutely sure of that:
	//if (mLastLine->mActionType == ACT_ELSE ||
	//	ACT_IS_IF(mLastLine->mActionType)
	//	...
	++mFileLineCount;
	if (AddLine(ACT_EXIT) != OK) // First exit.
		return -1;

	// Even if the last line of the script is already ACT_EXIT, always add another
	// one in case the script ends in a label.  That way, every label will have
	// a non-NULL target, which simplifies other aspects of script execution.
	// Making sure that all scripts end with an EXIT ensures that if the script
	// file ends with ELSEless IF or an ELSE, that IF's or ELSE's mRelatedLine
	// will be non-NULL, which further simplifies script execution:
	++mFileLineCount;
	if (AddLine(ACT_EXIT) != OK) // Second exit to guaranty non-NULL mRelatedLine(s).
		return -1;

	// Always do the blocks before the If/Else's because If/Else may rely on blocks:
	if (PreparseBlocks(mFirstLine) != NULL)
		if (PreparseIfElse(mFirstLine) != NULL)
		{
			// Use FindOrAdd, not Add, because the user may already have added it simply by
			// referring to it in the script:
			if (   !(g_ErrorLevel = FindOrAddVar("ErrorLevel"))   )
				return -1; // Error.  Above already displayed it for us.
			// Initialize the var state to zero right before running anything in the script:
			g_ErrorLevel->Assign(ERRORLEVEL_NONE);
			mIsReadyToExecute = true;

			// Initialize the random number generator:
			// Note: On 32-bit hardware, the generator module uses only 2506 bytes of static
			// data, so it doesn't seem worthwhile to put it in a class (so that the mem is
			// only allocated on first use of the generator).
			// This part is taken from the AutoIt3 source.  No comments were given there,
			// so I'm not sure if this initialization method is better than using
			// GetTickCount(), but I doubt it matters as long as it's at least 99.9999%
			// likely to be a different seed every time the program starts:
			struct _timeb timebuffer;
			_ftime(&timebuffer);
			init_genrand(timebuffer.millitm * (int)time(NULL));

			return mLineCount;
		}
	return -1; // Error.
}



inline int Script::CloseAndReturn(FILE *fp, int aReturnValue)
// Small inline to make LoadFromFile() code cleaner.
{
	fclose(fp);
	return aReturnValue;
}



size_t Script::GetLine(char *aBuf, int aMaxCharsToRead, FILE *fp)
{
	if (!aBuf || !fp) return -1;
	if (aMaxCharsToRead <= 0) return 0;
	if (feof(fp)) return -1; // Previous call to this function probably already read the last line.
	if (fgets(aBuf, aMaxCharsToRead, fp) == NULL) // end-of-file or error
	{
		*aBuf = '\0';  // Reset since on error, contents added by fgets() are indeterminate.
		return -1;
	}
	size_t aBuf_length = strlen(aBuf);
	if (!aBuf_length)
		return 0;
	if (aBuf[aBuf_length-1] == '\n')
		aBuf[--aBuf_length] = '\0';
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
			if (IS_SPACE_OR_TAB(*prevp)) // consider it to be a valid comment flag
			{
				*prevp = '\0';
				rtrim(aBuf); // Since it's our responsibility to return a fully trimmed string.
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
	char *cp;
	// Use strnicmp() so that a match is found as long as aBuf starts with the string in question.
	// e.g. so that "#SingleInstance, on" will still work too, but
	// "#a::run, something, "#SingleInstance" (i.e. a hotkey) will not be falsely detected
	// due to using a more lenient function such as stristr():
	#define IF_IS_DIRECTIVE_MATCH(directive) if (!strnicmp(aBuf, directive, strlen(directive)))
	IF_IS_DIRECTIVE_MATCH("#SingleInstance")
	{
		g_AllowOnlyOneInstance = true;
		return CONDITION_TRUE;
	}
	IF_IS_DIRECTIVE_MATCH("#AllowSameLineComments")  // i.e. There's no way to turn it off, only on.
	{
		g_AllowSameLineComments = true;
		return CONDITION_TRUE;
	}
	IF_IS_DIRECTIVE_MATCH("#InstallKeybdHook")
	{
		Hotkey::RequireHook(HOOK_KEYBD);
		return CONDITION_TRUE;
	}
	IF_IS_DIRECTIVE_MATCH("#InstallMouseHook")
	{
		Hotkey::RequireHook(HOOK_MOUSE);
		return CONDITION_TRUE;
	}
	IF_IS_DIRECTIVE_MATCH("#HotkeyModifierTimeout")
	{
		#define RETURN_IF_NO_CHAR \
		if (   !(cp = StrChrAny(aBuf, end_flags))   )\
			return CONDITION_TRUE;\
		if (   !*(cp = omit_leading_whitespace(cp))   )\
			return CONDITION_TRUE;
		RETURN_IF_NO_CHAR
		g_HotkeyModifierTimeout = atoi(cp);  // cp was set to the right position by the above macro
		return CONDITION_TRUE;
	}
	IF_IS_DIRECTIVE_MATCH("#HotkeyInterval")
	{
		RETURN_IF_NO_CHAR
		g_HotkeyThrottleInterval = atoi(cp);  // cp was set to the right position by the above macro
		if (g_HotkeyThrottleInterval < 10) // sanity check
			g_HotkeyThrottleInterval = 10;
		return CONDITION_TRUE;
	}
	IF_IS_DIRECTIVE_MATCH("#MaxHotkeysPerInterval")
	{
		RETURN_IF_NO_CHAR
		g_MaxHotkeysPerInterval = atoi(cp);  // cp was set to the right position by the above macro
		if (g_MaxHotkeysPerInterval <= 0) // sanity check
			g_MaxHotkeysPerInterval = 1;
		return CONDITION_TRUE;
	}

	// For the below series, it seems okay to allow the comment flag to contain other reserved chars,
	// such as DerefChar, since comments are evaluated, and then taken out of the game at an earlier
	// stage than DerefChar and the other special chars.
	IF_IS_DIRECTIVE_MATCH("#CommentFlag")
	{
		RETURN_IF_NO_CHAR
		if (!*(cp + 1))  // i.e. the length is 1
		{
			// Don't allow '#' since it's the preprocessor directive symbol being used here.
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
	IF_IS_DIRECTIVE_MATCH("#EscapeChar")
	{
		RETURN_IF_NO_CHAR
		if (   *cp == '#' || *cp == g_DerefChar || *cp == g_delimiter
			|| (g_CommentFlagLength == 1 && *cp == *g_CommentFlag)   )
			return ScriptError(ERR_DEFINE_CHAR);
		g_EscapeChar = *cp;
		return CONDITION_TRUE;
	}
	IF_IS_DIRECTIVE_MATCH("#DerefChar")
	{
		RETURN_IF_NO_CHAR
		if (   *cp == '#' || *cp == g_EscapeChar || *cp == g_delimiter
			|| (g_CommentFlagLength == 1 && *cp == *g_CommentFlag)   )
			return ScriptError(ERR_DEFINE_CHAR);
		g_DerefChar = *cp;
		return CONDITION_TRUE;
	}
	IF_IS_DIRECTIVE_MATCH("#Delimiter")
	{
		RETURN_IF_NO_CHAR
		if (   *cp == '#' || *cp == g_EscapeChar || *cp == g_DerefChar
			|| (g_CommentFlagLength == 1 && *cp == *g_CommentFlag)   )
			return ScriptError(ERR_DEFINE_CHAR);
		g_delimiter = *cp;
		return CONDITION_TRUE;
	}

	// Otherwise:
	return CONDITION_FALSE;
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
		return ScriptError("AddLabel(): Out of memory.");
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
	if (!aLineText || !*aLineText)
		return ScriptError("ParseAndAddLine() called incorrectly." PLEASE_REPORT);

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
	char *action_args = end_marker + 1;
	action_args = omit_leading_whitespace(action_args);
	// Now action_args is either the first delimiter or the first parameter (if it optional first
	// delimiter was omitted):
	if (*action_args == g_delimiter)
		// Find the start of the next token (or its ending delimiter if the token is blank such as ", ,"):
		for (++action_args; IS_SPACE_OR_TAB(*action_args); ++action_args);
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
	if (action_type == ACT_INVALID && old_action_type == OLD_INVALID)
		if (   (action_type = ConvertActionType(action_name)) == ACT_INVALID   )
			old_action_type = ConvertOldActionType(action_name);

	/////////////////////////////////////////////////////////////////////////////
	// Special handling for ACT_ASSIGN/ADD/SUB/MULT/DIV and IFEQUAL/GREATER/LESS.
	/////////////////////////////////////////////////////////////////////////////
	if (action_type == ACT_INVALID && old_action_type == OLD_INVALID)
	{
		// No match found, but is it a special type of action?
		if (!stricmp(action_name, "IF"))
		{
			char *operation = StrChrAny(action_args, "><!=");
			if (!operation)
				return ScriptError("Although this line is an IF, it lacks operator symbol(s).", aLineText);
				// Note: User can use whitespace to differentiate a literal symbol from
				// part of an operator, e.g. if var1 < =  <--- char is literal
			switch (*operation)
			{
			case '=': // But don't allow == to be "Equals" since the 2nd '=' might be literal.
				action_type = ACT_IFEQUAL;
				break;
			case '<':
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
					return ScriptError("When used this way, the symbol must be \"!=\" not \"!\".", aLineText);
				break;
			} // switch()
			// Set things up to be parsed as args later on:
			*operation = g_delimiter;
		}
		else // The action type is something other than an IF.
		{
			if (*action_args == '=')
				action_type = ACT_ASSIGN;
			else if (*action_args == '+' && *(action_args + 1) == '=')
				action_type = ACT_ADD;
			else if (*action_args == '-' && *(action_args + 1) == '=')
				action_type = ACT_SUB;
			else if (*action_args == '*' && *(action_args + 1) == '=')
				action_type = ACT_MULT;
			else if (*action_args == '/' && *(action_args + 1) == '=')
				action_type = ACT_DIV;
			if (action_type != ACT_INVALID)
			{
				// Set things up to be parsed as args later on:
				*action_args = g_delimiter;
				if (action_type != ACT_ASSIGN)
					*(action_args + 1) = ' ';  // Remove the "=" from consideration.
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

	////////////////////////////////////////////////////////////
	// Parse the parmeter string into a list of separate params.
	////////////////////////////////////////////////////////////
	// MaxParams has already been verified as being <= MAX_ARGS.
	// Any comma-delimited items beyond MaxParams will be included in a lump with the last param:
	int nArgs, mark;
	char *arg[MAX_ARGS], *arg_map[MAX_ARGS];
	ActionTypeType subaction_type = ACT_INVALID; // Must init these.
	ActionTypeType suboldaction_type = OLD_INVALID;
	char subaction_name[MAX_VAR_NAME_LENGTH + 1], *subaction_end_marker = NULL, *subaction_start = NULL;
	for (nArgs = mark = 0; action_args[mark] && nArgs < this_action->MaxParams; ++nArgs)
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
		if (nArgs == this_action->MaxParams - 1)
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
		for (; action_args[mark]; ++mark)
		{
			if (action_args[mark] == g_delimiter && !literal_map[mark])  // Match found: a non-literal delimiter.
			{
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
		default:
			return ScriptError("Unhandled Old-Command.", action_name);
		} // switch()
	}

	//////////////////////////////////////////////////////////////////////////////////////////////////
	// Handle AutoIt2-style IF-statements (i.e. the IF's action is on the same line as the condition).
	//////////////////////////////////////////////////////////////////////////////////////////////////
	// The check below: Don't bother if this IF (e.g. IfWinActive) has zero params or if the
	// subaction was already found above:
	if (nArgs && !subaction_type && !suboldaction_type && ACT_IS_IF(action_type))
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
	char end_flags[] = {' ', g_delimiter, '\t', '<', '>', '=', '+', '-', '*', '/', '!', '\0'}; // '\0' must be last.
	char *end_marker = StrChrAny(aBufSource, end_flags);
	if (end_marker) // Found a delimiter.
		if (end_marker > aBufSource) // The delimiter isn't very first char in aBufSource.
			--end_marker;
		else
		{
			// aBufSource starts with a delimiter (can't be whitespace since caller was
			// supposed to have already trimmed that?): probably a syntax error.
			if (aDisplayErrors)
				ScriptError("GetActionType(): Lines should not start with a delimiter.", aBufSource);
			return NULL;
		}
	else // No delimiter found, so set end_marker to the location of the last char in string.
		end_marker = aBufSource + strlen(aBufSource) - 1;
	// Now end_marker is the character just prior to the first delimiter or whitespace.
	end_marker = omit_trailing_whitespace(aBufSource, end_marker);
	size_t action_name_length = end_marker - aBufSource + 1;
	if (action_name_length < 1)
	{
		// Probably impossible due to trimming by the caller and all the other checks above:
		if (aDisplayErrors)
			ScriptError("GetActionType(): Parsing Error", aBufSource);
		return NULL;
	}
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
		return ScriptError("AddLine() called incorrectly.", aArgc > 0 ? aArg[0] : "");

	char error_msg[1024];
	ArgPurposeType arg_purpose;
	Var *target_var;
	int value;  // For temp use during validation.

	//////////////////////////////////////////////////////////
	// Build the new arg list in dynamic memory.
	// The allocated structs will be attached to the new line.
	//////////////////////////////////////////////////////////
	DerefType deref[MAX_DEREFS_PER_ARG];  // Will be used to temporarily store the var-deref locations in each arg.
	int deref_count;  // How many items are in deref array.
	ArgType *new_arg;  // We will allocate some dynamic memory for this, then hang it onto the new line.
	size_t deref_string_length;
	if (!aArgc)
		new_arg = NULL;  // Just need an empty array in this case.
	else
	{
		if (   !(new_arg = (ArgType *)SimpleHeap::Malloc(aArgc * sizeof(ArgType)))   )
			return ScriptError("AddLine(): Out of memory.");
		int i, j;
		for (i = 0; i < aArgc; ++i)
		{
			////////////////
			// FOR EACH ARG:
			////////////////

			// Before allocating memory for this Arg's text, first check if it's a pure
			// variable.  If it is, we store it differently (and there's no need to resolve
			// escape sequences in these cases, since var names can't contain them):
			arg_purpose = Line::ArgIsVar(aActionType, i);
			// Since some vars are optional, the below allows them all to be blank or
			// not present in the arg list.  If mandatory var is blank at this stage,
			// it's okay because all mandatory args are validated to be non-blank elsewhere:
			if (arg_purpose != IS_NOT_A_VAR && aArgc > i && *aArg[i])
			{
				if (   !(target_var = FindOrAddVar(aArg[i]))   )
					return FAIL;  // The above already displayed the error.
				// If this action type is something that modifies the contents of the var, ensure the var
				// isn't a special/reserved one:
				if (arg_purpose == IS_OUTPUT_VAR && VAR_IS_RESERVED(target_var))
					return ScriptError(ERR_VAR_IS_RESERVED, aArg[i]);
				// Rather than removing this arg from the list altogether -- which would distrub
				// the ordering and hurt the maintainability of the code -- the next best thing
				// in terms of saving memory is to store an empty string in place of the arg's
				// text if that arg is a pure variable (i.e. since the name of the variable is already
				// stored in the Var object, we don't need to store it twice):
				new_arg[i].text = arg_purpose; // Store a special, constant pointer value to flag it as a var.
				new_arg[i].deref = (DerefType *)target_var;
				continue;
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
			// of new_arg[i] to be modifiable (perhaps some obscure API calls require
			// modifiable strings?) can malloc a single-char for to contain
			// the empty string:
			if (   !(new_arg[i].text = SimpleHeap::Malloc(aArg[i]))   )
				return FAIL;  // It already displayed the error for us.


			////////////////////////////////////////////////////
			// Build the list of dereferenced vars for this arg.
			////////////////////////////////////////////////////
			// Now that any escaped g_DerefChars have been marked, scan new_arg.text to
			// determine where the variable dereferences are (if any).  In addition to helping
			// runtime performance, this also serves to validate the script at load-time
			// so that some errors can be caught early.  Note: new_arg[i].text is scanned rather
			// than aArg[i] because we want to establish pointers to the correct area of
			// memory:
			deref_count = 0;  // Init for each arg.
			// For each dereference:
			for (j = 0; ; ++j)  // Increment to skip over the symbol just found by the inner for().
			{
				// Find next non-literal g_DerefChar:
				for (; new_arg[i].text[j]
					&& (new_arg[i].text[j] != g_DerefChar || (aArgMap && aArgMap[i] && aArgMap[i][j])); ++j);
				if (!new_arg[i].text[j])
					break;
				// else: Match was found; this is the deref's open-symbol.
				if (deref_count >= MAX_DEREFS_PER_ARG)
					return ScriptError("The maximum number of variable dereferences has been"
						" exceeded in this parameter.", new_arg[i].text);
				deref[deref_count].marker = new_arg[i].text + j;  // Store the deref's starting location.
				// Find next g_DerefChar, even if it's a literal.
				for (++j; new_arg[i].text[j] && new_arg[i].text[j] != g_DerefChar; ++j);
				if (!new_arg[i].text[j])
					return ScriptError("This parameter contains a variable name"
						" that is missing its ending dereference symbol.", new_arg[i].text);
				// Otherwise: Match was found; this should be the deref's close-symbol.
				if (aArgMap && aArgMap[i] && aArgMap[i][j])  // But it's mapped as literal g_DerefChar.
					return ScriptError("This parmeter contains a variable name with"
						" an escaped dereference symbol, which is not allowed.", new_arg[i].text);
				deref_string_length = new_arg[i].text + j - deref[deref_count].marker + 1;
				if (deref_string_length - 2 > MAX_VAR_NAME_LENGTH) // -2 for the opening & closing g_DerefChars
					return ScriptError("This parmeter contains a variable name that is too long."
						, new_arg[i].text);
				deref[deref_count].length = (DerefLengthType)deref_string_length;
				if (   !(deref[deref_count].var = FindOrAddVar(deref[deref_count].marker + 1
					, deref[deref_count].length - 2))   )
					// The called function already displayed the error.
					return FAIL;
				++deref_count;
			} // for each dereference.

			///////////////////
			// More validation.
			///////////////////
			switch(aActionType)
			{
			case ACT_DETECTHIDDENWINDOWS:
			case ACT_DETECTHIDDENTEXT:
			case ACT_SETSTORECAPSLOCKMODE:
			case ACT_AUTOTRIM:
			case ACT_STRINGCASESENSE:
				if (i != 0) break;  // Should never happen in these cases.
				if (!deref_count)
					if (!Line::ConvertOnOff(new_arg[i].text))
						return ScriptError(ERR_ON_OFF, g_act[aActionType].Name);
				break;

			case ACT_SUSPEND:
				if (i != 0) break;  // Should never happen in this cases.
				if (!deref_count)
					if (!Line::ConvertOnOffTogglePermit(new_arg[i].text))
						return ScriptError(ERR_ON_OFF_TOGGLE_PERMIT, g_act[aActionType].Name);
				break;

			case ACT_PAUSE:
				if (i != 0) break;  // Should never happen in this cases.
				if (!deref_count)
					if (!Line::ConvertOnOffToggle(new_arg[i].text))
						return ScriptError(ERR_ON_OFF_TOGGLE, g_act[aActionType].Name);
				break;

			case ACT_SETNUMLOCKSTATE:
			case ACT_SETSCROLLLOCKSTATE:
			case ACT_SETCAPSLOCKSTATE:
				if (i != 0) break;  // Should never happen in these cases.
				if (!deref_count)
					if (!Line::ConvertOnOffAlways(new_arg[i].text))
						return ScriptError(ERR_ON_OFF_ALWAYS, g_act[aActionType].Name);
				break;

			case ACT_STRINGMID:
			case ACT_FILEREADLINE:
				if (i != 2) break; // i.e. 3rd param
				if (!deref_count)
				{
					value = atoi(new_arg[i].text);
					if (value <= 0)
						// This error is caught at load-time, but at runtime it's not considered
						// an error (i.e. if a variable resolves to zero or less, StringMid will
						// automatically consider it to be 1, though FileReadLine should consider
						// it an error):
						return ScriptError("The 3rd parameter be greater than zero.", new_arg[i].text);
				}
				break;

			case ACT_SOUNDSETWAVEVOLUME:
				if (i != 0) break; // i.e. 1st param
				if (!deref_count)
				{
					value = atoi(new_arg[i].text);
					if (value < 0 || value > 100)
						return ScriptError(ERR_PERCENT, new_arg[i].text);
				}
				break;

			case ACT_PIXELSEARCH:
				if (i != 7) break; // i.e. 8th param
				if (!deref_count)
				{
					value = atoi(new_arg[i].text);
					if (value < 0 || value > 255)
						return ScriptError("Parameter #8 must be number between 0 and 255, or a dereferenced variable.", new_arg[i].text);
				}
				break;

			case ACT_MOUSEMOVE:
				if (i != 2) break; // i.e. 3rd param
				#define VALIDATE_MOUSE_SPEED if (!deref_count)\
				{\
					value = atoi(new_arg[i].text);\
					if (value < 0 || value > MAX_MOUSE_SPEED)\
						return ScriptError(ERR_MOUSE_SPEED, new_arg[i].text);\
				}
				VALIDATE_MOUSE_SPEED
				break;
			case ACT_MOUSECLICK:
				if (i != 4) break; // i.e. 5th param
				VALIDATE_MOUSE_SPEED
				break;
			case ACT_MOUSECLICKDRAG:
				if (i != 5) break; // i.e. 6th param
				VALIDATE_MOUSE_SPEED
				break;
			case ACT_SETDEFAULTMOUSESPEED:
				if (i != 0) break; // i.e. 1st param
				VALIDATE_MOUSE_SPEED
				break;

			case ACT_FILECOPY:
			case ACT_FILEMOVE:
				if (i != 2) break; // i.e. 3rd param
				if (!deref_count)
				{
					value = atoi(new_arg[i].text);
					if (value != 0 && value != 1)
						return ScriptError("The 3rd parameter must be either blank, 0, 1, or a dereferenced variable."
							, new_arg[i].text);
				}
				break;
			case ACT_FILESELECTFILE:
				if (i != 1) break; // i.e. 2nd param
				if (!deref_count)
				{
					value = atoi(new_arg[i].text);
					if (value < 0 || value > 31)
						return ScriptError("The 2nd parameter must be either blank, a dereferenced variable,"
							" or a number between 0 and 31.", new_arg[i].text);
				}
				break;

			} // switch()

			//////////////////////////////////////////////////////////////
			// Allocate mem for this arg's list of dereferenced variables.
			//////////////////////////////////////////////////////////////
			if (deref_count)
			{
				// +1 for the "NULL-item" terminator:
				new_arg[i].deref = (DerefType *)SimpleHeap::Malloc((deref_count + 1) * sizeof(DerefType));
				if (!new_arg[i].deref)
					return ScriptError("AddLine(): Out of memory.");
				for (j = 0; j < deref_count; ++j)
				{
					new_arg[i].deref[j].marker = deref[j].marker;
					new_arg[i].deref[j].length = deref[j].length;
					new_arg[i].deref[j].var = deref[j].var;
				}
				// Terminate the list of derefs with a deref that has a NULL marker,
				// but only if the last one added isn't NULL (which it would be if it's
				// the fake-deref used to store an output-parameter's variable):
				if (deref_count && new_arg[i].deref[deref_count - 1].marker)
					new_arg[i].deref[deref_count].marker = NULL;
					// No need to increment deref_count since it's never used beyond this point.
			}
			else
				new_arg[i].deref = NULL;
		} // for each arg.
	} // else there are more than zero args.

	//////////////////////////////////////////////////////////////////////////////////////
	// Now the above has allocated some dynamic memory, the pointers to which we turn over
	// to Line's constructor so that they can be anchored to the new line.
	//////////////////////////////////////////////////////////////////////////////////////
	Line *line = new Line(g_script.mFileLineCount, aActionType, new_arg, aArgc);
	if (line == NULL)
		return ScriptError("AddLine(): Out of memory.");
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
	// Validate mandatory params and those that must be numeric.
	///////////////////////////////////////////////////////////////////
	if (!line->CheckForMandatoryArgs())
		return FAIL;  // It displayed the error for us.
	if (g_act[aActionType].NumericParams)
	{
		bool allow_negative;
		for (ActionTypeType *np = g_act[aActionType].NumericParams; *np; ++np)
		{
			if (line->mArgc >= *np)  // The arg exists.
			{
				if (!line->ArgHasDeref(*np)) // i.e. if it's a deref, we won't try to validate it now.
				{
					allow_negative = line->ArgAllowsNegative(*np);
					if (!IsPureNumeric(line->mArg[*np - 1].text, allow_negative))
					{
						if (aActionType == ACT_WINMOVE)
						{
							if (stricmp(line->mArg[*np - 1].text, "default"))
							{
								snprintf(error_msg, sizeof(error_msg), "\"%s\" requires parameter #%u to be"
									" either %snumeric, a dereferenced variable, blank, or the word Default."
									, g_act[line->mActionType].Name, *np, allow_negative ? "" : "non-negative ");
								return ScriptError(error_msg, line->mArg[*np - 1].text);
							}
							// else don't bother removing the word "default" since it won't save
							// any mem and since the command handler for ACT_WINMOVE must still
							// be able to understand the word "default" anyway, since that word
							// might be contained in a dereferenced variable.
						}
						else
						{
							snprintf(error_msg, sizeof(error_msg), "\"%s\" requires parameter #%u to be"
								" either %snumeric, blank (if blank is allowed), or a dereferenced variable."
								, g_act[line->mActionType].Name, *np, allow_negative ? "" : "non-negative ");
							return ScriptError(error_msg, line->mArg[*np - 1].text);
						}
					}
				}
			}
		} // for()
	} // if()

	///////////////////////////////////////////////////////////////////
	// Do any post-add validation & handling for specific action types.
	///////////////////////////////////////////////////////////////////
	switch(aActionType)
	{
	case ACT_SETTITLEMATCHMODE:
		if (line->mArgc > 0 && !line->ArgHasDeref(1) && !line->ConvertTitleMatchMode(LINE_RAW_ARG1))
			return ScriptError(ERR_TITLEMATCHMODE, LINE_RAW_ARG1);
		break;
	case ACT_MSGBOX:
		if (line->mArgc > 1) // i.e. this MsgBox is using the 4-param style.
			if (!line->ArgHasDeref(1)) // i.e. if it's a deref, we won't try to validate it now.
				if (!IsPureNumeric(LINE_RAW_ARG1))
					return ScriptError("When used with more than one parameter, MsgBox requires that"
						" the first parameter be numeric or a dereferenced variable.", LINE_RAW_ARG1);
		break;
	case ACT_IFMSGBOX:
		if (line->mArgc > 0 && !line->ArgHasDeref(1) && !line->ConvertMsgBoxResult(LINE_RAW_ARG1))
			return ScriptError(ERR_IFMSGBOX, LINE_RAW_ARG1);
		break;
	case ACT_GETKEYSTATE:
		if (line->mArgc > 1 && !line->ArgHasDeref(2) && !TextToVK(LINE_RAW_ARG2))
			return ScriptError("This is not a valid key or mouse button name.", LINE_RAW_ARG2);
		break;
	case ACT_DIV:
		if (!line->ArgHasDeref(2)) // i.e. if it's a deref, we won't try to validate it now.
			if (!atoi(LINE_RAW_ARG2))
				return ScriptError("This line would attempt to divide by zero.");
		break;
	case ACT_GROUPADD:
	case ACT_GROUPACTIVATE:
	case ACT_GROUPDEACTIVATE:
	case ACT_GROUPCLOSE:
	case ACT_GROUPCLOSEALL: // For all these, store a pointer to the group to help performance.
		if (!line->ArgHasDeref(1))
			if (   !(line->mAttribute = FindOrAddGroup(LINE_RAW_ARG1))   )
				return FAIL;  // The above already displayed the error.
		break;
	case ACT_RUN:
	case ACT_RUNWAIT:
		if (*LINE_RAW_ARG3 && !line->ArgHasDeref(3))
			if (line->ConvertRunMode(LINE_RAW_ARG3) == SW_SHOWNORMAL)
				return ScriptError(ERR_RUN_SHOW_MODE, LINE_RAW_ARG3);
		break;
	case ACT_MOUSECLICK:
	case ACT_MOUSECLICKDRAG:
		// Check that the button is valid (e.g. left/right/middle):
		if (!line->ArgHasDeref(1))
			if (!line->ConvertMouseButton(LINE_RAW_ARG1))
				return ScriptError(ERR_MOUSE_BUTTON, LINE_RAW_ARG1);
		if (!line->ValidateMouseCoords(LINE_RAW_ARG2, LINE_RAW_ARG3))
			return ScriptError(ERR_MOUSE_COORD, LINE_RAW_ARG2);
		if (aActionType == ACT_MOUSECLICKDRAG)
			if (!line->ValidateMouseCoords(LINE_RAW_ARG4, LINE_RAW_ARG5))
				return ScriptError(ERR_MOUSE_COORD, LINE_RAW_ARG4);
		break;
	case ACT_MOUSEMOVE:
		if (!line->ValidateMouseCoords(LINE_RAW_ARG1, LINE_RAW_ARG2))
			return ScriptError(ERR_MOUSE_COORD, LINE_RAW_ARG1);
		break;
	case ACT_REPEAT: // These types of loops are always "NORMAL".
		line->mAttribute = ATTR_LOOP_NORMAL;
	case ACT_LOOP:
		// If possible, determine the type of loop so that the preparser can better
		// validate some things:
		switch (line->mArgc)
		{
		case 0:
			line->mAttribute = ATTR_LOOP_NORMAL;
			break;
		case 1:
			if (line->ArgHasDeref(1)) // Impossible to know now what type of loop (only at runtime).
				line->mAttribute = ATTR_LOOP_UNKNOWN;
			else
				line->mAttribute = IsPureNumeric(LINE_RAW_ARG1) ? ATTR_LOOP_NORMAL : ATTR_LOOP_FILE;
			break;
		default:  // has 2 or more args.
			line->mAttribute = ATTR_LOOP_FILE;
			// Validate whatever we can rather than waiting for runtime validation:
			if (!line->ArgHasDeref(2) && Line::ConvertLoopMode(LINE_RAW_ARG2) == FILE_LOOP_INVALID)
				ScriptError(ERR_LOOP_FILE_MODE, LINE_RAW_ARG2);
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



Var *Script::FindOrAddVar(char *aVarName, size_t aVarNameLength)
// Returns the Var whose name matches aVarName.  If it doesn't exist, it is created.
{
	if (!aVarName || !*aVarName) return NULL;
	if (!aVarNameLength) // Caller didn't specify, so use the entire string.
		aVarNameLength = strlen(aVarName);
	for (Var *var = mFirstVar; var != NULL; var = var->mNextVar)
		if (!strlicmp(aVarName, var->mName, (UINT)aVarNameLength)) // Match found.
			return var;
	// Otherwise, no match found, so create a new var.
	if (AddVar(aVarName, aVarNameLength) != OK)
		return NULL;
	return mLastVar;
}



ResultType Script::AddVar(char *aVarName, size_t aVarNameLength)
// Returns OK or FAIL.
// The caller must already have verfied that this isn't a duplicate var.
{
	if (!aVarName || !*aVarName)
		return ScriptError("AddVar() called incorrectly.");
	if (!aVarNameLength) // Caller didn't specify, so use the entire string.
		aVarNameLength = strlen(aVarName);
	if (aVarNameLength > MAX_VAR_NAME_LENGTH)
		// Caller is supposed to prevent this.
		return ScriptError("AddVar(): Variable name is too long." PLEASE_REPORT);

	// Make a copy so that it can be modified for certain.  This will also serve
	// as the dynamic memory we pass to the constructor:
	char *new_name = SimpleHeap::Malloc(aVarNameLength + 1); // +1 for the terminator
	if (!new_name)
		return FAIL;  // It already displayed the error for us.
	strlcpy(new_name, aVarName, aVarNameLength + 1);  // Copies only the desired substring.
	if (!Var::ValidateName(new_name))
		return FAIL; // Above already displayed error for us.

	VarTypeType var_type;
	// Keeping the most common ones near the top helps performance a little:
	if (!stricmp(new_name, "clipboard")) var_type = VAR_CLIPBOARD;
	else if (!stricmp(new_name, "a_year")) var_type = VAR_YEAR;
	else if (!stricmp(new_name, "a_mon")) var_type = VAR_MON;
	else if (!stricmp(new_name, "a_mday")) var_type = VAR_MDAY;
	else if (!stricmp(new_name, "a_hour")) var_type = VAR_HOUR;
	else if (!stricmp(new_name, "a_min")) var_type = VAR_MIN;
	else if (!stricmp(new_name, "a_sec")) var_type = VAR_SEC;
	else if (!stricmp(new_name, "a_wday")) var_type = VAR_WDAY;
	else if (!stricmp(new_name, "a_yday")) var_type = VAR_YDAY;
	else if (!stricmp(new_name, "a_WorkingDir")) var_type = VAR_WORKINGDIR;
	else if (!stricmp(new_name, "a_ScriptName")) var_type = VAR_SCRIPTNAME;
	else if (!stricmp(new_name, "a_ScriptDir")) var_type = VAR_SCRIPTDIR;
	else if (!stricmp(new_name, "a_ScriptFullPath")) var_type = VAR_SCRIPTFULLPATH;
	else if (!stricmp(new_name, "a_NumBatchLines")) var_type = VAR_NUMBATCHLINES;
	else if (!stricmp(new_name, "a_OStype")) var_type = VAR_OSTYPE;
	else if (!stricmp(new_name, "a_OSversion")) var_type = VAR_OSVERSION;
	else if (!stricmp(new_name, "a_ThisHotkey")) var_type = VAR_THISHOTKEY;
	else if (!stricmp(new_name, "a_PriorHotkey")) var_type = VAR_PRIORHOTKEY;
	else if (!stricmp(new_name, "a_TimeSinceThisHotkey")) var_type = VAR_TIMESINCETHISHOTKEY;
	else if (!stricmp(new_name, "a_TimeSincePriorHotkey")) var_type = VAR_TIMESINCEPRIORHOTKEY;
	else if (!stricmp(new_name, "a_TickCount")) var_type = VAR_TICKCOUNT;
	else if (!stricmp(new_name, "a_Space")) var_type = VAR_SPACE;
	else var_type = VAR_NORMAL;

	Var *the_new_var = new Var(new_name, var_type);
	if (the_new_var == NULL)
		return ScriptError("AddVar(): Out of memory.");
	if (mFirstVar == NULL)
		mFirstVar = mLastVar = the_new_var;
	else
	{
		mLastVar->mNextVar = the_new_var;
		// This must be done after the above:
		mLastVar = the_new_var;
	}

	++mVarCount;
	return OK;
}



WinGroup *Script::FindOrAddGroup(char *aGroupName)
// Returns the Group whose name matches aGroupName.  If it doesn't exist, it is created.
{
	if (!aGroupName || !*aGroupName) return NULL;
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
	if (!aGroupName || !*aGroupName)
		return ScriptError("AddGroup() called incorrectly.");
	if (strlen(aGroupName) > MAX_VAR_NAME_LENGTH)
		return ScriptError("AddGroup(): Group name is too long.");
	if (!Var::ValidateName(aGroupName)) // Seems best to use same validation as var names.
		return FAIL; // Above already displayed error for us.

	char *new_name = SimpleHeap::Malloc(aGroupName);
	if (!new_name)
		return FAIL;  // It already displayed the error for us.

	WinGroup *the_new_group = new WinGroup(new_name);
	if (the_new_group == NULL)
		return ScriptError("AddGroup(): Out of memory.");
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



Line *Script::PreparseBlocks(Line *aStartingLine, int aFindBlockEnd, Line *aParentLine)
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
			// dynamically grow the stack to try to keep up:
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



Line *Script::PreparseIfElse(Line *aStartingLine, ExecUntilMode aMode, AttributeType aLoopType)
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
	AttributeType loop_type;
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

			loop_type = ATTR_NONE;
			if (aLoopType == ATTR_LOOP_FILE || line->mAttribute == ATTR_LOOP_FILE)
				// i.e. if either one is a file-loop, that's enough to establish
				// the fact that we're in a file loop.
				loop_type = ATTR_LOOP_FILE;
			else if (aLoopType == ATTR_LOOP_UNKNOWN || line->mAttribute == ATTR_LOOP_UNKNOWN)
				// ATTR_LOOP_UNKNOWN takes precedence over ATTR_LOOP_NORMAL because
				// we can't be sure if we're in a file loop, but it's correct to
				// assume that we are (otherwise, unwarranted syntax errors may be reported
				// later on in here).
				loop_type = ATTR_LOOP_UNKNOWN;
			else if (aLoopType == ATTR_LOOP_NORMAL || line->mAttribute == ATTR_LOOP_NORMAL)
				loop_type = ATTR_LOOP_NORMAL;

			// Check if the IF's action-line is something we want to recurse.  UPDATE: Always
			// recurse because other line types, such as Goto and Gosub, need to be preparsed
			// by this function even if they are the single-line actions of an IF or an ELSE:
			// Recurse this line rather than the next because we want
			// the called function to recurse again if this line is a ACT_BLOCK_BEGIN
			// or is itself an IF:
			line_temp = PreparseIfElse(line_temp, ONLY_ONE_LINE, loop_type);
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
				line = PreparseIfElse(line, ONLY_ONE_LINE, aLoopType);
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
			line = PreparseIfElse(line->mNextLine, UNTIL_BLOCK_END, aLoopType);
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
				 return line->PreparseError("Unexpected end-of-block (parsing single line).");
			if (UNTIL_BLOCK_END)
				// Return line rather than line->mNextLine because, if we're at the end of
				// the script, it's up to the caller to differentiate between that condition
				// and the condition where NULL is an error indicator.
				return line;
			// Otherwise, we found an end-block we weren't looking for.  This should be
			// impossible since the block pre-parsing already balanced all the blocks?
			return line->PreparseError("Unexpected end-of-block (parsing multiple lines).");
		case ACT_BREAK:
		case ACT_CONTINUE:
			if (!aLoopType)
				return line->PreparseError("This break or continue statement is not enclosed by a loop.");
			break;
		case ACT_FILESETDATEMODIFIED:
		case ACT_FILETOGGLEHIDDEN:
			if (aLoopType != ATTR_LOOP_FILE && aLoopType != ATTR_LOOP_UNKNOWN && !*LINE_RAW_ARG1)
				return line->PreparseError("When not enclosed in a file-loop, this command requires more parameters.");
			break;
		case ACT_GOTO:  // These two must be done here (i.e. *after* all the script lines have been added),
		case ACT_GOSUB: // so that labels both above and below each Gosub/Goto can be resolved.
			if (line->ArgHasDeref(1))
				// Since the jump-point contains a deref, it must be resolved at runtime:
				line->mRelatedLine = NULL;
			else
				if (!line->SetJumpTarget(false))
					return NULL; // Error was already displayed by called function.
			break;
		case ACT_GROUPADD: // This must be done here because it relies on all other lines already having been added.
			if (*LINE_RAW_ARG4 && !line->ArgHasDeref(4))
			{
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
		return mLastLine->PreparseError("The script ended while a block was still open."
			PLEASE_REPORT); // This is a bug because the preparser already verified all blocks are balanced.

	// If we were told to process a single line, we were recursed and it should have returned above,
	// so it's an error here (can happen if we were called with aStartingLine == NULL?):
	if (aMode == ONLY_ONE_LINE)
		return mLastLine->PreparseError("The script ended while an action was still expected.");

	// Otherwise, return something non-NULL to indicate success to the top-level caller:
	return mLastLine;
}


//-------------------------------------------------------------------------------------

// Init static vars:
Line *Line::sLog[] = {NULL};  // I think this initializes all the array elements.
int Line::sLogNext = 0;  // Start at the first element.
char *Line::sDerefBuf = NULL;  // Buffer to hold the values of any args that need to be dereferenced.
char *Line::sDerefBufMarker = NULL;
size_t Line::sDerefBufSize = 0;
char *Line::sArgDeref[MAX_ARGS]; // No init needed.
const char Line::sArgIsInputVar[1] = "";   // A special, constant pointer value we can use.
const char Line::sArgIsOutputVar[1] = "";  // A special, constant pointer value we can use.


ResultType Line::ExecUntil(ExecUntilMode aMode, modLR_type aModifiersLR, Line **apJumpToLine
	, WIN32_FIND_DATA *aCurrentFile)
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
	ResultType if_condition, result;

	for (Line *line = this; line != NULL;)
	{
		// If a previous command (line) had the clipboard open, perhaps because it directly accessed
		// the clipboard via Var::Contents(), we close it here:
		CLOSE_CLIPBOARD_IF_OPEN;
		g_script.mCurrLine = line;  // Simplifies error reporting when we get deep into function calls.
		line->Log();  // Maintains a circular queue of the lines most recently executed.

		// The below handles the message-loop checking regardless of whether
		// aMode is ONLY_ONE_LINE (i.e. recursed) or not (i.e. we're using
		// the for-loop to execute the script linearly):
		if (g.LinesPerCycle > 0 && g_script.mLinesExecutedThisCycle >= g.LinesPerCycle)
			// Sleep in between batches of lines, like AutoIt, to reduce
			// the chance that a maxed CPU will interfere with time-critical
			// apps such as games, video capture, or video playback.  Also,
			// check the message queue to see if the program has been asked
			// to exit by some outside force, or whether the user pressed
			// another hotkey to interrupt this one.  In the case of exit,
			// this call will never return.  Note: MsgSleep() will reset
			// mLinesExecutedThisCycle for us:
			MsgSleep(); // Exact length of sleep unimportant.

		// At this point, a pause may have been triggered either by the above MsgSleep()
		// or due to the action of a command (e.g. Pause, or perhaps tray menu "pause" was selected during Sleep):
		for (;;)
		{
			if (g.IsPaused)
				MsgSleep(INTERVAL_UNSPECIFIED, RETURN_AFTER_MESSAGES, false);
			else
				break;
		}

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
				result = line->mNextLine->ExecUntil(ONLY_ONE_LINE, aModifiersLR, &jump_to_line, aCurrentFile);
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
					line = line->mRelatedLine;  // Now line is either the IF's else or the end of the if-stmt.
					if (line == NULL)
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
				line = line->mRelatedLine;  // Set to IF's related line.
				if (line == NULL)
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
					result = line->mNextLine->ExecUntil(ONLY_ONE_LINE, aModifiersLR, &jump_to_line, aCurrentFile);
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
			if (line->mRelatedLine == NULL)
				if (!line->SetJumpTarget(true))
					return FAIL; // Error was already displayed by called function.
			// One or both of these lines can be NULL.  But the preparser should have
			// ensured that all we need to do is a simple compare to determine
			// whether this Goto should be handled by this layer or its caller
			// (i.e. if this Goto's target is not in our nesting level, it MUST be the
			// caller's responsibility to either handle it or pass it on to its
			// caller).
			if (aMode == ONLY_ONE_LINE || line->mParentLine != line->mRelatedLine->mParentLine)
			{
				if (apJumpToLine != NULL)
					*apJumpToLine = line->mRelatedLine; // Tell the caller to handle this jump.
				return OK;
			}
			// Otherwise, we will handle this Goto since it's in our nesting layer:
			line = line->mRelatedLine;
			break;  // Resume looping starting at the above line.
		case ACT_GOSUB:
			// A single gosub can cause an infinite loop if misused (i.e. recusive gosubs),
			// so be sure to do this to prevent the program from hanging:
			++g_script.mLinesExecutedThisCycle;
			if (line->mRelatedLine == NULL)
				if (!line->SetJumpTarget(true))
					return FAIL; // Error was already displayed by called function.
			// I'm pretty sure it's not valid for this call to ExecUntil() to tell us to jump
			// somewhere, because the called function, or a layer even deeper, should handle
			// the goto prior to returning to us?  So the last param is omitted:
			result = line->mRelatedLine->ExecUntil(UNTIL_RETURN, aModifiersLR, NULL, aCurrentFile);
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
			group->Activate(toupper(*LINE_ARG2) == 'R', NULL, (void **)&jump_to_line);
			if (jump_to_line)
			{
				if (!line->IsJumpValid(jump_to_line))
					// This check probably isn't necessary since IsJumpValid() is mostly
					// for Goto's.  But just in case the gosub's target label is some
					// crazy place:
					return FAIL;
				// This section is just like the Gosub code above, so maintain them together.
				result = jump_to_line->ExecUntil(UNTIL_RETURN, aModifiersLR, NULL, aCurrentFile);
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
			if (attr == ATTR_LOOP_UNKNOWN || attr == ATTR_NONE)
				// Since it couldn't be determined at load-time (probably due to derefs),
				// determine whether it's a file-loop or a normal/counter loop.
				// But don't change the value of line->mAttribute because that's our
				// indicator of whether this needs to be evaluated every time for
				// this particular loop (since the nature of the loop can change if the
				// contents of the variables dereferenced for this line change during runtime):
			{
				switch (line->mArgc)
				{
				case 0: attr = ATTR_LOOP_NORMAL; break;
				case 1: attr = IsPureNumeric(LINE_ARG1) ? ATTR_LOOP_NORMAL : ATTR_LOOP_FILE; break;
				case 2: attr = ATTR_LOOP_FILE; break;
				}
			}

			int iteration_limit = 0;
			bool is_infinite = line->mArgc < 1;
			if (!is_infinite)
				// Must be set to zero for ATTR_LOOP_FILE:
				iteration_limit = (attr == ATTR_LOOP_FILE) ? 0 : atoi(LINE_ARG1);

			if (line->mActionType == ACT_REPEAT && !iteration_limit)
				is_infinite = true;  // Because a 0 means infinite in AutoIt2 for the REPEAT command.

			FileLoopModeType file_loop_mode = (line->mArgc <= 1) ? FILE_LOOP_DEFAULT
				: ConvertLoopMode(LINE_ARG2);
			if (file_loop_mode == FILE_LOOP_INVALID)
				return line->LineError(ERR_LOOP_FILE_MODE ERR_ABORT, FAIL, LINE_ARG2);

			BOOL file_found = FALSE;
			HANDLE file_search = INVALID_HANDLE_VALUE;
			WIN32_FIND_DATA current_file;
			if (attr == ATTR_LOOP_FILE)
			{
				file_search = FindFirstFile(line->sArgDeref[0], &current_file);
				file_found = (file_search != INVALID_HANDLE_VALUE);
				for (; file_found && FileIsFilteredOut(current_file, file_loop_mode, line->sArgDeref[0])
					; file_found = FindNextFile(file_search, &current_file));
			}

			// Note: It seems best NOT to report warning if the loop iterates zero times
			// (e.g if no files are found by FindFirstFile() above), since that could
			// easily be an expected outcome.

			bool continue_main_loop = false;
			jump_to_line = NULL; // Init in case loop has zero iterations.
			for (int i = 0; is_infinite || file_found || i < iteration_limit; ++i)
			{
				// Preparser has ensured that every LOOP has a non-NULL next line.
				// aCurrentFile is sent as an arg so that more than one nested/recursive
				// file-loop can be supported:
				result = line->mNextLine->ExecUntil(ONLY_ONE_LINE, aModifiersLR, &jump_to_line
					, file_found ? &current_file : aCurrentFile); // inner loop takes precedence over outer.
				if (jump_to_line == line)
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
					continue_main_loop = true;
					break;
				}
				if (result == FAIL || result == EARLY_RETURN || result == EARLY_EXIT)
				{
					if (file_search != INVALID_HANDLE_VALUE)
					{
						FindClose(file_search);
						file_search = INVALID_HANDLE_VALUE;
					}
					return result;
				}
				if (jump_to_line != NULL && jump_to_line->mParentLine != line->mParentLine)
				{
					if (apJumpToLine != NULL) // In this case, it should always be non-NULL?
						*apJumpToLine = jump_to_line; // Tell the caller to handle this jump.
					if (file_search != INVALID_HANDLE_VALUE)
					{
						FindClose(file_search);
						file_search = INVALID_HANDLE_VALUE;
					}
					return OK;
				}
				if (result == LOOP_BREAK || jump_to_line != NULL)
					// In the case of jump_to_line != NULL: Above would already have returned
					// if the jump-point can't be handled by our layer, so this is ours.
					// This jump must be to somewhere outside of our loop's block (if it has one)
					// because otherwise the block's recursion layer (at a deeper level) would already
					// have handled it for us.  So the loop must be broken:
					break;
				// Otherwise, result is LOOP_CONTINUE (the current loop iteration was cut short)
				// or OK (the current iteration completed normally), so just continue on through
				// the loop.  But first do any end-of-iteration stuff:
				if (file_search != INVALID_HANDLE_VALUE)
					for (;;)
					{
						if (   !(file_found = FindNextFile(file_search, &current_file))   )
							break;
						if (FileIsFilteredOut(current_file, file_loop_mode, line->sArgDeref[0]))
							continue; // Ignore this one, get another one.
						else
							break;
					}
			}
			// The script's loop is now over.
			if (file_search != INVALID_HANDLE_VALUE)
			{
				FindClose(file_search);
				file_search = INVALID_HANDLE_VALUE;
			}
			if (continue_main_loop)
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
			if (jump_to_line != NULL)
			{
				// Since above didn't return, we're supposed to jump to the indicated place
				// and continue execution there:
				line = jump_to_line;
				break; // end this case of the switch().
			}
			// Since the above didn't return, either the loop has completed the specified number
			// of iterations or it was broken via "break".  In either case, we jump to the
			// line after our loop's structure and continue there:
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
			if (Hotkey::sHotkeyCount || Hotkey::HookIsActive())
				return EARLY_EXIT;  // It's "early" because only the very end of the script is the "normal" exit.
			else
				// This has been tested and it does return the error code indicated in ARG1, if present
				// (otherwise it returns 0, naturally) as expected:
				g_script.ExitApp(NULL, atoi(ARG1));  // Seems more reliable than PostQuitMessage().
		case ACT_EXITAPP: // Unconditional exit.
			g_script.ExitApp(NULL, atoi(ARG1));  // Seems more reliable than PostQuitMessage().
		case ACT_BLOCK_BEGIN:
			// Don't count block-begin/end against the total since they should be nearly instantaneous:
			//++g_script.mLinesExecutedThisCycle;
			// In this case, line->mNextLine is already verified non-NULL by the pre-parser:
			result = line->mNextLine->ExecUntil(UNTIL_BLOCK_END, aModifiersLR, &jump_to_line, aCurrentFile);
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
				return line->LineError("Unexpected end-of-block." PLEASE_REPORT ERR_ABORT);
			return OK; // It's the caller's responsibility to resume execution at the next line, if appropriate.
		case ACT_ELSE:
			// Shouldn't happen if the pre-parser and this function are designed properly?
			return line->LineError("This ELSE is unexpected." PLEASE_REPORT ERR_ABORT);
		default:
			++g_script.mLinesExecutedThisCycle;
			result = line->Perform(aModifiersLR, aCurrentFile);
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
	if (!ACT_IS_IF(mActionType))
		return LineError("EvaluateCondition() was called with a line that isn't a condition."
			PLEASE_REPORT ERR_ABORT);

	int if_condition;
	switch (mActionType)
	{
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
		if_condition = (WinActive(FOUR_ARGS) != NULL);
		break;
	case ACT_IFWINNOTACTIVE:
		if_condition = !WinActive(FOUR_ARGS);
		break;

	case ACT_IFEXIST:
		if_condition = DoesFilePatternExist(ARG1);
		break;
	case ACT_IFNOTEXIST:
		if_condition = !DoesFilePatternExist(ARG1);
		break;

	case ACT_IFINSTRING:
		#define STRING_SEARCH (g.StringCaseSense ? strstr(ARG1, ARG2) : stristr(ARG1, ARG2))
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
		#define BOTH_ARE_NUMERIC (*ARG1 && *ARG2 && IsPureNumeric(ARG1, true) && IsPureNumeric(ARG2, true))
		#define STRING_COMPARE (g.StringCaseSense ? strcmp(ARG1, ARG2) : stricmp(ARG1, ARG2))
		if (BOTH_ARE_NUMERIC)
			if_condition = atoi(ARG1) == atoi(ARG2);
		else
			if_condition = !STRING_COMPARE;
		break;
	case ACT_IFNOTEQUAL:
		if (BOTH_ARE_NUMERIC)
			if_condition = atoi(ARG1) != atoi(ARG2);
		else
			if_condition = STRING_COMPARE;
		break;
	case ACT_IFLESS:
		if (BOTH_ARE_NUMERIC)
			if_condition = atoi(ARG1) < atoi(ARG2);
		else
			if_condition = STRING_COMPARE < 0;
		break;
	case ACT_IFLESSOREQUAL:
		if (BOTH_ARE_NUMERIC)
			if_condition = atoi(ARG1) <= atoi(ARG2);
		else
			if_condition = STRING_COMPARE <= 0;
		break;
	case ACT_IFGREATER:
		if (BOTH_ARE_NUMERIC)
			if_condition = atoi(ARG1) > atoi(ARG2);
		else
			if_condition = STRING_COMPARE > 0;
		break;
	case ACT_IFGREATEROREQUAL:
		if (BOTH_ARE_NUMERIC)
			if_condition = atoi(ARG1) >= atoi(ARG2);
		else
			if_condition = STRING_COMPARE >= 0;
		break;
	case ACT_IFMSGBOX:
	{
		int mb_result = ConvertMsgBoxResult(ARG1);
		if (!mb_result)
			return LineError(ERR_IFMSGBOX ERR_ABORT, FAIL, ARG1);
		if_condition = (g.MsgBoxResult == mb_result);
		break;
	}
	default: // Should never happen, but return an error if it does.
		return LineError("EvaluateCondition(): Unhandled windowing action type." PLEASE_REPORT ERR_ABORT);
	}
	return if_condition ? CONDITION_TRUE : CONDITION_FALSE;
}



inline ResultType Line::Perform(modLR_type aModifiersLR, WIN32_FIND_DATA *aCurrentFile)
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
	VarSizeType space_needed; // For the commands that assign directly to an output var.
	ToggleValueType toggle;  // For commands that use on/off/neutral.
	int x, y;   // For mouse commands.
	int start_char_num, chars_to_extract;  // For String commands.
	size_t source_length; // For String commands.
	int math_result; // For math operations.
	vk_type vk; // For mouse commands and GetKeyState.
	HWND target_window;
	HANDLE running_process; // For RUNWAIT
	DWORD exit_code; // For RUNWAIT

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
		int wait_time = *ARG3 ? (1000 * atoi(ARG3)) : DEFAULT_WINCLOSE_WAIT;
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
		return IniDelete(THREE_ARGS);

	case ACT_REGREAD:
		return RegRead(ARG2, ARG3, ARG4, ARG5);
	case ACT_REGWRITE:
		return RegWrite(FIVE_ARGS);
	case ACT_REGDELETE:
		return RegDelete(THREE_ARGS);

	case ACT_SHUTDOWN:
		return Util_Shutdown(atoi(ARG1)) ? OK : FAIL;
	case ACT_SLEEP:
		MsgSleep(atoi(ARG1));
		return OK;
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
		g_ErrorLevel->Assign(SetEnvironmentVariable(ARG1, ARG2) ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);
		return OK;
	case ACT_ENVUPDATE:
	{
		// From the AutoIt3 source:
		// AutoIt3 uses SMTO_BLOCK (which prevents our thread from doing anything during the call)
		// vs. SMTO_NORMAL.  Since I'm not sure why, I'm leaving it that way for now:
		ULONG nResult;
		if (SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0, (LPARAM)"Environment", SMTO_BLOCK, 15000, &nResult))
			g_ErrorLevel->Assign(ERRORLEVEL_NONE);
		else
			g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		return OK;
	}
	case ACT_RUN:
		return g_script.ActionExec(ARG1, NULL, ARG2, true, ARG3);  // Be sure to pass NULL for 2nd param.
	case ACT_RUNWAIT:
		if (!g_script.ActionExec(ARG1, NULL, ARG2, true, ARG3, &running_process))
			return FAIL;
		// else fall through to the below.
	case ACT_CLIPWAIT:
	case ACT_WINWAIT:
	case ACT_WINWAITCLOSE:
	case ACT_WINWAITACTIVE:
	case ACT_WINWAITNOTACTIVE:
	{
		bool wait_indefinitely;
		int sleep_duration;
		DWORD start_time;
		if (   (mActionType != ACT_RUNWAIT && mActionType != ACT_CLIPWAIT && *ARG3)
			|| (mActionType == ACT_CLIPWAIT && *ARG1)   )
		{
			// Since the param containing the timeout value isn't blank, it must be numeric,
			// otherwise, the loading validation would have prevented the script from loading.
			wait_indefinitely = false;
			sleep_duration = atoi(mActionType == ACT_CLIPWAIT ? ARG1 : ARG3) * 1000; // Can be zero.
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
				if (WinActive(SAVED_WIN_ARGS))
				{
					DoWinDelay;
					return OK;
				}
				break;
			case ACT_WINWAITNOTACTIVE:
				if (!WinActive(SAVED_WIN_ARGS))
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
			case ACT_RUNWAIT:
				// Pretty nasty, but for now, nothing is done to prevent an infinite loop.
				// In the future, maybe OpenProcess() can be used to detect if a process still
				// exists (is there any other way?):
				// MSDN: "Warning: If a process happens to return STILL_ACTIVE (259) as an error code,
				// applications that test for this value could end up in an infinite loop."
				GetExitCodeProcess(running_process, &exit_code);
				if (exit_code != STATUS_PENDING)
				{
					CloseHandle(running_process);
					g_ErrorLevel->Assign((int)exit_code); // Use signed vs. unsigned, since that is more typical?
					return OK;
				}
				break;
			}

			// Must cast to int or any negative result will be lost due to DWORD type:
			if (wait_indefinitely || (int)(sleep_duration - (GetTickCount() - start_time)) > SLEEP_INTERVAL_HALF)
				// Last param 0 because we don't want it to restore the
				// current active window after the time expires (in case
				// our subroutine is suspended).  INTERVAL_UNSPECIFIED performs better:
				MsgSleep(INTERVAL_UNSPECIFIED, RETURN_AFTER_MESSAGES, false);
			else
			{
				g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Since it timed out, we override the default with this.
				return OK;  // Done waiting.
			}
		} // for()
		break; // Never executed, just for safety.
	}

	case ACT_WINMOVE:
		return mArgc > 2 ? WinMove(EIGHT_ARGS) : WinMove("", "", ARG1, ARG2);
	case ACT_WINMENUSELECTITEM:
		return WinMenuSelectItem(ELEVEN_ARGS);
	case ACT_CONTROLSEND:
		return ControlSend(SIX_ARGS, aModifiersLR);
	case ACT_CONTROLLEFTCLICK:
		return ControlLeftClick(FIVE_ARGS);
	case ACT_CONTROLFOCUS:
		return ControlFocus(FIVE_ARGS);
	case ACT_CONTROLSETTEXT:
		return ControlSetText(SIX_ARGS);
	case ACT_CONTROLGETTEXT:
		return ControlGetText(ARG2, ARG3, ARG4, ARG5, ARG6);
	case ACT_STATUSBARGETTEXT:
		return StatusBarGetText(ARG2, ARG3, ARG4, ARG5, ARG6);
	case ACT_STATUSBARWAIT:
		return StatusBarWait(EIGHT_ARGS);
	case ACT_WINSETTITLE:
		return mArgc > 1 ? WinSetTitle(FIVE_ARGS) : WinSetTitle("", "", ARG1);
	case ACT_WINGETTITLE:
		return WinGetTitle(ARG2, ARG3, ARG4, ARG5);
	case ACT_WINGETTEXT:
		return WinGetText(ARG2, ARG3, ARG4, ARG5);
	case ACT_WINGETPOS:
		return WinGetPos(ARG5, ARG6, ARG7, ARG8);
	case ACT_PIXELSEARCH:
		return PixelSearch(atoi(ARG3), atoi(ARG4), atoi(ARG5), atoi(ARG6), atoi(ARG7), atoi(ARG8));
	case ACT_PIXELGETCOLOR:
		return PixelGetColor(atoi(ARG2), atoi(ARG3));
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
		group->Deactivate(toupper(*ARG2) == 'R');  // Note: It will take care of DoWinDelay if needed.
		return OK;
	case ACT_GROUPCLOSE:
		if (   !(group = (WinGroup *)mAttribute)   )
			if (   !(group = g_script.FindOrAddGroup(ARG1))   )
				return FAIL;  // It already displayed the error for us.
		group->CloseAndGoToNext(toupper(*ARG2) == 'R');  // Note: It will take care of DoWinDelay if needed.
		return OK;
	case ACT_GROUPCLOSEALL:
		if (   !(group = (WinGroup *)mAttribute)   )
			if (   !(group = g_script.FindOrAddGroup(ARG1))   )
				return FAIL;  // It already displayed the error for us.
		group->CloseAll();  // Note: It will take care of DoWinDelay if needed.
		return OK;

	case ACT_STRINGLEFT:
		chars_to_extract = atoi(ARG3);
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
		return OUTPUT_VAR->Assign(ARG2, (VarSizeType)strnlen(ARG2, chars_to_extract));
	case ACT_STRINGRIGHT:
		chars_to_extract = atoi(ARG3);
		if (chars_to_extract < 0)
			chars_to_extract = 0;
		source_length = strlen(ARG2);
		if (chars_to_extract > (int)source_length)
			chars_to_extract = (int)source_length;
		// It will display any error that occurs:
		return OUTPUT_VAR->Assign(ARG2 + source_length - chars_to_extract, chars_to_extract);
	case ACT_STRINGMID:
		start_char_num = atoi(ARG3);
		if (start_char_num <= 0)
			// It's somewhat debatable, but it seems best not to report an error in this and
			// other cases.  The result here is probably enough to speak for itself, for script
			// debugging purposes:
			start_char_num = 1; // 1 is the position of the first char, unlike StringGetPos.
		chars_to_extract = atoi(ARG4);
		if (chars_to_extract < 0)
			chars_to_extract = 0;
		// Assign() is capable of doing what we want in this case:
		// It will display any error that occurs:
		if (strlen(ARG2) < (size_t)start_char_num)
			return OUTPUT_VAR->Assign();  // Set it to be blank in this case.
		else
			return OUTPUT_VAR->Assign(ARG2 + start_char_num - 1, chars_to_extract);
	case ACT_STRINGTRIMLEFT:
		chars_to_extract = atoi(ARG3);
		if (chars_to_extract < 0)
			chars_to_extract = 0;
		source_length = strlen(ARG2);
		if (chars_to_extract > (int)source_length) // This could be intentional, so don't display an error.
			chars_to_extract = (int)source_length;
		return OUTPUT_VAR->Assign(ARG2 + chars_to_extract, (VarSizeType)(source_length - chars_to_extract));
	case ACT_STRINGTRIMRIGHT:
		chars_to_extract = atoi(ARG3);
		if (chars_to_extract < 0)
			chars_to_extract = 0;
		source_length = strlen(ARG2);
		if (chars_to_extract > (int)source_length) // This could be intentional, so don't display an error.
			chars_to_extract = (int)source_length;
		return OUTPUT_VAR->Assign(ARG2, (VarSizeType)(source_length - chars_to_extract)); // It already displayed any error.
	case ACT_STRINGLOWER:
	case ACT_STRINGUPPER:
		space_needed = (VarSizeType)(strlen(ARG2) + 1);
		// Set up the var, enlarging it if necessary.  If the OUTPUT_VAR is of type VAR_CLIPBOARD,
		// this call will set up the clipboard for writing:
		if (OUTPUT_VAR->Assign(NULL, space_needed - 1) != OK)
			return FAIL;
		// Copy the input variable's text directly into the output variable:
		strlcpy(OUTPUT_VAR->Contents(), ARG2, space_needed);
		if (mActionType == ACT_STRINGLOWER)
			strlwr(OUTPUT_VAR->Contents());
		else
			strupr(OUTPUT_VAR->Contents());
		return OUTPUT_VAR->Close();  // In case it's the clipboard.
	case ACT_STRINGLEN:
		return OUTPUT_VAR->Assign((int)strlen(ARG2)); // It already displayed any error.
	case ACT_STRINGGETPOS:
	{
		g_ErrorLevel->Assign(ERRORLEVEL_NONE);  // Set default.
		bool search_from_the_right = (toupper(*ARG4) == 'R');
		int pos;  // An int, not size_t, so that it can be negative.
		if (!*ARG3) // It might be intentional, in obscure cases, to search for the empty string.
			pos = 0; // Empty string is always found immediately (first char from left).
		else
		{
			char *found;
			if (search_from_the_right)
				found = strrstr(ARG2, ARG3, g.StringCaseSense);
			else
				found = g.StringCaseSense ? strstr(ARG2, ARG3) : stristr(ARG2, ARG3);
			if (found)
				pos = (int)(found - ARG2);
			else
			{
				pos = -1;  // Another indicator in addition to the below.
				g_ErrorLevel->Assign(ERRORLEVEL_ERROR);  // So that it behaves like AutoIt2.
			}
		}
		return OUTPUT_VAR->Assign(pos); // It already displayed any error that may have occurred.
	}
	case ACT_STRINGREPLACE:
	{
		source_length = strlen(ARG2);
		space_needed = (VarSizeType)source_length + 1;  // Set default, or starting value for accumulation.
		VarSizeType final_space_needed = space_needed;
		bool do_replace = *ARG2 && *ARG3; // i.e. don't allow replacement of the empty string.
		bool replace_all = toupper(*ARG5) == 'A';
		UINT found_count = 0; // Set default.

		if (do_replace) 
		{
			// Note: It's okay if Search String is a subset of Replace String.
			// Example: Replacing all occurrences of "a" with "abc" would be
			// safe the way this StrReplaceAll() works (other implementations
			// might cause on infinite loop).
			size_t search_str_len = strlen(ARG3);
			size_t replace_str_len = strlen(ARG4);
			char *found_pos;
			for (found_count = 0, found_pos = ARG2;;)
			{
				if (   !(found_pos = g.StringCaseSense ? strstr(found_pos, ARG3) : stristr(found_pos, ARG3))   )
					break;
				++found_count;
				final_space_needed += (int)(replace_str_len - search_str_len); // Must cast to int.
				// Jump to the end of the string that was just found, in preparation
				// for the next iteration:
				found_pos += search_str_len;
				if (!replace_all) // Replacing only one, so we're done.
					break;
			}
			// Use the greater of the two because temporarily need more space in the output
			// var, because that is where the replacement will be conducted:
			if (final_space_needed > space_needed)
				space_needed = final_space_needed;
		}

		// AutoIt2: "If the search string cannot be found, the contents of <Output Variable>
		// will be the same as <Input Variable>."
		g_ErrorLevel->Assign(found_count ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR); // So that it behaves like AutoIt2.
	
		// Handle the output parameter.  This section is similar to that in PerformAssign().
		// Set up the var, enlarging it if necessary.  If the OUTPUT_VAR is of type VAR_CLIPBOARD,
		// this call will set up the clipboard for writing:
		if (OUTPUT_VAR->Assign(NULL, space_needed - 1) != OK)
			return FAIL;
		// Fetch the text directly into the var:
		if (space_needed == 1)
			*OUTPUT_VAR->Contents() = '\0';
		else
			strlcpy(OUTPUT_VAR->Contents(), ARG2, space_needed);
		OUTPUT_VAR->Length() = final_space_needed - 1;  // This will be the length after replacement is done.

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
				StrReplaceAll(OUTPUT_VAR->Contents(), ARG3, ARG4, g.StringCaseSense);
			else
				StrReplace(OUTPUT_VAR->Contents(), ARG3, ARG4, g.StringCaseSense);

		// UPDATE: This is NOT how AutoIt2 behaves, so don't do it:
		//if (g_script.mIsAutoIt2)
		//{
		//	trim(OUTPUT_VAR->Contents());  // Since this is how AutoIt2 behaves.
		//	OUTPUT_VAR->Length() = (VarSizeType)strlen(OUTPUT_VAR->Contents());
		//}

		// Consider the above to have been always successful unless the below returns an error:
		return OUTPUT_VAR->Close();  // In case it's the clipboard.
	}

	case ACT_GETKEYSTATE:
		if (vk = TextToVK(ARG2))
		{
			switch (toupper(*ARG3))
			{
			case 'T': // Whether a toggleable key such as CapsLock is currently turned on.
				return OUTPUT_VAR->Assign((GetKeyState(vk) & 0x00000001) ? "D" : "U");
			case 'P': // Physical state of key.
				if (g_hhkLowLevelKeybd)
					// Since the hook is installed, use its value rather than that from
					// GetAsyncKeyState(), which doesn't seem to work as advertised, at
					// least under WinXP:
					return OUTPUT_VAR->Assign(g_PhysicalKeyState[vk] ? "D" : "U");
				else
					return OUTPUT_VAR->Assign(IsPhysicallyDown(vk) ? "D" : "U");
			default: // Logical state of key
				return OUTPUT_VAR->Assign((GetKeyState(vk) & 0x8000) ? "D" : "U");
			}
		}
		return OUTPUT_VAR->Assign("");

	case ACT_RANDOM:
	{
		int rand_min = *ARG2 ? atoi(ARG2) : 0;
		int rand_max = *ARG3 ? atoi(ARG3) : INT_MAX;
		// Seems best not to use ErrorLevel for this command at all, since silly cases
		// such as Max > Min are too rare.  Swap the two values instead.
		if (rand_min > rand_max)
		{
			int rand_swap = rand_min;
			rand_min = rand_max;
			rand_max = rand_swap;
		}

		// Adapted from the AutoIt3 source:
#ifdef _MSC_VER
		// AutoIt3: __int64 is needed here to do the proper conversion from unsigned long to signed long
		int our_rand = int(__int64(genrand_int32()%(unsigned long)(rand_max - rand_min + 1)) + rand_min);
#else
		// My: Something seems fishy with this part, so if ever compile this on a non-MSC
		// compiler, might want to review:
		// AutoIt3: What to do when I do not have __int64
		// store in double (15 digits of precision will store 10 digit long)
		// Converting through a double is a lot slow than using __int64
		double fTemp;
		fTemp = fmod(genrand_int32(), rand_max - rand_min + 1) + rand_min;
		int our_rand = int(fTemp);
#endif
		return OUTPUT_VAR->Assign(our_rand); // It already displayed any error that may have occurred.
	}

	case ACT_ASSIGN:
		// Note: This line's args have not yet been dereferenced in this case.  The below
		// function will handle that if it is needed.
		return PerformAssign();  // It will report any errors for us.
	case ACT_DRIVESPACEFREE:
		return DriveSpaceFree(ARG2);
	case ACT_SOUNDSETWAVEVOLUME:
	{
		// Adapted from the AutoIt3 source.
		int volume = atoi(ARG1);
		if (volume < 0 || volume > 100)
		{
			g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
			return OK;  // Let ErrorLevel tell the story.
		}
		WORD wVolume = 0xFFFF * volume / 100;
		if (waveOutSetVolume(0, MAKELONG(wVolume, wVolume)) == MMSYSERR_NOERROR)
			g_ErrorLevel->Assign(ERRORLEVEL_NONE);
		else
			g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		return OK;
	}

	case ACT_FILESELECTFILE:
		return FileSelectFile(ARG2, ARG3);
	case ACT_FILECREATEDIR:
		return FileCreateDir(ARG1);
	case ACT_FILEREMOVEDIR:
		if (!*ARG1) // Consider an attempt to create or remove a blank dir to be an error.
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		return g_ErrorLevel->Assign(RemoveDirectory(ARG1) ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);
	case ACT_FILEREADLINE:
		return FileReadLine(ARG2, ARG3);
	case ACT_FILEAPPEND:
		return this->FileAppend(ARG2, ARG1);  // To avoid ambiguity in case there's another FileAppend().
	case ACT_FILEDELETE:
		return FileDelete(ARG1);
	case ACT_FILEMOVE:
		return FileMove(THREE_ARGS);
	case ACT_FILECOPY:
		return FileCopy(THREE_ARGS);

	// Like AutoIt2, if either OUTPUT_VAR or ARG1 aren't purely numeric, they
	// will be considered to be zero for all of the below math functions:
	case ACT_ADD:
		math_result = PureNumberToInt(OUTPUT_VAR->Contents()) + PureNumberToInt(ARG2);
		sprintf(buf_temp, "%d", math_result);
		OUTPUT_VAR->Assign(buf_temp);
		return OK;
	case ACT_SUB:
		math_result = PureNumberToInt(OUTPUT_VAR->Contents()) - PureNumberToInt(ARG2);
		sprintf(buf_temp, "%d", math_result);
		OUTPUT_VAR->Assign(buf_temp);
		return OK;
	case ACT_MULT:
		math_result = PureNumberToInt(OUTPUT_VAR->Contents()) * PureNumberToInt(ARG2);
		sprintf(buf_temp, "%d", math_result);
		OUTPUT_VAR->Assign(buf_temp);
		return OK;
	case ACT_DIV:
	{
		int value = PureNumberToInt(ARG2);
		if (!value)
			return LineError("This line would attempt to divide by zero (or a value that resolves"
				" to zero because it's non-numeric)." ERR_ABORT, FAIL, ARG2);
		math_result = PureNumberToInt(OUTPUT_VAR->Contents()) / value;
		sprintf(buf_temp, "%d", math_result);
		OUTPUT_VAR->Assign(buf_temp);
		return OK;
	}

	case ACT_FILETOGGLEHIDDEN:
	{
		// If there is a non-empty first param, it takes precedence over
		// a non-NULL aCurrentFile:
		char *filespec = ARG1;
		if (!*filespec && aCurrentFile)
			filespec = aCurrentFile->cFileName;
		if (!*filespec)
			// It's probably a deref'd var, otherwise the loader would have caught it.
			// It's probably best to abort the script/subroutine in case it's a loop
			// with many iterations, each of which would likely fail:
			return LineError("The filename provided is blank." ERR_ABORT);
		DWORD attr = GetFileAttributes(filespec);
		if (attr == 0xFFFFFFFF)  // failed
			return LineError("GetFileAttributes() failed.", WARN, filespec);
		if (attr & FILE_ATTRIBUTE_HIDDEN)
			attr &= ~FILE_ATTRIBUTE_HIDDEN;
		else
			attr |= FILE_ATTRIBUTE_HIDDEN;
		if (!SetFileAttributes(filespec, attr))
			return LineError("SetFileAttributes() failed.", WARN, filespec);
		return OK; // Always successful from a script-execution standpoint.
	}
	case ACT_FILESETDATEMODIFIED:
	{
		// If there is a non-empty first param, it takes precedence over
		// a non-NULL aCurrentFile:
		char *filespec = ARG1;
		if (!*filespec && aCurrentFile)
			filespec = aCurrentFile->cFileName;
		if (!*filespec)
			// It's probably a deref'd var, otherwise the loader would have caught it.
			// It's probably best to abort the script/subroutine in case it's a loop
			// with many iterations, each of which would likely fail:
			return LineError("The filename provided is blank." ERR_ABORT);
		if (!FileSetDateModified(filespec, ARG2))
			return LineError("This file or folder's modification date could not be changed.", WARN, filespec);
		return OK;
	}
	case ACT_KEYLOG:
	{
		if (*ARG1)
		{
			if (!stricmp(ARG1, "Off"))
				g_KeyLogToFile = false;
			else if (!stricmp(ARG1, "On"))
				g_KeyLogToFile = true;
			else if (!stricmp(ARG1, "Toggle"))
				g_KeyLogToFile = !g_KeyLogToFile;
			else // Assume the param is the target file to which to log the keys:
			{
				g_KeyLogToFile = true;
				KeyLogToFile(ARG1);
			}
			return OK;
		}
		// I was initially concerned that GetWindowText() can hang if the target window is
		// hung.  But at least on newer OS's, this doesn't seem to be a problem: MSDN says
		// "If the window does not have a caption, the return value is a null string. This
		// behavior is by design. It allows applications to call GetWindowText without hanging
		// if the process that owns the target window is hung. However, if the target window
		// is hung and it belongs to the calling application, GetWindowText will hang the
		// calling application."
		target_window = GetForegroundWindow();
		char win_title[50];  // Keep it small because MessageBox() will truncate the text if too long.
		if (target_window)
			GetWindowText(target_window, win_title, sizeof(win_title));
		else
			*win_title = '\0';
		char LRtext[128];
		snprintf(buf_temp, sizeof(buf_temp),
			"Window: %s"
			//"\r\nBlocks: %u"
			"\r\nKeybd hook: %s"
			"\r\nMouse hook: %s"
			"\r\nLast hotkey type: %s"
			"\r\nInterrupted subroutines: %d%s"
			"\r\nPaused subroutines: %d"
			"\r\nMsgBoxes: %d"
			"\r\nModifiers (GetKeyState() now) = %s"
			"\r\n"
			, win_title
			//, SimpleHeap::GetBlockCount()
			, g_hhkLowLevelKeybd == NULL ? "no" : "yes"
			, g_hhkLowLevelMouse == NULL ? "no" : "yes"
			, g_LastPerformedHotkeyType == HK_KEYBD_HOOK ? "keybd hook" : "not keybd hook"
			, g_nInterruptedSubroutines
			, g_nInterruptedSubroutines ? " (preempted: they will resume when the current subroutine finishes)" : ""
			, g_nPausedSubroutines
			, g_nMessageBoxes
			, ModifiersLRToText(GetModifierLRStateSimple(), LRtext));
		size_t length = strlen(buf_temp);
		GetHookStatus(buf_temp + length, sizeof(buf_temp) - length);
		ShowMainWindow(buf_temp);
		// It seems okay to allow more than one of these windows to be on the screen at a time.
		// That way, one window's contents can be compared with another's:
		//if (!MsgBox(buf_temp))
		//	return OK;  // It's usually safe to proceed with the script even if it failed.
		return OK;
	}
	case ACT_LISTLINES:
		ShowMainWindow(NULL, true); // Given a NULL, it defaults to showing the lines for us.
		// It seems okay to allow more than one of these windows to be on the screen at a time.
		// That way, one window's contents can be compared with another's:
		//if (!MsgBox(buf_temp))
		//	return OK;  // It's usually safe to proceed with the script even if it failed.
		return OK;
	case ACT_LISTVARS:
		g_script.ListVars(buf_temp, sizeof(buf_temp));
		ShowMainWindow(buf_temp);
		// It seems okay to allow more than one of these windows to be on the screen at a time.
		// That way, one window's contents can be compared with another's:
		//if (!MsgBox(buf_temp, MB_OK, "Variables (in order of appearance) & their current contents"))
		//	return OK;  // It's usually safe to proceed with the script even if it failed.
		return OK;
	case ACT_LISTHOTKEYS:
		Hotkey::ListHotkeys(buf_temp, sizeof(buf_temp));
		ShowMainWindow(buf_temp);
		// It seems okay to allow more than one of these windows to be on the screen at a time.
		// That way, one window's contents can be compared with another's:
		//if (!MsgBox(buf_temp, MB_OK, "   Type        Name"))
		//	return OK;  // It's usually safe to proceed with the script even if it failed.
		return OK;
	case ACT_MSGBOX:
	{
		int result;
		// If the MsgBox window can't be displayed for any reason, always return FAIL to
		// the caller because it would be unsafe to proceed with the execution of the
		// current script subroutine.  For example, if the script contains an IfMsgBox after,
		// this line, it's result would be unpredictable and might cause the subroutine to perform
		// the opposite action from what was intended (e.g. Delete vs. don't delete a file).
		result = (mArgc == 1) ? MsgBox(ARG1) : MsgBox(ARG3, atoi(ARG1), ARG2, atoi(ARG4));
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
		return InputBox(OUTPUT_VAR, ARG2, ARG3, toupper(*ARG4) == 'H'); // Last = whether to hide input.
	case ACT_SPLASHTEXTON:
	{
		// The SplashText command has been adapted from the AutoIt3 source.
		int W = *ARG1 ? atoi(ARG1) : 200;  // AutoIt3 default is 500.
		int H = *ARG2 ? atoi(ARG2) : 0;
		// Above: AutoIt3 default is 400 but I like 0 better because the default might
		// be accidentally triggered sometimes if the params are used wrong, or are blank
		// derefs.  In such cases, 400 would create a large topmost, unmovable, unclosable
		// window that would interfere with user's ability to see dialogs and other things.

		// Add some caption and frame size to window:
		W += GetSystemMetrics(SM_CXEDGE) * 2;
		int min_height = GetSystemMetrics(SM_CYCAPTION) + (GetSystemMetrics(SM_CYEDGE) * 2);
		if (g_script.mIsAutoIt2)
		{
			// I think this is probably something like how AutoIt2 does things:
			if (H < min_height)
				H = min_height;
		}
		else // Use the new, AutoIt3 method.
			H += min_height;

		RECT rect;
		SystemParametersInfo(SPI_GETWORKAREA, 0, &rect, 0);	// Get Desktop rect
		int xpos = (rect.right - W)/2;  // Center splash horizontally
		int ypos = (rect.bottom - H)/2; // Center splash vertically

		// My: Probably not to much overhead to do this, though it probably would
		// perform better to resize and "re-text" the existing window rather than
		// recreating it:
		#define DESTROY_SPLASH if (g_hWndSplash)\
		{\
			DestroyWindow(g_hWndSplash);\
			g_hWndSplash = NULL;\
		}
		DESTROY_SPLASH

		// Doesn't seem necessary to have it owned by the main window, but neither
		// does doing so seem to cause any harm.  Feels safer to have it be
		// an independent window.  Update: Must make it owned by the parent window
		// otherwise it will get its own tray icon, which is usually undesirable:
		g_hWndSplash = CreateWindowEx(WS_EX_TOPMOST, WINDOW_CLASS_NAME, ARG3  // ARG3 is the window title
			, WS_DISABLED|WS_POPUP|WS_CAPTION, xpos, ypos, W, H
			, g_hWnd, (HMENU)NULL, g_hInstance, NULL);

		GetClientRect(g_hWndSplash, &rect);	// get the client size

		// CREATE static label full size of client area
		HWND static_win = CreateWindowEx(0, "static", ARG4 // ARG4 is the window's text
			, WS_CHILD|WS_VISIBLE|SS_CENTER
			, 0, 0, rect.right - rect.left, rect.bottom - rect.top
			, g_hWndSplash, (HMENU)NULL, g_hInstance, NULL);

		char szFont[65];
		int CyPixels, nSize = 12, nWeight = 400;
		HFONT hfFont;
		HDC h_dc = CreateDC("DISPLAY", NULL, NULL, NULL);
		SelectObject(h_dc, (HFONT)GetStockObject(DEFAULT_GUI_FONT));		// Get Default Font Name
		GetTextFace(h_dc, sizeof(szFont) - 1, szFont); // -1 just in case, like AutoIt3.
		CyPixels = GetDeviceCaps(h_dc, LOGPIXELSY);			// For Some Font Size Math
		DeleteDC(h_dc);
		//strcpy(szFont,vParams[7].szValue());	// Font Name
		//nSize = vParams[8].nValue();		// Font Size
		//if ( vParams[9].nValue() >= 0 && vParams[9].nValue() <= 1000 )
		//	nWeight = vParams[9].nValue();			// Font Weight
		hfFont = CreateFont(0-(nSize*CyPixels)/72,0,0,0,nWeight,0,0,0,DEFAULT_CHARSET,
			OUT_TT_PRECIS,CLIP_DEFAULT_PRECIS,PROOF_QUALITY,FF_DONTCARE,szFont);	// Create Font

		SendMessage(static_win, WM_SETFONT, (WPARAM)hfFont, MAKELPARAM(TRUE, 0));	// Do Font
		ShowWindow(g_hWndSplash, SW_SHOWNOACTIVATE);				// Show the Splash
		return OK;
	}
	case ACT_SPLASHTEXTOFF:
		DESTROY_SPLASH
		return OK;
	case ACT_SEND:
		SendKeys(ARG1, aModifiersLR);
		return OK;
	case ACT_MOUSECLICKDRAG:
		if (   !(vk = ConvertMouseButton(ARG1))   )
			return LineError(ERR_MOUSE_BUTTON ERR_ABORT, FAIL, ARG1);
		if (!ValidateMouseCoords(ARG2, ARG3))
			return LineError(ERR_MOUSE_COORD ERR_ABORT, FAIL, ARG2);
		if (!ValidateMouseCoords(ARG4, ARG5))
			return LineError(ERR_MOUSE_COORD ERR_ABORT, FAIL, ARG4);
		// If no starting coords are specified, we tell the function to start at the
		// current mouse position:
		x = *ARG2 ? atoi(ARG2) : COORD_UNSPECIFIED;
		y = *ARG3 ? atoi(ARG3) : COORD_UNSPECIFIED;
		MouseClickDrag(vk, x, y, atoi(ARG4), atoi(ARG5), *ARG6 ? atoi(ARG6) : g.DefaultMouseSpeed);
		return OK;
	case ACT_MOUSECLICK:
		if (   !(vk = ConvertMouseButton(ARG1))   )
			return LineError(ERR_MOUSE_BUTTON ERR_ABORT, FAIL, ARG1);
		if (!ValidateMouseCoords(ARG2, ARG3))
			return LineError(ERR_MOUSE_COORD ERR_ABORT, FAIL, ARG2);
		x = *ARG2 ? atoi(ARG2) : COORD_UNSPECIFIED;
		y = *ARG3 ? atoi(ARG3) : COORD_UNSPECIFIED;
		MouseClick(vk, x, y, *ARG4 ? atoi(ARG4) : 1, *ARG5 ? atoi(ARG5) : g.DefaultMouseSpeed, *ARG6);
		return OK;
	case ACT_MOUSEMOVE:
		if (!ValidateMouseCoords(ARG1, ARG2))
			return LineError(ERR_MOUSE_COORD ERR_ABORT, FAIL, ARG1);
		x = *ARG1 ? atoi(ARG1) : COORD_UNSPECIFIED;
		y = *ARG2 ? atoi(ARG2) : COORD_UNSPECIFIED;
		MouseMove(x, y, *ARG3 ? atoi(ARG3) : g.DefaultMouseSpeed);
		return OK;
	case ACT_MOUSEGETPOS:
		return MouseGetPos();
//////////////////////////////////////////////////////////////////////////
	case ACT_SETDEFAULTMOUSESPEED:
		g.DefaultMouseSpeed = atoi(ARG1);
		// In case it was a deref, force it to be some default value if it's out of range:
		if (g.DefaultMouseSpeed < 0 || g.DefaultMouseSpeed > MAX_MOUSE_SPEED)
			g.DefaultMouseSpeed = DEFAULT_MOUSE_SPEED;
		return OK;
	case ACT_SETTITLEMATCHMODE:
		switch (ConvertTitleMatchMode(ARG1))
		{
		case FIND_IN_LEADING_PART:
			g.TitleFindAnywhere = false;
			return OK;
		case FIND_ANYWHERE:
			g.TitleFindAnywhere = true;
			return OK;
		case FIND_FAST:
			g.TitleFindFast = true;
			return OK;
		case FIND_SLOW:
			g.TitleFindFast = false;
			return OK;
		}
		return LineError(ERR_TITLEMATCHMODE2, FAIL, ARG1);
	case ACT_SETWINDELAY: g.WinDelay = atoi(ARG1); return OK;
	case ACT_SETKEYDELAY: g.KeyDelay = atoi(ARG1); return OK;
	case ACT_SETBATCHLINES:
		if (   !(g.LinesPerCycle = atoi(ARG1))   )
			// Don't interpret zero as "infinite" because zero can accidentally
			// occur if the dereferenced var was blank:
			g.LinesPerCycle = DEFAULT_BATCH_LINES;
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

	case ACT_FORCE_KEYBD_HOOK:
		// Anything other than "On" causes "Off" in this case:
		g_ForceKeybdHook = (ConvertOnOff(ARG1) == TOGGLED_ON) ? TOGGLED_ON : TOGGLED_OFF;
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
	case ACT_RELOADCONFIG:
		g_script.Reload();
		return OK;

	case ACT_INVALID: // Should be impossible because Script::AddLine() forbids it.
		return LineError("Perform(): Invalid action type." PLEASE_REPORT ERR_ABORT);
	}

	// Since above didn't return, this line's mActionType isn't handled here,
	// so caller called it wrong:
	return LineError("Perform(): Unhandled action type." PLEASE_REPORT ERR_ABORT);
}



ResultType Line::ExpandArgs()
{
	// Make two passes through this line's arg list.  This is done because the performance of
	// realloc() is worse than doing a free() and malloc() because the former does a memcpy()
	// in addition to the latter's steps.  In addition, realloc() as much as doubles the memory
	// load on the system during the brief time that both the old and the new blocks of memory exist.
	// First pass: determine how must space will be needed to do all the args and allocate
	// more memory if needed.  Second pass: dereference the args into the buffer.

	size_t space_needed = GetExpandedArgSize(true);  // First pass.

	if (space_needed > DEREF_BUF_MAX)
		return LineError("Dereferencing the variables in this line's parameters"
			" would exceed the allowed size of the temp buffer." ERR_ABORT);

	// Only allocate the buf at the last possible moment,
	// when it's sure the buffer will be used (improves performance when only a short
	// script with no derefs is being run):
	if (space_needed > sDerefBufSize)
	{
		// K&R: Integer division truncates the fractional part:
		size_t increments_needed = space_needed / DEREF_BUF_EXPAND_INCREMENT;
		if (space_needed % DEREF_BUF_EXPAND_INCREMENT)  // Need one more if above division did truncate it.
			++increments_needed;
		sDerefBufSize = increments_needed * DEREF_BUF_EXPAND_INCREMENT;
		if (sDerefBuf != NULL)
			// Do a free() and malloc(), which should be far more efficient than realloc(),
			// especially if there is a large amount of memory involved here:
			free(sDerefBuf);
		if (!(sDerefBufMarker = sDerefBuf = (char *)malloc(sDerefBufSize)))
		{
			sDerefBufSize = 0;  // Reset so that it can attempt a smaller amount next time.
			return LineError("Ran out of memory while attempting to dereference this line's parameters."
				ERR_ABORT);
		}
	}
	else
		sDerefBufMarker = sDerefBuf;  // Prior contents of buffer will be overwritten in any case.
	// Always init sDerefBufMarker even if zero iterations, because we want to enforce
	// the fact that its prior contents become invalid once this function is called.
	// It's also necessary due to the fact that all the old memory is discarded by
	// the above if more space was needed to accommodate this line:

	Var *the_only_var_of_this_arg;
	for (int iArg = 0; iArg < mArgc && iArg < MAX_ARGS; ++iArg) // Second pass.
	{
		// FOR EACH ARG:
		if (ARG_IS_OUTPUT_VAR(mArg[iArg]))  // Don't bother wasting the mem to deref output var.
		{
			// In case its "dereferenced" contents are ever directly examined, set it to be
			// the empty string.  This also allows the ARG to be passed a dummy param, which
			// makes things more convenient and maintainable in other places:
			sArgDeref[iArg] = "";
			continue;
		}
		if (   !(the_only_var_of_this_arg = ARG_IS_INPUT_VAR(mArg[iArg]))   ) // Arg isn't an input var.
		{
			#define NO_DEREF (!ArgHasDeref(iArg + 1))
			if (NO_DEREF)
			{
				sArgDeref[iArg] = mArg[iArg].text;  // Point the dereferenced arg to the arg text itself.
				continue;  // Don't need to use the buffer in this case.
			}
			// Now we know it has at least one deref.  If the second defer's marker is NULL,
			// the first is the only deref in this arg:
			#define SINGLE_ISOLATED_DEREF (!mArg[iArg].deref[1].marker\
				&& mArg[iArg].deref[0].length == strlen(mArg[iArg].text)) // and the arg contains no literal text
			if (SINGLE_ISOLATED_DEREF)
				the_only_var_of_this_arg = mArg[iArg].deref[0].var;
		}

		if (the_only_var_of_this_arg) // This arg resolves to only a single, naked var.
		{
			if (ArgMustBeDereferenced(the_only_var_of_this_arg))
			{
				// the_only_var_of_this_arg is either a reserved var or a normal var of
				// zero length (for which GetEnvironment() is called for), or is used
				// again in this line as an output variable.  In all these cases, it must
				// be expanded into the buffer rather than accessed directly:
				sArgDeref[iArg] = sDerefBufMarker; // Point it to its location in the buffer.
				sDerefBufMarker += the_only_var_of_this_arg->Get(sDerefBufMarker) + 1; // +1 for terminator.
			}
			else
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
				// to write the filespecs:
				sArgDeref[iArg] = the_only_var_of_this_arg->Contents();
		}
		else
		{
			sArgDeref[iArg] = sDerefBufMarker; // Point it to its location in the buffer.
			sDerefBufMarker = ExpandArg(sDerefBufMarker, iArg); // Expand the arg into that location.
		}
	}

	// Validate numeric params after runtime deref.
	// This section is similar to the one used to valididate the non-dereferenced numeric params
    // during load:
	if (g_act[mActionType].NumericParams)
	{
		bool allow_negative;
		for (ActionTypeType *np = g_act[mActionType].NumericParams; *np; ++np)
		{
			if (mArgc >= *np) // The arg exists.
			{
				allow_negative = ArgAllowsNegative(*np);
				if (!IsPureNumeric(sArgDeref[*np - 1], allow_negative))
				{
					if (mActionType == ACT_WINMOVE)
					{
						if (stricmp(sArgDeref[*np - 1], "default"))
							return LineError("This parameter of this line doesn't resolve to either a"
								" numeric value or the word Default as required.", FAIL, sArgDeref[*np - 1]);
						// else don't attempt to set the deref to be blank, to make parsing simpler,
						// because sArgDeref[*np - 1] might point directly to the contents of
						// a variable and we don't want to modify it in that case.
					}
					else
					{
						if (allow_negative)
							return LineError("This parameter of this line doesn't resolve to a"
								" numeric value as required.", FAIL, sArgDeref[*np - 1]);
						else
							return LineError("This parameter of this line doesn't resolve to a"
								" non-negative numeric value as required.", FAIL, sArgDeref[*np - 1]);
					}
				}
			}
		}
	}
	return OK;
}

	

inline VarSizeType Line::GetExpandedArgSize(bool aCalcDerefBufSize)
{
	int iArg;
	VarSizeType space_needed;
	DerefType *deref;
	Var *the_only_var_of_this_arg;

	// Note: the below loop is similar to the one in ExpandArgs(), so the two should be
	// maintained together:
	for (iArg = 0, space_needed = 0; iArg < mArgc && iArg < MAX_ARGS; ++iArg)
	{
		// FOR EACH ARG:
		// Accumulate the total of how much space we will need.
		if (ARG_IS_OUTPUT_VAR(mArg[iArg]))  // These should never be included in the space calculation.
			continue;
		// Always do this check before attempting to traverse the list of dereferences, since
		// such an attempt would be invalid in this case:
		if (   !(the_only_var_of_this_arg = ARG_IS_INPUT_VAR(mArg[iArg]))   ) // It's not an input var.
		{
			if (NO_DEREF)
			{
                if (!aCalcDerefBufSize) // i.e. we want the total size of what the args resolve to.
					space_needed += (VarSizeType)strlen(mArg[iArg].text) + 1;  // +1 for the zero terminator.
				// else don't increase space_needed, even by 1 for the zero terminator, because
				// the terminator isn't needed if the arg won't exist in the buffer at all.
				continue;
			}
			if (SINGLE_ISOLATED_DEREF)
				the_only_var_of_this_arg = mArg[iArg].deref[0].var;
		}
		if (the_only_var_of_this_arg)
		{
			if (!aCalcDerefBufSize || ArgMustBeDereferenced(the_only_var_of_this_arg))
				// Either caller wanted it included anyway, or it must always been deref'd:
				space_needed += the_only_var_of_this_arg->Get() + 1;  // +1 for the zero terminator.
			// else no extra space is needed in the buffer.
			continue;
		}
		// Otherwise: This arg has more than one deref, or a single deref with some literal text around it.
		space_needed += (VarSizeType)strlen(mArg[iArg].text);
		for (deref = mArg[iArg].deref; deref && deref->marker; ++deref)
		{
			space_needed -= deref->length;     // Substitute the length of the variable contents
			space_needed += deref->var->Get(); // for that of the literal text of the deref.
		}
		// +1 for this arg's zero terminator in the buffer:
		++space_needed;
	}
	return space_needed;
}



inline char *Line::ExpandArg(char *aBuf, int aArgIndex)
// Returns a pointer to the char in aBuf that occurs after the zero terminator
// (because that's the position where the caller would normally resume writing
// if there are more args, since the zero terminator must normally be retained
// between args).
// Caller must ensure that aBuf is large enough to accommodate the translation
// of the Arg.  No validation of above params is done, caller must do that.
{
	// This should never be called if the given arg is an output var, but we check just in case:
	if (ARG_IS_OUTPUT_VAR(mArg[aArgIndex]))
		// Just a warning because this function isn't set up to cause a true failure:
		LineError("ExpandArg() was called to expand an arg that contains only an output variable."
			PLEASE_REPORT, WARN);
	// Always do this part before attempting to traverse the list of dereferences, since
	// such an attempt would be invalid in this case:
	Var *the_only_var_of_this_arg = ARG_IS_VAR(mArg[aArgIndex]);
	if (the_only_var_of_this_arg)
		// +1 so that we return the position after the terminator, as required:
		return aBuf += the_only_var_of_this_arg->Get(aBuf) + 1;

	char *pText;
	DerefType *deref;
	for (pText = mArg[aArgIndex].text  // Start at the begining of this arg's text.
		, deref = mArg[aArgIndex].deref  // Start off by looking for the first deref.
		; deref && deref->marker; ++deref)  // A deref with a NULL marker terminates the list.
	{
		// FOR EACH DEREF IN AN ARG (if we're here, there's at least one):
		// Copy the chars that occur prior to deref->marker into the buffer:
		for (; pText < deref->marker; *aBuf++ = *pText++);

		// Now copy the contents of the dereferenced var.  For all cases, aBuf has already
		// been verified to be large enough, assuming the value hasn't changed between the
		// time we were called and the time the caller calculated the space needed.
		aBuf += deref->var->Get(aBuf);
		// Finally, jump over the dereference text:
		pText += deref->length;
	}
	// Copy any chars that occur after the final deref into the buffer:
	for (; *pText; *aBuf++ = *pText++);
	// Terminate the buffer, even if nothing was written into it:
	*aBuf++ = '\0';
	return aBuf; // Returns the position after the terminator.
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
	snprintf(aBuf, BUF_SPACE_REMAINING, "Script lines most recently executed (oldest first).  Press [F5] to refresh.\r\n\r\n");
	aBuf += strlen(aBuf);
	// Start at the oldest and continue up through the newest:
	int i, line_index;
	for (i = 0, line_index = sLogNext; i < LINE_LOG_SIZE; ++i, ++line_index)
	{
		if (line_index >= LINE_LOG_SIZE) // wrap around
			line_index = 0;
		if (sLog[line_index] == NULL)
			continue;
		aBuf = sLog[line_index]->ToText(aBuf, BUF_SPACE_REMAINING, true);
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
	// Start at the oldest and continue up through the newest:
	for (Line *line = line_start;;)
	{
		if (line == this)
			strlcpy(aBuf, "--->\t", BUF_SPACE_REMAINING);
		else
			strlcpy(aBuf, "\t", BUF_SPACE_REMAINING);
		aBuf += strlen(aBuf);
		aBuf = line->ToText(aBuf, BUF_SPACE_REMAINING, true);
		if (line == line_end)
			break;
		line = line->mNextLine;
	}
	return aBuf;
}



char *Line::ToText(char *aBuf, size_t aBufSize, bool aAppendNewline)
// Translates this line into its text equivalent, putting the result into aBuf and
// returning the position in aBuf of its new string terminator.
{
	if (!aBuf) return NULL;
	char *aBuf_orig = aBuf;
	snprintf(aBuf, BUF_SPACE_REMAINING, "%03u: ", mFileLineNumber);
	aBuf += strlen(aBuf);
	if (ACT_IS_ASSIGN(mActionType) || (ACT_IS_IF(mActionType) && mActionType < ACT_FIRST_COMMAND))
	{
		// Only these commands need custom conversion.
		snprintf(aBuf, BUF_SPACE_REMAINING, "%s%s %s %s"
			, ACT_IS_IF(mActionType) ? "IF " : "", VARARG1->mName, g_act[mActionType].Name, RAW_ARG2);
		aBuf += strlen(aBuf);
	}
	else
	{
		snprintf(aBuf, BUF_SPACE_REMAINING, "%s", g_act[mActionType].Name);
		aBuf += strlen(aBuf);
		for (int i = 0; i < mArgc; ++i)
		{
			// This method a little more efficient than using snprintfcat:
			snprintf(aBuf, BUF_SPACE_REMAINING, ",%s", ARG_IS_VAR(mArg[i]) ? VAR(mArg[i])->mName : mArg[i].text);
			aBuf += strlen(aBuf);
		}
	}
	if (aAppendNewline && BUF_SPACE_REMAINING >= 2)
	{
		*aBuf++ = '\r';
		*aBuf++ = '\n';
		*aBuf = '\0';
	}
	return aBuf;
}



void Line::ToggleSuspendState()
{
	if (g_IsSuspended) // toggle off the suspension
		Hotkey::AllActivate();
	else // suspend
		Hotkey::AllDeactivate(true);
	g_IsSuspended = !g_IsSuspended;
	g_script.UpdateTrayIcon();
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
		// active to call us):
		g.IsPaused = true;
		++g_nPausedSubroutines;
		g_script.UpdateTrayIcon();
		return OK;
	case TOGGLED_OFF:
		// Unpause the uppermost underlying paused subroutine.  If none of the interrupted subroutines
		// are paused, do nothing. This relies on the fact that the current script subroutine,
		// which got us here, cannot itself be currently paused by definition:
		if (g_nPausedSubroutines > 0)
			g_UnpauseWhenResumed = true;
		return OK;
	case NEUTRAL: // the user omitted the parameter entirely, which is considered the same as "toggle"
	case TOGGLE:
		// See TOGGLED_OFF comment above:
		if (g_nPausedSubroutines > 0)
			g_UnpauseWhenResumed = true;
		else
		{
			// Since there are no underlying paused subroutines, pause the current subroutine.
			// There must be a current subroutine (i.e. the script can't be idle) because the
			// it's the one that called us directly, and we know it's not paused:
			g.IsPaused = true;
			++g_nPausedSubroutines;
			g_script.UpdateTrayIcon();
		}
		return OK;
	default: // TOGGLE_INVALID or some other disallowed value.
		// We know it's a variable because otherwise the loading validation would have caught it earlier:
		return LineError("The variable in param #1 does not resolve to an allowed value.", FAIL, ARG1);
	}
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
	if (aErrorText == NULL)
		aErrorText = "Unknown Error";
	if (aExtraInfo == NULL)
		aExtraInfo = "";
	char buf[MSGBOX_TEXT_SIZE];
	snprintf(buf, sizeof(buf), "%s: %-1.500s\n\n"  // Keep it to a sane size in case it's huge.
		, aErrorType == WARN ? "Warning" : (aErrorType == CRITICAL_ERROR ? "Critical Error" : "Error")
		, aErrorText);
	if (*aExtraInfo)
		// Use format specifier to make sure really huge strings that get passed our
		// way, such as a var containing clipboard text, are kept to a reasonable size:
		snprintfcat(buf, sizeof(buf), "Info: %-1.100s%s\n\n"
		, aExtraInfo, strlen(aExtraInfo) > 100 ? "..." : "");
	char *buf_marker = buf + strlen(buf);
	buf_marker = VicinityToText(buf_marker, sizeof(buf) - (buf_marker - buf));
	if (aErrorType == CRITICAL_ERROR || (aErrorType == FAIL && !g_script.mIsReadyToExecute))
		strlcpy(buf_marker, g_script.mIsRestart ? ("\n" OLD_STILL_IN_EFFECT) : ("\n" WILL_EXIT)
			, sizeof(buf) - (buf_marker - buf));
	//buf_marker += strlen(buf_marker);
	g_script.mCurrLine = this;  // This needs to be set in some cases where the caller didn't.
	g_script.ShowInEditor();
	MsgBox(buf);
	if (aErrorType == CRITICAL_ERROR && g_script.mIsReadyToExecute)
		// Also ask the main message loop function to quit and announce to the system that
		// we expect it to quit.  In most cases, this is unnecessary because all functions
		// called to get to this point will take note of the CRITICAL_ERROR and thus keep
		// return immediately, all the way back to main.  However, there may cases
		// when this isn't true:
		// Note: Must do this only after MsgBox, since it appears that new dialogs can't
		// be created once it's done:
		PostQuitMessage(CRITICAL_ERROR);
	return aErrorType; // The caller told us whether it should be a critical error or not.
}



ResultType Script::ScriptError(char *aErrorText, char *aExtraInfo)
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
	if (aErrorText == NULL)
		aErrorText = "Unknown Error";
	if (aExtraInfo == NULL) // In case the caller explicitly called it with NULL.
		aExtraInfo = "";
	char buf[MSGBOX_TEXT_SIZE];
	snprintf(buf, sizeof(buf), "Error at line %u%s." // Don't call it "critical" because it's usually a syntax error.
		"\n\nLine Text: %-1.100s%s"
		"\nError: %-1.500s"
		"\n\n%s"
		, mFileLineCount, mFileLineCount ? "" : " (unknown)"
		, aExtraInfo // aExtraInfo defaults to "" so this is safe.
		, strlen(aExtraInfo) > 100 ? "..." : ""
		, aErrorText
		, mIsRestart ? OLD_STILL_IN_EFFECT : WILL_EXIT
		);
	ShowInEditor();
	MsgBox(buf);
	return FAIL; // See above for why it's better to return FAIL than CRITICAL_ERROR.
}



void Script::ShowInEditor()
{
	// Disabled for now:
	return;

	bool old_mode = g.TitleFindAnywhere;
	g.TitleFindAnywhere = true;
	HWND editor = WinExist(mFileName);
	g.TitleFindAnywhere = old_mode;
	if (!editor)
		return;
	char buf[256];
	GetWindowText(editor, buf, sizeof(buf));
	if (!stristr(buf, "metapad") && !stristr(buf, "notepad"))
		return;
	SetForegroundWindowEx(editor);  // This does a little bit of waiting for the window to become active.
	MsgSleep(100);  // Give it some extra time in case window animation is turned on.
	if (editor != GetForegroundWindow())
		return;
	strlcpy(buf, "^g", sizeof(buf));  // SendKeys requires it to be modifiable.
	SendKeys(buf);
	HWND goto_window;
	int i;
	for (i = 0; i < 25; ++i)
	{
		if (goto_window = WinActive("Go", "&Line")) // Works with both metapad and notepad.
			break;
		MsgSleep(20);
	}
	if (!goto_window)
		return;
	snprintf(buf, sizeof(buf), "%d{ENTER}"
		, mCurrLine ? mCurrLine->mFileLineNumber : mFileLineCount);
	SendKeys(buf);
	for (i = 0; i < 25; ++i)
	{
		MsgSleep(20); // It seems to need a certain minimum amount of time, so Sleep() before checking.
		if (editor == GetForegroundWindow())
			break;
	}
	if (editor != GetForegroundWindow())
		return;
	strlcpy(buf, "{home}+{end}", sizeof(buf));  // SendKeys requires it to be modifiable.
	SendKeys(buf);
}



char *Script::ListVars(char *aBuf, size_t aBufSize)
// Translates this script's list of variables into text equivalent, putting the result
// into aBuf and returning the position in aBuf of its new string terminator.
{
	if (!aBuf || aBufSize < 256) return NULL;
	char *aBuf_orig = aBuf;
	snprintf(aBuf, BUF_SPACE_REMAINING, "Variables (in order of appearance) & their current contents:\r\n\r\n");
	aBuf += strlen(aBuf);
	// Start at the oldest and continue up through the newest:
	for (Var *var = mFirstVar; var; var = var->mNextVar)
		if (var->mType == VAR_NORMAL) // Don't bother showing clipboard and other built-in vars.
			aBuf = var->ToText(aBuf, BUF_SPACE_REMAINING, true);
	return aBuf;
}



ResultType Script::ActionExec(char *aAction, char *aParams, char *aWorkingDir, bool aDisplayErrors
	, char *aRunShowMode, HANDLE *aProcess)
// Remember that aAction and aParams can both be NULL, so don't dereference without checking first.
// Note: Action & Params are parsed at runtime, here, rather than at load-time because the
// Run or RunWait command might contain a deferenced variable, and such can only be resolved
// at runtime.
{
	if (aProcess) // Init output param if the caller gave us memory to store it.
		*aProcess = NULL;

	// Launching nothing is always a success:
	if (!aAction || !*aAction) return OK;

	if (strlen(aAction) >= LINE_SIZE) // This can happen if user runs the contents of a very large variable.
	{
        if (aDisplayErrors)
			ScriptError("The string to be run is too long.");
		return FAIL;
	}

	// Make sure this is set to NULL because CreateProcess() won't work if it's the empty string:
	if (aWorkingDir && !*aWorkingDir)
		aWorkingDir = NULL;

	// Save the originals for later use:
	char *aAction_orig = aAction;
	char *aParams_orig = aParams;

	#define IS_VERB(str) (   !stricmp(str, "find") || !stricmp(str, "explore") || !stricmp(str, "open")\
		|| !stricmp(str, "edit") || !stricmp(str, "print") || !stricmp(str, "properties")   )
	bool action_is_system_verb = false;

	// Declare these here to ensure they're in scope for the entire function, since their
	// contents may be referred to indirectly:
	char action[LINE_SIZE], *first_phrase, *first_phrase_end, *second_phrase;

	// CreateProcess() requires that it be modifiable, so ensure that it is just in case
	// this function is ever changed to use CreateProcess() even when the caller gave
	// use some params:
	strlcpy(action, aAction, sizeof(action));
	aAction = action;

	if (aParams) // Caller specified the params (even an empty string counts, for this purpose).
		action_is_system_verb = IS_VERB(aAction);
	else // Caller wants us to try to parse params out of aAction.
	{
		aParams = "";  // Set default to be empty string in case the below doesn't find any params.

		// Find out the "first phrase" in the string.  This is done to support the special "find" and "explore"
		// operations as well as minmize the chance that executable names intended by the user to be parameters
		// will not be considered to be the program to run (e.g. for use with a compiler, perhaps).
		if (*aAction == '\"')
		{
			first_phrase = aAction + 1;  // Omit the double-quotes, for use with ProcessCreate() and such.
			first_phrase_end = strchr(first_phrase, '\"');
		}
		else
		{
			first_phrase = aAction;
			// Set first_phrase_end to be the location of the first whitespace char, if
			// one exists:
			first_phrase_end = StrChrAny(first_phrase, " \t"); // Find space or tab.
		}
		// Now first_phrase_end is either NULL, the position of the last double-quote in first-phrase,
		// or the position of the first whitespace char to the right of first_phrase.
		if (first_phrase_end)
		{
			// Split into two phrases for use with AddLine():
			*first_phrase_end = '\0';
			second_phrase = first_phrase_end + 1;
		}
		else // the entire string is considered to be the first_phrase, and there's no second:
			second_phrase = NULL;
		if (action_is_system_verb = IS_VERB(first_phrase))
		{
			aAction = first_phrase;
			aParams = second_phrase ? second_phrase : "";
		}
		else
		{
	// Rather than just consider the first phrase to be the exectable and the rest to be the param, we check it
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
			strlcpy(action, aAction_orig, sizeof(action));  // Restore the original value in case it was changed.
			aAction = action;  // Reset it to its original value, in case it was changed; and ensure it's modifiable.
			aParams = ""; // Init.
			char *action_extension;
			if (   !(action_extension = stristr(aAction, ".exe "))   )
				if (   !(action_extension = stristr(aAction, ".exe\""))   )
					if (   !(action_extension = stristr(aAction, ".bat "))   )
						if (   !(action_extension = stristr(aAction, ".bat\""))   )
							if (   !(action_extension = stristr(aAction, ".com "))   )
								if (   !(action_extension = stristr(aAction, ".com\""))   )
									// Not 100% sure that .cmd and .hta are genuine executables in every sense:
									if (   !(action_extension = stristr(aAction, ".cmd "))   )
										if (   !(action_extension = stristr(aAction, ".cmd\""))   )
											if (   !(action_extension = stristr(aAction, ".hta "))   )
												action_extension = stristr(aAction, ".hta\"");

			if (action_extension)
			{
				// If above isn't true, there's no extension: so assume the whole <aAction> is a document name
				// to be opened by the shell.  Otherwise we do this:
				// +4 for the 3-char extension with the period:
				char *exec_params = action_extension + 4;  // exec_params is now the start of params, or empty-string.
				if (*exec_params == '\"')
					// Exclude from exec_params since it's probably belongs to the action, not the params
					// (i.e. it's paired with another double-quote at the start):
					++exec_params;
				if (*exec_params) // otherwise, there doesn't appear to be any params.
				{
					// Terminate the <aAction> string in the right place.  For this to work correctly,
					// at least one space must exist between action & params (shortcoming?):
					*exec_params = '\0';
					++exec_params;
					ltrim(exec_params);
					aParams = exec_params;
				}
			}
		}
	}

	SHELLEXECUTEINFO sei;
	ZeroMemory(&sei, sizeof(SHELLEXECUTEINFO));
	sei.cbSize = sizeof(sei);
	// Below: "indicate that the hProcess member receives the process handle" and not to display error dialog:
	sei.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_FLAG_NO_UI;
	sei.lpDirectory = aWorkingDir; // OK if NULL or blank; that will cause current dir to be used.
	sei.nShow = (aRunShowMode && *aRunShowMode) ? Line::ConvertRunMode(aRunShowMode) : SW_SHOWNORMAL;
	// Check for special actions that are known to be system verbs:
	if (action_is_system_verb)
	{
		sei.lpVerb = aAction;
		if (!stricmp(aAction, "properties"))
			sei.fMask |= SEE_MASK_INVOKEIDLIST;  // Need to use this for the "properties" verb to work reliably.
		sei.lpFile = aParams;
		sei.lpParameters = NULL;
	}
	else
	{
		sei.lpVerb = "open";
		sei.lpFile = aAction;
		sei.lpParameters = aParams;
	}

	if (!ShellExecuteEx(&sei) || LOBYTE(LOWORD(sei.hInstApp)) <= 32) // Relies on short-circuit boolean order.
	{
		bool success = false;  // Init.
		// Fall back to the AutoIt method of using CreateProcess(), but only if caller didn't
		// originally give us some params (since in that case, ShellExecuteEx() should have worked
		// -- not to mention that CreateProcess() doesn't handle params as a separate string):
		if (!aParams_orig || !*aParams_orig)
		{
			sei.lpVerb = "";  // Set it for use in error reporting.
			aAction = action; // same
			aParams = "";     // same
			STARTUPINFO si = {0};  // Zero fill to be safer.
			si.cb = sizeof(si);
			si.lpReserved = si.lpDesktop = si.lpTitle = NULL;
			si.lpReserved2 = NULL;
			si.dwFlags = STARTF_USESHOWWINDOW;  // Tell it to use the value of wShowWindow below.
			si.wShowWindow = sei.nShow;
			PROCESS_INFORMATION pi = {0};
			// Since CreateProcess() requires that the 2nd param be modifiable, ensure that it is
			// (even if this is ANSI and not Unicode; it's just safer):
			strlcpy(action, aAction_orig, sizeof(action)); // i.e. we're running the original action from caller.
			// MSDN: "If [lpCurrentDirectory] is NULL, the new process is created with the same
			// current drive and directory as the calling process." (i.e. since caller may have
			// specified a NULL aWorkingDir):
			success = CreateProcess(NULL, aAction, NULL, NULL, FALSE, 0, NULL, aWorkingDir, &si, &pi);
			sei.hProcess = success ? pi.hProcess : NULL;  // Set this value for use later on.
		}

		if (!success)
		{
			if (aDisplayErrors)
			{
				char error_text[2048], verb_text[128];
				if (*sei.lpVerb && stricmp(sei.lpVerb, "open"))
					snprintf(verb_text, sizeof(verb_text), "\nVerb: <%s>", sei.lpVerb);
				else // Don't bother showing it if it's just "open".
					*verb_text = '\0';
				// Use format specifier to make sure it doesn't get too big for the error
				// function to display.  Also, due to above having tried CreateProcess()
				// as a last resort, aParams will always be blank so don't bother displaying it:
				snprintf(error_text, sizeof(error_text)
					, "Failed attempt to launch program or document:\nAction: <%-1.400s%s>%s"
					, aAction, strlen(aAction) > 400 ? "..." : "", verb_text);
				ScriptError(error_text);
			}
			return FAIL;
		}
	}

	if (aProcess) // The caller wanted the process handle, so provide it.
		*aProcess = sei.hProcess;
	return OK;
}
