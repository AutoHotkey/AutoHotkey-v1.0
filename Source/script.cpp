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
#include "script.h"
#include "globaldata.h" // for a lot of things
#include "util.h" // for strlcpy() etc.
#include "mt19937ar-cok.h" // for random number generator
#include "window.h" // for a lot of things
#include "application.h" // for MsgSleep()


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
	, mLoopFile(NULL), mLoopRegItem(NULL)
	, mThisHotkeyLabel(NULL), mPriorHotkeyLabel(NULL), mPriorHotkeyStartTime(0)
	, mFirstLabel(NULL), mLastLabel(NULL)
	, mFirstVar(NULL), mLastVar(NULL)
	, mLineCount(0), mLabelCount(0), mVarCount(0), mGroupCount(0)
	, mCurrFileNumber(0), mCurrLineNumber(0)
	, mFileSpec(""), mFileDir(""), mFileName(""), mOurEXE(""), mOurEXEDir(""), mMainWindowTitle("")
	, mIsReadyToExecute(false)
	, mIsRestart(false)
	, mIsAutoIt2(false)
	, mLinesExecutedThisCycle(0)
	, mLastSleepTime(GetTickCount())
{
	ZeroMemory(&mNIC, sizeof(mNIC));  // Constructor initializes this, to be safe.
	mNIC.hWnd = NULL;  // Set this as an indicator that it tray icon is not installed.
#ifdef _DEBUG
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

	

ResultType Script::CreateWindows(HINSTANCE aInstance)
// Returns OK or FAIL.
{
	if (!mMainWindowTitle || !*mMainWindowTitle) return FAIL;  // Init() must be called before this function.
	// Register a window class for the main window:
	HICON hIcon = LoadIcon(aInstance, MAKEINTRESOURCE(IDI_MAIN)); // LoadIcon((HINSTANCE) NULL, IDI_APPLICATION)
	WNDCLASSEX wc;
	ZeroMemory(&wc, sizeof(wc));
	wc.cbSize = sizeof(wc);
	wc.lpszClassName = WINDOW_CLASS_MAIN;
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

	// Note: the title below must be constructed the same was as is done by our
	// WinMain() (so that we can detect whether this script is already running)
	// which is why it's standardized in g_script.mMainWindowTitle.
	// Create the main window:
	if (   !(g_hWnd = CreateWindow(
		  WINDOW_CLASS_MAIN
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
	//ShowWindow(g_hWnd, SW_HIDE);
	//ShowWindow(g_hWnd, SW_HIDE); // 2nd call to be safe.
	// Update: Doing it this way prevents the launch of the program from "stealing focus"
	// (changing the foreground window to be nothing).  This allows scripts launched from the
	// start menu, for example, to immediately operate on the window that was foreground
	// prior to the start menu having been displayed:
	ShowWindow(g_hWnd, SW_MINIMIZE);
	ShowWindow(g_hWnd, SW_HIDE);

	g_hAccelTable = LoadAccelerators(aInstance, MAKEINTRESOURCE(IDR_ACCELERATOR1));

	////////////////////
	// Set up tray icon.
	////////////////////
	if (g_NoTrayIcon)
		mNIC.hWnd = NULL;  // Set this as an indicator that tray icon is not installed.
	else
	{
		ZeroMemory(&mNIC, sizeof(mNIC));  // To be safe.
		// Using NOTIFYICONDATA_V2_SIZE vs. sizeof(NOTIFYICONDATA) improves compatibility with Win9x maybe.
		// MSDN: "Using [NOTIFYICONDATA_V2_SIZE] for cbSize will allow your application to use NOTIFYICONDATA
		// with earlier Shell32.dll versions, although without the version 6.0 enhancements."
		// Update: Using V2 gives an compile error so trying V1.  Update: Trying sizeof(NOTIFYICONDATA)
		// for compatibility with VC++ 6.x.  This is also what AutoIt3 uses:
		mNIC.cbSize				= sizeof(NOTIFYICONDATA);  // NOTIFYICONDATA_V1_SIZE
		mNIC.hWnd				= g_hWnd;
		mNIC.uID				= 0;  // Icon ID (can be anything, like Timer IDs?)
		mNIC.uFlags				= NIF_MESSAGE | NIF_TIP | NIF_ICON;
		mNIC.uCallbackMessage	= AHK_NOTIFYICON;
		mNIC.hIcon				= LoadIcon(aInstance, MAKEINTRESOURCE(IDI_MAIN));
		strlcpy(mNIC.szTip, mFileName ? mFileName : NAME_P, sizeof(mNIC.szTip));
		if (!Shell_NotifyIcon(NIM_ADD, &mNIC))
		{
			mNIC.hWnd = NULL;  // Set this as an indicator that tray icon is not installed.
			return OK;  // But don't return FAIL (in case the user is using a different shell or something).
		}
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
		icon = IDI_PAUSE_SUSPEND;
	else if (g.IsPaused)
		icon = IDI_PAUSE;
	else if (g_IsSuspended)
		icon = IDI_SUSPEND;
	else
		icon = IDI_MAIN;
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
	// This is here in case a compiled script ever uses the Edit command.  Since the "Edit This
	// Script" menu item is not available for compiled scripts, it can't be called from there.
#ifdef AUTOHOTKEYSC
	return OK; // Do nothing.
#endif
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
				MsgBox("Could not open the file for editing using the associated \"edit\" action or Notepad.");
		}
	}
	return OK;
}



ResultType Script::Reload(bool aDisplayErrors)
{
	char current_dir[MAX_PATH];
	GetCurrentDirectory(sizeof(current_dir), current_dir);  // In case the user launched it in a non-default dir.
	// The new instance we're about to start will tell our process to stop, or it will display
	// a syntax error or some other error, in which case our process will still be running:
#ifdef AUTOHOTKEYSC
	// This is here in case a compiled script ever uses the Edit command.  Since the "Reload This
	// Script" menu item is not available for compiled scripts, it can't be called from there.
	return g_script.ActionExec(mOurEXE, "/restart", current_dir, aDisplayErrors);
#else
	char arg_string[MAX_PATH + 512];
	snprintf(arg_string, sizeof(arg_string), "/restart \"%s\"", mFileSpec);
	return g_script.ActionExec(mOurEXE, arg_string, current_dir, aDisplayErrors);
#endif
}



void Script::ExitApp(char *aBuf, int aExitCode)
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
		MessageBox(g_hWnd, buf, g_script.mFileSpec, MB_OK | MB_SETFOREGROUND | MB_APPLMODAL);
	}
	KeyHistoryToFile();  // Close the KeyHistory file if it's open.
	PostQuitMessage(aExitCode); // This might be needed to prevent hang-on-exit.
	Hotkey::AllDestructAndExit(*aBuf ? CRITICAL_ERROR : aExitCode); // Terminate the application.
	// Not as reliable: PostQuitMessage(CRITICAL_ERROR);
}



LineNumberType Script::LoadFromFile()
// Returns the number of non-comment lines that were loaded, or LOADING_FAILED on error.
{
	mIsReadyToExecute = false;
	if (!mFileSpec || !*mFileSpec) return LOADING_FAILED;

#ifndef AUTOHOTKEYSC  // Not in stand-alone mode, so read an external script file.
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
		char full_path[MAX_PATH * 2]; // In case OS supports extra-long filenames.
		char *filename_marker;
		GetFullPathName(aFileSpec, sizeof(full_path) - 1, full_path, &filename_marker);
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
	if ( oRead.Open(aFileSpec, "") != HS_EXEARC_E_OK)
	{
		MsgBox("Could not open the script inside the EXE.", 0, aFileSpec);
		return FAIL;
	}
	// AutoIt3: Read the script (the func allocates the memory for the buffer :) )
	if ( oRead.FileExtractToMem(">AUTOHOTKEY SCRIPT<", &script_buf, &nDataSize) != HS_EXEARC_E_OK)
	{
		oRead.Close();							// Close the archive
		MsgBox("Could not extract the script from the EXE into memory.", 0, aFileSpec);
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
	// Future: might be best to put a stat() in here for better handling.
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
	char buf[LINE_SIZE], *hotkey_flag, *cp;
	HookActionType hook_action;
	size_t buf_length;
	bool is_label, section_comment = false;

	// Init both for main file and any included files loaded by this function:
	mCurrFileNumber = source_file_number;  // source_file_number is kept on the stack due to recursion.
	LineNumberType line_number = mCurrLineNumber = 0;  // Keep a copy on the stack to help with recursion.

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
			{
				if (AddLine(ACT_RETURN) != OK)
					return CloseAndReturn(fp, script_buf, FAIL);
				mCurrLine = NULL;  // To signify that we're in transition, trying to load a new one.
			}
			if (AddLabel(buf) != OK) // Always add a label before adding the first line of its section.
				return CloseAndReturn(fp, script_buf, FAIL);
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
						return CloseAndReturn(fp, script_buf, FAIL);
				// Also add a Return that's implicit for a single-line hotkey:
				if (AddLine(ACT_RETURN) != OK)
					return CloseAndReturn(fp, script_buf, FAIL);
			}
			else
				hook_action = 0;
			// Set the new hotkey will jump to this label to begin execution:
			if (Hotkey::AddHotkey(mLastLabel, hook_action) != OK)
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
	char *cp, *directive_end;
	int value; // Helps detect values that are too large, since some of the target globals are UCHAR.

	// Use strnicmp() so that a match is found as long as aBuf starts with the string in question.
	// e.g. so that "#SingleInstance, on" will still work too, but
	// "#a::run, something, "#SingleInstance" (i.e. a hotkey) will not be falsely detected
	// due to using a more lenient function such as stristr().  UPDATE: Using strlicmp() now so
	// that overlapping names, such as #MaxThreads and #MaxThreadsPerHotkey won't get mixed up:
	if (   !(directive_end = StrChrAny(aBuf, end_flags))   )
		directive_end = aBuf + strlen(aBuf); // Point it to the zero terminator.
	#define IS_DIRECTIVE_MATCH(directive) (!strlicmp(aBuf, directive, (UINT)(directive_end - aBuf)))

	if (IS_DIRECTIVE_MATCH("#WinActivateForce"))
	{
		g_WinActivateForce = true;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#SingleInstance"))
	{
		g_AllowOnlyOneInstance = true;
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
		if (*cp == g_delimiter)\
		{\
			++cp;\
			if (   !*(cp = omit_leading_whitespace(cp))   )\
				return CONDITION_TRUE;\
		}
		RETURN_IF_NO_CHAR
		if (Line::ConvertOnOff(cp) == TOGGLED_OFF)
			g_ForceKeybdHook = false;
		// else leave the default to "true" as set above.
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#InstallKeybdHook"))
	{
		if (g_os.IsWin9x())
			MsgBox("#InstallKeybdHook is not supported on Windows 95/98/ME.  This line will be ignored.");
		else
			Hotkey::RequireHook(HOOK_KEYBD);
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#InstallMouseHook"))
	{
		if (g_os.IsWin9x())
			MsgBox("#InstallMouseHook is not supported on Windows 95/98/ME.  This line will be ignored.");
		else
			Hotkey::RequireHook(HOOK_MOUSE);
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#HotkeyModifierTimeout"))
	{
		RETURN_IF_NO_CHAR
		g_HotkeyModifierTimeout = atoi(cp);  // cp was set to the right position by the above macro
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#MaxThreads"))
	{
		RETURN_IF_NO_CHAR
		value = atoi(cp);  // cp was set to the right position by the above macro
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
		value = atoi(cp);  // cp was set to the right position by the above macro
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
		g_HotkeyThrottleInterval = atoi(cp);  // cp was set to the right position by the above macro
		if (g_HotkeyThrottleInterval < 10) // values under 10 wouldn't be useful due to timer granularity.
			g_HotkeyThrottleInterval = 10;
		return CONDITION_TRUE;
	}
	if (IS_DIRECTIVE_MATCH("#MaxHotkeysPerInterval"))
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
		// to what it already is, and even if this is attempted, it would just be ignored:
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

	#define DEFINE_END_FLAGS \
		char end_flags[] = {' ', g_delimiter, '\t', '<', '>', '=', '+', '-', '*', '/', '!', '\0'}; // '\0' must be last.
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
			aLineText[new_length++] = ',';
			aLineText[new_length++] = '1';
			aLineText[new_length] = '\0';
			action_args = aLineText;
		}
		else if (!stricmp(action_name, "IF"))
		{
			// Skip over the variable name so that the "is" and "is not" operators are properly supported:
			char *operation = StrChrAny(action_args, end_flags);
			if (!operation)
				operation = action_args + strlen(action_args); // Point it to the NULL terminator instead.
			else
				operation = omit_leading_whitespace(operation);
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
			case 'i':  // "is" or "is not"
			case 'I':
				if (toupper(*(operation + 1)) == 'S')
				{
					char *next_word = omit_leading_whitespace(operation + 2);
					if (strnicmp(next_word, "not", 3))
						action_type = ACT_IFIS;
					else
					{
						action_type = ACT_IFISNOT;
						// Remove the word "not" to set things up to be parsed as args further down.
						*next_word = ' ';
						*(next_word + 1) = ' ';
						*(next_word + 2) = ' ';
					}
					*(operation + 1) = ' '; // Remove the 'S' in "IS".  'I' is replaced with ',' later below.
				}
				else
					return ScriptError("The word IS was expected but not found.", aLineText);
				break;
			default:
				return ScriptError("Although this line is an IF, it lacks operator symbol(s).", aLineText);
				// Note: User can use whitespace to differentiate a literal symbol from
				// part of an operator, e.g. if var1 < =  <--- char is literal
			} // switch()
			// Set things up to be parsed as args later on:
			*operation = g_delimiter;
		}
		else // The action type is something other than an IF.
		{
			if (*action_args == '=')
				action_type = ACT_ASSIGN;
			else if (*action_args == '+' && (*(action_args + 1) == '=' || *(action_args + 1) == '+'))
				action_type = ACT_ADD;
			else if (*action_args == '-' && (*(action_args + 1) == '=' || *(action_args + 1) == '-'))
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
		const int max_msgbox_delimiters = 20;
		char *delimiter[max_msgbox_delimiters];
		int delimiter_count;
		for (mark = delimiter_count = 0; action_args[mark] && delimiter_count < max_msgbox_delimiters;)
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
			if (delimiter_count <= 1)
				// Force it to be 1-param mode.  In other words, we want to make MsgBox a very forgiving
				// command and have it rarely if ever report syntax errors:
				max_params_override = 1;
			else // It has more than 3 apparent params, but is the first param even numeric?
			{
				*delimiter[0] = '\0'; // Temporarily terminate action_args at the first delimiter.
				if (!IsPureNumeric(action_args, false, false, false)) // No floats allowed in this case.
					max_params_override = 1;
				*delimiter[0] = g_delimiter; // Restore the string.
				if (!max_params_override)
				{
					// The above has determined that the cmd isn't in 1-parameter mode.
					// If at this point it has exactly 3 apparent params, allow the command to be
					// processed normally without an override.  Otherwise, do more checking:
					if (delimiter_count > 2) // i.e. 3 or more delimiters, which means 4 or more params.
					{
						// If the last parameter isn't blank or pure numeric (i.e. even if it's a pure
						// deref, since trying to figure out what's a pure deref is somewhat complicated
						// at this early stage of parsing), assume the user didn't intend it to be the
						// MsgBox timeout (since that feature is rarely used):
						if (!IsPureNumeric(delimiter[delimiter_count-1] + 1, false, true, true))
							// Not blank and not a int or float.
							max_params_override = 3;
						// If it has more than 4 params or it has exactly 4 but the 4th isn't blank,
						// pure numeric, or a deref: assume it's being used in 3-parameter mode and
						// that all the other delimiters were intended to be literal.
					}
				}
			}
		}
	}


	////////////////////////////////////////////////////////////
	// Parse the parmeter string into a list of separate params.
	////////////////////////////////////////////////////////////
	// MaxParams has already been verified as being <= MAX_ARGS.
	// Any g_delimiter-delimited items beyond MaxParams will be included in a lump inside the last param:
	int nArgs;
	char *arg[MAX_ARGS], *arg_map[MAX_ARGS];
	ActionTypeType subaction_type = ACT_INVALID; // Must init these.
	ActionTypeType suboldaction_type = OLD_INVALID;
	char subaction_name[MAX_VAR_NAME_LENGTH + 1], *subaction_end_marker = NULL, *subaction_start = NULL;
	int max_params = max_params_override ? max_params_override : this_action->MaxParams;
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
		if (nArgs == max_params - 1)
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
		return ScriptError("AddLine() called incorrectly.", aArgc > 0 ? aArg[0] : "");

	char error_msg[1024];
	Var *target_var;

	//////////////////////////////////////////////////////////
	// Build the new arg list in dynamic memory.
	// The allocated structs will be attached to the new line.
	//////////////////////////////////////////////////////////
	DerefType deref[MAX_DEREFS_PER_ARG];  // Will be used to temporarily store the var-deref locations in each arg.
	int deref_count;  // How many items are in deref array.
	ArgStruct *new_arg;  // We will allocate some dynamic memory for this, then hang it onto the new line.
	size_t deref_string_length;
	if (!aArgc)
		new_arg = NULL;  // Just need an empty array in this case.
	else
	{
		if (   !(new_arg = (ArgStruct *)SimpleHeap::Malloc(aArgc * sizeof(ArgStruct)))   )
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
			new_arg[i].type = Line::ArgIsVar(aActionType, i);
			// Since some vars are optional, the below allows them all to be blank or
			// not present in the arg list.  If mandatory var is blank at this stage,
			// it's okay because all mandatory args are validated to be non-blank elsewhere:
			if (new_arg[i].type != ARG_TYPE_NORMAL)
			{
				if (!*aArg[i])
					// An optional input or output variable has been omitted, so indicate
					// that this arg is not a variable, just a normal empty arg.  Functions
					// such as ListLines() rely on this having been done because they assume,
					// for performance reasons, that args marked as variables really are
					// variables:
					new_arg[i].type = ARG_TYPE_NORMAL;
				else
				{
					// Does this input or output variable contain a dereference?  If so, it must
					// be resolved at runtime (to support old-style AutoIt2 arrays, etc.).
					// Find the first non-escaped dereference symbol:
					for (j = 0; aArg[i][j] && (aArg[i][j] != g_DerefChar || (aArgMap && aArgMap[i] && aArgMap[i][j])); ++j);
					if (!aArg[i][j])
					{
						// A non-escaped deref symbol wasn't found, therefore this variable does not
						// appear to be something that must be resolved dynamically at runtime.
						if (   !(target_var = FindOrAddVar(aArg[i]))   )
							return FAIL;  // The above already displayed the error.
						// If this action type is something that modifies the contents of the var, ensure the var
						// isn't a special/reserved one:
						if (new_arg[i].type == ARG_TYPE_OUTPUT_VAR && VAR_IS_RESERVED(target_var))
							return ScriptError(ERR_VAR_IS_RESERVED, aArg[i]);
						// Rather than removing this arg from the list altogether -- which would distrub
						// the ordering and hurt the maintainability of the code -- the next best thing
						// in terms of saving memory is to store an empty string in place of the arg's
						// text if that arg is a pure variable (i.e. since the name of the variable is already
						// stored in the Var object, we don't need to store it twice):
						new_arg[i].text = "";
						new_arg[i].deref = (DerefType *)target_var;
						continue;
					}
					// else continue on so that this input or output variable name's dynamic part (e.g. array%i%)
					// can be partially resolved:
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
	Line *line = new Line(mCurrFileNumber, mCurrLineNumber, aActionType, new_arg, aArgc);
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
		bool allow_negative, allow_float;
		for (ActionTypeType *np = g_act[aActionType].NumericParams; *np; ++np)
		{
			if (line->mArgc >= *np)  // The arg exists.
			{
				if (!line->ArgHasDeref(*np)) // i.e. if it's a deref, we won't try to validate it now.
				{
					allow_negative = line->ArgAllowsNegative(*np);
					allow_float = line->ArgAllowsFloat(*np);
					if (!IsPureNumeric(line->mArg[*np - 1].text, allow_negative, true, allow_float))
					{
						if (aActionType == ACT_WINMOVE)
						{
							if (stricmp(line->mArg[*np - 1].text, "default"))
							{
								snprintf(error_msg, sizeof(error_msg), "\"%s\" requires parameter #%u to be"
									" either %snumeric, a variable reference, blank, or the word Default."
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
								" either %snumeric, blank (if blank is allowed), or a variable reference."
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
	int value;    // For temp use during validation.
	FILETIME ft;  // same.
	switch(aActionType)
	{
	case ACT_AUTOTRIM:
	case ACT_STRINGCASESENSE:
	case ACT_DETECTHIDDENWINDOWS:
	case ACT_DETECTHIDDENTEXT:
	case ACT_BLOCKINPUT:
	case ACT_SETSTORECAPSLOCKMODE:
		if (line->mArgc > 0 && !line->ArgHasDeref(1) && !line->ConvertOnOff(LINE_RAW_ARG1))
			return ScriptError(ERR_ON_OFF, LINE_RAW_ARG1);
		break;

	case ACT_SUSPEND:
		if (line->mArgc > 0 && !line->ArgHasDeref(1) && !line->ConvertOnOffTogglePermit(LINE_RAW_ARG1))
			return ScriptError(ERR_ON_OFF_TOGGLE_PERMIT, LINE_RAW_ARG1);
		break;

	case ACT_PAUSE:
	case ACT_KEYHISTORY:
		if (line->mArgc > 0 && !line->ArgHasDeref(1) && !line->ConvertOnOffToggle(LINE_RAW_ARG1))
			return ScriptError(ERR_ON_OFF_TOGGLE, LINE_RAW_ARG1);
		break;

	case ACT_SETNUMLOCKSTATE:
	case ACT_SETSCROLLLOCKSTATE:
	case ACT_SETCAPSLOCKSTATE:
		if (line->mArgc > 0 && !line->ArgHasDeref(1) && !line->ConvertOnOffAlways(LINE_RAW_ARG1))
			return ScriptError(ERR_ON_OFF_ALWAYS, LINE_RAW_ARG1);
		break;

	case ACT_STRINGMID:
	case ACT_FILEREADLINE:
		if (line->mArgc > 2 && !line->ArgHasDeref(3))
		{
			if (!IsPureNumeric(LINE_RAW_ARG3, false, false, false) || !_atoi64(LINE_RAW_ARG3))
				// This error is caught at load-time, but at runtime it's not considered
				// an error (i.e. if a variable resolves to zero or less, StringMid will
				// automatically consider it to be 1, though FileReadLine would consider
				// it an error):
				return ScriptError("Parameter #3 must be a number greater than zero or a variable reference."
					, LINE_RAW_ARG3);
		}
		break;

	case ACT_STRINGGETPOS:
		if (   line->mArgc > 3 && !line->ArgHasDeref(4) && *LINE_RAW_ARG4
			&& (strlen(LINE_RAW_ARG4) > 1 || (toupper(*LINE_RAW_ARG4) != 'R' && *LINE_RAW_ARG4 != '1'))   )
			return ScriptError("If not blank, parameter #4 must be 1, R, or a variable reference.", LINE_RAW_ARG4);
		break;

	case ACT_STRINGREPLACE:
		if (line->mArgc > 4 && !line->ArgHasDeref(5) && *LINE_RAW_ARG5
			&& ((!*(LINE_RAW_ARG5 + 1) && *LINE_RAW_ARG5 != '1' && toupper(*LINE_RAW_ARG5) != 'A')
			|| (*(LINE_RAW_ARG5 + 1) && stricmp(LINE_RAW_ARG5, "all"))))
			return ScriptError("If not blank, parameter #5 must be 1, A, ALL, or a variable reference.", LINE_RAW_ARG5);
		break;

	case ACT_REGREAD:
		if (line->mArgc > 4) // The obsolete 5-param method is being used, wherein ValueType is the 2nd param.
		{
			if (!line->ArgHasDeref(3) && *LINE_RAW_ARG3 && !line->RegConvertRootKey(LINE_RAW_ARG3))
				return ScriptError(ERR_REG_KEY, LINE_RAW_ARG3);
		}
		else
			if (line->mArgc > 1 && !line->ArgHasDeref(2) && *LINE_RAW_ARG2 && !line->RegConvertRootKey(LINE_RAW_ARG2))
				return ScriptError(ERR_REG_KEY, LINE_RAW_ARG2);
		break;

	case ACT_REGWRITE:
		// Both of these checks require that at least two parameters be present.  Otherwise, the command
		// is being used in its registry-loop mode and is validated elsewhere:
		if (line->mArgc > 1)
		{
			if (!line->ArgHasDeref(1) && *LINE_RAW_ARG1 && !line->RegConvertValueType(LINE_RAW_ARG1))
				return ScriptError(ERR_REG_VALUE_TYPE, LINE_RAW_ARG1);
			if (!line->ArgHasDeref(2) && *LINE_RAW_ARG2 && !line->RegConvertRootKey(LINE_RAW_ARG2))
				return ScriptError(ERR_REG_KEY, LINE_RAW_ARG2);
		}
		break;

	case ACT_REGDELETE:
		if (line->mArgc > 0 && !line->ArgHasDeref(1) && *LINE_RAW_ARG1 && !line->RegConvertRootKey(LINE_RAW_ARG1))
			return ScriptError(ERR_REG_KEY, LINE_RAW_ARG1);
		break;

	case ACT_SOUNDSETWAVEVOLUME:
		if (line->mArgc > 0 && !line->ArgHasDeref(1))
		{
			value = atoi(LINE_RAW_ARG1);
			if (value < 0 || value > 100)
				return ScriptError(ERR_PERCENT, LINE_RAW_ARG1);
		}
		break;

	case ACT_SOUNDPLAY:
		if (line->mArgc > 1 && !line->ArgHasDeref(2) && *LINE_RAW_ARG2
			&& stricmp(LINE_RAW_ARG2, "wait") && stricmp(LINE_RAW_ARG2, "1"))
			return ScriptError("If not blank, parameter #2 must be 1, WAIT, or a variable reference.", LINE_RAW_ARG2);
		break;

	case ACT_PIXELSEARCH:
		if (line->mArgc > 7 && !line->ArgHasDeref(8) && *LINE_RAW_ARG8)
		{
			value = atoi(LINE_RAW_ARG8);
			if (value < 0 || value > 255)
				return ScriptError("Parameter #8 must be number between 0 and 255, blank, or a variable reference."
					, LINE_RAW_ARG8);
		}
		break;

	case ACT_MOUSEMOVE:
		if (line->mArgc > 2 && !line->ArgHasDeref(3) && *LINE_RAW_ARG3)
		{
			value = atoi(LINE_RAW_ARG3);
			if (value < 0 || value > MAX_MOUSE_SPEED)
				return ScriptError(ERR_MOUSE_SPEED, LINE_RAW_ARG3);
		}
		if (!line->ValidateMouseCoords(LINE_RAW_ARG1, LINE_RAW_ARG2))
			return ScriptError(ERR_MOUSE_COORD, LINE_RAW_ARG1);
		break;
	case ACT_MOUSECLICK:
		if (line->mArgc > 4 && !line->ArgHasDeref(5) && *LINE_RAW_ARG5)
		{
			value = atoi(LINE_RAW_ARG5);
			if (value < 0 || value > MAX_MOUSE_SPEED)
				return ScriptError(ERR_MOUSE_SPEED, LINE_RAW_ARG5);
		}
		if (line->mArgc > 5 && !line->ArgHasDeref(6) && *LINE_RAW_ARG6)
			if (strlen(LINE_RAW_ARG6) > 1 || !strchr("UD", toupper(*LINE_RAW_ARG6)))  // Up / Down
				return ScriptError(ERR_MOUSE_UPDOWN, LINE_RAW_ARG6);
		// Check that the button is valid (e.g. left/right/middle):
		if (!line->ArgHasDeref(1))
			if (!line->ConvertMouseButton(LINE_RAW_ARG1))
				return ScriptError(ERR_MOUSE_BUTTON, LINE_RAW_ARG1);
		if (!line->ValidateMouseCoords(LINE_RAW_ARG2, LINE_RAW_ARG3))
			return ScriptError(ERR_MOUSE_COORD, LINE_RAW_ARG2);
		break;
	case ACT_MOUSECLICKDRAG:
		if (line->mArgc > 5 && !line->ArgHasDeref(6) && *LINE_RAW_ARG6)
		{
			value = atoi(LINE_RAW_ARG6);
			if (value < 0 || value > MAX_MOUSE_SPEED)
				return ScriptError(ERR_MOUSE_SPEED, LINE_RAW_ARG6);
		}
		if (!line->ArgHasDeref(1))
			if (!line->ConvertMouseButton(LINE_RAW_ARG1))
				return ScriptError(ERR_MOUSE_BUTTON, LINE_RAW_ARG1);
		if (!line->ValidateMouseCoords(LINE_RAW_ARG2, LINE_RAW_ARG3))
			return ScriptError(ERR_MOUSE_COORD, LINE_RAW_ARG2);
		if (!line->ValidateMouseCoords(LINE_RAW_ARG4, LINE_RAW_ARG5))
			return ScriptError(ERR_MOUSE_COORD, LINE_RAW_ARG4);
		break;
	case ACT_SETDEFAULTMOUSESPEED:
		if (line->mArgc > 0 && !line->ArgHasDeref(1) && *LINE_RAW_ARG1)
		{
			value = atoi(LINE_RAW_ARG1);
			if (value < 0 || value > MAX_MOUSE_SPEED)
				return ScriptError(ERR_MOUSE_SPEED, LINE_RAW_ARG1);
		}
		break;

	case ACT_CONTROLCLICK:
		// Check that the button is valid (e.g. left/right/middle):
		if (!line->ArgHasDeref(4) && *LINE_RAW_ARG4) // i.e. it's allowed to be blank (defaults to left).
			if (!line->ConvertMouseButton(LINE_RAW_ARG4))
				return ScriptError(ERR_MOUSE_BUTTON, LINE_RAW_ARG4);
		if (line->mArgc > 5 && !line->ArgHasDeref(6) && *LINE_RAW_ARG6)
			if (strlen(LINE_RAW_ARG6) > 1 || !strchr("UD", toupper(*LINE_RAW_ARG6)))  // Up / Down
				return ScriptError(ERR_MOUSE_UPDOWN, LINE_RAW_ARG6);
		break;

	case ACT_ADD:
	case ACT_SUB:
		if (line->mArgc > 2)
		{
			if (!line->ArgHasDeref(3) && *LINE_RAW_ARG3)
				if (!strchr("SMHD", toupper(*LINE_RAW_ARG3)))  // (S)econds, (M)inutes, (H)ours, or (D)ays
					return ScriptError(ERR_COMPARE_TIMES, LINE_RAW_ARG3);
			if (aActionType == ACT_SUB && !line->ArgHasDeref(2) && *LINE_RAW_ARG2)
				if (!YYYYMMDDToFileTime(LINE_RAW_ARG2, &ft))
					return ScriptError(ERR_INVALID_DATETIME, LINE_RAW_ARG2);
		}
		break;

	case ACT_FILEINSTALL:
	case ACT_FILECOPY:
	case ACT_FILEMOVE:
	case ACT_FILECOPYDIR:
	case ACT_FILEMOVEDIR:
	case ACT_FILESELECTFOLDER:
		if (line->mArgc > 2 && !line->ArgHasDeref(3) && *LINE_RAW_ARG3)
		{
			if (strlen(LINE_RAW_ARG3) > 1 || (*LINE_RAW_ARG3 != '0' && *LINE_RAW_ARG3 != '1'))
				return ScriptError("Parameter #3 must be either blank, 0, 1, or a variable reference."
					, LINE_RAW_ARG3);
		}
		if (aActionType == ACT_FILEINSTALL)
		{
			if (line->mArgc > 0 && line->ArgHasDeref(1))
				return ScriptError("Parameter #1 must not contain references to variables."
					, LINE_RAW_ARG1);
		}
		break;

	case ACT_FILEREMOVEDIR:
		if (line->mArgc > 1 && !line->ArgHasDeref(2) && *LINE_RAW_ARG2)
		{
			if (strlen(LINE_RAW_ARG2) > 1 || (*LINE_RAW_ARG2 != '0' && *LINE_RAW_ARG2 != '1'))
				return ScriptError("Parameter #2 must be either blank, 0, 1, or a variable reference."
					, LINE_RAW_ARG2);
		}
		break;

	case ACT_FILESETATTRIB:
		if (line->mArgc > 0 && !line->ArgHasDeref(1) && *LINE_RAW_ARG1)
		{
			for (char *cp = LINE_RAW_ARG1; *cp; ++cp)
				if (!strchr("+-^RASHNOT", toupper(*cp)))
					return ScriptError("Parameter #1 contains unsupported file-attribute letters or symbols."
						, LINE_RAW_ARG1);
		}
		if (line->mArgc > 2 && !line->ArgHasDeref(3) && line->ConvertLoopMode(LINE_RAW_ARG3) == FILE_LOOP_INVALID)
			return ScriptError("If not blank, parameter #3 must be either 0, 1, 2, or a variable reference."
				, LINE_RAW_ARG3);
		if (line->mArgc > 3 && !line->ArgHasDeref(4) && *LINE_RAW_ARG4)
			if (strlen(LINE_RAW_ARG4) > 1 || (*LINE_RAW_ARG4 != '0' && *LINE_RAW_ARG4 != '1'))
				return ScriptError("Parameter #4 must be either blank, 0, 1, or a variable reference."
					, LINE_RAW_ARG4);
		break;

	case ACT_FILEGETTIME:
		if (line->mArgc > 2 && !line->ArgHasDeref(3) && *LINE_RAW_ARG3)
			if (strlen(LINE_RAW_ARG3) > 1 || !strchr("MCA", toupper(*LINE_RAW_ARG3)))
				return ScriptError(ERR_FILE_TIME, LINE_RAW_ARG3);
		break;

	case ACT_FILESETTIME:
		if (line->mArgc > 0 && !line->ArgHasDeref(1) && *LINE_RAW_ARG1)
			if (!YYYYMMDDToFileTime(LINE_RAW_ARG1, &ft))
				return ScriptError(ERR_INVALID_DATETIME, LINE_RAW_ARG1);
		if (line->mArgc > 2 && !line->ArgHasDeref(3) && *LINE_RAW_ARG3)
			if (strlen(LINE_RAW_ARG3) > 1 || !strchr("MCA", toupper(*LINE_RAW_ARG3)))
				return ScriptError(ERR_FILE_TIME, LINE_RAW_ARG3);
		if (line->mArgc > 3 && !line->ArgHasDeref(4) && line->ConvertLoopMode(LINE_RAW_ARG4) == FILE_LOOP_INVALID)
			return ScriptError("If not blank, parameter #4 must be either 0, 1, 2, or a variable reference."
				, LINE_RAW_ARG4);
		if (line->mArgc > 4 && !line->ArgHasDeref(5) && *LINE_RAW_ARG5)
			if (strlen(LINE_RAW_ARG5) > 1 || (*LINE_RAW_ARG5 != '0' && *LINE_RAW_ARG5 != '1'))
				return ScriptError("Parameter #5 must be either blank, 0, 1, or a variable reference."
					, LINE_RAW_ARG5);
		break;

	case ACT_FILEGETSIZE:
		if (line->mArgc > 2 && !line->ArgHasDeref(3) && *LINE_RAW_ARG3)
			if (strlen(LINE_RAW_ARG3) > 1 || !strchr("BKM", toupper(*LINE_RAW_ARG3))) // Allow B=Bytes as undocumented.
				return ScriptError("Parameter #3 must be either blank, K, M, or a variable reference."
					, LINE_RAW_ARG3);
		break;

	case ACT_FILESELECTFILE:
		if (line->mArgc > 1 && !line->ArgHasDeref(2) && *LINE_RAW_ARG2)
		{
			value = atoi(LINE_RAW_ARG2);
			if (value < 0 || value > 31)
				return ScriptError("Paremeter #2 must be either blank, a variable reference,"
					" or a number between 0 and 31 inclusive.", LINE_RAW_ARG2);
		}
		break;

	case ACT_SETTITLEMATCHMODE:
		if (line->mArgc > 0 && !line->ArgHasDeref(1) && !line->ConvertTitleMatchMode(LINE_RAW_ARG1))
			return ScriptError(ERR_TITLEMATCHMODE, LINE_RAW_ARG1);
		break;
	case ACT_SETFORMAT:
		if (line->mArgc > 0 && !line->ArgHasDeref(1) && stricmp(LINE_RAW_ARG1, "float"))
			return ScriptError("Parameter #1 must be the word FLOAT or a variable reference.", LINE_RAW_ARG1);
		// Size must be less than sizeof() minus 2 because need room to prepend the '%' and append
		// the 'f' to make it a valid format specifier string:
		if (line->mArgc > 1 && !line->ArgHasDeref(2) && strlen(LINE_RAW_ARG2) >= sizeof(g.FormatFloat) - 2)
			return ScriptError("Parameter #2 is too long.", LINE_RAW_ARG1);
		break;

	case ACT_MSGBOX:
		if (line->mArgc > 1) // i.e. this MsgBox is using the 3-param or 4-param style.
			if (!line->ArgHasDeref(1)) // i.e. if it's a deref, we won't try to validate it now.
				if (!IsPureNumeric(LINE_RAW_ARG1))
					return ScriptError("When used with more than one parameter, MsgBox requires that"
						" the 1st parameter be numeric or a variable reference.", LINE_RAW_ARG1);
		if (line->mArgc > 3)
			if (!line->ArgHasDeref(4)) // i.e. if it's a deref, we won't try to validate it now.
				if (!IsPureNumeric(LINE_RAW_ARG4, false, true, true))
					return ScriptError("MsgBox requires that the 4th parameter, if present, be numeric & positive,"
						" or a variable reference.", LINE_RAW_ARG4);
		break;
	case ACT_IFMSGBOX:
		if (line->mArgc > 0 && !line->ArgHasDeref(1) && !line->ConvertMsgBoxResult(LINE_RAW_ARG1))
			return ScriptError(ERR_IFMSGBOX, LINE_RAW_ARG1);
		break;
	case ACT_IFIS:
	case ACT_IFISNOT:
		if (line->mArgc > 1 && !line->ArgHasDeref(2) && !line->ConvertVariableTypeName(LINE_RAW_ARG2))
			// Don't refer to it as "Parameter #2" because this command isn't formatted/displayed that way:
			return ScriptError("The type name must be either NUMBER, INTEGER, FLOAT, TIME, DATE, DIGIT, ALPHA, ALNUM, SPACE, or a variable reference."
				, LINE_RAW_ARG2);
		break;
	case ACT_GETKEYSTATE:
		if (line->mArgc > 1 && !line->ArgHasDeref(2) && !TextToVK(LINE_RAW_ARG2))
			return ScriptError("This is not a valid key or mouse button name.", LINE_RAW_ARG2);
		break;
	case ACT_DIV:
		if (!line->ArgHasDeref(2)) // i.e. if it's a deref, we won't try to validate it now.
			if (!_atoi64(LINE_RAW_ARG2))
				return ScriptError("This line would attempt to divide by zero.");
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
			if (   !(line->mAttribute = FindOrAddGroup(LINE_RAW_ARG1))   )
				return FAIL;  // The above already displayed the error.
		if (aActionType == ACT_GROUPACTIVATE || aActionType == ACT_GROUPDEACTIVATE)
		{
			if (line->mArgc > 1 && !line->ArgHasDeref(2) && *LINE_RAW_ARG2)
				if (strlen(LINE_RAW_ARG2) > 1 || toupper(*LINE_RAW_ARG2) != 'R')
					return ScriptError("Parameter #2 must be either blank, R, or a variable reference."
						, LINE_RAW_ARG2);
		}
		else if (aActionType == ACT_GROUPCLOSE)
			if (line->mArgc > 1 && !line->ArgHasDeref(2) && *LINE_RAW_ARG2)
				if (strlen(LINE_RAW_ARG2) > 1 || !strchr("RA", toupper(*LINE_RAW_ARG2)))
					return ScriptError("Parameter #2 must be either blank, R, A, or a variable reference."
						, LINE_RAW_ARG2);
		break;

	case ACT_RUN:
	case ACT_RUNWAIT:
		if (*LINE_RAW_ARG3 && !line->ArgHasDeref(3))
			if (line->ConvertRunMode(LINE_RAW_ARG3) == SW_SHOWNORMAL)
				return ScriptError(ERR_RUN_SHOW_MODE, LINE_RAW_ARG3);
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
			{
				if (IsPureNumeric(LINE_RAW_ARG1, false))
					line->mAttribute = ATTR_LOOP_NORMAL;
				else
					line->mAttribute = line->RegConvertRootKey(LINE_RAW_ARG1) ? ATTR_LOOP_REG : ATTR_LOOP_FILE;
			}
			break;
		default:  // has 2 or more args.
			line->mAttribute = line->RegConvertRootKey(LINE_RAW_ARG1) ? ATTR_LOOP_REG : ATTR_LOOP_FILE;
			if (line->mAttribute == ATTR_LOOP_FILE)
			{
				// Validate whatever we can rather than waiting for runtime validation:
				if (!line->ArgHasDeref(2) && Line::ConvertLoopMode(LINE_RAW_ARG2) == FILE_LOOP_INVALID)
					return ScriptError(ERR_LOOP_FILE_MODE, LINE_RAW_ARG2);
				if (line->mArgc > 2 && !line->ArgHasDeref(3) && *LINE_RAW_ARG3)
					if (strlen(LINE_RAW_ARG3) > 1 || (*LINE_RAW_ARG3 != '0' && *LINE_RAW_ARG3 != '1'))
						return ScriptError("Parameter #3 must be either blank, 0, 1, or a variable reference."
							, LINE_RAW_ARG3);
			}
			else // Registry loop.
			{
				if (line->mArgc > 2 && !line->ArgHasDeref(3) && Line::ConvertLoopMode(LINE_RAW_ARG3) == FILE_LOOP_INVALID)
					return ScriptError(ERR_LOOP_REG_MODE, LINE_RAW_ARG3);
				if (line->mArgc > 3 && !line->ArgHasDeref(4) && *LINE_RAW_ARG4)
					if (strlen(LINE_RAW_ARG4) > 1 || (*LINE_RAW_ARG4 != '0' && *LINE_RAW_ARG4 != '1'))
						return ScriptError("Parameter #4 must be either blank, 0, 1, or a variable reference."
							, LINE_RAW_ARG4);
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



Var *Line::ResolveVarOfArg(int aArgIndex, bool aCreateIfNecessary)
// Args that are input or output variables are normally resolved at load-time, so that
// they contain a pointer to their Var object.  This is done for performance.  However,
// in order to support dynamically resolved variables names like AutoIt2 (e.g. arrays),
// we have to do some extra work here at runtime.
// Callers specify false for aCreateIfNecessary whenever the contents of the variable
// they're trying to find is unimportant.  For example, dynamically built input variables,
// such as "StringLen, length, array%i%", do not need to be created if they weren't
// previously assigned to (i.e. they weren't previously used as an output variable).
// In the above example, array elements that don't exist are not created unless they were
// specifically assigned a value at some previous time during runtime.
{
	if (aArgIndex >= mArgc) // The requested ARG isn't even present, so it can't have a variable.
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
	// function or should check for NULL upon return.

	char var_name[MAX_VAR_NAME_LENGTH + 1];  // Will hold the dynamically built name.
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
			#define DYNAMIC_TOO_LONG "This dynamically built variable name is too long."
			LineError(DYNAMIC_TOO_LONG, WARN, mArg[aArgIndex].text);
			return NULL;
		}
		// Now copy the contents of the dereferenced var.  For all cases, aBuf has already
		// been verified to be large enough, assuming the value hasn't changed between the
		// time we were called and the time the caller calculated the space needed.
		if (deref->var->Get() > (VarSizeType)(MAX_VAR_NAME_LENGTH - vni)) // The variable name would be too long!
		{
			LineError(DYNAMIC_TOO_LONG, WARN, mArg[aArgIndex].text);
			return NULL;
		}
		vni += deref->var->Get(var_name + vni);
		// Finally, jump over the dereference text:
		pText += deref->length;
	}
	// Copy any chars that occur after the final deref into the buffer:
	for (; *pText && vni < MAX_VAR_NAME_LENGTH; var_name[vni++] = *pText++);
	if (vni >= MAX_VAR_NAME_LENGTH && *pText) // The variable name would be too long!
	{
		LineError(DYNAMIC_TOO_LONG, WARN, mArg[aArgIndex].text);
		return NULL;
	}
	// Terminate the buffer, even if nothing was written into it:
	var_name[vni] = '\0';

	static Var empty_var("");
	Var *found_var;
	if (aCreateIfNecessary)
	{
		found_var = g_script.FindOrAddVar(var_name);
		if (mArg[aArgIndex].type == ARG_TYPE_OUTPUT_VAR && VAR_IS_RESERVED(found_var))
		{
			LineError(ERR_VAR_IS_RESERVED, WARN, var_name);
			return NULL;  // Don't return the var, preventing the caller from assigning to it.
		}
		else
			return found_var;
	}
	else
	{
		// Now we've dynamically build the variable name.  It's possible that the name is illegal,
		// so check that (the name is automatically checked by FindOrAddVar(), so we only need to
		// check it if we're not calling that):
		if (!Var::ValidateName(var_name))
			return NULL; // Above already displayed error for us.
		found_var = g_script.FindVar(var_name);
		// If not found: for performance reasons, don't create it because caller just wants an empty variable.
		return found_var ? found_var : &empty_var;
	}
}



Var *Script::FindOrAddVar(char *aVarName, size_t aVarNameLength)
// Returns the Var whose name matches aVarName.  If it doesn't exist, it is created.
{
	if (!aVarName || !*aVarName) return NULL;
	Var *var = FindVar(aVarName, aVarNameLength);
	if (var)
		return var;
	// Otherwise, no match found, so create a new var.
	if (AddVar(aVarName, aVarNameLength) != OK)
		return NULL;
	return mLastVar;
}



Var *Script::FindVar(char *aVarName, size_t aVarNameLength)
// Returns the Var whose name matches aVarName.  If it doesn't exist, NULL is returned.
{
	if (!aVarName || !*aVarName) return NULL;
	if (!aVarNameLength) // Caller didn't specify, so use the entire string.
		aVarNameLength = strlen(aVarName);
	for (Var *var = mFirstVar; var != NULL; var = var->mNextVar)
		if (!strlicmp(aVarName, var->mName, (UINT)aVarNameLength)) // Match found.
			return var;
	// Otherwise, no match found:
	return NULL;
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

	else if (!stricmp(new_name, "a_LoopFileName")) var_type = VAR_LOOPFILENAME;
	else if (!stricmp(new_name, "a_LoopFileShortName")) var_type = VAR_LOOPFILESHORTNAME;
	else if (!stricmp(new_name, "a_LoopFileDir")) var_type = VAR_LOOPFILEDIR;
	else if (!stricmp(new_name, "a_LoopFileFullPath")) var_type = VAR_LOOPFILEFULLPATH;
	else if (!stricmp(new_name, "a_LoopFileTimeModified")) var_type = VAR_LOOPFILETIMEMODIFIED;
	else if (!stricmp(new_name, "a_LoopFileTimeCreated")) var_type = VAR_LOOPFILETIMECREATED;
	else if (!stricmp(new_name, "a_LoopFileTimeAccessed")) var_type = VAR_LOOPFILETIMEACCESSED;
	else if (!stricmp(new_name, "a_LoopFileAttrib")) var_type = VAR_LOOPFILEATTRIB;
	else if (!stricmp(new_name, "a_LoopFileSize")) var_type = VAR_LOOPFILESIZE;
	else if (!stricmp(new_name, "a_LoopFileSizeKB")) var_type = VAR_LOOPFILESIZEKB;
	else if (!stricmp(new_name, "a_LoopFileSizeMB")) var_type = VAR_LOOPFILESIZEMB;

	else if (!stricmp(new_name, "a_LoopRegType")) var_type = VAR_LOOPREGTYPE;
	else if (!stricmp(new_name, "a_LoopRegKey")) var_type = VAR_LOOPREGKEY;
	else if (!stricmp(new_name, "a_LoopRegSubKey")) var_type = VAR_LOOPREGSUBKEY;
	else if (!stricmp(new_name, "a_LoopRegName")) var_type = VAR_LOOPREGNAME;
	else if (!stricmp(new_name, "a_LoopRegTimeModified")) var_type = VAR_LOOPREGTIMEMODIFIED;

	else if (!stricmp(new_name, "a_ThisHotkey")) var_type = VAR_THISHOTKEY;
	else if (!stricmp(new_name, "a_PriorHotkey")) var_type = VAR_PRIORHOTKEY;
	else if (!stricmp(new_name, "a_TimeSinceThisHotkey")) var_type = VAR_TIMESINCETHISHOTKEY;
	else if (!stricmp(new_name, "a_TimeSincePriorHotkey")) var_type = VAR_TIMESINCEPRIORHOTKEY;
	else if (!stricmp(new_name, "a_TickCount")) var_type = VAR_TICKCOUNT;
	else if (!stricmp(new_name, "a_Space")) var_type = VAR_SPACE;
	else if (!stricmp(new_name, "a_Tab")) var_type = VAR_TAB;
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



Line *Script::PreparseIfElse(Line *aStartingLine, ExecUntilMode aMode, AttributeType aLoopType1, AttributeType aLoopType2)
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
	AttributeType loop_type1, loop_type2;  // Although rare, a statement can be enclosed in both a file-loop and a reg-loop.
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
			loop_type1 = ATTR_NONE;
			if (aLoopType1 == ATTR_LOOP_FILE || line->mAttribute == ATTR_LOOP_FILE)
				// i.e. if either one is a file-loop, that's enough to establish
				// the fact that we're in a file loop.
				loop_type1 = ATTR_LOOP_FILE;
			else if (aLoopType1 == ATTR_LOOP_UNKNOWN || line->mAttribute == ATTR_LOOP_UNKNOWN)
				// ATTR_LOOP_UNKNOWN takes precedence over ATTR_LOOP_NORMAL because
				// we can't be sure if we're in a file loop, but it's correct to
				// assume that we are (otherwise, unwarranted syntax errors may be reported
				// later on in here).
				loop_type1 = ATTR_LOOP_UNKNOWN;
			else if (aLoopType1 == ATTR_LOOP_NORMAL || line->mAttribute == ATTR_LOOP_NORMAL)
				loop_type1 = ATTR_LOOP_NORMAL;

			// The section is the same as above except for registry vs. file loops:
			loop_type2 = ATTR_NONE;
			if (aLoopType2 == ATTR_LOOP_REG || line->mAttribute == ATTR_LOOP_REG)
				loop_type2 = ATTR_LOOP_REG;
			else if (aLoopType2 == ATTR_LOOP_UNKNOWN || line->mAttribute == ATTR_LOOP_UNKNOWN)
				loop_type2 = ATTR_LOOP_UNKNOWN;
			else if (aLoopType2 == ATTR_LOOP_NORMAL || line->mAttribute == ATTR_LOOP_NORMAL)
				loop_type2 = ATTR_LOOP_NORMAL;

			// Check if the IF's action-line is something we want to recurse.  UPDATE: Always
			// recurse because other line types, such as Goto and Gosub, need to be preparsed
			// by this function even if they are the single-line actions of an IF or an ELSE:
			// Recurse this line rather than the next because we want
			// the called function to recurse again if this line is a ACT_BLOCK_BEGIN
			// or is itself an IF:
			line_temp = PreparseIfElse(line_temp, ONLY_ONE_LINE, loop_type1, loop_type2);
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
				line = PreparseIfElse(line, ONLY_ONE_LINE, aLoopType1, aLoopType2);
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
			line = PreparseIfElse(line->mNextLine, UNTIL_BLOCK_END, aLoopType1, aLoopType2);
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
			if (!aLoopType1 && !aLoopType2)
				return line->PreparseError("This break or continue statement is not enclosed by a loop.");
			break;

		case ACT_REGREAD:
		case ACT_REGWRITE: // Unknown loops get the benefit of the doubt.
			if (aLoopType2 != ATTR_LOOP_REG && aLoopType2 != ATTR_LOOP_UNKNOWN && line->mArgc < 4)
				return line->PreparseError("When not enclosed in a registry-loop, this command requires 4 parameters.");
			break;
		case ACT_REGDELETE:
			if (aLoopType2 != ATTR_LOOP_REG && aLoopType2 != ATTR_LOOP_UNKNOWN && line->mArgc < 2)
				return line->PreparseError("When not enclosed in a registry-loop, this command requires 2 parameters.");
			break;

		case ACT_FILEGETATTRIB:
		case ACT_FILESETATTRIB:
		case ACT_FILEGETSIZE:
		case ACT_FILEGETVERSION:
		case ACT_FILEGETTIME:
		case ACT_FILESETTIME:
			if (aLoopType1 != ATTR_LOOP_FILE && aLoopType1 != ATTR_LOOP_UNKNOWN && !*LINE_RAW_ARG2)
				return line->PreparseError("When not enclosed in a file-loop, this command requires a 2nd parameter.");
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

char *Line::sSourceFile[MAX_SCRIPT_FILES]; // No init needed.
int Line::nSourceFiles = 0;  // Zero source files initially.  The main script will be the first.

char *Line::sDerefBuf = NULL;  // Buffer to hold the values of any args that need to be dereferenced.
char *Line::sDerefBufMarker = NULL;
size_t Line::sDerefBufSize = 0;
char *Line::sArgDeref[MAX_ARGS]; // No init needed.


ResultType Line::ExecUntil(ExecUntilMode aMode, modLR_type aModifiersLR, Line **apJumpToLine
	, WIN32_FIND_DATA *aCurrentFile, RegItemStruct *aCurrentRegItem)
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
	int sleep_duration;
	bool do_sleep;
	DWORD time_since_last_sleep;

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
		{
			// Sleep in between batches of lines, like AutoIt, to reduce
			// the chance that a maxed CPU will interfere with time-critical
			// apps such as games, video capture, or video playback.  Also,
			// check the message queue to see if the program has been asked
			// to exit by some outside force, or whether the user pressed
			// another hotkey to interrupt this one.  In the case of exit,
			// this call will never return.  Note: MsgSleep() will reset
			// mLinesExecutedThisCycle for us:
			do_sleep = true;
			sleep_duration = INTERVAL_UNSPECIFIED;  // Exact length of sleep unimportant.
		}
		else
		{
			// DWORD math ensures the correct answer even if latest tickcount has wrapped around past zero:
			time_since_last_sleep = GetTickCount() - g_script.mLastSleepTime;
			// Testing reveals that a value of 5 vs. 15 produces smoother mouse cursor movement while a
			// tight script loop is running with the mouse hook installed.  It also reduces CPU
			// utilitization from 60% to around 50% on my system (not important -- what is important
			// is that even a slow CPU should have enough time to run hundreds/thousands/millions of
			// script lines in 5ms [which is really 10ms on most systems due to timer granularity]).
			// In other words, having this forced sleep in here reduces the script speed by half
			// of what it would be if it were allowed to max the CPU.  This 50% reduction seems a fair
			// trade since most people would not want the the keyboard and mouse to lag at all while
			// scripts are running at "infinite" speed.  In any case, this is only done if the hook(s)
			// are installed.  Non-hook scripts are allowed the leeway to not sleep at all:
			if (Hotkey::HookIsActive() && time_since_last_sleep > 5)
			{
				// Try to reduce keyboard & mouse lag by forcing the process into the GetMessage()
				// state regardless of how high the value of BatchLines is.  The GetMessage() state
				// is the only means by which key & mouse events are ever sent to the hooks:
				do_sleep = true;
				sleep_duration = INTERVAL_UNSPECIFIED;
			}
			else if (time_since_last_sleep > 200)
			{
				// Do a minimal message check at least a few times a second so that the program's
				// windows and controls, such as the tray icon menu and the main window, continue
				// to be responsive even if BatchLines is very high or infinite:
				do_sleep = true;
				sleep_duration = -1;  // No sleep, just check messages.
				// Force an update since MsgSleep() won't do it for -1 sleep time:
				g_script.mLastSleepTime = GetTickCount();
			}
			else
				do_sleep = false;
		}

		if (do_sleep)
			MsgSleep(sleep_duration);

		// At this point, a pause may have been triggered either by the above MsgSleep()
		// or due to the action of a command (e.g. Pause, or perhaps tray menu "pause" was selected during Sleep):
		for (;;)
		{
			if (g.IsPaused)
				MsgSleep(INTERVAL_UNSPECIFIED, RETURN_AFTER_MESSAGES, false);
			else
				break;
		}

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
				result = line->mNextLine->ExecUntil(ONLY_ONE_LINE, aModifiersLR, &jump_to_line, aCurrentFile, aCurrentRegItem);
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
					result = line->mNextLine->ExecUntil(ONLY_ONE_LINE, aModifiersLR, &jump_to_line, aCurrentFile, aCurrentRegItem);
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
			result = line->mRelatedLine->ExecUntil(UNTIL_RETURN, aModifiersLR, NULL, aCurrentFile, aCurrentRegItem);
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
				result = jump_to_line->ExecUntil(UNTIL_RETURN, aModifiersLR, NULL, aCurrentFile, aCurrentRegItem);
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
			HKEY reg_root_key = NULL;
			if (attr == ATTR_LOOP_REG)
				reg_root_key = RegConvertRootKey(LINE_ARG1);
			else if (attr == ATTR_LOOP_UNKNOWN || attr == ATTR_NONE)
				// Since it couldn't be determined at load-time (probably due to derefs),
				// determine whether it's a file-loop, registry-loop or a normal/counter loop.
				// But don't change the value of line->mAttribute because that's our
				// indicator of whether this needs to be evaluated every time for
				// this particular loop (since the nature of the loop can change if the
				// contents of the variables dereferenced for this line change during runtime):
			{
				switch (line->mArgc)
				{
				case 0:
					attr = ATTR_LOOP_NORMAL;
					break;
				case 1:
					// Unlike at loadtime, allow it to be negative at runtime in case it was a variable
					// reference that resolved to a negative number, to indicate that 0 iterations
					// should be performed:
					if (IsPureNumeric(LINE_ARG1, true))
						attr = ATTR_LOOP_NORMAL;
					else
					{
						reg_root_key = RegConvertRootKey(LINE_ARG1);
						attr = reg_root_key ? ATTR_LOOP_REG : ATTR_LOOP_FILE;
					}
					break;
				default: // 2 or more args.
					reg_root_key = RegConvertRootKey(LINE_ARG1);
					attr = reg_root_key ? ATTR_LOOP_REG : ATTR_LOOP_FILE;
				}
			}

			bool recurse_subfolders = (attr == ATTR_LOOP_FILE && *LINE_ARG3 == '1' && !*(LINE_ARG3 + 1))
				|| (attr == ATTR_LOOP_REG && *LINE_ARG4 == '1' && !*(LINE_ARG4 + 1));

			__int64 iteration_limit = 0;
			bool is_infinite = line->mArgc < 1;
			if (!is_infinite)
				// Must be set to zero for ATTR_LOOP_FILE:
				iteration_limit = (attr == ATTR_LOOP_FILE || attr == ATTR_LOOP_REG) ? 0 : _atoi64(LINE_ARG1);

			if (line->mActionType == ACT_REPEAT && !iteration_limit)
				is_infinite = true;  // Because a 0 means infinite in AutoIt2 for the REPEAT command.

			FileLoopModeType file_loop_mode = FILE_LOOP_INVALID;
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

			bool continue_main_loop = false; // Init prior to below call.
			jump_to_line = NULL; // Init prior to below call.
			if (attr == ATTR_LOOP_REG)
				result = line->PerformLoopReg(aModifiersLR, aCurrentFile, continue_main_loop, jump_to_line
					, file_loop_mode, recurse_subfolders, reg_root_key, LINE_ARG2);
			else // All other loops types are handled this way:
				result = line->PerformLoop(aModifiersLR, aCurrentFile, aCurrentRegItem, continue_main_loop, jump_to_line
					, attr, file_loop_mode, recurse_subfolders, LINE_ARG1, iteration_limit, is_infinite);
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
			result = line->mNextLine->ExecUntil(UNTIL_BLOCK_END, aModifiersLR, &jump_to_line, aCurrentFile, aCurrentRegItem);
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
			result = line->Perform(aModifiersLR, aCurrentFile, aCurrentRegItem);
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

	pure_numeric_type var_is_pure_numeric, value_is_pure_numeric;
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
		#define STRING_COMPARE (g.StringCaseSense ? strcmp(ARG1, ARG2) : stricmp(ARG1, ARG2))
		#undef DETERMINE_NUMERIC_TYPES
		#define DETERMINE_NUMERIC_TYPES \
			value_is_pure_numeric = IsPureNumeric(ARG2, true, false, true);\
			var_is_pure_numeric = IsPureNumeric(ARG1, true, false, true);
		#define IF_EITHER_IS_NON_NUMERIC if (!value_is_pure_numeric || !var_is_pure_numeric)
		#define IF_EITHER_IS_FLOAT if (value_is_pure_numeric == PURE_FLOAT || var_is_pure_numeric == PURE_FLOAT)

		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_NON_NUMERIC
			if_condition = !STRING_COMPARE;
		else IF_EITHER_IS_FLOAT  // It might perform better to only do float conversions & math when necessary.
			if_condition = atof(ARG1) == atof(ARG2);
		else
			if_condition = _atoi64(ARG1) == _atoi64(ARG2);

		break;
	case ACT_IFNOTEQUAL:
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_NON_NUMERIC
			if_condition = STRING_COMPARE;
		else IF_EITHER_IS_FLOAT  // It might perform better to only do float conversions & math when necessary.
			if_condition = atof(ARG1) != atof(ARG2);
		else
			if_condition = _atoi64(ARG1) != _atoi64(ARG2);
		break;
	case ACT_IFLESS:
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_NON_NUMERIC
			if_condition = STRING_COMPARE < 0;
		else IF_EITHER_IS_FLOAT  // It might perform better to only do float conversions & math when necessary.
			if_condition = atof(ARG1) < atof(ARG2);
		else
			if_condition = _atoi64(ARG1) < _atoi64(ARG2);
		break;
	case ACT_IFLESSOREQUAL:
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_NON_NUMERIC
			if_condition = STRING_COMPARE <= 0;
		else IF_EITHER_IS_FLOAT  // It might perform better to only do float conversions & math when necessary.
			if_condition = atof(ARG1) <= atof(ARG2);
		else
			if_condition = _atoi64(ARG1) <= _atoi64(ARG2);
		break;
	case ACT_IFGREATER:
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_NON_NUMERIC
			if_condition = STRING_COMPARE > 0;
		else IF_EITHER_IS_FLOAT  // It might perform better to only do float conversions & math when necessary.
			if_condition = atof(ARG1) > atof(ARG2);
		else
			if_condition = _atoi64(ARG1) > _atoi64(ARG2);
		break;
	case ACT_IFGREATEROREQUAL:
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_NON_NUMERIC
			if_condition = STRING_COMPARE >= 0;
		else IF_EITHER_IS_FLOAT  // It might perform better to only do float conversions & math when necessary.
			if_condition = atof(ARG1) >= atof(ARG2);
		else
			if_condition = _atoi64(ARG1) >= _atoi64(ARG2);
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
			FILETIME ft;
			// Also insist on numeric, because even though YYYYMMDDToFileTime() will properly convert a
			// non-conformant string such as "2004.4", for future compatibility, we don't want to
			// report that such strings are valid times:
			if_condition = IsPureNumeric(ARG1, false, false, false) && YYYYMMDDToFileTime(ARG1, &ft);
			break;
		}
		case VAR_TYPE_DIGIT:
			if_condition = true;
			for (cp = ARG1; *cp; ++cp)
				if (!isdigit(*cp))
					if_condition = false;
			break;
		case VAR_TYPE_ALPHA:
			// Like AutoIt3, the empty string is considered to be alphabetic, which is only slightly debatable.
			if_condition = true;
			for (cp = ARG1; *cp; ++cp)
				if (!IsCharAlpha(*cp)) // Use this for to better support chars from non-English languages.
					if_condition = false;
			break;
		case VAR_TYPE_ALNUM:
			// Like AutoIt3, the empty string is considered to be alphabetic, which is only slightly debatable.
			if_condition = true;
			for (cp = ARG1; *cp; ++cp)
				if (!IsCharAlphaNumeric(*cp)) // Use this for to better support chars from non-English languages.
					if_condition = false;
			break;
		case VAR_TYPE_SPACE:
			if_condition = true;
			for (cp = ARG1; *cp; ++cp)
				if (!isspace(*cp))
					if_condition = false;
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
		return LineError("EvaluateCondition(): Unhandled windowing action type." PLEASE_REPORT ERR_ABORT);
	}
	return if_condition ? CONDITION_TRUE : CONDITION_FALSE;
}



ResultType Line::PerformLoop(modLR_type aModifiersLR, WIN32_FIND_DATA *apCurrentFile
	, RegItemStruct *apCurrentRegItem, bool &aContinueMainLoop, Line *&aJumpToLine
	, AttributeType aAttr, FileLoopModeType aFileLoopMode, bool aRecurseSubfolders, char *aFilePattern
	, __int64 aIterationLimit, bool aIsInfinite)
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
	for (__int64 i = 0; aIsInfinite || file_found || i < aIterationLimit; ++i)
	{
		// Execute once the body of the loop (either just one statement or a block of statements).
		// Preparser has ensured that every LOOP has a non-NULL next line.
		// apCurrentFile is sent as an arg so that more than one nested/recursive
		// file-loop can be supported:
		result = mNextLine->ExecUntil(ONLY_ONE_LINE, aModifiersLR, &jump_to_line
			, file_found ? &new_current_file : apCurrentFile // inner loop's file takes precedence over outer's.
			, apCurrentRegItem);
		if (result == FAIL || result == EARLY_RETURN || result == EARLY_EXIT || result == LOOP_BREAK)
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
		result = PerformLoop(aModifiersLR, NULL, apCurrentRegItem, aContinueMainLoop, aJumpToLine
			, aAttr, aFileLoopMode, aRecurseSubfolders, file_path
			, aIterationLimit, aIsInfinite);
		// result should never be LOOP_CONTINUE because the above call to PerformLoop() should have
		// handled that case.  However, it can be LOOP_BREAK if it encoutered the break command.
		if (result == FAIL || result == EARLY_RETURN || result == EARLY_EXIT || result == LOOP_BREAK)
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



ResultType Line::PerformLoopReg(modLR_type aModifiersLR, WIN32_FIND_DATA *apCurrentFile
	, bool &aContinueMainLoop, Line *&aJumpToLine, FileLoopModeType aFileLoopMode
	, bool aRecurseSubfolders, HKEY aRootKey, char *aRegSubkey)
{
	RegItemStruct reg_item(aRootKey, aRegSubkey);
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
		result = mNextLine->ExecUntil(ONLY_ONE_LINE, aModifiersLR, &jump_to_line, apCurrentFile, &reg_item);\
		if (result == FAIL || result == EARLY_RETURN || result == EARLY_EXIT || result == LOOP_BREAK)\
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
				// after the recusive call returns to us:
				snprintf(subkey_full_path, sizeof(subkey_full_path), "%s\\%s", reg_item.subkey, reg_item.name);
				// This section is very similar to the one in PerformLoop(), so see it for comments:
				result = PerformLoopReg(aModifiersLR, apCurrentFile, aContinueMainLoop, aJumpToLine
					, aFileLoopMode, aRecurseSubfolders, aRootKey, subkey_full_path);
				if (result == FAIL || result == EARLY_RETURN || result == EARLY_EXIT || result == LOOP_BREAK)
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



inline ResultType Line::Perform(modLR_type aModifiersLR, WIN32_FIND_DATA *aCurrentFile, RegItemStruct *aCurrentRegItem)
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
	pure_numeric_type var_is_pure_numeric, value_is_pure_numeric; // For math operations.
	vk_type vk; // For mouse commands and GetKeyState.
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
		int wait_time = *ARG3 ? (int)(1000 * atof(ARG3)) : DEFAULT_WINCLOSE_WAIT;
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
		if (mArgc < 2 && aCurrentRegItem) // Uses the registry loop's current item.
			// If aCurrentRegItem->name specifies a subkey rather than a value name, do this anyway
			// so that it will set ErrorLevel to ERROR and set the output variable to be blank:
			return RegRead(aCurrentRegItem->root_key, aCurrentRegItem->subkey, aCurrentRegItem->name);
		else if (mArgc < 5) // The new 4-parameter mode.
			return RegRead(RegConvertRootKey(ARG2), ARG3, ARG4);
		else // In 5-parameter mode, Arg2 is unused; it's only for backward compatibility with AutoIt2.
			return RegRead(RegConvertRootKey(ARG3), ARG4, ARG5);
	case ACT_REGWRITE:
		if (mArgc < 2 && aCurrentRegItem) // Uses the registry loop's current item.
			// If aCurrentRegItem->name specifies a subkey rather than a value name, do this anyway
			// so that it will set ErrorLevel to ERROR.  An error will also be indicated if
			// aCurrentRegItem->type is an unsupported type:
			return RegWrite(aCurrentRegItem->type, aCurrentRegItem->root_key, aCurrentRegItem->subkey, aCurrentRegItem->name, ARG1);
		else
			return RegWrite(RegConvertValueType(ARG1), RegConvertRootKey(ARG2), ARG3, ARG4, ARG5);
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
		else
			return RegDelete(RegConvertRootKey(ARG1), ARG2, ARG3);

	case ACT_SHUTDOWN:
		return Util_Shutdown(atoi(ARG1)) ? OK : FAIL;
	case ACT_SLEEP:
		// Only support 32-bit values for this command, since it seems unlikely anyone would to have
		// it sleep more than 24.8 days or so.  It also helps performance on 32-bit hardware because
		// MsgSleep() is so heavily called and checks the value of the first parameter frequently:
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
			sleep_duration = (int)(atof(mActionType == ACT_CLIPWAIT ? ARG1 : ARG3) * 1000); // Can be zero.
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
				if (exit_code != STATUS_PENDING)
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
					g_ErrorLevel->Assign((int)exit_code);
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
	case ACT_CONTROLCLICK:
		if (*ARG4)
		{
			if (   !(vk = ConvertMouseButton(ARG4))   )
				return LineError(ERR_MOUSE_BUTTON ERR_ABORT, FAIL, ARG4);
		}
		else // Default button when the param is blank or an reference to an empty var.
			vk = VK_LBUTTON;
		return ControlClick(vk, *ARG5 ? atoi(ARG5) : 1, *ARG6, ARG1, ARG2, ARG3, ARG7, ARG8);
	case ACT_CONTROLGETFOCUS:
		return ControlGetFocus(ARG2, ARG3, ARG4, ARG5);
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

	case ACT_STRINGLEFT:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		chars_to_extract = atoi(ARG3); // Use 32-bit signed to detect negatives and fit it VarSizeType.
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
		chars_to_extract = atoi(ARG3); // Use 32-bit signed to detect negatives and fit it VarSizeType.
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
		start_char_num = atoi(ARG3);
		if (start_char_num <= 0)
			// It's somewhat debatable, but it seems best not to report an error in this and
			// other cases.  The result here is probably enough to speak for itself, for script
			// debugging purposes:
			start_char_num = 1; // 1 is the position of the first char, unlike StringGetPos.
		chars_to_extract = atoi(ARG4); // Use 32-bit signed to detect negatives and fit it VarSizeType.
		if (chars_to_extract < 0)
			chars_to_extract = 0;
		// Assign() is capable of doing what we want in this case.
		// It will display any error that occurs:
		if ((int)strlen(ARG2) < start_char_num)
			return output_var->Assign();  // Set it to be blank in this case.
		else
			return output_var->Assign(ARG2 + start_char_num - 1, chars_to_extract);
	case ACT_STRINGTRIMLEFT:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		chars_to_extract = atoi(ARG3); // Use 32-bit signed to detect negatives and fit it VarSizeType.
		if (chars_to_extract < 0)
			chars_to_extract = 0;
		source_length = strlen(ARG2);
		if (chars_to_extract > (int)source_length) // This could be intentional, so don't display an error.
			chars_to_extract = (int)source_length;
		return output_var->Assign(ARG2 + chars_to_extract, (VarSizeType)(source_length - chars_to_extract));
	case ACT_STRINGTRIMRIGHT:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		chars_to_extract = atoi(ARG3); // Use 32-bit signed to detect negatives and fit it VarSizeType.
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
		if (mActionType == ACT_STRINGLOWER)
			strlwr(output_var->Contents());
		else
			strupr(output_var->Contents());
		return output_var->Close();  // In case it's the clipboard.
	case ACT_STRINGLEN:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		return output_var->Assign((__int64)strlen(ARG2)); // It already displayed any error.
	case ACT_STRINGGETPOS:
	{
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		g_ErrorLevel->Assign(ERRORLEVEL_NONE);  // Set default.
		bool search_from_the_right = (toupper(*ARG4) == 'R' || *ARG4 == '1');
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
		return output_var->Assign(pos); // It already displayed any error that may have occurred.
	}
	case ACT_STRINGREPLACE:
	{
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		source_length = strlen(ARG2);
		space_needed = (VarSizeType)source_length + 1;  // Set default, or starting value for accumulation.
		VarSizeType final_space_needed = space_needed;
		bool do_replace = *ARG2 && *ARG3; // i.e. don't allow replacement of the empty string.
		bool replace_all = toupper(*ARG5) == 'A' || *ARG5 == '1'; // i.e. more lenient if it's a dereferenced var.
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
				StrReplaceAll(output_var->Contents(), ARG3, ARG4, g.StringCaseSense);
			else
				StrReplace(output_var->Contents(), ARG3, ARG4, g.StringCaseSense);

		// UPDATE: This is NOT how AutoIt2 behaves, so don't do it:
		//if (g_script.mIsAutoIt2)
		//{
		//	trim(output_var->Contents());  // Since this is how AutoIt2 behaves.
		//	output_var->Length() = (VarSizeType)strlen(output_var->Contents());
		//}

		// Consider the above to have been always successful unless the below returns an error:
		return output_var->Close();  // In case it's the clipboard.
	}

	case ACT_GETKEYSTATE:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		if (vk = TextToVK(ARG2))
		{
			switch (toupper(*ARG3))
			{
			case 'T': // Whether a toggleable key such as CapsLock is currently turned on.
				return output_var->Assign((GetKeyState(vk) & 0x00000001) ? "D" : "U");
			case 'P': // Physical state of key.
				if (g_hhkLowLevelKeybd)
					// Since the hook is installed, use its value rather than that from
					// GetAsyncKeyState(), which doesn't seem to work as advertised, at
					// least under WinXP:
					return output_var->Assign(g_PhysicalKeyState[vk] ? "D" : "U");
				else
					return output_var->Assign(IsPhysicallyDown(vk) ? "D" : "U");
			default: // Logical state of key
				return output_var->Assign((GetKeyState(vk) & 0x8000) ? "D" : "U");
			}
		}
		return output_var->Assign("");

	case ACT_RANDOM:
	{
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		bool use_float = IsPureNumeric(ARG2, true, false, true) == PURE_FLOAT
			|| IsPureNumeric(ARG3, true, false, true) == PURE_FLOAT;
		if (use_float)
		{
			double rand_min = *ARG2 ? atof(ARG2) : 0;
			double rand_max = *ARG3 ? atof(ARG3) : INT_MAX;
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
	case ACT_SOUNDPLAY:
		return SoundPlay(ARG1, *ARG2 && !stricmp(ARG2, "wait") || !stricmp(ARG2, "1"));
	case ACT_FILEAPPEND:
		return this->FileAppend(ARG2, ARG1);  // To avoid ambiguity in case there's another FileAppend().
	case ACT_FILEREADLINE:
		return FileReadLine(ARG2, ARG3);
	case ACT_FILEDELETE:
		return FileDelete(ARG1);
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

	case ACT_FILESELECTFILE:
		return FileSelectFile(ARG2, ARG3, ARG4, ARG5);
	case ACT_FILESELECTFOLDER:
		return FileSelectFolder(ARG2, *ARG3 ? (atoi(ARG3) != 0) : true, ARG4);
	case ACT_FILECREATESHORTCUT:
		return FileCreateShortcut(SEVEN_ARGS);

	// Like AutoIt2, if either output_var or ARG1 aren't purely numeric, they
	// will be considered to be zero for all of the below math functions:
	case ACT_ADD:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;

		#undef DETERMINE_NUMERIC_TYPES
		#define DETERMINE_NUMERIC_TYPES \
			value_is_pure_numeric = IsPureNumeric(ARG2, true, false, true, true);\
			var_is_pure_numeric = IsPureNumeric(output_var->Contents(), true, false, true, true);
//		#define IF_EITHER_IS_NON_NUMERIC if (!value_is_pure_numeric || !var_is_pure_numeric)
		#define IF_EITHER_IS_FLOAT if (value_is_pure_numeric == PURE_FLOAT || var_is_pure_numeric == PURE_FLOAT)

		DETERMINE_NUMERIC_TYPES

		if (*ARG3 && strchr("SMHD", toupper(*ARG3))) // the command is being used to add a value to a date-time.
		{
			if (!value_is_pure_numeric) // It's considered to be zero, so the output_var is left unchanged:
				return OK;
			else
			{
				// Use double to support a floating point value for days, hours, minutes, etc:
				double nUnits = atof(ARG2);  // atof() returns a double, at least on MSVC++ 7.x
				FILETIME ft, ftNowUTC;
				if (*output_var->Contents())
				{
					if (!YYYYMMDDToFileTime(output_var->Contents(), &ft))
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
				FileTimeToYYYYMMDD(buf_temp, &ft, false);
				return output_var->Assign(buf_temp);
			}
		}
		else // The command is being used to do normal math (not date-time).
		{
			IF_EITHER_IS_FLOAT
				return output_var->Assign(atof(output_var->Contents()) + atof(ARG2));  // Overload: Assigns a double.
			else // Non-numeric variables or values are considered to be zero for the purpose of the calculation.
				return output_var->Assign(_atoi64(output_var->Contents()) + _atoi64(ARG2));  // Overload: Assigns an int.
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
				return output_var->Assign(atof(output_var->Contents()) - atof(ARG2));  // Overload: Assigns a double.
			else // Non-numeric variables or values are considered to be zero for the purpose of the calculation.
				return output_var->Assign(_atoi64(output_var->Contents()) - _atoi64(ARG2));  // Overload: Assigns an INT.
			break;
		}

		// If above didn't return, buf_temp now has the value to store:
		return output_var->Assign(buf_temp);

	case ACT_MULT:
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_FLOAT
			return output_var->Assign(atof(output_var->Contents()) * atof(ARG2));  // Overload: Assigns a double.
		else // Non-numeric variables or values are considered to be zero for the purpose of the calculation.
			return output_var->Assign(_atoi64(output_var->Contents()) * _atoi64(ARG2));  // Overload: Assigns an INT.

	case ACT_DIV:
	{
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;
		DETERMINE_NUMERIC_TYPES
		IF_EITHER_IS_FLOAT
		{
			double ARG2_as_float = atof(ARG2);  // Since atof() returns double, at least on MSVC++ 7.x
			// It's a little iffy to compare floats this way, but what's the alternative?:
			if (ARG2_as_float == (double)0.0)
				return LineError("This line would attempt to divide by zero." ERR_ABORT, FAIL, ARG2);
			return output_var->Assign(atof(output_var->Contents()) / ARG2_as_float);  // Overload: Assigns a double.
		}
		else // Non-numeric variables or values are considered to be zero for the purpose of the calculation.
		{
			__int64 ARG2_as_int = _atoi64(ARG2);
			if (!ARG2_as_int)
				return LineError("This line would attempt to divide by zero." ERR_ABORT, FAIL, ARG2);
			return output_var->Assign(_atoi64(output_var->Contents()) / ARG2_as_int);  // Overload: Assigns an INT.
		}
	}

	case ACT_KEYHISTORY:
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
		// Otherwise:
		return ShowMainWindow(MAIN_MODE_KEYHISTORY);
	case ACT_LISTLINES:
		return ShowMainWindow(MAIN_MODE_LINES);
	case ACT_LISTVARS:
		return ShowMainWindow(MAIN_MODE_VARS);
	case ACT_LISTHOTKEYS:
		return ShowMainWindow(MAIN_MODE_HOTKEYS);
	case ACT_MSGBOX:
	{
		int result;
		// If the MsgBox window can't be displayed for any reason, always return FAIL to
		// the caller because it would be unsafe to proceed with the execution of the
		// current script subroutine.  For example, if the script contains an IfMsgBox after,
		// this line, it's result would be unpredictable and might cause the subroutine to perform
		// the opposite action from what was intended (e.g. Delete vs. don't delete a file).
		result = (mArgc == 1) ? MsgBox(ARG1) : MsgBox(ARG3, atoi(ARG1), ARG2, atof(ARG4));
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
		return InputBox(output_var, ARG2, ARG3, toupper(*ARG4) == 'H'); // Last = whether to hide input.
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
		g_hWndSplash = CreateWindowEx(WS_EX_TOPMOST, WINDOW_CLASS_SPLASH, ARG3  // ARG3 is the window title
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
	case ACT_SETFORMAT:
	{
		// For now, it doesn't seem necessary to have runtime validation of the first parameter.
		// Just ignore the command if it's not valid:
		if (stricmp(ARG1, "float"))
			return OK;
		// -2 to allow room for the letter 'f' and the '%' that will be added:
		if (strlen(ARG2) >= sizeof(g.FormatFloat) - 2) // A variable that resolved to something too long.
			return OK; // Seems best not to bother with a runtime error for something so rare.
		// Make sure the formatted string wouldn't exceed the buffer size:
		__int64 width = _atoi64(ARG2);
		char *dot_pos = strchr(ARG2, '.');
		__int64 precision = dot_pos ? _atoi64(dot_pos + 1) : 0;
		if (width + precision + 2 > MAX_FORMATTED_NUMBER_LENGTH) // +2 to allow room for decimal point itself and a safety margin.
			return OK; // Don't change it.
		// Create as "%ARG2f".  Add a dot if none was specified so that "0" is the same as "0.", which
		// seems like the most user-friendly approach; it's also easier to document in the help file.
		// Note that %f can handle doubles in MSVC++:
		sprintf(g.FormatFloat, "%%%s%sf", ARG2, dot_pos ? "" : ".");
		return OK;
	}
	case ACT_SETCONTROLDELAY: g.ControlDelay = atoi(ARG1); return OK;
	case ACT_SETWINDELAY: g.WinDelay = atoi(ARG1); return OK;
	case ACT_SETKEYDELAY: g.KeyDelay = atoi(ARG1); return OK;
	case ACT_SETMOUSEDELAY: g.MouseDelay = atoi(ARG1); return OK;
	case ACT_SETBATCHLINES:
		// This value is signed 64-bits to support variable reference (i.e. containing a large int)
		// the user might throw at it:
		if (   !(g.LinesPerCycle = _atoi64(ARG1))   )
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
	case ACT_BLOCKINPUT:
		if (   (toggle = ConvertOnOff(ARG1, NEUTRAL)) != NEUTRAL   )
			BlockInput(toggle == TOGGLED_ON);
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
		if (mArg[iArg].type == ARG_TYPE_OUTPUT_VAR)  // Don't bother wasting the mem to deref output var.
		{
			// In case its "dereferenced" contents are ever directly examined, set it to be
			// the empty string.  This also allows the ARG to be passed a dummy param, which
			// makes things more convenient and maintainable in other places:
			sArgDeref[iArg] = "";
			continue;
		}

		the_only_var_of_this_arg = (mArg[iArg].type == ARG_TYPE_INPUT_VAR) ? ResolveVarOfArg(iArg, false) : NULL;
		if (!the_only_var_of_this_arg) // Arg isn't an input var.
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
			if (ArgMustBeDereferenced(the_only_var_of_this_arg, iArg))
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
		else // The arg must be expanded in the normal, lower-performance way.
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
				// It seems best to always allow floating point, even for commands that expect
				// an integer.  This is because the user may have done some math on a variable
				// such as SleepTime, such as dividing it, and thereby winding up with a float.
				// Rather than forcing the user to truncate SleepTime into an Int, it seems
				// best just to let atoi() convert the float into an int in these cases.
				// Note: The above only applies here, at runtime.  Incorrect literal floats are
				// still flagged as load-time errors:
				//allow_float = ArgAllowsFloat(*np);
				allow_negative = ArgAllowsNegative(*np);
				if (!IsPureNumeric(sArgDeref[*np - 1], allow_negative, true, true))
				{

					switch(mActionType)
					{
					case ACT_ADD:
					case ACT_SUB:
					case ACT_MULT:
					case ACT_DIV:
						// Don't report runtime errors for these (only loadtime) because they
						// indicate failure in a quieter, different way:
						break;
					case ACT_WINMOVE:
						if (stricmp(sArgDeref[*np - 1], "default"))
							return LineError("This parameter of this line doesn't resolve to either a"
								" numeric value or the word Default as required.", FAIL, sArgDeref[*np - 1]);
						// else don't attempt to set the deref to be blank, to make parsing simpler,
						// because sArgDeref[*np - 1] might point directly to the contents of
						// a variable and we don't want to modify it in that case.
						break;
					default:
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
		if (mArg[iArg].type == ARG_TYPE_OUTPUT_VAR)  // These should never be included in the space calculation.
			continue;
		// Always do this check before attempting to traverse the list of dereferences, since
		// such an attempt would be invalid in this case:
		the_only_var_of_this_arg = (mArg[iArg].type == ARG_TYPE_INPUT_VAR) ? ResolveVarOfArg(iArg, false) : NULL;
		if (!the_only_var_of_this_arg) // It's not an input var.
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
			if (!aCalcDerefBufSize || ArgMustBeDereferenced(the_only_var_of_this_arg, iArg))
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
	if (mArg[aArgIndex].type == ARG_TYPE_OUTPUT_VAR)
		// Just a warning because this function isn't set up to cause a true failure:
		LineError("ExpandArg() was called to expand an arg that contains only an output variable."
			PLEASE_REPORT, WARN);
	// Always do this part before attempting to traverse the list of dereferences, since
	// such an attempt would be invalid in this case:
	Var *the_only_var_of_this_arg = (mArg[aArgIndex].type == ARG_TYPE_INPUT_VAR) ? ResolveVarOfArg(aArgIndex, false) : NULL;
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
	snprintf(aBuf, BUF_SPACE_REMAINING, "%03u: ", mLineNumber);
	aBuf += strlen(aBuf);
	if (ACT_IS_ASSIGN(mActionType) || (ACT_IS_IF(mActionType) && mActionType < ACT_FIRST_COMMAND))
	{
		// Only these commands need custom conversion.
		snprintf(aBuf, BUF_SPACE_REMAINING, "%s%s %s %s"
			, ACT_IS_IF(mActionType) ? "IF " : ""
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



ResultType Line::BlockInput(bool aEnable)
// Adapted from the AutoIt3 source.
// Always returns OK for caller convenience.
{
	// Must be running 2000/ win98 for this function to be successful
	// We must dynamically load the function to retain compatibility with Win95
    // Get a handle to the DLL module that contains BlockInput
	HINSTANCE hinstLib = LoadLibrary("user32.dll");
    // If the handle is valid, try to get the function address.
	if (hinstLib != NULL)
	{
		typedef void (CALLBACK *BlockInput)(BOOL);
		BlockInput lpfnDLLProc = (BlockInput)GetProcAddress(hinstLib, "BlockInput");
		if (lpfnDLLProc != NULL)
			(*lpfnDLLProc)(aEnable ? TRUE : FALSE);
		// Free the DLL module.
		FreeLibrary(hinstLib);
	}
	return OK;
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
	ShowInEditor();
	MsgBox(buf);
	return FAIL; // See above for why it's better to return FAIL than CRITICAL_ERROR.
}



void Script::ShowInEditor()
{
	// IN LIGHT OF the addition of #include, this function will probably need to be revised so that
	// it targets the correct file that is the source of the error.
	// Disabled for now:
	return;

#ifdef AUTOHOTKEYSC
	return;
#endif

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
	snprintf(buf, sizeof(buf), "%d{ENTER}", mCurrLine ? mCurrLine->mLineNumber : mCurrLineNumber);
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
	char LRtext[128];
	snprintf(aBuf, aBufSize,
		"Window: %s"
		//"\r\nBlocks: %u"
		"\r\nKeybd hook: %s"
		"\r\nMouse hook: %s"
		"\r\nLast hotkey type: %s"
		"\r\nInterrupted threads: %d%s"
		"\r\nPaused threads: %d"
		"\r\nModifiers (GetKeyState() now) = %s"
		"\r\n"
		, win_title
		//, SimpleHeap::GetBlockCount()
		, g_hhkLowLevelKeybd == NULL ? "no" : "yes"
		, g_hhkLowLevelMouse == NULL ? "no" : "yes"
		, g_LastPerformedHotkeyType == HK_KEYBD_HOOK ? "keybd hook" : "not keybd hook"
		, g_nInterruptedSubroutines
		, g_nInterruptedSubroutines ? " (preempted: they will resume when the current thread finishes)" : ""
		, g_nPausedSubroutines
		, ModifiersLRToText(GetModifierLRStateSimple(), LRtext));
	aBuf += strlen(aBuf);
	GetHookStatus(aBuf, BUF_SPACE_REMAINING);
	aBuf += strlen(aBuf);
	snprintf(aBuf, BUF_SPACE_REMAINING, "\r\nPress [F5] to refresh.");
	aBuf += strlen(aBuf);
	return aBuf;
}



ResultType Script::ActionExec(char *aAction, char *aParams, char *aWorkingDir, bool aDisplayErrors
	, char *aRunShowMode, HANDLE *aProcess)
// Caller should specify NULL for aParams if it wants us to attempt to parse out params from
// within aAction.  Caller may specify empty string ("") instead to specify no params at all.
// Remember that aAction and aParams can both be NULL, so don't dereference without checking first.
// Note: For the Run & RunWait commands, aParams should always be NULL.  Params are parsed out of
// the aActionString at runtime, here, rather than at load-time because Run & RunWait might contain
// deferenced variable(s), which can only be resolved at runtime.
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
			if (   !(action_extension = stristr(parse_buf, ".exe "))   )
				if (   !(action_extension = stristr(parse_buf, ".exe\""))   )
					if (   !(action_extension = stristr(parse_buf, ".bat "))   )
						if (   !(action_extension = stristr(parse_buf, ".bat\""))   )
							if (   !(action_extension = stristr(parse_buf, ".com "))   )
								if (   !(action_extension = stristr(parse_buf, ".com\""))   )
									// Not 100% sure that .cmd and .hta are genuine executables in every sense:
									if (   !(action_extension = stristr(parse_buf, ".cmd "))   )
										if (   !(action_extension = stristr(parse_buf, ".cmd\""))   )
											if (   !(action_extension = stristr(parse_buf, ".hta "))   )
												action_extension = stristr(parse_buf, ".hta\"");

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
			new_process = pi.hProcess;
		}
//else
//MsgBox(command_line);
	}

	if (!success) // Either the above wasn't attempted, or the attempt failed.  So try ShellExecute().
	{
		SHELLEXECUTEINFO sei = {0};
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
			sei.lpVerb = "open";
			sei.lpFile = shell_action;
			sei.lpParameters = shell_params;
		}
		if (ShellExecuteEx(&sei) && LOBYTE(LOWORD(sei.hInstApp)) > 32) // Relies on short-circuit boolean order.
		{
			new_process = sei.hProcess;
			success = true;
		}
	}

	if (!success) // The above attempt(s) to launch failed.
	{
		if (aDisplayErrors)
		{
			char error_text[2048], verb_text[128];
			if (shell_action_is_system_verb && stricmp(shell_action, "open"))  // It's a verb, but not the default "open" verb.
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
			ScriptError(error_text);
		}
		return FAIL;
	}

	if (aProcess) // The caller wanted the process handle, so provide it.
		*aProcess = new_process;
	return OK;
}
