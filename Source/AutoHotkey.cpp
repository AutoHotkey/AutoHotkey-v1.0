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
#include "globaldata.h" // for access to many global vars
#include "application.h" // for MsgSleep()
#include "window.h" // For MsgBox() & SetForegroundLockTimeout()

// General note:
// The use of Sleep() should be avoided *anywhere* in the code.  Instead, call MsgSleep().
// The reason for this is that if the keyboard or mouse hook is installed, a straight call
// to Sleep() will cause user keystrokes & mouse events to lag because the message pump
// (GetMessage() or PeekMessage()) is the only means by which events are ever sent to the
// hook functions.


int WINAPI WinMain (HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
	// Init any globals not in "struct g" that need it:
	g_hInstance = hInstance;
	if (!GetCurrentDirectory(sizeof(g_WorkingDir), g_WorkingDir)) // Needed for the FileSelectFile() workaround.
		*g_WorkingDir = '\0';
	g_WorkingDirOrig = SimpleHeap::Malloc(g_WorkingDir); // Needed by the Reload command.

	// Set defaults, to be overridden by command line args we receive:
	bool restart_mode = false;

#ifdef AUTOHOTKEYSC
	char *script_filespec = __argv[0];  // i.e. the EXE name.  This is just a placeholder for now.
#else
	#ifdef _DEBUG
	//char *script_filespec = "C:\\Util\\AutoHotkey.ahk";
	char *script_filespec = "C:\\A-Source\\AutoHotkey\\ZZZZ Test Script.ahk";
	//char *script_filespec = "C:\\A-Source\\AutoHotkey\\Test\\Transform.ahk";
	#else
	char *script_filespec = NAME_P ".ini";  // Use this extension for better file association with editor(s).
	#endif
#endif

	// Examine command line args.  Rules:
	// Any special flags (e.g. /force and /restart) must appear prior to the script filespec.
	// The script filespec (if present) must be the first non-backslash arg.
	// All args that appear after the filespec are considered to be parameters for the script
	// and will be added as variables %1% %2% etc.
	// The above rules effectively make it impossible to autostart AutoHotkey.ini with parameters
	// unless the filename is explicitly given (shouldn't be an issue for 99.9% of people).
	char var_name[32]; // Small size since only numbers will be used (e.g. %1%, %2%).
	Var *var;
	bool switch_processing_is_complete = false;
	int script_param_num = 1;
	for (int i = 1; i < __argc; ++i) // Start at 1 because 0 contains the program name.
	{
		if (switch_processing_is_complete) // All args are now considered to be input parameters for the script.
		{
			snprintf(var_name, sizeof(var_name), "%d", script_param_num);
			if (   !(var = g_script.FindOrAddVar(var_name))   )
				return CRITICAL_ERROR;  // Realistically should never happen.
			var->Assign(__argv[i]);
			++script_param_num;
		}
		// Insist that switches be an exact match for the allowed values to cut down on ambiguity.
		// For example, if the user runs "CompiledScript.exe /find", we want /find to be considered
		// an input parameter for the script rather than a switch:
		else if (   (__argv[i][0] == '/' && toupper(__argv[i][1]) == 'R') || !stricmp(__argv[i], "/restart")   )
			restart_mode = true;
		else if (   (__argv[i][0] == '/' && toupper(__argv[i][1]) == 'F') || !stricmp(__argv[i], "/force")   )
			// Force the keybd/mouse hook(s) to be installed again even if another instance already did.
			g_ForceLaunch = true;
		else // since this is not a recognized switch, the end of the [Switches] section has been reached (by design).
		{
			switch_processing_is_complete = true;  // No more switches allowed after this point.
#ifdef AUTOHOTKEYSC
			--i; // Make the loop process this item again so that it will be treated as a script param.
#else
			script_filespec = __argv[i];  // The first unrecognized switch must be the script filespec, by design.
#endif
		}
	}

	// Like AutoIt2, store the number of script parameters in the script variable %0%, even if it's zero:
	if (   !(var = g_script.FindOrAddVar("0"))   )
		return CRITICAL_ERROR;  // Realistically should never happen.
	var->Assign(script_param_num - 1);

#ifndef AUTOHOTKEYSC
	size_t filespec_length = strlen(script_filespec);
	if (filespec_length >= CONVERSION_FLAG_LENGTH)
	{
		char *cp = script_filespec + filespec_length - CONVERSION_FLAG_LENGTH;
		// Now cp points to the first dot in the CONVERSION_FLAG of script_filespec (if it has one).
		if (!stricmp(cp, CONVERSION_FLAG))
			return Line::ConvertEscapeChar(script_filespec, '\\', '`', true);
	}
#endif

	global_init(&g);  // Set defaults prior to the below, since below might override them for AutoIt2 scripts.
	if (g_script.Init(script_filespec, restart_mode) != OK)  // Set up the basics of the script, using the above.
		return CRITICAL_ERROR;
	// Set g_default now, reflecting any changes made to "g" above, in case AutoExecSection(), below,
	// never returns, perhaps because it contains an infinite loop (intentional or not):
	CopyMemory(&g_default, &g, sizeof(global_struct));

	// Could use CreateMutex() but that seems pointless because we have to discover the
	// hWnd of the existing process so that we can close or restart it, so we would have
	// to do this check anyway, which serves both purposes.  Alt method is this:
	// Even if a 2nd instance is run with the /force switch and then a 3rd instance
	// is run without it, that 3rd instance should still be blocked because the
	// second created a 2nd handle to the mutex that won't be closed until the 2nd
	// instance terminates, so it should work ok:
	//CreateMutex(NULL, FALSE, script_filespec); // script_filespec seems a good choice for uniqueness.
	//if (!g_ForceLaunch && !restart_mode && GetLastError() == ERROR_ALREADY_EXISTS)

	// Init global arrays after chances to exit have passed:
	init_vk_to_sc();
	init_sc_to_vk();

	// Allocate from SimpleHeap() to reduce avg. memory load (i.e. since most scripts are small).
	// Do this before LoadFromFile() so that it will be the first thing allocated from the first
	// block, reducing the chance that a new block will have to be allocated just to handle this
	// larger-than-average request:
	if (g_KeyHistory = (KeyHistoryItem *)SimpleHeap::Malloc(MAX_HISTORY_KEYS * sizeof(KeyHistoryItem)))
		ZeroMemory(g_KeyHistory, MAX_HISTORY_KEYS * sizeof(KeyHistoryItem)); // Must be zeroed.
	else
	{
		// Realistically, this should never happen.  But it simplifies the code in many other
		// places if we make sure the KeyHistory array isn't NULL right away:
		MsgBox("Out of memory for key history.");
		return CRITICAL_ERROR;
	}

	LineNumberType load_result = g_script.LoadFromFile();
	if (load_result == LOADING_FAILED) // Error during load (was already displayed by the function call).
		return CRITICAL_ERROR;  // Should return this value because PostQuitMessage() also uses it.
	if (!load_result) // LoadFromFile() relies upon us to do this check.  No lines were loaded, so we're done.
		return 0;

	HWND w_existing = NULL;
	bool close_prior_instance = false;
	if (g_AllowOnlyOneInstance && !restart_mode && !g_ForceLaunch)
	{
		// Note: the title below must be constructed the same was as is done by our
		// CreateWindows(), which is why it's standardized in g_script.mMainWindowTitle:
		if (w_existing = FindWindow(WINDOW_CLASS_MAIN, g_script.mMainWindowTitle))
		{
			if (g_AllowOnlyOneInstance == SINGLE_INSTANCE_IGNORE)
				return 0;
			if (g_AllowOnlyOneInstance != SINGLE_INSTANCE_REPLACE)
				if (MsgBox("An older instance of this #SingleInstance script is already running."
					"  Replace it with this instance?", MB_YESNO, g_script.mFileName) == IDNO)
					return 0;
			// Otherwise:
			close_prior_instance = true;
		}
	}
	if (!close_prior_instance && restart_mode)
		if (w_existing = FindWindow(WINDOW_CLASS_MAIN, g_script.mMainWindowTitle))
			close_prior_instance = true;
	if (close_prior_instance)
	{
		// Now that the script has been validated and is ready to run,
		// close the prior instance.  We wait until now to do this so
		// that the prior instance's "restart" hotkey will still be
		// available to use again after the user has fixed the script.
		PostMessage(w_existing, WM_CLOSE, 0, 0);
		// Wait for it to close before we continue, so that it will deinstall any
		// hooks and unregister any hotkeys it has:
		for (int nInterval = 0; nInterval < 100; ++nInterval)
		{
			Sleep(20);  // No need to use MsgSleep() in this case.
			if (!IsWindow(w_existing))
				break;  // done waiting.
		}
		if (IsWindow(w_existing))
		{
			MsgBox("Could not close the previous instance of this script."  PLEASE_REPORT);
			return CRITICAL_ERROR;
		}
		// Give it a small amount of additional time to completely terminate, even though
		// its main window has already been destroyed:
		Sleep(100);
	}

	// Call this only after closing any existing instance of the program,
	// because otherwise the change to the "focus stealing" setting would never be undone:
	SetForegroundLockTimeout();

	// Create all our windows and the tray icon.  This is done after all other chances
	// to return early due to an error have passed, above.
	if (g_script.CreateWindows() != OK)
		return CRITICAL_ERROR;

	// Activate the hotkeys and any hooks that are required prior to executing the
	// top part (the auto-execute part) of the script so that they will be in effect
	// even if the top part is something that's very involved and requires user
	// interaction:
	Hotkey::AllActivate(true);       // We want these active now in case auto-execute never returns (e.g. loop)
	if (Hotkey::sJoyHotkeyCount)     // Joystick hotkeys requir the timer to be always on.
		SET_MAIN_TIMER
	g_script.AutoExecSection();      // Run the auto-execute part at the top of the script.
	if (!Hotkey::sHotkeyCount)       // No hotkeys are in effect.
		if (!Hotkey::HookIsActive()) // And the user hasn't requested a hook to be activated.
			if (!g_persistent)       // And the script doesn't contain the #Persistent directive.
				g_script.ExitApp();  // We're done.

	// The below is done even if AutoExecSectionTimeout() already set the values once.
	// This is because when the AutoExecute section finally does finish, by definition it's
	// supposed to store the global settings that are currently in effect as the default values.
	// In other words, the only purpose of AutoExecSectionTimeout() is to handle cases where
	// the AutoExecute section takes a long time to complete, or never completes (perhaps because
	// it is being used by the script as a "backround thread" of sorts):
	// Save the values of KeyDelay, WinDelay etc. in case they were changed by the auto-execute part
	// of the script.  These new defaults will be put into effect whenever a new hotkey subroutine
	// is launched.  Each launched subroutine may then change the values for its own purposes without
	// affecting the settings for other subroutines:
	global_clear_state(&g);  // Start with a "clean slate" in both g and g_default.
	CopyMemory(&g_default, &g, sizeof(global_struct));
	// After this point, the values in g_default should never be changed.

	// It seems best to set ErrorLevel to NONE after the auto-execute part of the script is done.
	// However, we do not set it to NONE right before launching each new hotkey subroutine because
	// it's more flexible that way (i.e. the user may want one hotkey subroutine to use the value of
	// ErrorLevel set by another).  This reset was also done by LoadFromFile(), but we do it again
	// here in case the auto-exectute section changed it:
	g_ErrorLevel->Assign(ERRORLEVEL_NONE);

	// Call it in this special mode to kick off the main event loop.
	// Be sure to pass something >0 for the first param or it will
	// return (and we never want this to return):
	MsgSleep(SLEEP_INTERVAL, WAIT_FOR_MESSAGES);
	return 0; // Never executed; avoids compiler warning.
}
