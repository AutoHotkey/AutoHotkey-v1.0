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

	// Set defaults, to be overridden by command line args we receive:
	bool restart_mode = false;

#ifdef AUTOHOTKEYSC
	char *script_filespec = __argv[0];  // i.e. the EXE name.  This is just a placeholder for now.
#else
	#ifdef _DEBUG
	//char *script_filespec = "C:\\A-Source\\AutoHotkey\\Find.aut";
	char *script_filespec = "C:\\Util\\AutoHotkey.ahk";
	//char *script_filespec = "C:\\A-Source\\AutoHotkey\\ZZZZ Test Script.ahk";
	#else
	char *script_filespec = NAME_P ".ini";  // Use this extension for better file associate with editor(s).
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
	for (int i = 1, script_param_num = 1; i < __argc; ++i) // Start at 1 because 0 contains the program name.
	{
		if (switch_processing_is_complete)
		{
			// All args after the script filespec are considered to be parameters for the script:
			snprintf(var_name, sizeof(var_name), "%d", script_param_num);
			if (var = g_script.FindOrAddVar(var_name))
				var->Assign(__argv[i]);
			++script_param_num;
		}
		else if (*__argv[i] == '/')
		{
			switch(toupper(__argv[i][1]))
			{
			case 'R': // Reload
				restart_mode = true;
				break;
			case 'F': // Force the keybd/mouse hook(s) to be installed again even if another instance already did.
				g_ForceLaunch = true;
				break;
			}
		}
		else // since this param does not start with the backslash, the end of the [Switches] section has been reached.
		{
			switch_processing_is_complete = true;  // No more switches allowed after this point.
#ifdef AUTOHOTKEYSC
			--i; // Make the loop process this item again so that it will be treated as a script param.
#else
			script_filespec = __argv[i];
#endif
		}
	}

#ifndef AUTOHOTKEYSC
	size_t filespec_length = strlen(script_filespec);
	if (filespec_length >= CONVERSION_FLAG_LENGTH)
	{
		char *cp = script_filespec + filespec_length - CONVERSION_FLAG_LENGTH;
		// Now cp points to the first dot in the CONVERSION_FLAG of script_filespec (if it has one).
		if (!stricmp(cp, CONVERSION_FLAG))
			return Line::ConvertEscapeChar(script_filespec, '\\', '`');
	}
#endif

	global_init(&g);  // Set defaults prior to the below, since below might override them for AutoIt2 scripts.
	if (g_script.Init(script_filespec, restart_mode) != OK)  // Set up the basics of the script, using the above.
		return CRITICAL_ERROR;

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

	int load_result = g_script.LoadFromFile();
	if (load_result < 0) // Error during load (was already displayed by the function call).
		return CRITICAL_ERROR;  // Should return this value because PostQuitMessage() also uses it.
	if (!load_result) // No lines were loaded, so we're done.
		return 0;

	// Note: the title below must be constructed the same was as is done by our
	// CreateWindows(), which is why it's standardized in g_script.mMainWindowTitle:
	HWND w_existing = FindWindow(WINDOW_CLASS_NAME, g_script.mMainWindowTitle);
	bool close_prior_instance = false;
	if (g_AllowOnlyOneInstance && w_existing && !restart_mode && !g_ForceLaunch)
	{
		// Use a more unique title for this dialog so that subsequent executions of this
		// program can easily find it (though they currently don't):
		//#define NAME_ALREADY_RUNNING NAME_PV " script already running"
		if (MsgBox("An older instance of this #SingleInstance script is already running."
			"  Replace it with this instance?", MB_YESNO, g_script.mFileName) == IDNO)
			return 0;
		else
			close_prior_instance = true;
	}
	if (!close_prior_instance && restart_mode && w_existing)
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
	if (g_script.CreateWindows(hInstance) != OK)
		return CRITICAL_ERROR;

	// Activate the hotkeys and any hooks that are required prior to executing the
	// top part (the auto-execute part) of the script so that they will be in effect
	// even if the top part is something that's very involved and requires user
	// interaction:
	Hotkey::AllActivate();
	g_script.ExecuteFromLine1();     // Run the auto-execute part at the top of the script.
	if (!Hotkey::sHotkeyCount)       // No hotkeys are in effect.
		if (!Hotkey::HookIsActive()) // And the user hasn't requested a hook to be activated.
			g_script.ExitApp();      // We're done.

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
	// ErrorLevel set by another):
	g_ErrorLevel->Assign(ERRORLEVEL_NONE);

	// Call it in this special mode to kick off the main event loop.
	// Be sure to pass something >0 for the first param or it will
	// return (and we never want this to return):
	MsgSleep(SLEEP_INTERVAL, WAIT_FOR_MESSAGES);
	return 0; // Never executed; avoids compiler warning.
}
