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
#ifdef _DEBUG
	//char *script_filespec = "C:\\A-Source\\AutoHotkey\\Find.aut";
	char *script_filespec = "C:\\Util\\AutoHotkey.ahk";
	//char *script_filespec = "C:\\A-Source\\AutoHotkey\\ZZZZ Test Script.ahk";
#else
	char *script_filespec = NAME_P ".ini";  // Use this extension for better file associate with editor(s).
#endif
	bool restart_mode = false;

	// Examine command line args:
	for (int i = 1; i < __argc; ++i) // Start at 1 because 0 contains the program name.
	{
		if (*__argv[i] == '/')
		{
			switch(toupper(__argv[i][1]))
			{
			case 'R': // Restart
				restart_mode = true;
				break;
			case 'F': // Force the keybd/mouse hook(s) to be installed again even if another instance already did.
				g_ForceLaunch = true;
				break;
			}
		}
		else // the last arg found without a leading slash will be the script filespec.
			script_filespec = __argv[i];
	}

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
	//CreateMutex(NULL, TRUE, script_filespec); // script_filespec seems a good choice for uniqueness.
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
	if (g_AllowOnlyOneInstance && w_existing && !restart_mode && !g_ForceLaunch)
	{
		// Use a more unique title for this dialog so that subsequent executions of this
		// program can easily find it (though they currently don't):
		//#define NAME_ALREADY_RUNNING NAME_PV " script already running"
		if (MsgBox("Another instance of this script is already running.  Close it?"
			"\n\nNote: You can force additional instances by running the program with"
			" the /force switch.", MB_YESNO, g_script.mFileName) == IDYES)
			PostMessage(w_existing, WM_CLOSE, 0, 0);
		return 0;
	}

	if (restart_mode && w_existing)
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

	// Call it in this special mode to kick off the main event loop.
	// Be sure to pass something >0 for the first param or it will
	// return (and we never want this to return):
	MsgSleep(SLEEP_INTERVAL, WAIT_FOR_MESSAGES);
	return 0; // Never executed; avoids compiler warning.
}
