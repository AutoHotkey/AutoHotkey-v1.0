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

#include <windows.h>
#include "script.h"
#include "window.h" // for IF_USE_FOREGROUND_WINDOW
#include "application.h" // for MsgSleep()
#include "C:\A-Source\AutoHotkey\VC++\AutoHotkeyNT\resource.h"  // For InputBox.


////////////////////
// Window related //
////////////////////

// Note that there's no action after the IF_USE_FOREGROUND_WINDOW line because that macro handles the action:
#define DETERMINE_TARGET_WINDOW \
HWND target_window;\
IF_USE_FOREGROUND_WINDOW(aTitle, aText, aExcludeTitle, aExcludeText)\
else if (*aTitle || *aText || *aExcludeTitle || *aExcludeText)\
	target_window = WinExist(aTitle, aText, aExcludeTitle, aExcludeText);\
else\
	target_window = g_ValidLastUsedWindow;


ResultType Line::WinMove(char *aTitle, char *aText, char *aX, char *aY
	, char *aWidth, char *aHeight, char *aExcludeTitle, char *aExcludeText)
{
	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return OK;
	// Adapted from the AutoIt3 source:
	RECT rect;
	GetWindowRect(target_window, &rect);
	MoveWindow(target_window
		, *aX && stricmp(aX, "default") ? atoi(aX) : rect.left  // X-position
		, *aY && stricmp(aY, "default") ? atoi(aY) : rect.top   // Y-position
		, *aWidth && stricmp(aWidth, "default") ? atoi(aWidth) : rect.right - rect.left
		, *aHeight && stricmp(aHeight, "default") ? atoi(aHeight) : rect.bottom - rect.top
		, TRUE);  // Do repaint.
	DoWinDelay;
	return OK;  // Always successful, like AutoIt.
}



ResultType Line::ControlSend(char *aControl, char *aKeysToSend, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText, modLR_type aModifiersLR)
{
	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return OK;
	HWND control_window = ControlExist(target_window, aControl);
	if (!control_window)
		return OK;
	SendKeys(aKeysToSend, aModifiersLR, control_window);
	return OK;
}



ResultType Line::ControlLeftClick(char *aControl, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
{
	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return OK;
	HWND control_window = ControlExist(target_window, aControl);
	if (!control_window)
		return OK;
	PostMessage(control_window, WM_LBUTTONDOWN, MK_LBUTTON, 0);
	PostMessage(control_window, WM_LBUTTONUP, 0, 0);
	return OK;
}



ResultType Line::ControlGetText(char *aControl, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
{
	DETERMINE_TARGET_WINDOW
	HWND control_window = target_window ? ControlExist(target_window, aControl) : NULL;
	// Even if control_window is NULL, we want to continue on so that the output
	// param is set to be the empty string, which is the proper thing to do
	// rather than leaving whatever was in there before.

	// Handle the output parameter.  This section is similar to that in
	// PerformAssign().  Note: Using GetWindowTextTimeout() vs. GetWindowText()
	// because it is able to get text from more types of controls (e.g. large edit controls):
	VarSizeType space_needed = control_window ? GetWindowTextTimeout(control_window) + 1 : 1; // 1 for terminator.

	// Set up the var, enlarging it if necessary.  If the OUTPUT_VAR is of type VAR_CLIPBOARD,
	// this call will set up the clipboard for writing:
	if (OUTPUT_VAR->Assign(NULL, space_needed - 1) != OK)
		return FAIL;  // It already displayed the error.
	// Fetch the text directly into the var.  Also set the length explicitly
	// in case actual size written was off from the esimated size (since
	// GetWindowTextLength() can return more space that will actually be required
	// in certain circumstances, see MS docs):
	if (control_window)
	{
		OUTPUT_VAR->Length() = (VarSizeType)GetWindowTextTimeout(control_window
			, OUTPUT_VAR->Contents(), space_needed);
		if (!OUTPUT_VAR->Length())
			// There was no text to get or GetWindowTextTimeout() failed.
			*OUTPUT_VAR->Contents() = '\0';  // Safe because Assign() gave us a non-constant memory area.
	}
	else
	{
		*OUTPUT_VAR->Contents() = '\0';
		OUTPUT_VAR->Length() = 0;
	}
	// Consider the above to be always successful, even if the window wasn't found, except
	// when below returns an error:
	return OUTPUT_VAR->Close();  // In case it's the clipboard.
}



ResultType Line::StatusBarGetText(char *aPart, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
{
	DETERMINE_TARGET_WINDOW
	HWND control_window = target_window ? ControlExist(target_window, "msctls_statusbar321") : NULL;
	StatusBarUtil(OUTPUT_VAR, control_window, atoi(aPart)); // It will handle any zero part# for us.
	return OK; // Even if it fails, seems best to return OK so that subroutine can continue.
}



ResultType Line::StatusBarWait(char *aTextToWaitFor, char *aSeconds, char *aPart, char *aTitle, char *aText
	, char *aInterval, char *aExcludeTitle, char *aExcludeText)
{
	DETERMINE_TARGET_WINDOW
	// Make a copy of any memory areas that are volatile (due to Deref buf being overwritten
	// if a new hotkey subroutine is launched while we are waiting) but whose contents we
	// need to refer to while we are waiting:
	char text_to_wait_for[4096];
	strlcpy(text_to_wait_for, aTextToWaitFor, sizeof(text_to_wait_for));
	HWND control_window = target_window ? ControlExist(target_window, "msctls_statusbar321") : NULL;
	StatusBarUtil(NULL, control_window, atoi(aPart) // It will handle a NULL control_window or zero part# for us.
		, text_to_wait_for, *aSeconds ? atoi(aSeconds)*1000 : -1 // Blank->indefinite.  0 means 500ms.
		, atoi(aInterval));
	return OK; // Even if it fails, seems best to return OK so that subroutine can continue.
}



ResultType Line::WinSetTitle(char *aTitle, char *aText, char *aNewTitle, char *aExcludeTitle, char *aExcludeText)
{
	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return OK;
	SetWindowText(target_window, aNewTitle);
	return OK;
}



ResultType Line::WinGetTitle(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	DETERMINE_TARGET_WINDOW
	// Even if target_window is NULL, we want to continue on so that the output
	// param is set to be the empty string, which is the proper thing to do
	// rather than leaving whatever was in there before.

	// Handle the output parameter.  See the comments in ACT_CONTROLGETTEXT for details.
	VarSizeType space_needed = target_window ? GetWindowTextLength(target_window) + 1 : 1; // 1 for terminator.
	if (OUTPUT_VAR->Assign(NULL, space_needed - 1) != OK)
		return FAIL;  // It already displayed the error.
	if (target_window)
	{
		OUTPUT_VAR->Length() = (VarSizeType)GetWindowText(target_window
			, OUTPUT_VAR->Contents(), space_needed);
		if (!OUTPUT_VAR->Length())
			// There was no text to get or GetWindowTextTimeout() failed.
			*OUTPUT_VAR->Contents() = '\0';  // Safe because Assign() gave us a non-constant memory area.
	}
	else
	{
		*OUTPUT_VAR->Contents() = '\0';
		OUTPUT_VAR->Length() = 0;
	}
	return OUTPUT_VAR->Close();  // In case it's the clipboard.
}



ResultType Line::WinGetText(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	DETERMINE_TARGET_WINDOW
	// Even if target_window is NULL, we want to continue on so that the output
	// variables are set to be the empty string, which is the proper thing to do
	// rather than leaving whatever was in there before:
	if (!target_window)
		return OUTPUT_VAR->Assign(); // Tell it not to free the memory by not calling with "".
	length_and_buf_type sab;
	sab.buf = NULL; // Tell it just to calculate the length this time around.
	sab.total_length = sab.capacity = 0; // Init
	EnumChildWindows(target_window, EnumChildGetText, (LPARAM)&sab);
	if (!sab.total_length)
		return OUTPUT_VAR->Assign(); // Tell it not to free the memory by omitting all params.
	// Set up the var, enlarging it if necessary.  If the OUTPUT_VAR is of type VAR_CLIPBOARD,
	// this call will set up the clipboard for writing:
	if (OUTPUT_VAR->Assign(NULL, (VarSizeType)sab.total_length) != OK)
		return FAIL;  // It already displayed the error.
	// Fetch the text directly into the var.  Also set the length explicitly
	// in case actual size written was off from the esimated size (since
	// GetWindowTextLength() can return more space that will actually be required
	// in certain circumstances, see MS docs):
	sab.buf = OUTPUT_VAR->Contents();
	sab.total_length = 0; // Init
	sab.capacity = OUTPUT_VAR->Capacity(); // Because capacity might be a little larger than we asked for.
	EnumChildWindows(target_window, EnumChildGetText, (LPARAM)&sab);
	OUTPUT_VAR->Length() = (VarSizeType)sab.total_length;  // In case it wound up being smaller than expected.
	if (!sab.total_length)
		// Something went wrong, so make sure we set to empty string.
		*OUTPUT_VAR->Contents() = '\0';  // Safe because Assign() gave us a non-constant memory area.
	return OUTPUT_VAR->Close();  // In case it's the clipboard.
}



BOOL CALLBACK EnumChildGetText(HWND aWnd, LPARAM lParam)
{
	if (!g.DetectHiddenText && !IsWindowVisible(aWnd))
		return TRUE;  // This child/control is hidden and user doesn't want it considered, so skip it.
	#define psab ((length_and_buf_type *)lParam)
	int length;
	if (psab->buf)
		length = GetWindowTextTimeout(aWnd, psab->buf + psab->total_length
			, (int)(psab->capacity - psab->total_length)); // Not +1.
	else
		length = GetWindowTextTimeout(aWnd);
	psab->total_length += length;
	if (length)
	{
		if (psab->buf)
		{
			if (psab->capacity - psab->total_length > 2) // Must be >2 due to zero terminator.
			{
				strcpy(psab->buf + psab->total_length, "\r\n"); // Something to delimit each control's text.
				psab->total_length += 2;
			}
			// else don't increment total_length
		}
		else
			psab->total_length += 2; // Since buf is NULL, accumulate the size that *would* be needed.
	}
	return TRUE; // Continue enumeration through all the windows.
}



ResultType Line::WinGetPos(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	DETERMINE_TARGET_WINDOW
	// Even if target_window is NULL, we want to continue on so that the output
	// variables are set to be the empty string, which is the proper thing to do
	// rather than leaving whatever was in there before.
	RECT rect;
	if (target_window)
		GetWindowRect(target_window, &rect);
	else // ensure it's initialized for possible later use:
		rect.bottom = rect.left = rect.right = rect.top = 0;

	ResultType result = OK; // Set default;

	if (VARRAW_ARG1)
		if (target_window)
		{
			if (!VARARG1->Assign((int)rect.left))  // X position
				result = FAIL;
		}
		else
			if (!VARARG1->Assign(""))
				result = FAIL;
	if (VARRAW_ARG2)
		if (target_window)
		{
			if (!VARARG2->Assign((int)rect.top))  // Y position
				result = FAIL;
		}
		else
			if (!VARARG2->Assign(""))
				result = FAIL;
	if (VARRAW_ARG3) // else user didn't want this value saved to an output param
		if (target_window)
		{
			if (!VARARG3->Assign((int)rect.right - rect.left))  // Width
				result = FAIL;
		}
		else
			if (!VARARG3->Assign("")) // Set it to be empty to signal the user that the window wasn't found.
				result = FAIL;
	if (VARRAW_ARG4)
		if (target_window)
		{
			if (!VARARG4->Assign((int)rect.bottom - rect.top))  // Height
				result = FAIL;
		}
		else
			if (!VARARG4->Assign(""))
				result = FAIL;

	return result;
}



/////////////////
// Main Window //
/////////////////

LRESULT CALLBACK MainWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
	switch (iMsg)
	{
	case WM_COMMAND: // If an application processes this message, it should return zero.
		// See if an item was selected from the tray menu:
		switch (LOWORD(wParam))
		{
		case ID_TRAY_OPEN:
			ShowMainWindow();
			return 0;
		case ID_TRAY_EDITSCRIPT:
		{
			TitleMatchModes old_mode = g.TitleMatchMode;
			g.TitleMatchMode = FIND_ANYWHERE_FAST;
			HWND hwnd = WinExist(g_script.mFileName);
			g.TitleMatchMode = old_mode;
			if (hwnd)
				SetForegroundWindowEx(hwnd);
			else
				if (!Script::ActionExec(g_script.mFileSpec, ""))
					MsgBox("Perhaps the file extension isn't associated with an application.");
			return 0;
		}
		case ID_TRAY_RELOADSCRIPT:
			g_script.Reload();
			return 0;
		case ID_TRAY_WINDOWSPY:
		{
			char buf[MAX_PATH + 512];
			if (!GetModuleFileName(NULL, buf, sizeof(buf))) // realistically, probably can't fail.
				break;
			char *last_backslash = strrchr(buf, '\\');
			if (!last_backslash)
				break;
			last_backslash[1] = '\0';
			snprintfcat(buf, sizeof(buf), "AU3_Spy.exe");
			Script::ActionExec(buf, "");
			return 0;
		}
		case ID_TRAY_HELP:
			MsgBox("Help's placeholder.");
			return 0;
		case ID_TRAY_SUSPEND:
			// Maybe can use CheckMenuItem() with this, somehow.
			return 0;
		case ID_TRAY_EXIT:
			g_script.ExitApp();
			return 0;
		} // Inner switch()
		break;

	case AHK_NOTIFYICON:  // Tray icon clicked on.
	{
        switch(lParam)
        {
		case WM_LBUTTONDBLCLK:
			ShowMainWindow();
			return 0;
		case WM_RBUTTONDOWN:
		{
			HMENU hMenu = LoadMenu(g_hInstance, MAKEINTRESOURCE(IDR_MENU1));
			// TrackPopupMenu cannot display the menu bar so get
			// the handle to the first shortcut menu.
 			hMenu = GetSubMenu(hMenu, 0);
			SetMenuDefaultItem(hMenu, ID_TRAY_OPEN, FALSE);
			POINT pt;
			GetCursorPos(&pt);
			SetForegroundWindow(hWnd); // Always call this right before TrackPopupMenu(), even window if hidden.
			// Set this so that if a new recursion layer is triggered by TrackPopupMenuEx's having
			// dispatched a hotkey message to our main window proc, IsCycleComplete() knows that
			// this layer here does not need to have its original foreground window restored to the foreground.
			// Also, this allows our main Window Proc to close the popup menu upon receive of any hotkey,
			// which is probably a good idea since most hotkeys change the foreground window and if that
			// happens, the menu cannot be dismissed (ever?) except by selecting one of the items in the
			// menu (which is often undesirable).
			g_TrayMenuIsVisible = true;
			TrackPopupMenuEx(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON,	pt.x, pt.y, hWnd, NULL);
			g_TrayMenuIsVisible = false;
			DestroyMenu(hMenu);
			PostMessage(hWnd, WM_NULL, 0, 0); // MSDN recommends this to prevent menu from closing on 2nd click.
			return 0;
		}
		} // Inner switch()
		break;
	} // case AHK_NOTIFYICON

	case AHK_DIALOG:  // User defined msg sent from MsgBox() or FileSelectFile().
	{
		// Always call this to close the clipboard if it was open (e.g. due to a script
		// line such as "MsgBox, %clipboard%" that got us here).  Seems better just to
		// do this rather than incurring the delay and overhead of a MsgSleep() call:
		CLOSE_CLIPBOARD_IF_OPEN;
		
		// Since we're here, it means the modal dialog's pump is now running and the script
		// that displayed the dialog is waiting for the dialog to finish.  Because of this,
		// the main timer should not be left enabled because otherwise timer messages will
		// just pile up in our thread's message queue (since our main msg pump isn't running),
		// which probably hurts performance.  The main timer is owned by the thread rather
		// than the main window because there seems to be cases where the timer message
		// is sent directly to this procedure, bypassing the main msg pump entirely,
		// which is not what we want.  UPDATE: Handling of the timer has been simplified,
		// so it should now be impossible for the timer to be active if we're here, so
		// this isn't necessary:
		//PURGE_AND_KILL_MAIN_TIMER

		// Ensure that the app's top-most window (the modal dialog) is the system's
		// foreground window.  This doesn't use FindWindow() since it can hang in rare
		// cases.  And GetActiveWindow, GetTopWindow, GetWindow, etc. don't seem appropriate.
		// So EnumWindows is probably the way to do it:
		HWND top_box = WinActivateOurTopDialog();
		if (top_box && (UINT)wParam > 0)
			// Caller told us to establish a timeout for this modal dialog (currently always MessageBox):
			SetTimer(top_box, g_nMessageBoxes, (UINT)wParam * 1000, DialogTimeout);
		// else: if !top_box: no error reporting currently.
		return 0;
	}

	case WM_HOTKEY: // As a result of this app having previously called RegisterHotkey().
	case AHK_HOOK_HOTKEY:  // Sent from this app's keyboard or mouse hook.
	{
		if (g_IgnoreHotkeys)
			// Used to prevent runaway hotkeys, or too many happening due to key-repeat feature.
			// It can also be used to prevent a call to MsgSleep() from accepting new hotkeys
			// in cases where the caller's activity might be interferred with by the launch
			// of a new hotkey subroutine, such as reading or writing to the clipboard:
			return 0;
		// Post it to the thread, just in case the OS tries to be "helpful" and
		// directly call the WindowProc (i.e. this function) rather than actually
		// posting the message.  We don't want to be called, we want the main loop
		// to handle this message:
		PostThreadMessage(GetCurrentThreadId(), iMsg, wParam, lParam);
		if (g_TrayMenuIsVisible)
		{
			// Ok this is a little strange, but the thought here is that if the tray menu is
			// displayed, it should be closed prior to executing any new hotkey.  This is
			// because hotkeys usually cause other windows to become active, and once that
			// happens, the tray menu cannot be closed except by choosing a menu item in it
			// (which is often undesirable):
			SendMessage(hWnd, WM_CANCELMODE, 0, 0);
			// The menu is now gone because the above should have called this function
			// recursively to close the it.  Now, rather than continuing in this
			// recursion layer, it seems best to return to the caller so that the menu
			// will be destroyed and g_TrayMenuIsVisible set to false.  After that is done,
			// the next call to MsgSleep() should notice the hotkey we posted above and
			// act upon it.
			// The above section has been tested and seems to work as expected.
			// UPDATE: Below doesn't work if there's a MsgBox() window displayed because
			// the caller to which we return is the MsgBox's msg pump, and that pump
			// ignores any messages for our thread so they just sit there.  So instead
			// of returning, call MsgSleep() without resetting the value of
			// g_TrayMenuIsVisible (so that it can use it).  When MsgSleep() returns,
			// we will return to our caller, which in this case should be TrackPopupMenuEx's
			// msg pump.  That pump should immediately return also since we've already
			// closed the menu.  And we will let it set the value of g_TrayMenuIsVisible
			// to false at that time rather than doing it here or in IsCycleComplete().
			// In keeping with the above, don't return:
			//return 0;
		}
		MsgSleep();  // Now call the main loop to handle the message we just posted (and any others).
		return 0; // Not sure if this is the correct return value.  It probably doesn't matter.
	}

	case WM_SYSCOMMAND:
		if (wParam == SC_CLOSE && hWnd == g_hWnd) // i.e. behave this way only for main window.
		{
			// The user has either clicked the window's "X" button, chosen "Close"
			// from the system (upper-left icon) menu, or pressed Alt-F4.  In all
			// these cases, we want to hide the window rather than actually closing
			// it.  If the user really wishes to exit the program, a File->Exit
			// menu option may be available, or use the Tray Icon, or launch another
			// instance which will close the previous, etc.
			ShowWindow(g_hWnd, SW_HIDE);
			return 0;
		}
		break;
	case WM_DESTROY:
		// MSDN: If an application processes this message, it should return zero.
		if (hWnd == g_hWnd) // i.e. not the SplashText window or anything other than the main.
		{
			// Once we do this, it appears that no new dialogs can be created
			// (perhaps no new windows of any kind?).  Also: Even if this function
			// was called by MessageBox()'s message loop, it appears that when we
			// call PostQuitMessage(), the MessageBox routine sees it and knows
			// to destroy itself, thus cascading the Quit state through any other
			// underlying MessageBoxes that may exist, until finally we wind up
			// back at our main message loop, which handles the WM_QUIT:
			PostQuitMessage(0);
			return 0;
		}
		// Otherwise, some window of ours other than our main window was destroyed
		// (impossible if we're here?)
		// Let DefWindowProc() handle it:
		break;
	case WM_CREATE:
		// MSDN: If an application processes this message, it should return zero to continue
		// creation of the window. If the application returns 1, the window is destroyed and
		// the CreateWindowEx or CreateWindow function returns a NULL handle.
		return 0;

	// Can't do this without ruining MsgBox()'s ShowWindow().
	// UPDATE: It doesn't do that anymore so leave this enabled for now.
	case WM_SIZE:
		if (hWnd == g_hWnd)
		{
			if (wParam == SIZE_MINIMIZED)
				// Minimizing the main window hides it.
				ShowWindow(g_hWnd, SW_HIDE);
			else
				MoveWindow(g_hWndEdit, 0, 0, LOWORD(lParam), HIWORD(lParam), TRUE);
			return 0; // The correct return value for this msg.
		}
		// Should probably never happen size SplashText window should never receive this msg?:
		break;

	case WM_SETFOCUS:
		if (hWnd == g_hWnd)
		{
			SetFocus(g_hWndEdit);  // Always focus the edit window, since it's the only navigable control.
			return 0;
		}
		break;

	} // end main switch

	return DefWindowProc(hWnd, iMsg, wParam, lParam);
}



ResultType ShowMainWindow(char *aContents, bool aJumpToBottom)
{
	ResultType result = OK; // Set default return value.
	// Update the text before doing anything else, since it might be a little less disruptive
	// while being quicker to do it while the window is hidden or non-foreground:
	char buf_temp[1024 * 8];
	if (!aContents)
	{
		Line::LogToText(buf_temp, sizeof(buf_temp));
		aContents = buf_temp;
		aJumpToBottom = true;
	}
	// else, aContents can be empty string, which clears the window.
	// Unlike SetWindowText(), this method seems to expand tab characters:
	if (SendMessage(g_hWndEdit, WM_SETTEXT, 0, (LPARAM)aContents) != TRUE) // FALSE or some non-TRUE value.
		result = FAIL;
	if (!IsWindowVisible(g_hWnd))
	{
		ShowWindow(g_hWnd, SW_SHOW);
		if (IsIconic(g_hWnd)) // This happens whenver the window was last hidden via the minimize button.
			ShowWindow(g_hWnd, SW_RESTORE);
	}
	if (g_hWnd != GetForegroundWindow())
		if (!SetForegroundWindow(g_hWnd))
			if (!SetForegroundWindowEx(g_hWnd))  // Only as a last resort, since it uses AttachThreadInput()
				result = FAIL;
	if (aJumpToBottom)
	{
		SendMessage(g_hWndEdit, EM_LINESCROLL , 0, 999999);
		//SendMessage(g_hWndEdit, EM_SETSEL, -1, -1);
		//SendMessage(g_hWndEdit, EM_SCROLLCARET, 0, 0);
	}
	return result;
}



//////////////
// InputBox //
//////////////

ResultType InputBox(Var *aOutputVar, char *aTitle, char *aText, bool aHideInput)
{
	// Note: for maximum compatibility with existing AutoIt2 scripts, do not
	// set ErrorLevel to ERRORLEVEL_ERROR when the user presses cancel.  Instead,
	// just set the output var to be blank.
	if (g_nInputBoxes >= MAX_INPUTBOXES)
	{
		// Have a maximum to help prevent runaway hotkeys due to key-repeat feature, etc.
		MsgBox("The maximum number of InputBoxes has been reached." ERR_ABORT);
		return FAIL;
	}
	if (!aOutputVar) return FAIL;
	if (!aText) aText = "";
	if (!aTitle || !*aTitle) aTitle = NAME_PV;
	// Limit the size of what we were given to prevent unreasonably huge strings from
	// possibly causing a failure in CreateDialog():
	char title[DIALOG_TITLE_SIZE];
	char text[2048];  // Probably can't fit more due to the limited size of the dialog's text area.
	strlcpy(title, aTitle, sizeof(title));
	strlcpy(text, aText, sizeof(text));
	g_InputBox[g_nInputBoxes].title = title;
	g_InputBox[g_nInputBoxes].text = text;
	g_InputBox[g_nInputBoxes].output_var = aOutputVar;
	g_InputBox[g_nInputBoxes].password_char = aHideInput ? '*' : '\0';
	g.WaitingForDialog = true;
	++g_nInputBoxes;
	// Specify NULL as the owner since we want to be able to have the main window in the foreground
	// even if there are InputBox windows:
	INT_PTR result = DialogBox(g_hInstance, MAKEINTRESOURCE(IDD_INPUTBOX), NULL, InputBoxProc);
	--g_nInputBoxes;
	g.WaitingForDialog = false;  // IsCycleComplete() relies on this.
	if (result == -1)
	{
		MsgBox("The InputBox window could not be displayed.");
		return FAIL;
	}
	// In other failure cases than the above, the error should have already been displayed
	// by InputBoxProc():
	return result == FAIL ? FAIL : OK;  // OK if user pressed the OK or Cancel button.
}



INT_PTR CALLBACK InputBoxProc(HWND hWndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
// MSDN:
// Typically, the dialog box procedure should return TRUE if it processed the message,
// and FALSE if it did not. If the dialog box procedure returns FALSE, the dialog
// manager performs the default dialog operation in response to the message.
{
	HWND hControl;
	int target_index = g_nInputBoxes - 1;  // Set default array index for g_InputBox[].
	switch(uMsg)
	{
	case WM_INITDIALOG:
	{
		// Clipboard may be open if its contents were used to build the text or title
		// of this dialog (e.g. "InputBox, out, %clipboard%").  It's best to do this before
		// anything that might take a relatively long time (e.g. SetForegroundWindowEx()):
		CLOSE_CLIPBOARD_IF_OPEN;
		// Caller has ensured that g_nInputBoxes > 0:
		#define CURR_INPUTBOX g_InputBox[target_index]
		CURR_INPUTBOX.hwnd = hWndDlg;
		SetWindowText(hWndDlg, CURR_INPUTBOX.title);
		if (hControl = GetDlgItem(hWndDlg, IDC_INPUTPROMPT))
			SetWindowText(hControl, CURR_INPUTBOX.text);
		if (hWndDlg != GetForegroundWindow()) // Normally it will be since the template has this property.
			SetForegroundWindowEx(hWndDlg);   // Try to force it to the foreground.
		if (CURR_INPUTBOX.password_char)
			SendDlgItemMessage(hWndDlg, IDC_INPUTEDIT, EM_SETPASSWORDCHAR, CURR_INPUTBOX.password_char, 0);
		return TRUE; // i.e. let the system set the keyboard focus to the first visible control.
	}
	case WM_COMMAND:
		// In this case, don't use (g_nInputBoxes - 1) as the index because it might
		// not correspond to the g_InputBox[] array element that belongs to hWndDlg.
		// This is because more than one input box can be on the screen at the same time.
		// If the user choses to work with on underneath instead of the most recent one,
		// we would be called with an hWndDlg whose index is less than the most recent
		// one's index (g_nInputBoxes - 1).  Instead, search the array for a match.
		// Work backward because the most recent one(s) are more likely to be a match:
		for (; target_index >= 0; --target_index)
			if (g_InputBox[target_index].hwnd == hWndDlg)
				break;
		if (target_index < 0)  // Should never happen if things are designed right.
			return FALSE;
		switch (LOWORD(wParam))
		{
		case IDOK:
		case IDCANCEL:
		{
			WORD return_value = LOWORD(wParam);  // Set default, i.e. ID_OK or ID_CANCEL
			if (   !(hControl = GetDlgItem(hWndDlg, IDC_INPUTEDIT))   )
				return_value = (WORD)FAIL;
			else
			{
				#define INPUTBOX_VAR CURR_INPUTBOX.output_var
				VarSizeType space_needed = (LOWORD(wParam) == IDCANCEL) ? 1 : GetWindowTextLength(hControl) + 1;
				// Set up the var, enlarging it if necessary.  If the OUTPUT_VAR is of type VAR_CLIPBOARD,
				// this call will set up the clipboard for writing:
				if (INPUTBOX_VAR->Assign(NULL, space_needed - 1) != OK)
					// It will have already displayed the error.  Displaying errors in a callback
					// function like this one isn't that good, since the callback won't return
					// to its caller in a timely fashion.  However, these type of errors are so
					// rare it's not a priority to change all the called functions (and the functions
					// they call) to skip the displaying of errors and just return FAIL instead.
					// In addition, this callback function has been tested with a MsgBox() call
					// inside and it doesn't seem to cause any crashes or undesirable behavior other
					// than the fact that the InputBox window is not dismissed until the MsgBox
					// window is dismissed:
					return_value = (WORD)FAIL;
				else
				{
					// Write to the variable:
					if (LOWORD(wParam) == IDCANCEL)
						// It's length was already set by the above call to Assign().
						*INPUTBOX_VAR->Contents() = '\0';
					else
					{
						INPUTBOX_VAR->Length() = (VarSizeType)GetWindowText(hControl
							, INPUTBOX_VAR->Contents(), space_needed);
						if (!INPUTBOX_VAR->Length())
							// There was no text to get or GetWindowTextTimeout() failed.
							*INPUTBOX_VAR->Contents() = '\0';  // Safe because Assign() gave us a non-constant memory area.
					}
					return_value = (WORD)INPUTBOX_VAR->Close();  // In case it's the clipboard.
				}
			}
			EndDialog(hWndDlg, return_value);
			return TRUE;
		} // case
		} // Inner switch()
	} // Outer switch()
	// Otherwise, let the dialog handler do its default action:
	return FALSE;
}



///////////////////
// Mouse related //
///////////////////

ResultType Line::MouseClickDrag(vk_type aVK, int aX1, int aY1, int aX2, int aY2, int aSpeed)
// Note: This is based on code in the AutoIt3 source.
{
	// Autoit3: Check for x without y
	// MY: In case this was called from a source that didn't already validate this:
	if (   (aX1 == INT_MIN && aY1 != INT_MIN) || (aX1 != INT_MIN && aY1 == INT_MIN)   )
		return FAIL;
	if (   (aX2 == INT_MIN && aY2 != INT_MIN) || (aX2 != INT_MIN && aY2 == INT_MIN)   )
		return FAIL;

	// Move the mouse to the start position if we're not starting in the current position:
	if (aX1 != COORD_UNSPECIFIED && aY1 != COORD_UNSPECIFIED) // Otherwise don't bother.
		MouseMove(aX1, aY1, aSpeed);

	// The drag operation fails unless speed is now >=2
	// My: I asked him about the above, saying "Have you discovered that insta-drags
	// almost always fail?" and he said "Yeah, it was weird, absolute lack of drag...
	// Don't know if it was my config or what."
	// My: But testing reveals "insta-drags" work ok, at least on my system, so
	// leaving them enabled.  User can easily increase the speed if there's
	// any problem:
	//if (aSpeed < 2)
	//	aSpeed = 2;

	// Always sleep a certain minimum amount of time between events to improve reliability,
	// but allow the user to specify a higher time if desired:
	#define MOUSE_SLEEP MsgSleep(g.KeyDelay > 10 ? g.KeyDelay : 10)

	// Do the drag operation
	switch (aVK)
	{
	case VK_LBUTTON:
		mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
		MOUSE_SLEEP;
		MouseMove(aX2, aY2, aSpeed);
		MOUSE_SLEEP;
		mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
		break;
	case VK_RBUTTON:
		mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
		MOUSE_SLEEP;
		MouseMove(aX2, aY2, aSpeed);
		MOUSE_SLEEP;
		mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
		break;
	case VK_MBUTTON:
		mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0);
		MOUSE_SLEEP;
		MouseMove(aX2, aY2, aSpeed);
		MOUSE_SLEEP;
		mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0);
		break;
	}
	// It seems best to always do this one too in case the script line that caused
	// us to be called here is followed immediately by another script line which
	// is either another mouse click or something that relies upon this mouse drag
	// having been completed:
	MOUSE_SLEEP;
	return OK;
}



ResultType Line::MouseClick(vk_type aVK, int aX, int aY, int aClickCount, int aSpeed)
// Note: This is based on code in the AutoIt3 source.
{
	// Autoit3: Check for x without y
	// MY: In case this was called from a source that didn't already validate this:
	if (   (aX == INT_MIN && aY != INT_MIN) || (aX != INT_MIN && aY == INT_MIN)   )
		// This was already validated during load so should never happen
		// unless this function was called directly from somewhere else
		// in the app, rather than by a script line:
		return FAIL;

	if (aClickCount <= 0)
		// Allow this to simply "do nothing", because it increases flexibility
		// in the case where the number of clicks is a dereferenced script variable
		// that may sometimes (by intent) resolve to zero:
		return OK;

	// Do we need to move the mouse?
	if (aX != COORD_UNSPECIFIED && aY != COORD_UNSPECIFIED) // Otherwise don't bother.
		MouseMove(aX, aY, aSpeed);

	for (int i = 0; i < aClickCount; ++i)
	{
		// Note: It seems best to always Sleep a certain minimum time between events
		// because the click-down event may cause the target app to do something which
		// changes the context or nature of the click-up event.  AutoIt3 has also been
		// revised to do this.
		switch (aVK)
		{
		case VK_LBUTTON:
			mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
			MOUSE_SLEEP;
			mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
			break;
		case VK_RBUTTON:
			mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
			MOUSE_SLEEP;
			mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
			break;
		case VK_MBUTTON:
			mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0);
			MOUSE_SLEEP;
			mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0);
			break;
		}
		// It seems best to always do this one too in case the script line that caused
		// us to be called here is followed immediately by another script line which
		// is either another mouse click or something that relies upon the mouse click
		// having been completed:
		MOUSE_SLEEP;
	}

	return OK;
}



void Line::MouseMove(int aX, int aY, int aSpeed)
// Note: This is based on code in the AutoIt3 source.
{
	POINT	ptCur;
	RECT	rect;
	int		xCur, yCur;
	int		delta;
	const	int	nMinSpeed = 32;

	if (aSpeed < 0)
		aSpeed = 0;  // 0 is the fastest.
	else
		if (aSpeed > MAX_MOUSE_SPEED)
			aSpeed = MAX_MOUSE_SPEED;

	GetWindowRect(GetForegroundWindow(), &rect);
	aX += rect.left; 
	aY += rect.top;

	// AutoIt3: Get size of desktop
	GetWindowRect(GetDesktopWindow(), &rect);

	// AutoIt3: Convert our coords to mouse_event coords
	aX = ((65535 * aX) / (rect.right - 1)) + 1;
	aY = ((65535 * aY) / (rect.bottom - 1)) + 1;


	// AutoIt3: Are we slowly moving or insta-moving?
	if (aSpeed == 0)
	{
		mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, aX, aY, 0, 0);
		MOUSE_SLEEP; // Should definitely do this in case the action immediately after this is a click.
		return;
	}

	// AutoIt3: Sanity check for speed
	if (aSpeed < 0 || aSpeed > MAX_MOUSE_SPEED)
		aSpeed = g.DefaultMouseSpeed;

	// AutoIt3: So, it's a more gradual speed that is needed :)
	GetCursorPos(&ptCur);
	xCur = ((ptCur.x * 65535) / (rect.right - 1)) + 1;
	yCur = ((ptCur.y * 65535) / (rect.bottom - 1)) + 1;

	while (xCur != aX || yCur != aY)
	{
		if (xCur < aX)
		{
			delta = (aX - xCur) / aSpeed;
			if (delta == 0 || delta < nMinSpeed)
				delta = nMinSpeed;
			if ((xCur + delta) > aX)
				xCur = aX;
			else
				xCur += delta;
		} 
		else 
			if (xCur > aX)
			{
				delta = (xCur - aX) / aSpeed;
				if (delta == 0 || delta < nMinSpeed)
					delta = nMinSpeed;
				if ((xCur - delta) < aX)
					xCur = aX;
				else
					xCur -= delta;
			}

		if (yCur < aY)
		{
			delta = (aY - yCur) / aSpeed;
			if (delta == 0 || delta < nMinSpeed)
				delta = nMinSpeed;
			if ((yCur + delta) > aY)
				yCur = aY;
			else
				yCur += delta;
		} 
		else 
			if (yCur > aY)
			{
				delta = (yCur - aY) / aSpeed;
				if (delta == 0 || delta < nMinSpeed)
					delta = nMinSpeed;
				if ((yCur - delta) < aY)
					yCur = aY;
				else
					yCur -= delta;
			}

		mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, xCur, yCur, 0, 0);
		MOUSE_SLEEP;
	}
}



ResultType Line::MouseGetPos()
// Returns OK or FAIL.
// This has been adapted from the AutoIt3 source.
{
	if (!VARRAW_ARG1 && !VARRAW_ARG2)
		// This is an error because it was previously verified that at least one is non-blank:
		return LineError("MouseGetPos() was called without any output vars." PLEASE_REPORT ERR_ABORT);

	RECT rect;
	POINT pt;
	GetCursorPos(&pt);  // Realistically, can't fail?
	HWND fore_win = GetForegroundWindow();
	if (fore_win)
		GetWindowRect(fore_win, &rect);
	else // ensure it's initialized for later calculations:
		rect.bottom = rect.left = rect.right = rect.top = 0;

	ResultType result = OK; // Set default;

	if (VARRAW_ARG1) // else the user didn't want the X coordinate, just the Y.
		if (!VARARG1->Assign((int)(pt.x - rect.left)))
			result = FAIL;
	if (VARRAW_ARG2) // else the user didn't want the Y coordinate, just the X.
		if (!VARARG2->Assign((int)(pt.y - rect.top)))
			result = FAIL;

	return result;
}



///////////////////////////////
// Related to other commands //
///////////////////////////////

ResultType Line::PerformAssign()
// Returns OK or FAIL.
{
	//if (OUTPUT_VAR == NULL)
	//	return LineError("PerformAssign() was called with a NULL target variable." PLEASE_REPORT ERR_ABORT);

	// Find out if OUTPUT_VAR (the var being assigned to) is dereferenced (mentioned) in
	// this line's first arg.  If it isn't, things are much simpler.  Note:
	// If OUTPUT_VAR is the clipboard, it can be used in the source deref(s) while also
	// being the target -- without having to use the deref buffer -- because the clipboard
	// has it's own temp buffer: the memory area to which the result is written.
	// The prior content of the clipboard remains available in its other memory area
	// until Commit() is called (i.e. long enough for our purposes):
	bool target_is_involved_in_source = false;
	if (OUTPUT_VAR->mType != VAR_CLIPBOARD && mArgc > 1)
		// It has a second arg, which in this case is the value to be assigned to the var.
		// Examine any derefs that the second arg has to see if OUTPUT_VAR is mentioned:
		for (DerefType *deref = mArg[1].deref; deref && deref->marker; ++deref)
			if (deref->var == OUTPUT_VAR)
			{
				target_is_involved_in_source = true;
				break;
			}

	// Note: It might be possible to improve performance in the case where
	// the target variable is large enough to accommodate the new source data
	// by moving memory around inside it.  For example: Var1 = xxxxxVar1
	// could be handled by moving the memory in Var1 to make room to insert
	// the literal string.  In addition to being quicker than the ExpandArgs()
	// method, this approach would avoid the possibility of needing to expand the
	// deref buffer just to handle the operation.  However, if that is ever done,
	// be sure to check that OUTPUT_VAR is mentioned only once in the list of derefs.
	// For example, something like this would probably be much easier to
	// implement by using ExpandArgs(): Var1 = xxxx Var1 Var2 Var1 xxxx.
	// So the main thing to be possibly later improved here is the case where
	// OUTPUT_VAR is mentioned only once in the deref list:
	VarSizeType space_needed;
	if (target_is_involved_in_source)
	{
		if (ExpandArgs() != OK)
			return FAIL;
		// ARG2 now contains the dereferenced (literal) contents of the text we want to assign.
		space_needed = (VarSizeType)strlen(ARG2) + 1;  // +1 for the zero terminator.
	}
	else
		space_needed = GetExpandedArgSize(false); // There's at most one arg to expand in this case.

	// Now above has ensured that space_needed is at least 1 (it should not be zero because even
	// the empty string uses up 1 char for its zero terminator).  The below relies upon this fact.

	if (space_needed <= 1) // Variable is being assigned the empty string (or a deref that resolves to it).
		return OUTPUT_VAR->Assign("");  // If the var is of large capacity, this will also free its memory.

	if (target_is_involved_in_source)
		// It was already dereferenced above, so use ARG2, which points to the
		// derefed contents of ARG2 (i.e. the data to be assigned):
		return OUTPUT_VAR->Assign(ARG2, space_needed - 1, g_script.mIsAutoIt2);

	// Otherwise:
	// If we're here, OUTPUT_VAR->mType must be clipboard or normal because otherwise
	// the validation during load would have prevented the script from loading:

	// First set everything up for the operation.  If OUTPUT_VAR is the clipboard, this
	// will prepare the clipboard for writing:
	if (OUTPUT_VAR->Assign(NULL, space_needed - 1) != OK)
		return FAIL;
	// Expand Arg2 directly into the var.  Also set the length explicitly
	// in case actual size written was off from the esimated size, perhaps
	// due to a failure or size discrepancy between the deref size-estimate
	// and the actual deref itself.  Note: If OUTPUT_VAR is the clipboard,
	// it's probably okay if the below actually writes less than the size of
	// the mem that has already been allocated for the new clipboard contents
	// That might happen due to a failure or size discrepancy between the
	// deref size-estimate and the actual deref itself:
	OUTPUT_VAR->Length() = (VarSizeType)(ExpandArg(OUTPUT_VAR->Contents(), 1) - OUTPUT_VAR->Contents() - 1);
	if (g_script.mIsAutoIt2)
	{
		trim(OUTPUT_VAR->Contents());  // Since this is how AutoIt2 behaves.
		OUTPUT_VAR->Length() = (VarSizeType)strlen(OUTPUT_VAR->Contents());
	}
	return OUTPUT_VAR->Close();  // i.e. Consider this function to be always successful unless this fails.
}


///////////////////////////////////////////////////////////////////////////////
// FROM AUTOIT3
///////////////////////////////////////////////////////////////////////////////
// Util_Shutdown()
// Shutdown or logoff the system
//
// Returns false if the function could not get the rights to shutdown 
///////////////////////////////////////////////////////////////////////////////

bool Util_Shutdown(int nFlag)
{

/* 
flags can be a combination of:
#define EWX_LOGOFF           0
#define EWX_SHUTDOWN         0x00000001
#define EWX_REBOOT           0x00000002
#define EWX_FORCE            0x00000004
#define EWX_POWEROFF         0x00000008 */

	HANDLE				hToken; 
	TOKEN_PRIVILEGES	tkp; 

	// If we are running NT, make sure we have rights to shutdown
	if (g_os.IsWinNT())
	{
		// Get a token for this process.
 		if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken)) 
			return false;						// Don't have the rights
 
		// Get the LUID for the shutdown privilege.
 		LookupPrivilegeValue(NULL, SE_SHUTDOWN_NAME, &tkp.Privileges[0].Luid); 
 
		tkp.PrivilegeCount = 1;  /* one privilege to set */
		tkp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED; 
 
		// Get the shutdown privilege for this process.
 		AdjustTokenPrivileges(hToken, FALSE, &tkp, 0, (PTOKEN_PRIVILEGES)NULL, 0); 
 
		// Cannot test the return value of AdjustTokenPrivileges.
 		if (GetLastError() != ERROR_SUCCESS) 
			return false;						// Don't have the rights
	}

	// if we are forcing the issue, AND this is 95/98 terminate all windows first
	if ( g_os.IsWin9x() && (nFlag & EWX_FORCE) ) 
	{
		nFlag ^= EWX_FORCE;	// remove this flag - not valid in 95
		EnumWindows((WNDENUMPROC) Util_ShutdownHandler, 0);
	}

	// ExitWindows
	if (ExitWindowsEx(nFlag, 0))
		return true;
	else
		return false;

} // Util_Shutdown()



///////////////////////////////////////////////////////////////////////////////
// FROM AUTOIT3
///////////////////////////////////////////////////////////////////////////////
BOOL Util_ShutdownHandler(HWND hwnd, DWORD lParam)
{
	// if the window is me, don't terminate!
	if (hwnd != g_hWnd && hwnd != g_hWndSplash)
		Util_WinKill(hwnd);

	// Continue the enumeration.
	return TRUE;

} // Util_ShutdownHandler()



///////////////////////////////////////////////////////////////////////////////
// FROM AUTOIT3
///////////////////////////////////////////////////////////////////////////////
void Util_WinKill(HWND hWnd)
{
	DWORD      pid = 0;
	LRESULT    lResult;
	HANDLE     hProcess;
	DWORD      dwResult;

	lResult = SendMessageTimeout(hWnd, WM_CLOSE, 0, 0, SMTO_ABORTIFHUNG, 500, &dwResult);	// wait 500ms

	if( !lResult )
	{
		// Use more force - Mwuahaha

		// Get the ProcessId for this window.
		GetWindowThreadProcessId( hWnd, &pid );

		// Open the process with all access.
		hProcess = OpenProcess(PROCESS_ALL_ACCESS, FALSE, pid);

		// Terminate the process.
		TerminateProcess(hProcess, 0);

		CloseHandle(hProcess);
	}

} // Util_WinKill()



ResultType Line::FileSelectFile(char *aOptions, char *aWorkingDir)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	if (g_nFileDialogs >= MAX_FILEDIALOGS)
	{
		// Have a maximum to help prevent runaway hotkeys due to key-repeat feature, etc.
		MsgBox("The maximum number of File Dialogs has been reached." ERR_ABORT);
		return FAIL;
	}
	char file_buf[64 * 1024]; // Large in case more than one file is allowed to be selected.

	// Use a more specific title so that the dialogs of different scripts can be distinguished
	// from one another, which may help script automation in rare cases:
	char dialog_title[512];
	snprintf(dialog_title, sizeof(dialog_title), "Select File - %s", g_script.mFileName);

	OPENFILENAME ofn;
	// This init method is used in more than one example:
	ZeroMemory(&ofn, sizeof(ofn));
	ofn.lStructSize = sizeof(ofn);
	ofn.hwndOwner = NULL; // i.e. no need to have main window forced into the background for this.
	// Must be terminated by two NULL characters.  One is explicit, the other automatic:
	ofn.lpstrTitle = dialog_title;
	ofn.lpstrFilter = "All Files (*.*)\0*.*\0Text Documents (*.txt)\0*.txt\0";
	ofn.lpstrFile = file_buf;
	ofn.nMaxFile = sizeof(file_buf) - 1; // -1 to be extra safe, like AutoIt3.
	// Specifying NULL will make it default to the last used directory (at least in Win2k):
	ofn.lpstrInitialDir = (aWorkingDir && *aWorkingDir) ? aWorkingDir : NULL;

	int options = atoi(aOptions);
	ofn.Flags = OFN_HIDEREADONLY | OFN_EXPLORER | OFN_NODEREFERENCELINKS;
	if (options & 0x10)
		ofn.Flags |= OFN_OVERWRITEPROMPT;
	if (options & 0x08)
		ofn.Flags |= OFN_CREATEPROMPT;
	if (options & 0x04)
		ofn.Flags |= OFN_ALLOWMULTISELECT;
	if (options & 0x02)
		ofn.Flags |= OFN_PATHMUSTEXIST;
	if (options & 0x01)
		ofn.Flags |= OFN_FILEMUSTEXIST;

	// This will attempt to force it to the foreground after it has been displayed, since the
	// dialog often will flash in the task bar instead of becoming foreground.
	// See MsgBox() for details:
	PostMessage(g_hWnd, AHK_DIALOG, (WPARAM)0, (LPARAM)0); // Must pass 0 for WPARAM in this case.

	g.WaitingForDialog = true;
	++g_nFileDialogs;
	// Below: OFN_CREATEPROMPT doesn't seem to work with GetSaveFileName(), so always
	// use GetOpenFileName() in that case:
	BOOL result = (ofn.Flags & OFN_OVERWRITEPROMPT) && !(ofn.Flags & OFN_CREATEPROMPT)
		? GetSaveFileName(&ofn) : GetOpenFileName(&ofn);
	--g_nFileDialogs;
	g.WaitingForDialog = false;  // IsCycleComplete() relies on this.

	if (!result) // User pressed CANCEL vs. OK to dismiss the dialog.
		// It seems best to clear the variable in these cases, since this is a scripting
		// language where performance is not the primary goal.  So do that and return OK,
		// but leave ErrorLevel set to ERRORLEVEL_ERROR.
		return OUTPUT_VAR->Assign(); // Tell it not to free the memory by not calling with "".
	else
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate that the user pressed OK vs. CANCEL.

	if (ofn.Flags & OFN_ALLOWMULTISELECT)
	{
		// Replace all the zero terminators with a delimiter, except the one for the last file
		// (the last file should be followed by two sequential zero terminators).
		// Use a delimiter that can't be confused with a real character inside a filename, i.e.
		// not a comma.  We only have room for one without getting into the complexity of having
		// to expand the string, so \r\n is disqualified for now.
		for (char *cp = ofn.lpstrFile;;)
		{
			for (; *cp; ++cp); // Find the next terminator.
			*cp = '\n'; // Replace zero-delimiter with a visible/printable delimiter, for the user.
			if (!*(cp + 1)) // This is the last file because it's double-terminated, so we're done.
				break;
		}
	}
	return OUTPUT_VAR->Assign(ofn.lpstrFile);
}



ResultType Line::FileCreateDir(char *aDirSpec)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	if (!aDirSpec || !*aDirSpec)
		return OK;  // Return OK because g_ErrorLevel tells the story.

	DWORD attr = GetFileAttributes(aDirSpec);
	if (attr != 0xFFFFFFFF)  // aDirSpec already exists.
	{
		if ((attr & FILE_ATTRIBUTE_DIRECTORY))
			g_ErrorLevel->Assign(ERRORLEVEL_NONE);  // Indicate success since it already exists as a dir.
		// else leave as failure, since aDirSpec exists as a file, not a dir.
		return OK;
	}

	// If it has a backslash, make sure all its parent directories exist before we attempt
	// to create this directory:
	char *last_backslash = strrchr(aDirSpec, '\\');
	if (last_backslash)
	{
		char parent_dir[MAX_PATH * 2];
		strlcpy(parent_dir, aDirSpec, last_backslash - aDirSpec + 1); // Omits the last backslash.
		FileCreateDir(parent_dir); // Recursively create all needed ancestor directories.
		if (*g_ErrorLevel->Contents() == *ERRORLEVEL_ERROR)
			return OK; // Return OK because ERRORLEVEL_ERROR is the indicator of failure.
	}

	// The above has recursively created all parent directories of aDirSpec if needed.
	// Now we can create aDirSpec.  Be sure to explicitly set g_ErrorLevel since it's value
	// is now indeterminate due to action above:
	return g_ErrorLevel->Assign(CreateDirectory(aDirSpec, NULL) ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);
}



ResultType Line::FileReadLine(char *aFilespec, char *aLineNumber)
// Returns OK or FAIL.  Will almost always return OK because if an error occurs,
// the script's ErrorLevel variable will be set accordingly.  However, if some
// kind of unexpected and more serious error occurs, such as variable-out-of-memory,
// that will cause FAIL to be returned.
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	UINT line_number = atoi(aLineNumber);
	if (line_number <= 0)
		return OK;  // Return OK because g_ErrorLevel tells the story.
	FILE *fp = fopen(aFilespec, "r");
	if (!fp)
		return OK;  // Return OK because g_ErrorLevel tells the story.

	// If the keyboard or mouse hook is installed, pause periodically during potentially long
	// operations such as this one, to give the msg pump a chance to process keyboard and
	// mouse events so that they don't lag.
	// 10000 causes barely perceptible lag when moving mouse cursor on Athlon XP 3000+,
	// so 1000 should be good for most CPUs.  Note: Tried PeekMessage(PM_NOREMOVE),
	// with and without WaitMessage(), but it didn't work.  So it seems necessary
	// to actually get into the GetMessage() wait-state.  One possibly drawback to this,
	// though likely extremely rare, is that a hotkey may fire while we're in the middle
	// of reading a file.  If that hotkey doesn't return in a reasonable amount of time,
	// the file we're reading will stay open for as long as this subroutine is suspended.
	// Pretty darn rare, and arguably the correct behavior in any case, so doesn't seem
	// cause for concern.  UPDATE: It seems best to do MsgSleep() periodically (though less
	// often) even if the hook isn't installed, so that the program will still be responsive
	// (e.g. its tray menu and other hotkeys) while conducting a file operation that takes
	// a very long time:
	int line_interval, sleep_duration;
	if (Hotkey::HookIsActive())
	{
		line_interval = 1000;
		sleep_duration = 10;
	}
	else
	{
		line_interval = 10000;
		sleep_duration = -1; // Since all we want to do is check messages.
	}

	char buf[64 * 1024];
	for (UINT i = 0; i < line_number; ++i)
	{
		if (i && !(i % line_interval))
			MsgSleep(sleep_duration); // See above comment.
		if (fgets(buf, sizeof(buf) - 1, fp) == NULL) // end-of-file or error
		{
			fclose(fp);
			return OK;  // Return OK because g_ErrorLevel tells the story.
		}
	}
	fclose(fp);

	size_t buf_length = strlen(buf);
	if (buf_length && buf[buf_length - 1] == '\n') // Remove any trailing newline for the user.
		buf[--buf_length] = '\0';
	if (!buf_length)
	{
		if (!OUTPUT_VAR->Assign()) // Explicitly call it this way so that it won't free the memory.
			return FAIL;
	}
	else
		if (!OUTPUT_VAR->Assign(buf, (VarSizeType)buf_length))
			return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return OK;
}



ResultType Line::FileAppend(char *aFilespec, char *aBuf)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	if (!aFilespec || !*aFilespec)
		return OK;  // Return OK because g_ErrorLevel tells the story.
	FILE *fp = fopen(aFilespec, "a");
	if (!fp)
		return OK;  // Return OK because g_ErrorLevel tells the story.
	if (!fputs(aBuf, fp)) // Success.
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	fclose(fp);
	return OK;
}



ResultType Line::FileDelete(char *aFilePattern)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	if (!aFilePattern || !*aFilePattern)
		return OK;  // Let ErrorLevel indicate an error, since this is probably not what the user intended.

	if (!StrChrAny(aFilePattern, "?*"))
	{
		if (DeleteFile(aFilePattern))
			g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		return OK; // ErrorLevel will indicate failure if the above didn't succeed.
	}

	// Otherwise aFilePattern contains wildcards, so we'll search for all matches and delete them.
	char file_path[MAX_PATH * 2];  // Give extra room in case OS supports extra-long files?
	char target_filespec[MAX_PATH * 2];
	if (strlen(aFilePattern) >= sizeof(file_path))
		return OK; // Return OK because this is non-critical.  Let the above ErrorLevel indicate the problem.
	strlcpy(file_path, aFilePattern, sizeof(file_path));
	char *last_backslash = strrchr(file_path, '\\');
	if (last_backslash)
		*(last_backslash + 1) = '\0'; // Leave the trailing backslash on it for consistency with below.
	else // Use current working directory, e.g. if user specified only *.*
		*file_path = '\0';

	WIN32_FIND_DATA current_file;
	HANDLE file_search = FindFirstFile(aFilePattern, &current_file);
	bool file_found = (file_search != INVALID_HANDLE_VALUE);
	int failure_count = 0;

	for (; file_found; file_found = FindNextFile(file_search, &current_file))
	{
		if (current_file.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) // skip any matching directories.
			continue;
		snprintf(target_filespec, sizeof(target_filespec), "%s%s", file_path, current_file.cFileName);
		if (!DeleteFile(target_filespec))
			++failure_count;
	}

	if (file_search != INVALID_HANDLE_VALUE) // In case the loop had zero iterations.
		FindClose(file_search);
	if (!failure_count)
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return OK;
}



ResultType Line::FileMove(char *aSource, char *aDest, char *aFlag)
// Adapted from the AutoIt3 source.
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	bool bOverwrite = aFlag && *aFlag == '1';
	if (MoveFile(aSource, aDest))
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return OK;
}


ResultType Line::FileCopy(char *aSource, char *aDest, char *aFlag)
// Adapted from the AutoIt3 source.
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	bool bOverwrite = aFlag && *aFlag == '1';
	if (Util_CopyFile(aSource, aDest, bOverwrite))
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return OK;
}



///////////////////////////////////////////////////////////////////////////////
// TAKEN FROM THE AUTOIT3 SOURCE
///////////////////////////////////////////////////////////////////////////////
// Util_CopyFile()
// Returns true if all files copied, else returns false
///////////////////////////////////////////////////////////////////////////////

bool Line::Util_CopyFile(const char *szInputSource, const char *szInputDest, bool bOverwrite)
{
	WIN32_FIND_DATA	findData;
	HANDLE			hSearch;
	bool			bLoop;
	bool			bDestExists;

	char			szSource[_MAX_PATH+1];
	char			szDest[_MAX_PATH+1];
	char			szExpandedDest[MAX_PATH+1];
	char			szTempPath[_MAX_PATH+1];

	char			szDrive[_MAX_PATH+1];
	char			szDir[_MAX_PATH+1];
	char			szFile[_MAX_PATH+1];
	char			szExt[_MAX_PATH+1];


	// Get local version of our source/dest
	strcpy(szSource, szInputSource);
	strcpy(szDest, szInputDest);

	// Split dest into file and extension
	_splitpath( szInputDest, szDrive, szDir, szFile, szExt );

	// If the filename and extension are both blank, sub with *.*
	if (szFile[0] == '\0' && szFile[0] == '\0')
	{
		strcpy(szFile, "*");
		strcpy(szExt,".*");
	}
	
	strcpy(szDest, szDrive);	strcat(szDest, szDir);
	strcat(szDest, szFile);	strcat(szDest, szExt);


	// Split source into file and extension
	_splitpath( szSource, szDrive, szDir, szFile, szExt );

	// If the filename and extension are both blank, sub with *.*
	if (szFile[0] == '\0' && szFile[0] == '\0')
	{
		strcpy(szFile, "*");
		strcpy(szExt,".*");
	}
	
	strcpy(szSource, szDrive);	strcat(szSource, szDir);
	strcat(szSource, szFile);	strcat(szSource, szExt);

	// Note we now rely on the SOURCE being the contents of SZDir, szFile, etc.


	// Does the source file exist?
	hSearch = FindFirstFile(szSource, &findData);
	bLoop = true;
	while (hSearch != INVALID_HANDLE_VALUE && bLoop == true)
	{
		// Make sure the returned handle is a file and not a directory before we
		// try and do copy type things on it!
		if ( (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != FILE_ATTRIBUTE_DIRECTORY)
		{
			// Expand the destination based on this found file
			Util_ExpandFilenameWildcard(findData.cFileName, szDest, szExpandedDest);
		
			// The find struct only returns the file NAME, we need to reconstruct the path!
			strcpy(szTempPath, szDrive);	
			strcat(szTempPath, szDir);
			strcat(szTempPath, findData.cFileName);

			// Does the destination exist?
			bDestExists = Util_DoesFileExist(szExpandedDest);

			// Copy the file - maybe
			if ( (bDestExists == true  && bOverwrite == true) || (bDestExists == false) )
			{
				if ( CopyFile(szTempPath, szExpandedDest, FALSE) != TRUE )
				{
					FindClose(hSearch);
					return false;						// Error copying one of the files
				}
			}
		} // End If

		if (FindNextFile(hSearch, &findData) == FALSE)
			bLoop = false;

	} // End while

	FindClose(hSearch);

	return true;

} // Util_CopyFile()


///////////////////////////////////////////////////////////////////////////////
// TAKEN FROM THE AUTOIT3 SOURCE
///////////////////////////////////////////////////////////////////////////////
// Util_ExpandFilenameWildcard()
///////////////////////////////////////////////////////////////////////////////

void Line::Util_ExpandFilenameWildcard(const char *szSource, const char *szDest, char *szExpandedDest)
{
	// copy one.two.three  *.txt     = one.two   .txt
	// copy one.two.three  *.*.txt   = one.two   .three  .txt
	// copy one.two.three  *.*.*.txt = one.two   .three  ..txt
	// copy one.two		   test      = test

	char	szFileTemp[_MAX_PATH+1];
	char	szExtTemp[_MAX_PATH+1];

	char	szSrcFile[_MAX_PATH+1];
	char	szSrcExt[_MAX_PATH+1];

	char	szDestDrive[_MAX_PATH+1];
	char	szDestDir[_MAX_PATH+1];
	char	szDestFile[_MAX_PATH+1];
	char	szDestExt[_MAX_PATH+1];

	// If the destination doesn't include a wildcard, send it back vertabim
	if (strchr(szDest, '*') == NULL)
	{
		strcpy(szExpandedDest, szDest);
		return;
	}

	// Split source and dest into file and extension
	_splitpath( szSource, szDestDrive, szDestDir, szSrcFile, szSrcExt );
	_splitpath( szDest, szDestDrive, szDestDir, szDestFile, szDestExt );

	// Source and Dest ext will either be ".nnnn" or "" or ".*", remove the period
	if (szSrcExt[0] == '.')
		strcpy(szSrcExt, &szSrcExt[1]);
	if (szDestExt[0] == '.')
		strcpy(szDestExt, &szDestExt[1]);

	// Start of the destination with the drive and dir
	strcpy(szExpandedDest, szDestDrive);
	strcat(szExpandedDest, szDestDir);

	// Replace first * in the destext with the srcext, remove any other *
	Util_ExpandFilenameWildcardPart(szSrcExt, szDestExt, szExtTemp);

	// Replace first * in the destfile with the srcfile, remove any other *
	Util_ExpandFilenameWildcardPart(szSrcFile, szDestFile, szFileTemp);

	// Concat the filename and extension if req
	if (szExtTemp[0] != '\0')
	{
		strcat(szFileTemp, ".");
		strcat(szFileTemp, szExtTemp);	
	}
	else
	{
		// Dest extension was blank SOURCE MIGHT NOT HAVE BEEN!
		if (szSrcExt[0] != '\0')
		{
			strcat(szFileTemp, ".");
			strcat(szFileTemp, szSrcExt);	
		}
	}

	// Now add the drive and directory bit back onto the dest
	strcat(szExpandedDest, szFileTemp);

} // Util_CopyFile



///////////////////////////////////////////////////////////////////////////////
// TAKEN FROM THE AUTOIT3 SOURCE
///////////////////////////////////////////////////////////////////////////////
// Util_ExpandFilenameWildcardPart()
///////////////////////////////////////////////////////////////////////////////

void Line::Util_ExpandFilenameWildcardPart(const char *szSource, const char *szDest, char *szExpandedDest)
{
	char	*lpTemp;
	int		i, j, k;

	// Replace first * in the dest with the src, remove any other *
	i = 0; j = 0; k = 0;
	lpTemp = strchr(szDest, '*');
	if (lpTemp != NULL)
	{
		// Contains at least one *, copy up to this point
		while(szDest[i] != '*')
			szExpandedDest[j++] = szDest[i++];
		// Skip the * and replace in the dest with the srcext
		while(szSource[k] != '\0')
			szExpandedDest[j++] = szSource[k++];
		// Skip any other *
		i++;
		while(szDest[i] != '\0')
		{
			if (szDest[i] == '*')
				i++;
			else
				szExpandedDest[j++] = szDest[i++];
		}
		szExpandedDest[j] = '\0';
	}
	else
	{
		// No wildcard, straight copy of destext
		strcpy(szExpandedDest, szDest);
	}

} // Util_ExpandFilenameWildcardPart()



///////////////////////////////////////////////////////////////////////////////
// TAKEN FROM THE AUTOIT3 SOURCE
///////////////////////////////////////////////////////////////////////////////
// Util_DoesFileExist()
// Returns true if file or directory exists
///////////////////////////////////////////////////////////////////////////////

bool Line::Util_DoesFileExist(const char *szFilename)
{
	if ( strchr(szFilename,'*')||strchr(szFilename,'?') )
	{
		WIN32_FIND_DATA	wfd;
		HANDLE			hFile;

	    hFile = FindFirstFile(szFilename, &wfd);

		if ( hFile == INVALID_HANDLE_VALUE )
			return false;

		FindClose(hFile);
		return true;
	}
    else
	{
		DWORD dwTemp;

		dwTemp = GetFileAttributes(szFilename);
		if ( dwTemp != 0xffffffff )
			return true;
		else
			return false;
	}

} // Util_DoesFileExist



ResultType Line::SetToggleState(vk_type aVK, ToggleValueType &ForceLock, char *aToggleText)
// Caller must have already validated that the args are correct.
// Always returns OK, for use as caller's return value.
{
	ToggleValueType toggle = ConvertOnOffAlways(aToggleText, NEUTRAL);
	switch (toggle)
	{
	case TOGGLED_ON:
	case TOGGLED_OFF:
		// Turning it on or off overrides any prior AlwaysOn or AlwaysOff setting.
		// Probably need to change the setting BEFORE attempting to toggle the
		// key state, otherwise the hook may prevent the state from being changed
		// if it was set to be AlwaysOn or AlwaysOff:
		ForceLock = NEUTRAL;
		ToggleKeyState(aVK, toggle);
		break;
	case ALWAYS_ON:
	case ALWAYS_OFF:
		ForceLock = (toggle == ALWAYS_ON) ? TOGGLED_ON : TOGGLED_OFF; // Must do this first.
		ToggleKeyState(aVK, ForceLock);
		// This will ensure that the hook is installed if it isn't already.  The hook is
		// currently needed to support keeping these keys AlwaysOn or AlwaysOff, though
		// there may be better ways to do it (such as registering them as a hotkey, but
		// that may introduce quite a bit of complexity):
		Hotkey::AllActivate();
		break;
	case NEUTRAL:
		// Note: No attempt is made to detect whether the keybd hook should be deinstalled
		// because it's no longer needed due to this change.  That would require some 
		// careful thought about the impact on the status variables in the Hotkey class, etc.,
		// so it can be left for a future enhancement:
		ForceLock = NEUTRAL;
		break;
	}
	return OK;
}



////////////////////////////////
// Misc lower level functions //
////////////////////////////////

ArgPurposeType Line::ArgIsVar(ActionTypeType aActionType, int aArgIndex)
{
	switch(aArgIndex)
	{
	case 0:  // Arg #1
		switch(aActionType)
		{
		case ACT_ASSIGN:
		case ACT_ADD:
		case ACT_SUB:
		case ACT_MULT:
		case ACT_DIV:
		case ACT_STRINGLEFT:
		case ACT_STRINGRIGHT:
		case ACT_STRINGMID:
		case ACT_STRINGTRIMLEFT:
		case ACT_STRINGTRIMRIGHT:
		case ACT_STRINGLEN:
		case ACT_STRINGREPLACE:
		case ACT_STRINGGETPOS:
		case ACT_CONTROLGETTEXT:
		case ACT_STATUSBARGETTEXT:
		case ACT_INPUTBOX:
		case ACT_RANDOM:
		case ACT_REGREAD:
		case ACT_FILEREADLINE:
		case ACT_FILESELECTFILE:
		case ACT_MOUSEGETPOS:
		case ACT_WINGETTITLE:
		case ACT_WINGETTEXT:
		case ACT_WINGETPOS:
			return (ArgPurposeType)IS_OUTPUT_VAR;

		case ACT_IFINSTRING:
		case ACT_IFNOTINSTRING:
		case ACT_IFEQUAL:
		case ACT_IFNOTEQUAL:
		case ACT_IFGREATER:
		case ACT_IFGREATEROREQUAL:
		case ACT_IFLESS:
		case ACT_IFLESSOREQUAL:
			return (ArgPurposeType)IS_INPUT_VAR;
		}
		break;

	case 1:  // Arg #2
		switch(aActionType)
		{
		case ACT_STRINGLEFT:
		case ACT_STRINGRIGHT:
		case ACT_STRINGMID:
		case ACT_STRINGTRIMLEFT:
		case ACT_STRINGTRIMRIGHT:
		case ACT_STRINGLEN:
		case ACT_STRINGREPLACE:
		case ACT_STRINGGETPOS:
			return (ArgPurposeType)IS_INPUT_VAR;

		case ACT_MOUSEGETPOS:
		case ACT_WINGETPOS:
			return (ArgPurposeType)IS_OUTPUT_VAR;
		}
		break;

	case 2:  // Arg #3
		if (aActionType == ACT_WINGETPOS)
			return (ArgPurposeType)IS_OUTPUT_VAR;
		break;

	case 3:  // Arg #4
		if (aActionType == ACT_WINGETPOS)
			return (ArgPurposeType)IS_OUTPUT_VAR;
		break;
	}
	// Otherwise:
	return IS_NOT_A_VAR;
}



ResultType Line::CheckForMandatoryArgs()
{
	switch(mActionType)
	{
	// For these, although we validate that at least one is non-blank here, it's okay at
	// runtime for them all to resolve to be blank, without an error being reported.
	// It's probably more flexible that way since the commands are equipped to handle
	// all-blank params.
	// Not these because they can be used with the "last-used window" mode:
	//case ACT_IFWINEXIST:
	//case ACT_IFWINNOTEXIST:
	case ACT_WINACTIVATEBOTTOM:
		if (!*RAW_ARG1 && !*RAW_ARG2 && !*RAW_ARG3 && !*RAW_ARG4)
			return LineError(ERR_WINDOW_PARAM);
		return OK;
	// Not these because they can have their window params all-blank to work in "last-used window" mode:
	//case ACT_IFWINACTIVE:
	//case ACT_IFWINNOTACTIVE:
	//case ACT_WINACTIVATE:
	//case ACT_WINWAITCLOSE:
	//case ACT_WINWAITACTIVE:
	//case ACT_WINWAITNOTACTIVE:
	case ACT_WINWAIT:
		if (!*RAW_ARG1 && !*RAW_ARG2 && !*RAW_ARG4 && !*RAW_ARG5) // ARG3 is omitted because it's the timeout.
			return LineError(ERR_WINDOW_PARAM);
		return OK;
	// Note: For ACT_WINMOVE, don't validate anything for mandatory args so that its two modes of
	// operation can be supported: 2-param mode and normal-param mode.
	case ACT_GROUPADD:
		if (!*RAW_ARG2 && !*RAW_ARG3 && !*RAW_ARG5 && !*RAW_ARG6) // ARG4 is the JumpToLine
			return LineError(ERR_WINDOW_PARAM);
		return OK;
	case ACT_CONTROLSEND:
		// Window params can all be blank in this case, but characters to send should
		// be non-blank (but it's ok if its a dereferenced var that resolves to blank
		// at runtime):
		if (!*RAW_ARG2)
			return LineError("Parameter #2 must not be blank.");
		return OK;
	case ACT_MOUSECLICKDRAG:
		// Even though we check for blanks at load-time, we don't bother to do so at runtime
		// (i.e. if a dereferenced var resolved to blank, it will be treated as a zero):
		if (!*RAW_ARG4 || !*RAW_ARG5)
			return LineError("Parameters 4 and 5 must specify a non-blank destination for the drag.");
		return OK;
	case ACT_MOUSEGETPOS:
		if (!VARRAW_ARG1 && !VARRAW_ARG2)
			return LineError(ERR_MISSING_OUTPUT_VAR);
		return OK;
	case ACT_WINGETPOS:
		if (!VARRAW_ARG1 && !VARRAW_ARG2 && !VARRAW_ARG3 && !VARRAW_ARG4)
			return LineError(ERR_MISSING_OUTPUT_VAR);
		return OK;
	}
	return OK;  // For when the command isn't mentioned in the switch().
}



bool Line::FileIsFilteredOut(WIN32_FIND_DATA &aCurrentFile, FileLoopModeType aFileLoopMode
	, char *aFilePath)
{
	if (   (!strcmp(aCurrentFile.cFileName, "..") || !strcmp(aCurrentFile.cFileName, "."))
		&& !(aFileLoopMode & FILE_LOOP_INCLUDE_SELF_AND_PARENT)   )
		return true;
	if (   (aCurrentFile.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)  // it is a folder.
		&& !(aFileLoopMode & (FILE_LOOP_INCLUDE_FOLDERS | FILE_LOOP_INCLUDE_FOLDERS_ONLY))   )
		return true;
	if (   !(aCurrentFile.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) // not a folder.
		&& (aFileLoopMode & FILE_LOOP_INCLUDE_FOLDERS_ONLY)   )
		return true;

	// Since file was found, also prepend the file's path to its name for the caller:
	if (!aFilePath || !*aFilePath) // don't bother.
		return false;

	char *last_backslash = strrchr(aFilePath, '\\');
	if (!last_backslash) // probably because a file search in the current dir, such as "*.*" was specified.
		return false;  // No need to prepend the path.
	size_t path_length = last_backslash - aFilePath + 1; // Exclude the wildcard part from the length.
	size_t filename_length = strlen(aCurrentFile.cFileName);
	if (filename_length + path_length >= MAX_PATH) // >= to allow room for the string terminator?
	{
		// This function isn't set up to cause a true FAIL condition, so just warn:
		LineError("When this filename's path is prepended, the result is too long.", WARN
			, aCurrentFile.cFileName);
		return true;  // Since we can't construct the full spec, tell it that this file was filtered after all.
	}
	// It's done this way to save stack space, since the recursion can get pretty deep.
	// Uses +1 to include the string's terminator:
	MoveMemory(aCurrentFile.cFileName + path_length, aCurrentFile.cFileName, filename_length + 1);
	MoveMemory(aCurrentFile.cFileName, aFilePath, path_length);
	return false;  // i.e. this file has not been filtered out.
}



ResultType Line::SetJumpTarget(bool aIsDereferenced)
{
	Label *label = g_script.FindLabel(aIsDereferenced ? ARG1 : RAW_ARG1);
	if (!label)
		if (aIsDereferenced)
			return LineError("This Goto/Gosub's target label does not exist.");
		else
			return LineError("This Goto/Gosub's target label does not exist." ERR_ABORT);
	mRelatedLine = label->mJumpToLine; // The script loader has ensured that this can't be NULL.
	// Seems best to do this even for GOSUBs even though it's a bit weird:
	return IsJumpValid(label->mJumpToLine);
	// Any error msg was already displayed by the above call.
}



ResultType Line::IsJumpValid(Line *aDestination)
// The caller has ensured that aDestination is not NULL.
// The caller relies on this function returning either OK or FAIL.
{
	// aDestination can be NULL if this Goto's target is the physical end of the script.
	// And such a destination is always valid, regardless of where aOrigin is.
	// UPDATE: It's no longer possible for the destination of a Goto or Gosub to be
	// NULL because the script loader has ensured that the end of the script always has
	// an extra ACT_EXIT that serves as an anchor for any final labels in the script:
	//if (aDestination == NULL)
	//	return OK;
	// The above check is also necessary to avoid dereferencing a NULL pointer below.

	// A Goto/Gosub can always jump to a point anywhere in the outermost layer
	// (i.e. outside all blocks) without restriction:
	if (aDestination->mParentLine == NULL)
		return OK;

	// So now we know this Goto/Gosub is attempting to jump into a block somewhere.  Is that
	// block a legal place to jump?:

	for (Line *ancestor = mParentLine; ancestor != NULL; ancestor = ancestor->mParentLine)
		if (aDestination->mParentLine == ancestor)
			// Since aDestination is in the same block as the Goto line itself (or a block
			// that encloses that block), it's allowed:
			return OK;
	// This can happen if the Goto's target is at a deeper level than it, or if the target
	// is at a more shallow level but is in some block totally unrelated to it!
	// Returns FAIL by default, which is what we want because that value is zero:
	return LineError("A Goto/Gosub/GroupActivate mustn't jump into a block that doesn't enclose it.");
	// Above currently doesn't attempt to detect runtime vs. load-time for the purpose of appending
	// ERR_ABORT.
}
