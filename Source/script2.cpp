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
#include <shlobj.h>  // for SHGetMalloc()
#include "script.h"
#include "window.h" // for IF_USE_FOREGROUND_WINDOW
#include "application.h" // for MsgSleep()
#include "resources\resource.h"  // For InputBox.


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



ResultType Line::PerformShowWindow(ActionTypeType aActionType, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
{
	// By design, the WinShow command must always unhide a hidden window, even if the user has
	// specified that hidden windows should not be detected.  So set this now so that the
	// DETERMINE_TARGET_WINDOW macro will make its calls in the right mode:
	bool need_restore = (aActionType == ACT_WINSHOW && !g.DetectHiddenWindows);
	if (need_restore)
		g.DetectHiddenWindows = true;
	DETERMINE_TARGET_WINDOW
	if (need_restore)
		g.DetectHiddenWindows = false;
	if (!target_window)
		return OK;

	int nCmdShow;
	switch (aActionType)
	{
		// SW_FORCEMINIMIZE: supported only in Windows 2000/XP and beyond: "Minimizes a window,
		// even if the thread that owns the window is hung. This flag should only be used when
		// minimizing windows from a different thread."
		// My: It seems best to use SW_FORCEMINIMIZE on OS's that support it because I have
		// observed ShowWindow() to hang (thus locking up our app's main thread) if the target
		// window is hung.
		// UPDATE: For now, not using "force" every time because it has undesirable side-effects such
		// as the window not being restored to its maximized state after it was minimized
		// this way.
		// Note: The use of IsHungAppWindow() (supported under Win2k+) is discouraged by MS,
		// so we won't use it here even though it probably performs much better.
		#define SW_INVALID -1
		case ACT_WINMINIMIZE:
			if (g_os.IsWin2000orLater())
				nCmdShow = IsWindowHung(target_window) ? SW_FORCEMINIMIZE : SW_MINIMIZE;
			else
				// If it's not Win2k or later, don't attempt to minimize hung windows because I
				// have an 80% expectation (i.e. untested) that our thread would hang because
				// the call to ShowWindow() would never return.  I have confirmed that SW_MINIMIZE
				// can lock up our thread on WinXP, which is why we revert to SW_FORCEMINIMIZE
				// above.
				nCmdShow = IsWindowHung(target_window) ? SW_INVALID : SW_MINIMIZE;
			break;
		case ACT_WINMAXIMIZE: nCmdShow = IsWindowHung(target_window) ? SW_INVALID : SW_MAXIMIZE; break;
		case ACT_WINRESTORE:  nCmdShow = IsWindowHung(target_window) ? SW_INVALID : SW_RESTORE;  break;
		// Seems safe to assume it's not hung in these cases, since I'm inclined to believe
		// (untested) that hiding and showing a hung window won't lock up our thread, and
		// there's a chance they may be effective even against hung windows, unlike the
		// others above (except ACT_WINMINIMIZE, which has a special FORCE method):
		case ACT_WINHIDE: nCmdShow = SW_HIDE; break;
		case ACT_WINSHOW: nCmdShow = SW_SHOW; break;
	}
	// UPDATE:  Trying ShowWindowAsync()
	// now, which should avoid the problems with hanging.  UPDATE #2: Went back to
	// not using Async() because sometimes the script lines that come after the one
	// that is doing this action here rely on this action having been completed
	// (e.g. a window being maximized prior to clicking somewhere inside it).
	if (nCmdShow != SW_INVALID)
	{
		// I'm not certain that SW_FORCEMINIMIZE works with ShowWindowAsync(), but
		// it probably does since there's absolutely no mention to the contrary
		// anywhere on MS's site or on the web.  But clearly, if it does work, it
		// does so only because Async() doesn't really post the message to the thread's
		// queue, instead opting for more agressive measures.  Thus, it seems best
		// to do it this way to have maximum confidence in it:
		//if (nCmdShow == SW_FORCEMINIMIZE) // Safer not to use ShowWindowAsync() in this case.
			ShowWindow(target_window, nCmdShow);
		//else
		//	ShowWindowAsync(target_window, nCmdShow);
//PostMessage(target_window, WM_SYSCOMMAND, SC_MINIMIZE, 0);
		DoWinDelay;
	}
	return OK;  // Return success for all the above cases.
}



ResultType Line::WinMove(char *aTitle, char *aText, char *aX, char *aY
	, char *aWidth, char *aHeight, char *aExcludeTitle, char *aExcludeText)
{
	// So that compatibility is retained, don't set ErrorLevel for commands that are native to AutoIt2
	// but that AutoIt2 doesn't use ErrorLevel with (such as this one).
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



ResultType Line::WinMenuSelectItem(char *aTitle, char *aText, char *aMenu1, char *aMenu2
	, char *aMenu3, char *aMenu4, char *aMenu5, char *aMenu6, char *aMenu7
	, char *aExcludeTitle, char *aExcludeText)
// Adapted from the AutoIt3 source.
{
	// Set up a temporary array make it easier to traverse nested menus & submenus
	// in a loop.  Also add a NULL at the end to simplify the loop a little:
	char *menu_param[] = {aMenu1, aMenu2, aMenu3, aMenu4, aMenu5, aMenu6, aMenu7, NULL};

	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return OK;  // Let ErrorLevel tell the story.

	HMENU hMenu = GetMenu(target_window);
	if (!hMenu) // Window has no menu bar.
		return OK;  // Let ErrorLevel tell the story.

	int menu_item_count = GetMenuItemCount(hMenu);
	if (menu_item_count <= 0) // Menu bar has no menus.
		return OK;  // Let ErrorLevel tell the story.
	
	#define MENU_ITEM_IS_SUBMENU 0xFFFFFFFF
	UINT menu_id = MENU_ITEM_IS_SUBMENU;
	char menu_text[1024];
	bool match_found;
	size_t menu_param_length;
	int pos, target_menu_pos;
	for (int i = 0; menu_param[i] && *menu_param[i]; ++i)
	{
		if (!hMenu)  // The nesting of submenus ended prior to the end of the list of menu search terms.
			return OK;  // Let ErrorLevel tell the story.
		menu_param_length = strlen(menu_param[i]);
		target_menu_pos = (menu_param[i][menu_param_length - 1] == '&') ? atoi(menu_param[i]) - 1 : -1;
		if (target_menu_pos >= 0)
		{
			if (target_menu_pos >= menu_item_count)  // Invalid menu position (doesn't exist).
				return OK;  // Let ErrorLevel tell the story.
			#define UPDATE_MENU_VARS(menu_pos) \
			menu_id = GetMenuItemID(hMenu, menu_pos);\
			if (menu_id == MENU_ITEM_IS_SUBMENU)\
				menu_item_count = GetMenuItemCount(hMenu = GetSubMenu(hMenu, menu_pos));\
			else\
			{\
				menu_item_count = 0;\
				hMenu = NULL;\
			}
			UPDATE_MENU_VARS(target_menu_pos)
		}
		else // Searching by text rather than numerical position.
		{
			for (match_found = false, pos = 0; pos < menu_item_count; ++pos)
			{
				GetMenuString(hMenu, pos, menu_text, sizeof(menu_text) - 1, MF_BYPOSITION);
				match_found = !strnicmp(menu_text, menu_param[i], strlen(menu_param[i]));
				//match_found = stristr(menu_text, menu_param[i]);
				if (!match_found)
				{
					// Try again to find a match, this time without the ampersands used to indicate
					// a menu item's shortcut key:
					StrReplace(menu_text, "&", "");
					match_found = !strnicmp(menu_text, menu_param[i], strlen(menu_param[i]));
					//match_found = stristr(menu_text, menu_param[i]);
				}
				if (match_found)
				{
					UPDATE_MENU_VARS(pos)
					break;
				}
			} // inner for()
			if (!match_found) // The search hierarchy (nested menus) specified in the params could not be found.
				return OK;  // Let ErrorLevel tell the story.
		} // else
	} // outer for()

	// This would happen if the outer loop above had zero iterations due to aMenu1 being NULL or blank,
	// or if the caller specified a submenu as the target (which doesn't seem valid since an app would
	// next expect to ever receive a message for a submenu?):
	if (menu_id == MENU_ITEM_IS_SUBMENU)
		return OK;  // Let ErrorLevel tell the story.

	// Since the above didn't return, the specified search hierarchy was completely found.
	PostMessage(target_window, WM_COMMAND, (WPARAM)menu_id, 0);
	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return OK;
}



ResultType Line::ControlSend(char *aControl, char *aKeysToSend, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText, modLR_type aModifiersLR)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return OK;
	HWND control_window = ControlExist(target_window, aControl);
	if (!control_window)
		return OK;
	SendKeys(aKeysToSend, aModifiersLR, control_window);
	// But don't do WinDelay because KeyDelay should have been in effect for the above.
	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return OK;
}



ResultType Line::ControlLeftClick(char *aControl, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return OK;
	HWND control_window = ControlExist(target_window, aControl);
	if (!control_window)
		return OK;
	PostMessage(control_window, WM_LBUTTONDOWN, MK_LBUTTON, 0);
	PostMessage(control_window, WM_LBUTTONUP, 0, 0);
	DoControlDelay;  // It seems safer and more flexible to do this even for Control commands.
	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return OK;
}



ResultType Line::ControlGetFocus(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	output_var->Assign();  // Set default: blank for the output variable.
	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return OK;  // Let ErrorLevel and the blank output variable tell the story.

	// Unlike many of the other Control commands, this one requires AttachThreadInput():
	bool is_attached_my_to_fore = false, is_attached_fore_to_target = false;
	DWORD fore_thread, my_thread, target_thread;
	fore_thread = GetWindowThreadProcessId(GetForegroundWindow(), NULL);
	my_thread  = GetCurrentThreadId();
	target_thread = GetWindowThreadProcessId(target_window, NULL);
	if (fore_thread && my_thread != fore_thread)
		is_attached_my_to_fore = AttachThreadInput(my_thread, fore_thread, TRUE) != 0;
	if (fore_thread && target_thread && fore_thread != target_thread)
		is_attached_fore_to_target = AttachThreadInput(fore_thread, target_thread, TRUE) != 0;

	class_and_hwnd_type cah;
	cah.hwnd = GetFocus();  // Do this now that our thread is attached to the target window's.

	// Very important to detach any threads whose inputs were attached above,
	// prior to returning, otherwise the next attempt to attach thread inputs
	// for these particular windows may result in a hung thread or other
	// undesirable effect:
	if (is_attached_my_to_fore)
		AttachThreadInput(my_thread, fore_thread, FALSE);
	if (is_attached_fore_to_target)
		AttachThreadInput(fore_thread, target_thread, FALSE);

	if (!cah.hwnd)
		return OK;  // Let ErrorLevel and the blank output variable tell the story.

	char class_name[1024];
	cah.class_name = class_name;
	if (!GetClassName(cah.hwnd, class_name, sizeof(class_name)))
		return OK;  // Let ErrorLevel and the blank output variable tell the story.
	
	cah.class_count = 0;  // Init for the below.
	cah.is_found = false; // Same.
	EnumChildWindows(target_window, EnumChildFocusFind, (LPARAM)&cah);
	if (!cah.is_found)
		return OK;  // Let ErrorLevel and the blank output variable tell the story.
	// Append the class sequence number onto the class name set the output param to be that value:
	size_t class_name_length = strlen(class_name);
	snprintf(class_name + class_name_length, sizeof(class_name) - class_name_length - 1, "%d", cah.class_count);
	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return output_var->Assign(class_name);
}



BOOL CALLBACK EnumChildFocusFind(HWND aWnd, LPARAM lParam)
{
	#define cah ((class_and_hwnd_type *)lParam)
	char class_name[1024];
	if (!GetClassName(aWnd, class_name, sizeof(class_name)))
		return TRUE;  // Continue the enumeration.
	if (!strcmp(class_name, cah->class_name)) // Class names match.
	{
		++cah->class_count;
		if (aWnd == cah->hwnd)  // The caller-specified window has been found.
		{
			cah->is_found = true;
			return FALSE;
		}
	}
	return TRUE; // Continue enumeration until a match is found or there aren't any windows remaining.
}



ResultType Line::ControlFocus(char *aControl, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return OK;
	HWND control_window = ControlExist(target_window, aControl);
	if (!control_window)
		return OK;

	// Unlike many of the other Control commands, this one requires AttachThreadInput()
	// to have any realistic chance of success (though sometimes it may work by pure
	// chance even without it):
	bool is_attached_my_to_fore = false, is_attached_fore_to_target = false;
	DWORD fore_thread, my_thread, target_thread;
	fore_thread = GetWindowThreadProcessId(GetForegroundWindow(), NULL);
	my_thread  = GetCurrentThreadId();
	target_thread = GetWindowThreadProcessId(target_window, NULL);
	if (fore_thread && my_thread != fore_thread)
		is_attached_my_to_fore = AttachThreadInput(my_thread, fore_thread, TRUE) != 0;
	if (fore_thread && target_thread && fore_thread != target_thread)
		is_attached_fore_to_target = AttachThreadInput(fore_thread, target_thread, TRUE) != 0;

	if (SetFocus(control_window))
	{
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		DoControlDelay;  // It seems safer and more flexible to do this even for Control commands.
	}

	// Very important to detach any threads whose inputs were attached above,
	// prior to returning, otherwise the next attempt to attach thread inputs
	// for these particular windows may result in a hung thread or other
	// undesirable effect:
	if (is_attached_my_to_fore)
		AttachThreadInput(my_thread, fore_thread, FALSE);
	if (is_attached_fore_to_target)
		AttachThreadInput(fore_thread, target_thread, FALSE);

	return OK;
}



ResultType Line::ControlSetText(char *aControl, char *aNewText, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return OK;
	HWND control_window = ControlExist(target_window, aControl);
	if (!control_window)
		return OK;
	// SendMessage must be used, not PostMessage(), at least for some (probably most) apps.
	// Also: No need to call IsWindowHung() because SendMessageTimeout() should return
	// immediately if the OS already "knows" the window is hung:
	DWORD result;
	SendMessageTimeout(control_window, WM_SETTEXT, (WPARAM)0, (LPARAM)aNewText
		, SMTO_ABORTIFHUNG, 5000, &result);
	DoControlDelay;  // It seems safer and more flexible to do this even for Control commands.
	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return OK;
}



ResultType Line::ControlGetText(char *aControl, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	DETERMINE_TARGET_WINDOW
	HWND control_window = target_window ? ControlExist(target_window, aControl) : NULL;
	// Even if control_window is NULL, we want to continue on so that the output
	// param is set to be the empty string, which is the proper thing to do
	// rather than leaving whatever was in there before.

	// Handle the output parameter.  This section is similar to that in
	// PerformAssign().  Note: Using GetWindowTextTimeout() vs. GetWindowText()
	// because it is able to get text from more types of controls (e.g. large edit controls):
	VarSizeType space_needed = control_window ? GetWindowTextTimeout(control_window) + 1 : 1; // 1 for terminator.

	// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
	// this call will set up the clipboard for writing:
	if (output_var->Assign(NULL, space_needed - 1) != OK)
		return FAIL;  // It already displayed the error.
	// Fetch the text directly into the var.  Also set the length explicitly
	// in case actual size written was off from the esimated size (since
	// GetWindowTextLength() can return more space that will actually be required
	// in certain circumstances, see MS docs):
	if (control_window)
	{
		output_var->Length() = (VarSizeType)GetWindowTextTimeout(control_window
			, output_var->Contents(), space_needed);
		if (!output_var->Length())
			// There was no text to get or GetWindowTextTimeout() failed.
			*output_var->Contents() = '\0';  // Safe because Assign() gave us a non-constant memory area.
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	}
	else
	{
		*output_var->Contents() = '\0';
		output_var->Length() = 0;
		// And leave g_ErrorLevel set to ERRORLEVEL_ERROR to distinguish a non-existent control
		// from a one that does exist but returns no text.
	}
	// Consider the above to be always successful, even if the window wasn't found, except
	// when below returns an error:
	return output_var->Close();  // In case it's the clipboard.
}



ResultType Line::StatusBarGetText(char *aPart, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	// Note: ErrorLevel is handled by StatusBarUtil(), below.
	DETERMINE_TARGET_WINDOW
	HWND control_window = target_window ? ControlExist(target_window, "msctls_statusbar321") : NULL;
	// Call this even if control_window is NULL because in that case, it will set the output var to
	// be blank for us:
	StatusBarUtil(output_var, control_window, atoi(aPart)); // It will handle any zero part# for us.
	return OK; // Even if it fails, seems best to return OK so that subroutine can continue.
}



ResultType Line::StatusBarWait(char *aTextToWaitFor, char *aSeconds, char *aPart, char *aTitle, char *aText
	, char *aInterval, char *aExcludeTitle, char *aExcludeText)
{
	// Note: ErrorLevel is handled by StatusBarUtil(), below.
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
// Like AutoIt, this function and others like it always return OK, even if the target window doesn't
// exist or there action doesn't actually succeed.
{
	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return OK;
	SetWindowText(target_window, aNewTitle);
	return OK;
}



ResultType Line::WinGetTitle(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	DETERMINE_TARGET_WINDOW
	// Even if target_window is NULL, we want to continue on so that the output
	// param is set to be the empty string, which is the proper thing to do
	// rather than leaving whatever was in there before.

	// Handle the output parameter.  See the comments in ACT_CONTROLGETTEXT for details.
	VarSizeType space_needed = target_window ? GetWindowTextLength(target_window) + 1 : 1; // 1 for terminator.
	if (output_var->Assign(NULL, space_needed - 1) != OK)
		return FAIL;  // It already displayed the error.
	if (target_window)
	{
		output_var->Length() = (VarSizeType)GetWindowText(target_window
			, output_var->Contents(), space_needed);
		if (!output_var->Length())
			// There was no text to get or GetWindowTextTimeout() failed.
			*output_var->Contents() = '\0';  // Safe because Assign() gave us a non-constant memory area.
	}
	else
	{
		*output_var->Contents() = '\0';
		output_var->Length() = 0;
	}
	return output_var->Close();  // In case it's the clipboard.
}



ResultType Line::WinGetText(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	DETERMINE_TARGET_WINDOW
	// Even if target_window is NULL, we want to continue on so that the output
	// variables are set to be the empty string, which is the proper thing to do
	// rather than leaving whatever was in there before:
	if (!target_window)
		return output_var->Assign(); // Tell it not to free the memory by not calling with "".

	length_and_buf_type sab;
	sab.buf = NULL; // Tell it just to calculate the length this time around.
	sab.total_length = sab.capacity = 0; // Init
	EnumChildWindows(target_window, EnumChildGetText, (LPARAM)&sab);

	if (!sab.total_length)
	{
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		return output_var->Assign(); // Tell it not to free the memory by omitting all params.
	}

	// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
	// this call will set up the clipboard for writing:
	if (output_var->Assign(NULL, (VarSizeType)sab.total_length) != OK)
		return FAIL;  // It already displayed the error.

	// Fetch the text directly into the var.  Also set the length explicitly
	// in case actual size written was off from the esimated size (since
	// GetWindowTextLength() can return more space that will actually be required
	// in certain circumstances, see MS docs):
	sab.buf = output_var->Contents();
	sab.total_length = 0; // Init
	sab.capacity = output_var->Capacity(); // Because capacity might be a little larger than we asked for.
	EnumChildWindows(target_window, EnumChildGetText, (LPARAM)&sab);

	output_var->Length() = (VarSizeType)sab.total_length;  // In case it wound up being smaller than expected.
	if (sab.total_length)
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	else
		// Something went wrong, so make sure we set to empty string.
		*output_var->Contents() = '\0';  // Safe because Assign() gave us a non-constant memory area.
	return output_var->Close();  // In case it's the clipboard.
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
	Var *output_var_x = ResolveVarOfArg(0);  // Ok if NULL.
	Var *output_var_y = ResolveVarOfArg(1);  // Ok if NULL.
	Var *output_var_width = ResolveVarOfArg(2);  // Ok if NULL.
	Var *output_var_height = ResolveVarOfArg(3);  // Ok if NULL.

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

	if (output_var_x)
		if (target_window)
		{
			if (!output_var_x->Assign((int)rect.left))  // X position
				result = FAIL;
		}
		else
			if (!output_var_x->Assign(""))
				result = FAIL;
	if (output_var_y)
		if (target_window)
		{
			if (!output_var_y->Assign((int)rect.top))  // Y position
				result = FAIL;
		}
		else
			if (!output_var_y->Assign(""))
				result = FAIL;
	if (output_var_width) // else user didn't want this value saved to an output param
		if (target_window)
		{
			if (!output_var_width->Assign((int)rect.right - rect.left))  // Width
				result = FAIL;
		}
		else
			if (!output_var_width->Assign("")) // Set it to be empty to signal the user that the window wasn't found.
				result = FAIL;
	if (output_var_height)
		if (target_window)
		{
			if (!output_var_height->Assign((int)rect.bottom - rect.top))  // Height
				result = FAIL;
		}
		else
			if (!output_var_height->Assign(""))
				result = FAIL;

	return result;
}



ResultType Line::PixelSearch(int aLeft, int aTop, int aRight, int aBottom, int aColor, int aVariation)
{
	Var *output_var_x = ResolveVarOfArg(0);  // Ok if NULL.
	Var *output_var_y = ResolveVarOfArg(1);  // Ok if NULL.

	g_ErrorLevel->Assign(ERRORLEVEL_ERROR2); // Set default ErrorLevel.  2 means error other than "color not found".
	if (output_var_x)
		output_var_x->Assign();  // Init to empty string regardless of whether we succeed here.
	if (output_var_y)
		output_var_y->Assign(); // Same.

	// Always adjust coords to reflect the position of the foreground window because AutoHotkey
	// doesn't yet support AutoIt3's absolute-screen-coords mode:
	RECT rect;
	GetWindowRect(GetForegroundWindow(), &rect);
	aLeft   += rect.left;
	aTop    += rect.top;
	aRight  += rect.left;  // Add left vs. right because we're adjusting based on the position of the window.
	aBottom += rect.top;   // Same.

	HDC hdc = GetDC(NULL);
	if (!hdc)
		return OK;  // Let ErrorLevel tell the story.

	if (aVariation < 0) aVariation = 0;
	if (aVariation > 255) aVariation = 255;
	BYTE search_red = GetRValue(aColor);
	BYTE search_green = GetGValue(aColor);
	BYTE search_blue = GetBValue(aColor);
	BYTE red_low, red_high, green_low, green_high, blue_low, blue_high;
	if (aVariation <= 0)  // User wanted an exact match.
	{
		red_low = red_high = search_red;
		green_low = green_high = search_green;
		blue_low = blue_high = search_blue;
	}
	else
	{
		// Allow colors to vary within the spectrum of intensity, rather than having them
		// wrap around (which doesn't seem to make much sense).  For example, if the user specified
		// a variation of 5, but the red component of aColor is only 0x01, we don't want red_low to go
		// below zero, which would cause it to wrap around to a very intense red color:
		red_low = (aVariation > search_red) ? 0 : search_red - aVariation;
		green_low = (aVariation > search_green) ? 0 : search_green - aVariation;
		blue_low = (aVariation > search_blue) ? 0 : search_blue - aVariation;
		red_high = (aVariation > 0xFF - search_red) ? 0xFF : search_red + aVariation;
		green_high = (aVariation > 0xFF - search_green) ? 0xFF : search_green + aVariation;
		blue_high = (aVariation > 0xFF - search_blue) ? 0xFF : search_blue + aVariation;
	}

	int xpos, ypos, color;
	BYTE red, green, blue;
	bool match_found;
	ResultType result = OK;
	for (xpos = aLeft; xpos <= aRight; ++xpos)
	{
		for (ypos = aTop; ypos <= aBottom; ++ypos)
		{
			color = GetPixel(hdc, xpos, ypos);
			if (aVariation <= 0)  // User wanted an exact match.
				match_found = (color == aColor);
			else  // User specified that some variation in each of the RGB components is allowable.
			{
				red = GetRValue(color);
				green = GetGValue(color);
				blue = GetBValue(color);
				match_found = (red >= red_low && red <= red_high
					&& green >= green_low && green <= green_high
					&& blue >= blue_low && blue <= blue_high);
			}
			if (match_found) // This pixel matches one of the specified color(s).
			{
				ReleaseDC(NULL, hdc);
				if (output_var_x)
					// Adjust coords to make them relative to the position of the target window:
					if (!output_var_x->Assign(xpos - rect.left))
						result = FAIL;
				if (output_var_y)
					if (!output_var_y->Assign(ypos - rect.top))
						result = FAIL;
				if (result == OK)
					g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
				return result;
			}
		}
	}

	// If the above didn't return, the pixel wasn't found in the specified region.
	// So leave ErrorLevel set to "error" to indicate that:
	ReleaseDC(NULL, hdc);
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // This value indicates "color not found".
	return OK;
}



ResultType Line::PixelGetColor(int aX, int aY)
// This has been adapted from the AutoIt3 source.
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	output_var->Assign(); // Init to empty string regardless of whether we succeed here.

	RECT rect;
	GetWindowRect(GetForegroundWindow(), &rect);
	aX += rect.left;
	aY += rect.top;

	HDC hdc = GetDC(NULL);
	if (!hdc)
		return OK;  // Let ErrorLevel tell the story.
	ResultType result = output_var->Assign((int)GetPixel(hdc, aX, aY));
	ReleaseDC(NULL, hdc);

	if (result == OK)
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return result;  // i.e. only return failure if something unexpected happens when assigning to the variable.
}



/////////////////
// Main Window //
/////////////////

LRESULT CALLBACK MainWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
	char buf[2048], *last_backslash;  // For various uses.

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
			g_script.Edit();
			return 0;
		case ID_TRAY_RELOADSCRIPT:
			g_script.Reload();
			return 0;
		case ID_TRAY_WINDOWSPY:
			strlcpy(buf, g_script.mOurEXE, sizeof(buf)); // Make a modifiable copy.
			last_backslash = strrchr(buf, '\\');
			if (!last_backslash)
				break;
			last_backslash[1] = '\0';
			snprintfcat(buf, sizeof(buf), "AU3_Spy.exe");
			g_script.ActionExec(buf, "");
			return 0;
		case ID_TRAY_HELP:
			strlcpy(buf, g_script.mOurEXE, sizeof(buf)); // Make a modifiable copy.
			last_backslash = strrchr(buf, '\\');
			if (!last_backslash)
				break;
			last_backslash[1] = '\0';
			snprintfcat(buf, sizeof(buf), "AutoHotkey.chm");
			g_script.ActionExec(buf, "");
			return 0;
		case ID_TRAY_SUSPEND:
			Line::ToggleSuspendState();
			return 0;
		case ID_TRAY_PAUSE:
			if (!g.IsPaused && g_IsIdle)
			{
				MsgBox("The script cannot be paused while it is doing nothing.  If you wish to prevent new"
					" hotkey subroutines from running, use Suspend instead.");
				// i.e. we don't want idle scripts to ever be in a paused state.
				return 0;
			}
			if (g.IsPaused)
				--g_nPausedSubroutines;
			else
				++g_nPausedSubroutines;
			g.IsPaused = !g.IsPaused;
			g_script.UpdateTrayIcon();
			return 0;
		case ID_TRAY_EXIT:
			g_script.ExitApp();  // More reliable than PostQuitMessage(), which has been known to fail in rare cases.
			return 0;
		} // Inner switch()
		break;

	case AHK_NOTIFYICON:  // Tray icon clicked on.
	{
        switch(lParam)
        {
// Don't allow the main window to be opened this way by a compiled EXE, since it will display
// the lines most recently executed (and many people who compile scripts don't want their users
// to see the contents of the script:
#ifndef AUTOHOTKEYSC
		case WM_LBUTTONDBLCLK:
			ShowMainWindow();
			return 0;
#endif
		case WM_RBUTTONDOWN:
		{
			HMENU hMenu = LoadMenu(g_hInstance, MAKEINTRESOURCE(IDR_MENU1));
			// TrackPopupMenu cannot display the menu bar so get
			// the handle to the first shortcut menu.
			if (!hMenu)
				return 0;
 			if (   !(hMenu = GetSubMenu(hMenu, 0))   )  // It seems this must be done prior to some of the below.
				return 0;
#ifdef AUTOHOTKEYSC
			DeleteMenu(hMenu, ID_TRAY_OPEN, MF_BYCOMMAND);
			DeleteMenu(hMenu, ID_TRAY_HELP, MF_BYCOMMAND);
			DeleteMenu(hMenu, ID_TRAY_WINDOWSPY, MF_BYCOMMAND);
			// Doesn't make sense to have this option since script can't
			// be edited once it's inside the EXE:
			DeleteMenu(hMenu, ID_TRAY_RELOADSCRIPT, MF_BYCOMMAND);
			DeleteMenu(hMenu, ID_TRAY_EDITSCRIPT, MF_BYCOMMAND);
			// Delete the left-over menu separators.  This isn't a very
			// maintainable way to do it, but I'm doing something wrong
			// with the other method, so using this for now:
			DeleteMenu(hMenu, 1, MF_BYPOSITION);
			DeleteMenu(hMenu, 0, MF_BYPOSITION);
			// Trying to code this is a way that easier to maintain:
			//int menu_item_count = GetMenuItemCount(hMenu);
			//MENUITEMINFO mii;
			//for (int i = 0; i < menu_item_count; ++i)
			//{
			//	if (GetMenuItemInfo(hMenu, i, TRUE, &mii))
			//		if (mii.fType == MFT_SEPARATOR)
			//			DeleteMenu(hMenu, i, MF_BYPOSITION);
			//}
#else
			SetMenuDefaultItem(hMenu, ID_TRAY_OPEN, FALSE);
#endif
			CheckMenuItem(hMenu, ID_TRAY_SUSPEND, g_IsSuspended ? MF_CHECKED : MF_UNCHECKED);
			CheckMenuItem(hMenu, ID_TRAY_PAUSE, g.IsPaused ? MF_CHECKED : MF_UNCHECKED);
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
		if (g_TrayMenuIsVisible || (iMsg == AHK_HOOK_HOTKEY && lParam && g_hWnd == GetForegroundWindow()))
		{
			// Ok this is a little strange, but the thought here is that if the tray menu is
			// displayed, it should be closed prior to executing any new hotkey.  This is
			// because hotkeys usually cause other windows to become active, and once that
			// happens, the tray menu cannot be closed except by choosing a menu item in it
			// (which is often undesirable).  This is also done if the hook told us that
			// this event is something that may have invoked the tray menu or a context
			// menu in our own main window, because otherwise such menus can't be dismissed
			// by another mouseclick or (in the case of the tray menu) they don't work reliably.
			// However, I'm not sure that this AHK_HOOK_HOTKEY workaround won't cause more
			// problems than it solves (i.e. WM_CANCELMODE might be called in some cases
			// when it doesn't need to be, and then might have some undesirable side-effect
			// since I believe it has other purposes besides dismissing menus).  But such
			// side effects should be minimal since WM_CANCELMODE mode is only set if
			// this AutoHotkey instance's own main window is active:
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

	case AHK_KEYLOG:
		KeyLogToFile(NULL, ((KeyLogItem *)wParam)->event_type, ((KeyLogItem *)wParam)->key_up
			, ((KeyLogItem *)wParam)->vk, ((KeyLogItem *)wParam)->sc);
		return 0;

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
	if (!aTitle || !*aTitle)
		// If available, the script's filename seems a much better title in case the user has
		// more than one script running:
		aTitle = (g_script.mFileName && *g_script.mFileName) ? g_script.mFileName : NAME_PV;
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
				// For AutoIt2 (.aut) script:
				// If the user presses the cancel button, we set the output variable to be blank so
				// that there is a way to detect that the cancel button was pressed.  This is because
				// the InputBox command does not set ErrorLevel for .aut scripts (to maximize backward
				// compatibility).
				// For non-AutoIt2 scripts: The output variable is set to whatever the user entered,
				// even if the user pressed the cancel button.  This allows the cancel button to specify
				// that a different operation should be performed on the entered text.
				#define SET_OUTPUT_VAR_TO_BLANK (LOWORD(wParam) == IDCANCEL && g_script.mIsAutoIt2)
				#define INPUTBOX_VAR CURR_INPUTBOX.output_var
				if (!g_script.mIsAutoIt2)
					g_ErrorLevel->Assign(LOWORD(wParam) == IDCANCEL ? ERRORLEVEL_ERROR : ERRORLEVEL_NONE);
				VarSizeType space_needed = SET_OUTPUT_VAR_TO_BLANK ? 1 : GetWindowTextLength(hControl) + 1;
				// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
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
					if (SET_OUTPUT_VAR_TO_BLANK)
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
	if (aX1 != COORD_UNSPECIFIED && aY1 != COORD_UNSPECIFIED)
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
	#define MOUSE_SLEEP SLEEP_AND_IGNORE_HOTKEYS(g.KeyDelay > 10 ? g.KeyDelay : 10)

	// Do the drag operation
	switch (aVK)
	{
	case VK_LBUTTON:
		MouseEvent(MOUSEEVENTF_LEFTDOWN);
		MOUSE_SLEEP;
		MouseMove(aX2, aY2, aSpeed);
		MOUSE_SLEEP;
		MouseEvent(MOUSEEVENTF_LEFTUP);
		break;
	case VK_RBUTTON:
		MouseEvent(MOUSEEVENTF_RIGHTDOWN);
		MOUSE_SLEEP;
		MouseMove(aX2, aY2, aSpeed);
		MOUSE_SLEEP;
		MouseEvent(MOUSEEVENTF_RIGHTUP);
		break;
	case VK_MBUTTON:
		MouseEvent(MOUSEEVENTF_MIDDLEDOWN);
		MOUSE_SLEEP;
		MouseMove(aX2, aY2, aSpeed);
		MOUSE_SLEEP;
		MouseEvent(MOUSEEVENTF_MIDDLEUP);
		break;
	}
	// It seems best to always do this one too in case the script line that caused
	// us to be called here is followed immediately by another script line which
	// is either another mouse click or something that relies upon this mouse drag
	// having been completed:
	MOUSE_SLEEP;
	return OK;
}



ResultType Line::MouseClick(vk_type aVK, int aX, int aY, int aClickCount, int aSpeed, char aEventType)
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

	// The chars 'U' (up) and 'D' (down), if specified, will restrict the clicks
	// to being only DOWN or UP (so that the mouse button can be held down, for
	// example):
	aEventType = toupper(aEventType);

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
			if (aEventType != 'U')
			{
				MouseEvent(MOUSEEVENTF_LEFTDOWN);
				MOUSE_SLEEP;
			}
			if (aEventType != 'D')
			{
				MouseEvent(MOUSEEVENTF_LEFTUP);
				// It seems best to always do this one too in case the script line that caused
				// us to be called here is followed immediately by another script line which
				// is either another mouse click or something that relies upon the mouse click
				// having been completed:
				MOUSE_SLEEP;
			}
			break;
		case VK_RBUTTON:
			if (aEventType != 'U')
			{
				MouseEvent(MOUSEEVENTF_RIGHTDOWN);
				MOUSE_SLEEP;
			}
			if (aEventType != 'D')
			{
				MouseEvent(MOUSEEVENTF_RIGHTUP);
				MOUSE_SLEEP;
			}
			break;
		case VK_MBUTTON:
			if (aEventType != 'U')
			{
				MouseEvent(MOUSEEVENTF_MIDDLEDOWN);
				MOUSE_SLEEP;
			}
			if (aEventType != 'D')
			{
				MouseEvent(MOUSEEVENTF_MIDDLEUP);
				MOUSE_SLEEP;
			}
			break;
		}
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
		MouseEvent(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, aX, aY);
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

		MouseEvent(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, xCur, yCur);
		MOUSE_SLEEP;
	}
}



ResultType Line::MouseGetPos()
// Returns OK or FAIL.
// This has been adapted from the AutoIt3 source.
{
	// Caller should already have ensured that at least one of these will be non-NULL.
	// The only time this isn't true is for dynamically-built variable names.  In that
	// case, we don't worry about it if it's NULL, since the user will already have been
	// warned:
	Var *output_var_x = ResolveVarOfArg(0);  // Ok if NULL.
	Var *output_var_y = ResolveVarOfArg(1);  // Ok if NULL.

	RECT rect;
	POINT pt;
	GetCursorPos(&pt);  // Realistically, can't fail?
	HWND fore_win = GetForegroundWindow();
	if (fore_win)
		GetWindowRect(fore_win, &rect);
	else // ensure it's initialized for later calculations:
		rect.bottom = rect.left = rect.right = rect.top = 0;

	ResultType result = OK; // Set default;

	if (output_var_x) // else the user didn't want the X coordinate, just the Y.
		if (!output_var_x->Assign((int)(pt.x - rect.left)))
			result = FAIL;
	if (output_var_y) // else the user didn't want the Y coordinate, just the X.
		if (!output_var_y->Assign((int)(pt.y - rect.top)))
			result = FAIL;

	return result;
}



///////////////////////////////
// Related to other commands //
///////////////////////////////

ResultType Line::PerformAssign()
// Returns OK or FAIL.
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	// Find out if output_var (the var being assigned to) is dereferenced (mentioned) in this line's
	// second arg, which is the value to be assigned.  If it isn't, things are much simpler.
	// Note: Since Arg#2 for this function is never an output or an input variable, it is not
	// necessary to check whether its the same variable as Arg#1 for this determination.
	// Note: If output_var is the clipboard, it can be used in the source deref(s) while also
	// being the target -- without having to use the deref buffer -- because the clipboard
	// has it's own temp buffer: the memory area to which the result is written.
	// The prior content of the clipboard remains available in its other memory area
	// until Commit() is called (i.e. long enough for our purposes):
	bool target_is_involved_in_source = false;
	if (output_var->mType != VAR_CLIPBOARD && mArgc > 1)
		// It has a second arg, which in this case is the value to be assigned to the var.
		// Examine any derefs that the second arg has to see if output_var is mentioned:
		for (DerefType *deref = mArg[1].deref; deref && deref->marker; ++deref)
			if (deref->var == output_var)
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
	// be sure to check that output_var is mentioned only once in the list of derefs.
	// For example, something like this would probably be much easier to
	// implement by using ExpandArgs(): Var1 = xxxx Var1 Var2 Var1 xxxx.
	// So the main thing to be possibly later improved here is the case where
	// output_var is mentioned only once in the deref list:
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
		return output_var->Assign("");  // If the var is of large capacity, this will also free its memory.

	if (target_is_involved_in_source)
		// It was already dereferenced above, so use ARG2, which points to the
		// derefed contents of ARG2 (i.e. the data to be assigned).
		// Seems better to trim even if not AutoIt2, since that's currently the only way easy way
		// to trim things:
		return output_var->Assign(ARG2, space_needed - 1, g.AutoTrim); // , g_script.mIsAutoIt2);

	// Otherwise:
	// If we're here, output_var->mType must be clipboard or normal because otherwise
	// the validation during load would have prevented the script from loading:

	// First set everything up for the operation.  If output_var is the clipboard, this
	// will prepare the clipboard for writing:
	if (output_var->Assign(NULL, space_needed - 1) != OK)
		return FAIL;
	// Expand Arg2 directly into the var.  Also set the length explicitly
	// in case actual size written was off from the esimated size, perhaps
	// due to a failure or size discrepancy between the deref size-estimate
	// and the actual deref itself.  Note: If output_var is the clipboard,
	// it's probably okay if the below actually writes less than the size of
	// the mem that has already been allocated for the new clipboard contents
	// That might happen due to a failure or size discrepancy between the
	// deref size-estimate and the actual deref itself:
	output_var->Length() = (VarSizeType)(ExpandArg(output_var->Contents(), 1) - output_var->Contents() - 1);
	if (g.AutoTrim)
	{
		trim(output_var->Contents());
		output_var->Length() = (VarSizeType)strlen(output_var->Contents());
	}
	return output_var->Close();  // i.e. Consider this function to be always successful unless this fails.
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



ResultType Line::DriveSpaceFree(char *aPath)
// Adapted from the AutoIt3 source.
// Because of NTFS's ability to mount volumes into a directory, a path might not necessarily
// have the same amount of free space as its root drive.  However, I'm not sure if this
// method here actually takes that into account.
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	output_var->Assign(); // Init to empty string regardless of whether we succeed here.

	if (!aPath || !*aPath) return OK;  // Let ErrorLevel tell the story.  Below relies on this check.

	char buf[MAX_PATH * 2];
	strlcpy(buf, aPath, sizeof(buf));
	size_t length = strlen(buf);
	if (buf[length - 1] != '\\') // AutoIt3: Attempt to fix the parameter passed.
	{
		if (length + 1 >= sizeof(buf)) // No room to fix it.
			return OK; // Let ErrorLevel tell the story.
		buf[length++] = '\\';
		buf[length] = '\0';
	}

	SetErrorMode(SEM_FAILCRITICALERRORS);  // AutoIt3: So a:\ does not ask for disk

#ifdef _MSC_VER									// AutoIt3: Need a MinGW solution for __int64
	ULARGE_INTEGER	uiTotal, uiFree, uiUsed;
	DWORD			dwSectPerClust, dwBytesPerSect, dwFreeClusters, dwTotalClusters;
	static FARPROC  pGetDiskFreeSpaceEx = GetProcAddress(GetModuleHandle("kernel32.dll"), "GetDiskFreeSpaceExA");
	int				test = ERROR_SUCCESS;

	// MSDN: "The GetDiskFreeSpaceEx function returns correct values for all volumes, including those
	// that are greater than 2 gigabytes."
	if (pGetDiskFreeSpaceEx)  // Function is available (unpatched Win95 and WinNT might not have it).
	{
		if (!GetDiskFreeSpaceEx(buf,&uiFree,&uiTotal,&uiUsed)) // AutoIt3 just calls it by name like this.
			return OK; // Let ErrorLevel tell the story.
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		// Casting this way limits us to 2,097,152 gigabytes in size:
		return output_var->Assign(   (int)((unsigned __int64)(uiFree.QuadPart) / (unsigned __int64)(1024*1024))   );
	}
	else
	{
		if (!GetDiskFreeSpace(buf, &dwSectPerClust, &dwBytesPerSect, &dwFreeClusters, &dwTotalClusters))
			return OK; // Let ErrorLevel tell the story.
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		return output_var->Assign(   (int)((unsigned __int64)(dwFreeClusters * dwSectPerClust * dwBytesPerSect)
			/ (unsigned __int64)(1024*1024))   );
	}
#endif	

	return OK;
}



ResultType Line::SoundPlay(char *aFilespec, bool aSleepUntilDone)
{
	// Adapted from the AutoIt3 source.
	// See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/multimed/htm/_win32_play.asp
	// for some documentation mciSendString() and related.
	#define SOUNDPLAY_ALIAS "AHK_PlayMe"
	char buf[MAX_PATH * 2];
	mciSendString("status " SOUNDPLAY_ALIAS " mode", buf, sizeof(buf), NULL);
	if (*buf) // "playing" or "stopped" (so close it before trying to re-open with a new aFilespec).
		mciSendString("close " SOUNDPLAY_ALIAS, NULL, 0, NULL);
	snprintf(buf, sizeof(buf), "open \"%s\" alias " SOUNDPLAY_ALIAS, aFilespec);
	if (mciSendString(buf, NULL, 0, NULL)) // Failure.
	{
		g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		return OK;  // OK since this isn't a critical error.
	}
	if (mciSendString("play " SOUNDPLAY_ALIAS, NULL, 0, NULL)) // Failure.
	{
		g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		return OK;  // OK since this isn't a critical error.
	}
	// Otherwise, the sound is now playing.
	g_ErrorLevel->Assign(ERRORLEVEL_NONE);
	if (!aSleepUntilDone)
		return OK;
	// Otherwise we're to wait until the sound is done.  To allow our app to remain responsive
	// during this time, and so that the keyboard and mouse hook (if installed) won't cause
	// key & mouse lag, we do this in a loop rather than AutoIt3's method:
	// "mciSendString("play " SOUNDPLAY_ALIAS " wait",NULL,0,NULL)"
	for (;;)
	{
		mciSendString("status " SOUNDPLAY_ALIAS " mode", buf, sizeof(buf), NULL);
		if (!*buf) // Probably can't happen given the state we're in.
			break;
		if (!strcmp(buf, "stopped")) // The sound is done playing.
		{
			mciSendString("close " SOUNDPLAY_ALIAS, NULL, 0, NULL);
			break;
		}
		// Sleep a little longer than normal because I'm not sure how much overhead
		// and CPU utilization the above incurs:
		MsgSleep(20);
	}
	return OK;
}



ResultType Line::FileSelectFile(char *aOptions, char *aWorkingDir)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
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
		return output_var->Assign(); // Tell it not to free the memory by not calling with "".
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
	return output_var->Assign(ofn.lpstrFile);
}



ResultType Line::FileSelectFolder(char *aRootDir, bool aAllowCreateFolder, char *aGreeting)
// Adapted from the AutoIt3 source.
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	if (!aRootDir) aRootDir = "";
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	if (!output_var->Assign())  // Initialize the output variable.
		return FAIL;

	if (g_nFolderDialogs >= MAX_FOLDERDIALOGS)
	{
		// Have a maximum to help prevent runaway hotkeys due to key-repeat feature, etc.
		MsgBox("The maximum number of Folder Dialogs has been reached." ERR_ABORT);
		return FAIL;
	}

	LPMALLOC pMalloc;
    if (SHGetMalloc(&pMalloc) != NOERROR)	// Initialize
		return OK;  // Let ErrorLevel tell the story.

	BROWSEINFO browseInfo;
	if (*aRootDir)
	{
		IShellFolder *pDF;
		if (SHGetDesktopFolder(&pDF) == NOERROR)
		{
			LPITEMIDLIST pIdl = NULL;
			ULONG        chEaten;
			ULONG        dwAttributes;
			OLECHAR olePath[MAX_PATH];			// wide-char version of path name
			MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, aRootDir, -1, olePath, sizeof(olePath));
			pDF->ParseDisplayName(NULL, NULL, olePath, &chEaten, &pIdl, &dwAttributes);
			pDF->Release();
			browseInfo.pidlRoot = pIdl;
		}
	}
	else
		browseInfo.pidlRoot = NULL;  // Since aRootDir, this should make it use "My Computer" as the root dir.

	int iImage = 0;
	browseInfo.iImage = iImage;
	browseInfo.hwndOwner = NULL;  // i.e. no need to have main window forced into the background for this.
	char greeting[1024];
	if (aGreeting && *aGreeting)
		strlcpy(greeting, aGreeting, sizeof(greeting));
	else
		snprintf(greeting, sizeof(greeting), "Select Folder - %s", g_script.mFileName);
	browseInfo.lpszTitle = greeting;
	browseInfo.lpfn = NULL;
	browseInfo.ulFlags = 0x0040 | (aAllowCreateFolder ? 0 : 0x0200);

	char Result[2048];
	browseInfo.pszDisplayName = Result;  // This will hold the user's choice.

	// This will attempt to force it to the foreground after it has been displayed, since the
	// dialog often will flash in the task bar instead of becoming foreground.
	// See MsgBox() for details:
	PostMessage(g_hWnd, AHK_DIALOG, (WPARAM)0, (LPARAM)0); // Must pass 0 for WPARAM in this case.

	g.WaitingForDialog = true;
	++g_nFolderDialogs;
	LPITEMIDLIST lpItemIDList = SHBrowseForFolder(&browseInfo);  // Spawn Dialog
	--g_nFolderDialogs;
	g.WaitingForDialog = false;  // IsCycleComplete() relies on this.

	if (!lpItemIDList)
		return OK;  // Let ErrorLevel tell the story.

	*Result = '\0';  // Reuse this var, this time to old the result of the below:
	SHGetPathFromIDList(lpItemIDList, Result);
	pMalloc->Free(lpItemIDList);
	pMalloc->Release();

	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return output_var->Assign(Result);
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
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	UINT line_number = atoi(aLineNumber);
	if (line_number <= 0)
		return OK;  // Return OK because g_ErrorLevel tells the story.
	FILE *fp = fopen(aFilespec, "r");
	if (!fp)
		return OK;  // Return OK because g_ErrorLevel tells the story.

	// Remember that once the first call to MsgSleep() is done, a new hotkey subroutine
	// may fire and suspend what we're doing here.  Such a subroutine might also overwrite
	// the values our params, some of which may be in the deref buffer.  So be sure not
	// to refer to those strings once MsgSleep() has been done, below.  Alternatively,
	// a copy of such params can be made using our own stack space.

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
			MsgSleep(sleep_duration); // See above comment. Also, it seems okay to allow new hotkeys during the sleep.
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
		if (!output_var->Assign()) // Explicitly call it this way so that it won't free the memory.
			return FAIL;
	}
	else
		if (!output_var->Assign(buf, (VarSizeType)buf_length))
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
	char file_path[MAX_PATH * 2];  // Give extra room in case OS supports extra-long filenames?
	char target_filespec[MAX_PATH * 2];
	if (strlen(aFilePattern) >= sizeof(file_path))
		return OK; // Return OK because this is non-critical.  Let the above ErrorLevel indicate the problem.
	strlcpy(file_path, aFilePattern, sizeof(file_path));
	char *last_backslash = strrchr(file_path, '\\');
	// Remove the filename and/or wildcard part.   But leave the trailing backslash on it for
	// consistency with below:
	if (last_backslash)
		*(last_backslash + 1) = '\0';
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
	g_ErrorLevel->Assign(failure_count); // i.e. indicate success if there were no failures.
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



ResultType Line::FileGetAttrib(char *aFilespec)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default
	output_var->Assign(); // Init to be blank, in case of failure.

	if (!aFilespec || !*aFilespec)
		return OK;  // Let ErrorLevel indicate an error, since this is probably not what the user intended.

	DWORD attr = GetFileAttributes(aFilespec);
	if (attr == 0xFFFFFFFF)  // Failure, probably because file doesn't exist.
		return OK;  // Let ErrorLevel tell the story.

	g_ErrorLevel->Assign(ERRORLEVEL_NONE);
	char attr_string[128];
	return output_var->Assign(FileAttribToStr(attr_string, attr));
}



int Line::FileSetAttrib(char *aAttributes, char *aFilePattern, FileLoopModeType aOperateOnFolders
	, bool aDoRecurse, bool aCalledRecursively)
// Returns the number of files and folders that could not be changed due to an error.
// Current limitation: It will not recurse into subfolders unless their names also match
// aFilePattern.
{
	// Making this static is not friendly to reentrant calls to this function (i.e. calls maded
	// as a consequence of the current script subroutine being interrupted by another during
	// this instance's MsgSleep()).  However, it doesn't seem to be that much of a consequence
	// since the exact interval period of the MsgSleep()'s isn't that important.  It's also
	// pretty unlikely that the interrupting subroutine will also just happen to call this
	// function rather than some other.  On the other hand, we don't do failure_count as a
	// static because we want its value to be 100% accurate even if reentrancy occurs:
	static int files_considered;
	if (!aCalledRecursively)  // i.e. Only need to do this if we're not called by ourself:
	{
		files_considered = 0;
		g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default
		if (!aFilePattern || !*aFilePattern)
			return 0;  // Let ErrorLevel indicate an error, since this is probably not what the user intended.
		if (aOperateOnFolders == FILE_LOOP_INVALID) // In case runtime dereference of a var was an invalid value.
			aOperateOnFolders = FILE_LOOP_FILES_ONLY;  // Set default.
	}

	// Related to the comment at the top: Since the script subroutine that resulted in the call to
	// this function can be interrupted during our MsgSleep(), make a copy of any params that might
	// currently point directly to the deref buffer.  This is done because their contents might
	// be overwritten by the interrupting subroutine:
	char attributes[64];
	strlcpy(attributes, aAttributes, sizeof(attributes));
	char file_pattern[MAX_PATH * 2];  // In case OS supports extra-long filenames.
	strlcpy(file_pattern, aFilePattern, sizeof(file_pattern));

	// Give extra room in some of these vars, in case OS supports extra-long filenames:
	char file_path[MAX_PATH * 2];
	char target_filespec[MAX_PATH * 2];
	if (strlen(file_pattern) >= sizeof(file_path))
		return 0; // Let the above ErrorLevel indicate the problem.
	strlcpy(file_path, file_pattern, sizeof(file_path));
	char *last_backslash = strrchr(file_path, '\\');
	// Remove the filename and/or wildcard part.   But leave the trailing backslash on it for
	// consistency with below:
	if (last_backslash)
		*(last_backslash + 1) = '\0';
	else // Use current working directory, e.g. if user specified only *.*
		*file_path = '\0';

	// For use with aDoRecurse, get just the naked file name/pattern:
	char *naked_filename_or_pattern = strrchr(file_pattern, '\\');
	if (naked_filename_or_pattern)
		++naked_filename_or_pattern;
	else
		naked_filename_or_pattern = file_pattern;

	if (!StrChrAny(naked_filename_or_pattern, "?*"))
		// Since no wildcards, always operate on this single item even if it's a folder.
		aOperateOnFolders = FILE_LOOP_FILES_AND_FOLDERS;

	int batch_size, sleep_duration;
	if (Hotkey::HookIsActive())
	{
		batch_size = 100; // i.e. try to minimize keyboard & mouse lag by sleeping more often.
		sleep_duration = 10;
	}
	else
	{
		batch_size = 1000;
		sleep_duration = -1; // Since all we want to do is check messages.
	}

	WIN32_FIND_DATA current_file;
	HANDLE file_search = FindFirstFile(file_pattern, &current_file);
	bool file_found = (file_search != INVALID_HANDLE_VALUE);
	char *cp;
	int failure_count = 0;
	enum attrib_modes {ATTRIB_MODE_NONE, ATTRIB_MODE_ADD, ATTRIB_MODE_REMOVE, ATTRIB_MODE_TOGGLE};
	attrib_modes mode = ATTRIB_MODE_NONE;

	for (; file_found; file_found = FindNextFile(file_search, &current_file), ++files_considered)
	{
		if (files_considered && !(files_considered % batch_size))
			MsgSleep(sleep_duration); // Allows the program to be responsive even during a very long operation.
		if (current_file.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
		{
			if (!strcmp(current_file.cFileName, "..") || !strcmp(current_file.cFileName, "."))
				continue; // Never operate upon or recurse into these.
			// Regardless of whether this folder will be recursed into, if this param
			// is false, this folder's own attributes will not be affected:
			if (aOperateOnFolders == FILE_LOOP_FILES_ONLY)
				continue;
		}
		else // It's a file, not a folder.
			if (aOperateOnFolders == FILE_LOOP_FOLDERS_ONLY)
				continue;
		for (cp = attributes; *cp; ++cp)
		{
			switch (toupper(*cp))
			{
			case '+': mode = ATTRIB_MODE_ADD; break;
			case '-': mode = ATTRIB_MODE_REMOVE; break;
			case '^': mode = ATTRIB_MODE_TOGGLE; break;
			// Note that D (directory) and C (compressed) are currently not supported:
			case 'R':
				if (mode == ATTRIB_MODE_ADD)
					current_file.dwFileAttributes |= FILE_ATTRIBUTE_READONLY;
				else if (mode == ATTRIB_MODE_REMOVE)
					current_file.dwFileAttributes &= ~FILE_ATTRIBUTE_READONLY;
				else if (mode == ATTRIB_MODE_TOGGLE)
					current_file.dwFileAttributes ^= FILE_ATTRIBUTE_READONLY;
				break;
			case 'A':
				if (mode == ATTRIB_MODE_ADD)
					current_file.dwFileAttributes |= FILE_ATTRIBUTE_ARCHIVE;
				else if (mode == ATTRIB_MODE_REMOVE)
					current_file.dwFileAttributes &= ~FILE_ATTRIBUTE_ARCHIVE;
				else if (mode == ATTRIB_MODE_TOGGLE)
					current_file.dwFileAttributes ^= FILE_ATTRIBUTE_ARCHIVE;
				break;
			case 'S':
				if (mode == ATTRIB_MODE_ADD)
					current_file.dwFileAttributes |= FILE_ATTRIBUTE_SYSTEM;
				else if (mode == ATTRIB_MODE_REMOVE)
					current_file.dwFileAttributes &= ~FILE_ATTRIBUTE_SYSTEM;
				else if (mode == ATTRIB_MODE_TOGGLE)
					current_file.dwFileAttributes ^= FILE_ATTRIBUTE_SYSTEM;
				break;
			case 'H':
				if (mode == ATTRIB_MODE_ADD)
					current_file.dwFileAttributes |= FILE_ATTRIBUTE_HIDDEN;
				else if (mode == ATTRIB_MODE_REMOVE)
					current_file.dwFileAttributes &= ~FILE_ATTRIBUTE_HIDDEN;
				else if (mode == ATTRIB_MODE_TOGGLE)
					current_file.dwFileAttributes ^= FILE_ATTRIBUTE_HIDDEN;
				break;
			case 'N':  // Docs say it's valid only when used alone.  But let the API handle it if this is not so.
				if (mode == ATTRIB_MODE_ADD)
					current_file.dwFileAttributes |= FILE_ATTRIBUTE_NORMAL;
				else if (mode == ATTRIB_MODE_REMOVE)
					current_file.dwFileAttributes &= ~FILE_ATTRIBUTE_NORMAL;
				else if (mode == ATTRIB_MODE_TOGGLE)
					current_file.dwFileAttributes ^= FILE_ATTRIBUTE_NORMAL;
				break;
			case 'O':
				if (mode == ATTRIB_MODE_ADD)
					current_file.dwFileAttributes |= FILE_ATTRIBUTE_OFFLINE;
				else if (mode == ATTRIB_MODE_REMOVE)
					current_file.dwFileAttributes &= ~FILE_ATTRIBUTE_OFFLINE;
				else if (mode == ATTRIB_MODE_TOGGLE)
					current_file.dwFileAttributes ^= FILE_ATTRIBUTE_OFFLINE;
				break;
			case 'T':
				if (mode == ATTRIB_MODE_ADD)
					current_file.dwFileAttributes |= FILE_ATTRIBUTE_TEMPORARY;
				else if (mode == ATTRIB_MODE_REMOVE)
					current_file.dwFileAttributes &= ~FILE_ATTRIBUTE_TEMPORARY;
				else if (mode == ATTRIB_MODE_TOGGLE)
					current_file.dwFileAttributes ^= FILE_ATTRIBUTE_TEMPORARY;
				break;
			}
		}
		snprintf(target_filespec, sizeof(target_filespec), "%s%s", file_path, current_file.cFileName);
		if (!SetFileAttributes(target_filespec, current_file.dwFileAttributes))
			++failure_count;
	}

	if (file_search != INVALID_HANDLE_VALUE) // In case the loop had zero iterations.
		FindClose(file_search);

	if (aDoRecurse)
	{
		// Now that the base directory of aFilePattern has been processed, recurse into
		// any subfolders to look for the same pattern.  Create a new pattern (don't reuse
		// file_pattern since naked_filename_or_pattern points into it) to be *.* so that
		// we can find all subfolders, not just those matching the original pattern:
		char all_file_pattern[MAX_PATH * 2];  // In case OS supports extra-long filenames.
		snprintf(all_file_pattern, sizeof(all_file_pattern), "%s*.*", file_path);
		file_search = FindFirstFile(all_file_pattern, &current_file);
		file_found = (file_search != INVALID_HANDLE_VALUE);
		for (; file_found; file_found = FindNextFile(file_search, &current_file), ++files_considered)
		{
			if (files_considered && !(files_considered % batch_size))
				MsgSleep(sleep_duration); // Allows the program to be responsive even during a very long operation.
			if (!(current_file.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
				continue;
			if (!strcmp(current_file.cFileName, "..") || !strcmp(current_file.cFileName, "."))
				continue; // Never recurse into these.
			// This will build the string CurrentDir+SubDir+FilePatternOrName.
			// If FilePatternOrName doesn't contain a wildcard, the recursion
			// process will attempt to operate on the originally-specified
			// single filename or folder name if it occurs anywhere else in the
			// tree, e.g. recursing C:\Temp\temp.txt would affect all occurences
			// of temp.txt both in C:\Temp and any subdirectories it might contain:
			snprintf(target_filespec, sizeof(target_filespec), "%s%s\\%s"
				, file_path, current_file.cFileName, naked_filename_or_pattern);
			failure_count += FileSetAttrib(attributes, target_filespec, aOperateOnFolders, aDoRecurse, true);
		}
		if (file_search != INVALID_HANDLE_VALUE) // In case the loop had zero iterations.
			FindClose(file_search);
	}

	if (!aCalledRecursively) // i.e. Only need to do this if we're returning to top-level caller:
		g_ErrorLevel->Assign(failure_count); // i.e. indicate success if there were no failures.
	return failure_count;
}



ResultType Line::FileGetTime(char *aFilespec, char aWhichTime)
// Adapted from the AutoIt3 source.
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default
	output_var->Assign(); // Init to be blank, in case of failure.

	if (!aFilespec || !*aFilespec)
		return OK;  // Let ErrorLevel indicate an error, since this is probably not what the user intended.

	// Open existing file.  Uses CreateFile() rather than OpenFile for an expectation
	// of greater compatibility for all files, and folder support too.
	// FILE_FLAG_NO_BUFFERING might improve performance because all we're doing is
	// fetching one of the file's attributes.  FILE_FLAG_BACKUP_SEMANTICS must be
	// used, otherwise fetching the time of a directory under NT and beyond will
	// not succeed.  Win95 (not sure about Win98/ME) does not support this, but it
	// should be harmless to specify it even if the OS is Win95:
	HANDLE file_handle = CreateFile(aFilespec, GENERIC_READ, FILE_SHARE_READ, (LPSECURITY_ATTRIBUTES)NULL
		, OPEN_EXISTING, FILE_FLAG_NO_BUFFERING | FILE_FLAG_BACKUP_SEMANTICS, NULL);
	if (file_handle == INVALID_HANDLE_VALUE)
		return OK;  // Let ErrorLevel Tell the story.

	FILETIME ftCreationTime, ftLastAccessTime, ftLastWriteTime;
	if (GetFileTime(file_handle, &ftCreationTime, &ftLastAccessTime, &ftLastWriteTime))
		CloseHandle(file_handle);
	else
	{
		CloseHandle(file_handle);
		return OK;  // Let ErrorLevel Tell the story.
	}

	FILETIME local_file_time;
	switch (toupper(aWhichTime))
	{
	case 'C': // File's creation time.
		FileTimeToLocalFileTime(&ftCreationTime, &local_file_time);
		break;
	case 'A': // File's last access time.
		FileTimeToLocalFileTime(&ftLastAccessTime, &local_file_time);
		break;
	default:  // 'M', unspecified, or some other value.  Use the file's modification time.
		FileTimeToLocalFileTime(&ftLastWriteTime, &local_file_time);
	}

    g_ErrorLevel->Assign(ERRORLEVEL_NONE);  // Indicate success.
	char local_file_time_string[128];
	return output_var->Assign(FileTimeToYYYYMMDD(local_file_time_string, &local_file_time));
}



int Line::FileSetTime(char *aYYYYMMDD, char *aFilePattern, char aWhichTime
	, FileLoopModeType aOperateOnFolders, bool aDoRecurse, bool aCalledRecursively)
// Returns the number of files and folders that could not be changed due to an error.
// Current limitation: It will not recurse into subfolders unless their names also match
// aFilePattern.
{
	// Making this static is not friendly to reentrant calls to this function (i.e. calls maded
	// as a consequence of the current script subroutine being interrupted by another during
	// this instance's MsgSleep()).  However, it doesn't seem to be that much of a consequence
	// since the exact interval period of the MsgSleep()'s isn't that important.  It's also
	// pretty unlikely that the interrupting subroutine will also just happen to call this
	// function rather than some other.  On the other hand, we don't do failure_count as a
	// static because we want its value to be 100% accurate even if reentrancy occurs:
	static int files_considered;
	if (!aCalledRecursively)  // i.e. Only need to do this if we're not called by ourself:
	{
		files_considered = 0;
		g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default
		if (!aFilePattern || !*aFilePattern)
			return 0;  // Let ErrorLevel indicate an error, since this is probably not what the user intended.
		if (aOperateOnFolders == FILE_LOOP_INVALID) // In case runtime dereference of a var was an invalid value.
			aOperateOnFolders = FILE_LOOP_FILES_ONLY;  // Set default.
	}

	// Related to the comment at the top: Since the script subroutine that resulted in the call to
	// this function can be interrupted during our MsgSleep(), make a copy of any params that might
	// currently point directly to the deref buffer.  This is done because their contents might
	// be overwritten by the interrupting subroutine:
	char yyyymmdd[64]; // Even do this one since its value is passed recursively in calls to self.
	strlcpy(yyyymmdd, aYYYYMMDD, sizeof(yyyymmdd));
	char file_pattern[MAX_PATH * 2];  // In case OS supports extra-long filenames.
	strlcpy(file_pattern, aFilePattern, sizeof(file_pattern));

	FILETIME ft, ftUTC;
	if (*yyyymmdd)
	{
		// Convert the arg into the time struct as local (non-UTC) time:
		if (!YYYYMMDDToFileTime(yyyymmdd, &ft))
			return 0;  // Let ErrorLevel tell the story.
		// Convert from local to UTC:
		if (!LocalFileTimeToFileTime(&ft, &ftUTC))
			return 0;  // Let ErrorLevel tell the story.
	}
	else // User wants to use the current time (i.e. now) as the new timestamp.
		GetSystemTimeAsFileTime(&ftUTC);

	// This following section is very similar to that in FileSetAttrib and FileDelete:
	// Give extra room in some of these vars, in case OS supports extra-long filenames:
	char file_path[MAX_PATH * 2];
	char target_filespec[MAX_PATH * 2];
	if (strlen(aFilePattern) >= sizeof(file_path))
		return 0; // Return OK because this is non-critical.  Let the above ErrorLevel indicate the problem.
	strlcpy(file_path, aFilePattern, sizeof(file_path));
	char *last_backslash = strrchr(file_path, '\\');
	// Remove the filename and/or wildcard part.   But leave the trailing backslash on it for
	// consistency with below:
	if (last_backslash)
		*(last_backslash + 1) = '\0';
	else // Use current working directory, e.g. if user specified only *.*
		*file_path = '\0';

	// For use with aDoRecurse, get just the naked file name/pattern:
	char *naked_filename_or_pattern = strrchr(file_pattern, '\\');
	if (naked_filename_or_pattern)
		++naked_filename_or_pattern;
	else
		naked_filename_or_pattern = file_pattern;

	if (!StrChrAny(naked_filename_or_pattern, "?*"))
		// Since no wildcards, always operate on this single item even if it's a folder.
		aOperateOnFolders = FILE_LOOP_FILES_AND_FOLDERS;

	int batch_size, sleep_duration;
	if (Hotkey::HookIsActive())
	{
		batch_size = 100; // i.e. try to minimize keyboard & mouse lag by sleeping more often.
		sleep_duration = 10;
	}
	else
	{
		batch_size = 1000;
		sleep_duration = -1; // Since all we want to do is check messages.
	}

	WIN32_FIND_DATA current_file;
	HANDLE file_search = FindFirstFile(aFilePattern, &current_file);
	bool file_found = (file_search != INVALID_HANDLE_VALUE);
	int failure_count = 0;
	HANDLE hFile;

	for (; file_found; file_found = FindNextFile(file_search, &current_file), ++files_considered)
	{
		if (files_considered && !(files_considered % batch_size))
			MsgSleep(sleep_duration); // Allows the program to be responsive even during a very long operation.
		if (current_file.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
		{
			if (!strcmp(current_file.cFileName, "..") || !strcmp(current_file.cFileName, "."))
				continue; // Never operate upon or recurse into these.
			// Regardless of whether this folder was recursed-into by the above, if this param
			// is false, its own attributes will not be affected:
			if (aOperateOnFolders == FILE_LOOP_FILES_ONLY)
				continue;
		}
		else // It's a file, not a folder.
			if (aOperateOnFolders == FILE_LOOP_FOLDERS_ONLY)
				continue;

		snprintf(target_filespec, sizeof(target_filespec), "%s%s", file_path, current_file.cFileName);
		hFile = CreateFile(target_filespec, GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE
			, (LPSECURITY_ATTRIBUTES)NULL, OPEN_EXISTING
			, FILE_FLAG_NO_BUFFERING | FILE_FLAG_BACKUP_SEMANTICS, NULL);
		if (hFile == INVALID_HANDLE_VALUE)
		{
			++failure_count;
			continue;
		}
		switch (toupper(aWhichTime))
		{
		case 'C': // File's creation time.
			if (!SetFileTime(hFile, &ftUTC, NULL, NULL))
				++failure_count;
			break;
		case 'A': // File's last access time.
			if (!SetFileTime(hFile, NULL, &ftUTC, NULL))
				++failure_count;
			break;
		default:  // 'M', unspecified, or some other value.  Use the file's modification time.
			if (!SetFileTime(hFile, NULL, NULL, &ftUTC))
				++failure_count;
		}
		CloseHandle(hFile);
	}

	if (file_search != INVALID_HANDLE_VALUE) // In case the loop had zero iterations.
		FindClose(file_search);

	// This section is identical to that in FileSetAttrib() except for the recursive function
	// call itself, so see comments there for details:
	if (aDoRecurse) 
	{
		char all_file_pattern[MAX_PATH * 2];
		snprintf(all_file_pattern, sizeof(all_file_pattern), "%s*.*", file_path);
		file_search = FindFirstFile(all_file_pattern, &current_file);
		file_found = (file_search != INVALID_HANDLE_VALUE);
		for (; file_found; file_found = FindNextFile(file_search, &current_file), ++files_considered)
		{
			if (files_considered && !(files_considered % batch_size))
				MsgSleep(sleep_duration);
			if (!(current_file.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
				continue;
			if (!strcmp(current_file.cFileName, "..") || !strcmp(current_file.cFileName, "."))
				continue;
			snprintf(target_filespec, sizeof(target_filespec), "%s%s\\%s"
				, file_path, current_file.cFileName, naked_filename_or_pattern);
			failure_count += FileSetTime(yyyymmdd, target_filespec, aWhichTime, aOperateOnFolders, aDoRecurse, true);
		}
		if (file_search != INVALID_HANDLE_VALUE) // In case the loop had zero iterations.
			FindClose(file_search);
	}

	if (!aCalledRecursively) // i.e. Only need to do this if we're returning to top-level caller:
		g_ErrorLevel->Assign(failure_count); // i.e. indicate success if there were no failures.
	return failure_count;
}



ResultType Line::FileGetSize(char *aFilespec, char *aGranularity)
// Adapted from the AutoIt3 source.
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default
	output_var->Assign(); // Init to be blank, in case of failure.

	if (!aFilespec || !*aFilespec)
		return OK;  // Let ErrorLevel indicate an error, since this is probably not what the user intended.

	// Open existing file.  Uses CreateFile() rather than OpenFile for an expectation
	// of greater compatibility for all files, and folder support too.
	// FILE_FLAG_NO_BUFFERING might improve performance because all we're doing is
	// fetching one of the file's attributes.  FILE_FLAG_BACKUP_SEMANTICS must be
	// used, otherwise fetching the time of a directory under NT and beyond will
	// not succeed.  Win95 (not sure about Win98/ME) does not support this, but it
	// should be harmless to specify it even if the OS is Win95:
	HANDLE file_handle = CreateFile(aFilespec, GENERIC_READ, FILE_SHARE_READ, (LPSECURITY_ATTRIBUTES)NULL
		, OPEN_EXISTING, FILE_FLAG_NO_BUFFERING | FILE_FLAG_BACKUP_SEMANTICS, NULL);
	if (file_handle == INVALID_HANDLE_VALUE)
		return OK;  // Let ErrorLevel Tell the story.
	unsigned __int64 size = GetFileSize64(file_handle);
	if (size == 0xFFFFFFFFFFFFFFFF) // failure
		return OK;  // Let ErrorLevel Tell the story.
	CloseHandle(file_handle);

	switch(toupper(*aGranularity))
	{
	case 'K': // KB
		size /= 1024;
		break;
	case 'M': // MB
		size /= (1024 * 1024);
		break;
	// default: // i.e. either 'B' for bytes, or blank, or some other unknown value, so default to bytes.
		// do nothing
	}

    g_ErrorLevel->Assign(ERRORLEVEL_NONE);  // Indicate success.
	return output_var->Assign((int)(size > MAXDWORD ? -1 : size));
	// Currently, the above is basically subject to a 2 gig limit, I believe, after which the
	// size will appear to be negative.  Beyond a 4 gig limit, the value will probably wrap around
	// to zero and start counting from there as file sizes grow beyond 4 gig (UPDATE: The size
	// is now set to -1 [the maximum DWORD when expressed as a signed int] whenever >4 gig).
	// There's not much sense in putting values larger than 2 gig into the var as a text string
	// containing a positive number because such a var could never be properly handled by anything
	// that compares to it (e.g. IfGreater) or does math on it (e.g. EnvAdd), since those operations
	// use atoi() to convert the string.  So as a future enhancement (unless the whole program is
	// revamped to use 64bit ints or something) might add an optional param to the end to indicate
	// size should be returned in K(ilobyte) or M(egabyte).  However, this is sorta bad too since
	// adding a param can break existing scripts which use filenames containing commas (delimiters)
	// with this command.  Therefore, I think I'll just add the K/M param now.
	// Also, the above assigns an int because unsigned ints should never be stored in script
	// variables.  This is because an unsigned variable larger than INT_MAX would not be properly
	// converted by atoi(), which is current standard method for variables to be auto-converted
	// from text back to a number whenever that is needed.
}



ResultType Line::FileGetVersion(char *aFilespec)
// Adapted from the AutoIt3 source.
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default
	output_var->Assign(); // Init to be blank, in case of failure.

	if (!aFilespec || !*aFilespec)
		return OK;  // Let ErrorLevel indicate an error, since this is probably not what the user intended.

	DWORD dwUnused, dwSize;
	if (   !(dwSize = GetFileVersionInfoSize(aFilespec, &dwUnused))   )
		return OK;  // Let ErrorLevel tell the story.

	BYTE *pInfo = (BYTE*)malloc(dwSize);  // Allocate the size retrieved by the above.

	// Read the version resource
	GetFileVersionInfo((LPSTR)aFilespec, 0, dwSize, (LPVOID)pInfo);

	// Locate the fixed information
	VS_FIXEDFILEINFO *pFFI;
	UINT uSize;
	if (!VerQueryValue(pInfo, "\\", (LPVOID *)&pFFI, &uSize))
	{
		free(pInfo);
		return OK;  // Let ErrorLevel tell the story.
	}

	// extract the fields you want from pFFI
	UINT iFileMS = (UINT)pFFI->dwFileVersionMS;
	UINT iFileLS = (UINT)pFFI->dwFileVersionLS;
	char version_string[128];  // AutoIt3: 43+1 is the maximum size, but leave a little room to increase confidence.
	snprintf(version_string, sizeof(version_string), "%u.%u.%u.%u"
		, (iFileMS >> 16), (iFileMS & 0xFFFF), (iFileLS >> 16), (iFileLS & 0xFFFF));

	free(pInfo);

    g_ErrorLevel->Assign(ERRORLEVEL_NONE);  // Indicate success.
	return output_var->Assign(version_string);
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

int Line::ConvertEscapeChar(char *aFilespec, char aOldChar, char aNewChar)
{
	if (!aFilespec || !*aFilespec) return 1;  // Non-zero is failure in this case.
	if (aOldChar == aNewChar)
	{
		MsgBox("Conversion: The OldChar must not be the same as the NewChar.");
		return 1;
	}
	FILE *f1 = fopen(aFilespec, "r");
	if (!f1)
	{
		MsgBox(aFilespec, 0, "Could not open source file for conversion:");
		return 1; // Failure
	}
	char new_filespec[MAX_PATH * 2];
	strlcpy(new_filespec, aFilespec, sizeof(new_filespec));
	StrReplace(new_filespec, CONVERSION_FLAG, "-NEW" EXT_AUTOHOTKEY, false);
	FILE *f2 = fopen(new_filespec, "w");
	if (!f2)
	{
		fclose(f1);
		MsgBox(new_filespec, 0, "Could not open target file for conversion:");
		return 1; // Failure
	}

	char buf[LINE_SIZE], *cp, next_char;
	size_t line_count, buf_length;
	for (line_count = 0;;)
	{
		if (   -1 == (buf_length = ConvertEscapeCharGetLine(buf, (int)(sizeof(buf) - 1), f1))   )
			break;
		++line_count;

		for (cp = buf; ; ++cp)  // Increment to skip over the symbol just found by the inner for().
		{
			for (; *cp && *cp != aOldChar && *cp != aNewChar; ++cp);  // Find the next escape char.

			if (!*cp) // end of string.
				break;

			if (*cp == aNewChar)
			{
				if (buf_length < sizeof(buf) - 1)  // buffer safety check
				{
					// Insert another of the same char to make it a pair, so that it becomes the
					// literal version of this new escape char (e.g. ` --> ``)
					MoveMemory(cp + 1, cp, strlen(cp) + 1);
					*cp = aNewChar;
					// Increment so that the loop will resume checking at the char after this new pair.
					// Example: `` becomes ````
					++cp;  // Only +1 because the outer for-loop will do another increment.
					++buf_length;
				}
				continue;
			}

			// Otherwise *cp == aOldChar:
			next_char = *(cp + 1);
			if (next_char == aOldChar)
			{
				// This is a double-escape (e.g. \\ in AutoIt2).  Replace it with a single
				// character of the same type:
				MoveMemory(cp, cp + 1, strlen(cp + 1) + 1);
				--buf_length;
			}
			else
				// It's just a normal escape sequence.  Even if it's not a valid escape sequence,
				// convert it anyway because it's more failsafe to do so (the script parser will
				// handle such things much better than we can when the script is run):
				*cp = aNewChar;
		}
		fputs(buf, f2);
	}

	fclose(f1);
	fclose(f2);
	MsgBox("The file was successfully converted.");
	return 0;  // Return 0 on success in this case.
}



size_t Line::ConvertEscapeCharGetLine(char *aBuf, int aMaxCharsToRead, FILE *fp)
{
	if (!aBuf || !fp) return -1;
	if (aMaxCharsToRead <= 0) return 0;
	if (feof(fp)) return -1; // Previous call to this function probably already read the last line.
	if (fgets(aBuf, aMaxCharsToRead, fp) == NULL) // end-of-file or error
	{
		*aBuf = '\0';  // Reset since on error, contents added by fgets() are indeterminate.
		return -1;
	}
	return strlen(aBuf);
}



ArgTypeType Line::ArgIsVar(ActionTypeType aActionType, int aArgIndex)
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
		case ACT_STRINGLOWER:
		case ACT_STRINGUPPER:
		case ACT_STRINGLEN:
		case ACT_STRINGREPLACE:
		case ACT_STRINGGETPOS:
		case ACT_GETKEYSTATE:
		case ACT_CONTROLGETFOCUS:
		case ACT_CONTROLGETTEXT:
		case ACT_STATUSBARGETTEXT:
		case ACT_INPUTBOX:
		case ACT_RANDOM:
		case ACT_INIREAD:
		case ACT_REGREAD:
		case ACT_DRIVESPACEFREE:
		case ACT_FILEREADLINE:
		case ACT_FILEGETATTRIB:
		case ACT_FILEGETTIME:
		case ACT_FILEGETSIZE:
		case ACT_FILEGETVERSION:
		case ACT_FILESELECTFILE:
		case ACT_FILESELECTFOLDER:
		case ACT_MOUSEGETPOS:
		case ACT_WINGETTITLE:
		case ACT_WINGETTEXT:
		case ACT_WINGETPOS:
		case ACT_PIXELGETCOLOR:
		case ACT_PIXELSEARCH:
			return ARG_TYPE_OUTPUT_VAR;

		case ACT_IFINSTRING:
		case ACT_IFNOTINSTRING:
		case ACT_IFEQUAL:
		case ACT_IFNOTEQUAL:
		case ACT_IFGREATER:
		case ACT_IFGREATEROREQUAL:
		case ACT_IFLESS:
		case ACT_IFLESSOREQUAL:
			return ARG_TYPE_INPUT_VAR;
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
		case ACT_STRINGLOWER:
		case ACT_STRINGUPPER:
		case ACT_STRINGLEN:
		case ACT_STRINGREPLACE:
		case ACT_STRINGGETPOS:
			return ARG_TYPE_INPUT_VAR;

		case ACT_MOUSEGETPOS:
		case ACT_WINGETPOS:
		case ACT_PIXELSEARCH:
			return ARG_TYPE_OUTPUT_VAR;
		}
		break;

	case 2:  // Arg #3
		if (aActionType == ACT_WINGETPOS)
			return ARG_TYPE_OUTPUT_VAR;
		break;

	case 3:  // Arg #4
		if (aActionType == ACT_WINGETPOS)
			return ARG_TYPE_OUTPUT_VAR;
		break;
	}
	// Otherwise:
	return ARG_TYPE_NORMAL;
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
	case ACT_WINMENUSELECTITEM:
		// Window params can all be blank in this case, but the first menu param should
		// be non-blank (but it's ok if its a dereferenced var that resolves to blank
		// at runtime):
		if (!*RAW_ARG3)
			return LineError("Parameter #3 must not be blank.");
		return OK;
	case ACT_MOUSECLICKDRAG:
		// Even though we check for blanks at load-time, we don't bother to do so at runtime
		// (i.e. if a dereferenced var resolved to blank, it will be treated as a zero):
		if (!*RAW_ARG4 || !*RAW_ARG5)
			return LineError("Parameters 4 and 5 must specify a non-blank destination for the drag.");
		return OK;
	case ACT_MOUSEGETPOS:
		if (!ARG_HAS_VAR(1) && !ARG_HAS_VAR(2))
			return LineError(ERR_MISSING_OUTPUT_VAR);
		return OK;
	case ACT_WINGETPOS:
		if (!ARG_HAS_VAR(1) && !ARG_HAS_VAR(2) && !ARG_HAS_VAR(3) && !ARG_HAS_VAR(4))
			return LineError(ERR_MISSING_OUTPUT_VAR);
		return OK;
	case ACT_PIXELSEARCH:
		if (!*RAW_ARG3 || !*RAW_ARG4 || !*RAW_ARG5 || !*RAW_ARG6 || !*RAW_ARG7)
			return LineError("Parameters 3 through 7 must not be blank.");
		return OK;
	case ACT_REGREAD:
		if (!*RAW_ARG3 || !*RAW_ARG4) // But #5 can be blank in case the command is being used in 4-param mode.
			return LineError("Parameters 3 and 4 must not be blank.");
		return OK;
	}
	return OK;  // For when the command isn't mentioned in the switch().
}



bool Line::FileIsFilteredOut(WIN32_FIND_DATA &aCurrentFile, FileLoopModeType aFileLoopMode, char *aFilePath)
{
	if (aCurrentFile.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) // It's a folder.
	{
		if (aFileLoopMode == FILE_LOOP_FILES_ONLY
			|| !strcmp(aCurrentFile.cFileName, "..") || !strcmp(aCurrentFile.cFileName, "."))
			return true; // Exclude this folder by returning true.
	}
	else // it's not a folder.
		if (aFileLoopMode == FILE_LOOP_FOLDERS_ONLY)
			return true; // Exclude this file by returning true.

	// Since file was found, also prepend the file's path to its name for the caller:
	if (!aFilePath || !*aFilePath) // don't bother.
		return false;
	char temp[MAX_PATH];
	strlcpy(temp, aCurrentFile.cFileName, sizeof(temp));
	snprintf(aCurrentFile.cFileName, sizeof(aCurrentFile.cFileName), "%s%s", aFilePath, temp);
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
