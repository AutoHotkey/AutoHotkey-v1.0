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
#include "application.h" // for MsgSleep()
#include "window.h" // for SetForegroundWindowEx()


ResultType Script::PerformGui(char *aCommand, char *aParam2, char *aParam3, char *aParam4)
{
	int window_index = g.DefaultGuiIndex; // Which window to operate upon.  Initialized to thread's default.
	char *options; // This will contain something that is meaningful only when gui_command == GUI_CMD_OPTIONS.
	GuiCommands gui_command = Line::ConvertGuiCommand(aCommand, &window_index, &options);
	if (gui_command == GUI_CMD_INVALID)
		return ScriptError(ERR_GUICOMMAND ERR_ABORT, aCommand);
	if (window_index < 0 || window_index >= MAX_GUI_WINDOWS)
		return ScriptError("The window number must be between 1 and " MAX_GUI_WINDOWS_STR
			"." ERR_ABORT, aCommand);

	// First completely handle any sub-command that doesn't require the window to exist.
	// In other words, don't auto-create the window before doing this command like we do
	// for the others:
	if (gui_command == GUI_CMD_DESTROY)
		return GuiType::Destroy(window_index);

	// If the window doesn't currently exist, don't auto-create it for those commands for
	// which it wouldn't make sense. Note that things like FONT and COLOR are allowed to
	// auto-create the window, since those commands can be legitimately used prior to the
	// first "Gui Add" command.  Also, it seems best to allow SHOW even though all it will
	// do is create and display an empty window.
	if (!g_gui[window_index])
	{
		switch(gui_command)
		{
		case GUI_CMD_SUBMIT:
		case GUI_CMD_CANCEL:
			return OK; // Nothing needs to be done since the window object doesn't exist.
		}
		// Otherwise: Create the object and (later) its window, since all the other sub-commands below need it:
		if (   !(g_gui[window_index] = new GuiType(window_index))   )
			return FAIL; // No error displayed since extremely rare.
		// Probably better to increment here rather than in constructor in case GuiType objects
		// are ever created outside of the g_gui array (such as for temp local variables):
		++GuiType::sObjectCount; // This count is maintained to help performance in the main event loop and other places.
	}

	GuiType &gui = *g_gui[window_index];  // For performance and convenience.

	// Now handle any commands that should be handled prior to creation of the window in the case
	// where the window doesn't already exist:
	if (gui_command == GUI_CMD_OPTIONS)
	{
		char *next_option, *option_end;
		bool option_is_being_removed;
		int owner_window_index;
		DWORD style_orig = gui.mStyle;
		for (next_option = options; *next_option; next_option = omit_leading_whitespace(option_end))
		{
			// Below: assume option is being added in the absence of either sign.  However, the first
			// option in the list must begin with +/- otherwise the cmd would never have been properly
			// detected as GUI_CMD_OPTIONS in the first place.
			option_is_being_removed = (*next_option == '-');
			next_option = omit_leading_whitespace(++next_option);
			if (   !(option_end = StrChrAny(next_option, " \t"))   )  // Space or tab.
				option_end = next_option + strlen(next_option); // Set to position of zero terminator instead.
			if (!strnicmp(next_option, "Owner", 5))
			{
				if (gui.mHwnd) // OS provides no way to change an existing window's owner.
					continue;   // Currently no effect, as documented.
				if (option_is_being_removed)
					gui.mOwner = NULL;
				else
				{
					if (option_end - next_option > 5) // Length is greater than 5, so it has a number (e.g. Owned1).
					{
						// Using ATOI() vs. atoi() seems okay in these cases since spaces are required
						// between options:
						owner_window_index = ATOI(next_option + 5) - 1;
						if (owner_window_index >= 0 && owner_window_index < MAX_GUI_WINDOWS
							&& owner_window_index != window_index  // Window can't own itself!
							&& g_gui[owner_window_index] && g_gui[owner_window_index]->mHwnd) // Relies on short-circuit boolean order.
							gui.mOwner = g_gui[owner_window_index]->mHwnd;
						else
							return ScriptError("The owner window is not valid or does not yet exist." ERR_ABORT, next_option);
					}
					else
						gui.mOwner = g_hWnd; // Make a window owned (by script's main window) omits its taskbar button.
				}
			}
			else if (!strnicmp(next_option, "Resize", 6))
			{
				if (option_is_being_removed)
					gui.mStyle &= ~(WS_SIZEBOX|WS_MAXIMIZEBOX); // Remove either or both.
				else
					gui.mStyle |= WS_SIZEBOX|WS_MAXIMIZEBOX;
			}
		}

		// Besides reducing the code size and complexity, another reason all changes to style are made
		// here rather than above is that multiple changes might have been made above to the style,
		// and there's no point in redrawing/updating the window for each one:
		if (gui.mHwnd && gui.mStyle != style_orig)
		{
			// Since window already exists but its style has changed, attempt to update it dynamically.
			SetWindowLong(gui.mHwnd, GWL_STYLE, gui.mStyle);
			// Hiding then showing is the only way I've discovered to make it update.  If the window
			// is not updated, a strange effect occurs where the window is still visible but can no
			// longer be used at all (clicks pass right through it).  This show/hide method is less
			// desirable due to possible side effects caused to any script that happens to be watching
			// for its existence/non-existence, so it would be nice if some better way can be discovered
			// do do this.
			// SetWindowPos is also necessary, otherwise the frame thickness entirely around the window
			// does not get updated (just parts of it):
			SetWindowPos(gui.mHwnd, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
			ShowWindow(gui.mHwnd, SW_HIDE);
			ShowWindow(gui.mHwnd, SW_SHOWNOACTIVATE); // i.e. don't activate it if it wasn't before.
			// None of the following methods alone is enough, at least not when the window is currently active:
			// 1) InvalidateRect() on the window's entire rect.
			// 2) SendMessage(gui.mHwnd, WM_NCPAINT, 1, 0);  // 1 = Repaint entire frame.
			// 3) RedrawWindow(gui.mHwnd, NULL, NULL, RDW_INVALIDATE|RDW_FRAME|RDW_UPDATENOW);
			// 4) SetWindowPos(gui.mHwnd, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
		}

		// Continue on to create the window so that code is simplified in other places by
		// using the assumption that "if gui[i] object exists, so does its window".
		// Another important reason this is done is that if an owner window were to be destroyed
		// before the window it owns is actually created, the WM_DESTROY logic would have to check
		// for any windows owned by the window being destroyed and update them.
	}

	// Create the window if needed.  Since it should not be possible for our window to get destroyed
	// without our knowning about it (via the explicit handling in its window proc), it shouldn't
	// be necessary to check the result of IsWindow(gui.mHwnd):
	if (!gui.mHwnd)
		if (!gui.Create())
		{
			GuiType::Destroy(window_index); // Get rid of the object so that it stays in sync with the window's existence.
			return ScriptError("Could not create window." ERR_ABORT);
		}

	// After creating the window, return from any commands that were fully handled above:
	if (gui_command == GUI_CMD_OPTIONS)
		return OK;

	GuiControls gui_control_type = GUI_CONTROL_INVALID;

	switch(gui_command)
	{
	case GUI_CMD_ADD:
		if (   !(gui_control_type = Line::ConvertGuiControl(aParam2))   )
			return ScriptError(ERR_GUICONTROL ERR_ABORT, aParam2);
		return gui.AddControl(gui_control_type, aParam3, aParam4); // It already displayed any error.

	case GUI_CMD_MENU:
		UserMenu *menu;
		if (*aParam2)
		{
			// By design, the below will give a slightly misleading error if the specified menu is the
			// TRAY menu, since it should be obvious that it cannot be used as a menu bar (since it
			// must always be of the popup type):
			if (   !(menu = FindMenu(aParam2)) || menu == g_script.mTrayMenu   ) // Relies on short-circuit boolean.
				return ScriptError(ERR_MENU ERR_ABORT, aParam2);
			menu->Create(MENU_TYPE_BAR);  // Ensure the menu physically exists and is the "non-popup" type (for a menu bar).
		}
		else
			menu = NULL;
		SetMenu(gui.mHwnd, menu ? menu->mMenu : NULL);  // Add or remove the menu.
		return OK;

	case GUI_CMD_SHOW:
		return gui.Show(aParam2, aParam3);

	case GUI_CMD_SUBMIT:
		return gui.Submit(stricmp(aParam2, "NoHide"));

	case GUI_CMD_CANCEL:
		return gui.Cancel();

	case GUI_CMD_FONT:
		return gui.SetCurrentFont(aParam2, aParam3);

	case GUI_CMD_COLOR:
		// AssignColor() takes care of deleting old brush, etc.
		// In this case, a blank for either param means "leaving existing color alone", in which
		// case AssignColor() is not called since it would assume CLR_NONE then.
		if (*aParam2)
			AssignColor(aParam2, gui.mBackgroundColorWin, gui.mBackgroundBrushWin);
		if (*aParam3)
			AssignColor(aParam3, gui.mBackgroundColorCtl, gui.mBackgroundBrushCtl);
		if (IsWindowVisible(gui.mHwnd))
		{
			// Force the window to repaint so that colors take effect immediately.
			// UpdateWindow() isn't enough sometimes/always, so do something more aggressive:
			RECT client_rect;
			GetClientRect(gui.mHwnd, &client_rect);
			InvalidateRect(gui.mHwnd, &client_rect, TRUE);
		}
		return OK;

	} // switch()

	return FAIL;  // Should never be reached, but avoids compiler warning and improves bug detection.
}



ResultType Line::GuiControl(char *aCommand, char *aControlID, char *aParam3)
{
	char *options; // This will contain something that is meaningful only when gui_command == GUICONTROL_CMD_OPTIONS.
	int window_index = g.DefaultGuiIndex; // Which window to operate upon.  Initialized to thread's default.
	GuiControlCmds guicontrol_cmd = Line::ConvertGuiControlCmd(aCommand, &window_index, &options);
	if (guicontrol_cmd == GUICONTROL_CMD_INVALID)
		// This is caught at load-time 99% of the time and can only occur here if the sub-command name
		// is contained in a variable reference.  Since it's so rare, the handling of it is debatable,
		// but to keep it simple just set ErrorLevel:
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	if (window_index < 0 || window_index >= MAX_GUI_WINDOWS || !g_gui[window_index]) // Relies on short-circuit boolean order.
		// This departs from the tradition used by PerformGui() but since this type of error is rare,
		// and since use ErrorLevel adds a little bit of flexibility (since the script's curretn thread
		// is not unconditionally aborted), this seems best:
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);

	GuiType &gui = *g_gui[window_index];  // For performance and convenience.
	UINT control_index = gui.FindControl(aControlID);
	if (control_index >= gui.mControlCount) // Not found.
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	GuiControlType &control = gui.mControl[control_index];   // For performance and convenience.

	// Beyond this point, errors are rare so set the default to "no error":
	g_ErrorLevel->Assign(ERRORLEVEL_NONE);

	char *malloc_buf;
	RECT rect;
	WPARAM checked;
	UINT radio_start, radio_end;

	switch(guicontrol_cmd)
	{
	case GUICONTROL_CMD_OPTIONS: // Currently does nothing.
		return OK;
	case GUICONTROL_CMD_CONTENTS:
		switch (control.type)
		{
		case GUI_CONTROL_EDIT:
			malloc_buf = (*aParam3 && (GetWindowLong(control.hwnd, GWL_STYLE) & ES_MULTILINE))
				? TranslateLFtoCRLF(aParam3) : NULL; // Automatic translation, as documented.
			SetWindowText(control.hwnd,  malloc_buf ? malloc_buf : aParam3);
			if (malloc_buf)
				free(malloc_buf);
			return OK;

		case GUI_CONTROL_PIC:
			// Update: The below doesn't work, so it will be documented that a picture control
			// should be always be referred to by its original filename even if the picture changes.
			// Set the text unconditionally even if the picture can't be loaded.  This text must
			// be set to allow GuiControl(Get) to be able to operate upon the picture without
			// needing to indentify it via something like "Static14".
			//SetWindowText(control.hwnd, aParam3);
			//SendMessage(control.hwnd, WM_SETTEXT, 0, (LPARAM)aParam3);
			// Although all HBITMAPs are freed upon program termination, if the program changes
			// the picture frequently, memory/resources would continue to rise in the meantime
			// unless this is done:
			if (control.hbitmap)
				DeleteObject(control.hbitmap);
			GetWindowRect(control.hwnd, &rect);
			// See comments in AddControl():
			if (control.hbitmap = LoadPicture(aParam3, rect.right - rect.left, rect.bottom - rect.top)) // Assign.
				// LoadPicture() uses CopyImage() to scale the image, which seems to provide better scaling
				// quality than using MoveWindow() (followed by redrawing the parent window) on the static
				// control that contains the image.
				SendMessage(control.hwnd, (UINT)STM_SETIMAGE, (WPARAM)IMAGE_BITMAP, (LPARAM)control.hbitmap);
			else
				return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
			return OK;

		case GUI_CONTROL_CHECKBOX:
		case GUI_CONTROL_RADIO:
			if (IsPureNumeric(aParam3, true, false))
			{
				checked = ATOI(aParam3);
				if (!checked || checked == 1 || (control.type == GUI_CONTROL_CHECKBOX && checked == -1))
				{
					if (checked == -1)
						checked = BST_INDETERMINATE;
					//else the "checked" var is already set correctly.
					if (control.type == GUI_CONTROL_RADIO && checked == BST_CHECKED)
					{
						// If this radio's group has more than one radio, ensure all other radios
						// in the group are unchecked.  This has been verified to be necessary even
						// when the BS_AUTORADIOBUTTON style is present:
						if (gui.FindGroup(control_index, radio_start, radio_end) > 1)
						{
							for (UINT u = radio_start; u < radio_end; ++u)
								if (gui.mControl[u].type == GUI_CONTROL_RADIO) // Since a group can have non-radios in it.
									SendMessage(gui.mControl[u].hwnd, BM_SETCHECK
										, (u == control_index) ? BST_CHECKED : BST_UNCHECKED, 0);
							return OK;
						}
					}
					// Otherwise, single radio or checkbox is being affected:
					SendMessage(control.hwnd, BM_SETCHECK, checked, 0);
					return OK;
				}
				//else the default SetWindowText() action will be taken below.
			}
			// else assume it's the text/caption for the item, so the default SetWindowText() action will be taken below.
			break;

		case GUI_CONTROL_DROPDOWNLIST:
		case GUI_CONTROL_COMBOBOX:
		case GUI_CONTROL_LISTBOX:
			if (*aParam3 == '|') // The signal to overwrite rather than append to the list.
			{
				++aParam3;  // Exclude the initial pipe from further consideration.
				SendMessage(control.hwnd, (control.type == GUI_CONTROL_LISTBOX) ? LB_RESETCONTENT : CB_RESETCONTENT
					, 0, 0); // Delete all items currently in the list.
			}
			gui.AddControlContent(control, aParam3, 0);
			return OK;
		} // inner switch() for control's type
		// Since above didn't return, it's either:
		// 1) A control that uses the standard SetWindowText() method such as GUI_CONTROL_TEXT,
		//    GUI_CONTROL_GROUPBOX, or GUI_CONTROL_BUTTON.
		// 2) A radio or checkbox whose caption is being changed instead of its checked state.
		SetWindowText(control.hwnd, aParam3);
		return OK;

	case GUICONTROL_CMD_MOVE:
	{
		int xpos = COORD_UNSPECIFIED;
		int ypos = COORD_UNSPECIFIED;
		int width = COORD_UNSPECIFIED;
		int height = COORD_UNSPECIFIED;
		for (char *cp = aParam3; *cp; ++cp)
		{
			switch(toupper(*cp))
			{
			// For options such as W, H, X and Y:
			// Use atoi() vs. ATOI() to avoid interpreting something like 0x01B as hex when in fact
			// the B was meant to be an option letter (though in this case, none of the hex digits are
			// currently used as option letters):
			case 'W':
				width = atoi(cp + 1);
				break;
			case 'H':
				height = atoi(cp + 1);
				break;
			case 'X':
				xpos = atoi(cp + 1);
				break;
			case 'Y':
				ypos = atoi(cp + 1);
				break;
			}
		}
		RECT rect;
		GetWindowRect(control.hwnd, &rect); // Failure seems to rare to check for.
		POINT pt = {rect.left, rect.top};
		ScreenToClient(gui.mHwnd, &pt);
		if (!MoveWindow(control.hwnd
			, xpos == COORD_UNSPECIFIED ? pt.x : xpos
			, ypos == COORD_UNSPECIFIED ? pt.y : ypos
			, width == COORD_UNSPECIFIED ? rect.right - rect.left : width
			, height == COORD_UNSPECIFIED ? rect.bottom - rect.top : height
			, TRUE))  // Do repaint.
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		return OK;
	}

	case GUICONTROL_CMD_FOCUS:
		return SetFocus(control.hwnd) ? OK : g_ErrorLevel->Assign(ERRORLEVEL_ERROR);

	case GUICONTROL_CMD_ENABLE:
		EnableWindow(control.hwnd, TRUE);
		return OK;

	case GUICONTROL_CMD_DISABLE:
		EnableWindow(control.hwnd, FALSE);
		return OK;

	case GUICONTROL_CMD_SHOW:
		ShowWindow(control.hwnd, SW_SHOWNOACTIVATE);
		return OK;

	case GUICONTROL_CMD_HIDE:
		ShowWindow(control.hwnd, SW_HIDE);
		return OK;

	case GUICONTROL_CMD_CHOOSE:
	case GUICONTROL_CMD_CHOOSESTRING:
	{
		UINT msg, x_msg, y_msg;
		switch(control.type)
		{
		case GUI_CONTROL_DROPDOWNLIST:
		case GUI_CONTROL_COMBOBOX:
			msg = (guicontrol_cmd == GUICONTROL_CMD_CHOOSE) ? CB_SETCURSEL : CB_SELECTSTRING;
			x_msg = CBN_SELCHANGE;
			y_msg = CBN_SELENDOK;
			break;
		case GUI_CONTROL_LISTBOX:
			msg = (guicontrol_cmd == GUICONTROL_CMD_CHOOSE) ? LB_SETCURSEL : LB_SELECTSTRING;
			x_msg = LBN_SELCHANGE;
			y_msg = LBN_DBLCLK;
			break;
		default:  // Not a supported control type.
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		}
		int extra_actions = 0; // Set default.
		if (*aParam3 == '|') // First extra action.
		{
			++aParam3; // Omit this pipe char from further consideration below.
			++extra_actions;
		}
		if (*aParam3 == '|') // Second extra action.
		{
			++aParam3; // Omit this pipe char from further consideration below.
			++extra_actions;
		}
		if (guicontrol_cmd == GUICONTROL_CMD_CHOOSESTRING)
		{
			if (SendMessage(control.hwnd, msg, 1, (LPARAM)aParam3) == CB_ERR) // CB_ERR == LB_ERR
				return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		}
		else
		{
			int selection_index = ATOI(aParam3) - 1;
			if (selection_index < 0)
				return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
			if (SendMessage(control.hwnd, msg, selection_index, 0) == CB_ERR) // CB_ERR == LB_ERR
				return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		}
		int control_id = GUI_INDEX_TO_ID(control_index);
		if (extra_actions > 0)
			SendMessage(gui.mHwnd, WM_COMMAND, (WPARAM)MAKELONG(control_id, x_msg), (LPARAM)control.hwnd);
		if (extra_actions > 1)
			SendMessage(gui.mHwnd, WM_COMMAND, (WPARAM)MAKELONG(control_id, y_msg), (LPARAM)control.hwnd);
		return OK;
	} // case
	} // switch()

	return FAIL;  // Should never be reached, but avoids compiler warning and improves bug detection.
}



ResultType Line::GuiControlGet(char *aCommand, char *aControlID, char *aParam3)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL; // ErrorLevel is not used in this case since it's an unexpected and critical error.
	output_var->Assign(); // Set default to be blank for all commands, for consistency.

	int window_index = g.DefaultGuiIndex; // Which window to operate upon.  Initialized to thread's default.
	GuiControlGetCmds guicontrolget_cmd = Line::ConvertGuiControlGetCmd(aCommand, &window_index);
	if (guicontrolget_cmd == GUICONTROLGET_CMD_INVALID)
		// This is caught at load-time 99% of the time and can only occur here if the sub-command name
		// is contained in a variable reference.  Since it's so rare, the handling of it is debatable,
		// but to keep it simple just set ErrorLevel:
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	if (window_index < 0 || window_index >= MAX_GUI_WINDOWS || !g_gui[window_index]) // Relies on short-circuit boolean order.
		// This departs from the tradition used by PerformGui() but since this type of error is rare,
		// and since use ErrorLevel adds a little bit of flexibility (since the script's curretn thread
		// is not unconditionally aborted), this seems best:
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);

	GuiType &gui = *g_gui[window_index];  // For performance and convenience.
	if (!*aControlID) // In this case, default to the name of the output variable, as documented.
		aControlID = output_var->mName;

	// Beyond this point, errors are rare so set the default to "no error":
	g_ErrorLevel->Assign(ERRORLEVEL_NONE);

	// Handle GUICONTROLGET_CMD_FOCUS early since it doesn't need a specified ControlID:
	if (guicontrolget_cmd == GUICONTROLGET_CMD_FOCUS)
	{
		class_and_hwnd_type cah;
		cah.hwnd = GetFocus();
		if (!cah.hwnd)
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		if (!gui.FindControl(cah.hwnd))
		{
			cah.hwnd = GetParent(cah.hwnd); // Since ComboBoxes have children, try looking up parent if its a child.
			if (!(GetWindowLong(cah.hwnd, GWL_STYLE) & WS_CHILD))
				return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
			if (!gui.FindControl(cah.hwnd)) // Probably because GUI window isn't foreground.
				return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		}
		// This section is the same as that in ControlGetFocus():
		char class_name[WINDOW_CLASS_SIZE];
		cah.class_name = class_name;
		if (!GetClassName(cah.hwnd, class_name, sizeof(class_name) - 5)) // -5 to allow room for sequence number.
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		cah.class_count = 0;  // Init for the below.
		cah.is_found = false; // Same.
		EnumChildWindows(gui.mHwnd, EnumChildFindSeqNum, (LPARAM)&cah);
		if (!cah.is_found) // Should be impossible due to FindControl() having already found it above.
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		// Append the class sequence number onto the class name set the output param to be that value:
		snprintfcat(class_name, sizeof(class_name), "%d", cah.class_count);
		return output_var->Assign(class_name); // And leave ErrorLevel set to NONE.
	}

	UINT control_index = gui.FindControl(aControlID);
	if (control_index >= gui.mControlCount) // Not found.
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	GuiControlType &control = gui.mControl[control_index];   // For performance and convenience.

	switch(guicontrolget_cmd)
	{
	case GUICONTROLGET_CMD_CONTENTS:
		// Because the below returns FAIL only if a critical error occurred, g_ErrorLevel is
		// left at NONE as set above for all cases.
		return GuiType::ControlGetContents(*output_var, control, !stricmp(aParam3, "Text"));

	case GUICONTROLGET_CMD_POS:
	{
		// In this case, output_var itself is not used directly, but is instead used to:
		// 1) Help performance by giving us the location in the linked list of variables of
		//    where to find the X/Y/W/H "array elements".
		// 2) Simplify the code by avoiding the need to classify GuiControlGet's param #1
		//    as something that is only sometimes a variable.
		RECT rect;
		GetWindowRect(control.hwnd, &rect);
		POINT pt = {rect.left, rect.top};
		ScreenToClient(gui.mHwnd, &pt);  // Failure seems too rare to check for.
		// Make it longer than Max var name so that FindOrAddVar() will be able to spot and report
		// var names that are too long:
		char var_name[MAX_VAR_NAME_LENGTH + 20];
		Var *var;
		snprintf(var_name, sizeof(var_name), "%sX", output_var->mName);
		if (   !(var = g_script.FindOrAddVar(var_name, 0, output_var))   ) // Called with output_var to enhance performance.
			return FAIL;  // It will have already displayed the error.
		var->Assign(pt.x);
		snprintf(var_name, sizeof(var_name), "%sY", output_var->mName);
		if (   !(var = g_script.FindOrAddVar(var_name, 0, output_var))   ) // Called with output_var to enhance performance.
			return FAIL;  // It will have already displayed the error.
		var->Assign(pt.y);
		snprintf(var_name, sizeof(var_name), "%sW", output_var->mName);
		if (   !(var = g_script.FindOrAddVar(var_name, 0, output_var))   ) // Called with output_var to enhance performance.
			return FAIL;  // It will have already displayed the error.
		var->Assign(rect.right - rect.left);
		snprintf(var_name, sizeof(var_name), "%sH", output_var->mName);
		if (   !(var = g_script.FindOrAddVar(var_name, 0, output_var))   ) // Called with output_var to enhance performance.
			return FAIL;  // It will have already displayed the error.
		return var->Assign(rect.bottom - rect.top);
	}

	case GUICONTROLGET_CMD_ENABLED:
		// See commment below.
		return output_var->Assign(IsWindowEnabled(control.hwnd) ? "1" : "0");

	case GUICONTROLGET_CMD_VISIBLE:
		// From working on Window Spy, I seem to remember that IsWindowVisible() uses different standards
		// for determining visibility than simply checking for WS_VISIBLE is the control and its parent
		// window.  If so, it might be undocumented in MSDN.  It is mentioned here to explain why
		// this "visible" sub-cmd is kept separate from some figure command such as "GuiControlGet, Out, Style":
		// 1) The style method is cumbersome to script with since it requires bitwise operates afterward.
		// 2) IsVisible() uses a different standard of detection than simply checking WS_VISIBLE.
		return output_var->Assign(IsWindowVisible(control.hwnd) ? "1" : "0");
	} // switch()

	return FAIL;  // Should never be reached, but avoids compiler warning and improves bug detection.
}



/////////////////
// Static members
/////////////////
FontType GuiType::sFont[MAX_GUI_FONTS]; // Not intialized to help catch bugs.
int GuiType::sFontCount = 0;
int GuiType::sObjectCount = 0;



ResultType GuiType::Destroy(UINT aWindowIndex)
// Rather than deal with the confusion of an object destroying itself, this method is static
// and designed to deal with one particular window index in the g_gui array.
{
	if (aWindowIndex >= MAX_GUI_WINDOWS)
		return FAIL;
	if (!g_gui[aWindowIndex]) // It's already in the right state.
		return OK;
	GuiType &gui = *g_gui[aWindowIndex];  // For performance and convenience.
	UINT u;
	if (gui.mHwnd)
	{
		// First destroy any windows owned by this window, since they will be auto-destroyed
		// anyway due to their being owned.  By destroying them explicitly, the Destroy()
		// function is called recursively which keeps everything relatively neat.
		for (u = 0; u < MAX_GUI_WINDOWS; ++u)
			if (g_gui[u] && g_gui[u]->mOwner == gui.mHwnd)
				GuiType::Destroy(u);
		if (IsWindow(gui.mHwnd)) // If WM_DESTROY called us, the window might already be partially destroyed.
		{
			// If this window is using a menu bar but that menu is also used by some other window, first
			// detatch the menu so that it doesn't get auto-destroyed with the window.  This is done
			// unconditionally since such a menu will be automatically destroyed when the script exits
			// or when the menu is destroyed explicitly via the Menu command.  It also prevents any
			// submenus attached to the menu bar from being destroyed, since those submenus might be
			// also used by other menus (however, this is not really an issue since menus destroyed
			// would be automatically re-created upon next use).  But in the case of a window that
			// is currently using a menu bar, destroying that bar in conjunction with the destruction
			// of some other window might cause bad side effects on some/all OSes.
			ShowWindow(gui.mHwnd, SW_HIDE);  // Hide it to prevent re-drawing due to menu removal.
			SetMenu(gui.mHwnd, NULL);
			if (!gui.mDestroyWindowHasBeenCalled)
			{
				gui.mDestroyWindowHasBeenCalled = true;  // Signal the WM_DESTROY routine not to call us.
				DestroyWindow(gui.mHwnd);  // The WindowProc is immediately called and it now destroys the window.
			}
			// else WM_DESTROY was called by a function other than this one (possibly auto-destruct due to
			// being owned by script's main window), so it would be bad to call DestroyWindow() again since
			// it's already in progress.
		}
	}
	if (gui.mBackgroundBrushWin)
		DeleteObject(gui.mBackgroundBrushWin);
	if (gui.mBackgroundBrushCtl)
		DeleteObject(gui.mBackgroundBrushCtl);
	// It seems best to delete the bitmaps whenever the control changes to a new image or
	// whenever the control is destroyed.  Otherwise, if a control or its parent window is
	// destroyed and recreated many times, memory allocation would continue to grow from
	// all the abandoned pointers:
	for (u = 0; u < gui.mControlCount; ++u)
		if (gui.mControl[u].hbitmap)
			DeleteObject(gui.mControl[u].hbitmap);
	// Not necessary since the object itself is about to be destroyed:
	//gui.mHwnd = NULL;
	//gui.mControlCount = 0; // All child windows (controls) are automatically destroyed with parent.
	delete g_gui[aWindowIndex]; // After this, the gui var itself is invalid so should not be referenced.
	g_gui[aWindowIndex] = NULL;
	--sObjectCount; // This count is maintained to help performance in the main event loop and other places.
	// For simplicity and performance, any fonts used by a destroyed window are destroyed
	// only when the program terminates.  Another reason for this is that sometimes a destroyed window
	// is soon recreated to use the same fonts it did before.
	return OK;
}



ResultType GuiType::Create()
{
	if (mHwnd) // It already exists
		return FAIL;  // Seems best for now, since it shouldn't really be called this way.

	// Use a separate class for GUI, which gives it a separate WindowProc and allows it to be more
	// distinct when used with the ahk_class method of addressing windows.
	static sGuiInitialized = false;
	if (!sGuiInitialized)
	{
		HICON hIcon = LoadIcon(g_hInstance, MAKEINTRESOURCE(IDI_MAIN));
		WNDCLASSEX wc = {0};
		wc.cbSize = sizeof(wc);
		wc.lpszClassName = WINDOW_CLASS_GUI;
		wc.hInstance = g_hInstance;
		wc.lpfnWndProc = GuiWindowProc;
		wc.hIcon = hIcon;
		wc.hIconSm = hIcon;
		//wc.style = 0;  // CS_HREDRAW | CS_VREDRAW
		wc.hCursor = LoadCursor((HINSTANCE) NULL, IDC_ARROW);
		wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
		wc.cbWndExtra = DLGWINDOWEXTRA;  // So that it will be the type that uses DefDlgProc() vs. DefWindowProc().
		if (!RegisterClassEx(&wc))
		{
			MsgBox("RegisterClass() GUI failed.");
			return FAIL;
		}

		// This section may be needed later:
		//INITCOMMONCONTROLSEX icce;
		//icce.dwSize = sizeof(INITCOMMONCONTROLSEX);
		//icce.dwICC = ICC_TAB_CLASSES  // Tab and Tooltip
		//	| ICC_UPDOWN_CLASS   // up-down control
		//	| ICC_PROGRESS_CLASS // progress bar
		//	| ICC_DATE_CLASSES;  // date and time picker
		//InitCommonControlsEx(&icce);

		sGuiInitialized = true;
	}

	// WS_EX_APPWINDOW: "Forces a top-level window onto the taskbar when the window is minimized."
	// But it doesn't since the window is currently always unowned, there is not yet any need to use it.
	if (   !(mHwnd = CreateWindowEx(0, WINDOW_CLASS_GUI, g_script.mFileName, mStyle, 0, 0, 0, 0
		, mOwner, NULL, g_hInstance, NULL))   )
		return FAIL;

	if ((mStyle & WS_SYSMENU) || !mOwner)
	{
		// Setting the small icon puts it in the upper left corner of the dialog window.
		// Setting the big icon makes the dialog show up correctly in the Alt-Tab menu (but big seems to
		// have no effect unless the window is unowned, i.e. it has a button on the task bar).
		LPARAM main_icon = (LPARAM)(g_script.mCustomIcon ? g_script.mCustomIcon
			: LoadIcon(g_hInstance, MAKEINTRESOURCE(IDI_MAIN)));
		if (mStyle & WS_SYSMENU)
			SendMessage(mHwnd, WM_SETICON, ICON_SMALL, main_icon);
		if (!mOwner)
			SendMessage(mHwnd, WM_SETICON, ICON_BIG, main_icon);
	}


	char label_name[1024];  // Labels are unlimited in length, so use a size to cover anything realistic.

	// Find the label to run automatically when the form closes (if any):
	if (mWindowIndex > 0) // Prepend the window number for windows other than the first.
		snprintf(label_name, sizeof(label_name), "%dGuiClose", mWindowIndex + 1);
	else
		strcpy(label_name, "GuiClose");
	mLabelForClose = g_script.FindLabel(label_name);  // OK if NULL (closing the window is the same as "gui, cancel").

	// Find the label to run automatically when the the user presses Escape (if any):
	if (mWindowIndex > 0) // Prepend the window number for windows other than the first.
		snprintf(label_name, sizeof(label_name), "%dGuiEscape", mWindowIndex + 1);
	else
		strcpy(label_name, "GuiEscape");
	mLabelForEscape = g_script.FindLabel(label_name);  // OK if NULL (pressing ESCAPE does nothing).

	return OK;
}



void GuiType::UpdateMenuBars(HMENU aMenu)
// Caller has changed aMenu and wants the change visibly reflected in any windows that that
// use aMenu as a menu bar.  For example, if a menu item has been disabled, the grey-color
// won't show up immediately unless the window is refreshed.
{
	for (int i = 0; i < MAX_GUI_WINDOWS; ++i)
		if (g_gui[i] && g_gui[i]->mHwnd && GetMenu(g_gui[i]->mHwnd) == aMenu && IsWindowVisible(g_gui[i]->mHwnd))
		{
			// Neither of the below two calls by itself is enough for all types of changes.
			// Thought it's possible that every type of change only needs one or the other, both
			// are done for simplicity:
			// This first line is necessary at least for cases where the height of the menu bar
			// (the number of rows needed to display all its items) has changed as a result
			// of the caller's change.  In addition, I believe SetWindowPos() must be called
			// before RedrawWindow() to prevent artifacts in some cases:
			SetWindowPos(g_gui[i]->mHwnd, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
			// This line is necessary at least when a single new menu item has been added:
			RedrawWindow(g_gui[i]->mHwnd, NULL, NULL, RDW_INVALIDATE|RDW_FRAME|RDW_UPDATENOW);
			// RDW_UPDATENOW: Seems best so that the window is in an visible updated state when function
			// returns.  This is because if the menu bar happens to be two rows or its size is changed
			// in any other way, the window dimensions themselves might change, and the caller might
			// rely on such a change being visibly finished for PixelGetColor, etc.
			//Not enough: UpdateWindow(g_gui[i]->mHwnd);
		}
}



ResultType GuiType::AddControl(GuiControls aControlType, char *aOptions, char *aText)
// Caller must have ensured that mHwnd is non-NULL (i.e. that the parent window already exists).
{
	if (mControlCount >= MAX_CONTROLS_PER_GUI)
		return g_script.ScriptError("Too many controls." ERR_ABORT); // Short msg since so rare.

	// If this is the first control, set the default margin for the window based on the size
	// of the current font:
	if (!mControlCount)
	{
		mMarginX = (int)(1.25 * sFont[mCurrentFontIndex].point_size);  // Seems to be a good rule of thumb.
		mMarginY = (int)(0.75 * sFont[mCurrentFontIndex].point_size);  // Also seems good.
		mPrevX = mMarginX;  // This makes first control be positioned correctly if it lacks both X & Y coords.
	}

	////////////////////////////////////////////////////////////////////////////////////////
	// Set defaults for the various options, to be overridden individually by any specified.
	////////////////////////////////////////////////////////////////////////////////////////
	GuiControlType &control = mControl[mControlCount];
	ZeroMemory(&control, sizeof(GuiControlType));
	control.color = mCurrentColor; // Default to the most recently set color.
	int width = COORD_UNSPECIFIED;
	int height = COORD_UNSPECIFIED;
	int x = COORD_UNSPECIFIED;
	int y = COORD_UNSPECIFIED;
	float row_count = 0;
	int choice = 0;  // Which item of a DropDownList/ComboBox/ListBox to initially choose.
	int checked = BST_UNCHECKED;  // Default starting state of checkbox.
	char password_char = '\0';  // Indicates "no password" for an edit character.
	DWORD style = WS_CHILD|WS_VISIBLE; // Starting default for all control types.
	char var_name[MAX_VAR_NAME_LENGTH + 20] = "";  // Make it longer than MAX so that AddVar() will detect and display errors for us.
	char label_name[1024] = "";  // Subroutine labels are nearly unlimited in length, so use a size to cover anything realistic.
	bool start_new_section = (mSectionX == COORD_UNSPECIFIED); // Whether pos. should be stored for use by other controls.

	//////////////////////////////////////////////////
	// Manage any automatic behavior for radio groups.
	//////////////////////////////////////////////////
	if (aControlType == GUI_CONTROL_RADIO)
	{
		if (!mInRadioGroup)
			style |= WS_GROUP|WS_TABSTOP; // Otherwise it lacks a tabstop by default.
			// The mInRadioGroup flag will be changed accordingly after the control is successfully created.
		//else no tabstop by default
	}
	else // Not a radio.
		if (mInRadioGroup) // Close out the prior radio group by giving this control the WS_GROUP style.
			style |= WS_GROUP;

	/////////////////////////////////////////////////
	// Set control-specific defaults for any options.
	/////////////////////////////////////////////////
	switch(aControlType)
	{
	case GUI_CONTROL_BUTTON:
	case GUI_CONTROL_CHECKBOX:
	case GUI_CONTROL_DROPDOWNLIST:
	case GUI_CONTROL_COMBOBOX:
	case GUI_CONTROL_LISTBOX:
	case GUI_CONTROL_EDIT:
		style |= WS_TABSTOP;  // Set default.
		break;
	// Nothing for these currently:
	//case GUI_CONTROL_TEXT:
	//case GUI_CONTROL_PIC:
	//case GUI_CONTROL_GROUPBOX:
	//case GUI_CONTROL_RADIO:
	}

	/////////////////////////////
	// Parse the list of options.
	/////////////////////////////
	// Vars for temporary use within the loop:
	size_t length;
	char *end_of_name, *space_pos;
	char color_str[32];

	for (char *cp = aOptions; *cp; ++cp)
	{
		switch(toupper(*cp))
		{

		case 'C':
			if (!strnicmp(cp, "CheckedGray", 11)) // *** MUST CHECK this prior to "checked" to avoid ambiguity/overlap.
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 10 vs. 11 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 10;
				if (aControlType == GUI_CONTROL_CHECKBOX) // Radios can't have the 3rd/gray state.
					checked = BST_INDETERMINATE;
			}
			else if (!strnicmp(cp, "Checked", 7))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 6 vs. 7 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 6;
				if (aControlType == GUI_CONTROL_CHECKBOX || aControlType == GUI_CONTROL_RADIO)
					checked = BST_CHECKED;
			}
			else if (!strnicmp(cp, "Check3", 6)) // Enable tri-state checkbox.
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 5 vs. 6 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 5;
				if (aControlType == GUI_CONTROL_CHECKBOX) // Radios can't have the 3rd/gray state.
					style |= BS_AUTO3STATE;
			}
			else if (!strnicmp(cp, "center", 6))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 5 vs. 6 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 5;
				switch (aControlType)
				{
					case GUI_CONTROL_TEXT:
					case GUI_CONTROL_PIC:  // Has no actual effect currently.
						style |= SS_CENTER;
						break;
					case GUI_CONTROL_GROUPBOX: // Changes alighment of its label.
					case GUI_CONTROL_BUTTON:   // Probably has no effect in this case, since it's centered by default?
					case GUI_CONTROL_CHECKBOX: // Puts gap between box and label.
					case GUI_CONTROL_RADIO:
						style |= BS_CENTER;
						break;
					case GUI_CONTROL_EDIT:
						style |= ES_CENTER;
						break;
					// Not applicable for:
					//case GUI_CONTROL_DROPDOWNLIST:
					//case GUI_CONTROL_COMBOBOX:
					//case GUI_CONTROL_LISTBOX:
				}
			}
			else if (!strnicmp(cp, "choose", 6))
			{
				// "CHOOSE" provides an easier way to conditionally select a different item at the time
				// the control is added.  Example: gui, add, ListBox, vMyList Choose%choice%, %MyItemList%
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 5 vs. 6 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 5;
				choice = atoi(cp + 1); // This variable is later ignored if not applicable for this control type.
				if (choice < 0) // Invalid: number should be 1 or greater.
					choice = 0;
			}
			else // Assume it's a color.
			{
				strlcpy(color_str, cp + 1, sizeof(color_str));
				if (space_pos = StrChrAny(color_str, " \t"))  // space or tab
					*space_pos = '\0';
				//else a color name can still be present if it's at the end of the string.
				control.color = ColorNameToBGR(color_str);
				if (control.color == CLR_NONE) // A matching color name was not found, so assume it's in hex format.
				{
					if (strlen(color_str) > 6)
						color_str[6] = '\0';  // Shorten to exactly 6 chars, which happens if no space/tab delimiter is present.
					control.color = rgb_to_bgr(strtol(color_str, NULL, 16));
					// if color_str does not contain something hex-numeric, black (0x00) will be assumed,
					// which seems okay given how rare such a problem would be.
				}
				// Skip over the color string to avoid interpreting hex digits or color names as option letters:
				cp += strlen(color_str);
			}
			break;

		case 'D':
			if (!strnicmp(cp, "default", 7))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 6 vs. 7 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 6;
				if (aControlType == GUI_CONTROL_BUTTON)
					style |= BS_DEFPUSHBUTTON;
				//else ignore this option for other types
			}
			else if (!strnicmp(cp, "disabled", 8))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 7 vs. 8 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 7;
				style |= WS_DISABLED;
			}
			break;

		case 'L':
			if (!strnicmp(cp, "left", 4))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 3 vs. 4 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 3;
				switch (aControlType)
				{
					case GUI_CONTROL_GROUPBOX: // Probably not applicable, but just in case.
					case GUI_CONTROL_BUTTON:
					case GUI_CONTROL_CHECKBOX: // Not applicable unless used like this: "right left" (due to BS_RIGHTBUTTON).
					case GUI_CONTROL_RADIO:    // Same.
						style |= BS_LEFT;
						break;
					// Not applicable for:
					//case GUI_CONTROL_TEXT:   SS_LEFT seems to be acquired automatically according to Spy++.
					//case GUI_CONTROL_PIC:
					//case GUI_CONTROL_EDIT:
					//case GUI_CONTROL_DROPDOWNLIST:
					//case GUI_CONTROL_COMBOBOX:
					//case GUI_CONTROL_LISTBOX:
				}
			}
			break;

		case 'N':
			if (!strnicmp(cp, "NoTab", 5))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 4 vs. 5 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 4;
				style &= ~WS_TABSTOP;
			}
			break;

		case 'P':
			if (!strnicmp(cp, "password", 8)) // Password vs. pass to make scripts more self-documenting.
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 7 vs. 8 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 7;
				if (*(cp + 1))
				{
					++cp; // Skip over the password char to prevent it from being seen as an option letter.
					// Allow a space to be the masking character, since it's conceivable that might
					// be wanted in cases where someone doesn't wany anyone to know they're typing a password.
					password_char = *cp;  // Later ignored if this control isn't an edit. Can be '\0'.
				}
				else
					password_char = '*'; // Use default.
				if (aControlType == GUI_CONTROL_EDIT)
					style |= ES_PASSWORD;
			}
			break;

		case 'S':
			if (!strnicmp(cp, "section", 7))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 6 vs. 7 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 6;
				start_new_section = true;
			}
			break;
		// For option letters that accept a string, such as G and V:
		// Don't allow syntax "v varname" (spaces between option letter and its string) because
		// that might lead to ambiguity, either in this option letter or in others for which
		// we don't want this one to set a confusing precedent.  The ambiguity would come about
		// if the option letter is intended to specify a blank string, e.g. "v l", where v
		// specifies a blank string and l is the next option letter rather than being v's string.
		case 'G': // "Gosub" a label (as a new thread) when something actionable happens to the control.
			if (!*(cp + 1)) // Avoids reading beyond end of string due to loop's additional increment.
				break;
			++cp;
			if (   !(end_of_name = StrChrAny(cp, " \t"))   )
				end_of_name = cp + strlen(cp);
			length = end_of_name - cp;
			if (length)
			{
				// For reasons of potential future use and compatibility, don't allow subroutines to be
				// assigned to control types that have no present use for them.  Note: GroupBoxes do
				// no support click-detection anyway, even if the BS_NOTIFY style is given to them
				// (this has been verified twice):
				if (aControlType == GUI_CONTROL_EDIT || aControlType == GUI_CONTROL_GROUPBOX)
					return g_script.ScriptError("This control type should not have an associated subroutine." ERR_ABORT, cp);
				if (length >= sizeof(label_name)) // Prevent buffer overflow.  Truncation is reported when label isn't found.
					length = sizeof(label_name) - 1;
				strlcpy(label_name, cp, length + 1);
			}
			// Skip over the text of the name so that it isn't interpreted as option letters.  -1 to avoid
			// the loop's addition ++cp from reading beyond the length of the string:
			cp = end_of_name - 1;
			break;

		case 'V': // Variable in which to store control's contents or selection.
			if (!*(cp + 1)) // Avoids reading beyond end of string due to loop's additional increment.
				break;
			++cp;
			if (   !(end_of_name = StrChrAny(cp, " \t"))   )
				end_of_name = cp + strlen(cp);
			length = end_of_name - cp;
			if (length)
			{
				// For reasons of potential future use and compatibility, don't allow variables to be
				// assigned to control types that have no present use for them:
				switch (aControlType)
				{
				case GUI_CONTROL_TEXT:
				case GUI_CONTROL_PIC:
				case GUI_CONTROL_GROUPBOX:
				case GUI_CONTROL_BUTTON:
					return g_script.ScriptError("This control type should not have an associated variable." ERR_ABORT, cp);
				}
				if (length >= sizeof(var_name)) // Prevent buffer overflow.  AddVar() will report the "too long" for us.
					length = sizeof(var_name) - 1;
				strlcpy(var_name, cp, length + 1);
			}
			// Skip over the text of the name so that it isn't interpreted as option letters.  -1 to avoid
			// the loop's addition ++cp from reading beyond the length of the string:
			cp = end_of_name - 1;
			break;

		// For options such as W, H, X and Y:
		// Use atoi() vs. ATOI() to avoid interpreting something like 0x01B as hex when in fact
		// the B was meant to be an option letter:
		// DIMENSIONS:
		case 'W':
			width = atoi(cp + 1);
			break;

		case 'H':
			if (!strnicmp(cp, "hidden", 6))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 5 vs. 6 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 5;
				style &= ~WS_VISIBLE;
			}
			else
				height = atoi(cp + 1);
			break;

		case 'R':
			if (!strnicmp(cp, "right", 5))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 4 vs. 5 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 4;
				switch (aControlType)
				{
					case GUI_CONTROL_TEXT:
					case GUI_CONTROL_PIC:  // Has no actual effect currently.
						style |= SS_RIGHT;
						break;
					case GUI_CONTROL_GROUPBOX:
					case GUI_CONTROL_BUTTON:
					case GUI_CONTROL_CHECKBOX:
					case GUI_CONTROL_RADIO:
						style |= BS_RIGHT;
						// And by default, put button itself to the right of its label since that seems
						// likely to be far more common/desirable (there can be a more obscure option
						// later to change this default):
						if (aControlType == GUI_CONTROL_CHECKBOX || aControlType == GUI_CONTROL_RADIO)
							style |= BS_RIGHTBUTTON;
						break;
					case GUI_CONTROL_EDIT:
						style |= ES_RIGHT;
						break;
					// Not applicable for:
					//case GUI_CONTROL_DROPDOWNLIST:
					//case GUI_CONTROL_COMBOBOX:
					//case GUI_CONTROL_LISTBOX:
				}
			}
			else if (!strnicmp(cp, "ReadOnly", 8))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 7 vs. 8 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 7;
				if (aControlType == GUI_CONTROL_EDIT)
					style |= ES_READONLY;
				//else ignore this option for other types
			}
			// Otherwise:
			// The number of rows desired in the control.  Use atof() so that fractional rows are allowed.
			row_count = (float)atof(cp + 1); // Don't need double precision.
			break;

		case 'X':
			if (!*(cp + 1)) // Avoids reading beyond end of string due to loop's additional increment.
				break;
			++cp;
			if (*cp == '+')
			{
				x = mPrevX + mPrevWidth + atoi(cp + 1);
				if (y == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
					y = mPrevY;  // Since moving in the X direction, retain the same Y as previous control.
			}
			// For the M and P sub-options, not that the +/- prefix is optional.  The number is simply
			// read in as-is (though the use of + is more self-documenting in this case than omitting
			// the sign entirely).
			else if (toupper(*cp) == 'M') // Use the X margin
			{
				x = mMarginX + atoi(cp + 1);
				if (y == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
					y = mMaxExtentDown + mMarginY;
			}
			else if (toupper(*cp) == 'P') // Use the previous control's X position.
			{
				x = mPrevX + atoi(cp + 1);
				if (y == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
					y = mPrevY;  // Since moving in the X direction, retain the same Y as previous control.
			}
			else if (toupper(*cp) == 'S') // Use the saved X position
			{
				x = mSectionX + atoi(cp + 1);
				if (y == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
					y = mMaxExtentDownSection + mMarginY;  // In this case, mMarginY is the padding between controls.
			}
			else
			{
				x = atoi(cp);
				if (y == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
					y = mMaxExtentDown + mMarginY;
			}
			break;

		case 'Y':
			if (!*(cp + 1)) // Avoids reading beyond end of string due to loop's additional increment.
				break;
			++cp;
			if (*cp == '+')
			{
				y = mPrevY + mPrevHeight + atoi(cp + 1);
				if (x == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
					x = mPrevX;  // Since moving in the Y direction, retain the same X as previous control.
			}
			// For the M and P sub-options, not that the +/- prefix is optional.  The number is simply
			// read in as-is (though the use of + is more self-documenting in this case than omitting
			// the sign entirely).
			else if (toupper(*cp) == 'M') // Use the Y margin
			{
				y = mMarginY + atoi(cp + 1);
				if (x == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
					x = mMaxExtentRight + mMarginX;
			}
			else if (toupper(*cp) == 'P') // Use the previous control's Y position.
			{
				y = mPrevY + atoi(cp + 1);
				if (x == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
					x = mPrevX;  // Since moving in the Y direction, retain the same X as previous control.
			}
			else if (toupper(*cp) == 'S') // Use the saved Y position
			{
				y = mSectionY + atoi(cp + 1);
				if (x == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
					x = mMaxExtentRightSection + mMarginX; // In this case, mMarginX is the padding between controls.
			}
			else
			{
				y = atoi(cp);
				if (x == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
					x = mMaxExtentRight + mMarginX;
			}
			break;
		// Otherwise: Ignore other characters, such as the digits that occur after the P/X/Y option letters.
		} // switch()
	} // for()

	////////////////////////////////////////////////////////////////////
	// Set the control's associated variable and/or label, if available.
	////////////////////////////////////////////////////////////////////
	UINT u;  // Used in various places.
	if (*var_name)
	{
		if (   !(control.output_var = g_script.FindOrAddVar(var_name))   )
			return FAIL;  // It already displayed the error.
		// Check if any other control (visible or not, to avoid the complexity of a hidden control
		// needing to be dupe-checked every time it becomes visible) on THIS gui window has the
		// same variable.  That's an error because not only doesn't it make sense to do that,
		// but it might be useful to uniquely identify a control by its variable name (when making
		// changes to it, etc.)  Note that if this is the first control being added, mControlCount
		// is now zero because this control has not yet actually been added.  That is why
		// "u < mControlCount" is used:
		for (u = 0; u < mControlCount; ++u)
			if (mControl[u].output_var == control.output_var)
				return g_script.ScriptError("The same variable cannot be used for more than one control per window."
					ERR_ABORT, var_name);
	}
	//else:
	// It seems best to allow an input control to lack a variable, in which case its contents will be
	// lost when the form is closed (unless fetched beforehand with something like ControlGetText).
	// This also allows layout editors and other script generators to omit the variable and yet still
	// be able to generate a runnable script.

	if (*label_name)
	{
		if (   !(control.jump_to_label = g_script.FindLabel(label_name))   )
		{
			// If there is no explicit label, fall back to a special action if one is available
			// for this keyword:
			if (!stricmp(label_name, "Cancel"))
				control.implicit_action = GUI_IMPLICIT_CANCEL;
			//else if (!stricmp(label_name, "Clear"))
			//	control.implicit_action = GUI_IMPLICIT_CLEAR;
			else // Since a non-special label was explicitly specified, display an error that it couldn't be found.
				return g_script.ScriptError(ERR_CONTROLLABEL ERR_ABORT, label_name);
		}
		// Apply the SS_NOTIFY style *only* if the control actually has an associated action.
		// This is because otherwise the control would steal all clicks for any other controls
		// drawn on top of it (e.g. a picture control with some edit fields drawn on top of it).
		// See comments in the creation of GUI_CONTROL_PIC for more info:
		if (aControlType == GUI_CONTROL_TEXT || aControlType == GUI_CONTROL_PIC)
			style |= SS_NOTIFY;
	}
	else if (aControlType == GUI_CONTROL_BUTTON) // Check whether the automatic/implicit label exists.
	{
		if (mWindowIndex > 0) // Prepend the window number for windows other than the first.
			_itoa(mWindowIndex + 1, label_name, 10);
		else
			*label_name = '\0';
		snprintfcat(label_name, sizeof(label_name), "Button%s", aText);
		// Remove spaces and ampersands.  Although ampersands are legal in labels, it seems
		// more friendly not to use them in the automatic-label label name.  Note that a button
		// or menu item can contain a literal ampersand by using two ampersands, such as
		// "Save && Exit".  In this example, the auto-label would be named "ButtonSaveExit".
		StrReplaceAll(label_name, " ", "");
		StrReplaceAll(label_name, "&", "");
		StrReplaceAll(label_name, "\r", "");
		StrReplaceAll(label_name, "\n", "");
		control.jump_to_label = g_script.FindLabel(label_name);  // OK if NULL (the button will do nothing).
	}

	////////////////////////////////////////////////////////////////////////////////////////////
	// Automatically set the control's position in the client area if no position was specified.
	////////////////////////////////////////////////////////////////////////////////////////////
	if (x == COORD_UNSPECIFIED && y == COORD_UNSPECIFIED)
	{
		// Since both coords were unspecified, proceed downward from the previous control, using a default margin.
		x = mPrevX;
		y = mPrevY + mPrevHeight + mMarginY;  // Don't use mMaxExtentDown in this is a new column.
		if (aControlType == GUI_CONTROL_TEXT && mControlCount && mControl[mControlCount - 1].type == GUI_CONTROL_TEXT)
			// Since this text control is being auto-positioned immediately below another, provide extra
			// margin space so that any edit control later added to its right in "vertical progression"
			// mode will line up with it.
			y += GUI_CTL_VERTICAL_DEADSPACE;
	}
	// Can't happen due to the logic in the options-parsing section:
	//else if (x == COORD_UNSPECIFIED)
	//	x = mPrevX;
	//else if (y == COORD_UNSPECIFIED)
	//	y = mPrevY;


	//////////////////////////////////////////////////////////////////////////////////
	// For certain types of controls, provide a standard height if none was specified.
	//////////////////////////////////////////////////////////////////////////////////
	if (aControlType == GUI_CONTROL_EDIT && (style & WS_VSCROLL))
		style |= GUI_EDIT_DEFAULT_STYLE_MULTI;  // This will need to be revised to cooperate with future options.

	if (height == COORD_UNSPECIFIED && row_count <= 0)
	{
		switch(aControlType)
		{
		case GUI_CONTROL_DROPDOWNLIST:
		case GUI_CONTROL_COMBOBOX: // In these cases, row-count is defined as the number of items to display in the list.
		case GUI_CONTROL_LISTBOX:
			row_count = 3;  // Actual height will be calculated below using this.
			break;
		case GUI_CONTROL_GROUPBOX:
			// Seems more appropriate to give GUI_CONTROL_GROUPBOX exactly two rows: the first for the
			// title of the group-box and the second for its content (since it might contain controls
			// placed horizontally end-to-end, and thus only need one row).
			row_count = 2;
			break;
		case GUI_CONTROL_EDIT:
			// If there's no default text in the control from which to later calc the height, use 1 row.
			if (!*aText)
				row_count = 1;
			break;
		// Types not included
		// ------------------
		//case GUI_CONTROL_TEXT:     Rows are based on control's contents.
		//case GUI_CONTROL_PIC:      N/A
		//case GUI_CONTROL_BUTTON:   Rows are based on control's contents.
		//case GUI_CONTROL_CHECKBOX: Same
		//case GUI_CONTROL_RADIO:    Same
		}
	}

	////////////////////////////////////////////////////////////////////////////////////////////
	// In case the control being added requires an HDC to calculate its size, provide the means.
	////////////////////////////////////////////////////////////////////////////////////////////
	HDC hdc = NULL;
	HFONT hfont_old = NULL;
	TEXTMETRIC tm; // Used in more than one place.
	// To improve maintainability, always use this macro to deal with the above.
	// HDC will be released much further below when it is no longer needed.
	// Remember to release DC if ever need to return FAIL in the middle of auto-sizing/positioning.
	#define GUI_SET_HDC \
		if (!hdc)\
		{\
			hdc = GetDC(mHwnd);\
			hfont_old = (HFONT)SelectObject(hdc, sFont[mCurrentFontIndex].hfont);\
		}

	//////////////////////////////////////////////////////////////////////////////////////
	// If a row-count was specified or made available by the above defaults, calculate the
	// control's actual height (to be used when creating the window).  Note: If both
	// row_count and height were explicitly specified, row_count takes precedence.
	//////////////////////////////////////////////////////////////////////////////////////
	if (row_count > 0)
	{
		// For GroupBoxes, add 1 to any row_count greater than 1 so that the title itself is
		// already included automatically.  In other words, the R-value specified by the user
		// should be the number of rows available INSIDE the box.
		// For DropDownLists and ComboBoxes, 1 is added because row_count is defined as the
		// number of rows shown in the drop-down portion of the control, so we need one extra
		// (used in later calculations) for the always visible portion of the control.
		switch (aControlType)
		{
		case GUI_CONTROL_DROPDOWNLIST:
		case GUI_CONTROL_COMBOBOX:
		case GUI_CONTROL_GROUPBOX:
			++row_count;
		}
		GUI_SET_HDC
		GetTextMetrics(hdc, &tm);
		// Calc the height by adding up the font height for each row, and including the space between lines
		// (tmExternalLeading) if there is more than one line.  0.5 is used in two places to prevent
		// negatives in one, and round the overall result in the other.
		height = (int)((tm.tmHeight * row_count) + (tm.tmExternalLeading * ((int)(row_count + 0.5) - 1)) + 0.5);
		switch (aControlType)
		{
		case GUI_CONTROL_EDIT:
		case GUI_CONTROL_DROPDOWNLIST:
		case GUI_CONTROL_COMBOBOX:
		case GUI_CONTROL_LISTBOX:
			height += GUI_CTL_VERTICAL_DEADSPACE;
			break;
		case GUI_CONTROL_BUTTON:
			// Provide a extra space for top/bottom margin together, proportional to the current font
			// size so that it looks better with very large or small fonts.  The +2 seems to make
			// it look just right on all font sizes, especially the default GUI size of 8 where the
			// height should be about 23 to be standard(?)
			height += sFont[mCurrentFontIndex].point_size + 2;
			break;
		case GUI_CONTROL_GROUPBOX: // Since groups usually contain other controls, the below sizing seems best.
			// Use row_count-2 because of the +1 added above for GUI_CONTROL_GROUPBOX.
			// The current font's height is added in to provide an upper/lower margin in the box
			// proportional to the current font size, which makes it look better in most cases:
			height += (GUI_CTL_VERTICAL_DEADSPACE * ((int)(row_count + 0.5) - 2)) + (2 * sFont[mCurrentFontIndex].point_size);
			break;
		// Types not included
		// ------------------
		//case GUI_CONTROL_TEXT:     Uses basic height calculated above the switch().
		//case GUI_CONTROL_PIC:      Uses basic height calculated above the switch() (seems OK even for pic).
		//case GUI_CONTROL_CHECKBOX: Uses basic height calculated above the switch().
		//case GUI_CONTROL_RADIO:    Same.
		}
	}

	if (height == COORD_UNSPECIFIED || width == COORD_UNSPECIFIED)
	{
		int extra_width = 0;
		UINT draw_format = DT_CALCRECT;

		switch (aControlType)
		{
		case GUI_CONTROL_EDIT:
			if (!*aText) // Only auto-calculate edit's dimensions if there is text to do it with.
				break;
			// Since edit controls leave approximate 1 avg-char-width margin on the right side,
			// and probably exactly 4 pixels on the left counting its border and the internal
			// margin), adjust accordingly so that DrawText() will calculate the correct
			// control height based on word-wrapping.  Note: Can't use EM_GETRECT because
			// control doesn't exist yet (though that might be an alternative approach for
			// the future):
			GUI_SET_HDC
			GetTextMetrics(hdc, &tm);
			extra_width += 4 + tm.tmAveCharWidth;
			// Determine (if possible) whether there will be a vertical scrollbar present:
			if (row_count >= 1.5 || (style & WS_VSCROLL) || strchr(aText, '\n'))
				extra_width += GetSystemMetrics(SM_CXVSCROLL);
			// DT_EDITCONTROL: "the average character width is calculated in the same manner as for an edit control"
			draw_format |= DT_EDITCONTROL; // Might help some aspects of the estimate conducted below.
			// and now fall through and have the dimensions calculated based on what's in the control.
		case GUI_CONTROL_TEXT:
		case GUI_CONTROL_BUTTON:
		case GUI_CONTROL_CHECKBOX:
		case GUI_CONTROL_RADIO:
		{
			GUI_SET_HDC
			if (aControlType == GUI_CONTROL_CHECKBOX || aControlType == GUI_CONTROL_RADIO)
			{
				// Both Checkbox and Radio seem to have the same spacing characteristics:
				// Expand to allow room for button itself, its border, and the space between
				// the button and the first character of its label (this space seems to
				// be the same as tmAveCharWidth).  +2 seems to be needed to make it work
				// for the various sizes of Courier New vs. Verdana that I tested.  The
				// alternative, (2 * GetSystemMetrics(SM_CXEDGE)), seems to add a little
				// too much width (namely 4 vs. 2).
				GetTextMetrics(hdc, &tm);
				extra_width += GetSystemMetrics(SM_CXMENUCHECK) + tm.tmAveCharWidth + 2;
			}
			if (width != COORD_UNSPECIFIED) // Since a width was given, auto-expand the height via word-wrapping.
				draw_format |= DT_WORDBREAK;
			RECT draw_rect;
			draw_rect.left = 0;
			draw_rect.top = 0;
			draw_rect.right = (width == COORD_UNSPECIFIED) ? 0 : width - extra_width; // extra_width
			draw_rect.bottom = (height == COORD_UNSPECIFIED) ? 0 : height;
			// If no text, "H" is used in case the function requires a non-empty string to give consistent results:
			int draw_height = DrawText(hdc, *aText ? aText : "H", -1, &draw_rect, draw_format);
			int draw_width = draw_rect.right - draw_rect.left;
			// Even if either height or width was already explicitly specified above, it seems best to
			// override it if DrawText() says it's not big enough.  REASONING: It seems too rare that
			// someone would want to use an explicit height/width to selectively hide part of a control's
			// contents, presumably for revelation later.  If that is truly desired, ControlMove or
			// similar can be used to resize the control afterward.  In addition, by specifying BOTH
			// width and height/rows, none of these calculates happens anyway, so that's another way
			// this override can be overridden:
			if (height == COORD_UNSPECIFIED || draw_height > height)
			{
				height = draw_height;
				if (aControlType == GUI_CONTROL_EDIT)
					height += GUI_CTL_VERTICAL_DEADSPACE;
				else if (aControlType == GUI_CONTROL_BUTTON)
					height += sFont[mCurrentFontIndex].point_size + 2;  // +2 makes it standard height.
			}
			if (width == COORD_UNSPECIFIED || draw_width > width)
			{
				width = draw_width + extra_width;
				if (aControlType == GUI_CONTROL_BUTTON)
					// Allow room for border and an internal margin proportional to the font height.
					// Button's border is 3D by default, so SM_CXEDGE vs. SM_CXBORDER is used?
					width += 2 * GetSystemMetrics(SM_CXEDGE) + sFont[mCurrentFontIndex].point_size;
			}
			break;
		} // case

		// Types not included
		// ------------------
		//case GUI_CONTROL_PIC:           If needed, it is given some default dimensions at the time of creation.
		//case GUI_CONTROL_GROUPBOX:      Seems too rare than anyone would want its width determined by its text.
		//case GUI_CONTROL_EDIT:          It is included, but only if it has default text inside it.
		//case GUI_CONTROL_DROPDOWNLIST:  These last 3 are given (below) a standard width based on font size.
		//case GUI_CONTROL_COMBOBOX:
		//case GUI_CONTROL_LISTBOX:

		} // switch()
	}

	//////////////////////////////////////////////////////////////////////////////////////////
	// If the width was not specified and the above did not already determine it (which should
	// only be possible for the cases contained in the switch-stmt below), provide a default.
	//////////////////////////////////////////////////////////////////////////////////////////
	if (width == COORD_UNSPECIFIED)
	{
		switch(aControlType)
		{
		case GUI_CONTROL_DROPDOWNLIST:
		case GUI_CONTROL_COMBOBOX:
		case GUI_CONTROL_LISTBOX:
		case GUI_CONTROL_EDIT:
			width = GUI_STANDARD_WIDTH;
			break;
		case GUI_CONTROL_GROUPBOX:
			// Since groups contain other controls, allow room inside them for a margin based on current
			// font size.  Multiplying by 3 seems to yield about the right margin amount based on
			// other dialogs I've seen.
			width = GUI_STANDARD_WIDTH + (3 * sFont[mCurrentFontIndex].point_size);
			break;
		// Types not included
		// ------------------
		//case GUI_CONTROL_TEXT:     Exact width should already have been calculated based on contents.
		//case GUI_CONTROL_PIC:      Calculated based on actual pic size if no explicit width was given.
		//case GUI_CONTROL_BUTTON:   Exact width should already have been calculated based on contents.
		//case GUI_CONTROL_CHECKBOX: Same.
		//case GUI_CONTROL_RADIO:    Same.
		}
	}

	/////////////////////////////////////////////////////////////////////////////////////////
	// For edit controls: If the above didn't already determine how many rows it should have,
	// auto-detect that by comparing the current font size with the specified height.
	/////////////////////////////////////////////////////////////////////////////////////////
	if (aControlType == GUI_CONTROL_EDIT && !(style & ES_MULTILINE))
	{
		if (row_count <= 0) // Determine the row-count to auto-detect multi-line vs. single-line.
		{
			GUI_SET_HDC
			GetTextMetrics(hdc, &tm);
			int height_beyond_first_row = height - GUI_CTL_VERTICAL_DEADSPACE - tm.tmHeight;
			if (height_beyond_first_row > 0)
				row_count = 1 + ((float)height_beyond_first_row / (tm.tmHeight + tm.tmExternalLeading));
			else
				row_count = 1;
		}
		// Set the type based on the value of row_count determined above or earlier:
		if (row_count < 1.5)
			style |= GUI_EDIT_DEFAULT_STYLE_SINGLE;
		else
			style |= GUI_EDIT_DEFAULT_STYLE_MULTI;
	}

	// If either height or width is still undetermined, leave it set to COORD_UNSPECIFIED since that
	// is a large negative number and should thus help catch bugs.  In other words, the above
	// hueristics should be designed to handle all cases and always resolve height/width to something,
	// with the possible exception of things that auto-size based on external content such as
	// GUI_CONTROL_PIC.

	////////////////////////////////
	// Release the HDC if necessary.
	////////////////////////////////
	if (hdc)
	{
		if (hfont_old)
		{
			SelectObject(hdc, hfont_old); // Necessary to avoid memory leak.
			hfont_old = NULL;
		}
		ReleaseDC(mHwnd, hdc);
		hdc = NULL;
	}

	//////////////////////
	//
	// CREATE THE CONTROL.
	//
	//////////////////////
	bool font_was_set = false;
	bool retrieve_dimensions = false;
	int item_height, min_list_height;
	char *malloc_buf;
	HMENU control_id = (HMENU)(size_t)GUI_INDEX_TO_ID(mControlCount); // Cast to size_t avoids compiler warning.

	switch(aControlType)
	{
	case GUI_CONTROL_TEXT:
		// Seems best to omit SS_NOPREFIX by default so that ampersand can be used to create shortcut keys.
		control.hwnd = CreateWindowEx(0, "static", aText, style
			, x, y, width, height, mHwnd, control_id, g_hInstance, NULL);
		break;

	case GUI_CONTROL_PIC:
		if (width == COORD_UNSPECIFIED)
			width = 0;  // Use zero to tell LoadPicture() to keep original width.
		if (height == COORD_UNSPECIFIED)
			height = 0;  // Use zero to tell LoadPicture() to keep original height.
		if (control.hwnd = CreateWindowEx(0, "static", aText, style|SS_BITMAP
			, x, y, width, height  // OK if zero, control creation should still succeed.
			, mHwnd, control_id, g_hInstance, NULL))
		{
			// In light of the below, it seems best to delete the bitmaps whenever the control changes
			// to a new image or whenever the control is destroyed.  Otherwise, if a control or its
			// parent window is destroyed and recreated many times, memory allocation would continue
			// to grow from all the abandoned pointers.
			// MSDN: "When you are finished using a bitmap...loaded without specifying the LR_SHARED flag,
			// you can release its associated memory by calling...DeleteObject."
			// MSDN: "The system automatically deletes these resources when the process that created them
			// terminates, however, calling the appropriate function saves memory and decreases the size
			// of the process's working set."
			if (   !(control.hbitmap = LoadPicture(aText, width, height))   )
				break;  // By design, no error is reported.  The picture is simply not displayed.
			// LoadPicture() uses CopyImage() to scale the image, which seems to provide better scaling
			// quality than using MoveWindow() (followed by redrawing the parent window) on the static
			// control that contains the image.
			SendMessage(control.hwnd, (UINT)STM_SETIMAGE, (WPARAM)IMAGE_BITMAP, (LPARAM)control.hbitmap);
			// UPDATE ABOUT THE BELOW: Rajat says he can't get the Smart GUI working without
			// the controls retaining their original numbering/z-order.  This has to do with the fact
			// that TEXT controls and PIC controls are both static.  If only PIC controls were reordered,
			// that's not enought to make things work for him.  If only TEXT controls were ALSO reordered
			// to be at the top of the list (so that all statics retain their original ordering with
			// respect to each other) I have a 90% expectation (based on testing) that prefix/shortcut
			// keys inside static text controls would jump to the wrong control because their associated
			// control seems to be based solely on z-order.  ANOTHER REASON NOT to ever change the z-order
			// of controls automatically: Every time a picture is added, it would increment the z-order
			// number of all other control types by 1.  This is bad especially for text controls because
			// they are static just like pictures, and thus their Class+NN "unique ID" would change (as
			// seen by the ControlXXX commands) if a picture were ever added after a window was shown.
			// Older note: calling SetWindowPos() with HWND_TOPMOST doesn't seem to provide any useful
			// effect that I could discern (by contrast, HWND_TOP does actually move a control to the
			// top of the z-order).
			// The below is OBSOLETE and its code further below is commented out:
			// Facts about how overlapping controls are drawn vs. which one receives mouse clicks:
			// 1) The first control created is at the top of the Z-order, the second is next, and so on.
			// 2) Controls get drawn in order based on their Z-order (i.e. the first control is drawn first
			//    and any later controls that overlap are drawn on top of it).
			// 3) When a user clicks a point that contains two overlapping controls and each control is
			//    capable of capturing clicks, the one closer to the top captures the click even though it
			//    was drawn beneath (overlapped by) the other control.
			// Because of this behavior, the following policy seems best:
			// 1) Move all static images to the top of the Z-order so that other controls are always
			//    drawn on top of them.  This is done because it seems to be the behavior that would
			//    be desired at least 90% of the time.
			// 2) Do not do the same for static text and GroupBoxes because it seems too rare that
			//    overlapping would be done in such cases, and even if it is done it seems more
			//    flexible to allow the order in which the controls were created to determine how they
			//    overlap and which one get the clicks.
			//
			// Rather push static pictures to the top in the reverse order they were created -- 
			// which might be a little more intuitive since the ones created first would then always
			// be "behind" ones created later -- for simplicity, we do it here at the time the control
			// is created.  This avoids complications such as a picture being added after the
			// window is shown for the first time and then not getting sent to the top where it
			// should be.  Update: The control is now kept in its original order for two reasons:
			// 1) The reason mentioned above (that later images should overlap earlier images).
			// 2) Allows Rajat's SmartGUI Creator to support picture controls (due to its grid background pic).
			// First find the last picture control in the array.  The last one should already be beneath
			// all the others in the z-order due to having been originally added using the below method.
			// Might eventually need to switch to EnumWindows() if it's possible for Z-order to be
			// altered, or controls to be insertted between others, by any commands in the future.
			//UINT index_of_last_picture = UINT_MAX;
			//for (u = 0; u < mControlCount; ++u)
			//{
			//	if (mControl[u].type == GUI_CONTROL_PIC)
			//		index_of_last_picture = u;
			//}
			//if (index_of_last_picture == UINT_MAX) // There are no other pictures, so put this one on top.
			//	SetWindowPos(control.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_NOMOVE|SWP_NOSIZE);
			//else // Put this picture after the last picture in the z-order.
			//	SetWindowPos(control.hwnd, mControl[index_of_last_picture].hwnd, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_NOMOVE|SWP_NOSIZE);
			//// Adjust to control's actual size in case it changed for any reason (failure to load picture, etc.)
			retrieve_dimensions = true;
		}
		break;

	case GUI_CONTROL_GROUPBOX:
		// In this case, BS_MULTILINE will obey literal newlines in the text, but it does not automatically
		// wrap the text, at least on XP.  Since it's strange-looking to have multiple lines, newlines
		// should be rarely present anyway.  Also, BS_NOTIFY seems to have no effect on GroupBoxes (it
		// never sends any BN_CLICKED/BN_DBLCLK messages).  This has been verified twice.
		control.hwnd = CreateWindowEx(0, "button", aText, style|BS_MULTILINE|BS_GROUPBOX
			, x, y, width, height, mHwnd, control_id, g_hInstance, NULL);
		break;

	case GUI_CONTROL_BUTTON:
		// For all "button" type controls, BS_MULTILINE is included by default so that any literal
		// newlines in the button's name will start a new line of text as the user intended.
		// In addition, this causes automatic wrapping to occur if the user specified a width
		// too small to fit the entire line.
		if (control.hwnd = CreateWindowEx(WS_EX_WINDOWEDGE, "button", aText, style|BS_MULTILINE
			, x, y, width, height, mHwnd, control_id, g_hInstance, NULL))
		{
			if (style & BS_DEFPUSHBUTTON)
			{
				// First remove the style from the old default button, if there is one:
				if (mDefaultButtonIndex < mControlCount)
					SetWindowLong(mControl[mDefaultButtonIndex].hwnd, GWL_STYLE
						, GetWindowLong(mControl[mDefaultButtonIndex].hwnd, GWL_STYLE) & ~BS_DEFPUSHBUTTON);
				mDefaultButtonIndex = mControlCount;
			}
		}
		break;

	case GUI_CONTROL_CHECKBOX:
		// Note: Having both of these styles causes the control to be displayed incorrectly, so ensure
		// they are mutually exclusive:
		if (!(style & BS_AUTO3STATE))
			style |= BS_AUTOCHECKBOX;
		// The BS_NOTIFY style is not a good idea for checkboxes because although it causes the control
		// to send BN_DBLCLK messages, any rapid clicks by the user on (for example) a tri-state checkbox
		// are seen only as one click for the purpose of changing the box's state.
		if (control.hwnd = CreateWindowEx(0, "button", aText, style|BS_MULTILINE
			, x, y, width, height, mHwnd, control_id, g_hInstance, NULL))
			if (checked != BST_UNCHECKED) // Set the specified state.
				SendMessage(control.hwnd, BM_SETCHECK, checked, 0);
		break;

	case GUI_CONTROL_RADIO:
		// The BS_NOTIFY style is probably better not applied by default to radios because although it
		// causes the control to send BN_DBLCLK messages, each double-click by the user is seen only
		// as one click for the purpose of cosmetically making the button appear like it is being
		// clicked rapidly.  Update: the usefulness of double-clicking a radio button seems to
		// outweigh the rare cosmetic deficiency of rapidly clicking a radio button, so it seems
		// better to allow it:
		if (control.hwnd = CreateWindowEx(0, "button", aText, style|BS_MULTILINE|BS_AUTORADIOBUTTON|BS_NOTIFY
			, x, y, width, height, mHwnd, control_id, g_hInstance, NULL))
		{
			if (checked != BST_UNCHECKED) // Set the specified state.
				SendMessage(control.hwnd, BM_SETCHECK, checked, 0);
			mInRadioGroup = true; // Only set now that creation was succesful.
		}
		break;

	case GUI_CONTROL_DROPDOWNLIST:
	case GUI_CONTROL_COMBOBOX:
		// Uses var_name as its title to give it a more friendly name by which it can be referred later:
		// Note: CBS_SIMPLE does not work if CBS_DROPDOWNLIST is also specified.
		if (control.hwnd = CreateWindowEx(WS_EX_CLIENTEDGE, "Combobox", var_name, style|WS_VSCROLL
			|CBS_AUTOHSCROLL|(aControlType == GUI_CONTROL_DROPDOWNLIST ? CBS_DROPDOWNLIST : CBS_DROPDOWN)
			, x, y, width, height, mHwnd, control_id, g_hInstance, NULL))
		{
			// 0 for last param: No redraw, since it's hidden:
			#define GUI_SETFONT \
			{\
				SendMessage(control.hwnd, WM_SETFONT, (WPARAM)sFont[mCurrentFontIndex].hfont, 0);\
				font_was_set = true;\
			}
			// Set font unconditionally to simplify calculations, which help ensure that at least one item
			// in the DropDownList/Combo is visible when the list drops down:
			GUI_SETFONT // Set font in preparation for asking it how tall each item is.
			item_height = (int)SendMessage(control.hwnd, CB_GETITEMHEIGHT, 0, 0);
			// Note that at this stage, height should contain a explicitly-specified height or height
			// estimate from the above, even if row_count is greater than 0.
			// Add 4 to make up for border between the always-visible control and the drop list:
			min_list_height = (2 * item_height) + GUI_CTL_VERTICAL_DEADSPACE + 4;
			if (height < min_list_height) // Adjust so that at least 1 item can be shown.
				height = min_list_height;
			else if (row_count > 0)
				// Now that we know the true item height (since the control has been created and we asked
				// it), resize the control to try to get it to the match the specified number of rows.
				height = (int)(row_count * item_height) + GUI_CTL_VERTICAL_DEADSPACE + 4;
			MoveWindow(control.hwnd, x, y, width, height, FALSE);
			// Since combo's size is created to accomodate its drop-down height, adjust our true height
			// to its actual collapsed size.  This true height is used for auto-positioning the next
			// control, if it uses auto-positioning.  It might be possible for it's width to be different
			// also, such as if it snaps to a certain minimize width if one too small was specified,
			// so that is recalculated too:
			retrieve_dimensions = true;
		}
		break;

	case GUI_CONTROL_LISTBOX:
		// Uses var_name as its title to give it a more friendly name to which it can be referred later:
		if (control.hwnd = CreateWindowEx(WS_EX_CLIENTEDGE, "Listbox", var_name, style|WS_VSCROLL|WS_BORDER
			|LBS_NOTIFY  // Omit LBS_STANDARD because it includes LBS_SORT, which we don't want as a default style.
			, x, y, width, height, mHwnd, control_id, g_hInstance, NULL))
		{
			// For now, it seems best to always override a height that would cause zero items to be
			// displayed.  This is because there is a very thin control visible even if the height
			// is explicitly set to zero, which seems pointless (there are other ways to draw thin
			// looking objects for unusual purposes anyway).
			// Set font unconditionally to simplify calculations, which help ensure that at least one item
			// in the DropDownList/Combo is visible when the list drops down:
			GUI_SETFONT // Set font in preparation for asking it how tall each item is.
			item_height = (int)SendMessage(control.hwnd, LB_GETITEMHEIGHT, 0, 0);
			// Note that at this stage, height should contain a explicitly-specified height or height
			// estimate from the above, even if row_count is greater than 0.
			min_list_height = item_height + GUI_CTL_VERTICAL_DEADSPACE;
			if (height < min_list_height) // Adjust so that at least 1 item can be shown.
				height = min_list_height;
			else if (row_count > 0)
				// Now that we know the true item height (since the control has been created and we asked
				// it), resize the control to try to get it to the match the specified number of rows.
				height = (int)(row_count * item_height) + GUI_CTL_VERTICAL_DEADSPACE;
			MoveWindow(control.hwnd, x, y, width, height, FALSE);
			// Since by default, the OS adjusts list's height to prevent a partial item from showing
			// (LBS_NOINTEGRALHEIGHT), fetch the actual height for possible use in positioning the
			// next control:
			retrieve_dimensions = true;
		}
		break;

	case GUI_CONTROL_EDIT:
		// malloc() is done because I think edit controls in NT/2k/XP support more than 64K?
		// Mem alloc errors are so rare (since the text is usually less than 32K/64K) that no error is displayed.
		// Instead, the un-translated text is put in directly.  Also, translation is not done for
		// single-line edits since they can't display linebreaks correctly anyway.
		malloc_buf = (*aText && (style & ES_MULTILINE)) ? TranslateLFtoCRLF(aText) : NULL;
		if (control.hwnd = CreateWindowEx(WS_EX_CLIENTEDGE, "edit", malloc_buf ? malloc_buf : aText, style|WS_BORDER
			, x, y, width, height, mHwnd, control_id, g_hInstance, NULL))
		{
			if (style & ES_PASSWORD && password_char != '*') // Override default.
				SendMessage(control.hwnd, EM_SETPASSWORDCHAR, (WPARAM)password_char, 0);
		}
		if (malloc_buf)
			free(malloc_buf);
		break;
	}

	// Below also serves as a bug check, i.e. GUI_CONTROL_INVALID or some unknown type.
	if (!control.hwnd)
		return g_script.ScriptError("The control could not be created." ERR_ABORT);
	// Otherwise the above control creation succeeed.
	control.type = aControlType;
	++mControlCount;

	///////////////////////////////////////////////////
	// Add any content to the control and set its font.
	///////////////////////////////////////////////////
	AddControlContent(control, aText, choice);
	// Must set the font even if mCurrentFontIndex > 0, otherwise the bold SYSTEM_FONT will be used
	if (!font_was_set && control.type != GUI_CONTROL_PIC)
		GUI_SETFONT

	if (retrieve_dimensions) // Update to actual size for use later below.
	{
		RECT rect;
		GetWindowRect(control.hwnd, &rect);
		height = rect.bottom - rect.top;
		width = rect.right - rect.left;
	}

	////////////////////////////////////////////////////////////////////////////////////////////////////
	// Save the details of this control's position for posible use in auto-positioning the next control.
	////////////////////////////////////////////////////////////////////////////////////////////////////
	mPrevX = x;
	mPrevY = y;
	mPrevWidth = width;
	mPrevHeight = height;
	int right = x + width;
	int bottom = y + height;
	if (right > mMaxExtentRight)
		mMaxExtentRight = right;
	if (bottom > mMaxExtentDown)
		mMaxExtentDown = bottom;

	if (start_new_section)
	{
		mSectionX = x;
		mSectionY = y;
		mMaxExtentRightSection = right;
		mMaxExtentDownSection = bottom;
	}
	else
	{
		if (right > mMaxExtentRightSection)
			mMaxExtentRightSection = right;
		if (bottom > mMaxExtentDownSection)
			mMaxExtentDownSection = bottom;
	}

	return OK;
}



void GuiType::AddControlContent(GuiControlType &aControl, char *aContent, int aChoice)
// Caller must ensure that aContent is a writable memory area, since this function temporarily
// alters the string.
{
	if (!*aContent)
		return;

	UINT msg_add, msg_select;

	switch (aControl.type)
	{
	case GUI_CONTROL_DROPDOWNLIST:
	case GUI_CONTROL_COMBOBOX:
		msg_add = CB_ADDSTRING;
		msg_select = CB_SETCURSEL;
		break;
	case GUI_CONTROL_LISTBOX:
		msg_add = LB_ADDSTRING;
		msg_select = LB_SETCURSEL;
		break;
	default: // Do nothing for any other control type that doesn't require content to be added this way.
		return;
	}

	bool temporarily_terminated;
	char *this_field, *next_field;
	int item_count;

	// Check *this_field at the top too, in case list ends in delimiter.
	for (this_field = aContent, item_count = 0; *this_field; ++item_count)
	{
		// Decided to use pipe as delimiter, rather than comma, because it makes the script more readable.
		// For example, it's easier to pick out the list of choices at a glance rather than having to
		// figure out where the commas delimit the beginning and end of "real" parameters vs. those params
		// that are a self-contained CSV list.  Of course, the pipe character itself is "sacrificed" and
		// cannot be used literally due to this method.  That limitation could be removed in a future
		// version by allowing a different delimiter to be optionally specified.
		if (next_field = strchr(this_field, '|')) // Assign
		{
			*next_field = '\0';  // Temporarily terminate (caller has ensured this is safe).
			temporarily_terminated = true;
		}
		else
		{
			next_field = this_field + strlen(this_field);  // Point it to the end of the string.
			temporarily_terminated = false;
		}
		SendMessage(aControl.hwnd, msg_add, 0, (LPARAM)this_field); // In this case, ignore any errors, namely CB_ERR/LB_ERR and CB_ERRSPACE).
		if (temporarily_terminated)
		{
			*next_field = '|';  // Restore the original char.
			++next_field;
			if (*next_field == '|')  // An item ending in two delimiters is a default (pre-selected) item.
			{
				SendMessage(aControl.hwnd, msg_select, (WPARAM)item_count, 0);  // Select this item.
				++next_field;  // Now this could be a third '|', which would in effect be an empty item.
				// It can also be the zero terminator if the list ends in a delimiter, e.g. item1|item2||
			}
		}
		this_field = next_field;
	} // for()

	// Have aChoice take precedence over any double-piped item(s) that appeared in the list:
	if (aChoice > 0)
		SendMessage(aControl.hwnd, msg_select, (WPARAM)(aChoice - 1), 0);  // Select this item.
}



ResultType GuiType::Show(char *aOptions, char *aText)
{
	if (!mHwnd)
		return OK;  // Make this a harmless attempt.

	int x = COORD_UNSPECIFIED;
	int y = COORD_UNSPECIFIED;
	int width = COORD_UNSPECIFIED;
	int height = COORD_UNSPECIFIED;

	for (char *cp = aOptions; *cp; ++cp)
	{
		switch(toupper(*cp))
		{
		// For options such as W, H, X and Y:
		// Use atoi() vs. ATOI() to avoid interpreting something like 0x01B as hex when in fact
		// the B was meant to be an option letter:
		// DIMENSIONS:
		case 'C':
			if (!strnicmp(cp, "center", 6))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 5 vs. 6 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 5;
				x = COORD_CENTERED;
				y = COORD_CENTERED;
			}
			break;
		case 'W':
			width = atoi(cp + 1);
			break;
		case 'H':
			// Allow any width/height to be specified so that the window can be "rolled up" to its title bar:
			height = atoi(cp + 1);
			break;
		case 'X':
			if (!strnicmp(cp + 1, "center", 6))
			{
				cp += 6; // 6 in this case since we're working with cp + 1
				x = COORD_CENTERED;
			}
			else
				x = atoi(cp + 1);
			break;
		case 'Y':
			if (!strnicmp(cp + 1, "center", 6))
			{
				cp += 6; // 6 in this case since we're working with cp + 1
				y = COORD_CENTERED;
			}
			else
				y = atoi(cp + 1);
			break;
		// Otherwise: Ignore other characters, such as the digits that occur after the P/X/Y option letters.
		} // switch()
	} // for()

	RECT work_rect;
	SystemParametersInfo(SPI_GETWORKAREA, 0, &work_rect, 0);  // Get Desktop rect excluding task bar.
	int work_width = work_rect.right - work_rect.left;  // Note that "left" won't be zero if task bar is on left!
	int work_height = work_rect.bottom - work_rect.top;  // Note that "top" won't be zero if task bar is on top!

	int width_orig = width;
	int height_orig = height;
	if (width == COORD_UNSPECIFIED)
		width = mMaxExtentRight + mMarginX;
	if (height == COORD_UNSPECIFIED)
		height = mMaxExtentDown + mMarginY;

	if (mFirstShowing) // By default, center the window if this is the first time it's being shown.
	{
		if (x == COORD_UNSPECIFIED)
			x = COORD_CENTERED;
		if (y == COORD_UNSPECIFIED)
			y = COORD_CENTERED;
	}
	mFirstShowing = false;

	// The above has determined the height/width of the client area.  From that area, determine
	// the window's new rect, including title bar, borders, etc.
	// If the window has a border or caption this also changes top & left *slightly* from zero.
	RECT rect = {0, 0, width, height}; // left,top,right,bottom
	AdjustWindowRectEx(&rect, GetWindowLong(mHwnd, GWL_STYLE), GetMenu(mHwnd) ? TRUE : FALSE
		, GetWindowLong(mHwnd, GWL_EXSTYLE));
	width = rect.right - rect.left;  // rect.left might be slightly less than zero.
	height = rect.bottom - rect.top; // rect.top might be slightly less than zero.

	// Seems best to restrict window size to the size of the desktop whenever explicit sizes
	// weren't given,  since most users would probably want that:
	if (width_orig == COORD_UNSPECIFIED && width > work_width)
		width = work_width;
	if (height_orig == COORD_UNSPECIFIED && height > work_height)
		height = work_height;

	if (x == COORD_CENTERED || y == COORD_CENTERED) // Center it, based on its dimensions determined above.
	{
		// This does not currently handle multi-monitor systems explicitly, since those calculations
		// require API functions that don't exist in Win95/NT (and thus would have to be loaded
		// dynamically to allow the program to launch).  Therefore, windows will likely wind up
		// being centered across the total dimensions of all monitors, which usually results in
		// half being on one monitor and half in the other.  This doesn't seem too terrible and
		// might even be what the user wants in some cases (i.e. for really big windows).
		if (x == COORD_CENTERED)
			x = work_rect.left + ((work_width - width) / 2);;
		if (y == COORD_CENTERED)
			y = work_rect.top + ((work_height - height) / 2);
	}

	BOOL is_visible = IsWindowVisible(mHwnd);
	RECT old_rect;
	GetWindowRect(mHwnd, &old_rect);
	int old_width = old_rect.right - old_rect.left;
	int old_height = old_rect.bottom - old_rect.top;

	if (width != old_width || height != old_height || (x != COORD_UNSPECIFIED && x != old_rect.left)
		|| (y != COORD_UNSPECIFIED && y != old_rect.bottom))
	{
		MoveWindow(mHwnd, x == COORD_UNSPECIFIED ? old_rect.left : x, y == COORD_UNSPECIFIED ? old_rect.top : y
			, width, height, is_visible);  // Do repaint if visible.
	}

	// Change the title before displaying it (makes transition a little nicer):
	if (*aText)
		SetWindowText(mHwnd, aText);

	if (!is_visible)
		ShowWindow(mHwnd, SW_SHOW);
	if (mHwnd != GetForegroundWindow()) // Normally it will be foreground since the template has this property.
		SetForegroundWindowEx(mHwnd);   // Try to force it to the foreground.

	return OK;
}



ResultType GuiType::PerformImplicitAction(GuiImplicitActions aImplicitAction)
{
	switch(aImplicitAction)
	{
	//case GUI_IMPLICIT_CLEAR:
	//	return Clear();
	case GUI_IMPLICIT_CANCEL:
		return Cancel();
	}
	// Otherwise, do nothing (caller may rely on this):
	return OK;
}



ResultType GuiType::Clear() // Not implemented yet.
{
	//if (!mHwnd) // Operating on a non-existent GUI has no effect.
	//	return OK;
	return OK;
}



ResultType GuiType::Cancel()
{
	if (mHwnd)
		ShowWindow(mHwnd, SW_HIDE);
	return OK;
}



ResultType GuiType::Close()
// If there is a GuiClose label defined in for this event, launch it as a new thread.
// In this case, don't close or hide the window.  It's up to the subroutine to do that
// if it wants to.
// If there is no label, treat it the same as Cancel().
{
	if (!mLabelForClose)
		return Cancel();
	// See lengthy comments in Event() about this section:
	POST_AHK_GUI_ACTION(mHwnd, AHK_GUI_CLOSE, GUI_EVENT_NORMAL);
	MsgSleep(-1);
	return OK;
}



ResultType GuiType::Escape() // Similar to close, except typically called when the user presses ESCAPE.
// If there is a GuiEscape label defined in for this event, launch it as a new thread.
// In this case, don't close or hide the window.  It's up to the subroutine to do that
// if it wants to.
// If there is no label, treat it the same as Cancel().
{
	if (!mLabelForEscape) // The user preference (via votes on forum poll) is to do nothing by default.
		return OK;
	// See lengthy comments in Event() about this section:
	POST_AHK_GUI_ACTION(mHwnd, AHK_GUI_ESCAPE, GUI_EVENT_NORMAL);
	MsgSleep(-1);
	return OK;
}



ResultType GuiType::Submit(bool aHideIt)
// Caller has ensured that all controls have valid, non-NULL hwnds:
{
	if (!mHwnd) // Operating on a non-existent GUI has no effect.
		return OK;

	// Handle all non-radio controls:
	UINT u;
	for (u = 0; u < mControlCount; ++u)
		if (mControl[u].output_var && mControl[u].type != GUI_CONTROL_RADIO)
			ControlGetContents(*mControl[u].output_var, mControl[u], false);

	// Handle GUI_CONTROL_RADIO separately so that any radio group that has a single variable
	// to share among all its members can be given special treatment:
	int group_radios = 0;          // The number of radio buttons in the current group.
	int group_radios_with_var = 0; // The number of the above that have an output var.
	Var *group_var = NULL;         // The last-found output variable of the current group.
	int selection_number = 0;      // Which radio in the current group is selected (0 if none).
	Var *output_var;
	char temp[32];
    
	// The below uses <= so that it goes one beyond the limit.  This allows the final radio group
	// (if any) to be noticed in the case where the very last control in the window is a radio button.
	// This is because in such a case, there is no "terminating control" having the WS_GROUP style:
	for (u = 0; u <= mControlCount; ++u)
	{
		// WS_GROUP is used to determine where one group ends and the next begins -- rather than using
		// seeing if the control's type is radio -- because in the future it may be possible for a radio
		// group to have other controls interspersed within it and yet still be a group for the purpose
		// of "auto radio button (only one selected at a time)" behavior:
		if (u == mControlCount || GetWindowLong(mControl[u].hwnd, GWL_STYLE) & WS_GROUP) // New group. Relies on short-circuit boolean order.
		{
			// If the prior group had exactly one output var but more than one radio in it, that
			// var is shared among all radios (as of v1.0.20).  Otherwise:
			// 1) If it has zero radios and/or zero variables: already fully handled by other logic.
			// 2) It has multiple variables: the default values assigned in the loop are simply retained.
			// 3) It has exactly one radio in it and that radio has an output var: same as above.
			if (group_radios_with_var == 1 && group_radios > 1)
			{
				// Multiple buttons selected.  Since this is so rare, don't give it a distinct value.
				// Instead, treat this the same as "none selected":
				if (selection_number == -1)
					selection_number = 0;
				// Convert explicitly to decimal so that g.FormatIntAsHex is not obeyed.
				// This is so that this result matches the decimal format tradition set by
				// the "1" and "0" strings normally used for radios and checkboxes:
				_itoa(selection_number, temp, 10); // selection_number can be legitimately zero.
				group_var->Assign(temp); // group_var should not be NULL since group_radios_with_var == 1
			}
			if (u == mControlCount) // The last control in the window is a radio and its group was just processed.
				break;
			group_radios = group_radios_with_var = selection_number = 0;
		}
		if (mControl[u].type == GUI_CONTROL_RADIO)
		{
			++group_radios;
			if (output_var = mControl[u].output_var) // Assign.
			{
				++group_radios_with_var;
				group_var = output_var; // If this group winds up having only one var, this will be it.
			}
			// Assign default value for now.  It will be overridden if this turns out to be the
			// only variable in this group:
			if (SendMessage(mControl[u].hwnd, BM_GETCHECK, 0, 0) == BST_CHECKED)
			{
				if (selection_number) // Multiple buttons selected, so flag this as an indeterminate state.
					selection_number = -1;
				else
					selection_number = group_radios;
				if (output_var)
					output_var->Assign("1");
			}
			else
				if (output_var)
					output_var->Assign("0");
		}
	} // for()

	if (aHideIt)
		ShowWindow(mHwnd, SW_HIDE);
	return OK;
}



ResultType GuiType::ControlGetContents(Var &aOutputVar, GuiControlType aControl, bool aGetText)
{
	if (!aGetText)
	{
		// The caller wants the contents of the control, which is often different from its
		// caption/text.  Any control types not mentioned in the switch() below will fall through
		// into the section at the bottom that applies the standard GetWindowText() method.

		LRESULT index, length;

		switch (aControl.type)
		{
		case GUI_CONTROL_CHECKBOX:
		case GUI_CONTROL_RADIO:
			// Submit() handles GUI_CONTROL_RADIO on its own, but other callers might need us to handle it.
			// In addition, rather than handling multi-radio groups that share a single output variable
			// in a special way, it's kept simple here because:
			// 1) It's more flexible (there might be cases when the user wants to get the value of
			//    a single radio in the group, not the group's currently-checked button).
			// 2) The multi-radio handling seems too complex to be justified given how rarely users would
			//    want such behavior (since "Submit, NoHide" can be used as a substitute).
			switch (SendMessage(aControl.hwnd, BM_GETCHECK, 0, 0))
			{
			case BST_CHECKED:
				return aOutputVar.Assign("1");
			case BST_UNCHECKED:
				return aOutputVar.Assign("0");
			case BST_INDETERMINATE:
				// Seems better to use a value other than blank because blank might sometimes represent the
				// state of an unintialized or unfetched control.  In other words, a blank variable often
				// has an external meaning that transcends the more specific meaning often desirable when
				// retrieving the state of the control:
				return aOutputVar.Assign("-1");
			}
			return FAIL; // Shouldn't be reached since ZERO(BST_UNCHECKED) is returned on failure.

		case GUI_CONTROL_COMBOBOX: // Not needed for GUI_CONTROL_DROPDOWNLIST since GetWindowText() is enough
			index = SendMessage(aControl.hwnd, CB_GETCURSEL, 0, 0); // Get index of currently selected item.
			if (index == CB_ERR) // There is no selection (or very rarely, some other type of problem).
				break; // Break out of the switch rather than returning so that the GetWindowText() method can be applied.
			length = SendMessage(aControl.hwnd, CB_GETLBTEXTLEN, (WPARAM)index, 0);
			if (length == CB_ERR) // Given the way it was called, this should be impossible based on MSDN docs.
				return aOutputVar.Assign();
			// In unusual cases, MSDN says the indicated length might be longer than it actually winds up
			// being when the item's text is retrieved.  This should be harmless, since there are many
			// other precedents where a variable is sized to something larger than it winds up carrying.
			// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
			// this call will set up the clipboard for writing:
			if (aOutputVar.Assign(NULL, (VarSizeType)length) != OK)
				return FAIL;  // It already displayed the error.
			length = SendMessage(aControl.hwnd, CB_GETLBTEXT, (WPARAM)index, (LPARAM)aOutputVar.Contents());
			if (length == CB_ERR) // Given the way it was called, this should be impossible based on MSDN docs.
			{
				aOutputVar.Close(); // In case it's the clipboard.
				return aOutputVar.Assign();
			}
			aOutputVar.Length() = (VarSizeType)length;  // Update it to the actual length, which can vary from the estimate.
			return aOutputVar.Close(); // In case it's the clipboard.

		case GUI_CONTROL_LISTBOX:
			index = SendMessage(aControl.hwnd, LB_GETCURSEL, 0, 0); // Get index of currently selected item.
			if (index == LB_ERR) // There is no selection (or very rarely, some other type of problem).
				return aOutputVar.Assign();
			length = SendMessage(aControl.hwnd, LB_GETTEXTLEN, (WPARAM)index, 0);
			if (length == LB_ERR) // Given the way it was called, this should be impossible based on MSDN docs.
				return aOutputVar.Assign();
			// In unusual cases, MSDN says the indicated length might be longer than it actually winds up
			// being when the item's text is retrieved.  This should be harmless, since there are many
			// other precedents where a variable is sized to something larger than it winds up carrying.
			// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
			// this call will set up the clipboard for writing:
			if (aOutputVar.Assign(NULL, (VarSizeType)length) != OK)
				return FAIL;  // It already displayed the error.
			length = SendMessage(aControl.hwnd, LB_GETTEXT, (WPARAM)index, (LPARAM)aOutputVar.Contents());
			if (length == LB_ERR) // Given the way it was called, this should be impossible based on MSDN docs.
			{
				aOutputVar.Close(); // In case it's the clipboard.
				return aOutputVar.Assign();
			}
			aOutputVar.Length() = (VarSizeType)length;  // Update it to the actual length, which can vary from the estimate.
			return aOutputVar.Close(); // In case it's the clipboard.
		} // switch()
	} // if (!aGetText)

	// Since the above didn't return, at lest one of the following is true:
	// 1) aGetText is true (the caller wanted the simple GetWindowText() method applied unconditionally).
	// 2) This control's type is not mentioned in the switch because it does not require special handling.
	//   e.g.  GUI_CONTROL_EDIT, GUI_CONTROL_DROPDOWNLIST, and others that use a simple GetWindowText().
	// 3) This control is a ComboBox, but it lacks a selected item, so any text entered by the user
	//    into the control's edit field is fetched instead.

	// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
	// this call will set up the clipboard for writing:
	int length = GetWindowTextLength(aControl.hwnd); // Might be zero, which is properly handled below.
	if (aOutputVar.Assign(NULL, (VarSizeType)length) != OK)
		return FAIL;  // It already displayed the error.
	// Update length using the actual length, rather than the estimate provided by GetWindowTextLength():
	if (   !(aOutputVar.Length() = (VarSizeType)GetWindowText(aControl.hwnd, aOutputVar.Contents(), (int)(length + 1)))   )
		// There was no text to get.  Set to blank explicitly just to be sure.
		*aOutputVar.Contents() = '\0';  // Safe because Assign() gave us a non-constant memory area.
	else if (aControl.type == GUI_CONTROL_EDIT) // Auto-translate CRLF to LF for better compatibility with other script commands.
	{
		StrReplaceAll(aOutputVar.Contents(), "\r\n", "\n");
		aOutputVar.Length() = (VarSizeType)strlen(aOutputVar.Contents());
	}
	return aOutputVar.Close();  // In case it's the clipboard.
}



UINT GuiType::FindControl(char *aControlID)
// Find the index of the control that matches the string, which can be either:
// 1) The name of a control's associated output variable.
// 2) Class+NN
// 3) Control's title/caption.
// Returns -1 if not found.
{
	// To keep things simple, the first search method is always conducted: It looks for a
	// matching variable name, but only among the variables used by this particular window's
	// controls (i.e. avoid ambiguity by NOT having earlier matched up aControlID against
	// all variable names in the entire script, perhaps in PreparseBlocks() or something):
	UINT u;
	for (u = 0; u < mControlCount; ++u)
		if (mControl[u].output_var && !stricmp(mControl[u].output_var->mName, aControlID)) // Relies on short-circuit boolean order.
			return u;  // Match found.
	// Otherwise: No match found, so fall back to standard control class and/or text finding method.
	HWND control_hwnd = ControlExist(mHwnd, aControlID);
	if (!control_hwnd)
		return -1; // No match found.
	for (u = 0; u < mControlCount; ++u)
		if (mControl[u].hwnd == control_hwnd)
			return u;  // Match found.
	// Otherwise: No match found.  At this stage, should be impossible if design is correct.
	return -1;
}



int GuiType::FindGroup(UINT aControlIndex, UINT &aGroupStart, UINT &aGroupEnd)
// Caller must provide a valid aControlIndex for an existing control.
// Returns the number of radio buttons inside the group. In addition, it provides start and end
// values to the caller via aGroupStart/End, where Start is the index of the first control in
// the group and End is the index of the control *after* the last control in the group (if none,
// aGroupEnd is set to mControlCount).
// NOTE: This returns the range covering the entire group, and it is possible for the group
// to contain non-radio type controls.  Thus, the caller should check each control inside the
// returned range to make sure it's a radio before operating upon it.
{
	// Work backwards in the control array until the first member of the group is found or the
	// first array index, whichever comes first (the first array index is the top control in the
	// Z-Order and thus treated by the OS as an implicit start of a new group):
	int group_radios = 0; // This and next are both init'd to 0 not 1 because the first loop checks aControlIndex itself.
	for (aGroupStart = aControlIndex;; --aGroupStart)
	{
		if (mControl[aGroupStart].type == GUI_CONTROL_RADIO)
			++group_radios;
		if (!aGroupStart || GetWindowLong(mControl[aGroupStart].hwnd, GWL_STYLE) & WS_GROUP)
			break;
	}
	// Now find the control after the last control (or mControlCount if none).  Must start at +1
	// because if aControlIndex's control has the WS_GROUP style, we don't want to count that
	// as the end of the group (because in fact that would be the beginning of the group).
	for (aGroupEnd = aControlIndex + 1; aGroupEnd < mControlCount; ++aGroupEnd)
	{
		// Unlike the previous loop, this one must do this check prior to the next one:
		if (GetWindowLong(mControl[aGroupEnd].hwnd, GWL_STYLE) & WS_GROUP)
			break;
		if (mControl[aGroupEnd].type == GUI_CONTROL_RADIO)
			++group_radios;
	}
	return group_radios;
}



ResultType GuiType::SetCurrentFont(char *aOptions, char *aFontName)
{
	COLORREF color;
	int font_index = FindOrCreateFont(aOptions, aFontName, &sFont[mCurrentFontIndex], &color);
	if (color != CLR_NONE) // Even if the above call failed, it returns a color if one was specified.
		mCurrentColor = color;
	if (font_index >= 0) // Success.
	{
		mCurrentFontIndex = font_index;
		return OK;
	}
	// Failure of the above is rare because it falls back to default typeface if the one specified
	// isn't found.  It will have already displayed the error:
	return FAIL;
}



int GuiType::FindOrCreateFont(char *aOptions, char *aFontName, FontType *aFoundationFont, COLORREF *aColor)
// Returns the index of existing or new font within the sFont array (an index is returned so that
// caller can see that this is the default-gui-font whenever index 0 is returned).  Returns -1
// on error, but still sets *aColor to be the color name, if any was specified in aOptions.
// To prevent a large number of font handles from being created (such as one for each control
// that uses something other than GUI_DEFAULT_FONT), it seems best to conserve system resources
// by creating new fonts only when called for.  Therefore, this function will first check if
// the specified font already exists within the array of fonts.  If not found, a new font will
// be added to the array.
{
	// Set default output parameter in case of early return:
	if (aColor) // Caller wanted color returned in an output parameter.
		*aColor = CLR_NONE; // Because we want CLR_DEFAULT to indicate a real color.

	HDC hdc;

	if (!*aOptions && !*aFontName)
	{
		// Relies on the fact that first item in the font array is always the default font.
		// If there are fonts, the default font should be the first one (index 0).
		// If not, we create it here:
		if (!sFontCount)
		{
			// Otherwise, for simplifying other code sections, create an entry in the array for the default font:
			// Doesn't seem likely that DEFAULT_GUI_FONT face/size will change while a script is running,
			// or even while the system is running for that matter.  I think it's always an 8 or 9 point
			// font regardless of desktop's appearance/theme settings.
			ZeroMemory(&sFont[sFontCount], sizeof(FontType));
			// SYSTEM_FONT seems to be the bold one that is used in a dialog window by default.
			// MSDN: "It is not necessary (but it is not harmful) to delete stock objects by calling DeleteObject."
			sFont[sFontCount].hfont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
			// Get attributes of DEFAULT_GUI_FONT (name, size, etc.)
			hdc = GetDC(HWND_DESKTOP);
			HFONT hfont_old = (HFONT)SelectObject(hdc, sFont[sFontCount].hfont);
			GetTextFace(hdc, sizeof(sFont[sFontCount].name) - 1, sFont[sFontCount].name);
			TEXTMETRIC tm;
			GetTextMetrics(hdc, &tm);
			// Convert height to points.  Use MulDiv's build-in rounding to get a more accurate result.
			// This is confirmed to be the correct forumla to convert tm's height to font point size,
			// and it does yield 8 for DEFAULT_GUI_FONT as it should:
			sFont[sFontCount].point_size = MulDiv(tm.tmHeight - tm.tmInternalLeading, 72, GetDeviceCaps(hdc, LOGPIXELSY));
			sFont[sFontCount].weight = tm.tmWeight;
			// Probably unnecessary for default font, but just to be consistent:
			sFont[sFontCount].italic = (bool)tm.tmItalic;
			sFont[sFontCount].underline = (bool)tm.tmUnderlined;
			sFont[sFontCount].strikeout = (bool)tm.tmStruckOut;
			SelectObject(hdc, hfont_old); // Necessary to avoid memory leak.
			ReleaseDC(HWND_DESKTOP, hdc);
			++sFontCount;
		}
		// Tell caller to return to default color, since this is documented behavior when
		// returning to default font:
		if (aColor) // Caller wanted color returned in an output parameter.
			*aColor = CLR_DEFAULT;
		return 0;  // Always returns 0 since that is always the index of the default font.
	}

	// Otherwise, a non-default name/size, etc. is being requested.  Find or create a font to match it.
	if (!aFoundationFont) // Caller didn't specify a font whose attributes should be used as default.
	{
		if (sFontCount > 0)
			aFoundationFont = &sFont[0]; // Use default if it exists.
		else
			return -1; // No error displayed since shouldn't happen if things are designed right.
	}

	// Copy the current default font's attributes into our local font structure.
	// The caller must ensure that mCurrentFontIndex array element exists:
	FontType font = *aFoundationFont;
	if (*aFontName)
		strlcpy(font.name, aFontName, sizeof(font.name));
	COLORREF color = CLR_NONE; // Because we want to treat CLR_DEFAULT as a real color.

	// Temp vars:
	char color_str[32], *space_pos;

	for (char *cp = aOptions; *cp; ++cp)
	{
		switch(toupper(*cp))
		{
		case 'B':
			if (!strnicmp(cp, "bold", 4))
			{
				font.weight = FW_BOLD;
				cp += 3;  // Skip over the word itself to prevent next interation from seeing it as option letters.
			}
			break;

		case 'I':
			if (!strnicmp(cp, "italic", 6))
			{
				font.italic = true;
				cp += 5;  // Skip over the word itself to prevent next interation from seeing it as option letters.
			}
			break;

		case 'N':
			if (!strnicmp(cp, "norm", 4))
			{
				font.italic = false;
				font.underline = false;
				font.strikeout = false;
				font.weight = FW_NORMAL;
				cp += 3;  // Skip over the word itself to prevent next interation from seeing it as option letters.
			}
			break;

		case 'U':
			if (!strnicmp(cp, "underline", 9))
			{
				font.underline = true;
				cp += 8;  // Skip over the word itself to prevent next interation from seeing it as option letters.
			}
			break;

		case 'C': // Color
			strlcpy(color_str, cp + 1, sizeof(color_str));
			if (space_pos = StrChrAny(color_str, " \t"))  // space or tab
				*space_pos = '\0';
			//else a color name can still be present if it's at the end of the string.
			color = ColorNameToBGR(color_str);
			if (color == CLR_NONE) // A matching color name was not found, so assume it's in hex format.
			{
				if (strlen(color_str) > 6)
					color_str[6] = '\0';  // Shorten to exactly 6 chars, which happens if no space/tab delimiter is present.
				color = rgb_to_bgr(strtol(color_str, NULL, 16));
				// if color_str does not contain something hex-numeric, black (0x00) will be assumed,
				// which seems okay given how rare such a problem would be.
			}
			// Skip over the color string to avoid interpreting hex digits or color names as option letters:
			cp += strlen(color_str);
			break;

		// For options such as S and W:
		// Use atoi()/atof() vs. ATOI()/ATOF() to avoid interpreting something like 0x01B as hex when in fact
		// the B was meant to be an option letter:
		case 'S':
			// Seems best to allow fractional point sizes via atof, though it might usually get rounded
			// by the OS anyway (at the time font is created):
			if (!strnicmp(cp, "strike", 6))
			{
				font.strikeout = true;
				cp += 5;  // Skip over the word itself to prevent next interation from seeing it as option letters.
			}
			else
				font.point_size = (int)(atof(cp + 1) + 0.5);  // Round to nearest int.
			break;

		case 'W':
			font.weight = atoi(cp + 1);
			break;

		// Otherwise: Ignore other characters, such as the digits that occur after the P/X/Y option letters.
		} // switch()
	} // for()

	if (aColor) // Caller wanted color returned in an output parameter.
		*aColor = color;

	hdc = GetDC(HWND_DESKTOP);
	// Fetch the value every time in case it can change while the system is running (e.g. due to changing
	// display to TV-Out, etc).  In addition, this HDC is needed by 
	int pixels_per_point_y = GetDeviceCaps(hdc, LOGPIXELSY);

	// The reason it's done this way is that CreateFont() does not always (ever?) fail if given a
	// non-existent typeface:
	if (!FontExist(hdc, font.name)) // Fall back to foundation font's type face, as documented.
		strcpy(font.name, aFoundationFont ? aFoundationFont->name : sFont[0].name);

	ReleaseDC(HWND_DESKTOP, hdc);
	hdc = NULL;

	// Now that the attributes of the requested font are known, see if such a font already
	// exists in the array:
	int font_index = FindFont(font);
	if (font_index != -1) // Match found.
		return font_index;

	// Since above didn't return, create the font if there's room.
	if (sFontCount >= MAX_GUI_FONTS)
	{
		g_script.ScriptError("Too many fonts." ERR_ABORT);  // Short msg since so rare.
		return -1;
	}

	// MulDiv() is usually better because it has automatic rounding, getting the target font
	// closer to the size specified:
	if (   !(font.hfont = CreateFont(-MulDiv(font.point_size, pixels_per_point_y, 72), 0, 0, 0
		, font.weight, font.italic, font.underline, font.strikeout
		, DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, FF_DONTCARE, font.name))   )
		// OUT_DEFAULT_PRECIS/OUT_TT_PRECIS ... DEFAULT_QUALITY/PROOF_QUALITY
	{
		g_script.ScriptError("Can't create font." ERR_ABORT);  // Short msg since so rare.
		return -1;
	}

	sFont[sFontCount++] = font; // Copy the newly created font's attributes into the next array element.
	return sFontCount - 1; // The index of the newly created font.
}



int GuiType::FindFont(FontType &aFont)
{
	for (int i = 0; i < sFontCount; ++i)
		if (!stricmp(sFont[i].name, aFont.name)
			&& sFont[i].point_size == aFont.point_size
			&& sFont[i].weight == aFont.weight
			&& sFont[i].italic == aFont.italic
			&& sFont[i].underline == aFont.underline
			&& sFont[i].strikeout == aFont.strikeout) // Match found.
			return i;
	return -1;  // Indicate failure.
}



LRESULT CALLBACK GuiWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
	GuiType *pgui;
	GuiControlType *pcontrol;

	switch (iMsg)
	{
	// Let the default handler take care of WM_CREATE.

	case WM_COMMAND:
	{
		// First find which of the GUI windows is receiving this event, if none (probably impossible
		// the way things are set up currently), let DefaultProc handle it:
		if (   !(pgui = GuiType::FindGui(hWnd))   )
			break;
		WORD wParam_loword = LOWORD(wParam);
		if (wParam_loword >= ID_USER_FIRST)
		{
			// Since all control id's are less than ID_USER_FIRST, this message is either
			// a user defined menu item ID or a bogus message due to it corresponding to
			// a non-existent menu item or a main/tray menu item (which should never be
			// received or processed here).
			HandleMenuItem(wParam_loword, pgui->mWindowIndex);
			return 0; // Return unconditionally since it's not in the correct range to be a control ID.
		}
		// Since this even is not a menu item, see if it's for a control inside the window.
		if (wParam_loword == IDOK) // The user pressed ENTER, regardless of whether there is a default button.
		{
			// If there is no default button, just let default proc handle it:
			if (pgui->mDefaultButtonIndex >= pgui->mControlCount)
				break;
			// Otherwise, take the same action as clicking the button:
			pgui->Event(pgui->mDefaultButtonIndex, BN_CLICKED);
			return 0;
		}
		else if (wParam_loword == IDCANCEL) // The user pressed ESCAPE.
		{
			pgui->Escape();
			return 0;
		}
		UINT control_index = wParam_loword - CONTROL_ID_FIRST; // Convert from ID to array index.
		if (control_index < pgui->mControlCount // Relies on short-circuit eval order.
			&& pgui->mControl[control_index].hwnd == (HWND)lParam) // Handles match (this filters out bogus msgs).
			pgui->Event(control_index, HIWORD(wParam));
		//else ignore this msg.
		return 0;
	}

	case WM_SYSCOMMAND:
		if (wParam == SC_CLOSE)
		{
			if (   !(pgui = GuiType::FindGui(hWnd))   )
				break; // Let DefDlgProc() handle it.
			pgui->Close();
			return 0;
		}
		break;

	case WM_ERASEBKGND:
		if (   !(pgui = GuiType::FindGui(hWnd))   )
			break; // Let DefDlgProc() handle it.
		if (!pgui->mBackgroundBrushWin) // Let default proc handle it.
			break;
		// Can't use SetBkColor(), need an real brush to fill it.
		RECT clipbox;
		GetClipBox((HDC)wParam, &clipbox);
		FillRect((HDC)wParam, &clipbox, pgui->mBackgroundBrushWin);
		return 1; // "An application should return nonzero if it erases the background."

	case WM_CTLCOLORSTATIC: // WM_CTLCOLORDLG is probably never received by this type of dialog window.
	case WM_CTLCOLORLISTBOX:
	case WM_CTLCOLOREDIT:
	// MSDN: Buttons with the BS_PUSHBUTTON, BS_DEFPUSHBUTTON, or BS_PUSHLIKE styles do not use the
	// returned brush. Buttons with these styles are always drawn with the default system colors.
	// Drawing push buttons requires several different brushes-face, highlight and shadow-but the
	// WM_CTLCOLORBTN message allows only one brush to be returned. To provide a custom appearance
	// for push buttons, use an owner-drawn button.
	// Thus, WM_CTLCOLORBTN is commented out for now since it doesn't seem to have any effect on the
	// types of buttons used so far (not even checkboxes, whose text is probably handled by
	// WM_CTLCOLORSTATIC:
	//case WM_CTLCOLORBTN:
		if (   !(pgui = GuiType::FindGui(hWnd))   )
			break;
		if (   !(pcontrol = pgui->FindControl((HWND)lParam))   )
			break;
		if (pcontrol->type == GUI_CONTROL_COMBOBOX) // But GUI_CONTROL_DROPDOWNLIST partially works.
			// Setting the colors of combo boxes won't work without subclassing the child controls
			// via Get/SetWindowLong & GWL_WNDPROC.  That would be fairly complicated because when
			// the incoming HWND/LParam is one of the children of a combobox, we would have to have
			// a way of looking up that child to see that it really is for a combo (unless we assumed
			// all "unfound" hWnds are children of combos, which would work if there are ever any
			// other controls that also would need to be subclassed, such as date-time).  This is
			// because the combo's original WindowProc should always be called unless its childrens'
			// messages were fully handled here (which is true only in certain cases).
			break;
		if (pcontrol->color != CLR_DEFAULT)
			SetTextColor((HDC)wParam, pcontrol->color);
		if (iMsg == WM_CTLCOLORSTATIC)
		{
			if (pgui->mBackgroundBrushWin)
			{
				// Since we're processing this msg rather than passing it on to the default proc, must set
				// background color unconditionally, otherwise plain white will likely be used:
				SetBkColor((HDC)wParam, pgui->mBackgroundColorWin);
				// Always return a real HBRUSH so that Windows knows we altered the HDC for it to use:
				return (LRESULT)pgui->mBackgroundBrushWin;
			}
		}
		else // It's for the interior of a non-static control.  Use the control background color (if there is one).
		{
			if (pgui->mBackgroundBrushCtl)
			{
				SetBkColor((HDC)wParam, pgui->mBackgroundColorCtl);
				// Always return a real HBRUSH so that Windows knows we altered the HDC for it to use:
				return (LRESULT)pgui->mBackgroundBrushCtl;
			}
		}
		// Since above didn't return a custom HBRUSH, we must return one here -- rather than letting the
		// default proc handle this message -- if the color of the text itself was changed.  This is so
		// that the OS will know that the DC has been altered:
		if (pcontrol->color != CLR_DEFAULT)
		{
			// Whenever the default proc won't be handling this message, the background color must be set
			// explicitly if something other than plain white is needed.  This must be done even for
			// non-static controls because otherwise the area doesn't get filled correctly:
			if (iMsg == WM_CTLCOLORSTATIC)
			{
				SetBkColor((HDC)wParam, GetSysColor(COLOR_BTNFACE));
				return (LRESULT)GetSysColorBrush(COLOR_BTNFACE);
			}
			else
			{
				// I'm pretty sure that COLOR_WINDOW is the color used by default for the background of
				// all standard controls, such as ListBox, ComboBox, Edit, etc.  Although it's usually
				// white, it can be different depending on theme/appearance settings:
				SetBkColor((HDC)wParam, GetSysColor(COLOR_WINDOW));
				return (LRESULT)GetSysColorBrush(COLOR_WINDOW);
			}
		}
		//else since no colors were changed, let default proc handle it.
		break;

	case WM_CLOSE: // For now, take the same action as SC_CLOSE.
		if (   !(pgui = GuiType::FindGui(hWnd))   )
			break; // Let DefDlgProc() handle it.
		pgui->Close();
		return 0;

	case WM_DESTROY:
		// Update to below: If a GUI window is owned by the script's main window (via "gui +owner"),
		// it can be destroyed automatically.  Because of this and the fact that it's difficult to
		// keep track of all the ways a window can be destroyed, it seems best for peace-of-mind to
		// have this WM_DESTROY handler 
		// Older Note: Let default-proc handle WM_DESTROY because with the current design, it should
		// be impossible for a window to be destroyed without the object "knowing about it" and
		// updating itself (then destroying itself) accordingly.  The object methods always
		// destroy (recursively) any windows it owns, so once again it "knows about it".
		if (pgui = GuiType::FindGui(hWnd)) // Assign.
			if (!pgui->mDestroyWindowHasBeenCalled)
			{
				pgui->mDestroyWindowHasBeenCalled = true; // Tell it not to call DestroyWindow(), just clean up everything else.
				GuiType::Destroy(pgui->mWindowIndex);
			}
		// Above: if mDestroyWindowHasBeenCalled==true, we were called by Destroy(), so don't call Destroy() again recursively.
		// And in any case, pass it on to DefDlgProc() in case it does any extra cleanup:
		break;

	// Cases for WM_ENTERMENULOOP and WM_EXITMENULOOP:
	HANDLE_MENU_LOOP

	} // switch()

	// This will handle anything not already fully handled and returned from above:
	return DefDlgProc(hWnd, iMsg, wParam, lParam);
}



void GuiType::Event(UINT aControlIndex, WORD aNotifyCode)
// Handles events within a GUI window that caused one of its controls to change in a meaningful way,
// or that is an event that could trigger an external action, such as clicking a button or icon.
{
	if (aControlIndex >= MAX_CONTROLS_PER_GUI) // Caller probably already checked, but just to be safe.
		return;
	GuiControlType &control = mControl[aControlIndex];
	if (!control.jump_to_label && !control.implicit_action)
		return;

	// Update: The below is now checked by MsgSleep() at the time the launch actually would occur:
	// If this control already has a thread running in its label, don't create a new thread to avoid
	// problems of buried threads, or a stack of suspended threads that might be resumed later
	// at an unexpected time. Users of timer subs that take a long time to run should be aware, as
	// documented in the help file, that long interruptions are then possible.
	//if (g_nThreads >= g_MaxThreadsTotal || aControl->label_is_running)
	//	continue

	GuiEventType gui_event = GUI_EVENT_NORMAL;  // Set default.  Don't use NONE since that means "not a GUI thread".

	// Explicitly cover all control types in the swithc() rather than relying solely on 
	// aNotifyCode in case it's ever possible for the code to be context-sensitive
	// depending on the type of control.
	switch(control.type)
	{
	case GUI_CONTROL_BUTTON:
	case GUI_CONTROL_CHECKBOX:
	case GUI_CONTROL_RADIO:
		// Must include BN_DBLCLK or these control types won't be responsive to rapid consecutive clicks.
		// Update: The above is true only if the button has the BS_NOTIFY option, and now it doesn't so
		// checking for BN_DBLCLK is no longer necessary.  Update: Double-clicks are now detected in
		// case that style every winds up on any of the above control types (currently it's the default
		// on GUI_CONTROL_RADIO anyway):
		switch (aNotifyCode)
		{
		case BN_CLICKED: // Must explicitly list this case since the default label does a return.
			break;
		case BN_DBLCLK:
			gui_event = GUI_EVENT_DBLCLK;
			break;
		default:
			return;
		}
		break;
	case GUI_CONTROL_DROPDOWNLIST:
	case GUI_CONTROL_COMBOBOX:
		switch (aNotifyCode)
		{
		case CBN_SELCHANGE: // Must explicitly list this case since the default label does a return.
			break;
		case CBN_DBLCLK: // Needed in case CBS_SIMPLE (i.e. list always visible) is ever possible.
			gui_event = GUI_EVENT_DBLCLK;
			break;
		default:
			return;
		}
		break;
	case GUI_CONTROL_LISTBOX:
		switch (aNotifyCode)
		{
		case LBN_SELCHANGE: // Must explicitly list this case since the default label does a return.
			break;
		case LBN_DBLCLK:
			gui_event = GUI_EVENT_DBLCLK;
			break;
		default:
			return;
		}
		break;
	case GUI_CONTROL_TEXT:
	case GUI_CONTROL_PIC:
		// Update: Unlike buttons, it's all-or-none for static controls.  Testing shows that if
		// STN_DBLCLK is not checked for and the user clicks rapidly, half the clicks will be
		// ignored:
		// Based on experience with BN_DBLCLK, it's likely that STN_DBLCLK must be included or else
		// these control types won't be responsive to rapid consecutive clicks:
		switch (aNotifyCode)
		{
		case STN_CLICKED: // Must explicitly list this case since the default label does a return.
			break;
		case STN_DBLCLK:
			gui_event = GUI_EVENT_DBLCLK;
			break;
		default:
			return;
		}
		break;
	default: // For other controls or other types of messages, take no action.
		return;
	}

	// Rather than launching the thread directly from here, it seems best to always post it to our
	// thread to be handled there.  Here are the reasons:
	// 1) We don't want to be in the situation where a thread launched here would return first
	//    to a dialog's message pump rather than MsgSleep's pump.  That's because our thread
	//    might have queued messages that would then be misrouted or lost because the dialog's
	//    dispatched them incorrectly, or didn't know what to do with them because they had a
	//    NULL hwnd.
	// 2) If the script happens to be uninterrutible, we would want to "re-queue" the messages
	//    in this fashion anyway (doing so avoids possible conflict if the current quasi-thread
	//    is in the middle of a critical operation, such as trying to open the clipboard for 
	//    a command it is right in the middle of executing).
	// 3) "Re-queuing" this event *only* for case #2 might cause problems with losing the original
	//     sequence of events that occurred in the GUI.  For example, if some events were re-queued
	//     due to uninterruptibility, but other more recent ones were not (because the thread became
	//     interruptible again at a critical moment), the more recent events would take effect before
	//     the older ones!  Requeuing all events ensures that when they do take effect, they do so
	//     in their original order.
	//
	// More explanation about Case #1 above.  Consider this order of events: 1) Current thread is
	// waiting for dialog, thus that dialog's msg pump is running. 2) User presses a button on GUI
	// window, and then another while the prev. button's thread is still uninterruptible (perhaps
	// this happened due to automating the form with the Send command).  3) Due to uninterruptibility,
	// this event would be re-queued to the thread's msg queue.  4) If the first thread ends before any
	// call to MsgSleep(), we'll be back at the dialog's msg pump again, and thus the requeued message
	// would be misrouted or discarded by its automatic dispatching.
	//
	// Info about why events are buffered when script is uninterruptible:
 	// It seems best to buffer GUI events that would trigger a new thread, rather than ignoring
	// them or allowing them to launch unconditionally.  Ignoring them is bad because lost events
	// might cause a GUI to get out of sync with how its controls are designed to update and
	// interact with each other.  Allowing them to launch anyway is bad in the case where
	// a critical operation for another thread is underway (such as attempting to open the
	// clipboard).  Thus, post this event back to our thread so that even if a msg pump
	// other than our own is running (such as MsgBox or another dialog), the msg will stay
	// buffered (the msg might bounce around fiercely if we kept posting it to our WindowProc,
	// since any other msg pump would not have the buffering filter in place that ours has).

	POST_AHK_GUI_ACTION(mHwnd, (WPARAM)aControlIndex, (LPARAM)gui_event);
	MsgSleep(-1);
	// The above MsgSleep() is for the case when there is a dialog message pump nearer on the
	// call stack than an instance of MsgSleep(), in which case the message would be dispatched
	// to our GUI Window Proc, which doesn't know what to do with it and would just discard it.
	// If the script happens to be uninterruptible, that shouldn't be a problem because it implies
	// that there is an instance of MsgSleep() nearer on the call stack than any dialog's message
	// pump.  Search for "call stack" to find more comments about this.
}
