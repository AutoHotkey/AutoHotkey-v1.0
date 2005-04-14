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
#include "script.h"
#include "globaldata.h" // for a lot of things
#include "application.h" // for MsgSleep()
#include "window.h" // for SetForegroundWindowEx()


ResultType Script::PerformGui(char *aCommand, char *aParam2, char *aParam3, char *aParam4)
{
	int window_index = g.GuiDefaultWindowIndex; // Which window to operate upon.  Initialized to thread's default.
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
	switch(gui_command)
	{
	case GUI_CMD_DESTROY:
		return GuiType::Destroy(window_index);

	case GUI_CMD_DEFAULT:
		// Change the "default" member, not g.GuiWindowIndex because that contains the original
		// window number reponsible for launching this thread, which should not be changed because it is
		// used to produce the contents of A_Gui.  Also, it's okay if the specify window index doesn't
		// currently exist.
		g.GuiDefaultWindowIndex = window_index;
		return OK;
	}


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
		case GUI_CMD_FLASH:
		case GUI_CMD_MINIMIZE:
		case GUI_CMD_MAXIMIZE:
		case GUI_CMD_RESTORE:
			return OK; // Nothing needs to be done since the window object doesn't exist.
		}
		// Otherwise: Create the object and (later) its window, since all the other sub-commands below need it:
		if (   !(g_gui[window_index] = new GuiType(window_index))   )
			return FAIL; // No error displayed since extremely rare.
		if (   !(g_gui[window_index]->mControl = (GuiControlType *)malloc(GUI_CONTROL_BLOCK_SIZE * sizeof(GuiControlType)))   )
		{
			delete g_gui[window_index];
			g_gui[window_index] = NULL;
			return FAIL; // No error displayed since extremely rare.
		}
		g_gui[window_index]->mControlCapacity = GUI_CONTROL_BLOCK_SIZE;
		// Probably better to increment here rather than in constructor in case GuiType objects
		// are ever created outside of the g_gui array (such as for temp local variables):
		++GuiType::sObjectCount; // This count is maintained to help performance in the main event loop and other places.
	}

	GuiType &gui = *g_gui[window_index];  // For performance and convenience.

	// Now handle any commands that should be handled prior to creation of the window in the case
	// where the window doesn't already exist:
	bool set_last_found_window = false;
	if (gui_command == GUI_CMD_OPTIONS)
		if (!gui.ParseOptions(options, set_last_found_window))
			return FAIL;  // It already displayed the error.

	// Create the window if needed.  Since it should not be possible for our window to get destroyed
	// without our knowning about it (via the explicit handling in its window proc), it shouldn't
	// be necessary to check the result of IsWindow(gui.mHwnd):
	if (!gui.mHwnd && !gui.Create())
	{
		GuiType::Destroy(window_index); // Get rid of the object so that it stays in sync with the window's existence.
		return ScriptError("Could not create window." ERR_ABORT);
	}

	if (set_last_found_window)
		g.hWndLastUsed = gui.mHwnd;

	// After creating the window, return from any commands that were fully handled above:
	if (gui_command == GUI_CMD_OPTIONS)
		return OK;

	GuiControls gui_control_type = GUI_CONTROL_INVALID;
	int index;

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

	case GUI_CMD_MINIMIZE:
		// If the window is hidden, it is unhidden as a side-effect (this happens even for SW_SHOWMINNOACTIVE).
		ShowWindow(gui.mHwnd, SW_MINIMIZE);
		return OK;

	case GUI_CMD_MAXIMIZE:
		ShowWindow(gui.mHwnd, SW_MAXIMIZE); // If the window is hidden, it is unhidden as a side-effect.
		return OK;

	case GUI_CMD_RESTORE:
		ShowWindow(gui.mHwnd, SW_RESTORE); // If the window is hidden, it is unhidden as a side-effect.
		return OK;

	case GUI_CMD_FONT:
		return gui.SetCurrentFont(aParam2, aParam3);

	case GUI_CMD_TAB:
		if (!*aParam2 && !*aParam3) // Both the tab control number and the tab number were omitted.
			gui.mCurrentTabControlIndex = MAX_TAB_CONTROLS; // i.e. "no tab"
		else
		{
			if (*aParam3) // Which tab control. Must be processed prior to Param2 since it might change mCurrentTabControlIndex.
			{
				index = ATOI(aParam3) - 1;
				if (index < 0 || index > MAX_TAB_CONTROLS - 1)
					return ScriptError("Paramter #3 is out of bounds." ERR_ABORT, aParam2);
				gui.mCurrentTabControlIndex = index;
			}
			if (*aParam2) // Index of a particular tab inside a control.
			{
				// Unlike "GuiControl, Choose", in this case, don't allow negatives since that would just
				// generate an error msg further below:
				if (IsPureNumeric(aParam2, false, false))
				{
					index = ATOI(aParam2) - 1;
					if (index < 0 || index > MAX_TABS_PER_CONTROL - 1)
						return ScriptError("Paramter #2 is out of bounds." ERR_ABORT, aParam2);
				}
				else
				{
					index = -1;  // Set default to be "failure".
					GuiControlType *tab_control = gui.FindTabControl(gui.mCurrentTabControlIndex);
					if (tab_control)
						index = gui.FindTabIndexByName(*tab_control, aParam2); // Returns -1 on failure.
					if (index == -1)
						return ScriptError("This tab name doesn't exist yet." ERR_ABORT, aParam2);
				}
				gui.mCurrentTabIndex = index;
				if (!*aParam3 && gui.mCurrentTabControlIndex == MAX_TAB_CONTROLS)
					// Provide a default: the most recently added tab control.  If there are no
					// tab controls, assume the index is the first tab control (i.e. a tab control
					// to be created in the future):
					gui.mCurrentTabControlIndex = gui.mTabControlCount ? gui.mTabControlCount - 1 : 0;
			}
		}
		return OK;
		
	case GUI_CMD_COLOR:
		// AssignColor() takes care of deleting old brush, etc.
		// In this case, a blank for either param means "leaving existing color alone", in which
		// case AssignColor() is not called since it would assume CLR_NONE then.
		if (*aParam2)
			AssignColor(aParam2, gui.mBackgroundColorWin, gui.mBackgroundBrushWin);
		if (*aParam3)
			AssignColor(aParam3, gui.mBackgroundColorCtl, gui.mBackgroundBrushCtl);
		if (IsWindowVisible(gui.mHwnd))
			// Force the window to repaint so that colors take effect immediately.
			// UpdateWindow() isn't enough sometimes/always, so do something more aggressive:
			InvalidateRect(gui.mHwnd, NULL, TRUE);
		return OK;

	case GUI_CMD_FLASH:
		// Note that FlashWindowEx() would have to be loaded dynamically since it is not available
		// on Win9x/NT.  But for now, just this simple method is provided.  In the future, more
		// sophisticated paramters can be made available to flash the window a given number of times
		// and at a certain frequency, and other options such as only-taskbar-button or only-caption.
		// Set FlashWindowEx() for more ideas:
		FlashWindow(gui.mHwnd, stricmp(aParam2, "Off") ? TRUE : FALSE);
		return OK;

	} // switch()

	return FAIL;  // Should never be reached, but avoids compiler warning and improves bug detection.
}



ResultType Line::GuiControl(char *aCommand, char *aControlID, char *aParam3)
{
	char *options; // This will contain something that is meaningful only when gui_command == GUICONTROL_CMD_OPTIONS.
	int window_index = g.GuiDefaultWindowIndex; // Which window to operate upon.  Initialized to thread's default.
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
	GuiIndexType control_index = gui.FindControl(aControlID);
	if (control_index >= gui.mControlCount) // Not found.
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	GuiControlType &control = gui.mControl[control_index];   // For performance and convenience.

	// Beyond this point, errors are rare so set the default to "no error":
	g_ErrorLevel->Assign(ERRORLEVEL_NONE);

	char *malloc_buf;
	RECT rect;
	WPARAM checked;
	GuiControlType *tab_control;

	switch(guicontrol_cmd)
	{

	case GUICONTROL_CMD_OPTIONS:
	{
		GuiControlOptionsType go; // Its contents not currently used here, but it might be in the future.
		gui.ControlInitOptions(go, control);
		return gui.ControlParseOptions(options, go, control, control_index);
	}

	case GUICONTROL_CMD_CONTENTS:
	case GUICONTROL_CMD_TEXT:
		switch (control.type)
		{
		case GUI_CONTROL_EDIT:
			// Note that TranslateLFtoCRLF() will return the original buffer we gave it if no translation
			// is needed.  Otherwise, it will return a new buffer which we are responsible for freeing
			// when done (or NULL if it failed to allocate the memory).
			malloc_buf = (*aParam3 && (GetWindowLong(control.hwnd, GWL_STYLE) & ES_MULTILINE))
				? TranslateLFtoCRLF(aParam3) : aParam3; // Automatic translation, as documented.
			SetWindowText(control.hwnd,  malloc_buf ? malloc_buf : aParam3); // malloc_buf is checked again in case the mem alloc failed.
			if (malloc_buf && malloc_buf != aParam3)
				free(malloc_buf);
			return OK;

		case GUI_CONTROL_PIC:
		{
			// Update: The below doesn't work, so it will be documented that a picture control
			// should be always be referred to by its original filename even if the picture changes.
			// Set the text unconditionally even if the picture can't be loaded.  This text must
			// be set to allow GuiControl(Get) to be able to operate upon the picture without
			// needing to indentify it via something like "Static14".
			//SetWindowText(control.hwnd, aParam3);
			//SendMessage(control.hwnd, WM_SETTEXT, 0, (LPARAM)aParam3);

			// Set default options, to be possibly overridden by any options actually present:
			// Fixed for v1.0.23: Below should use GetClientRect() vs. GetWindowRect(), otherwise
			// a size too large will be returned if the control has a border:
			GetClientRect(control.hwnd, &rect);
			int width = rect.right - rect.left;
			int height = rect.bottom - rect.top;
			int icon_index = -1;  // Tell it to use the default, which avoids the ExtractIcon() method.

			// The below must be done only after the above, because setting the control's picture handle
			// to NULL sometimes or always shrinks the control down to zero dimensions:
			// Although all HBITMAPs are freed upon program termination, if the program changes
			// the picture frequently, memory/resources would continue to rise in the meantime
			// unless this is done:
			if (control.union_hbitmap)
				if (control.attrib & GUI_CONTROL_ATTRIB_ALTBEHAVIOR) // union_hbitmap is an icon or cursor.
				{
					// The control's image is set to NULL is done for the following reasons:
					// 1) It turns off the control's animation timer in case the new image is not animated.
					// 2) It feels a little bit safter to destroy the image only after it has been removed
					//    from the control.
					// NOTE: IMAGE_ICON or IMAGE_CURSOR must be passed, not IMAGE_BITMAP.  Otherwise the
					// animated property of the control (via a timer that the control created) will remain
					// in effect for the next image, even if it isn't animated, which results in a
					// flashing/redrawing effect:
					SendMessage(control.hwnd, STM_SETIMAGE, IMAGE_CURSOR, NULL);
					DestroyIcon((HICON)control.union_hbitmap); // Works on cursors too.  See notes in LoadPicture().
				}
				else // union_hbitmap is a bitmap
				{
					SendMessage(control.hwnd, STM_SETIMAGE, IMAGE_BITMAP, NULL);
					DeleteObject(control.union_hbitmap);
				}

			// Parse any options that are present in front of the filename:
			char *next_option = omit_leading_whitespace(aParam3);
			if (*next_option == '*') // Options are present.  Must check this here and in the for-loop to avoid omitting legitimate whitespace in a filename that starts with spaces.
			{
				char *option_end, orig_char;
				for (; *next_option == '*'; next_option = omit_leading_whitespace(option_end))
				{
					// Find the end of this option item:
					if (   !(option_end = StrChrAny(next_option, " \t"))   )  // Space or tab.
						option_end = next_option + strlen(next_option); // Set to position of zero terminator instead.
					// Permanently terminate in between options to help eliminate ambiguity for words contained
					// inside other words, and increase confidence in decimal and hexadecimal conversion.
					orig_char = *option_end;
					*option_end = '\0';
					++next_option; // Skip over the asterisk.  It might point to a zero terminator now.
					if (!strnicmp(next_option, "Icon", 4))
						icon_index = ATOI(next_option + 4) - 1; // LoadPicture() correctly handles any negative value.
					else
					{
						switch (toupper(*next_option))
						{
						case 'W':
							width = ATOI(next_option + 1);
							break;
						case 'H':
							height = ATOI(next_option + 1);
							break;
						// If not one of the above, such as zero terminator or a number, just ignore it.
						}
					}

					*option_end = orig_char; // Undo the temporary termination so that loop's omit_leading() will work.
				} // for() each item in option list

				// The below assigns option_end + 1 vs. next_option in case the filename is contained in a
				// variable ref and/ that filename contains leading spaces.  Example:
				// GuiControl,, MyPic, *w100 *h-1 %FilenameWithLeadingSpaces%
				// Update: Windows XP and perhaps other OSes will load filenames-containing-leading-spaces
				// even if those spaces are omitted.  However, I'm not sure whether all API calls that
				// use filenames do this, so it seems best to include those spaces wheneve possible.
				aParam3 = *option_end ? option_end + 1 : option_end; // Set aParam3 to the start of the image's filespec.
			} 
			//else options are not present, so do not set aParam3 to be next_option because that would
			// omit legitimate spaces and tabs that might exist at the beginning of a real filename (file
			// names can start with spaces).

			// See comments in AddControl():
			int image_type;
			if (   !(control.union_hbitmap = LoadPicture(aParam3, width, height, image_type, icon_index
				, control.attrib & GUI_CONTROL_ATTRIB_ALTSUBMIT))   )
				return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
			DWORD style = GetWindowLong(control.hwnd, GWL_STYLE);
			DWORD style_image_type = style & 0x0F;
			style &= ~0x0F;  // Purge the low-order four bits in case style-image-type needs to be altered below.
			if (image_type == IMAGE_BITMAP)
			{
				if (style_image_type != SS_BITMAP)
					SetWindowLong(control.hwnd, GWL_STYLE, style | SS_BITMAP);
			}
			else // Icon or Cursor.
				if (style_image_type != SS_ICON) // Must apply SS_ICON or such handles cannot be displayed.
					SetWindowLong(control.hwnd, GWL_STYLE, style | SS_ICON);
			// LoadPicture() uses CopyImage() to scale the image, which seems to provide better scaling
			// quality than using MoveWindow() (followed by redrawing the parent window) on the static
			// control that contains the image.
			SendMessage(control.hwnd, STM_SETIMAGE, image_type, (LPARAM)control.union_hbitmap);
			if (image_type == IMAGE_BITMAP)
				control.attrib &= ~GUI_CONTROL_ATTRIB_ALTBEHAVIOR;  // Flag it as a bitmap so that DeleteObject vs. DestroyIcon will be called for it.
			else // Cursor or Icon, which are functionally identical.
				control.attrib |= GUI_CONTROL_ATTRIB_ALTBEHAVIOR;
			return OK;
		}

		case GUI_CONTROL_CHECKBOX:
		case GUI_CONTROL_RADIO:
			if (guicontrol_cmd == GUICONTROL_CMD_CONTENTS && IsPureNumeric(aParam3, true, false))
			{
				checked = ATOI(aParam3);
				if (!checked || checked == 1 || (control.type == GUI_CONTROL_CHECKBOX && checked == -1))
				{
					if (checked == -1)
						checked = BST_INDETERMINATE;
					//else the "checked" var is already set correctly.
					if (control.type == GUI_CONTROL_RADIO)
					{
						gui.ControlCheckRadioButton(control, control_index, checked);
						return OK;
					}
					// Otherwise, we're operating upon a checkbox.
					SendMessage(control.hwnd, BM_SETCHECK, checked, 0);
					return OK;
				}
				//else the default SetWindowText() action will be taken below.
			}
			// else assume it's the text/caption for the item, so the default SetWindowText() action will be taken below.
			break;

		case GUI_CONTROL_HOTKEY:
			SendMessage(control.hwnd, HKM_SETHOTKEY, gui.TextToHotkey(aParam3), 0); // This will set it to "None" if aParam3 is blank.
			break;
		
		case GUI_CONTROL_SLIDER:
			// Confirmed this fact from MSDN: That the control automatically deals with out-of-range values
			// by setting slider to min or max:
			if (*aParam3 == '+') // Apply as delta from its current position.
			{
				int slider_value = ATOI(aParam3 + 1);
				if (control.attrib & GUI_CONTROL_ATTRIB_ALTBEHAVIOR)
					slider_value = -slider_value;  // Delta moves to opposite direction if control is inverted.
				SendMessage(control.hwnd, TBM_SETPOS, TRUE
					, SendMessage(control.hwnd, TBM_GETPOS, 0, 0) + slider_value);
				// Above uses +1 to omit the plus sign, which allows a negative delta via +-5.
				// -5 is not treated as a delta because that would be ambigous with an absolute position.
				// In any case, it seems like too much code to be justified.
			}
			else
				SendMessage(control.hwnd, TBM_SETPOS, TRUE, gui.ControlInvertSliderIfNeeded(control, ATOI(aParam3)));
				// Above msg has no return value.
			return OK;

		case GUI_CONTROL_PROGRESS:
			// Confirmed through testing (PBM_DELTAPOS was also tested): The control automatically deals
			// with out-of-range values by setting bar to min or max.  
			if (*aParam3 == '+')
				// This allows a negative delta, e.g. via +-5.  Nothing fancier is done since the need
				// to go backwards in a progress bar is rare.
				SendMessage(control.hwnd, PBM_DELTAPOS, ATOI(aParam3 + 1), 0);
			else
				SendMessage(control.hwnd, PBM_SETPOS, ATOI(aParam3), 0);
			return OK;

		case GUI_CONTROL_DROPDOWNLIST:
		case GUI_CONTROL_COMBOBOX:
		case GUI_CONTROL_LISTBOX:
		case GUI_CONTROL_TAB:
			// Seems best not to do the below due to the extreme rarity of anyone wanting to change a
			// ListBox or ComboBox's hidden caption.  That can be done via ControlSetText if it is
			// ever needed.  The advantage of not doing this is that the "TEXT" command can be used
			// as a gentle, slight-variant of GUICONTROL_CMDCONTENTS, i.e. without needing to worry
			// what the target control's type is:
			//if (guicontrol_cmd == GUICONTROL_CMD_TEXT)
			//	break;
			bool list_replaced;
			if (*aParam3 == '|') // The signal to overwrite rather than append to the list.
			{
				list_replaced = true;
				++aParam3;  // Exclude the initial pipe from further consideration.
				if (control.type == GUI_CONTROL_TAB)
					TabCtrl_DeleteAllItems(control.hwnd);
				else
					SendMessage(control.hwnd, (control.type == GUI_CONTROL_LISTBOX)
						? LB_RESETCONTENT : CB_RESETCONTENT, 0, 0); // Delete all items currently in the list.
			}
			else
				list_replaced = false;
			gui.ControlAddContents(control, aParam3, 0);
			if (control.type == GUI_CONTROL_TAB && list_replaced)
			{

				// In case replacement tabs deleted the currently active tab, update the tab.
				// The "false" param will cause focus to jump to first item in z-order if
				// a the control that previously had focus was inside a tab that was just
				// deleted (seems okay since this kind of operation is fairly rare):
				gui.ControlUpdateCurrentTab(control, false);
				// Must invalidate part of parent window to get controls to redraw correctly, at least
				// in the following case: Tab that is currently active still exists and is still active
				// after the tab-rebuild done above.
				// For simplicitly, invalidate the whole thing since changing the quantity/names of tabs
				// while the window is visible is rare.  NOTE: It might be necessary to invalidate
				// the entire window *anyway* in case some of this tab's controls exist outside its
				// boundaries (e.g. TCS_BUTTONS).  Another reason is the fact that there have been
				// problems retrieving an accurate client area for tab controls when they have certain
				// styles such as TCS_VERTICAL:
				InvalidateRect(gui.mHwnd, NULL, TRUE); // TRUE = Seems safer to erase, not knowing all possible overlaps.
			}
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

		GetWindowRect(control.hwnd, &rect); // Failure seems too rare to check for.
		POINT dest_pt = {rect.left, rect.top};
		ScreenToClient(gui.mHwnd, &dest_pt); // Set default x/y target position, to be possibly overridden below.
		if (xpos != COORD_UNSPECIFIED)
			dest_pt.x = xpos;
		if (ypos != COORD_UNSPECIFIED)
			dest_pt.y = ypos;

		if (!MoveWindow(control.hwnd, dest_pt.x, dest_pt.y
			, width == COORD_UNSPECIFIED ? rect.right - rect.left : width
			, height == COORD_UNSPECIFIED ? rect.bottom - rect.top : height
			, TRUE))  // Do repaint.
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);

		if (control.type == GUI_CONTROL_SLIDER) // It seems buddies don't move automatically, so trigger the move.
		{
			HWND buddy1 = (HWND)SendMessage(control.hwnd, TBM_GETBUDDY, TRUE, 0);
			HWND buddy2 = (HWND)SendMessage(control.hwnd, TBM_GETBUDDY, FALSE, 0);
			if (buddy1)
			{
				SendMessage(control.hwnd, TBM_SETBUDDY, TRUE, (LPARAM)buddy1);
				// It doesn't always redraw the buddies correctly, at least on XP, so do it manually:
				InvalidateRect(buddy1, NULL, TRUE);
			}
			if (buddy2)
			{
				SendMessage(control.hwnd, TBM_SETBUDDY, FALSE, (LPARAM)buddy2);
				InvalidateRect(buddy2, NULL, TRUE);
			}
		}

		// This must be done, at least in cases such as GroupBox under certain themes/conditions.
		// More than just control.hwnd must be invalided, otherwise the interior of the GroupBox retains
		// a ghost image of whatever was in it before the move:
		GetWindowRect(control.hwnd, &rect); // Limit it to only that part of the client area that is receiving the rect.
		MapWindowPoints(NULL, gui.mHwnd, (LPPOINT)&rect, 2); // Convert rect to client coordinates (not the same as GetClientRect()).
		InvalidateRect(gui.mHwnd, &rect, TRUE); // Seems safer to use TRUE, not knowing all possible overlaps, etc.
		return OK;
	}

	case GUICONTROL_CMD_FOCUS:
		return SetFocus(control.hwnd) ? OK : g_ErrorLevel->Assign(ERRORLEVEL_ERROR);

	case GUICONTROL_CMD_ENABLE:
	case GUICONTROL_CMD_DISABLE:
		// GUI_CONTROL_ATTRIB_EXPLICITLY_DISABLED is maintained for use with tab controls.  It allows controls
		// on inactive tabs to be marked for later enabling.  It also allows explicitly disabled controls to
		// stay disabled even when their tab/page becomes active. It is updated unconditionally for simplicity
		// and maintainability.  
		if (guicontrol_cmd == GUICONTROL_CMD_ENABLE)
			control.attrib &= ~GUI_CONTROL_ATTRIB_EXPLICITLY_DISABLED;
		else
			control.attrib |= GUI_CONTROL_ATTRIB_EXPLICITLY_DISABLED;
		if ((tab_control = gui.FindTabControl(control.tab_control_index)) // It belongs to a tab control that already exists.
			&& ((GetWindowLong(tab_control->hwnd, GWL_STYLE) & WS_DISABLED) // But its tab control is disabled...
				|| TabCtrl_GetCurSel(tab_control->hwnd) != control.tab_index))
				// ... or either there is no current tab/page (or there are no tabs at all) or the one selected
				// is not this control's: Do not disable or re-enable the control in this case.
			return OK;
		// Since above didn't return, act upon the enabled/disable:
		EnableWindow(control.hwnd, guicontrol_cmd == GUICONTROL_CMD_ENABLE ? TRUE : FALSE);
		if (control.type == GUI_CONTROL_TAB) // This control is a tab control.
			// Update the control so that its current tab's controls will all be enabled or disabled (now
			// that the tab control itself has just been enabled or disabled):
			gui.ControlUpdateCurrentTab(control, false);
		return OK;

	case GUICONTROL_CMD_SHOW:
	case GUICONTROL_CMD_HIDE:
		// GUI_CONTROL_ATTRIB_EXPLICITLY_HIDDEN is maintained for use with tab controls.  It allows controls
		// on inactive tabs to be marked for later showing.  It also allows explicitly hidden controls to
		// stay hidden even when their tab/page becomes active. It is updated unconditionally for simplicity
		// and maintainability.  
		if (guicontrol_cmd == GUICONTROL_CMD_SHOW)
			control.attrib &= ~GUI_CONTROL_ATTRIB_EXPLICITLY_HIDDEN;
		else
			control.attrib |= GUI_CONTROL_ATTRIB_EXPLICITLY_HIDDEN;
		if ((tab_control = gui.FindTabControl(control.tab_control_index)) // It belongs to a tab control that already exists.
			&& (!(GetWindowLong(tab_control->hwnd, GWL_STYLE) & WS_VISIBLE) // But its tab control is hidden...
				|| TabCtrl_GetCurSel(tab_control->hwnd) != control.tab_index))
				// ... or either there is no current tab/page (or there are no tabs at all) or the one selected
				// is not this control's: Do not show or hide the control in this case.
			return OK;
		// Since above didn't return, act upon the show/hide:
		ShowWindow(control.hwnd, guicontrol_cmd == GUICONTROL_CMD_SHOW ? SW_SHOWNOACTIVATE : SW_HIDE);
		if (control.type == GUI_CONTROL_TAB) // This control is a tab control.
			// Update the control so that its current tab's controls will all be shown or hidden (now
			// that the tab control itself has just been shown or hidden):
			gui.ControlUpdateCurrentTab(control, false);
		return OK;

	case GUICONTROL_CMD_CHOOSE:
	case GUICONTROL_CMD_CHOOSESTRING:
	{
		int selection_index;
		int extra_actions = 0; // Set default.
		if (*aParam3 == '|') // First extra action.
		{
			++aParam3; // Omit this pipe char from further consideration below.
			++extra_actions;
		}
		if (control.type == GUI_CONTROL_TAB)
		{
			// Generating the TCN_SELCHANGING and TCN_SELCHANGE messages manually is fairly complex since they
			// need a struct and who knows whether it's even valid for sources other than the tab controls
			// themselves to generate them.  I would use TabCtrl_SetCurFocus(), but that is shot down by
			// the fact that it only generates TCN_SELCHANGING and TCN_SELCHANGE if the tab control lacks
			// the TCS_BUTTONS style, which would make it an incomplete/inconsistent solution.  But I guess
			// it's better than nothing as long as it's documented.
			// MSDN: "If the tab control does not have the TCS_BUTTONS style, changing the focus also changes
			// selected tab. In this case, the tab control sends the TCN_SELCHANGING and TCN_SELCHANGE
			// notification messages to its parent window. 
			// Automatically switch to CHOOSESTRING if parameter isn't numeric:
			if (guicontrol_cmd == GUICONTROL_CMD_CHOOSE && !IsPureNumeric(aParam3, true, false))
				guicontrol_cmd = GUICONTROL_CMD_CHOOSESTRING;
			if (guicontrol_cmd == GUICONTROL_CMD_CHOOSESTRING)
				selection_index = gui.FindTabIndexByName(control, aParam3); // Returns -1 on failure.
			else
				selection_index = ATOI(aParam3) - 1;
			if (selection_index < 0 || selection_index > MAX_TABS_PER_CONTROL - 1)
				return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
			int previous_selection_index = TabCtrl_GetCurSel(control.hwnd);
			if (!extra_actions || (GetWindowLong(control.hwnd, GWL_STYLE) & TCS_BUTTONS))
			{
				if (TabCtrl_SetCurSel(control.hwnd, selection_index) == -1)
					return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
				// In this case but not the "else" below, must update the tab to show the proper controls:
				if (previous_selection_index != selection_index)
					gui.ControlUpdateCurrentTab(control, extra_actions > 0); // And set focus if the more forceful extra_actions was done.
			}
			else // There is an extra_action and it's not TCS_BUTTONS, so extra_action is possible via TabCtrl_SetCurFocus.
			{
				TabCtrl_SetCurFocus(control.hwnd, selection_index); // No return value, so check for success below.
				if (TabCtrl_GetCurSel(control.hwnd) != selection_index)
					return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
			}
			return OK;
		}
		// Otherwise, it's not a tab control, but a ListBox/DropDownList/Combo or other control:
		if (*aParam3 == '|' && control.type != GUI_CONTROL_TAB) // Second extra action.
		{
			++aParam3; // Omit this pipe char from further consideration below.
			++extra_actions;
		}
		if (guicontrol_cmd == GUICONTROL_CMD_CHOOSE && !IsPureNumeric(aParam3, true, false))
			guicontrol_cmd = GUICONTROL_CMD_CHOOSESTRING;
		UINT msg, x_msg, y_msg;
		switch(control.type)
		{
		case GUI_CONTROL_TAB:
			break; // i.e. don't do the "default" section.
		case GUI_CONTROL_DROPDOWNLIST:
		case GUI_CONTROL_COMBOBOX:
			msg = (guicontrol_cmd == GUICONTROL_CMD_CHOOSE) ? CB_SETCURSEL : CB_SELECTSTRING;
			x_msg = CBN_SELCHANGE;
			y_msg = CBN_SELENDOK;
			break;
		case GUI_CONTROL_LISTBOX:
			if (GetWindowLong(control.hwnd, GWL_STYLE) & (LBS_EXTENDEDSEL|LBS_MULTIPLESEL))
			{
				if (guicontrol_cmd == GUICONTROL_CMD_CHOOSE)
					msg = LB_SETSEL;
				else
					// MSDN: Do not use [LB_SELECTSTRING] with a list box that has the LBS_MULTIPLESEL or the
					// LBS_EXTENDEDSEL styles:
					msg = LB_FINDSTRING;
			}
			else // single-select listbox
				if (guicontrol_cmd == GUICONTROL_CMD_CHOOSE)
					msg = LB_SETCURSEL;
				else
					msg = LB_SELECTSTRING;
			x_msg = LBN_SELCHANGE;
			y_msg = LBN_DBLCLK;
			break;
		default:  // Not a supported control type.
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		} // switch(control.type)

		if (guicontrol_cmd == GUICONTROL_CMD_CHOOSESTRING)
		{
			if (msg == LB_FINDSTRING)
			{
				// This msg is needed for multi-select listbox because LB_SELECTSTRING is not supported
				// in this case.
				LRESULT found_item = SendMessage(control.hwnd, msg, -1, (LPARAM)aParam3);
				if (found_item == CB_ERR) // CB_ERR == LB_ERR
					return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
				if (SendMessage(control.hwnd, LB_SETSEL, TRUE, found_item) == CB_ERR) // CB_ERR == LB_ERR
					return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
			}
			else
				if (SendMessage(control.hwnd, msg, 1, (LPARAM)aParam3) == CB_ERR) // CB_ERR == LB_ERR
					return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		}
		else // Choose by position vs. string.
		{
			selection_index = ATOI(aParam3) - 1;
			if (selection_index < 0)
				return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
			if (msg == LB_SETSEL) // Multi-select, so use the cumulative method.
			{
				if (SendMessage(control.hwnd, msg, TRUE, selection_index) == CB_ERR) // CB_ERR == LB_ERR
					return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
			}
			else
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

	case GUICONTROL_CMD_FONT:
		// Done in regardless of USES_FONT_AND_TEXT_COLOR to allow future OSes or common control updates
		// to be given an explicit font, even though it would have no effect currently:
		SendMessage(control.hwnd, WM_SETFONT, (WPARAM)gui.sFont[gui.mCurrentFontIndex].hfont, 0);
		if (USES_FONT_AND_TEXT_COLOR(control.type)) // Must check this to avoid trashing union_hbitmap.
			control.union_color = gui.mCurrentColor;
		InvalidateRect(control.hwnd, NULL, TRUE); // Required for refresh, at least for edit controls, probably some others.
		return OK;

	} // switch()

	return FAIL;  // Should never be reached, but avoids compiler warning and improves bug detection.
}



ResultType Line::GuiControlGet(char *aCommand, char *aControlID, char *aParam3)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL; // ErrorLevel is not used in this case since it's an unexpected and critical error.
	output_var->Assign(); // Set default to be blank for all commands, for consistency.

	int window_index = g.GuiDefaultWindowIndex; // Which window to operate upon.  Initialized to thread's default.
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
		if (!cah.hwnd || !gui.FindControl(cah.hwnd)) // Relies on short-circuit boolean order.
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
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

	GuiIndexType control_index = gui.FindControl(aControlID);
	if (control_index >= gui.mControlCount) // Not found.
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	GuiControlType &control = gui.mControl[control_index];   // For performance and convenience.

	switch(guicontrolget_cmd)
	{
	case GUICONTROLGET_CMD_CONTENTS:
		// Because the below returns FAIL only if a critical error occurred, g_ErrorLevel is
		// left at NONE as set above for all cases.
		return gui.ControlGetContents(*output_var, control, aParam3);

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



ResultType GuiType::Destroy(GuiIndexType aWindowIndex)
// Rather than deal with the confusion of an object destroying itself, this method is static
// and designed to deal with one particular window index in the g_gui array.
{
	if (aWindowIndex >= MAX_GUI_WINDOWS)
		return FAIL;
	if (!g_gui[aWindowIndex]) // It's already in the right state.
		return OK;
	GuiType &gui = *g_gui[aWindowIndex];  // For performance and convenience.
	GuiIndexType u, object_count;
	if (gui.mHwnd)
	{
		// First destroy any windows owned by this window, since they will be auto-destroyed
		// anyway due to their being owned.  By destroying them explicitly, the Destroy()
		// function is called recursively which keeps everything relatively neat.
		for (u = 0, object_count = 0; u < MAX_GUI_WINDOWS; ++u)
		{
			if (g_gui[u])
			{
				if (g_gui[u]->mOwner == gui.mHwnd)
					GuiType::Destroy(u);
				if (sObjectCount == ++object_count) // No need to keep searching.
					break;
			}
		}
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
	if (gui.mHdrop)
		DragFinish(gui.mHdrop);
	// It seems best to delete the bitmaps whenever the control changes to a new image or
	// whenever the control is destroyed.  Otherwise, if a control or its parent window is
	// destroyed and recreated many times, memory allocation would continue to grow from
	// all the abandoned pointers:
	for (u = 0; u < gui.mControlCount; ++u)
	{
		if (gui.mControl[u].type == GUI_CONTROL_PIC && gui.mControl[u].union_hbitmap)
		{
			if (gui.mControl[u].attrib & GUI_CONTROL_ATTRIB_ALTBEHAVIOR)
				DestroyIcon((HICON)gui.mControl[u].union_hbitmap); // Works on cursors too.  See notes in LoadPicture().
			else // union_hbitmap is a bitmap rather than an icon or cursor.
				DeleteObject(gui.mControl[u].union_hbitmap);
			//else do nothing, since it isn't the right type to have a valid union_hbitmap member.
		}
	}
	// Not necessary since the object itself is about to be destroyed:
	//gui.mHwnd = NULL;
	//gui.mControlCount = 0; // All child windows (controls) are automatically destroyed with parent.
	free(gui.mControl); // Free the control array, which was previously malloc'd.
	delete g_gui[aWindowIndex]; // After this, the var "gui" is invalid so should not be referenced, i.e. the next line.
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
		sGuiInitialized = true;
	}

	char label_name[1024];  // Labels are unlimited in length, so use a size to cover anything realistic.

	// Find the label to run automatically when the form closes (if any):
	if (mWindowIndex > 0) // Prepend the window number for windows other than the first.
		snprintf(label_name, sizeof(label_name), "%dGuiClose", mWindowIndex + 1);
	else
		strcpy(label_name, "GuiClose");
	mLabelForClose = g_script.FindLabel(label_name);  // OK if NULL (closing the window is the same as "gui, cancel").

	// Find the label to run automatically when the user presses Escape (if any):
	if (mWindowIndex > 0) // Prepend the window number for windows other than the first.
		snprintf(label_name, sizeof(label_name), "%dGuiEscape", mWindowIndex + 1);
	else
		strcpy(label_name, "GuiEscape");
	mLabelForEscape = g_script.FindLabel(label_name);  // OK if NULL (pressing ESCAPE does nothing).

	// Find the label to run automatically when the user resizes the window (if any):
	if (mWindowIndex > 0) // Prepend the window number for windows other than the first.
		snprintf(label_name, sizeof(label_name), "%dGuiSize", mWindowIndex + 1);
	else
		strcpy(label_name, "GuiSize");
	mLabelForSize = g_script.FindLabel(label_name);  // OK if NULL.

	// Find the label to run automatically when files are dropped onto the window:
	if (mWindowIndex > 0) // Prepend the window number for windows other than the first.
		snprintf(label_name, sizeof(label_name), "%dGuiDropFiles", mWindowIndex + 1);
	else
		strcpy(label_name, "GuiDropFiles");
	if (mLabelForDropFiles = g_script.FindLabel(label_name))  // OK if NULL (dropping files).
		mExStyle |= WS_EX_ACCEPTFILES; // Makes the window accept drops. Otherwise, the WM_DROPFILES msg is not received.

	// The above is done prior to creating the window so that mLabelForDropFiles can determine
	// whether to add the WS_EX_ACCEPTFILES style.

	// WS_EX_APPWINDOW: "Forces a top-level window onto the taskbar when the window is minimized."
	// But it doesn't since the window is currently always unowned, there is not yet any need to use it.
	if (   !(mHwnd = CreateWindowEx(mExStyle, WINDOW_CLASS_GUI, g_script.mFileName, mStyle, 0, 0, 0, 0
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

	return OK;
}



void GuiType::UpdateMenuBars(HMENU aMenu)
// Caller has changed aMenu and wants the change visibly reflected in any windows that that
// use aMenu as a menu bar.  For example, if a menu item has been disabled, the grey-color
// won't show up immediately unless the window is refreshed.
{
	int i, object_count;
	for (i = 0, object_count = 0; i < MAX_GUI_WINDOWS; ++i)
	{
		if (g_gui[i])
		{
			if (g_gui[i]->mHwnd && GetMenu(g_gui[i]->mHwnd) == aMenu && IsWindowVisible(g_gui[i]->mHwnd))
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
			if (sObjectCount == ++object_count) // No need to keep searching.
				break;
		}
	}
}



ResultType GuiType::AddControl(GuiControls aControlType, char *aOptions, char *aText)
// Caller must have ensured that mHwnd is non-NULL (i.e. that the parent window already exists).
{
	#define TOO_MANY_CONTROLS "Too many controls." ERR_ABORT // Short msg since so rare.
	if (mControlCount >= MAX_CONTROLS_PER_GUI)
		return g_script.ScriptError(TOO_MANY_CONTROLS);
	if (mControlCount >= mControlCapacity) // The section below on the above check already having been done.
	{
		// realloc() to keep the array contiguous, which allows better-performing methods to be
		// used to access the list of controls in various places.
		// Expand the array by one block:
		GuiControlType *realloc_temp;  // Needed since realloc returns NULL on failure but leaves original block allocated!
		if (   !(realloc_temp = (GuiControlType *)realloc(mControl
			, (mControlCapacity + GUI_CONTROL_BLOCK_SIZE) * sizeof(GuiControlType)))   )
			return g_script.ScriptError(TOO_MANY_CONTROLS); // A non-specific msg since this error is so rare.
		mControl = realloc_temp;
		mControlCapacity += GUI_CONTROL_BLOCK_SIZE;
	}
	if (aControlType == GUI_CONTROL_TAB && mTabControlCount == MAX_TAB_CONTROLS)
		return g_script.ScriptError("Too many tab controls." ERR_ABORT); // Short msg since so rare.

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
	control.type = aControlType; // Improves maintainability to do this early.

	if (aControlType == GUI_CONTROL_TAB)
	{
		// For now, don't allow a tab control to be create inside another tab control because it raises
		// doubt and probably would create complications.  If it ever is allowed, note that
		// control.tab_index stores this tab control's index (0 for the first tab control, 1 for the
		// second, etc.) -- this is done for performance reasons.
		control.tab_control_index = MAX_TAB_CONTROLS;
		control.tab_index = mTabControlCount; // Store its control-index to help look-up performance in other sections.
	}
	else
	{
		control.tab_control_index = mCurrentTabControlIndex;
		control.tab_index = mCurrentTabIndex;
	}
	GuiControlOptionsType opt;
	ControlInitOptions(opt, control);
	// aOpt.checked is already okay since BST_UNCHECKED == 0
	// Similarly, the zero-init above also set the right values for password_char, new_section, etc.

	/////////////////////////////////////////////////
	// Set control-specific defaults for any options.
	/////////////////////////////////////////////////
	opt.style_add |= WS_VISIBLE;  // Starting default for all control types.

	// Radio buttons are handled separately here, outside the switch() further below:
	if (aControlType == GUI_CONTROL_RADIO)
	{
		// The BS_NOTIFY style is probably better not applied by default to radios because although it
		// causes the control to send BN_DBLCLK messages, each double-click by the user is seen only
		// as one click for the purpose of cosmetically making the button appear like it is being
		// clicked rapidly.  Update: the usefulness of double-clicking a radio button seems to
		// outweigh the rare cosmetic deficiency of rapidly clicking a radio button, so it seems
		// better to provide it as a default that can be overridden via explicit option.
		opt.style_add |= BS_MULTILINE|BS_NOTIFY;  // No WS_TABSTOP here since that is applied elsewhere depending on radio group nature.
		if (!mInRadioGroup)
			opt.style_add |= WS_GROUP; // Tabstop must be handled later below.
			// The mInRadioGroup flag will be changed accordingly after the control is successfully created.
		//else by default, no WS_TABSTOP or WS_GROUP.  However, WS_GROUP can be applied manually via the
		// options list to split this radio group off from one immediately prior to it.
	}
	else // Not a radio.
		if (mInRadioGroup) // Close out the prior radio group by giving this control the WS_GROUP style.
			opt.style_add |= WS_GROUP; // This might not be necessary on all OSes, but it seems traditional / best-practice.

	// Set control's default text color:
	if (USES_FONT_AND_TEXT_COLOR(aControlType))
		control.union_color = mCurrentColor; // Default to the most recently set color.
	else if (aControlType == GUI_CONTROL_PROGRESS) // This must be done to detect custom Progress color.
		control.union_color = CLR_DEFAULT; // Set progress to default color avoids unnecessary stripping of theme.
	//else don't change union_color since it shares the same address as union_hbitmap.

	switch (aControlType)
	{
	// Some controls also have the WS_EX_CLIENTEDGE exstyle by default because they look pretty strange
	// without them.  This seems to be the standard default used by most applications.
	// Note: It seems that WS_BORDER is hardly ever used in practice with controls, just parent windows.
	case GUI_CONTROL_GROUPBOX:
		opt.style_add |= BS_MULTILINE;
		break;
	case GUI_CONTROL_BUTTON:
	case GUI_CONTROL_CHECKBOX:
		opt.style_add |= WS_TABSTOP|BS_MULTILINE;
		break;
	case GUI_CONTROL_DROPDOWNLIST:
		opt.style_add |= WS_TABSTOP|WS_VSCROLL;  // CBS_DROPDOWNLIST is forcibly applied later. WS_VSCROLL is necessary.
		break;
	case GUI_CONTROL_COMBOBOX:
		// CBS_DROPDOWN is set as the default here to allow the flexibilty for it to be changed to
		// CBS_SIMPLE.  CBS_SIMPLE is allowed for ComboBox but not DropDownList because CBS_SIMPLE
		// has an edit control just like a combo, which DropDownList isn't equipped to handle via Submit().
		// Also, if CBS_AUTOHSCROLL is omitted, typed text cannot go beyond the visible width of the
		// edit control, so it seems best to havethat as a default also:
		opt.style_add |= WS_TABSTOP|WS_VSCROLL|CBS_AUTOHSCROLL|CBS_DROPDOWN;  // WS_VSCROLL is necessary.
		break;
	case GUI_CONTROL_LISTBOX:
		// Omit LBS_STANDARD because it includes LBS_SORT, which we don't want as a default style.
		opt.style_add |= WS_TABSTOP|WS_VSCROLL;  // WS_VSCROLL seems the most desirable default.
		opt.exstyle_add |= WS_EX_CLIENTEDGE;
		break;
	case GUI_CONTROL_EDIT:
		opt.style_add |= WS_TABSTOP;
		opt.exstyle_add |= WS_EX_CLIENTEDGE;
		break;
	case GUI_CONTROL_HOTKEY:
	case GUI_CONTROL_SLIDER:
		opt.style_add |= WS_TABSTOP;
		break;
	case GUI_CONTROL_PROGRESS:
		opt.style_add |= PBS_SMOOTH; // The smooth ones seem preferable as a default.  Theme is removed later below.
		break;
	case GUI_CONTROL_TAB:
		opt.style_add |= WS_TABSTOP|TCS_MULTILINE;
		break;
	// Nothing extra for these currently:
	//case GUI_CONTROL_RADIO: This one is handled separately above the switch().
	//case GUI_CONTROL_TEXT:
	//case GUI_CONTROL_PIC:
	}

	/////////////////////////////
	// Parse the list of options.
	/////////////////////////////
	if (!ControlParseOptions(aOptions, opt, control))
		return FAIL;  // It already displayed the error.
	DWORD style = opt.style_add & ~opt.style_remove;
	DWORD exstyle = opt.exstyle_add & ~opt.exstyle_remove;
	if (!mControlCount) // Always start new section for very first control, so override any false value from the above.
		opt.start_new_section = true;

	//////////////////////////////////////////
	// Force any mandatory styles into effect.
	//////////////////////////////////////////
	style |= WS_CHILD;  // All control types must have this, even if script attempted to remove it explicitly.
	switch (aControlType)
	{
	case GUI_CONTROL_GROUPBOX:
		// There doesn't seem to be any flexibility lost by forcing the buttons to be the right type,
		// and doing so improves maintainability and peace-of-mind:
		style = (style & ~BS_TYPEMASK) | BS_GROUPBOX;  // Force it to be the right type of button.
		break;
	case GUI_CONTROL_BUTTON:
		if (style & BS_DEFPUSHBUTTON)
			style = (style & ~BS_TYPEMASK) | BS_DEFPUSHBUTTON; // Done to ensure the lowest four bits are pure.
		else
			style &= ~BS_TYPEMASK;  // Force it to be the right type of button --> BS_PUSHBUTTON == 0
		break;
	case GUI_CONTROL_CHECKBOX:
		// Note: BS_AUTO3STATE and BS_AUTOCHECKBOX are mutually exclusive due to their overlap within
		// the bit field:
		if (style & BS_AUTO3STATE)
			style = (style & ~BS_TYPEMASK) | BS_AUTO3STATE; // Done to ensure the lowest four bits are pure.
		else
			style = (style & ~BS_TYPEMASK) | BS_AUTOCHECKBOX;  // Force it to be the right type of button.
		break;
	case GUI_CONTROL_RADIO:
		style = (style & ~BS_TYPEMASK) | BS_AUTORADIOBUTTON;  // Force it to be the right type of button.
		// This below must be handled here rather than in the set-defaults section because this
		// radio might be the first of its group due to the script having explicitly specified the word
		// Group in options (useful to make two adjacent radio groups).
		if (style & WS_GROUP && !(opt.style_remove & WS_TABSTOP))
			style |= WS_TABSTOP;
		// Otherwise it lacks a tabstop by default.
		break;
	case GUI_CONTROL_DROPDOWNLIST:
		style |= CBS_DROPDOWNLIST;  // This works because CBS_DROPDOWNLIST == CBS_SIMPLE|CBS_DROPDOWN
		break;
	case GUI_CONTROL_COMBOBOX:
		if (style & CBS_SIMPLE) // i.e. CBS_SIMPLE has been added to the original default, so assume it is SIMPLE.
			style = (style & ~0x0F) | CBS_SIMPLE; // Done to ensure the lowest four bits are pure.
		else
			style = (style & ~0x0F) | CBS_DROPDOWN; // Done to ensure the lowest four bits are pure.
		break;
	case GUI_CONTROL_LISTBOX:
		style |= LBS_NOTIFY;  // There doesn't seem to be any flexibility lost by forcing this style.
		break;
	case GUI_CONTROL_EDIT:
		// This is done for maintainability and peace-of-mind, though it might not strictly be required
		// to be done at this stage:
		if (opt.row_count > 1.5 || strchr(aText, '\n')) // Multiple rows or contents contain newline.
			style |= (ES_MULTILINE & ~opt.style_remove); // Add multiline unless it was explicitly removed.
		// This next check is relied upon by other things.  If this edit has the multiline style either
		// due to the above check or any other reason, provide other default styles if those styles
		// weren't explicitly removed in the options list:
		if (style & ES_MULTILINE) // If allowed, enable vertical scrollbar and capturing of ENTER keystrokes.
			// Safest to include ES_AUTOVSCROLL, though it appears to have no effect on XP.  See also notes below:
			#define EDIT_MULTILINE_DEFAULT (WS_VSCROLL|ES_WANTRETURN|ES_AUTOVSCROLL)
			style |= EDIT_MULTILINE_DEFAULT & ~opt.style_remove;
			// In addition, word-wrapping is implied unless explicitly disabled via -wrap in options.
			// This is because -wrap adds the ES_AUTOHSCROLL style.
		// else: Single-line edit.  ES_AUTOHSCROLL will be applied later below if all the other checks
		// fail to auto-detect this edit as a multi-line edit.
		// Notes: ES_MULTILINE is required for any CRLFs in the default value to display correctly.
		// If ES_MULTILINE is in effect: "If you do not specify ES_AUTOHSCROLL, the control automatically
		// wraps words to the beginning of the next line when necessary."
		// Also, ES_AUTOVSCROLL seems to have no additional effect, perhaps because this window type
		// is considered to be a dialog. MSDN: "When the multiline edit control is not in a dialog box
		// and the ES_AUTOVSCROLL style is specified, the edit control shows as many lines as possible
		// and scrolls vertically when the user presses the ENTER key. If you do not specify ES_AUTOVSCROLL,
		// the edit control shows as many lines as possible and beeps if the user presses the ENTER key when
		// no more lines can be displayed."
		break;
	case GUI_CONTROL_TAB:
		style |= WS_CLIPSIBLINGS; // MSDN: Both the parent window and the tab control must have the WS_CLIPSIBLINGS window style.
		// TCS_OWNERDRAWFIXED is required to implement custom Text color in the tabs.
		// For some reason, it's also required for TabWindowProc's WM_ERASEBACKGROUND to be able to
		// override the background color of the control's interior, at least when an XP theme is in effect.
		// (which is currently impossible since theme is always removed from a tab).
		// Even if that weren't the case, would still want owner-draw because otherwise the background
		// color of the tab-text would be different than the tab's interior, which testing shows looks
		// pretty strange.
		if (   (mBackgroundBrushWin && !(control.attrib & GUI_CONTROL_ATTRIB_BACKGROUND_DEFAULT))
			|| control.union_color != CLR_DEFAULT   )
			style |= TCS_OWNERDRAWFIXED;
		else
			style &= ~TCS_OWNERDRAWFIXED;
		break;

	// Nothing extra for these currently:
	//case GUI_CONTROL_TEXT:  Ensuring SS_BITMAP and such are absent seems too over-protective.
	//case GUI_CONTROL_PIC:   SS_BITMAP/SS_ICON are applied after the control isn't created so that it doesn't try to auto-load a resource.
	//case GUI_CONTROL_HOTKEY:
	//case GUI_CONTROL_SLIDER:
	//case GUI_CONTROL_PROGRESS:
	}

	////////////////////////////////////////////////////////////////////////////////////////////
	// If the above didn't already set a label for this control and this control type qualifies,
	// check if an automatic/implicit label exists for it in the script.
	////////////////////////////////////////////////////////////////////////////////////////////
	if (aControlType == GUI_CONTROL_BUTTON
		&& !control.jump_to_label && !(control.attrib & GUI_CONTROL_ATTRIB_IMPLICIT_CANCEL))
	{
		char label_name[1024]; // Subroutine labels are nearly unlimited in length, so use a size to cover anything realistic.
		if (mWindowIndex > 0) // Prepend the window number for windows other than the first.
			_itoa(mWindowIndex + 1, label_name, 10);
		else
			*label_name = '\0';
		snprintfcat(label_name, sizeof(label_name), "Button%s", aText);
		// Remove spaces and ampersands.  Although ampersands are legal in labels, it seems
		// more friendly to omit them in the automatic-label label name.  Note that a button
		// or menu item can contain a literal ampersand by using two ampersands, such as
		// "Save && Exit" (in this example, the auto-label would be named "ButtonSaveExit").
		StrReplaceAll(label_name, " ", "", true);
		StrReplaceAll(label_name, "&", "", true);
		StrReplaceAll(label_name, "\r", "", true); // Done separate from \n in case they're ever unpaired.
		StrReplaceAll(label_name, "\n", "", true);
		control.jump_to_label = g_script.FindLabel(label_name);  // OK if NULL (the button will do nothing).
	}

	GuiControlType *owning_tab_control = FindTabControl(control.tab_control_index); // For use in various places.

	////////////////////////////////////////////////////////////////////////////////////////////
	// Automatically set the control's position in the client area if no position was specified.
	////////////////////////////////////////////////////////////////////////////////////////////
	if (opt.x == COORD_UNSPECIFIED && opt.y == COORD_UNSPECIFIED)
	{
		if (owning_tab_control && !GetControlCountOnTabPage(control.tab_control_index, control.tab_index)) // Relies on short-circuit boolean.
		{
			// Since this control belongs to a tab control and that tab control already exists,
			// Position it relative to the tab control's client area upper-left corner if this
			// is the first control on this particular tab/page:
			POINT pt = GetPositionOfTabClientArea(*owning_tab_control);
			// Since both coords were unspecified, position this control at the upper left corner of its page.
			opt.x = pt.x + mMarginX;
			opt.y = pt.y + mMarginY;
		}
		else
		{
			// Since both coords were unspecified, proceed downward from the previous control, using a default margin.
			opt.x = mPrevX;
			opt.y = mPrevY + mPrevHeight + mMarginY;  // Don't use mMaxExtentDown in this is a new column.
		}
		if (aControlType == GUI_CONTROL_TEXT && mControlCount && mControl[mControlCount - 1].type == GUI_CONTROL_TEXT)
			// Since this text control is being auto-positioned immediately below another, provide extra
			// margin space so that any edit control, dropdownlist, or other "tall input" control later added
			// to its right in "vertical progression" mode will line up with it:
			opt.y += GUI_CTL_VERTICAL_DEADSPACE;
	}
	// Can't happen due to the logic in the options-parsing section:
	//else if (opt.x == COORD_UNSPECIFIED)
	//	opt.x = mPrevX;
	//else if (y == COORD_UNSPECIFIED)
	//	opt.y = mPrevY;


	/////////////////////////////////////////////////////////////////////////////////////
	// For certain types of controls, provide a standard row_count if none was specified.
	/////////////////////////////////////////////////////////////////////////////////////
	bool calc_control_height_from_row_count = true; // Set default for all control types.

	if (opt.height == COORD_UNSPECIFIED && opt.row_count <= 0)
	{
		switch(aControlType)
		{
		case GUI_CONTROL_DROPDOWNLIST:
		case GUI_CONTROL_COMBOBOX: // For these 2, row-count is defined as the number of items to display in the list.
			// Update: Unfortunately, heights taller than the desktop do not work: pre-v6 common controls
			// misbehave when the height is too tall to fit on the screen.  So the below comment is
			// obsolete and kept only for reference:
			// Since no height or row-count was given, make the control very tall so that OSes older than
			// XP will behavior similar to XP: they will let the desktop height determine how tall the
			// control can be. One exception to this is a true CBS_SIMPLE combo, which has appearance
			// and functionality similar to a ListBox.  In that case, a default row-count is provided
			// since that is more appropriate than having a really tall box hogging the window.
			// Because CBS_DROPDOWNLIST includes both CBS_SIMPLE and CBS_DROPDOWN, a true "simple"
			// (which is similar to a listbox) must omit CBS_DROPDOWN:
			opt.row_count = 3;  // Actual height will be calculated below using this.
			// Avoid doing various calculations later below if the XP+ will ignore the height anyway.
			// CBS_NOINTEGRALHEIGHT is checked in case that style was explicitly applied to the control
			// by the script. This is because on XP+ it will cause the combo/DropDownList to obey the
			// specified default height set above.  Also, for a pure CBS_SIMPLE combo, the OS always
			// obeys height:
			if ((!(style & CBS_SIMPLE) || (style & CBS_DROPDOWN)) // Not a pure CBS_SIMPLE.
				&& g_os.IsWinXPorLater() // ... and the OS is XP+.
				&& !(style & CBS_NOINTEGRALHEIGHT)) // ... and XP won't obey the height.
				calc_control_height_from_row_count = false; // Don't bother calculating the height (i.e. override the default).
			break;
		case GUI_CONTROL_LISTBOX:
			opt.row_count = 3;  // Actual height will be calculated below using this.
			break;
		case GUI_CONTROL_GROUPBOX:
			// Seems more appropriate to give GUI_CONTROL_GROUPBOX exactly two rows: the first for the
			// title of the group-box and the second for its content (since it might contain controls
			// placed horizontally end-to-end, and thus only need one row).
			opt.row_count = 2;
			break;
		case GUI_CONTROL_EDIT:
			// If there's no default text in the control from which to later calc the height, use a basic default.
			if (!*aText)
				opt.row_count = (style & ES_MULTILINE) ? 3.0F : 1.0F;
			break;
		case GUI_CONTROL_HOTKEY:
			opt.row_count = 1;
		case GUI_CONTROL_SLIDER:
			// Make vertical trackbars tall by default.  If their width has also been omitted, they
			// will be made narrow by default in a later section.
			if (style & TBS_VERT)
				opt.row_count = 5.0F;
			else
				opt.height = ControlGetDefaultSliderThickness(style, opt.thickness);
			break;
		case GUI_CONTROL_PROGRESS:
			// Make vertical progress bars tall by default.  If their width has also been omitted, they
			// will be made narrow by default in a later section.
			if (style & PBS_VERTICAL)
				opt.row_count = 5.0F;
			else
				// Height vs. row_count is specified to ensure the same thickness for both vertical
				// and horizontal progress bars:
				#define PROGRESS_DEFAULT_THICKNESS (2 * sFont[mCurrentFontIndex].point_size)
				opt.height = PROGRESS_DEFAULT_THICKNESS;
			break;
		case GUI_CONTROL_TAB:
			opt.row_count = 10;
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
	else // Either a row_count or a height was explicitly specified.
		// If OS is XP+, must apply the CBS_NOINTEGRALHEIGHT style for these reasons:
		// 1) The app now has a manifest, which tells OS to use common controls v6.
		// 2) Common controls v6 will not obey the the user's specified height for the control's
		//    list portion unless the CBS_NOINTEGRALHEIGHT style is present.
		if ((aControlType == GUI_CONTROL_DROPDOWNLIST || aControlType == GUI_CONTROL_COMBOBOX) && g_os.IsWinXPorLater())
			style |= CBS_NOINTEGRALHEIGHT; // Forcibly applied, even if removed in options.

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
	if (opt.row_count > 0)
	{
		// For GroupBoxes, add 1 to any row_count greater than 1 so that the title itself is
		// already included automatically.  In other words, the R-value specified by the user
		// should be the number of rows available INSIDE the box.
		// For DropDownLists and ComboBoxes, 1 is added because row_count is defined as the
		// number of rows shown in the drop-down portion of the control, so we need one extra
		// (used in later calculations) for the always visible portion of the control.
		switch (aControlType)
		{
		case GUI_CONTROL_GROUPBOX:
		case GUI_CONTROL_DROPDOWNLIST:
		case GUI_CONTROL_COMBOBOX:
			++opt.row_count;
			break;
		}
		if (calc_control_height_from_row_count)
		{
			GUI_SET_HDC
			GetTextMetrics(hdc, &tm);
			// Calc the height by adding up the font height for each row, and including the space between lines
			// (tmExternalLeading) if there is more than one line.  0.5 is used in two places to prevent
			// negatives in one, and round the overall result in the other.
			opt.height = (int)((tm.tmHeight * opt.row_count) + (tm.tmExternalLeading * ((int)(opt.row_count + 0.5) - 1)) + 0.5);
			switch (aControlType)
			{
			case GUI_CONTROL_DROPDOWNLIST:
			case GUI_CONTROL_COMBOBOX:
			case GUI_CONTROL_LISTBOX:
			case GUI_CONTROL_EDIT:
			case GUI_CONTROL_HOTKEY:
				opt.height += GUI_CTL_VERTICAL_DEADSPACE;
				if (style & WS_HSCROLL)
					opt.height += GetSystemMetrics(SM_CYHSCROLL);
				break;
			case GUI_CONTROL_BUTTON:
				// Provide a extra space for top/bottom margin together, proportional to the current font
				// size so that it looks better with very large or small fonts.  The +2 seems to make
				// it look just right on all font sizes, especially the default GUI size of 8 where the
				// height should be about 23 to be standard(?)
				opt.height += sFont[mCurrentFontIndex].point_size + 2;
				break;
			case GUI_CONTROL_GROUPBOX: // Since groups usually contain other controls, the below sizing seems best.
				// Use row_count-2 because of the +1 added above for GUI_CONTROL_GROUPBOX.
				// The current font's height is added in to provide an upper/lower margin in the box
				// proportional to the current font size, which makes it look better in most cases:
				opt.height += mMarginY * (2 + ((int)(opt.row_count + 0.5) - 2));
				break;
			case GUI_CONTROL_TAB:
				opt.height += mMarginY * (2 + ((int)(opt.row_count + 0.5) - 1)); // -1 vs. the -2 used above.
				break;
			// Types not included
			// ------------------
			//case GUI_CONTROL_TEXT:     Uses basic height calculated above the switch().
			//case GUI_CONTROL_PIC:      Uses basic height calculated above the switch() (seems OK even for pic).
			//case GUI_CONTROL_CHECKBOX: Uses basic height calculated above the switch().
			//case GUI_CONTROL_RADIO:    Same.
			//case GUI_CONTROL_SLIDER:   Same.
			//case GUI_CONTROL_PROGRESS: Same.
			} // switch
		}
		else // calc_control_height_from_row_count == false
			// Assign a default just to allow the control to be created successfully. 13 is the default
			// height of a text/radio control for the typical 8 point font size, but the exact value
			// shouldn't matter (within reason) since calc_control_height_from_row_count is telling us this type of
			// control will not obey the height anyway.  Update: It seems better to use a small constant
			// value to help catch bugs while still allowing the control to be created:
			opt.height = 30;  // formerly: (int)(13 * opt.row_count);
	}

	if (opt.height == COORD_UNSPECIFIED || opt.width == COORD_UNSPECIFIED)
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
			// Determine whether there will be a vertical scrollbar present.  If ES_MULTILINE hasn't
			// already been applied or auto-detected above, it's possible that a scrollbar will be
			// added later due to the text auto-wrapping.  In that case, the calculated height may
			// be incorrect due to the additional wrapping caused by the width taken up by the
			// scrollbar.  Since this combination of circumstances is rare, and since there are easy
			// workarounds, it's just documented here as a limitation:
			if (style & WS_VSCROLL)
				extra_width += GetSystemMetrics(SM_CXVSCROLL);
			// DT_EDITCONTROL: "the average character width is calculated in the same manner as for an edit control"
			// It might help some aspects of the estimate conducted below.
			// Also include DT_EXPANDTABS under the assumption that if there are tabs present, the user
			// intended for them to be there because a multiline edit would expand them (rather than trying
			// to worry about whether this control *might* become auto-multiline after this point.
			draw_format |= DT_EXPANDTABS|DT_EDITCONTROL;
			// and now fall through and have the dimensions calculated based on what's in the control.
		case GUI_CONTROL_TEXT:
		case GUI_CONTROL_BUTTON:
		case GUI_CONTROL_CHECKBOX:
		case GUI_CONTROL_RADIO:
		{
			GUI_SET_HDC
			if (aControlType == GUI_CONTROL_TEXT)
				draw_format |= DT_EXPANDTABS; // Buttons can't expand tabs, so don't add this for them.
			else if (aControlType == GUI_CONTROL_CHECKBOX || aControlType == GUI_CONTROL_RADIO)
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
			if (opt.width != COORD_UNSPECIFIED) // Since a width was given, auto-expand the height via word-wrapping.
				draw_format |= DT_WORDBREAK;
			RECT draw_rect;
			draw_rect.left = 0;
			draw_rect.top = 0;
			draw_rect.right = (opt.width == COORD_UNSPECIFIED) ? 0 : opt.width - extra_width; // extra_width
			draw_rect.bottom = (opt.height == COORD_UNSPECIFIED) ? 0 : opt.height;
			// If no text, "H" is used in case the function requires a non-empty string to give consistent results:
			int draw_height = DrawText(hdc, *aText ? aText : "H", -1, &draw_rect, draw_format);
			int draw_width = draw_rect.right - draw_rect.left;
			// Even if either height or width was already explicitly specified above, it seems best to
			// override it if DrawText() says it's not big enough.  REASONING: It seems too rare that
			// someone would want to use an explicit height/width to selectively hide part of a control's
			// contents, presumably for revelation later.  If that is truly desired, ControlMove or
			// similar can be used to resize the control afterward.  In addition, by specifying BOTH
			// width and height/rows, none of these calculations happens anyway, so that's another way
			// this override can be overridden:
			if (opt.height == COORD_UNSPECIFIED || draw_height > opt.height)
			{
				opt.height = draw_height;
				if (aControlType == GUI_CONTROL_EDIT)
				{
					opt.height += GUI_CTL_VERTICAL_DEADSPACE;
					if (style & WS_HSCROLL)
						opt.height += GetSystemMetrics(SM_CYHSCROLL);
				}
				else if (aControlType == GUI_CONTROL_BUTTON)
					opt.height += sFont[mCurrentFontIndex].point_size + 2;  // +2 makes it standard height.
			}
			if (opt.width == COORD_UNSPECIFIED || draw_width > opt.width)
			{
				opt.width = draw_width + extra_width;
				if (aControlType == GUI_CONTROL_BUTTON)
					// Allow room for border and an internal margin proportional to the font height.
					// Button's border is 3D by default, so SM_CXEDGE vs. SM_CXBORDER is used?
					opt.width += 2 * GetSystemMetrics(SM_CXEDGE) + sFont[mCurrentFontIndex].point_size;
			}
			break;
		} // case

		// Types not included
		// ------------------
		//case GUI_CONTROL_PIC:           If needed, it is given some default dimensions at the time of creation.
		//case GUI_CONTROL_GROUPBOX:      Seems too rare than anyone would want its width determined by its text.
		//case GUI_CONTROL_EDIT:          It is included, but only if it has default text inside it.
		//case GUI_CONTROL_TAB:           Seems too rare than anyone would want its width determined by tab-count.
		//case GUI_CONTROL_DROPDOWNLIST:  These last 6 are given (later below) a standard width based on font size.
		//case GUI_CONTROL_COMBOBOX:      In addition, their height has already been determined further above.
		//case GUI_CONTROL_LISTBOX:
		//case GUI_CONTROL_HOTKEY:
		//case GUI_CONTROL_SLIDER:
		//case GUI_CONTROL_PROGRESS:
		} // switch()
	}

	//////////////////////////////////////////////////////////////////////////////////////////
	// If the width was not specified and the above did not already determine it (which should
	// only be possible for the cases contained in the switch-stmt below), provide a default.
	//////////////////////////////////////////////////////////////////////////////////////////
	if (opt.width == COORD_UNSPECIFIED)
	{
		switch(aControlType)
		{
		case GUI_CONTROL_DROPDOWNLIST:
		case GUI_CONTROL_COMBOBOX:
		case GUI_CONTROL_LISTBOX:
		case GUI_CONTROL_HOTKEY:
		case GUI_CONTROL_EDIT:
			opt.width = GUI_STANDARD_WIDTH;
			break;
		case GUI_CONTROL_SLIDER:
			// Make vertical trackbars narrow by default.  For vertical trackbars: there doesn't seem
			// to be much point in defaulting the width to something proportional to font size because
			// the thumb only seems to have two sizes and doesn't auto-grow any larger than that.
			opt.width = (style & TBS_VERT) ? ControlGetDefaultSliderThickness(style, opt.thickness) : GUI_STANDARD_WIDTH;
			break;
		case GUI_CONTROL_PROGRESS:
			opt.width = (style & PBS_VERTICAL) ? PROGRESS_DEFAULT_THICKNESS : GUI_STANDARD_WIDTH;
			break;
		case GUI_CONTROL_GROUPBOX:
			// Since groups and tabs contain other controls, allow room inside them for a margin based
			// on current font size.
			opt.width = GUI_STANDARD_WIDTH + (2 * mMarginX);
			break;
		case GUI_CONTROL_TAB:
			// Tabs tend to be wide so that that tabs can all fit on the top row, and because they
			// are usually used to fill up the entire window.  Therefore, default them to the ability
			// to hold two columns of standard-width controls:
			opt.width = (2 * GUI_STANDARD_WIDTH) + (3 * mMarginX);  // 3 vs. 2 to allow space in between columns.
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
	// auto-detect that by comparing the current font size with the specified height. At this
	// stage, the above has already ensured that an Edit has at least a height or a row_count.
	/////////////////////////////////////////////////////////////////////////////////////////
	if (aControlType == GUI_CONTROL_EDIT && !(style & ES_MULTILINE))
	{
		if (opt.row_count <= 0) // Determine the row-count to auto-detect multi-line vs. single-line.
		{
			GUI_SET_HDC
			GetTextMetrics(hdc, &tm);
			int height_beyond_first_row = opt.height - GUI_CTL_VERTICAL_DEADSPACE - tm.tmHeight;
			if (style & WS_HSCROLL)
				height_beyond_first_row -= GetSystemMetrics(SM_CYHSCROLL);
			if (height_beyond_first_row > 0)
			{
				opt.row_count = 1 + ((float)height_beyond_first_row / (tm.tmHeight + tm.tmExternalLeading));
				// This section is a near exact match for one higher above.  Search for comment
				// "Add multiline unless it was explicitly removed" for a full explanation and keep
				// the below in sync with that section above:
				if (opt.row_count > 1.5)
				{
					style |= (ES_MULTILINE & ~opt.style_remove); // Add multiline unless it was explicitly removed.
					// Do the below only if the above actually added multiline:
					if (style & ES_MULTILINE) // If allowed, enable vertical scrollbar and capturing of ENTER keystrokes.
						style |= EDIT_MULTILINE_DEFAULT & ~opt.style_remove;
					// else: Single-line edit.  ES_AUTOHSCROLL will be applied later below if all the other checks
					// fail to auto-detect this edit as a multi-line edit.
				}
			}
			else // there appears to be only one row.
				opt.row_count = 1;
				// And ES_AUTOHSCROLL will be applied later below if all the other checks
				// fail to auto-detect this edit as a multi-line edit.
		}
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
	bool do_strip_theme = !mUseTheme;   // Set defaults.
	bool font_was_set = false;          // "
	bool retrieve_dimensions = false;   // "
	int item_height, min_list_height;
	RECT rect;
	char *malloc_buf;
	HMENU control_id = (HMENU)(size_t)GUI_INDEX_TO_ID(mControlCount); // Cast to size_t avoids compiler warning.

	// If a control is being added to a tab, even if the parent window is hidden (since it might
	// have been hidden by Gui, Cancel), make sure the control isn't visible unless it's on a
	// visible tab.
	// The below alters style vs. style_remove, since later below style_remove is checked to
	// find out if the control was explicitly hidden vs. hidden by the automatic action here:
	if (control.tab_control_index < MAX_TAB_CONTROLS) // This control belongs to a tab control (must check this even though FindTabControl() does too).
	{
		if (owning_tab_control) // Its tab control exists...
		{
			if (!(GetWindowLong(owning_tab_control->hwnd, GWL_STYLE) & WS_VISIBLE) // Don't use IsWindowVisible().
				|| TabCtrl_GetCurSel(owning_tab_control->hwnd) != control.tab_index)
				// ... but it's not set to the page/tab that contains this control, or the entire tab control is hidden.
				style &= ~WS_VISIBLE;
		}
		else // Its tab control does not exist, so this control is kept hidden until such time that it does.
			style &= ~WS_VISIBLE;
	}
	// else do nothing.

	switch(aControlType)
	{
	case GUI_CONTROL_TEXT:
		// Seems best to omit SS_NOPREFIX by default so that ampersand can be used to create shortcut keys.
		control.hwnd = CreateWindowEx(exstyle, "static", aText, style
			, opt.x, opt.y, opt.width, opt.height, mHwnd, control_id, g_hInstance, NULL);
		break;

	case GUI_CONTROL_PIC:
		if (opt.width == COORD_UNSPECIFIED)
			opt.width = 0;  // Use zero to tell LoadPicture() to keep original width.
		if (opt.height == COORD_UNSPECIFIED)
			opt.height = 0;  // Use zero to tell LoadPicture() to keep original height.
		// Must set its caption to aText so that documented ability to refer to a picture by its original
		// filename is possible:
		if (control.hwnd = CreateWindowEx(exstyle, "static", aText, style
			, opt.x, opt.y, opt.width, opt.height  // OK if zero, control creation should still succeed.
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
			// LoadPicture() uses CopyImage() to scale the image, which seems to provide better scaling
			// quality than using MoveWindow() (followed by redrawing the parent window) on the static
			// control that contains the image.
			int image_type;
			// In the below, opt.icon_number is set to zero if there was no preference specified, which is
			// important because it tells LoadPicture() that it can safely use LoadImage() vs. ExtractIcon()
			// on a .ico, .cur, or .ani file, which in turn allows the icon/cursor to be animated when
			// otherwise it wouldn't be:
			if (   !(control.union_hbitmap = LoadPicture(aText, opt.width, opt.height, image_type, opt.icon_number - 1
				, control.attrib & GUI_CONTROL_ATTRIB_ALTSUBMIT))   )
				break;  // By design, no error is reported.  The picture is simply not displayed, nor is its
						// style set to SS_BITMAP/SS_ICON, which allows the control to have the specified size
						// yet lack an image (SS_BITMAP/SS_ICON tend to cause the control to auto-size to
						// zero dimensions).
			// For image to display correctly, must apply SS_ICON for cursors/icons and SS_BITMAP for bitmaps.
			// This style change is made *after* the control is created so that the act of creation doesn't
			// attempt to load the image from a resource (which as documented by SS_ICON/SS_BITMAP, would happen
			// since text is interpreted as the name of a bitmap in the resource file).
			SetWindowLong(control.hwnd, GWL_STYLE, style | (image_type == IMAGE_BITMAP ? SS_BITMAP : SS_ICON));
			// Above uses ~0x0F to ensure the lowest four/five bits are pure.
			// Also note that it does not seem correct to use SS_TYPEMASK if bitmaps/icons can also have
			// any of the following styles.  MSDN confirms(?) this by saying that SS_TYPEMASK is out of date
			// and should not be used:
			//#define SS_ETCHEDHORZ       0x00000010L
			//#define SS_ETCHEDVERT       0x00000011L
			//#define SS_ETCHEDFRAME      0x00000012L
			SendMessage(control.hwnd, STM_SETIMAGE, (WPARAM)image_type, (LPARAM)control.union_hbitmap);
			if (image_type == IMAGE_BITMAP)
				control.attrib &= ~GUI_CONTROL_ATTRIB_ALTBEHAVIOR;  // Flag it as a bitmap so that DeleteObject vs. DestroyIcon will be called for it.
			else // Cursor or Icon, which are functionally identical for our purposes.
				control.attrib |= GUI_CONTROL_ATTRIB_ALTBEHAVIOR;
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
			//    and any later controls that overlap are drawn on top of it, except for controls that
			//    have WS_CLIPSIBLINGS).
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
			//GuiIndexType index_of_last_picture = UINT_MAX;
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
		control.hwnd = CreateWindowEx(exstyle, "button", aText, style
			, opt.x, opt.y, opt.width, opt.height, mHwnd, control_id, g_hInstance, NULL);
		break;

	case GUI_CONTROL_BUTTON:
		// For all "button" type controls, BS_MULTILINE is included by default so that any literal
		// newlines in the button's name will start a new line of text as the user intended.
		// In addition, this causes automatic wrapping to occur if the user specified a width
		// too small to fit the entire line.
		if (control.hwnd = CreateWindowEx(exstyle, "button", aText, style
			, opt.x, opt.y, opt.width, opt.height, mHwnd, control_id, g_hInstance, NULL))
		{
			if (style & BS_DEFPUSHBUTTON)
			{
				// First remove the style from the old default button, if there is one:
				if (mDefaultButtonIndex < mControlCount)
				{
					// MSDN says this is necessary in some cases:
					// Since the window might be visbible at this point, send BM_SETSTYLE rather than
					// SetWindowLong() so that the button will get redrawn.  Update: The redraw doesn't
					// actually seem to happen, but this is kept because it also serves to change the
					// default button appearance later, which is received in the WindowProc via WM_COMMAND:
					SendMessage(mControl[mDefaultButtonIndex].hwnd, BM_SETSTYLE
						, (WPARAM)LOWORD((GetWindowLong(mControl[mDefaultButtonIndex].hwnd, GWL_STYLE) & ~BS_DEFPUSHBUTTON))
						, MAKELPARAM(TRUE, 0)); // Redraw = yes. It's probably smart enough not to do it if the window is hidden.
					// The below attempts to get the old button to lose its default-border failed.  This might
					// be due to the fact that if the window hasn't yet been shown for the first time, its
					// client area isn't yet the right size, so the OS decides that no update is needed since
					// the control is probably outside the boundaries of the window:
					//InvalidateRect(mHwnd, NULL, TRUE);
					//GetClientRect(mControl[mDefaultButtonIndex].hwnd, &client_rect);
					//InvalidateRect(mControl[mDefaultButtonIndex].hwnd, &client_rect, TRUE);
					//ShowWindow(mHwnd, SW_SHOWNOACTIVATE); // i.e. don't activate it if it wasn't before.
					//ShowWindow(mHwnd, SW_HIDE);
					//UpdateWindow(mHwnd);
					//SendMessage(mHwnd, WM_NCPAINT, 1, 0);
					//RedrawWindow(mHwnd, NULL, NULL, RDW_INVALIDATE|RDW_FRAME|RDW_UPDATENOW);
					//SetWindowPos(mHwnd, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
				}
				mDefaultButtonIndex = mControlCount;
				SendMessage(mHwnd, DM_SETDEFID, (WPARAM)GUI_INDEX_TO_ID(mDefaultButtonIndex), 0);
				// This is necessary to make the default button have its visual style when the dialog is first shown:
				SendMessage(mControl[mDefaultButtonIndex].hwnd, BM_SETSTYLE
					, (WPARAM)LOWORD((GetWindowLong(mControl[mDefaultButtonIndex].hwnd, GWL_STYLE) | BS_DEFPUSHBUTTON))
					, MAKELPARAM(TRUE, 0)); // Redraw = yes. It's probably smart enough not to do it if the window is hidden.
			}
		}
		break;

	case GUI_CONTROL_CHECKBOX:
		// The BS_NOTIFY style is not a good idea for checkboxes because although it causes the control
		// to send BN_DBLCLK messages, any rapid clicks by the user on (for example) a tri-state checkbox
		// are seen only as one click for the purpose of changing the box's state.
		if (control.hwnd = CreateWindowEx(exstyle, "button", aText, style
			, opt.x, opt.y, opt.width, opt.height, mHwnd, control_id, g_hInstance, NULL))
		{
			if (opt.checked != BST_UNCHECKED) // Set the specified state.
				SendMessage(control.hwnd, BM_SETCHECK, opt.checked, 0);
		}
		break;

	case GUI_CONTROL_RADIO:
		control.hwnd = CreateWindowEx(exstyle, "button", aText, style
			, opt.x, opt.y, opt.width, opt.height, mHwnd, control_id, g_hInstance, NULL);
		// opt.checked is handled later below.
		break;

	case GUI_CONTROL_DROPDOWNLIST:
	case GUI_CONTROL_COMBOBOX:
		// It has been verified that that EM_LIMITTEXT has no effect when sent directly
		// to a ComboBox hwnd; however, it might work if sent to its edit-child. But for now,
		// a Combobox can only be limited to its visible width.  Later, there might
		// be a way to send a message to its child control to limit its width directly.
		if (opt.limit && control.type == GUI_CONTROL_COMBOBOX)
			style &= ~CBS_AUTOHSCROLL;
		// Since the control's variable can change, it seems best to pass in the empty string
		// as the control's caption, rather than the name of the variable.  The name of the variable
		// isn't that useful anymore anyway since GuiControl(Get) can access controls directly by
		// their current output-var names:
		if (control.hwnd = CreateWindowEx(exstyle, "Combobox", "", style
			, opt.x, opt.y, opt.width, opt.height, mHwnd, control_id, g_hInstance, NULL))
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
			if (calc_control_height_from_row_count)
			{
				item_height = (int)SendMessage(control.hwnd, CB_GETITEMHEIGHT, 0, 0);
				// Note that at this stage, height should contain a explicitly-specified height or height
				// estimate from the above, even if row_count is greater than 0.
				// The below calculation may need some fine tuning:
				int cbs_extra_height = ((style & CBS_SIMPLE) && !(style & CBS_DROPDOWN)) ? 4 : 2;
				min_list_height = (2 * item_height) + GUI_CTL_VERTICAL_DEADSPACE + cbs_extra_height;
				if (opt.height < min_list_height) // Adjust so that at least 1 item can be shown.
					opt.height = min_list_height;
				else if (opt.row_count > 0)
					// Now that we know the true item height (since the control has been created and we asked
					// it), resize the control to try to get it to the match the specified number of rows.
					// +2 seems to be the exact amount needed to prevent partial rows from showing
					// on all font sizes and themes when NOINTEGRALHEIGHT is in effect:
					opt.height = (int)(opt.row_count * item_height) + GUI_CTL_VERTICAL_DEADSPACE + cbs_extra_height;
			}
			MoveWindow(control.hwnd, opt.x, opt.y, opt.width, opt.height, TRUE); // Repaint, since it might be visible.
			// Since combo's size is created to accomodate its drop-down height, adjust our true height
			// to its actual collapsed size.  This true height is used for auto-positioning the next
			// control, if it uses auto-positioning.  It might be possible for it's width to be different
			// also, such as if it snaps to a certain minimize width if one too small was specified,
			// so that is recalculated too:
			retrieve_dimensions = true;
		}
		break;

	case GUI_CONTROL_LISTBOX:
		// See GUI_CONTROL_COMBOBOX above for why empty string is passed in as the caption:
		if (control.hwnd = CreateWindowEx(exstyle, "Listbox", "", style
			, opt.x, opt.y, opt.width, opt.height, mHwnd, control_id, g_hInstance, NULL))
		{
			if (opt.tabstop_count)
				SendMessage(control.hwnd, LB_SETTABSTOPS, opt.tabstop_count, (LPARAM)opt.tabstop);
			// For now, it seems best to always override a height that would cause zero items to be
			// displayed.  This is because there is a very thin control visible even if the height
			// is explicitly set to zero, which seems pointless (there are other ways to draw thin
			// looking objects for unusual purposes anyway).
			// Set font unconditionally to simplify calculations, which help ensure that at least one item
			// in the DropDownList/Combo is visible when the list drops down:
			GUI_SETFONT // Set font in preparation for asking it how tall each item is.
			item_height = (int)SendMessage(control.hwnd, LB_GETITEMHEIGHT, 0, 0);
			// Note that at this stage, height should contain a explicitly-specified height or height
			// estimate from the above, even if opt.row_count is greater than 0.
			min_list_height = item_height + GUI_CTL_VERTICAL_DEADSPACE;
			if (style & WS_HSCROLL)
				// Assume bar will be actually appear even though it won't in the rare case where
				// its specified pixel-width is smaller than the width of the window:
				min_list_height += GetSystemMetrics(SM_CYHSCROLL);
			if (opt.height < min_list_height) // Adjust so that at least 1 item can be shown.
				opt.height = min_list_height;
			else if (opt.row_count > 0)
			{
				// Now that we know the true item height (since the control has been created and we asked
				// it), resize the control to try to get it to the match the specified number of rows.
				opt.height = (int)(opt.row_count * item_height) + GUI_CTL_VERTICAL_DEADSPACE;
				if (style & WS_HSCROLL)
					// Assume bar will be actually appear even though it won't in the rare case where
					// its specified pixel-width is smaller than the width of the window:
				opt.height += GetSystemMetrics(SM_CYHSCROLL);
			}
			MoveWindow(control.hwnd, opt.x, opt.y, opt.width, opt.height, TRUE); // Repaint, since it might be visible.
			// Since by default, the OS adjusts list's height to prevent a partial item from showing
			// (LBS_NOINTEGRALHEIGHT), fetch the actual height for possible use in positioning the
			// next control:
			retrieve_dimensions = true;
		}
		break;

	case GUI_CONTROL_EDIT:
		if (!(style & ES_MULTILINE)) // ES_MULTILINE was not explicitly or automatically specified.
		{
			if (opt.limit < 0) // This is the signal to limit input length to visible width of field.
				// But it can only work if the Edit isn't a multiline.
				style &= ~(WS_HSCROLL|ES_AUTOHSCROLL); // Enable the limiting style.
			else // Since this is a single-line edit, add AutoHScroll if it wasn't explicitly removed.
				style |= ES_AUTOHSCROLL & ~opt.style_remove;
				// But no attempt is made to turn off WS_VSCROLL or ES_WANTRETURN since those might have some
				// usefulness even in a single-line edit?  In any case, it seems too overprotective to do so.
		}
		// malloc() is done because edit controls in NT/2k/XP support more than 64K.
		// Mem alloc errors are so rare (since the text is usually less than 32K/64K) that no error is displayed.
		// Instead, the un-translated text is put in directly.  Also, translation is not done for
		// single-line edits since they can't display linebreaks correctly anyway.
		// Note that TranslateLFtoCRLF() will return the original buffer we gave it if no translation
		// is needed.  Otherwise, it will return a new buffer which we are responsible for freeing
		// when done (or NULL if it failed to allocate the memory).
		malloc_buf = (*aText && (style & ES_MULTILINE)) ? TranslateLFtoCRLF(aText) : aText;
		if (control.hwnd = CreateWindowEx(exstyle, "edit", malloc_buf ? malloc_buf : aText, style  // malloc_buf is checked again in case mem alloc failed.
			, opt.x, opt.y, opt.width, opt.height, mHwnd, control_id, g_hInstance, NULL))
		{
			// As documented in MSDN, setting a password char will have no effect for multi-line edits
			// since they do not support password/mask char.
			// It seems best to allow password_char to be a literal asterisk so that there's a way to
			// have asterisk vs. bullet/closed-circle on OSes that default to bullet.
			if ((style & ES_PASSWORD) && opt.password_char) // Override default.
				SendMessage(control.hwnd, EM_SETPASSWORDCHAR, (WPARAM)opt.password_char, 0);
			// For the below, note that EM_LIMITTEXT == EM_SETLIMITTEXT.
			if (opt.limit < 0)
				opt.limit = 0;
			//else leave it as the zero (unlimited) or positive (restricted) limit already set.
			// Now set the limit. Specifying a limit of zero opens the control to its maximum text capacity,
			// which removes the 32K size restriction.  Testing shows that this does not increase the actual
			// amount of memory used for controls containing small amounts of text.  All it does is allow
			// the control to allocate more memory as the user enters text.  By specifying zero, a max
			// of 64K becomes available on Windows 9x, and perhaps as much as 4 GB on NT/2k/XP.
			SendMessage(control.hwnd, EM_LIMITTEXT, (WPARAM)opt.limit, 0);
			if (opt.tabstop_count)
				SendMessage(control.hwnd, EM_SETTABSTOPS, opt.tabstop_count, (LPARAM)opt.tabstop);
		}
		if (malloc_buf && malloc_buf != aText)
			free(malloc_buf);
		break;

	case GUI_CONTROL_HOTKEY:
		// In this case, not only doesn't the caption appear anywhere, it's not set either (or at least
		// not retrievable via GetWindowText()):
		if (control.hwnd = CreateWindowEx(exstyle, HOTKEY_CLASS, "", style
			, opt.x, opt.y, opt.width, opt.height, mHwnd, control_id, g_hInstance, NULL))
		{
			if (*aText)
				SendMessage(control.hwnd, HKM_SETHOTKEY, TextToHotkey(aText), 0);
			if (opt.limit > 0)
				SendMessage(control.hwnd, HKM_SETRULES, opt.limit, MAKELPARAM(HOTKEYF_CONTROL|HOTKEYF_ALT, 0));
				// Above must also specify Ctrl+Alt or some other default, otherwise the restriction will have
				// no effect.
		}
		break;

	case GUI_CONTROL_SLIDER:
		if (control.hwnd = CreateWindowEx(exstyle, TRACKBAR_CLASS, "", style
			, opt.x, opt.y, opt.width, opt.height, mHwnd, control_id, g_hInstance, NULL))
		{
			ControlSetSliderOptions(control, opt); // Fix for v1.0.25.08: This must be done prior to the below.
			// The control automatically deals with out-of-range values by setting slider to min or max.
			// MSDN: "If this value is outside the control's maximum and minimum range, the position
			// is set to the maximum or minimum value."
			if (*aText)
				SendMessage(control.hwnd, TBM_SETPOS, TRUE, ControlInvertSliderIfNeeded(control, ATOI(aText)));
				// Above msg has no return value.
			//else leave it at the OS's default starting position (probably always the far left or top of the range).
		}
		break;

	case GUI_CONTROL_PROGRESS:
		if (control.hwnd = CreateWindowEx(exstyle, PROGRESS_CLASS, "", style
			, opt.x, opt.y, opt.width, opt.height, mHwnd, control_id, g_hInstance, NULL))
		{
			ControlSetProgressOptions(control, opt, style); // Fix for v1.0.27.01: This must be done prior to the below.
			// This has been confirmed though testing, even when the range is dynamically changed
			// after the control is created to something that no longer includes the bar's current
			// position: The control automatically deals with out-of-range values by setting bar to
			// min or max.
			if (*aText)
				SendMessage(control.hwnd, PBM_SETPOS, ATOI(aText), 0);
			//else leave it at the OS's default starting position (probably always the far left or top of the range).
			do_strip_theme = false;  // The above would have already stripped it if needed, so don't do it again.
		}
		break;

	case GUI_CONTROL_TAB:
		if (control.hwnd = CreateWindowEx(exstyle, WC_TABCONTROL, "", style
			, opt.x, opt.y, opt.width, opt.height, mHwnd, control_id, g_hInstance, NULL))
		{
			// For v1.0.23, theme is removed unconditionally for Tab controls because if an XP theme is
			// in effect, causing a non-solid background (such as an off-white gradient/fade), there are
			// many complications to getting the sub-controls' background to match the gradient.
			// The small advantages (styled tab appearance, yellow-bar hot-tracking, and the dubious
			// cosmetic appeal of the gradient itself) do not seem to outweigh the added complications.
			// The main approaches to supporting a themed tab control in the future are:
			// 1) Making a brush from a bitmap/snapshot of the background and applying that to Radios,
			//    Checkboxes, and GroupBoxes (and possibly other future control types).
			// 2) Using CreateDialog() or such to make a dialog window (child of main window, not
			//    child of the tab control).  The tab's controls are then made children of this dialog
			//    and automatically get the right background appearance by virtue of a call to
			//    EnableThemeDialogTexture().  It seems this call only works on true dialogs and their
			//    children.
			// See this and especially its reponses: http://www.codeproject.com/wtl/ThemedDialog.asp#xx727162xx
			do_strip_theme = true;
			// After a new tab control is created, default all subsequently created controls to belonging
			// to the first tab of this tab control: 
			mCurrentTabControlIndex = mTabControlCount;
			mCurrentTabIndex = 0;
			++mTabControlCount;
			// Override the tab's window-proc so that custom background color becomes possible:
			if (!g_TabClassProc)
				g_TabClassProc = (WNDPROC)GetClassLong(control.hwnd, GCL_WNDPROC);
			SetWindowLong(control.hwnd, GWL_WNDPROC, (LONG)TabWindowProc);
			// Doesn't work to remove theme background from tab:
			//MyEnableThemeDialogTexture(control.hwnd, ETDT_DISABLE);
			// This attempt to apply theme to the entire dialog window also has no effect, probably
			// because ETDT_ENABLETAB only works with true dialog windows (e.g. CreateDialog()):
			//MyEnableThemeDialogTexture(mHwnd, ETDT_ENABLETAB);
			// The above require the following line:
			//#include <uxtheme.h> // For EnableThemeDialogTexture()'s constants.
		}
		break;
	}

	// Below also serves as a bug check, i.e. GUI_CONTROL_INVALID or some unknown type.
	if (!control.hwnd)
		return g_script.ScriptError("The control could not be created." ERR_ABORT);
	// Otherwise the above control creation succeeed.
	++mControlCount;

	if (control.type == GUI_CONTROL_RADIO)
	{
		if (opt.checked != BST_UNCHECKED)
			ControlCheckRadioButton(control, mControlCount - 1, opt.checked); // Also handles alteration of the group's tabstop, etc.
		//else since the control has just been created, there's no need to uncheck it or do any actions
		// related to unchecking, such as tabstop adjustment.
		mInRadioGroup = true; // Set here, only after creation was successful.
	}
	else
		mInRadioGroup = false;

	// Check style_remove vs. style because this control might be hidden just because it was added
	// to a tab that isn't active:
	if (opt.style_remove & WS_VISIBLE)
		control.attrib |= GUI_CONTROL_ATTRIB_EXPLICITLY_HIDDEN;  // For use with tab controls.
	if (opt.style_add & WS_DISABLED)
		control.attrib |= GUI_CONTROL_ATTRIB_EXPLICITLY_DISABLED;

	// Strip theme from the control if called for:
	// It is stripped for radios, checkboxes, and groupboxes if they have a custom text color.
	// Otherwise the transparency and/or custom text color will not be obeyed on XP, at least when
	// a non-Classic theme is active.  For GroupBoxes, when a theme is active, it will obey 
	// custom background color but not a custom text color.  The same is true for radios and checkboxes.
	if (do_strip_theme || (control.union_color != CLR_DEFAULT && (control.type == GUI_CONTROL_CHECKBOX
		|| control.type == GUI_CONTROL_RADIO || control.type == GUI_CONTROL_GROUPBOX)) // GroupBox needs it too.
		|| (control.type == GUI_CONTROL_GROUPBOX && (control.attrib & GUI_CONTROL_ATTRIB_BACKGROUND_TRANS))   ) // Tested and found to be necessary.)
		MySetWindowTheme(control.hwnd, L"", L"");

	///////////////////////////////////////////////////
	// Add any content to the control and set its font.
	///////////////////////////////////////////////////
	ControlAddContents(control, aText, opt.choice);

	// Must set the font even if mCurrentFontIndex > 0, otherwise the bold SYSTEM_FONT will be used.
	// Note: Neither the slider's buddies nor itself are affected by the font setting, so it's not applied.
	// However, the buddies are affected at the time they are created if they are a type that uses a font.
	if (!font_was_set && USES_FONT_AND_TEXT_COLOR(control.type))
		GUI_SETFONT

	if (control.type == GUI_CONTROL_TAB && opt.row_count > 0)
	{
		// Now that the tabs have been added (possibly creating more than one row of tabs), resize so that
		// the interior of the control has the actual number of rows specified.
		GetClientRect(control.hwnd, &rect); // MSDN: "the coordinates of the upper-left corner are (0,0)"
		// This is a workaround for the fact that TabCtrl_AdjustRect() seems to give an invalid
		// height when the tabs are at the bottom, at least on XP.  Unfortunately, this workaround
		// does not work when the tabs or on the left or right side, so don't even bother with that
		// adjustment (it's very rare that a tab control would have an unspecified width anyway).
		bool bottom_is_in_effect = (style & TCS_BOTTOM) && !(style & TCS_VERTICAL);
		if (bottom_is_in_effect)
			SetWindowLong(control.hwnd, GWL_STYLE, style & ~TCS_BOTTOM);
		// Insist on a taller client area (or same height in the case of TCS_VERTICAL):
		TabCtrl_AdjustRect(control.hwnd, TRUE, &rect); // Calculate new window height.
		if (bottom_is_in_effect)
			SetWindowLong(control.hwnd, GWL_STYLE, style);
		opt.height = rect.bottom - rect.top;  // Update opt.height for here and for later use below.
		// The below is commented out because TabCtrl_AdjustRect() is unable to cope with tabs on
		// the left or right sides.  It would be rarely used anyway.
		//if (style & TCS_VERTICAL && width_was_originally_unspecified)
		//	// Also make the interior wider in this case, to make the interior as large as intended.
		//	// It is a known limitation that this adjustment does not occur when the script did not
		//	// specify a row_count or omitted height and row_count.
		//	opt.width = rect.right - rect.left;
		MoveWindow(control.hwnd, opt.x, opt.y, opt.width, opt.height, TRUE); // Repaint, since parent might be visible.
	}

	if (retrieve_dimensions) // Update to actual size for use later below.
	{
		GetWindowRect(control.hwnd, &rect);
		opt.height = rect.bottom - rect.top;
		opt.width = rect.right - rect.left;

		if (aControlType == GUI_CONTROL_LISTBOX && (style & WS_HSCROLL))
		{
			if (opt.hscroll_pixels < 0) // Calculate a default based on control's width.
				// Since horizontal scrollbar is relatively rarely used, no fancy method
				// such as calculating scrolling-width via LB_GETTEXTLEN & current font's
				// average width is used.
				opt.hscroll_pixels = 3 * opt.width;
			// If hscroll_pixels is now zero or smaller than the width of the control, the
			// scrollbar will not be shown.  But the message is still sent unconditionally
			// in case it has some desirable side-effects:
			SendMessage(control.hwnd, LB_SETHORIZONTALEXTENT, (WPARAM)opt.hscroll_pixels, 0);
		}
	}

	////////////////////////////////////////////////////////////////////////////////////////////////////
	// Save the details of this control's position for posible use in auto-positioning the next control.
	////////////////////////////////////////////////////////////////////////////////////////////////////
	mPrevX = opt.x;
	mPrevY = opt.y;
	mPrevWidth = opt.width;
	mPrevHeight = opt.height;
	int right = opt.x + opt.width;
	int bottom = opt.y + opt.height;
	if (right > mMaxExtentRight)
		mMaxExtentRight = right;
	if (bottom > mMaxExtentDown)
		mMaxExtentDown = bottom;

	if (opt.start_new_section) // Always start new section for very first control.
	{
		mSectionX = opt.x;
		mSectionY = opt.y;
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



ResultType GuiType::ParseOptions(char *aOptions, bool &aSetLastFoundWindow)
// This function is similar to ControlParseOptions() further below, so should be maintained alongside it.
// Caller must have already initialized aSetLastFoundWindow with desired starting values.
// Caller must ensure that aOptions is a modifiable string, since this method temporarily alters it.
{
	int owner_window_index;
	DWORD style_orig = mStyle;
	DWORD exstyle_orig = mExStyle;

	char *next_option, *option_end, orig_char;
	bool adding; // Whether this option is beeing added (+) or removed (-).

	for (next_option = aOptions; *next_option; next_option = omit_leading_whitespace(option_end))
	{
		if (*next_option == '-')
		{
			adding = false;
			// omit_leading_whitespace() is not called, which enforces the fact that the option word must
			// immediately follow the +/- sign.  This is done to allow the flexibility to have options
			// omit the plus/minus sign, and also to reserve more flexibility for future option formats.
			++next_option;  // Point it to the option word itself.
		}
		else
		{
			// Assume option is being added in the absence of either sign.  However, when we were
			// called by GuiControl(), the first option in the list must begin with +/- otherwise the cmd
			// would never have been properly detected as GUICONTROL_CMD_OPTIONS in the first place.
			adding = true;
			if (*next_option == '+')
				++next_option;  // Point it to the option word itself.
			//else do not increment, under the assumption that the plus has been omitted from a valid
			// option word and is thus an implicit plus.
		}

		if (!*next_option) // In case the entire option string ends in a naked + or -.
			break;
		// Find the end of this option item:
		if (   !(option_end = StrChrAny(next_option, " \t"))   )  // Space or tab.
			option_end = next_option + strlen(next_option); // Set to position of zero terminator instead.
		if (option_end == next_option)
			continue; // i.e. the string contains a + or - with a space or tab after it, which is intentionally ignored.

		// Temporarily terminate to help eliminate ambiguity for words contained inside other words,
		// such as "Checked" inside of "CheckedGray":
		orig_char = *option_end;
		*option_end = '\0';

		// Attributes and option words:
		if (!strnicmp(next_option, "Owner", 5))
		{
			if (mHwnd) // OS provides no way to change an existing window's owner.
				continue;   // Currently no effect, as documented.
			if (!adding)
				mOwner = NULL;
			else
			{
				if (option_end - next_option > 5) // Length is greater than 5, so it has a number (e.g. Owned1).
				{
					// Using ATOI() vs. atoi() seems okay in these cases since spaces are required
					// between options:
					owner_window_index = ATOI(next_option + 5) - 1;
					if (owner_window_index >= 0 && owner_window_index < MAX_GUI_WINDOWS
						&& owner_window_index != mWindowIndex  // Window can't own itself!
						&& g_gui[owner_window_index] && g_gui[owner_window_index]->mHwnd) // Relies on short-circuit boolean order.
						mOwner = g_gui[owner_window_index]->mHwnd;
					else
						return g_script.ScriptError("The owner window is not valid or does not yet exist."
							ERR_ABORT, next_option);
				}
				else
					mOwner = g_hWnd; // Make a window owned (by script's main window) omits its taskbar button.
			}
		}

		else if (!stricmp(next_option, "AlwaysOnTop"))
		{
			// If the window already exists, SetWindowLong() isn't enough.  Must use SetWindowPos()
			// to make it take effect.
			if (mHwnd)
				SetWindowPos(mHwnd, adding ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0
					, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE); // SWP_NOACTIVATE prevents the side-effect of activating the window, which is undesirable if only its style is changing.
			else // Must do the below ONLY if the window doesn't exist, other will window will be broken.
				if (adding) mExStyle |= WS_EX_TOPMOST; else mStyle = mExStyle & ~WS_EX_TOPMOST;

		}

		else if (!stricmp(next_option, "Border"))
			if (adding) mStyle |= WS_BORDER; else mStyle &= ~WS_BORDER;

		else if (!stricmp(next_option, "Caption"))
			// To remove title bar successfully, the WS_POPUP style must also be applied:
			if (adding) mStyle |= WS_CAPTION; else mStyle = mStyle & ~WS_CAPTION | WS_POPUP;

		else if (!stricmp(next_option, "Disabled"))
		{
			if (mHwnd)
				EnableWindow(mHwnd, adding ? FALSE : TRUE);  // Must not not apply WS_DISABLED directly because that breaks the window.
			else
				if (adding) mStyle |= WS_DISABLED; else mStyle = mStyle & ~WS_DISABLED;
		}

		else if (!stricmp(next_option, "LastFound"))
			aSetLastFoundWindow = true; // Regardless of whether "adding" is true or false.

		else if (!stricmp(next_option, "MaximizeBox")) // See above comment.
			if (adding) mStyle |= WS_MAXIMIZEBOX|WS_SYSMENU; else mStyle &= ~WS_MAXIMIZEBOX;

		else if (!stricmp(next_option, "MinimizeBox"))
			// WS_MINIMIZEBOX requires WS_SYSMENU to take effect.  It can be explicitly omitted
			// via "+MinimizeBox -SysMenu" if that functionality is ever needed.
			if (adding) mStyle |= WS_MINIMIZEBOX|WS_SYSMENU; else mStyle &= ~WS_MINIMIZEBOX;

		else if (!stricmp(next_option, "Resize")) // Minus removes either or both.
			if (adding) mStyle |= WS_SIZEBOX|WS_MAXIMIZEBOX; else mStyle &= ~(WS_SIZEBOX|WS_MAXIMIZEBOX);

		else if (!stricmp(next_option, "SysMenu"))
			if (adding) mStyle |= WS_SYSMENU; else mStyle &= ~WS_SYSMENU;

		else if (!stricmp(next_option, "Theme"))
			mUseTheme = adding;
			// But don't apply/remove theme from parent window because that is usually undesirable.
			// This is because even old apps running on XP still have the new parent window theme,
			// at least for their title bar and title bar buttons (except console apps, maybe).

		else if (!stricmp(next_option, "ToolWindow"))
			// WS_EX_TOOLWINDOW provides narrower title bar, omits task bar button, and omits
			// entry in the alt-tab menu.
			if (adding) mExStyle |= WS_EX_TOOLWINDOW; else mStyle &= ~WS_EX_TOOLWINDOW;

		// This one should be near the bottom since "E" is fairly vague and might be contained at the start
		// of future option words such as Edge, Exit, etc.
		else if (toupper(*next_option) == 'E') // Extended style
		{
			++next_option; // Skip over the E itself.
			if (IsPureNumeric(next_option, false, false)) // Disallow whitespace in case option string ends in naked "E".
			{
				// Pure numbers are assumed to be style additions or removals:
				DWORD given_exstyle = ATOU(next_option); // ATOU() for unsigned.
				if (adding)
					mExStyle |= given_exstyle;
				else
					mExStyle &= ~given_exstyle;
			}
		}

		else // Handle things that are more general than the above, such as single letter options and pure numbers:
		{
			if (IsPureNumeric(next_option)) // Above has already verified that *next_option can't be whitespace.
			{
				// Pure numbers are assumed to be style additions or removals:
				DWORD given_style = ATOU(next_option); // ATOU() for unsigned.
				if (adding)
					mStyle |= given_style;
				else
					mStyle &= ~given_style;
			}
		}

		// If the item was not handled by the above, ignore it because it is unknown.

		*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.

	} // for() each item in option list

	// Besides reducing the code size and complexity, another reason all changes to style are made
	// here rather than above is that multiple changes might have been made above to the style,
	// and there's no point in redrawing/updating the window for each one:
	if (mHwnd && (mStyle != style_orig || mExStyle != exstyle_orig))
	{
		// v1.0.27.01: Must do this prior to SetWindowLong() because sometimes SetWindowLong()
		// traumatizes the window (such as "Gui -Caption"), making it effectively invisible
		// even though its non-functional remnant is still on the screen:
		bool is_visible = IsWindowVisible(mHwnd) && !IsIconic(mHwnd);

		// Since window already exists but its style has changed, attempt to update it dynamically.
		if (mStyle != style_orig)
			SetWindowLong(mHwnd, GWL_STYLE, mStyle);
		if (mExStyle != exstyle_orig)
			SetWindowLong(mHwnd, GWL_EXSTYLE, mExStyle);

		if (is_visible)
		{
			// Hiding then showing is the only way I've discovered to make it update.  If the window
			// is not updated, a strange effect occurs where the window is still visible but can no
			// longer be used at all (clicks pass right through it).  This show/hide method is less
			// desirable due to possible side effects caused to any script that happens to be watching
			// for its existence/non-existence, so it would be nice if some better way can be discovered
			// to do this.
			// SetWindowPos is also necessary, otherwise the frame thickness entirely around the window
			// does not get updated (just parts of it):
			SetWindowPos(mHwnd, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
			ShowWindow(mHwnd, SW_HIDE);
			ShowWindow(mHwnd, SW_SHOWNA); // i.e. don't activate it if it wasn't before. Note that SW_SHOWNA avoids restoring the window if it is currently minimized or maximized (unlike SW_SHOWNOACTIVATE).
			// None of the following methods alone is enough, at least not when the window is currently active:
			// 1) InvalidateRect(mHwnd, NULL, TRUE);
			// 2) SendMessage(mHwnd, WM_NCPAINT, 1, 0);  // 1 = Repaint entire frame.
			// 3) RedrawWindow(mHwnd, NULL, NULL, RDW_INVALIDATE|RDW_FRAME|RDW_UPDATENOW);
			// 4) SetWindowPos(mHwnd, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
		}
		// Continue on to create the window so that code is simplified in other places by
		// using the assumption that "if gui[i] object exists, so does its window".
		// Another important reason this is done is that if an owner window were to be destroyed
		// before the window it owns is actually created, the WM_DESTROY logic would have to check
		// for any windows owned by the window being destroyed and update them.
	}

	return OK;
}



ResultType GuiType::ControlParseOptions(char *aOptions, GuiControlOptionsType &aOpt, GuiControlType &aControl
	, GuiIndexType aControlIndex)
// Caller must have already initialized aOpt with zeroes or any other desired starting values.
// Caller must ensure that aOptions is a modifiable string, since this method temporarily alters it.
{
	char *next_option, *option_end, orig_char;
	bool adding; // Whether this option is beeing added (+) or removed (-).
	GuiControlType *tab_control;
	RECT rect;
	POINT pt;

	for (next_option = aOptions; *next_option; next_option = omit_leading_whitespace(option_end))
	{
		if (*next_option == '-')
		{
			adding = false;
			// omit_leading_whitespace() is not called, which enforces the fact that the option word must
			// immediately follow the +/- sign.  This is done to allow the flexibility to have options
			// omit the plus/minus sign, and also to reserve more flexibility for future option formats.
			++next_option;  // Point it to the option word itself.
		}
		else
		{
			// Assume option is being added in the absence of either sign.  However, when we were
			// called by GuiControl(), the first option in the list must begin with +/- otherwise the cmd
			// would never have been properly detected as GUICONTROL_CMD_OPTIONS in the first place.
			adding = true;
			if (*next_option == '+')
				++next_option;  // Point it to the option word itself.
			//else do not increment, under the assumption that the plus has been omitted from a valid
			// option word and is thus an implicit plus.
		}

		if (!*next_option) // In case the entire option string ends in a naked + or -.
			break;
		// Find the end of this option item:
		if (   !(option_end = StrChrAny(next_option, " \t"))   )  // Space or tab.
			option_end = next_option + strlen(next_option); // Set to position of zero terminator instead.
		if (option_end == next_option)
			continue; // i.e. the string contains a + or - with a space or tab after it, which is intentionally ignored.

		// Temporarily terminate to help eliminate ambiguity for words contained inside other words,
		// such as "Checked" inside of "CheckedGray":
		orig_char = *option_end;
		*option_end = '\0';

		// Attributes:
		if (!stricmp(next_option, "Section")) // Adding and removing are treated the same in this case.
			aOpt.start_new_section = true;    // Ignored by caller when control already exists.
		else if (!stricmp(next_option, "AltSubmit"))
			if (adding) aControl.attrib |= GUI_CONTROL_ATTRIB_ALTSUBMIT; else aControl.attrib &= ~GUI_CONTROL_ATTRIB_ALTSUBMIT;

		// Content of control (these are currently only effective if the control is being newly created):
		else if ((aControl.type == GUI_CONTROL_CHECKBOX || aControl.type == GUI_CONTROL_RADIO)
			&& !strnicmp(next_option, "Checked", 7))
		{
			// As of v1.0.26, Checked/Hidden/Disabled can be followed by an optional 1/0/-1 so that
			// there is a way for a script to set the starting state by reading from an INI or registry
			// entry that contains 1 or 0 instead of needing the literal word "checked" stored in there.
			// Otherwise, a script would have to do something like the following before every "Gui Add":
			// if Box1Enabled
			//    Enable = Enabled
			// else
			//    Enable =
			// Gui Add, checkbox, %Enable%, My checkbox.
			if (next_option[7]) // There's more after the word, namely a 1, 0, or -1.
			{
				aOpt.checked = ATOI(next_option + 7);
				if (aOpt.checked == -1)
					aOpt.checked = BST_INDETERMINATE;
			}
			else
				if (adding) aOpt.checked = BST_CHECKED; else aOpt.checked = BST_UNCHECKED;
		}
		else if (aControl.type == GUI_CONTROL_CHECKBOX && !stricmp(next_option, "CheckedGray")) // Radios can't have the 3rd/gray state.
			if (adding) aOpt.checked = BST_INDETERMINATE; else aOpt.checked = BST_UNCHECKED;
		else if (!strnicmp(next_option, "Choose", 6)) // Caller should ignore aOpt.choice if it isn't applicable for this control type.
		{
			// "CHOOSE" provides an easier way to conditionally select a different item at the time
			// the control is added.  Example: gui, add, ListBox, vMyList Choose%choice%, %MyItemList%
			if (adding)
			{
				aOpt.choice = ATOI(next_option + 6);
				if (aOpt.choice < 1) // Invalid: number should be 1 or greater.
					aOpt.choice = 0; // Flag it as invalid.
			}
			//else do nothing (not currently implemented)
		}

		// Styles (general):
		else if (!stricmp(next_option, "Border"))
			if (adding) aOpt.style_add |= WS_BORDER; else aOpt.style_remove |= WS_BORDER;
		else if (!stricmp(next_option, "VScroll")) // Seems harmless in this case not to check aControl.type to ensure it's an input-capable control.
			if (adding) aOpt.style_add |= WS_VSCROLL; else aOpt.style_remove |= WS_VSCROLL;
		else if (!strnicmp(next_option, "HScroll", 7)) // Seems harmless in this case not to check aControl.type to ensure it's an input-capable control.
		{
			if (adding)
			{
				// MSDN: "To respond to the LB_SETHORIZONTALEXTENT message, the list box must have
				// been defined with the WS_HSCROLL style."
				aOpt.style_add |= WS_HSCROLL;
				next_option += 7;
				aOpt.hscroll_pixels = *next_option ? ATOI(next_option) : -1;  // -1 signals it to use a default based on control's width.
			}
			else
				aOpt.style_remove |= WS_HSCROLL;
		}
		else if (!stricmp(next_option, "Tabstop")) // Seems harmless in this case not to check aControl.type to ensure it's an input-capable control.
			if (adding) aOpt.style_add |= WS_TABSTOP; else aOpt.style_remove |= WS_TABSTOP;
		else if (!stricmp(next_option, "NoTab")) // Supported for backward compatibility and it might be more ergonomic for "Gui Add".
			if (adding) aOpt.style_remove |= WS_TABSTOP; else aOpt.style_add |= WS_TABSTOP;
		else if (!stricmp(next_option, "Group")) // This overlaps with g-label, but seems well worth it in this case.
			if (adding) aOpt.style_add |= WS_GROUP; else aOpt.style_remove |= WS_GROUP;
		else if (!strnicmp(next_option, "Disabled", 8))
		{
			// As of v1.0.26, Checked/Hidden/Disabled can be followed by an optional 1/0/-1 so that
			// there is a way for a script to set the starting state by reading from an INI or registry
			// entry that contains 1 or 0 instead of needing the literal word "checked" stored in there.
			// Otherwise, a script would have to do something like the following before every "Gui Add":
			// if Box1Enabled
			//    Enable = Enabled
			// else
			//    Enable =
			// Gui Add, checkbox, %Enable%, My checkbox.
			if (next_option[8] && !ATOI(next_option + 8)) // If it's Disabled0, invert the mode to become "enabled".
				adding = !adding;
			if (aControl.hwnd) // More correct to call EnableWindow and let it set the style.  Do not set the style explicitly in this case since that might break it.
				EnableWindow(aControl.hwnd, adding ? FALSE : TRUE);
			else
				if (adding) aOpt.style_add |= WS_DISABLED; else aOpt.style_remove |= WS_DISABLED;
		}
		else if (!strnicmp(next_option, "Hidden", 6))
		{
			// As of v1.0.26, Checked/Hidden/Disabled can be followed by an optional 1/0/-1 so that
			// there is a way for a script to set the starting state by reading from an INI or registry
			// entry that contains 1 or 0 instead of needing the literal word "checked" stored in there.
			// Otherwise, a script would have to do something like the following before every "Gui Add":
			// if Box1Enabled
			//    Enable = Enabled
			// else
			//    Enable =
			// Gui Add, checkbox, %Enable%, My checkbox.
			if (next_option[6] && !ATOI(next_option + 6)) // If it's Hidden0, invert the mode to become "show".
				adding = !adding;
			if (aControl.hwnd) // More correct to call ShowWindow() and let it set the style.  Do not set the style explicitly in this case since that might break it.
				ShowWindow(aControl.hwnd, adding ? SW_HIDE : SW_SHOWNOACTIVATE);
			else
				if (adding) aOpt.style_remove |= WS_VISIBLE; else aOpt.style_add |= WS_VISIBLE;
		}
		else if (!stricmp(next_option, "Wrap"))
		{
			switch(aControl.type)
			{
			case GUI_CONTROL_TEXT: // This one is a little tricky but the below should be appropriate in most cases:
				if (adding) aOpt.style_remove |= 0x0F; else aOpt.style_add |= SS_LEFTNOWORDWRAP;
				break;
			case GUI_CONTROL_GROUPBOX:
			case GUI_CONTROL_BUTTON:
			case GUI_CONTROL_CHECKBOX:
			case GUI_CONTROL_RADIO:
				if (adding) aOpt.style_add |= BS_MULTILINE; else aOpt.style_remove |= BS_MULTILINE;
				break;
			case GUI_CONTROL_EDIT: // Must be a multi-line now or shortly in the future or these will have no effect.
				if (adding) aOpt.style_remove |= WS_HSCROLL|ES_AUTOHSCROLL; else aOpt.style_add |= ES_AUTOHSCROLL;
				// WS_HSCROLL is removed because with it, wrapping is automatically off.
				break;
			case GUI_CONTROL_TAB:
				if (adding) aOpt.style_add |= TCS_MULTILINE; else aOpt.style_remove |= TCS_MULTILINE;
				// WS_HSCROLL is removed because with it, wrapping is automatically off.
				break;
			// N/A for these:
			//case GUI_CONTROL_PIC:
			//case GUI_CONTROL_DROPDOWNLIST:
			//case GUI_CONTROL_COMBOBOX:
			//case GUI_CONTROL_LISTBOX:
			}
		}
		else if (!strnicmp(next_option, "Background", 10))
		{
			next_option += 10;  // To help maintainability, point it to the optional suffix here.
			if (aControl.type == GUI_CONTROL_PROGRESS)
			{
				// Note that GUI_CONTROL_ATTRIB_BACKGROUND_DEFAULT and GUI_CONTROL_ATTRIB_BACKGROUND_TRANS
				// don't apply to Progress controls because the window proc never receives CTLCOLOR messages
				// for them.
				if (adding)
				{
					aOpt.progress_color_bk = ColorNameToBGR(next_option);
					if (aOpt.progress_color_bk == CLR_NONE) // A matching color name was not found, so assume it's in hex format.
						// It seems strtol() automatically handles the optional leading "0x" if present:
						aOpt.progress_color_bk = rgb_to_bgr(strtol(next_option, NULL, 16));
						// if next_option did not contain something hex-numeric, black (0x00) will be assumed,
						// which seems okay given how rare such a problem would be.
				}
				else // Removing
					aOpt.progress_color_bk = CLR_DEFAULT;
			}
			else
			{
				if (adding)
				{
					aControl.attrib &= ~GUI_CONTROL_ATTRIB_BACKGROUND_DEFAULT;
					if (!stricmp(next_option, "Trans"))
						aControl.attrib |= GUI_CONTROL_ATTRIB_BACKGROUND_TRANS; // This is mutually exclusive of the above anyway.
					else
						aControl.attrib &= ~GUI_CONTROL_ATTRIB_BACKGROUND_TRANS;
					// In the future, something like the below can help support background colors for individual controls.
					//COLORREF background_color = ColorNameToBGR(next_option + 10);
					//if (background_color == CLR_NONE) // A matching color name was not found, so assume it's in hex format.
					//	// It seems strtol() automatically handles the optional leading "0x" if present:
					//	background_color = rgb_to_bgr(strtol(next_option, NULL, 16));
					//	// if next_option did not contain something hex-numeric, black (0x00) will be assumed,
					//	// which seems okay given how rare such a problem would be.
				}
				else
				{
					// Note that "-BackgroundTrans" is not supported, since Trans is considered to be
					// a color value for the purpose of expanding this feature in the future to support
					// custom background colors on a per-control basis.  In other words, the trans factor
					// can be turned off by using "-Background" or "+BackgroundBlue", etc.
					aControl.attrib &= ~GUI_CONTROL_ATTRIB_BACKGROUND_TRANS;
					aControl.attrib |= GUI_CONTROL_ATTRIB_BACKGROUND_DEFAULT;
				}
			}
		}

		// Picture
		else if (!strnicmp(next_option, "Icon", 4)) // Caller should ignore aOpt.icon_number if it isn't applicable for this control type.
		{
			if (adding)
				aOpt.icon_number = ATOI(next_option + 4);
			//else do nothing (not currently implemented)
		}

		// Button
		else if (aControl.type == GUI_CONTROL_BUTTON && !stricmp(next_option, "Default"))
			if (adding) aOpt.style_add |= BS_DEFPUSHBUTTON; else aOpt.style_remove |= BS_DEFPUSHBUTTON;
		else if (aControl.type == GUI_CONTROL_CHECKBOX && !stricmp(next_option, "Check3")) // Radios can't have the 3rd/gray state.
			if (adding) aOpt.style_add |= BS_AUTO3STATE; else aOpt.style_remove |= BS_AUTO3STATE;

		// Edit (and upper/lowercase for combobox/ddl, and others)
		else if (!stricmp(next_option, "ReadOnly"))
		{
			if (aControl.type == GUI_CONTROL_EDIT)
				if (adding) aOpt.style_add |= ES_READONLY; else aOpt.style_remove |= ES_READONLY;
			else if (aControl.type == GUI_CONTROL_LISTBOX)
				if (adding) aOpt.style_add |= LBS_NOSEL; else aOpt.style_remove |= LBS_NOSEL;
		}
		else if (!stricmp(next_option, "Multi"))
		{
			// It was named "multi" vs. multiline and/or "MultiSel" because it seems easier to
			// remember in these cases.  In fact, any time two styles can be combined into one
			// name whose actual function depends on the control type, it seems likely to make
			// things easier to remember.
			if (aControl.type == GUI_CONTROL_EDIT)
				if (adding) aOpt.style_add |= ES_MULTILINE; else aOpt.style_remove |= ES_MULTILINE;
			else if (aControl.type == GUI_CONTROL_LISTBOX)
				if (adding) aOpt.style_add |= LBS_EXTENDEDSEL; else aOpt.style_remove |= LBS_EXTENDEDSEL;
		}
		else if (aControl.type == GUI_CONTROL_EDIT && !stricmp(next_option, "WantReturn"))
			if (adding) aOpt.style_add |= ES_WANTRETURN; else aOpt.style_remove |= ES_WANTRETURN;
		else if (aControl.type == GUI_CONTROL_EDIT && !stricmp(next_option, "Number"))
			if (adding) aOpt.style_add |= ES_NUMBER; else aOpt.style_remove |= ES_NUMBER;
		else if (!stricmp(next_option, "Lowercase"))
		{
			if (aControl.type == GUI_CONTROL_EDIT)
				if (adding) aOpt.style_add |= ES_LOWERCASE; else aOpt.style_remove |= ES_LOWERCASE;
			else if (aControl.type == GUI_CONTROL_COMBOBOX || aControl.type == GUI_CONTROL_DROPDOWNLIST)
				if (adding) aOpt.style_add |= CBS_LOWERCASE; else aOpt.style_remove |= CBS_LOWERCASE;
		}
		else if (!stricmp(next_option, "Uppercase"))
		{
			if (aControl.type == GUI_CONTROL_EDIT)
				if (adding) aOpt.style_add |= ES_UPPERCASE; else aOpt.style_remove |= ES_UPPERCASE;
			else if (aControl.type == GUI_CONTROL_COMBOBOX || aControl.type == GUI_CONTROL_DROPDOWNLIST)
				if (adding) aOpt.style_add |= CBS_UPPERCASE; else aOpt.style_remove |= CBS_UPPERCASE;
		}
		else if (aControl.type == GUI_CONTROL_EDIT && !strnicmp(next_option, "Password", 8))
		{
			// Allow a space to be the masking character, since it's conceivable that might
			// be wanted in cases where someone doesn't wany anyone to know they're typing a password.
			// Simplest to assign unconditionally, regardless of whether adding or removing:
			aOpt.password_char = next_option[8];  // Can be '\0', which indicates "use OS default".
			if (adding)
			{
				aOpt.style_add |= ES_PASSWORD;
				if (aControl.hwnd) // Update the existing edit.
				{
					// Don't know how to achieve the black circle on XP *after* the control has
					// been created.  Maybe it's impossible.  Thus, provide default since otherwise
					// pass-char will be removed vs. added:
					if (!aOpt.password_char)
						aOpt.password_char = '*';
					SendMessage(aControl.hwnd, EM_SETPASSWORDCHAR, (WPARAM)aOpt.password_char, 0);
				}
			}
			else
			{
				aOpt.style_remove |= ES_PASSWORD;
				if (aControl.hwnd) // Update the existing edit.
					SendMessage(aControl.hwnd, EM_SETPASSWORDCHAR, 0, 0);
			}
		}
		else if (!strnicmp(next_option, "Limit", 5)) // This is used for Hotkey controls also.
		{
			if (adding)
			{
				next_option += 5;
				aOpt.limit = *next_option ? ATOI(next_option) : -1;  // -1 signals it to limit input to visible width of field.
				// aOpt.limit will later be ignored for some control types.
			}
			else
				aOpt.limit = INT_MIN; // Signal it to remove the limit.
		}

		// Combo/DropDownList/ListBox
		else if (aControl.type == GUI_CONTROL_COMBOBOX && !stricmp(next_option, "Simple")) // DDL is not equipped to handle this style.
			if (adding) aOpt.style_add |= CBS_SIMPLE; else aOpt.style_remove |= CBS_SIMPLE;
		else if (!stricmp(next_option, "Sort"))
		{
			if (aControl.type == GUI_CONTROL_LISTBOX)
				if (adding) aOpt.style_add |= LBS_SORT; else aOpt.style_remove |= LBS_SORT;
			else if (aControl.type == GUI_CONTROL_COMBOBOX || aControl.type == GUI_CONTROL_DROPDOWNLIST)
				if (adding) aOpt.style_add |= CBS_SORT; else aOpt.style_remove |= CBS_SORT;
		}

		// Slider
		else if (aControl.type == GUI_CONTROL_SLIDER && !stricmp(next_option, "Invert")) // Not called "Reverse" to avoid confusion with the non-functional style of that name.
			if (adding) aControl.attrib |= GUI_CONTROL_ATTRIB_ALTBEHAVIOR; else aControl.attrib &= ~GUI_CONTROL_ATTRIB_ALTBEHAVIOR;
		else if (aControl.type == GUI_CONTROL_SLIDER && !stricmp(next_option, "NoTicks"))
			if (adding) aOpt.style_add |= TBS_NOTICKS; else aOpt.style_remove |= TBS_NOTICKS;
		else if (aControl.type == GUI_CONTROL_SLIDER && !strnicmp(next_option, "TickInterval", 12))
		{
			if (adding)
			{
				aOpt.style_add |= TBS_AUTOTICKS;
				aOpt.tick_interval = ATOI(next_option + 12);
			}
			else
			{
				aOpt.style_remove |= TBS_AUTOTICKS;
				aOpt.tick_interval = -1;  // Signal it to remove the ticks later below (if the window exists).
			}
		}
		else if (aControl.type == GUI_CONTROL_SLIDER && !strnicmp(next_option, "Line", 4))
		{
			if (adding)
				aOpt.line_size = ATOI(next_option + 4);
			//else removal not supported.
		}
		else if (aControl.type == GUI_CONTROL_SLIDER && !strnicmp(next_option, "Page", 4))
		{
			if (adding)
				aOpt.page_size = ATOI(next_option + 4);
			//else removal not supported.
		}
		else if (aControl.type == GUI_CONTROL_SLIDER && !strnicmp(next_option, "Thick", 5))
		{
			if (adding)
			{
				aOpt.style_add |= TBS_FIXEDLENGTH;
				aOpt.thickness = ATOI(next_option + 5);
			}
			else // Removing the style is enough to reset its appearance on both XP Theme and Classic Theme.
				aOpt.style_remove |= TBS_FIXEDLENGTH;
		}
		else if (aControl.type == GUI_CONTROL_SLIDER && !strnicmp(next_option, "ToolTip", 7))
		{
			if (adding)
			{
				aOpt.tip_side = -1;  // Set default.
				switch(toupper(next_option[7]))
				{
				case 'T': aOpt.tip_side = TBTS_TOP; break;
				case 'L': aOpt.tip_side = TBTS_LEFT; break;
				case 'B': aOpt.tip_side = TBTS_BOTTOM; break;
				case 'R': aOpt.tip_side = TBTS_RIGHT; break;
				}
				if (aOpt.tip_side < 0)
					aOpt.tip_side = 0; // Restore to the value that means "use default side".
				else
					++aOpt.tip_side; // Offset by 1, since zero is reserved as "use default side".
				aOpt.style_add |= TBS_TOOLTIPS;
			}
			else
				aOpt.style_remove |= TBS_TOOLTIPS;
		}
		else if (aControl.type == GUI_CONTROL_SLIDER && !strnicmp(next_option, "Buddy", 5))
		{
			if (adding)
			{
				next_option += 5;
				char which_buddy = *next_option;
				if (which_buddy) // i.e. it's not the zero terminator
				{
					++next_option; // Now it should point to the variable name of the buddy control.
					Var *var = g_script.FindVar(next_option); // Not FindOrAdd() in this case.
					if (var)
					{
						// Below relies on GuiIndexType underflow:
						for (GuiIndexType u = mControlCount - 1; u < mControlCount; --u) // Search in reverse for better avg-case performance.
							if (mControl[u].output_var == var)
								if (which_buddy == '1')
									aOpt.buddy1 = &mControl[u];
								else // assume '2'
									aOpt.buddy2 = &mControl[u];
					}
				}
			}
			//else removal not supported.
		}

		// Progress and Slider
		else if (!stricmp(next_option, "Vertical"))
		{
			// Seems best not to recognize Vertical for Tab controls since Left and Right
			// already cover it very well.
			if (aControl.type == GUI_CONTROL_SLIDER)
				if (adding) aOpt.style_add |= TBS_VERT; else aOpt.style_remove |= TBS_VERT;
			else if (aControl.type == GUI_CONTROL_PROGRESS)
				if (adding) aOpt.style_add |= PBS_VERTICAL; else aOpt.style_remove |= PBS_VERTICAL;
			//else do nothing, not a supported type
		}
		else if (!strnicmp(next_option, "Range", 5)) // Caller should ignore aOpt.range_min/max if it isn't applicable for this control type.
		{
			if (adding)
			{
				next_option += 5; // Helps with omitting the first minus sign, if any, below.
				aOpt.range_min = ATOI(next_option);
				char *cp = strchr(next_option + 1, '-');  // +1 to omit the min's minus sign, if it has one.
				if (cp)
					aOpt.range_max = ATOI(cp + 1);
			}
			//else do nothing (not currently implemented)
		}

		// Progress
		else if (aControl.type == GUI_CONTROL_PROGRESS && !stricmp(next_option, "Smooth"))
			if (adding) aOpt.style_add |= PBS_SMOOTH; else aOpt.style_remove |= PBS_SMOOTH;

		// Tab control
		else if (aControl.type == GUI_CONTROL_TAB && !stricmp(next_option, "Buttons"))
			if (adding) aOpt.style_add |= TCS_BUTTONS; else aOpt.style_remove |= TCS_BUTTONS;
		else if (aControl.type == GUI_CONTROL_TAB && !stricmp(next_option, "Bottom"))
			if (adding)
			{
				aOpt.style_add |= TCS_BOTTOM;
				aOpt.style_remove |= TCS_VERTICAL;
			}
			else
				aOpt.style_remove |= TCS_BOTTOM;

		// Styles (alignment/justification):
		else if (!stricmp(next_option, "Center"))
			if (adding)
			{
				switch (aControl.type)
				{
				case GUI_CONTROL_SLIDER:
					if (adding) aOpt.style_add |= TBS_BOTH;
					aOpt.style_remove |= TBS_LEFT;
					break;
				case GUI_CONTROL_TEXT:
					aOpt.style_add |= SS_CENTER;
					aOpt.style_remove |= SS_RIGHT; // Mutually exclusive since together they are invalid.
					break;
				case GUI_CONTROL_GROUPBOX: // Changes alignment of its label.
				case GUI_CONTROL_BUTTON:   // Probably has no effect in this case, since it's centered by default?
				case GUI_CONTROL_CHECKBOX: // Puts gap between box and label.
				case GUI_CONTROL_RADIO:
					aOpt.style_add |= BS_CENTER;
					// But don't remove BS_LEFT or BS_RIGHT since BS_CENTER is defined as a combination of them.
					break;
				case GUI_CONTROL_EDIT:
					aOpt.style_add |= ES_CENTER;
					aOpt.style_remove |= ES_RIGHT; // Mutually exclusive since together they are (probably) invalid.
					break;
				// Not applicable for:
				//case GUI_CONTROL_PIC: SS_CENTERIMAGE is currently not used due to auto-pic-scaling/fitting.
				//case GUI_CONTROL_DROPDOWNLIST:
				//case GUI_CONTROL_COMBOBOX:
				//case GUI_CONTROL_LISTBOX:
				}
			}
			else // Removing.
			{
				switch (aControl.type)
				{
				case GUI_CONTROL_SLIDER:
					aOpt.style_remove |= TBS_BOTH;
					break;
				case GUI_CONTROL_TEXT:
					aOpt.style_remove |= SS_CENTER; // Seems okay since SS_ICON shouldn't be present for this control type.
					break;
				case GUI_CONTROL_GROUPBOX:
				case GUI_CONTROL_BUTTON:
				case GUI_CONTROL_CHECKBOX:
				case GUI_CONTROL_RADIO:
					// BS_CENTER is a tricky one since it is a combination of BS_LEFT and BS_RIGHT.
					// If the control exists and has either BS_LEFT or BS_RIGHT (but not both), do
					// nothing:
					if (aControl.hwnd)
					{
						if (GetWindowLong(aControl.hwnd, GWL_STYLE) & BS_CENTER) // i.e. it has both BS_LEFT and BS_RIGHT
							aOpt.style_remove |= BS_CENTER;
						//else nothing needs to be done.
					}
					else
						if (aOpt.style_add & BS_CENTER) // i.e. Both BS_LEFT and BS_RIGHT are set to be added.
							aOpt.style_add &= ~BS_CENTER; // Undo it, which later helps avoid the need to apply style_add prior to style_remove.
						//else nothing needs to be done.
					break;
				case GUI_CONTROL_EDIT:
					aOpt.style_remove |= ES_CENTER;
					break;
				// Not applicable for:
				//case GUI_CONTROL_PIC: SS_CENTERIMAGE is currently not used due to auto-pic-scaling/fitting.
				//case GUI_CONTROL_DROPDOWNLIST:
				//case GUI_CONTROL_COMBOBOX:
				//case GUI_CONTROL_LISTBOX:
				}
			}

		else if (!stricmp(next_option, "Right"))
			if (adding)
			{
				switch (aControl.type)
				{
				case GUI_CONTROL_SLIDER:
					aOpt.style_remove |= TBS_LEFT|TBS_BOTH;
					break;
				case GUI_CONTROL_TEXT:
					aOpt.style_add |= SS_RIGHT;
					aOpt.style_remove |= SS_CENTER; // Mutually exclusive since together they are invalid.
					break;
				case GUI_CONTROL_GROUPBOX:
				case GUI_CONTROL_BUTTON:
				case GUI_CONTROL_CHECKBOX:
				case GUI_CONTROL_RADIO:
					aOpt.style_add |= BS_RIGHT;
					// Doing this indirectly removes BS_CENTER, and does it in a way that makes unimportant
					// the order in which style_add and style_remove are applied later:
					aOpt.style_remove |= BS_LEFT;
					// And by default, put button itself to the right of its label since that seems
					// likely to be far more common/desirable (there can be a more obscure option
					// later to change this default):
					if (aControl.type == GUI_CONTROL_CHECKBOX || aControl.type == GUI_CONTROL_RADIO)
						aOpt.style_add |= BS_RIGHTBUTTON;
					break;
				case GUI_CONTROL_EDIT:
					aOpt.style_add |= ES_RIGHT;
					aOpt.style_remove |= ES_CENTER; // Mutually exclusive since together they are (probably) invalid.
					break;
				case GUI_CONTROL_TAB:
					aOpt.style_add |= TCS_VERTICAL|TCS_MULTILINE|TCS_RIGHT;
					break;
				// Not applicable for:
				//case GUI_CONTROL_PIC: SS_RIGHTJUST is currently not used due to auto-pic-scaling/fitting.
				//case GUI_CONTROL_DROPDOWNLIST:
				//case GUI_CONTROL_COMBOBOX:
				//case GUI_CONTROL_LISTBOX:
				}
			}
			else // Removing.
			{
				switch (aControl.type)
				{
				case GUI_CONTROL_SLIDER:
					aOpt.style_add |= TBS_LEFT;
					aOpt.style_remove |= TBS_BOTH; // Debatable.
					break;
				case GUI_CONTROL_TEXT:
					aOpt.style_remove |= SS_RIGHT; // Seems okay since SS_ICON shouldn't be present for this control type.
					break;
				case GUI_CONTROL_GROUPBOX:
				case GUI_CONTROL_BUTTON:
				case GUI_CONTROL_CHECKBOX:
				case GUI_CONTROL_RADIO:
					// BS_RIGHT is a tricky one since it is included inside BS_CENTER.
					// Thus, if the control exists and has BS_CENTER, do nothing since
					// BS_RIGHT can't be in effect if BS_CENTER already is:
					if (aControl.hwnd)
					{
						if (!(GetWindowLong(aControl.hwnd, GWL_STYLE) & BS_CENTER))
							aOpt.style_remove |= BS_RIGHT;
					}
					else
						if (!(aOpt.style_add & BS_CENTER))
							aOpt.style_add &= ~BS_RIGHT;  // A little strange, but seems correct since control hasn't even been created yet.
						//else nothing needs to be done because BS_RIGHT is already in effect removed since
						//BS_CENTER makes BS_RIGHT impossible to manifest.
					break;
				case GUI_CONTROL_EDIT:
					aOpt.style_remove |= ES_RIGHT;
					break;
				case GUI_CONTROL_TAB:
					aOpt.style_remove |= TCS_VERTICAL|TCS_RIGHT;
					break;
				// Not applicable for:
				//case GUI_CONTROL_PIC: SS_RIGHTJUST is currently not used due to auto-pic-scaling/fitting.
				//case GUI_CONTROL_DROPDOWNLIST:
				//case GUI_CONTROL_COMBOBOX:
				//case GUI_CONTROL_LISTBOX:
				}
			}

		else if (!stricmp(next_option, "Left"))
		{
			if (adding)
			{
				switch (aControl.type)
				{
				case GUI_CONTROL_SLIDER:
					aOpt.style_add |= TBS_LEFT;
					aOpt.style_remove |= TBS_BOTH;
					break;
				case GUI_CONTROL_TEXT:
					aOpt.style_remove |= SS_RIGHT|SS_CENTER;  // Removing these exposes the default of 0, which is LEFT.
					break;
				case GUI_CONTROL_CHECKBOX:
				case GUI_CONTROL_GROUPBOX:
				case GUI_CONTROL_BUTTON:
				case GUI_CONTROL_RADIO:
					aOpt.style_add |= BS_LEFT;
					// Doing this indirectly removes BS_CENTER, and does it in a way that makes unimportant
					// the order in which style_add and style_remove are applied later:
					aOpt.style_remove |= BS_RIGHT;
					// And by default, put button itself to the left of its label since that seems
					// likely to be far more common/desirable (there can be a more obscure option
					// later to change this default):
					if (aControl.type == GUI_CONTROL_CHECKBOX || aControl.type == GUI_CONTROL_RADIO)
						aOpt.style_remove |= BS_RIGHTBUTTON;
					break;
				case GUI_CONTROL_EDIT:
					aOpt.style_remove |= ES_RIGHT|ES_CENTER;  // Removing these exposes the default of 0, which is LEFT.
					break;
				case GUI_CONTROL_TAB:
					aOpt.style_add |= TCS_VERTICAL|TCS_MULTILINE;
					aOpt.style_remove |= TCS_RIGHT;
					break;
				// Not applicable for:
				//case GUI_CONTROL_PIC: SS_CENTERIMAGE is currently not used due to auto-pic-scaling/fitting.
				//case GUI_CONTROL_DROPDOWNLIST:
				//case GUI_CONTROL_COMBOBOX:
				//case GUI_CONTROL_LISTBOX:
				}
			}
			else // Removing.
			{
				switch (aControl.type)
				{
				case GUI_CONTROL_SLIDER:
					aOpt.style_remove |= TBS_LEFT|TBS_BOTH; // Removing TBS_BOTH is debatable, but "-left" is pretty rare/obscure anyway.
					break;
				case GUI_CONTROL_GROUPBOX:
				case GUI_CONTROL_BUTTON:
				case GUI_CONTROL_CHECKBOX:
				case GUI_CONTROL_RADIO:
					// BS_LEFT is a tricky one since it is included inside BS_CENTER.
					// Thus, if the control exists and has BS_CENTER, do nothing since
					// BS_LEFT can't be in effect if BS_CENTER already is:
					if (aControl.hwnd)
					{
						if (!(GetWindowLong(aControl.hwnd, GWL_STYLE) & BS_CENTER))
							aOpt.style_remove |= BS_LEFT;
					}
					else
						if (!(aOpt.style_add & BS_CENTER))
							aOpt.style_add &= ~BS_LEFT;  // A little strange, but seems correct since control hasn't even been created yet.
						//else nothing needs to be done because BS_LEFT is already in effect removed since
						//BS_CENTER makes BS_LEFT impossible to manifest.
					break;
				case GUI_CONTROL_TAB:
					aOpt.style_remove |= TCS_VERTICAL;
					break;
				// Not applicable for these since their LEFT attributes are zero and thus cannot be removed:
				//case GUI_CONTROL_TEXT:
				//case GUI_CONTROL_PIC:
				//case GUI_CONTROL_DROPDOWNLIST:
				//case GUI_CONTROL_COMBOBOX:
				//case GUI_CONTROL_LISTBOX:
				//case GUI_CONTROL_EDIT:
				}
			}
		} // else if

		else
		{
			// THE BELOW SHOULD BE DONE LAST so that they don't steal phrases/words that should be detected
			// as option words above.  An existing example is H for Hidden (above) or Height (below).
			// Additional examples:
			// if "visible" and "resize" ever become valid option words, the below would otherwise wrongly
			// detect them as variable=isible and row_count=esize, respectively.

			if (IsPureNumeric(next_option)) // Above has already verified that *next_option can't be whitespace.
			{
				// Pure numbers are assumed to be style additions or removals:
				DWORD given_style = ATOU(next_option); // ATOU() for unsigned.
				if (adding) aOpt.style_add |= given_style; else aOpt.style_remove |= given_style;
				*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.
				continue;
			}

			++next_option;  // Above has already verified that next_option isn't the empty string.
			if (!*next_option)
			{
				// The option word consists of only one character, so ignore allow except the below
				// since mandatory arg should immediately follow it.  Example: An isolated letter H
				// should do nothing rather than cause the height to be set to zero.
				switch (toupper(next_option[-1]))
				{
				case 'C':
					if (!adding && aControl.type != GUI_CONTROL_PIC && aControl.union_color != CLR_DEFAULT)
					{
						aControl.union_color = CLR_DEFAULT; // i.e. treat "-C" as return to the default color.
						aOpt.color_changed = true;
					}
					break;
				case 'G':
					aControl.jump_to_label = NULL;
					break;
				case 'V':
					aControl.output_var = NULL;
					break;
				}
				*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.
				continue;
			}

			// Since above didn't "continue", there is text after the option letter, so take action accordingly.
			switch (toupper(next_option[-1]))
			{
			case 'G': // "Gosub" a label when this control is clicked or changed.
				// For reasons of potential future use and compatibility, don't allow subroutines to be
				// assigned to control types that have no present use for them.  Note: GroupBoxes do
				// no support click-detection anyway, even if the BS_NOTIFY style is given to them
				// (this has been verified twice):
				if (aControl.type == GUI_CONTROL_EDIT || aControl.type == GUI_CONTROL_GROUPBOX
					|| aControl.type == GUI_CONTROL_PROGRESS || aControl.type == GUI_CONTROL_HOTKEY)
					// If control's hwnd exists, we were called from a caller who wants ErrorLevel set
					// instead of a message displayed:
					return aControl.hwnd ? g_ErrorLevel->Assign(ERRORLEVEL_ERROR)
						: g_script.ScriptError("This control type should not have an associated subroutine."
							ERR_ABORT, next_option - 1);
				Label *candidate_label;
				if (   !(candidate_label = g_script.FindLabel(next_option))   )
				{
					// If there is no explicit label, fall back to a special action if one is available
					// for this keyword:
					if (!stricmp(next_option, "Cancel"))
						aControl.attrib |= GUI_CONTROL_ATTRIB_IMPLICIT_CANCEL;
					// When the below is added, it should probably be made mutually exclusive of the above, probably
					// by devoting two bits in the field for a total of three possible implicit actions (since the
					// fourth is reserved as 00 = no action):
					//else if (!stricmp(label_name, "Clear")) -->
					//	control.options |= GUI_CONTROL_ATTRIB_IMPLICIT_CLEAR;
					else // Since a non-special label was explicitly specified, it's an error that it can't be found.
						return aControl.hwnd ? g_ErrorLevel->Assign(ERRORLEVEL_ERROR)
							: g_script.ScriptError(ERR_CONTROLLABEL ERR_ABORT, next_option - 1);
				}
				// Apply the SS_NOTIFY style *only* if the control actually has an associated action.
				// This is because otherwise the control would steal all clicks for any other controls
				// drawn on top of it (e.g. a picture control with some edit fields drawn on top of it).
				// See comments in the creation of GUI_CONTROL_PIC for more info:
				if (aControl.type == GUI_CONTROL_TEXT || aControl.type == GUI_CONTROL_PIC)
					aOpt.style_add |= SS_NOTIFY;
				aControl.jump_to_label = candidate_label; // Will be NULL if something like gCancel (implicit was used).
				break;

			case 'T': // Tabstop (the kind that exists inside a multi-line edit control or ListBox).
				if (aOpt.tabstop_count < GUI_MAX_TABSTOPS)
					aOpt.tabstop[aOpt.tabstop_count++] = ATOU(next_option);
				if (aControl.type == GUI_CONTROL_LISTBOX)
					aOpt.style_add |= LBS_USETABSTOPS; // Required to allow the ListBox to respond to LB_SETTABSTOPS.
				//else ignore ones beyond the maximum.
				break;

			case 'V': // Variable
				// It seems best to allow an input-control to lack a variable, in which case its contents will be
				// lost when the form is closed (unless fetched beforehand with something like ControlGetText).
				// This is because it allows layout editors and other script generators to omit the variable
				// and yet still be able to generate a runnable script.
				Var *candidate_var;
				if (   !(candidate_var = g_script.FindOrAddVar(next_option))   )
					// For now, this is always a critical error that stops the current quasi-thread rather
					// than setting ErrorLevel (if ErrorLevel is called for).  This is because adding a
					// variable can cause one of any number of different errors to be displayed, and changing
					// all those functions to have a silent mode doesn't seem worth the trouble given how
					// rarely 1) a control needs to get a new variable; 2) that variable name is too long
					// or not valid.
					return FAIL;  // It already displayed the error (e.g. name too long). Existing var (if any) is retained.
				// Check if any other control (visible or not, to avoid the complexity of a hidden control
				// needing to be dupe-checked every time it becomes visible) on THIS gui window has the
				// same variable.  That's an error because not only doesn't it make sense to do that,
				// but it might be useful to uniquely identify a control by its variable name (when making
				// changes to it, etc.)  Note that if this is the first control being added, mControlCount
				// is now zero because this control has not yet actually been added.  That is why
				// "u < mControlCount" is used:
				GuiIndexType u;
				for (u = 0; u < mControlCount; ++u)
					if (mControl[u].output_var == candidate_var)
						return aControl.hwnd ? g_ErrorLevel->Assign(ERRORLEVEL_ERROR)
							: g_script.ScriptError("The same variable cannot be used for more than one control per window."
								ERR_ABORT, next_option - 1);
				aControl.output_var = candidate_var;
				break;

			case 'E':  // Extended style
				if (IsPureNumeric(next_option, false, false)) // Disallow whitespace in case option string ends in naked "E".
				{
					// Pure numbers are assumed to be style additions or removals:
					DWORD given_exstyle = ATOU(next_option); // ATOU() for unsigned.
					if (adding) aOpt.exstyle_add |= given_exstyle; else aOpt.exstyle_remove |= given_exstyle;
				}
				break;

			case 'C':  // Color
				if (aControl.type == GUI_CONTROL_PIC) // Don't trash the union's hbitmap member.
					break;
				COLORREF new_color;
				new_color = ColorNameToBGR(next_option);
				if (new_color == CLR_NONE) // A matching color name was not found, so assume it's in hex format.
					// It seems strtol() automatically handles the optional leading "0x" if present:
					new_color = rgb_to_bgr(strtol(next_option, NULL, 16));
					// if next_option did not contain something hex-numeric, black (0x00) will be assumed,
					// which seems okay given how rare such a problem would be.
				if (aControl.union_color != new_color)
				{
					aControl.union_color = new_color;
					aOpt.color_changed = true;
				}
				break;

			case 'W':
				if (toupper(*next_option) == 'P') // Use the previous control's value.
					aOpt.width = mPrevWidth + ATOI(next_option + 1);
				else
					aOpt.width = ATOI(next_option);
				break;

			case 'H':
				if (toupper(*next_option) == 'P') // Use the previous control's value.
					aOpt.height = mPrevHeight + ATOI(next_option + 1);
				else
					aOpt.height = ATOI(next_option);
				break;

			case 'X':
				if (*next_option == '+')
				{
					if (tab_control = FindTabControl(aControl.tab_control_index)) // Assign.
					{
						// Since this control belongs to a tab control and that tab control already exists,
						// Position it relative to the tab control's client area upper-left corner if this
						// is the first control on this particular tab/page:
						if (!GetControlCountOnTabPage(aControl.tab_control_index, aControl.tab_index))
						{
							pt = GetPositionOfTabClientArea(*tab_control);
							aOpt.x = pt.x + ATOI(next_option + 1);
							if (aOpt.y == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
								aOpt.y = pt.y + mMarginY;
							break;
						}
						// else fall through and do it the standard way.
					}
					// Since above didn't break, do it the standard way.
					aOpt.x = mPrevX + mPrevWidth + ATOI(next_option + 1);
					if (aOpt.y == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.y = mPrevY;  // Since moving in the X direction, retain the same Y as previous control.
				}
				// For the M and P sub-options, not that the +/- prefix is optional.  The number is simply
				// read in as-is (though the use of + is more self-documenting in this case than omitting
				// the sign entirely).
				else if (toupper(*next_option) == 'M') // Use the X margin
				{
					aOpt.x = mMarginX + ATOI(next_option + 1);
					if (aOpt.y == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.y = mMaxExtentDown + mMarginY;
				}
				else if (toupper(*next_option) == 'P') // Use the previous control's X position.
				{
					aOpt.x = mPrevX + ATOI(next_option + 1);
					if (aOpt.y == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.y = mPrevY;  // Since moving in the X direction, retain the same Y as previous control.
				}
				else if (toupper(*next_option) == 'S') // Use the saved X position
				{
					aOpt.x = mSectionX + ATOI(next_option + 1);
					if (aOpt.y == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.y = mMaxExtentDownSection + mMarginY;  // In this case, mMarginY is the padding between controls.
				}
				else
				{
					aOpt.x = ATOI(next_option);
					if (aOpt.y == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.y = mMaxExtentDown + mMarginY;
				}
				break;

			case 'Y':
				if (*next_option == '+')
				{
					if (tab_control = FindTabControl(aControl.tab_control_index)) // Assign.
					{
						// Since this control belongs to a tab control and that tab control already exists,
						// Position it relative to the tab control's client area upper-left corner if this
						// is the first control on this particular tab/page:
						if (!GetControlCountOnTabPage(aControl.tab_control_index, aControl.tab_index))
						{
							pt = GetPositionOfTabClientArea(*tab_control);
							aOpt.y = pt.y + ATOI(next_option + 1);
							if (aOpt.x == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
								aOpt.x = pt.x + mMarginX;
							break;
						}
						// else fall through and do it the standard way.
					}
					// Since above didn't break, do it the standard way.
					aOpt.y = mPrevY + mPrevHeight + ATOI(next_option + 1);
					if (aOpt.x == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.x = mPrevX;  // Since moving in the Y direction, retain the same X as previous control.
				}
				// For the M and P sub-options, not that the +/- prefix is optional.  The number is simply
				// read in as-is (though the use of + is more self-documenting in this case than omitting
				// the sign entirely).
				else if (toupper(*next_option) == 'M') // Use the Y margin
				{
					aOpt.y = mMarginY + ATOI(next_option + 1);
					if (aOpt.x == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.x = mMaxExtentRight + mMarginX;
				}
				else if (toupper(*next_option) == 'P') // Use the previous control's Y position.
				{
					aOpt.y = mPrevY + ATOI(next_option + 1);
					if (aOpt.x == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.x = mPrevX;  // Since moving in the Y direction, retain the same X as previous control.
				}
				else if (toupper(*next_option) == 'S') // Use the saved Y position
				{
					aOpt.y = mSectionY + ATOI(next_option + 1);
					if (aOpt.x == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.x = mMaxExtentRightSection + mMarginX; // In this case, mMarginX is the padding between controls.
				}
				else
				{
					aOpt.y = ATOI(next_option);
					if (aOpt.x == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.x = mMaxExtentRight + mMarginX;
				}
				break;

			case 'R': // The number of rows desired in the control.  Use ATOF() so that fractional rows are allowed.
				aOpt.row_count = (float)ATOF(next_option); // Don't need double precision.
				break;
			} // switch()
		} // Final "else" in the "else if" ladder.

		// If the item was not handled by the above, ignore it because it is unknown.

		*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.

	} // for() each item in option list

	// If the control has already been created, apply the new style and exstyle here, if any:
	if (aControl.hwnd)
	{
		DWORD current_style = GetWindowLong(aControl.hwnd, GWL_STYLE);
		DWORD new_style = (current_style | aOpt.style_add) & ~aOpt.style_remove;

		// Fix for v1.0.24:
		// Certain styles can't be applied with a simple bit-or.  The below section is a subset of
		// a similar section in AddControl() to make sure that these styles are propertly handled:
		switch (aControl.type)
		{
		case GUI_CONTROL_PIC:
			// Fixed for v1.0.25.11 to prevent SS_ICON from getting changed to SS_BITMAP:
			new_style = (new_style & ~0x0F) | (current_style & 0x0F); // Done to ensure the lowest four bits are pure.
			break;
		case GUI_CONTROL_GROUPBOX:
			// There doesn't seem to be any flexibility lost by forcing the buttons to be the right type,
			// and doing so improves maintainability and peace-of-mind:
			new_style = (new_style & ~BS_TYPEMASK) | BS_GROUPBOX;  // Force it to be the right type of button.
			break;
		case GUI_CONTROL_BUTTON:
			if (new_style & BS_DEFPUSHBUTTON)
				new_style = (new_style & ~BS_TYPEMASK) | BS_DEFPUSHBUTTON; // Done to ensure the lowest four bits are pure.
			else
				new_style &= ~BS_TYPEMASK;  // Force it to be the right type of button --> BS_PUSHBUTTON == 0
			break;
		case GUI_CONTROL_CHECKBOX:
			// Note: BS_AUTO3STATE and BS_AUTOCHECKBOX are mutually exclusive due to their overlap within
			// the bit field:
			if (new_style & BS_AUTO3STATE)
				new_style = (new_style & ~BS_TYPEMASK) | BS_AUTO3STATE; // Done to ensure the lowest four bits are pure.
			else
				new_style = (new_style & ~BS_TYPEMASK) | BS_AUTOCHECKBOX;  // Force it to be the right type of button.
			break;
		case GUI_CONTROL_RADIO:
			new_style = (new_style & ~BS_TYPEMASK) | BS_AUTORADIOBUTTON;  // Force it to be the right type of button.
			break;
		case GUI_CONTROL_DROPDOWNLIST:
			new_style |= CBS_DROPDOWNLIST;  // This works because CBS_DROPDOWNLIST == CBS_SIMPLE|CBS_DROPDOWN
			break;
		case GUI_CONTROL_COMBOBOX:
			if (new_style & CBS_SIMPLE) // i.e. CBS_SIMPLE has been added to the original default, so assume it is SIMPLE.
				new_style = (new_style & ~0x0F) | CBS_SIMPLE; // Done to ensure the lowest four bits are pure.
			else
				new_style = (new_style & ~0x0F) | CBS_DROPDOWN; // Done to ensure the lowest four bits are pure.
			break;
		// Nothing extra for these currently:
		//case GUI_CONTROL_LISTBOX: i.e. allow LBS_NOTIFY to be removed in case anyone really wants to do that.
		//case GUI_CONTROL_EDIT:
		//case GUI_CONTROL_TEXT:  Ensuring SS_BITMAP and such are absent seems too over-protective.
		//case GUI_CONTROL_HOTKEY:
		//case GUI_CONTROL_SLIDER:
		//case GUI_CONTROL_PROGRESS:
		//case GUI_CONTROL_TAB: i.e. allow WS_CLIPSIBLINGS to be removed (Rajat needs this) and also TCS_OWNERDRAWFIXED in case anyone really wants to.
		}

		// This needs to be done prior to applying the updated style since it sometimes adds
		// more style attributes:
		if (aOpt.limit) // A char length-limit was specified or de-specified for an edit/combo field.
		{
			// These styles are applied last so that multiline vs. singleline will already be resolved
			// and known, since all options have now been processed.
			if (aControl.type == GUI_CONTROL_EDIT)
			{
				// For the below, note that EM_LIMITTEXT == EM_SETLIMITTEXT.
				if (aOpt.limit < 0)
				{
					// Either limit to visible width of field, or remove existing limit.
					// In both cases, first remove the control's internal limit in case it
					// was something really small before:
					SendMessage(aControl.hwnd, EM_LIMITTEXT, 0, 0);
					// Limit > INT_MIN but less than zero is the signal to limit input length to visible
					// width of field. But it can only work if the edit isn't a multiline.
					if (aOpt.limit != INT_MIN && !(new_style & ES_MULTILINE))
						new_style &= ~(WS_HSCROLL|ES_AUTOHSCROLL); // Enable the limit-to-visible-width style.
				}
				else // greater than zero, since zero itself it checked in one of the enclosing IFs above.
					SendMessage(aControl.hwnd, EM_LIMITTEXT, aOpt.limit, 0); // Set a hard limit.
			}
			else if (aControl.type == GUI_CONTROL_HOTKEY)
			{
				if (aOpt.limit < 0) // This is the signal to remove any existing limit.
					SendMessage(aControl.hwnd, HKM_SETRULES, 0, 0);
				else // greater than zero, since zero itself it checked in one of the enclosing IFs above.
					SendMessage(aControl.hwnd, HKM_SETRULES, aOpt.limit, MAKELPARAM(HOTKEYF_CONTROL|HOTKEYF_ALT, 0));
					// Above must also specify Ctrl+Alt or some other default, otherwise the restriction will have
					// no effect.
			}
			// Altering the limit after the control exists appears to be ineffective, so this is commented out:
			//else if (aControl.type == GUI_CONTROL_COMBOBOX)
			//	// It has been verified that that EM_LIMITTEXT has no effect when sent directly
			//	// to a ComboBox hwnd; however, it might work if sent to its edit-child.
			//	// For now, a Combobox can only be limited to its visible width.  Later, there might
			//	// be a way to send a message to its child control to limit its width directly.
			//	if (aOpt.limit == INT_MIN) // remove existing limit
			//		new_style |= CBS_AUTOHSCROLL;
			//	else
			//		new_style &= ~CBS_AUTOHSCROLL;
			//	// i.e. SetWindowLong() cannot manifest the above style after the window exists.
		}

		bool style_change_ok;
		bool style_needed_changing = false; // Set default.

		if (current_style != new_style)
		{
			style_needed_changing = true; // Either style or exstyle is changing.
			style_change_ok = false; // Starting assumption.
			switch (aControl.type)
			{
			case GUI_CONTROL_BUTTON:
				// BM_SETSTYLE is much more likely to have an effect for buttons than SetWindowLong().
				SendMessage(mControl[mDefaultButtonIndex].hwnd, BM_SETSTYLE, (WPARAM)LOWORD(new_style)
					, MAKELPARAM(TRUE, 0)); // Redraw = yes, though it seems to be ineffective sometimes. It's probably smart enough not to do it if the window is hidden.
				if ((new_style & BS_DEFPUSHBUTTON) && !(current_style & BS_DEFPUSHBUTTON))
				{
					mDefaultButtonIndex = aControlIndex;
					// This will alter the control id received via WM_COMMAND when the user presses ENTER:
					SendMessage(mHwnd, DM_SETDEFID, (WPARAM)GUI_INDEX_TO_ID(mDefaultButtonIndex), 0);
				}
				else if (!(new_style & BS_DEFPUSHBUTTON) && (current_style & BS_DEFPUSHBUTTON))
				{
					// Remove the default button (rarely needed so that's why there is current no
					// "Gui, NoDefaultButton" command:
					mDefaultButtonIndex = -1;
					// This will alter the control id received via WM_COMMAND when the user presses ENTER:
					SendMessage(mHwnd, DM_SETDEFID, (WPARAM)IDOK, 0); // restore to default
				}
				break;

			case GUI_CONTROL_LISTBOX:
				if ((new_style & WS_HSCROLL) && !(current_style & WS_HSCROLL)) // Scroll bar being added.
				{
					if (aOpt.hscroll_pixels < 0) // Calculate a default based on control's width.
					{
						// Since horizontal scrollbar is relatively rarely used, no fancy method
						// such as calculating scrolling-width via LB_GETTEXTLEN and current font's
						// average width is used.
						GetWindowRect(aControl.hwnd, &rect);
						aOpt.hscroll_pixels = 3 * (rect.right - rect.left);
					}
					// If hscroll_pixels is now zero or smaller than the width of the control,
					// the scrollbar will not be shown:
					SendMessage(aControl.hwnd, LB_SETHORIZONTALEXTENT, (WPARAM)aOpt.hscroll_pixels, 0);
				}
				else if (!(new_style & WS_HSCROLL) && (current_style & WS_HSCROLL)) // Scroll bar being removed.
					SendMessage(aControl.hwnd, LB_SETHORIZONTALEXTENT, 0, 0);
				break;
			} // switch()
			SetLastError(0); // Prior to SetWindowLong(), as recommended by MSDN.
			// Call this even for buttons because BM_SETSTYLE only handles the LOWORD part of the style:
			if (SetWindowLong(aControl.hwnd, GWL_STYLE, new_style) || !GetLastError()) // This is the precise way to detect success according to MSDN.
			{
				// Even if it indicated success, sometimes it failed anyway.  Find out for sure:
				if (GetWindowLong(aControl.hwnd, GWL_STYLE) != current_style)
					style_change_ok = true; // Even a partial change counts as a success.
			}
		}

		DWORD current_exstyle = GetWindowLong(aControl.hwnd, GWL_EXSTYLE);
		DWORD new_exstyle = (current_exstyle | aOpt.exstyle_add) & ~aOpt.exstyle_remove;
		if (current_exstyle != new_exstyle)
		{
			if (!style_needed_changing)
			{
				style_needed_changing = true; // Either style or exstyle is changing.
				style_change_ok = false; // Starting assumption.
			}
			//ELSE don't change the value of style_change_ok because we want it to retain the value set by
			// the GWL_STYLE change above; i.e. a partial success on either style or exstyle counts as a full
			// success.
			SetLastError(0); // Prior to SetWindowLong(), as recommended by MSDN.
			if (SetWindowLong(aControl.hwnd, GWL_EXSTYLE, new_exstyle) || !GetLastError()) // This is the precise way to detect success according to MSDN.
			{
				// Even if it indicated success, sometimes it failed anyway.  Find out for sure:
				if (GetWindowLong(aControl.hwnd, GWL_EXSTYLE) != current_exstyle)
					style_change_ok = true; // Even a partial change counts as a success.
			}
		}

  		// Redrawing the controls is required in some cases, such as a checkbox losing its 3-state
		// style while it has a gray checkmark in it (which incidentally in this case only changes
		// the appearance of the control, not the internal stored value in this case).
		bool do_invalidate_rect = style_needed_changing && style_change_ok; // Set default.

		// Do the below only after applying the styles above since part of it requires that the style be
		// updated and applied above.
		switch (aControl.type)
		{
		case GUI_CONTROL_SLIDER:
			ControlSetSliderOptions(aControl, aOpt);
			if (aOpt.style_remove & TBS_TOOLTIPS)
				SendMessage(aControl.hwnd, TBM_SETTOOLTIPS, NULL, 0); // i.e. removing the TBS_TOOLTIPS style is not enough.
			break;
		case GUI_CONTROL_PROGRESS:
			ControlSetProgressOptions(aControl, aOpt, new_style);
			// Above strips theme if required by new options.  It also applies new colors.
			break;
		case GUI_CONTROL_EDIT:
			if (aOpt.tabstop_count)
			{
				SendMessage(aControl.hwnd, EM_SETTABSTOPS, aOpt.tabstop_count, (LPARAM)aOpt.tabstop);
				// MSDN: "If the application is changing the tab stops for text already in the edit control,
				// it should call the InvalidateRect function to redraw the edit control window."
				do_invalidate_rect = true; // Override the default.
			}
			break;
		}

		if (do_invalidate_rect)
			InvalidateRect(aControl.hwnd, NULL, TRUE); // Assume there's text in the control.

		if (style_needed_changing && !style_change_ok) // Override the default errorlevel set by our caller, GuiControl().
			g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	} // aControl.hwnd is not NULL

	return OK;
}



void GuiType::ControlInitOptions(GuiControlOptionsType &aOpt, GuiControlType &aControl)
// Not done as class to avoid code-size overhead of initializer list, etc.
{
	ZeroMemory(&aOpt, sizeof(GuiControlOptionsType));
	aOpt.x = aOpt.y = aOpt.width = aOpt.height = COORD_UNSPECIFIED;
	aOpt.progress_color_bk = aControl.hwnd ? CLR_INVALID : CLR_DEFAULT; // If control doesn't yet exist, default it to CLR_DEFAULT;
	// Above: If it stays unaltered, CLR_INVALID means "leave color as it is".  This is for
	// use with "GuiControl, +option" so that ControlSetProgressOptions() knows that
	// the "+/-Background" item was not among the options in the list.  The reason this is needed
	// for background color but not bar color is that bar_color is stored as a control attribute,
	// but to save memory, background color is not.  In addition, there is no good way to ask the
	// control what its background color currently is.
}



void GuiType::ControlAddContents(GuiControlType &aControl, char *aContent, int aChoice)
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
		msg_select = (GetWindowLong(aControl.hwnd, GWL_STYLE) & (LBS_EXTENDEDSEL|LBS_MULTIPLESEL))
			? LB_SETSEL : LB_SETCURSEL;
		break;
	case GUI_CONTROL_TAB:
		break;
	default:    // Do nothing for any other control type that doesn't require content to be added this way.
		return; // e.g. GUI_CONTROL_SLIDER, which the caller should handle.
	}

	bool temporarily_terminated;
	char *this_field, *next_field;
	LRESULT item_index;

	// For tab controls:
	int requested_tab_index = 0;
	TCITEM tci;
	tci.mask = TCIF_TEXT | TCIF_IMAGE;
	tci.iImage = -1; 

	// Check *this_field at the top too, in case list ends in delimiter.
	for (this_field = aContent; *this_field;)
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

		// Add the item:
		if (aControl.type == GUI_CONTROL_TAB)
		{
			if (requested_tab_index > MAX_TABS_PER_CONTROL - 1) // Unlikely, but indicate failure if so.
				item_index = -1;
			else
			{
				tci.pszText = this_field;
				item_index = TabCtrl_InsertItem(aControl.hwnd, requested_tab_index, &tci);
				if (item_index != -1) // item_index is used later below as an indicator of success.
					++requested_tab_index;
			}
		}
		else
			item_index = SendMessage(aControl.hwnd, msg_add, 0, (LPARAM)this_field); // In this case, ignore any errors, namely CB_ERR/LB_ERR and CB_ERRSPACE).
			// For the above, item_index must be retrieved and used as the item's index because it might
			// be different than expected if the control's SORT style is in effect.

		if (temporarily_terminated)
		{
			*next_field = '|';  // Restore the original char.
			++next_field;
			if (*next_field == '|')  // An item ending in two delimiters is a default (pre-selected) item.
			{
				if (item_index >= 0) // The item was successfully added.
				{
					if (aControl.type == GUI_CONTROL_TAB)
						// MSDN: "A tab control does not send a TCN_SELCHANGING or TCN_SELCHANGE notification message
						// when a tab is selected using the TCM_SETCURSEL message."
						TabCtrl_SetCurSel(aControl.hwnd, item_index);
					else if (msg_select == LB_SETSEL) // Multi-select box requires diff msg to have a cumulative effect.
						SendMessage(aControl.hwnd, msg_select, (WPARAM)TRUE, (LPARAM)item_index);
					else
						SendMessage(aControl.hwnd, msg_select, (WPARAM)item_index, 0);  // Select this item.
				}
				++next_field;  // Now this could be a third '|', which would in effect be an empty item.
				// It can also be the zero terminator if the list ends in a delimiter, e.g. item1|item2||
			}
		}
		this_field = next_field;
	} // for()

	// Have aChoice take precedence over any double-piped item(s) that appeared in the list:
	if (aChoice <= 0)
		return;
	--aChoice;

	if (aControl.type == GUI_CONTROL_TAB)
		// MSDN: "A tab control does not send a TCN_SELCHANGING or TCN_SELCHANGE notification message
		// when a tab is selected using the TCM_SETCURSEL message."
		TabCtrl_SetCurSel(aControl.hwnd, aChoice);
	else if (msg_select == LB_SETSEL) // Multi-select box requires diff msg to have a cumulative effect.
		SendMessage(aControl.hwnd, msg_select, (WPARAM)TRUE, (LPARAM)aChoice);
	else
		SendMessage(aControl.hwnd, msg_select, (WPARAM)aChoice, 0);  // Select this item.
}



ResultType GuiType::Show(char *aOptions, char *aText)
{
	if (!mHwnd)
		return OK;  // Make this a harmless attempt.

	// In the future, it seems best to rely on mShowIsInProgress to prevent the Window Proc from ever
	// doing a MsgSleep() to launch a script subroutine.  This is because if anything we do in this
	// function results in a launch of the Window Proc (such as MoveWindow and ShowWindow), our
	// activity here might be interrupted in a destructive way.  For example, if a script subroutine
	// is launched while we're in the middle of something here, our activity is suspended until
	// the subroutine completes and the call stack collapses back to here.  But if that subroutine
	// recursively calls us while the prior call is still in progress, the mShowIsInProgress would
	// be set to false when that layer completes, leaving it false when it really should be true
	// because our layer isn't done yet.
	mShowIsInProgress = true; // Signal WM_SIZE to queue the GuiSize launch.  We'll unqueue via MsgSleep() when we're done.

	// Change the title to get that out of the way.  But in any case, the title must be changed before the
	// following:
	// 1) Before the window is shown (to make transition a little nicer).
	// 2) v1.0.25: Before MoveWindow(), because otherwise the GuiSize label (if any) will be launched
	//    while the the window still has its old title (or no title, if this is the first showing), which
	//    would not be desirable 99% of the time.
	if (*aText)
		SetWindowText(mHwnd, aText);

	int x = COORD_UNSPECIFIED;
	int y = COORD_UNSPECIFIED;
	int width = COORD_UNSPECIFIED;
	int height = COORD_UNSPECIFIED;
	bool auto_size = false;

	// There is evidence that SW_SHOWNORMAL might be better than SW_SHOW for the first showing because
	// someone reported that a window appears centered on the screen for its first showing even if some
	// other position was specified.  In addition, MSDN says (without explanation): "An application should
	// specify [SW_SHOWNORMAL] when displaying the window for the first time."  However, SW_SHOWNORMAL is
	// avoided after the first showing of the window because that would probably also do a "restore" on the
	// window if it was maximized previously.  Note that the description of SW_SHOWNORMAL is virtually the
	// same as that of SW_RESTORE in MSDN.  UPDATE: mFirstGuiShowCmd is used here instead of mFirstActivation
	// because it seems more flexible to have "Gui Show" behave consistently (SW_SHOW) every time after
	// the first use of "Gui Show".  UPDATE: Since SW_SHOW seems to have no effect on minimized windows,
	// at least on XP, and since such a minimized window will be restored by action of SetForegroundWindowEx(),
	// it seems best to unconditionally use SW_SHOWNORMAL, rather than "mFirstGuiShowCmd ? SW_SHOWNORMAL : SW_SHOW".
	// This is done so that the window will be restored and thus have a better chance of being successfully
	// activated (and thus not requiring the call to SetForegroundWindowEx()).
	int show_mode = SW_SHOWNORMAL; // Set default.
	// Note that although SW_SHOW un-minimizes a window (at least on XP), it does not un-maximize it
	// (unlike SW_SHOWNORMAL, which seems functionally identical to SW_RESTORE).

	for (char *cp = aOptions; *cp; ++cp)
	{
		switch(toupper(*cp))
		{
		// For options such as W, H, X and Y:
		// Use atoi() vs. ATOI() to avoid interpreting something like 0x01B as hex when in fact
		// the B was meant to be an option letter:
		// DIMENSIONS:
		case 'A':
			if (!strnicmp(cp, "AutoSize", 8))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 7 vs. 8 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 7;
				auto_size = true;
			}
			break;
		case 'C':
			if (!strnicmp(cp, "Center", 6))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 5 vs. 6 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 5;
				x = COORD_CENTERED;
				y = COORD_CENTERED;
			}
			break;
		case 'M':
			if (!strnicmp(cp, "Minimize", 8)) // Seems best to reserve "Min" for other things, such as Min W/H. "Minimize" is also more self-documenting.
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 7 vs. 8 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 7;
				show_mode = SW_MINIMIZE;  // Seems more typically useful/desirable than SW_SHOWMINIMIZED.
			}
			else if (!strnicmp(cp, "Maximize", 8))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 7 vs. 8 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 7;
				show_mode = SW_MAXIMIZE;  // SW_MAXIMIZE == SW_SHOWMAXIMIZED
			}
			break;
		case 'N':
			if (!strnicmp(cp, "NA", 2))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 1 vs. 2 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 1;
				show_mode = SW_SHOWNA;
			}
			else if (!strnicmp(cp, "NoActivate", 10))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 9 vs. 10 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 9;
				show_mode = SW_SHOWNOACTIVATE;
			}
			break;
		case 'R':
			if (!strnicmp(cp, "Restore", 7))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 6 vs. 7 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 6;
				show_mode = SW_RESTORE;
			}
			break;
		case 'W':
			width = atoi(cp + 1);
			break;
		case 'H':
			if (!strnicmp(cp, "Hide", 4))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				// 3 vs. 4 to avoid the loop's addition ++cp from reading beyond the length of the string:
				cp += 3;
				show_mode = SW_HIDE;
			}
			else
				// Allow any width/height to be specified so that the window can be "rolled up" to its title bar:
				height = atoi(cp + 1);
			break;
		case 'X':
			if (!strnicmp(cp + 1, "Center", 6))
			{
				cp += 6; // 6 in this case since we're working with cp + 1
				x = COORD_CENTERED;
			}
			else
				x = atoi(cp + 1);
			break;
		case 'Y':
			if (!strnicmp(cp + 1, "Center", 6))
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

	int width_orig = width;
	int height_orig = height;

	// The following section must be done prior to any calls to GetWindow/ClientRect(mHwnd) because
	// neither of them can retrieve the correct diminsnions of a minmized window.  Similarly, if
	// the window is maximized but is about to be restored, do that prior to getting any of mHwnd's
	// rectangles because we want to use the restored size as the basis for centering, resizing, etc.
	// If show_mode is "hide", move the window only after hiding it (to reduce screen flicker).
	// If the window is being restored from a minimized or maximized state, move the window only
	// after restoring it; otherwise, any resize to be done by the MoveWindow() might not take effect.
	// Note that SW_SHOWNOACTIVATE is very similar to SW_RESTORE in its effects.
	bool show_was_done = false;
	if (show_mode == SW_HIDE // Hiding a window or restoring a window known to be minimized/maximized.
		|| (show_mode == SW_RESTORE || SW_SHOWNOACTIVATE) && (IsZoomed(mHwnd) || IsIconic(mHwnd)))
	{
		ShowWindow(mHwnd, show_mode);
		show_was_done = true;
	}
	// Note that SW_RESTORE and SW_SHOWNOACTIVATE will show a window if it's hidden.  Therefore, just
	// because the window is not in need of restoring doesn't mean the ShowWindow() call is skipped.
	// That is why show_was_done is left false in such cases.

	// Due to the checking above, if the window is minimized/maximized now, that means it will still be
	// minimized/maximized when this function is done.  As a result, it's not really valid to call
	// MoveWindow() for any purpose (auto-centering, auto-sizing, new position, new size, etc.).
	// The below is especially necessary for minimized windows because it avoid calculating window
	// dimensions, auto-centering, etc. based on incorrect values returned by GetWindow/ClientRect(mHwnd).
	// Update: For flexibililty, it seems best to allow a maximized window to be moved, which might be
	// valid on a multi-monitor system.  This maintains flexibility and doesn't appear to give up
	// anything because the script can do an explicit "Gui, Show, Restore" prior to a
	// "Gui, Show, x33 y44 w400" to be sure the window is restored before the operation (or combine
	// both of those commands into one).
	bool allow_move_window = !IsIconic(mHwnd);

	if (allow_move_window)
	{
		if (auto_size) // Check this one first so that it takes precedence over mFirstGuiShowCmd below.
		{
			// Find out a different set of max extents rather than using mMaxExtentRight/Down, which should
			// not be altered because they are used to position any subsequently added controls.
			RECT rect;
			width = 0;
			height = 0;
			for (GuiIndexType u = 0; u < mControlCount; ++u)
				if (GetWindowLong(mControl[u].hwnd, GWL_STYLE) & WS_VISIBLE) // Don't use IsWindowVisible() in case parent window is hidden.
				{
					GetWindowRect(mControl[u].hwnd, &rect);
					MapWindowPoints(NULL, mHwnd, (LPPOINT)&rect, 2); // Convert rect to client coordinates (not the same as GetClientRect()).
					if (rect.right > width)
						width = rect.right;
					if (rect.bottom > height)
						height = rect.bottom;
				}
			if (width > 0)
				width += mMarginX;
			if (height > 0)
				height += mMarginY;
		}
		else if (width == COORD_UNSPECIFIED || height == COORD_UNSPECIFIED)
		{
			if (mFirstGuiShowCmd) // By default, center the window if this is the first use of "Gui Show" (even "Gui Show, Hide").
			{
				if (width == COORD_UNSPECIFIED)
					width = mMaxExtentRight + mMarginX;
				if (height == COORD_UNSPECIFIED)
					height = mMaxExtentDown + mMarginY;
			}
			else
			{
				RECT rect;
				GetClientRect(mHwnd, &rect);
				if (width == COORD_UNSPECIFIED) // Keep the current client width, as documented.
					width = rect.right - rect.left;
				if (height == COORD_UNSPECIFIED) // Keep the current client height, as documented.
					height = rect.bottom - rect.top;
			}
		}
	} // if (allow_move_window)

	if (mFirstGuiShowCmd)
	{
		// Update any tab controls to show only their correct pane.  This should only be necessary
		// upon the first "Gui Show" (even "Gui, Show, Hide") of the window because subsequent switches
		// of the control's tab should result in a TCN_SELCHANGE notification.
		for (GuiIndexType u = 0; u < mControlCount; ++u)
			if (mControl[u].type == GUI_CONTROL_TAB)
				ControlUpdateCurrentTab(mControl[u], false); // Pass false so that default/z-order focus is used across entire window.
		// By default, center the window if this is the first time it's being shown:
		if (x == COORD_UNSPECIFIED)
			x = COORD_CENTERED;
		if (y == COORD_UNSPECIFIED)
			y = COORD_CENTERED;
	}

	BOOL is_visible = IsWindowVisible(mHwnd);

	if (allow_move_window)
	{
		// The above has determined the height/width of the client area.  From that area, determine
		// the window's new rect, including title bar, borders, etc.
		// If the window has a border or caption this also changes top & left *slightly* from zero.
		RECT rect = {0, 0, width, height}; // left,top,right,bottom
		AdjustWindowRectEx(&rect, GetWindowLong(mHwnd, GWL_STYLE), GetMenu(mHwnd) ? TRUE : FALSE
			, GetWindowLong(mHwnd, GWL_EXSTYLE));
		width = rect.right - rect.left;  // rect.left might be slightly less than zero.
		height = rect.bottom - rect.top; // rect.top might be slightly less than zero.

		RECT work_rect;
		SystemParametersInfo(SPI_GETWORKAREA, 0, &work_rect, 0);  // Get desktop rect excluding task bar.
		int work_width = work_rect.right - work_rect.left;  // Note that "left" won't be zero if task bar is on left!
		int work_height = work_rect.bottom - work_rect.top; // Note that "top" won't be zero if task bar is on top!

		// Seems best to restrict window size to the size of the desktop whenever explicit sizes
		// weren't given, since most users would probably want that.  But only on first use of
		// "Gui Show" (even "Gui, Show, Hide"):
		if (mFirstGuiShowCmd)
		{
			if (width_orig == COORD_UNSPECIFIED && width > work_width)
				width = work_width;
			if (height_orig == COORD_UNSPECIFIED && height > work_height)
				height = work_height;
		}

		if (x == COORD_CENTERED || y == COORD_CENTERED) // Center it, based on its dimensions determined above.
		{
			// This does not currently handle multi-monitor systems explicitly, since those calculations
			// require API functions that don't exist in Win95/NT (and thus would have to be loaded
			// dynamically to allow the program to launch).  Therefore, windows will likely wind up
			// being centered across the total dimensions of all monitors, which usually results in
			// half being on one monitor and half in the other.  This doesn't seem too terrible and
			// might even be what the user wants in some cases (i.e. for really big windows).
			if (x == COORD_CENTERED)
				x = work_rect.left + ((work_width - width) / 2);
			if (y == COORD_CENTERED)
				y = work_rect.top + ((work_height - height) / 2);
		}

		RECT old_rect;
		GetWindowRect(mHwnd, &old_rect);
		int old_width = old_rect.right - old_rect.left;
		int old_height = old_rect.bottom - old_rect.top;

		// Avoid calling MoveWindow() if nothing changed because it might repaint/redraw even if window size/pos
		// didn't change:
		if (width != old_width || height != old_height || (x != COORD_UNSPECIFIED && x != old_rect.left)
			|| (y != COORD_UNSPECIFIED && y != old_rect.bottom))
		{
			MoveWindow(mHwnd, x == COORD_UNSPECIFIED ? old_rect.left : x, y == COORD_UNSPECIFIED ? old_rect.top : y
				, width, height, is_visible);  // Do repaint if window is visible.
		}
	} // if (allow_move_window)

	// Note that for SW_MINIMIZE and SW_MAXIMZE, the MoveWindow() above should be done prior to ShowWindow()
	// so that the window will "remember" its new size upon being restored later.
	if (!show_was_done)
		ShowWindow(mHwnd, show_mode);

	bool we_did_the_first_activation = false; // Set default.

	switch(show_mode)
	{
	case SW_SHOW:
	case SW_SHOWNORMAL:
	case SW_MAXIMIZE:
	case SW_RESTORE:
		if (mHwnd != GetForegroundWindow())
			SetForegroundWindowEx(mHwnd);   // In the above modes, try to force it to the foreground as documented.
		if (mFirstActivation)
		{
			// Since the window has never before been active, any of the above qualify as "first activation".
			// Thus, we are no longer at the first activation:
			mFirstActivation = false;
			we_did_the_first_activation = true; // And we're the ones who did the first activation.
		}
		break;
	// No action for these:
	//case SW_MINIMIZE:
	//case SW_SHOWNA:
	//case SW_SHOWNOACTIVATE:
	//case SW_HIDE:
	}

	// No attempt is made to handle the fact that Gui windows can be shown or activated via WinShow and
	// WinActivate.  In such cases, if the tab control itself is focused, mFirstActivation will stil focus
	// a control inside the tab rather than leaving the tab control focused.  Similarly, if the window
	// was shown with NA or NOACTIVATE or MINIMIZE, when the first use of an activation mode of "Gui Show"
	// is done, even if it's far into the future, long after the user has activated and navigated in the
	// window, the same "first activation" behavior will be done anyway.  This is documented here as a
	// known limitation, since fixing it would probably add an unreasonable amount of complexity.
	HWND focused_control_hwnd;
	if (we_did_the_first_activation && mTabControlCount // Window probably must be visible and active for GetFocus() to work.
		&& (focused_control_hwnd = GetFocus())) // Assign
	{
		// Since this is the first activation, if the focus wound up on a tab control itself as a result
		// of the above, focus the first control of that tab since that is traditional.  HOWEVER, do not
		// instead default tab controls to lacking WS_TABSTOP since it is traditional for them to have
		// that property, probably to aid accessibility.
		GuiControlType *focused_control = FindControl(focused_control_hwnd);
		if (focused_control && focused_control->type == GUI_CONTROL_TAB)
		{
			// v1.0.27: The following must be done, at least in some cases, because otherwise
			// controls outside of the tab control will not get drawn correctly.  I suspect this
			// is because at the exact moment execution reaches the line below, the window is in
			// a transitional state, with some WM_PAINT and other messages waiting in the queue
			// for it.  If those messages are not processed prior to ControlUpdateCurrentTab()'s
			// use of WM_SETREDRAW, they might get dropped out of the queue and lost.
			UpdateWindow(mHwnd);
			ControlUpdateCurrentTab(*focused_control, true);
		}
	}

	mFirstGuiShowCmd = false;

	// It seems best to reset this prior to SLEEP below, but after the above line (for code clarity) since
	// otherwise it might get stuck in a true state if the SLEEP results in the launch of a script
	// subroutine that takes a long time to complete:
	mShowIsInProgress = false;

	// Update for v1.0.25: The below is now done last to prevent the GuiSize label (if any) from launching
	// while this function is still incomplete; in other words, don't allow the GuiSize label to launch
	// until after all of the above members and actions have been completed.
	// This is done for the same reason it's done for ACT_SPLASHTEXTON.  If it weren't done, whenever
	// a command that blocks (fully uses) the main thread such as "Drive Eject" immediately follows
	// "Gui Show", the GUI window might not appear until afterward because our thread never had a
	// chance to call its WindowProc with all the messages needed to actually show the window:
	SLEEP_WITHOUT_INTERRUPTION(-1)
	// UpdateWindow() would probably achieve the same effect as the above, but it feels safer to do
	// the above because it ensures that our message queue is empty prior to returning to our caller.

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
	GuiIndexType u;
	for (u = 0; u < mControlCount; ++u)
		if (mControl[u].output_var && mControl[u].type != GUI_CONTROL_RADIO)
			ControlGetContents(*mControl[u].output_var, mControl[u], "Submit");

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
				// Instead, treat this the same as "none selected".  Update for v1.0.24: It is no longer
				// directly possible to have multiple radios selected by having the word "checked" in
				// more than one of their "Gui Add" commands.  However, there are probably other ways
				// to get multiple buttons checked (perhaps the Control command), so this handling
				// for multiple selections is left intact.
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



ResultType GuiType::ControlGetContents(Var &aOutputVar, GuiControlType &aControl, char *aMode)
{
	char buf[1024]; // For various uses.
	bool submit_mode = !stricmp(aMode, "Submit");

	// First handle any control types that behave the same regardless of aMode:
	switch (aControl.type)
	{
	case GUI_CONTROL_SLIDER: // Doesn't seem useful to ever retrieve the control's actual caption, which is invisible.
		return aOutputVar.Assign(ControlInvertSliderIfNeeded(aControl, (int)SendMessage(aControl.hwnd, TBM_GETPOS, 0, 0)));
		// Above assigns it as a signed value because testing shows a slider can have part or all of its
		// available range consist of negative values.  32-bit values are supported if the range is set
		// with the right messages.
	case GUI_CONTROL_PROGRESS:
		return submit_mode ? OK : aOutputVar.Assign((int)SendMessage(aControl.hwnd, PBM_GETPOS, 0, 0));
		// Above does not save to control during submit mode, since progress bars do not receive
		// user input so it seems wasteful 99% of the time.  "GuiControlGet, MyProgress" can be used instead.
	case GUI_CONTROL_HOTKEY:
		// Testing shows that neither GetWindowText() nor WM_GETTEXT can pull anything out of a hotkey
		// control, so the only type of retrieval that can be offered is the HKM_GETHOTKEY method:
		HotkeyToText((WORD)SendMessage(aControl.hwnd, HKM_GETHOTKEY, 0, 0), buf);
		return aOutputVar.Assign(buf);
	}

	if (stricmp(aMode, "Text")) // Non-text, i.e. don't unconditionally use the simple GetWindowText() method.
	{
		// The caller wants the contents of the control, which is often different from its
		// caption/text.  Any control types not mentioned in the switch() below will fall through
		// into the section at the bottom that applies the standard GetWindowText() method.

		LRESULT index, length, item_length;

		switch (aControl.type)
		{
		case GUI_CONTROL_TEXT:
		case GUI_CONTROL_PIC:
		case GUI_CONTROL_GROUPBOX:
		case GUI_CONTROL_BUTTON:
		case GUI_CONTROL_PROGRESS:
			if (submit_mode) // In submit mode, do not waste memory & cpu time to save the above.
				return OK;
				// There doesn't seem to be a strong/net advantage to setting the vars to be blank
				// because even if that were done, it seems it would not do much to reserve flexibility
				// for future features in which these associated variables are used for a purpose other
				// than uniquely identifying the control with GuiControl & GuiControlGet.
			break;

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

		case GUI_CONTROL_DROPDOWNLIST:
			if (aControl.attrib & GUI_CONTROL_ATTRIB_ALTSUBMIT) // Caller wanted the position, not the text retrieved.
			{
				index = SendMessage(aControl.hwnd, CB_GETCURSEL, 0, 0); // Get index of currently selected item.
				if (index == CB_ERR) // Maybe happens only if DROPDOWNLIST has no items at all, so ErrorLevel is not changed.
					return aOutputVar.Assign();
				return aOutputVar.Assign((int)index + 1);
			}
			break; // Fall through to the normal GetWindowText() method, which works for DDLs but not ComboBoxes.

		case GUI_CONTROL_COMBOBOX:
			index = SendMessage(aControl.hwnd, CB_GETCURSEL, 0, 0); // Get index of currently selected item.
			if (index == CB_ERR) // There is no selection (or very rarely, some other type of problem).
				break; // Break out of the switch rather than returning so that the GetWindowText() method can be applied.
			if (aControl.attrib & GUI_CONTROL_ATTRIB_ALTSUBMIT) // Caller wanted the position, not the text retrieved.
				return aOutputVar.Assign((int)index + 1);
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
			if (GetWindowLong(aControl.hwnd, GWL_STYLE) & (LBS_EXTENDEDSEL|LBS_MULTIPLESEL))
			{
				LRESULT item_count = SendMessage(aControl.hwnd, LB_GETSELCOUNT, 0, 0);
				if (item_count <= 0)  // <=0 to check for LB_ERR too (but it should be impossible in this case).
					return aOutputVar.Assign();
				int *item = (int *)malloc(item_count * sizeof(int)); // dynamic since there can be a very large number of items.
				if (!item)
					return aOutputVar.Assign();
				item_count = SendMessage(aControl.hwnd, LB_GETSELITEMS, (WPARAM)item_count, (LPARAM)item);
				if (item_count <= 0)  // 0 or LB_ERR, but both these conditions should be impossible in this case.
				{
					free(item);
					return aOutputVar.Assign();
				}
				LRESULT u;
				if (aControl.attrib & GUI_CONTROL_ATTRIB_ALTSUBMIT) // Caller wanted the positions, not the text retrieved.
				{
					// Accumulate the length of delimited list of positions.
					// length is initialized to item_count - 1 to account for all the delimiter
					// characters in the list, one delim after each item except the last:
					for (length = item_count - 1, u = 0; u < item_count; ++u)
					{
						_itoa(item[u] + 1, buf, 10);  // +1 to convert from zero-based to 1-based.
						length += strlen(buf);
					}
				}
				else
				{
					// Accumulate the length of delimited list of selected items (not positions in this case).
					// See above loop for more comments.
					for (length = item_count - 1, u = 0; u < item_count; ++u)
					{
						item_length = SendMessage(aControl.hwnd, LB_GETTEXTLEN, (WPARAM)item[u], 0);
						if (item_length == LB_ERR) // Realistically impossible based on MSDN.
						{
							free(item);
							return aOutputVar.Assign();
						}
						length += item_length;
					}
				}
				// In unusual cases, MSDN says the indicated length might be longer than it actually winds up
				// being when the item's text is retrieved.  This should be harmless, since there are many
				// other precedents where a variable is sized to something larger than it winds up carrying.
				// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
				// this call will set up the clipboard for writing:
				if (aOutputVar.Assign(NULL, (VarSizeType)length) != OK)
					return FAIL;  // It already displayed the error.
				char *cp = aOutputVar.Contents(); // Init for both of the loops below.
				if (aControl.attrib & GUI_CONTROL_ATTRIB_ALTSUBMIT) // Caller wanted the positions, not the text retrieved.
				{
					// In this case, the original length estimate should be the same as the actual, so
					// it is not re-accumulated.
					// See above loop for more comments.
					for (u = 0; u < item_count; ++u)
					{
						_itoa(item[u] + 1, cp, 10);  // +1 to convert from zero-based to 1-based.
						cp += strlen(cp);  // Point it to the terminator in preparation for the next write.
						if (u < item_count - 1)
							*cp++ = '|'; // Add delimiter after each item except the last (helps parsing loop).
					}
				}
				else
				{
					// See above loop for more comments.
					for (length = item_count - 1, u = 0; u < item_count; ++u)
					{
						item_length = SendMessage(aControl.hwnd, LB_GETTEXT, (WPARAM)item[u], (LPARAM)cp);
						if (length == LB_ERR) // Given the way it was called, this should be impossible based on MSDN docs.
						{
							aOutputVar.Close(); // In case it's the clipboard.
							free(item);
							return aOutputVar.Assign();
						}
						length += item_length; // Accumulate actual vs. estimated length.
						cp += item_length;  // Point it to the terminator in preparation for the next write.
						if (u < item_count - 1)
							*cp++ = '|'; // Add delimiter after each item except the last (helps parsing loop).
						// Above:
						// A hard-coded pipe delimiter is used for now because it seems fairly easy to
						// add an option later for a custom delimtier (such as '\n') via an Param4 of
						// GuiControlGetText and/or an option-word in "Gui Add".  The reason pipe is
						// used as a delimiter is that it allows the selection to be easily inserted
						// into another ListBox because it's already in the right format with the
						// right delimiter.  In addition, literal pipes should be rare since that is
						// the delimiter used when insertting and appending entries into a ListBox.
					}
				}
				free(item);
			}
			else // Single-select ListBox style.
			{
				index = SendMessage(aControl.hwnd, LB_GETCURSEL, 0, 0); // Get index of currently selected item.
				if (index == LB_ERR) // There is no selection (or very rarely, some other type of problem).
					return aOutputVar.Assign();
				if (aControl.attrib & GUI_CONTROL_ATTRIB_ALTSUBMIT) // Caller wanted the position, not the text retrieved.
					return aOutputVar.Assign((int)index + 1);
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
			}
			aOutputVar.Length() = (VarSizeType)length;  // Update it to the actual length, which can vary from the estimate.
			return aOutputVar.Close(); // In case it's the clipboard.

		case GUI_CONTROL_TAB:
			index = TabCtrl_GetCurSel(aControl.hwnd); // Get index of currently selected item.
			if (index == -1) // There is no selection (maybe happens only if it has no tabs at all), so ErrorLevel is not changed.
				return aOutputVar.Assign();
			if (aControl.attrib & GUI_CONTROL_ATTRIB_ALTSUBMIT) // Caller wanted the index, not the text retrieved.
				return aOutputVar.Assign((int)index + 1);
			// Otherwise: Get the stored name/caption of this tab:
			TCITEM tci;
			tci.mask = TCIF_TEXT;
			tci.pszText = buf;
			tci.cchTextMax = sizeof(buf) - 1; // MSDN example uses -1.
			if (TabCtrl_GetItem(aControl.hwnd, index, &tci))
				return aOutputVar.Assign(tci.pszText);
			return aOutputVar.Assign();
		// Types specifically not handled here.  They will be handled by the section below this switch():
		//case GUI_CONTROL_EDIT:
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
		// Since edit controls tend to have many hard returns in them, use "true" for the last param to
		// enhance performance.  This performance gain is extreme when the control contains thousands
		// of CRLFs:
		StrReplaceAll(aOutputVar.Contents(), "\r\n", "\n", false);
		aOutputVar.Length() = (VarSizeType)strlen(aOutputVar.Contents());
	}
	return aOutputVar.Close();  // In case it's the clipboard.
}



GuiIndexType GuiType::FindControl(char *aControlID)
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
	GuiIndexType u;
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



int GuiType::FindGroup(GuiIndexType aControlIndex, GuiIndexType &aGroupStart, GuiIndexType &aGroupEnd)
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
				// For v1.0.22, this is no longer done because want to support an optional leading 0x
				// if it is present, e.g. 0xFFAABB.  It seems strtol() automatically handles the
				// optional leading "0x" if present:
				//if (strlen(color_str) > 6)
				//	color_str[6] = '\0';  // Shorten to exactly 6 chars, which happens if no space/tab delimiter is present.
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
	GuiIndexType control_index;
	WORD wParam_loword;
	RECT rect;
	bool text_color_was_changed;
	char buf[1024];

	switch (iMsg)
	{
	// Let the default handler take care of WM_CREATE.

	case WM_SIZE:
		if (   !(pgui = GuiType::FindGui(hWnd))   )
			break; // Let default proc handle it.
		if (pgui->mLabelForSize) // There is an event handler in the script.
		{
			pgui->mSizeType = wParam;
			pgui->mSizeWidthHeight = lParam; // A slight aid to performance to only divide it into halves upon demand (later).
			POST_AHK_GUI_ACTION(pgui->mHwnd, AHK_GUI_SIZE, GUI_EVENT_NORMAL);
			if (!pgui->mShowIsInProgress) // v1.0.25
				MsgSleep(-1); // See Gui::Event() for details about this.
			//else don't do the MsgSleep now.  Let Gui::Show() do it so that the launch of the GuiSize
			// label does not occur until after Show() has unhidden the window and completely finished
			// its various other activities.  This avoids the need to turn on DetectHiddenWindows in the
			// script for the first launch of GuiSize.  It also avoids an unnecessary MsgSleep() call,
			// since Gui::Show() will be doing that for us.
		}
		return 0; // "If an application processes this message, it should return zero."
		// Testing shows that the window still resizes correctly (controls are revealed as the window
		// is expanded) even if the event isn't passed on to the default proc.

	case WM_COMMAND:
	{
		// First find which of the GUI windows is receiving this event, if none (probably impossible
		// the way things are set up currently), let DefaultProc handle it:
		if (   !(pgui = GuiType::FindGui(hWnd))   )
			break;
		wParam_loword = LOWORD(wParam);
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
		// Note: It is not necessary to check for IDOK because:
		// 1) If there is no default button, the IDOK message is ignored.
		// 2) If there is a default button, we should never receive IDOK because BM_SETSTYLE (sent earlier)
		// will have altered the message we receive to be the ID of the actual default button.
		if (wParam_loword == IDCANCEL) // The user pressed ESCAPE.
		{
			pgui->Escape();
			return 0;
		}
		GuiIndexType control_index = GUI_ID_TO_INDEX(wParam_loword); // Convert from ID to array index.
		if (control_index < pgui->mControlCount // Relies on short-circuit boolean order.
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

	case WM_NOTIFY:
	{
		if (   !(pgui = GuiType::FindGui(hWnd))   )
			break; // Let DefDlgProc() handle it.
		NMHDR &nmhdr = *(LPNMHDR)lParam;
		switch(nmhdr.code)
		{
		case TCN_SELCHANGING:
		case TCN_SELCHANGE:
			control_index = (GuiIndexType)GUI_ID_TO_INDEX(nmhdr.idFrom); // Convert from ID to array index.
			if (control_index < pgui->mControlCount // Relies on short-circuit eval order.
				&& pgui->mControl[control_index].hwnd == nmhdr.hwndFrom) // Handles match (this filters out bogus msgs).
			{
				if (nmhdr.code == TCN_SELCHANGE)
				{
					pgui->ControlUpdateCurrentTab(pgui->mControl[control_index], true);
					pgui->Event(control_index, nmhdr.code);
				}
				else // TCN_SELCHANGING
					if (pgui->mControl[control_index].output_var && pgui->mControl[control_index].jump_to_label)
						pgui->ControlGetContents(*pgui->mControl[control_index].output_var, pgui->mControl[control_index]);
			}
			break; // inner switch()
		}
		break; // outer switch()
	}

	case WM_HSCROLL: // These two should only be received for slider controls (trackbars).
	case WM_VSCROLL:
		if (   !(pgui = GuiType::FindGui(hWnd))   )
			break; // Let default proc handle it.
		pgui->Event(GUI_HWND_TO_INDEX((HWND)lParam), LOWORD(wParam));
		return 0; // "If an application processes this message, it should return zero."
	
	case WM_ERASEBKGND:
		if (   !(pgui = GuiType::FindGui(hWnd))   )
			break; // Let DefDlgProc() handle it.
		if (!pgui->mBackgroundBrushWin) // Let default proc handle it.
			break;
		// Can't use SetBkColor(), need an real brush to fill it.
		GetClipBox((HDC)wParam, &rect);
		FillRect((HDC)wParam, &rect, pgui->mBackgroundBrushWin);
		return 1; // "An application should return nonzero if it erases the background."

	// The below seems to be the equivalent of the above (but MSDN indicates it will only work
	// if there is no WM_ERASEBKGND handler).  Although it might perform a little better,
	// the above is kept in effect to avoid introducing problems without a good reason:
	//case WM_CTLCOLORDLG:
	//	if (   !(pgui = GuiType::FindGui(hWnd))   )
	//		break; // Let DefDlgProc() handle it.
	//	if (!pgui->mBackgroundBrushWin) // Let default proc handle it.
	//		break;
	//	SetBkColor((HDC)wParam, pgui->mBackgroundColorWin);
	//	return (LRESULT)pgui->mBackgroundBrushWin;

	// It seems that scrollbars belong to controls (such as Edit and ListBox) do not send us
	// WM_CTLCOLORSCROLLBAR (unlike the static messages we receive for radio and checkbox).
	// Therefore, this section is commented out since it has no effect (it might be useful
	// if a control's class window-proc is ever overridden with a new proc):
	//case WM_CTLCOLORSCROLLBAR:
	//	if (   !(pgui = GuiType::FindGui(hWnd))   )
	//		break;
	//	if (pgui->mBackgroundBrushWin)
	//	{
	//		// Since we're processing this msg rather than passing it on to the default proc, must set
	//		// background color unconditionally, otherwise plain white will likely be used:
	//		SetTextColor((HDC)wParam, pgui->mBackgroundColorWin);
	//		SetBkColor((HDC)wParam, pgui->mBackgroundColorWin);
	//		// Always return a real HBRUSH so that Windows knows we altered the HDC for it to use:
	//		return (LRESULT)pgui->mBackgroundBrushWin;
	//	}
	//	break;

	case WM_CTLCOLORSTATIC:
	case WM_CTLCOLORLISTBOX:
	case WM_CTLCOLOREDIT:
		// MSDN: Buttons with the BS_PUSHBUTTON, BS_DEFPUSHBUTTON, or BS_PUSHLIKE styles do not use the
		// returned brush. Buttons with these styles are always drawn with the default system colors.
		// This is because "drawing push buttons requires several different brushes-face, highlight and
		// shadow". In short, to provide a custom appearance for push buttons, use an owner-drawn button.
		// Thus, WM_CTLCOLORBTN not handled here because it doesn't seem to have any effect on the
		// types of buttons used so far.  This has been confirmed: Even when a theme is in effect,
		// checkboxes, radios, and groupboxes do not receive WM_CTLCOLORBTN, but they do receive
		// WM_CTLCOLORSTATIC.
		if (   !(pgui = GuiType::FindGui(hWnd))   )
			break;
		if (   !(pcontrol = pgui->FindControl((HWND)lParam))   )
			break;
		if (pcontrol->type == GUI_CONTROL_COMBOBOX) // But GUI_CONTROL_DROPDOWNLIST partially works.
			// Setting the colors of combo boxes won't work without overriding the ComboBox window proc,
			// which introduces complexities because there is no knowing exactly what the default
			// window proc of a ComboBox really does in all OSes and under all visual themes.
			// Overriding it is likely to cause problems, or at the very least require testing across
			// various OSes and themes (XP vs. classic).
			break;
		if (text_color_was_changed = (pcontrol->type != GUI_CONTROL_PIC && pcontrol->union_color != CLR_DEFAULT))
			SetTextColor((HDC)wParam, pcontrol->union_color);

		if (pcontrol->attrib & GUI_CONTROL_ATTRIB_BACKGROUND_TRANS)
		{
			switch (pcontrol->type)
			{
			case GUI_CONTROL_CHECKBOX: // Checkbox and radios with trans background have problems with
			case GUI_CONTROL_RADIO:    // their focus rects being drawn incorrectly.
			case GUI_CONTROL_LISTBOX:  // ListBox and Edit are also a problem, at least under some theme settings.
			case GUI_CONTROL_EDIT:
			case GUI_CONTROL_SLIDER:   // Slider is a problem under both classic and XP themes.
				break;  // Ignore the TRANS setting for the above control types.
			// Types not included above because they support transparent background or because the attempt
			// to make the background transparent has no effect:
			//case GUI_CONTROL_TEXT:         Supported via WM_CTLCOLORSTATIC
			//case GUI_CONTROL_PIC:          Supported via WM_CTLCOLORSTATIC
			//case GUI_CONTROL_GROUPBOX:     Supported via WM_CTLCOLORSTATIC
			//case GUI_CONTROL_BUTTON:       Can't reach this point because WM_CTLCOLORBTN is not handled above.
			//case GUI_CONTROL_DROPDOWNLIST: Can't reach this point because WM_CTLCOLORxxx is never received for it.
			//case GUI_CONTROL_COMBOBOX:     I believe WM_CTLCOLOREDIT is not received for it.
			//case GUI_CONTROL_PROGRESS:     Can't reach this point because WM_CTLCOLORxxx is never received for it.
			//case GUI_CONTROL_HOTKEY:       Same (verified).
			//case GUI_CONTROL_TAB:          Same.
			default:
				SetBkMode((HDC)wParam, TRANSPARENT);
				return (LRESULT)GetStockObject(NULL_BRUSH);
			}
			//else ignore the TRANS setting, since it causes the ListBox (at least in Classic theme)
			// to appear to be multi-select even though it isn't.  And it causes Edit to have a
			// black background.
		}
		if (pcontrol->attrib & GUI_CONTROL_ATTRIB_BACKGROUND_DEFAULT) // i.e. TRANS (above) takes precedence over this.
		{
			if (!text_color_was_changed) // No need to return a brush since no changes are needed.  Let def. proc. handle it.
				break;
			if (iMsg == WM_CTLCOLORSTATIC)
			{
				SetBkColor((HDC)wParam, GetSysColor(COLOR_BTNFACE)); // Use default window color for static controls.
				return (LRESULT)GetSysColorBrush(COLOR_BTNFACE);
			}
			else
			{
				SetBkColor((HDC)wParam, GetSysColor(COLOR_WINDOW)); // Use default control-bkgnd color for others.
				return (LRESULT)GetSysColorBrush(COLOR_WINDOW);
			}
		}

		if (iMsg == WM_CTLCOLORSTATIC)
		{
			// If this static control both belongs to a tab control and is within its physical boundaries,
			// match its background to the tab control's.  This is only necessary if the tab control has
			// the GUI_CONTROL_ATTRIB_BACKGROUND_DEFAULT property, since otherwise its background would
			// be the same as the window's:
			bool override_to_default_color = pgui->ControlOverrideBkColor(*pcontrol);
			if (pgui->mBackgroundBrushWin && !override_to_default_color)
			{
				// Since we're processing this msg rather than passing it on to the default proc, must set
				// background color unconditionally, otherwise plain white will likely be used:
				SetBkColor((HDC)wParam, pgui->mBackgroundColorWin);
				// Always return a real HBRUSH so that Windows knows we altered the HDC for it to use:
				return (LRESULT)pgui->mBackgroundBrushWin;
			}
			// else continue on through so that brush can be returned if text_color_was_changed == true.
		}
		else // WM_CTLCOLORLISTBOX or WM_CTLCOLOREDIT: The interior of a non-static control.  Use the control background color (if there is one).
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
		if (text_color_was_changed)
		{
			// Whenever the default proc won't be handling this message, the background color must be set
			// explicitly if something other than plain white is needed.  This must be done even for
			// non-static controls because otherwise the area doesn't get filled correctly:
			if (iMsg == WM_CTLCOLORSTATIC)
			{
				// COLOR_BTNFACE is hard-coded because here because it is also the hard-coded background
				// color of the GUI window class:
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

	case WM_DRAWITEM:
	{
		// WM_DRAWITEM msg is never received if there are no GUI windows containing a tab
		// control with custom tab colors.  The TCS_OWNERDRAWFIXED style is what causes
		// this message to be received.
		if (   !(pgui = GuiType::FindGui(hWnd))   )
			break;
		LPDRAWITEMSTRUCT lpdis = (LPDRAWITEMSTRUCT)lParam;
		control_index = (GuiIndexType)GUI_ID_TO_INDEX(lpdis->CtlID); // Convert from ID to array index.
		if (control_index >= pgui->mControlCount // Relies on short-circuit eval order.
			|| pgui->mControl[control_index].hwnd != lpdis->hwndItem  // Handles do not match (this filters out bogus msgs).
			|| pgui->mControl[control_index].type != GUI_CONTROL_TAB) // In case this msg can be received for other types.
			break;
		GuiControlType &control = pgui->mControl[control_index]; // For performance & convenience.
		if (pgui->mBackgroundBrushWin && !(control.attrib & GUI_CONTROL_ATTRIB_BACKGROUND_DEFAULT))
		{
			FillRect(lpdis->hDC, &lpdis->rcItem, pgui->mBackgroundBrushWin); // Fill the tab itself.
			SetBkColor(lpdis->hDC, pgui->mBackgroundColorWin); // Set the text's background color.
		}
		else // Must do this anyway, otherwise there is an unwanted thin white line and possibly other problems.
			FillRect(lpdis->hDC, &lpdis->rcItem, (HBRUSH)GetClassLong(control.hwnd, GCL_HBRBACKGROUND));
		// else leave background colors to default, in the case where only the text itself has a custom color.
		// Get the stored name/caption of this tab:
		TCITEM tci;
		tci.mask = TCIF_TEXT;
		tci.pszText = buf;
		tci.cchTextMax = sizeof(buf) - 1; // MSDN example uses -1.
		// Set text color if needed:
        COLORREF prev_color = CLR_INVALID;
		if (control.union_color != CLR_DEFAULT)
			prev_color = SetTextColor(lpdis->hDC, control.union_color);
		// Draw the text.  Note that rcItem contains the dimensions of a tab that has already been sized
		// to handle the amount of text in the tab at the specified WM_SETFONT font size, which makes
		// this much easier.
		if (TabCtrl_GetItem(lpdis->hwndItem, lpdis->itemID, &tci))
		{
			// The text is centered horizontally and vertically because that seems to be how the
			// control acts without the TCS_OWNERDRAWFIXED style.  DT_NOPREFIX is not specified
			// because that is not how the control acts without the TCS_OWNERDRAWFIXED style
			// (ampersands do cause underlined letters, even though they currently have no effect).
			if (TabCtrl_GetCurSel(control.hwnd) != lpdis->itemID)
				lpdis->rcItem.top += 3; // For some reason, the non-current tabs' rects are a little off.
			DrawText(lpdis->hDC, tci.pszText, (int)strlen(tci.pszText), &lpdis->rcItem
				, DT_CENTER|DT_VCENTER|DT_SINGLELINE); // DT_VCENTER requires DT_SINGLELINE.
			// Cruder method, probably not always accurate depending on theme/display-settings/etc.:
			//TextOut(lpdis->hDC, lpdis->rcItem.left + 5, lpdis->rcItem.top + 3, tci.pszText, (int)strlen(tci.pszText));
		}
		if (prev_color != CLR_INVALID) // Put the previous color back into effect for this DC.
			SetTextColor(lpdis->hDC, prev_color);
		break;
	}

	case WM_DROPFILES:
	{
		if (   !(pgui = GuiType::FindGui(hWnd))   )
			break; // Let DefDlgProc() handle it.
		HDROP hdrop = (HDROP)wParam;
		if (!pgui->mLabelForDropFiles || pgui->mHdrop)
		{
			// There is no event handler in the script, or this window is still processing a prior drop.
			// Ignore this drop and free its memory.
			DragFinish(hdrop);
			return 0; // "An application should return zero if it processes this message."
		}
		// Otherwise: Indicate that this window is now processing the drop.  DragFinish() will be called later.
		pgui->mHdrop = hdrop;
		point_and_hwnd_type pah = {0};
		// DragQueryPoint()'s return value is non-zero if the drop occurred in the client area.
		// However, that info seems too rarely needed to justify storing it anywhere:
		DragQueryPoint(hdrop, &pah.pt);
		ClientToScreen(pgui->mHwnd, &pah.pt); // EnumChildFindPoint() requires screen coords.
		EnumChildWindows(pgui->mHwnd, EnumChildFindPoint, (LPARAM)&pah); // Find topmost control containing point.
		// Look up the control in case the drop occurred in a child of a child, such as the edit portion
		// of a ComboBox (FindControl will take that into account):
		pcontrol = pah.hwnd_found ? pgui->FindControl(pah.hwnd_found) : NULL;
		control_index = pcontrol ? GUI_HWND_TO_INDEX(pcontrol->hwnd) : MAX_CONTROLS_PER_GUI;
		// Above: MAX_CONTROLS_PER_GUI indicates to GetGuiControl() that there is no control in this case.
		POST_AHK_GUI_ACTION(pgui->mHwnd, AHK_GUI_DROPFILES, control_index);
		MsgSleep(-1); // See Gui::Event() for details about this.
		return 0; // "An application should return zero if it processes this message."
	}

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



LRESULT CALLBACK TabWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
	// Variables are kept separate up here for future expansion of this function (to handle
	// more iMsgs/cases, etc.):
	GuiType *pgui;
	GuiControlType *pcontrol;
	HWND parent_window;

	if (iMsg == WM_ERASEBKGND)
	{
		parent_window = GetParent(hWnd);
		// Relies on short-circuit boolean order:
		if (   (pgui = GuiType::FindGui(parent_window)) && (pcontrol = pgui->FindControl(hWnd))
			&& pgui->mBackgroundBrushWin && !(pcontrol->attrib & GUI_CONTROL_ATTRIB_BACKGROUND_DEFAULT)   )
		{
			// Can't use SetBkColor(), need an real brush to fill it.
			RECT clipbox;
			GetClipBox((HDC)wParam, &clipbox);
			FillRect((HDC)wParam, &clipbox, pgui->mBackgroundBrushWin);
			return 1; // "An application should return nonzero if it erases the background."
		}
		//else let default proc handle it.
	}

	// This will handle anything not already fully handled and returned from above:
	return CallWindowProc(g_TabClassProc, hWnd, iMsg, wParam, lParam);
}



void GuiType::Event(GuiIndexType aControlIndex, UINT aNotifyCode)
// Handles events within a GUI window that caused one of its controls to change in a meaningful way,
// or that is an event that could trigger an external action, such as clicking a button or icon.
{
	if (aControlIndex >= mControlCount) // Caller probably already checked, but just to be safe.
		return;
	GuiControlType &control = mControl[aControlIndex];
	if (!control.jump_to_label && !(control.attrib & GUI_CONTROL_ATTRIB_IMPLICIT_CANCEL))
		return;

	// Update: The below is now checked by MsgSleep() at the time the launch actually would occur:
	// If this control already has a thread running in its label, don't create a new thread to avoid
	// problems of buried threads, or a stack of suspended threads that might be resumed later
	// at an unexpected time. Users of timer subs that take a long time to run should be aware, as
	// documented in the help file, that long interruptions are then possible.
	//if (g_nThreads >= g_MaxThreadsTotal || (aControl->attrib & GUI_CONTROL_ATTRIB_LABEL_IS_RUNNING))
	//	continue

	GuiEventType gui_event = GUI_EVENT_NORMAL;  // Set default.  Don't use NONE since that means "not a GUI thread".

	// Explicitly cover all control types in the switch() rather than relying solely on
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
		case BN_CLICKED: // Must explicitly list this case since the default label below does a return.
			// Fix for v1.0.24: The below excludes from consideration messages from radios that are
			// being unchecked.  This prevents a radio group's g-label from being fired twice when the
			// user navigates to a new radio via the arrow keys.  It also filters out the BN_CLICKED that
			// occurs when the user tabs over to a radio group that lacks a selected button.  Such
			// behavior seems like it would be desirable most of the time.
			if (control.type == GUI_CONTROL_RADIO && SendMessage(control.hwnd, BM_GETCHECK, 0, 0) == BST_UNCHECKED)
				return;
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
		case CBN_SELCHANGE:  // Must explicitly list this case since the default label does a return.
		case CBN_EDITCHANGE: // Added for v1.0.24 to support detection of changes in a ComboBox's edit portion.
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

	case GUI_CONTROL_SLIDER:
		switch (aNotifyCode)
		{
		case TB_ENDTRACK: // WM_KEYUP (the user released a key that sent a relevant virtual key code)
			// Unfortunately, the control does not generate a TB_ENDTRACK notification when the slider
			// was moved via the mouse wheel.  This is documented here as a known limitation.  The
			// workaround is to use AltSubmit.
			break;
		default:
			// Namely the following:
			//case TB_THUMBPOSITION: // Mouse wheel or WM_LBUTTONUP following a TB_THUMBTRACK notification message
			//case TB_THUMBTRACK:    // Slider movement (the user dragged the slider)
			//case TB_LINEUP:        // VK_LEFT or VK_UP
			//case TB_LINEDOWN:      // VK_RIGHT or VK_DOWN
			//case TB_PAGEUP:        // VK_PRIOR (the user clicked the channel above or to the left of the slider)
			//case TB_PAGEDOWN:      // VK_NEXT (the user clicked the channel below or to the right of the slider)
			//case TB_TOP:           // VK_HOME
			//case TB_BOTTOM:        // VK_END
			if (!(control.attrib & GUI_CONTROL_ATTRIB_ALTSUBMIT)) // Ignore this event.
				return;
			// Otherwise:
			gui_event = aNotifyCode + 48; // Signal it to store an ASCII character (digit) in A_GuiControlEvent.
		}
		if (control.output_var)
			ControlGetContents(*control.output_var, control);
		break;

	case GUI_CONTROL_TAB: // aNotifyCode == TCN_SELCHANGE should be the only possibility.
		break; // Must explicitly list this case since the default label does a return.

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



WORD GuiType::TextToHotkey(char *aText)
// Returns a WORD (not a DWORD -- MSDN is wrong about that) compatible with the HKM_SETHOTKEY message:
// LOBYTE is the virtual key.
// HIBYTE is a set of modifiers:
	//HOTKEYF_ALT ALT key
	//HOTKEYF_CONTROL CONTROL key
	//HOTKEYF_SHIFT SHIFT key
	//HOTKEYF_EXT Extended key
{
	BYTE modifiers = 0; // Set default.
	for (bool done = false; *aText; ++aText)
	{
		switch (*aText)
		{
		case '!': modifiers |= HOTKEYF_ALT; break;
		case '^': modifiers |= HOTKEYF_CONTROL; break;
		case '+': modifiers |= HOTKEYF_SHIFT; break;
		default: done = true;  // Some other character type, so it marks the end of the modifiers.
		}
		if (done) // This must be checked prior here otherwise the loop's ++aText will increment one too many.
			break;
	}
	BYTE vk = TextToVK(aText);
    if (!vk)
		return 0;  // Indicate total failure because a hotkey control can't contain just modifiers without a VK.
	// Find out if the HOTKEYF_EXT flag should be set.
	sc_type sc = TextToSC(aText); // Better than g_vk_to_sc[] since g_vk_to_sc has both an a & b to choose from.
	if (!sc) // Since not found above, default to the primary scan code.
		sc = g_vk_to_sc[vk].a;
	if (sc & 0x100) // It's extended.
		modifiers |= HOTKEYF_EXT;
	return MAKEWORD(vk, modifiers);
}



char *GuiType::HotkeyToText(WORD aHotkey, char *aBuf)
// Caller has ensured aBuf is large enough to hold any hotkey name.
// Returns aBuf.
{
	BYTE modifiers = HIBYTE(aHotkey); // In this case, both the VK and the modifiers are bytes, not words.
	char *cp = aBuf;
	if (modifiers & HOTKEYF_SHIFT)
		*cp++ = '+';
	if (modifiers & HOTKEYF_CONTROL)
		*cp++ = '^';
	if (modifiers & HOTKEYF_ALT)
		*cp++ = '!';
	BYTE vk = LOBYTE(aHotkey);

	// For translating the virtual key, the following notes apply:
	// These are unlikely, and in any case don't have a non-extended counterpart so no special handling:
	//g_vk_to_sc[VK_CANCEL].a |= 0x0100; // Ctrl-break
	//g_vk_to_sc[VK_SNAPSHOT].a |= 0x0100;  // PrintScreen

	// These do not have an extended counterpart, i.e. their VK is unique.
	//g_vk_to_sc[VK_DIVIDE].a |= 0x0100; // NumpadDivide (slash)
	//g_vk_to_sc[VK_NUMLOCK].a |= 0x0100;

	// All of the following are handled properly via the scan code logic below:
	//g_vk_to_sc[VK_INSERT].b = g_vk_to_sc[VK_INSERT].a | 0x0100;
	//g_vk_to_sc[VK_PRIOR].b = g_vk_to_sc[VK_PRIOR].a | 0x0100; // PgUp
	//g_vk_to_sc[VK_NEXT].b = g_vk_to_sc[VK_NEXT].a | 0x0100;  // PgDn
	//g_vk_to_sc[VK_HOME].b = g_vk_to_sc[VK_HOME].a | 0x0100;
	//g_vk_to_sc[VK_END].b = g_vk_to_sc[VK_END].a | 0x0100;
	//g_vk_to_sc[VK_UP].b = g_vk_to_sc[VK_UP].a | 0x0100;
	//g_vk_to_sc[VK_DOWN].b = g_vk_to_sc[VK_DOWN].a | 0x0100;
	//g_vk_to_sc[VK_LEFT].b = g_vk_to_sc[VK_LEFT].a | 0x0100;
	//g_vk_to_sc[VK_RIGHT].b = g_vk_to_sc[VK_RIGHT].a | 0x0100;

	// Same note as above but these cannot be typed by the user, only programmatically inserted via
	// initial value of "Gui Add" or via "GuiControl,, MyHotkey, ^Delete":
	//g_vk_to_sc[VK_DELETE].b = g_vk_to_sc[VK_DELETE].a | 0x0100;
	//g_vk_to_sc[VK_RETURN].b = g_vk_to_sc[VK_RETURN].a | 0x0100;
	// Note: NumpadEnter (not Enter) is extended, unlike Home/End/Pgup/PgDn/Arrows, which are
	// NON-extended on the keypad.

	if (modifiers & HOTKEYF_EXT) // Try to find the extended version of this VK if it has two versions.
	{
		sc_type sc = g_vk_to_sc[vk].b;
		if (!(sc & 0x100)) // No "b" scan code at all or it's not extended.  Try "a".
			sc = g_vk_to_sc[vk].a;
		if (sc & 0x100) // It has a non-zero scan code and it's extended.
		{
			SCToKeyName(sc, cp, 100);
			return aBuf;
		}
	}
	// Since above didn't return, use a simple lookup on VK, since it gives preference to non-extended keys.
	VKToKeyName(vk, 0, cp, 100);
	// The above call might be produce an unknown key-name via GetKeyName().  Since it seems so rare and
	// the exact return value (e.g. SC vs. VK) is uncertain/debatable: For now, it seems best to
	// leave it as its native-language name rather than attempting to convert it to an SC or VK that
	// can be compatible with GetKeyState or the Hotkey command:
	//if (!TextToVK(cp))
	//	sprintf(cp, "vk%02X", vk);
	return aBuf;
}



void GuiType::ControlCheckRadioButton(GuiControlType &aControl, GuiIndexType aControlIndex, WPARAM aCheckType)
{
	GuiIndexType radio_start, radio_end;
	FindGroup(aControlIndex, radio_start, radio_end); // Even if the return value is 1, do the below because it ensures things like tabstop are in the right state.
	if (aCheckType == BST_CHECKED)
		// This will check the specified button and uncheck all the others in the group.
		// There is at least one other reason to call CheckRadioButton() rather than doing something
		// manually: It prevents an unwanted firing of the radio's g-label upon WM_ACTIVATE,
		// at least when a radio group is first in the window's z-order and the radio group has
		// an initially selected button:
		CheckRadioButton(mHwnd, GUI_INDEX_TO_ID(radio_start), GUI_INDEX_TO_ID(radio_end - 1), GUI_INDEX_TO_ID(aControlIndex));
	else // Uncheck it.
	{
		// If the group was originally created with the tabstop style, unchecking the button that currently
		// has that style would also remove the tabstop style because apparently that's how radio buttons
		// respond to being unchecked.  Compensate for this by giving the first radio in the group the
		// tabstop style. Update: The below no longer checks to see if the radio group has the tabstop style
		// because it's fairly pointless.  This is because when the user checks/clicks a radio button,
		// the control automatically acquires the tabstop style.  In other words, even though -TabStop is
		// allowed in a radio's options, it will be overridden the first time a user selects a radio button.
		HWND first_radio_in_group = NULL;
		// Find the first radio in this control group:
		for (GuiIndexType u = radio_start; u < radio_end; ++u)
			if (mControl[u].type == GUI_CONTROL_RADIO) // Since a group can have non-radio controls in it.
			{
				first_radio_in_group = mControl[u].hwnd;
				break;
			}
		// The below can't be done until after the above because it would remove the tabstop style
		// if the specified radio happens to be the one that has the tabstop style.  This is usually
		// the case since the button is usually on (since the script is now turning it off here),
		// and other logic has ensured that the on-button is the one with the tabstop style:
		SendMessage(aControl.hwnd, BM_SETCHECK, BST_UNCHECKED, 0);
		if (first_radio_in_group) // Apply the tabstop style to it.
			SetWindowLong(first_radio_in_group, GWL_STYLE, WS_TABSTOP | GetWindowLong(first_radio_in_group, GWL_STYLE));
	}
}



int GuiType::ControlGetDefaultSliderThickness(DWORD aStyle, int aThumbThickness)
{
	if (aThumbThickness <= 0)
		aThumbThickness = 20;  // Set default.
	// Provide a small margin on both sides, otherwise the bar is sometimes truncated.
	aThumbThickness += 5; // 5 looks better than 4 in most styles/themes.
	if (aStyle & TBS_NOTICKS) // This takes precedence over TBS_BOTH (which if present will still make the thumb flat vs. pointed).
		return aThumbThickness;
	if (aStyle & TBS_BOTH)
		return aThumbThickness + 16;
	return aThumbThickness + 8;
}



int GuiType::ControlInvertSliderIfNeeded(GuiControlType &aControl, int aPosition)
// Caller has ensured that aControl.type is slider.
{
	return (aControl.attrib & GUI_CONTROL_ATTRIB_ALTBEHAVIOR)
		? (SendMessage(aControl.hwnd, TBM_GETRANGEMAX, 0, 0) - aPosition) + SendMessage(aControl.hwnd, TBM_GETRANGEMIN, 0, 0)
		: aPosition;  // No inversion necessary.
}



void GuiType::ControlSetSliderOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt)
// Caller has ensured that aControl.type is slider.
{
	if (aOpt.range_min || aOpt.range_max) // Must check like this because although it valid for one to be zero, both should not be.
	{
		// Don't use TBM_SETRANGE because then only 16-bit values are supported:
		SendMessage(aControl.hwnd, TBM_SETRANGEMIN, FALSE, aOpt.range_min); // No redraw
		SendMessage(aControl.hwnd, TBM_SETRANGEMAX, TRUE, aOpt.range_max); // Redraw.
	}
	if (aOpt.tick_interval)
	{
		if (aOpt.tick_interval < 0) // This is the signal to remove the existing tickmarks.
			SendMessage(aControl.hwnd, TBM_CLEARTICS, TRUE, 0);
		else // greater than zero, since zero itself it checked in one of the enclose IFs above.
			SendMessage(aControl.hwnd, TBM_SETTICFREQ, aOpt.tick_interval, 0);
	}
	if (aOpt.line_size > 0) // Removal is not supported, so only positive values are considered.
		SendMessage(aControl.hwnd, TBM_SETLINESIZE, 0, aOpt.line_size);
	if (aOpt.page_size > 0) // Removal is not supported, so only positive values are considered.
		SendMessage(aControl.hwnd, TBM_SETPAGESIZE, 0, aOpt.page_size);
	if (aOpt.thickness > 0)
		SendMessage(aControl.hwnd, TBM_SETTHUMBLENGTH, aOpt.thickness, 0);
	if (aOpt.tip_side)
		SendMessage(aControl.hwnd, TBM_SETTIPSIDE, aOpt.tip_side - 1, 0); // -1 to convert back to zero base.

	// Buddy positioning is left primitive and automatic even when auto-position of this slider
	// or the controls that come after it is in effect.  This is because buddy controls seem too
	// rarely used (due to their lack of positioning options), which is the reason why extra code
	// isn't added here and in "GuiControl Move" to treat the buddies as part of the control rect
	// (i.e. as an entire unit), nor any code for auto-positioning the entire unit when the control
	// is created.  If such code were added, it would require even more code if the slider itself
	// has an automatic position, because the positions of its buddies would have to be recorded,
	// then after they are set as buddies (which moves them) the slider is moved up or left to the
	// position where its buddies used to be.  Otherwise, there would be a gap left during
	// auto-layout.
	// For these, removal is not supported, only changing, since removal seems too rarely needed:
	if (aOpt.buddy1)
		SendMessage(aControl.hwnd, TBM_SETBUDDY, TRUE, (LPARAM)aOpt.buddy1->hwnd);  // Left/top
	if (aOpt.buddy2)
		SendMessage(aControl.hwnd, TBM_SETBUDDY, FALSE, (LPARAM)aOpt.buddy2->hwnd); // Right/bottom
}



void GuiType::ControlSetProgressOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt, DWORD aStyle)
// Caller has ensured that aControl.type is Progress.
// Caller has ensured that aOpt.progress_color_bk is CLR_INVALID if no change should be made to the
// bar's current background color.
{
	// If any options are present that cannot be manifest while a visual theme is in effect, ensure any
	// such theme is removed from the control (currently, once removed it is never put back on):
	// Override the default so that colors/smooth can be manifest even when non-classic theme is in effect.
	if (aControl.union_color != CLR_DEFAULT
		|| !(aOpt.progress_color_bk == CLR_DEFAULT || aOpt.progress_color_bk == CLR_INVALID)
		|| (aStyle & PBS_SMOOTH))
		MySetWindowTheme(aControl.hwnd, L"", L"");

	if (aOpt.range_min || aOpt.range_max) // Must check like this because although it valid for one to be zero, both should not be.
	{
		if (aOpt.range_min >= 0 && aOpt.range_min <= 0xFFFF && aOpt.range_max >= 0 && aOpt.range_max <= 0xFFFF)
			// Since the values fall within the bounds for Win95/NT to support, use the old method
			// in case Win95/NT lacks MSIE 3.0:
			SendMessage(aControl.hwnd, PBM_SETRANGE, 0, MAKELPARAM(aOpt.range_min, aOpt.range_max));
		else
			SendMessage(aControl.hwnd, PBM_SETRANGE32, aOpt.range_min, aOpt.range_max);
	}

	if (aOpt.color_changed)
		SendMessage(aControl.hwnd, PBM_SETBARCOLOR, 0, aControl.union_color);

	switch (aOpt.progress_color_bk)
	{
	case CLR_DEFAULT:
		// If background color is default, mBackgroundColorWin won't take effect if there is a visual theme
		// in effect for this control.  But do the below anyway because we don't want to strip the theme off
		// the control just to make the bar's background match the window or tab control.  But we do want
		// it to match if the theme happens to be absent (due to OS not supporting it, classic theme being
		// in effect, or -theme being in effect):
		SendMessage(aControl.hwnd, PBM_SETBKCOLOR, 0, ControlOverrideBkColor(aControl) ? GetSysColor(COLOR_BTNFACE)
			: mBackgroundColorWin);
		break;
	case CLR_INVALID: // Do nothing in this case because caller didn't want existing bkgnd color changed.
		break;
	default: // Custom background color.  In this case, theme would already have been stripped above.
		SendMessage(aControl.hwnd, PBM_SETBKCOLOR, 0, aOpt.progress_color_bk);
	}
}



bool GuiType::ControlOverrideBkColor(GuiControlType &aControl)
// Caller has ensured that aControl.type is something for which the window's or tab control's background
// should apply (e.g. Progress or Text).
{
	GuiControlType *ptab_control;
	if (!mTabControlCount || !(ptab_control = FindTabControl(aControl.tab_control_index))
		|| !(ptab_control->attrib & GUI_CONTROL_ATTRIB_BACKGROUND_DEFAULT)) // Relies on short-circuit boolean order.
		return false;  // Override not needed because control isn't on a tab, or it's tab has same color as window.
	// Does this control lie mostly inside the tab?  Note that controls can belong to a tab page even though
	// they aren't physically located inside the page.
	RECT overlap_rect, tab_rect, control_rect;
	GetWindowRect(ptab_control->hwnd, &tab_rect);
	GetWindowRect(aControl.hwnd, &control_rect);
	IntersectRect(&overlap_rect, &tab_rect, &control_rect);
	// Returns true if more than 50% of control's area is inside the tab:
	return (overlap_rect.right - overlap_rect.left) * (overlap_rect.bottom - overlap_rect.top)
		> 0.5 * (control_rect.right - control_rect.left) * (control_rect.bottom - control_rect.top);
}



void GuiType::ControlUpdateCurrentTab(GuiControlType &aTabControl, bool aFocusFirstControl)
// Handles the selection of a new tab in a tab control.
{
	int curr_tab_index = TabCtrl_GetCurSel(aTabControl.hwnd);
	if (curr_tab_index == -1) // No tab is selected.  Maybe only happens if the tab control has no tabs at all.
		return;

	// Fix for v1.0.23:
	// If the tab control lacks the visible property, hide all its controls on all its tabs.
	// Don't use IsWindowVisible() because that would say that tab is hidden just because the parent
	// window is hidden, which is not desirable because then when the parent is shown, the shower
	// would always have to remember to call us.  This "hide all" behavior is done here rather than
	// attempting to "rely on everyone else to do their jobs of keeping the controls in the right state"
	// because it improves maintainability:
	DWORD tab_style = GetWindowLong(aTabControl.hwnd, GWL_STYLE);
	bool hide_all = !(tab_style & WS_VISIBLE); // Regardless of whether mHwnd is visible or not.
	bool disable_all = (tab_style & WS_DISABLED); // Don't use IsWindowEnabled() because it might return false if parent is disabled?
	// Say that the focus was already set correctly if the entire tab control is hidden or caller said
	// not to focus it:
	bool focus_was_set;
	bool parent_is_visible = IsWindowVisible(mHwnd);
	bool parent_is_visible_and_not_minimized = parent_is_visible && !IsIconic(mHwnd);
	if (hide_all || disable_all)
		focus_was_set = true;  // Tell the below not to set focus, since all tab controls are hidden or disabled.
	else if (aFocusFirstControl)  // Note that SetFocus() has an effect even if the parent window is hidden. i.e. next time the window is shown, the control will be focused.
		focus_was_set = false; // Tell it to focus the first control on the new page.
	else
	{
		HWND focused_hwnd;
		GuiControlType *focused_control;
		// If the currently focused control is somewhere in this tab control (but not the tab control
		// itself, because arrow-key navigation relies on tabs stay focused while the user is pressing
		// left and right-arrow), override the fact that aFocusFirstControl is false so that when the
		// page changes, its first control will be focused:
		focus_was_set = !(   parent_is_visible && (focused_hwnd = GetFocus())
			&& (focused_control = FindControl(focused_hwnd))
			&& focused_control->tab_control_index == aTabControl.tab_index   );
	}

	bool will_be_visible, will_be_enabled, has_visible_style, has_enabled_style, member_of_current_tab, control_state_altered;
	DWORD style;
	RECT rect, tab_rect;
	POINT *rect_pt = (POINT *)&rect; // i.e. have rect_pt be an array of the two points already within rect.

	GetWindowRect(aTabControl.hwnd, &tab_rect);

	// Update: Don't do the below because it causes a tab to look focused even when it isn't in cases
	// where a control was focused while drawing was suspended.  This is because the below omits the
	// tab rows themselves from the InvalidateRect() further below:
	// Tabs on left (TCS_BUTTONS only) require workaround, at least on XP.  Otherwise tab_rect.left will be
	// much too large.  Because of this, include entire tab rect if it can't be "deflated" reliably:
	//if (!(tab_style & TCS_VERTICAL) || (tab_style & TCS_RIGHT) || !(tab_style & TCS_BUTTONS))
	//	TabCtrl_AdjustRect(aTabControl.hwnd, FALSE, &tab_rect); // Reduce it to just the area without the tabs, since the tabs have already been redrawn.

	// For a likely cleaner transition between tabs, disable redrawing until the switch is complete.
	// Doing it this way also serves to refresh a tab whose controls have just been disabled via
	// something like "GuiControl, Disable, MyTab", which would otherwise not happen because unlike
	// ShowWindow(), EnableWindow() apparently does not cause a repaint to occur.
	// Fix for v1.0.25.14: Don't send the message below (and its counterpart later on) because that
	// sometimes or always, as a side-effect, shows the window if it's hidden:
	if (parent_is_visible_and_not_minimized)
		SendMessage(mHwnd, WM_SETREDRAW, FALSE, 0);
	bool invalidate_entire_parent = false; // Set default.

	// Even if mHwnd is hidden, set styles to Show/Hide and Enable/Disable any controls that need it.
	for (GuiIndexType u = 0; u < mControlCount; ++u)
	{
		// Note aTabControl.tab_index stores aTabControl's tab_control_index (true only for type GUI_CONTROL_TAB).
		if (mControl[u].tab_control_index != aTabControl.tab_index) // This control is not in this tab control.
			continue;
		GuiControlType &control = mControl[u]; // Probably helps performance; certainly improves conciseness.
		member_of_current_tab = (control.tab_index == curr_tab_index);
		will_be_visible = !hide_all && member_of_current_tab && !(control.attrib & GUI_CONTROL_ATTRIB_EXPLICITLY_HIDDEN);
		will_be_enabled = !disable_all && member_of_current_tab && !(control.attrib & GUI_CONTROL_ATTRIB_EXPLICITLY_DISABLED);
		// Don't use IsWindowVisible() because if the parent window is hidden, I think that will
		// always say that the controls are hidden too.  In any case, IsWindowVisible() does not
		// work correctly for this when the window is first shown:
		style = GetWindowLong(control.hwnd, GWL_STYLE);
		has_visible_style =  style & WS_VISIBLE;
		has_enabled_style = !(style & WS_DISABLED);
		// Showing/hiding/enabling/disabling only when necessary might cut down on redrawing:
		control_state_altered = false;  // Set default.
		if (will_be_visible)
		{
			if (!has_visible_style)
			{
				ShowWindow(control.hwnd, SW_SHOWNOACTIVATE);
				control_state_altered = true;
			}
		}
		else
			if (has_visible_style)
			{
				ShowWindow(control.hwnd, SW_HIDE);
				control_state_altered = true;
			}
		if (will_be_enabled)
		{
			if (!has_enabled_style)
			{
				EnableWindow(control.hwnd, TRUE);
				control_state_altered = true;
			}
		}
		else
			if (has_enabled_style)
			{
				// Note that it seems to make sense to disable even text/pic/groupbox controls because
				// they can receive clicks and double clicks (except GroupBox).
				EnableWindow(control.hwnd, FALSE);
				control_state_altered = true;
			}

		if (control_state_altered)
		{
			// If this altered control lies at least partially outside the tab's interior,
			// set it up to do the full repaint of the parent window:
			GetWindowRect(control.hwnd, &rect);
			if (!(PtInRect(&tab_rect, rect_pt[0]) && PtInRect(&tab_rect, rect_pt[1])))
				invalidate_entire_parent = true;
		}
		// The aboves use of show/hide across a wide range of controls may be necessary to support things
		// such as the dynamic removal of tabs via "GuiControl,, MyTab, |NewTabSet1|NewTabSet2", i.e. if the
		// newly added removed tab was active, it's controls should now be hidden.
		// The below sets focus to the first input-capable control, which seems standard for the tab-control
		// dialogs I've seen.
		if (!focus_was_set && member_of_current_tab && will_be_visible && will_be_enabled
			&& GUI_CONTROL_TYPE_CAN_BE_FOCUSED(control.type))
		{
			// Fix for v1.0.24: Don't check the return value of SetFocus() because sometimes it returns
			// NULL even when the call will wind up succeeding.  For example, if the user clicks on
			// the second tab in a tab control, SetFocus() will probably return NULL because there
			// is not previously focused control at the instant the call is made.  This is because
			// the control that had focus has likely already been hidden and thus lost focus before
			// we arrived at this stage:
			SetFocus(control.hwnd); // Note that this has an effect even if the parent window is hidden. i.e. next time the parent is shown, this control will be focused.
			focus_was_set = true; // i.e. SetFocus() only for the FIRST control that meets the above criteria.
		}
	}

	if (parent_is_visible_and_not_minimized) // Fix for v1.0.25.14.  See further above for details.
		SendMessage(mHwnd, WM_SETREDRAW, TRUE, 0); // Re-enable drawing before below so that tab can be focused below.

	// In case tab is empty or there is no control capable of receiving focus, focus the tab itself
	// instead.  This allows the Ctrl-Pgdn/Pgup keyboard shortcuts to continue to navigate within
	// this tab control rather than having the focus get kicked backed outside the tab control
	// -- which I think happens when the tab contains no controls or only text controls (pic controls
	// seem okay for some reason), i.e. if the control with focus is hidden, the dialog falls back to
	// giving the focus to the the first focus-capable control in the z-order.
	if (!focus_was_set)
		SetFocus(aTabControl.hwnd); // Note that this has an effect even if the parent window is hidden. i.e. next time the parent is shown, this control will be focused.

	// UPDATE: Below is now only done when necessary to cut down on flicker:
	// Seems best to invalidate the entire client area because otherwise, if any of the tab's controls lie
	// outside of its interior (this is common for TCS_BUTTONS style), they would not get repainted properly.
	// In addition, tab controls tend to occupy the majority of their parent's client area anyway:
	if (parent_is_visible_and_not_minimized)
	{
		if (invalidate_entire_parent)
			InvalidateRect(mHwnd, NULL, TRUE); // TRUE seems safer.
		else
		{
			MapWindowPoints(NULL, mHwnd, (LPPOINT)&tab_rect, 2); // Convert rect to client coordinates (not the same as GetClientRect()).
			InvalidateRect(mHwnd, &tab_rect, TRUE); // Seems safer to use TRUE, not knowing all possible overlaps, etc.
		}
	}
}



GuiControlType *GuiType::FindTabControl(TabControlIndexType aTabControlIndex)
{
	if (aTabControlIndex == MAX_TAB_CONTROLS)
		// This indicates it's not a member of a tab control. Callers rely on this check.
		return NULL;
	TabControlIndexType tab_control_index = 0;
	for (GuiIndexType u = 0; u < mControlCount; ++u)
		if (mControl[u].type == GUI_CONTROL_TAB)
			if (tab_control_index == aTabControlIndex)
				return &mControl[u];
			else
				++tab_control_index;
	return NULL; // Since above didn't return, indicate failure.
}



int GuiType::FindTabIndexByName(GuiControlType &aTabControl, char *aName)
// Find the first tab in this tab control whose leading-part-of-name matches aName.
// Return int vs. TabIndexType so that failure can be indicated.
{
	int tab_count = TabCtrl_GetItemCount(aTabControl.hwnd);
	if (!tab_count)
		return -1; // No match.
	if (!*aName)
		return 0;  // First item (index 0) matches empty string.
	TCITEM tci;
	tci.mask = TCIF_TEXT;
	char buf[1024];
	tci.pszText = buf;
	tci.cchTextMax = sizeof(buf) - 1; // MSDN example uses -1.
	size_t aName_length = strlen(aName);
	for (int i = 0; i < tab_count; ++i)
		if (TabCtrl_GetItem(aTabControl.hwnd, i, &tci))
			if (!strnicmp(tci.pszText, aName, aName_length))
				return i; // Match found.
	return -1; // No match found.
}



int GuiType::GetControlCountOnTabPage(TabControlIndexType aTabControlIndex, TabIndexType aTabIndex)
{
	int count = 0;
	for (GuiIndexType u = 0; u < mControlCount; ++u)
		if (mControl[u].tab_index == aTabIndex && mControl[u].tab_control_index == aTabControlIndex) // This boolean order helps performance.
			++count;
	return count;
}



POINT GuiType::GetPositionOfTabClientArea(GuiControlType &aTabControl)
// Gets position of tab control relative to parent window's client area.
{
	RECT rect, entire_rect;
	GetWindowRect(aTabControl.hwnd, &entire_rect);
	POINT pt = {entire_rect.left, entire_rect.top};
	ScreenToClient(mHwnd, &pt);
	GetClientRect(aTabControl.hwnd, &rect); // Used because the coordinates of its upper-left corner are (0,0).
	DWORD style = GetWindowLong(aTabControl.hwnd, GWL_STYLE);
	// Tabs on left (TCS_BUTTONS only) require workaround, at least on XP.  Otherwise pt.x will be much too large.
	// This has been confirmed to be true even when theme has been stripped off the tab control.
	bool workaround = !(style & TCS_RIGHT) && (style & (TCS_VERTICAL | TCS_BUTTONS)) == (TCS_VERTICAL | TCS_BUTTONS);
	if (workaround)
		SetWindowLong(aTabControl.hwnd, GWL_STYLE, style & ~TCS_BUTTONS);
	TabCtrl_AdjustRect(aTabControl.hwnd, FALSE, &rect); // Retrieve the area beneath the tabs.
	if (workaround)
	{
		SetWindowLong(aTabControl.hwnd, GWL_STYLE, style);
		pt.x += 5 * TabCtrl_GetRowCount(aTabControl.hwnd); // Adjust for the fact that buttons are wider than tabs.
	}
	pt.x += rect.left - 2;  // -2 because testing shows that X (but not Y) is off by exactly 2.
	pt.y += rect.top;
	return pt;
}



ResultType GuiType::SelectAdjacentTab(GuiControlType &aTabControl, bool aMoveToRight, bool aFocusFirstControl
	, bool aWrapAround)
{
	int tab_count = TabCtrl_GetItemCount(aTabControl.hwnd);
	if (!tab_count)
		return FAIL;
	int selected_tab = TabCtrl_GetCurSel(aTabControl.hwnd);
	if (selected_tab == -1) // Not sure how this can happen in this case (since it has at least one tab).
		selected_tab = aMoveToRight ? 0 : tab_count - 1; // Select the first or last tab.
	else
	{
		if (aMoveToRight) // e.g. Ctrl-PgDn or Ctrl-Tab, right-arrow
		{
			++selected_tab;
			if (selected_tab >= tab_count) // wrap around to the start
			{
				if (!aWrapAround)
					return FAIL; // Indicate that tab was not selected due to non-wrap.
				selected_tab = 0;
			}
		}
		else // Ctrl-PgUp or Ctrl-Shift-Tab
		{
			--selected_tab;
			if (selected_tab < 0) // wrap around to the end
			{
				if (!aWrapAround)
					return FAIL; // Indicate that tab was not selected due to non-wrap.
				selected_tab = tab_count - 1;
			}
		}
	}
	// MSDN: "A tab control does not send a TCN_SELCHANGING or TCN_SELCHANGE notification message
	// when a tab is selected using the TCM_SETCURSEL message."
	TabCtrl_SetCurSel(aTabControl.hwnd, selected_tab);
	ControlUpdateCurrentTab(aTabControl, aFocusFirstControl);
	return OK;
}
