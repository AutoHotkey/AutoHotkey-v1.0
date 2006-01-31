/*
AutoHotkey

Copyright 2003-2006 Chris Mallett (support@autohotkey.com)

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
#include <olectl.h> // for OleLoadPicture()
#include <winioctl.h> // For PREVENT_MEDIA_REMOVAL and CD lock/unlock.
#include "qmath.h" // Used by Transform() [math.h incurs 2k larger code size just for ceil() & floor()]
#include "mt19937ar-cok.h" // for sorting in random order
#include "script.h"
#include "window.h" // for IF_USE_FOREGROUND_WINDOW
#include "application.h" // for MsgSleep()
#include "resources\resource.h"  // For InputBox.

////////////////////
// Window related //
////////////////////

ResultType Line::Splash(char *aOptions, char *aSubText, char *aMainText, char *aTitle, char *aFontName
	, char *aImageFile, bool aSplashImage)
{
	int window_index = 0;  // Set the default window to operate upon (the first).
	char *options, *image_filename = aImageFile;  // Set default.
	bool turn_off = false;
	bool show_it_only = false;
	int bar_pos;
	bool bar_pos_has_been_set = false;
	bool options_consist_of_bar_pos_only = false;

	if (aSplashImage)
	{
		options = aOptions;
		if (*aImageFile)
		{
			char *colon_pos = strchr(aImageFile, ':');
			char *image_filename_omit_leading_whitespace = omit_leading_whitespace(image_filename); // Added in v1.0.38.04 per someone's suggestion.
			if (colon_pos)
			{
				char window_number_str[32];  // Allow extra room in case leading spaces or in hex format, e.g. 0x02
				size_t length_to_copy = colon_pos - aImageFile;
				if (length_to_copy < sizeof(window_number_str))
				{
					strlcpy(window_number_str, aImageFile, length_to_copy + 1);
					if (IsPureNumeric(window_number_str, false, false, true)) // Seems best to allow float at runtime.
					{
						// Note that filenames can start with spaces, so omit_leading_whitespace() is only
						// used if the string is entirely blank:
						image_filename = colon_pos + 1;
						image_filename_omit_leading_whitespace = omit_leading_whitespace(image_filename); // Update to reflect the change above.
						if (!*image_filename_omit_leading_whitespace)
							image_filename = image_filename_omit_leading_whitespace;
						window_index = ATOI(window_number_str) - 1;
						if (window_index < 0 || window_index >= MAX_SPLASHIMAGE_WINDOWS)
							return LineError("Max window number is " MAX_SPLASHIMAGE_WINDOWS_STR "." ERR_ABORT
								, FAIL, aOptions);
					}
				}
			}
			if (!stricmp(image_filename_omit_leading_whitespace, "Off")) // v1.0.38.04: Ignores leading whitespace per someone's suggestion.
				turn_off = true;
			else if (!stricmp(image_filename_omit_leading_whitespace, "Show")) // v1.0.38.04: Ignores leading whitespace per someone's suggestion.
				show_it_only = true;
		}
	}
	else // Progress Window.
	{
		if (   !(options = strchr(aOptions, ':'))   )
			options = aOptions;
		else
		{
			window_index = ATOI(aOptions) - 1;
			if (window_index < 0 || window_index >= MAX_PROGRESS_WINDOWS)
				return LineError("Max window number is " MAX_PROGRESS_WINDOWS_STR "." ERR_ABORT, FAIL, aOptions);
			++options;
		}
		options = omit_leading_whitespace(options); // Added in v1.0.38.04 per someone's suggestion.
		if (!stricmp(options, "Off"))
            turn_off = true;
		else if (!stricmp(options, "Show"))
			show_it_only = true;
		else
		{
			// Allow floats at runtime for flexibility (i.e. in case aOptions was in a variable reference).
			// But still use ATOI for the conversion:
			if (IsPureNumeric(options, true, false, true)) // Negatives are allowed as of v1.0.25.
			{
				bar_pos = ATOI(options);
				bar_pos_has_been_set = true;
				options_consist_of_bar_pos_only = true;
			}
			//else leave it set to the default.
		}
	}

	SplashType &splash = aSplashImage ? g_SplashImage[window_index] : g_Progress[window_index];

	// In case it's possible for the window to get destroyed by other means (WinClose?).
	// Do this only after the above options were set so that the each window's settings
	// will be remembered until such time as "Command, Off" is used:
	if (splash.hwnd && !IsWindow(splash.hwnd))
		splash.hwnd = NULL;

	if (show_it_only)
	{
		if (splash.hwnd && !IsWindowVisible(splash.hwnd))
			ShowWindow(splash.hwnd,  SW_SHOWNOACTIVATE); // See bottom of this function for comments on SW_SHOWNOACTIVATE.
		//else for simplicity, do nothing.
		return OK;
	}

	if (!turn_off && splash.hwnd && !*image_filename && (options_consist_of_bar_pos_only || !*options)) // The "modify existing window" mode is in effect.
	{
		// If there is an existing window, just update its bar position and text.
		// If not, do nothing since we don't have the original text of the window to recreate it.
		// Since this is our thread's window, it shouldn't be necessary to use SendMessageTimeout()
		// since the window cannot be hung since by definition our thread isn't hung.  Also, setting
		// a text item from non-blank to blank is not supported so that elements can be omitted from an
		// update command without changing the text that's in the window.  The script can specify %a_space%
		// to explicitly make an element blank.
		if (!aSplashImage && bar_pos_has_been_set && splash.bar_pos != bar_pos) // Avoid unnecessary redrawing.
		{
			splash.bar_pos = bar_pos;
			if (splash.hwnd_bar)
				SendMessage(splash.hwnd_bar, PBM_SETPOS, (WPARAM)bar_pos, 0);
		}
		// SendMessage() vs. SetWindowText() is used for controls so that tabs are expanded.
		// For simplicity, the hwnd_text1 control is not expanded dynamically if it is currently of
		// height zero.  The user can recreate the window if a different height is needed.
		if (*aMainText && splash.hwnd_text1)
			SendMessage(splash.hwnd_text1, WM_SETTEXT, 0, (LPARAM)(aMainText));
		if (*aSubText)
			SendMessage(splash.hwnd_text2, WM_SETTEXT, 0, (LPARAM)(aSubText));
		if (*aTitle)
			SetWindowText(splash.hwnd, aTitle); // Use the simple method for parent window titles.
		return OK;
	}

	// Otherwise, destroy any existing window first:
	if (splash.hwnd)
		DestroyWindow(splash.hwnd);
	if (splash.hfont1) // Destroy font only after destroying the window that uses it.
		DeleteObject(splash.hfont1);
	if (splash.hfont2)
		DeleteObject(splash.hfont2);
	if (splash.hbrush)
		DeleteObject(splash.hbrush);
	if (splash.pic)
		splash.pic->Release();
	ZeroMemory(&splash, sizeof(splash)); // Set the above and all other fields to zero.

	if (turn_off)
		return OK;
	// Otherwise, the window needs to be created or recreated.

	if (!*aTitle) // Provide default title.
		aTitle = (g_script.mFileName && *g_script.mFileName) ? g_script.mFileName : "";

	// Since there is often just one progress/splash window, and it defaults to always-on-top,
	// it seems best to default owned to be true so that it doesn't get its own task bar button:
	bool owned = true;          // Whether this window is owned by the main window.
	bool centered_main = true;  // Whether the main text is centered.
	bool centered_sub = true;   // Whether the sub text is centered.
	bool initially_hidden = false;  // Whether the window should kept hidden (for later showing by the script).
	int style = WS_DISABLED|WS_POPUP|WS_CAPTION;  // WS_CAPTION implies WS_BORDER
	int exstyle = WS_EX_TOPMOST;
	int xpos = COORD_UNSPECIFIED;
	int ypos = COORD_UNSPECIFIED;
	int range_min = 0, range_max = 0;  // For progress bars.
	int font_size1 = 0; // 0 is the flag to "use default size".
	int font_size2 = 0;
	int font_weight1 = FW_DONTCARE;  // Flag later logic to use default.
	int font_weight2 = FW_DONTCARE;  // Flag later logic to use default.
	COLORREF bar_color = CLR_DEFAULT;
	splash.color_bk = CLR_DEFAULT;
	splash.color_text = CLR_DEFAULT;
	splash.height = COORD_UNSPECIFIED;
	if (aSplashImage)
	{
		#define SPLASH_DEFAULT_WIDTH 300
		splash.width = COORD_UNSPECIFIED;
		splash.object_height = COORD_UNSPECIFIED;
	}
	else // Progress window.
	{
		splash.width = SPLASH_DEFAULT_WIDTH;
		splash.object_height = 20;
	}
	splash.object_width = COORD_UNSPECIFIED;  // Currently only used for SplashImage, not Progress.
	if (*aMainText || *aSubText || !aSplashImage)
	{
		splash.margin_x = 10;
		splash.margin_y = 5;
	}
	else // Displaying only a naked image, so don't use borders.
		splash.margin_x = splash.margin_y = 0;

	for (char *cp2, *cp = options; *cp; ++cp)
	{
		switch(toupper(*cp))
		{
		case 'A':  // Non-Always-on-top.  Synonymous with A0 in early versions.
			// Decided against this enforcement.  In the enhancement mentioned below is ever done (unlikely),
			// it seems that A1 can turn always-on-top on and A0 or A by itself can turn it off:
			//if (cp[1] == '0') // The zero is required to allow for future enhancement: modify attrib. of existing window.
			exstyle &= ~WS_EX_TOPMOST;
			break;
		case 'B': // Borderless and/or Titleless
			style &= ~WS_CAPTION;
			if (cp[1] == '1')
				style |= WS_BORDER;
			else if (cp[1] == '2')
				style |= WS_DLGFRAME;
			break;
		case 'C': // Colors
			if (!cp[1]) // Avoids out-of-bounds when the loop's own ++cp is done.
				break;
			++cp; // Always increment to omit the next char from consideration by the next loop iteration.
			switch(toupper(*cp))
			{
			case 'B': // Bar color.
			case 'T': // Text color.
			case 'W': // Window/Background color.
			{
				char color_str[32];
				strlcpy(color_str, cp + 1, sizeof(color_str));
				char *space_pos = StrChrAny(color_str, " \t");  // space or tab
				if (space_pos)
					*space_pos = '\0';
				//else a color name can still be present if it's at the end of the string.
				COLORREF color = ColorNameToBGR(color_str);
				if (color == CLR_NONE) // A matching color name was not found, so assume it's in hex format.
				{
					if (strlen(color_str) > 6)
						color_str[6] = '\0';  // Shorten to exactly 6 chars, which happens if no space/tab delimiter is present.
					color = rgb_to_bgr(strtol(color_str, NULL, 16));
					// if color_str does not contain something hex-numeric, black (0x00) will be assumed,
					// which seems okay given how rare such a problem would be.
				}
				switch (toupper(*cp))
				{
				case 'B':
					bar_color = color;
					break;
				case 'T':
					splash.color_text = color;
					break;
				case 'W':
					splash.color_bk = color;
					splash.hbrush = CreateSolidBrush(color); // Used for window & control backgrounds.
					break;
				}
				// Skip over the color string to avoid interpreting hex digits or color names as option letters:
				cp += strlen(color_str);
				break;
			}
			default:
				centered_sub = (*cp != '0');
				centered_main = (cp[1] != '0');
			}
		case 'F':
			if (!cp[1]) // Avoids out-of-bounds when the loop's own ++cp is done.
				break;
			++cp; // Always increment to omit the next char from consideration by the next loop iteration.
			switch(toupper(*cp))
			{
			case 'M':
				if ((font_size1 = atoi(cp + 1)) < 0)
					font_size1 = 0;
				break;
			case 'S':
				if ((font_size2 = atoi(cp + 1)) < 0)
					font_size2 = 0;
				break;
			}
			break;
		case 'M': // Movable and (optionally) resizable.
			style &= ~WS_DISABLED;
			if (cp[1] == '1')
				style |= WS_SIZEBOX;
			if (cp[1] == '2')
				style |= WS_SIZEBOX|WS_MINIMIZEBOX|WS_MAXIMIZEBOX|WS_SYSMENU;
			break;
		case 'P': // Starting position of progress bar [v1.0.25]
			bar_pos = atoi(cp + 1);
			bar_pos_has_been_set = true;
			break;
		case 'R': // Range of progress bar [v1.0.25]
			if (!cp[1]) // Ignore it because we don't want cp to ever point to the NULL terminator due to the loop's increment.
				break;
			range_min = ATOI(++cp); // Increment cp to point it to range_min.
			if (cp2 = strchr(cp + 1, '-'))  // +1 to omit the min's minus sign, if it has one.
			{
				cp = cp2;
				if (!cp[1]) // Ignore it because we don't want cp to ever point to the NULL terminator due to the loop's increment.
					break;
				range_max = ATOI(++cp); // Increment cp to point it to range_max, which can be negative as in this example: R-100--50
			}
			break;
		case 'T': // Give it a task bar button by making it a non-owned window.
			owned = false;
			break;
		// For options such as W, X and Y:
		// Use atoi() vs. ATOI() to avoid interpreting something like 0x01B as hex when in fact
		// the B was meant to be an option letter:
		case 'W':
			if (!cp[1]) // Avoids out-of-bounds when the loop's own ++cp is done.
				break;
			++cp; // Always increment to omit the next char from consideration by the next loop iteration.
			switch(toupper(*cp))
			{
			case 'M':
				if ((font_weight1 = atoi(cp + 1)) < 0)
					font_weight1 = 0;
				break;
			case 'S':
				if ((font_weight2 = atoi(cp + 1)) < 0)
					font_weight2 = 0;
				break;
			default:
				splash.width = atoi(cp);
			}
			break;
		case 'H':
			if (!strnicmp(cp, "Hide", 4)) // Hide vs. Hidden is debatable.
			{
				initially_hidden = true;
				cp += 3; // +3 vs. +4 due to the loop's own ++cp.
			}
			else // Allow any width/height to be specified so that the window can be "rolled up" to its title bar:
				splash.height = atoi(cp + 1);
			break;
		case 'X':
			xpos = atoi(cp + 1);
			break;
		case 'Y':
			ypos = atoi(cp + 1);
			break;
		case 'Z':
			if (!cp[1]) // Avoids out-of-bounds when the loop's own ++cp is done.
				break;
			++cp; // Always increment to omit the next char from consideration by the next loop iteration.
			switch(toupper(*cp))
			{
			case 'B':  // for backward compatibility with interim releases of v1.0.14
			case 'H':
				splash.object_height = atoi(cp + 1); // Allow it to be zero or negative to omit the object.
				break;
			case 'W':
				if (aSplashImage)
					splash.object_width = atoi(cp + 1); // Allow it to be zero or negative to omit the object.
				//else for Progress, don't allow width to be changed since a zero would omit the bar.
				break;
			case 'X':
				splash.margin_x = atoi(cp + 1);
				break;
			case 'Y':
				splash.margin_y = atoi(cp + 1);
				break;
			}
			break;
		// Otherwise: Ignore other characters, such as the digits that occur after the P/X/Y option letters.
		} // switch()
	} // for()

	HDC hdc = CreateDC("DISPLAY", NULL, NULL, NULL);
	int pixels_per_point_y = GetDeviceCaps(hdc, LOGPIXELSY);

	// Get name and size of default font.
	HFONT hfont_default = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
	HFONT hfont_old = (HFONT)SelectObject(hdc, hfont_default);
	char default_font_name[65];
	GetTextFace(hdc, sizeof(default_font_name) - 1, default_font_name);
	TEXTMETRIC tm;
	GetTextMetrics(hdc, &tm);
	int default_gui_font_height = tm.tmHeight;

	// If both are zero or less, reset object height/width for maintainability and sizing later.
	// However, if one of them is -1, the caller is asking for that dimension to be auto-calc'd
	// to "keep aspect ratio" with the the other specified dimension:
	if (   splash.object_height < 1 && splash.object_height != COORD_UNSPECIFIED
		&& splash.object_width < 1 && splash.object_width != COORD_UNSPECIFIED
		|| !splash.object_height || !splash.object_width   )
		splash.object_height = splash.object_width = 0;

	// If there's an image, handle it first so that automatic-width can be applied (if no width was specified)
	// for later font calculations:
	HANDLE hfile_image;
	if (aSplashImage && *image_filename && splash.object_height
		&& (hfile_image = CreateFile(image_filename, GENERIC_READ, 0, NULL, OPEN_EXISTING, 0, NULL)) != INVALID_HANDLE_VALUE)
	{
		// If any of these calls fail (rare), just omit the picture from the window.
		DWORD file_size = GetFileSize(hfile_image, NULL);
		if (file_size != -1)
		{
			HGLOBAL hglobal = GlobalAlloc(GMEM_MOVEABLE, file_size); // MSDN: alloc memory based on file size
			if (hglobal)
			{
				LPVOID pdata = GlobalLock(hglobal);
				if (pdata)
				{
					DWORD bytes_to_read = 0;
					// MSDN: read file and store in global memory:
					if (ReadFile(hfile_image, pdata, file_size, &bytes_to_read, NULL))
					{
						// MSDN: create IStream* from global memory:
						LPSTREAM pstm = NULL;
						if (SUCCEEDED(CreateStreamOnHGlobal(hglobal, TRUE, &pstm)) && pstm)
						{
							// MSDN: Create IPicture from image file:
							if (FAILED(OleLoadPicture(pstm, file_size, FALSE, IID_IPicture, (LPVOID *)&splash.pic)))
								splash.pic = NULL;
							pstm->Release();
							long hm_width, hm_height;
							if (splash.object_height == -1 && splash.object_width > 0)
							{
								// Caller wants height calculated based on the specified width (keep aspect ratio).
								splash.pic->get_Width(&hm_width);
								splash.pic->get_Height(&hm_height);
								if (hm_width) // Avoid any chance of divide-by-zero.
									splash.object_height = (int)(((double)hm_height / hm_width) * splash.object_width + .5); // Round.
							}
							else if (splash.object_width == -1 && splash.object_height > 0)
							{
								// Caller wants width calculated based on the specified height (keep aspect ratio).
								splash.pic->get_Width(&hm_width);
								splash.pic->get_Height(&hm_height);
								if (hm_height) // Avoid any chance of divide-by-zero.
									splash.object_width = (int)(((double)hm_width / hm_height) * splash.object_height + .5); // Round.
							}
							else
							{
								if (splash.object_height == COORD_UNSPECIFIED)
								{
									splash.pic->get_Height(&hm_height);
									// Convert himetric to pixels:
									splash.object_height = MulDiv(hm_height, pixels_per_point_y, HIMETRIC_INCH);
								}
								if (splash.object_width == COORD_UNSPECIFIED)
								{
									splash.pic->get_Width(&hm_width);
									splash.object_width = MulDiv(hm_width, GetDeviceCaps(hdc, LOGPIXELSX), HIMETRIC_INCH);
								}
							}
							if (splash.width == COORD_UNSPECIFIED)
								splash.width = splash.object_width + (2 * splash.margin_x);
						}
					}
					GlobalUnlock(hglobal);
				}
			}
		}
		CloseHandle(hfile_image);
	}

	// If width is still unspecified -- which should only happen if it's a SplashImage window with
	// no image, or there was a problem getting the image above -- set it to be the default.
	if (splash.width == COORD_UNSPECIFIED)
		splash.width = SPLASH_DEFAULT_WIDTH;
	// Similarly, object_height is set to zero if the object is not present:
	if (splash.object_height == COORD_UNSPECIFIED)
		splash.object_height = 0;

	// Lay out client area.  If height is COORD_UNSPECIFIED, use a temp value for now until
	// it can be later determined.
	RECT client_rect, draw_rect;
	SetRect(&client_rect, 0, 0, splash.width, splash.height == COORD_UNSPECIFIED ? 500 : splash.height);

	// Create fonts based on specified point sizes.  A zero value for font_size1 & 2 are correctly handled
	// by CreateFont():
	if (*aMainText)
	{
		// If a zero size is specified, it should use the default size.  But the default brought about
		// by passing a zero here is not what the system uses as a default, so instead use a font size
		// that is 25% larger than the default size (since the default size itself is used for aSubtext).
		// On a typical system, the default GUI font's point size is 8, so this will make it 10 by default.
		// Also, it appears that changing the system's font size in Control Panel -> Display -> Appearance
		// does not affect the reported default font size.  Thus, the default is probably 8/9 for most/all
		// XP systems and probably other OSes as well.
		// By specifying PROOF_QUALITY the nearest matching font size should be chosen, which should avoid
		// any scaling artifacts that might be caused if default_gui_font_height is not 8.
		if (   !(splash.hfont1 = CreateFont(font_size1 ? -MulDiv(font_size1, pixels_per_point_y, 72) : (int)(1.25 * default_gui_font_height)
			, 0, 0, 0, font_weight1 ? font_weight1 : FW_SEMIBOLD, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_PRECIS
			, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, FF_DONTCARE, *aFontName ? aFontName : default_font_name))   )
			// Call it again with default font in case above failed due to non-existent aFontName.
			// Update: I don't think this actually does any good, at least on XP, because it appears
			// that CreateFont() does not fail merely due to a non-existent typeface.  But it is kept
			// in case it ever fails for other reasons:
			splash.hfont1 = CreateFont(font_size1 ? -MulDiv(font_size1, pixels_per_point_y, 72) : (int)(1.25 * default_gui_font_height)
				, 0, 0, 0, font_weight1 ? font_weight1 : FW_SEMIBOLD, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_PRECIS
				, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, FF_DONTCARE, default_font_name);
		// To avoid showing a runtime error, fall back to the default font if other font wasn't created:
		SelectObject(hdc, splash.hfont1 ? splash.hfont1 : hfont_default);
		// Calc height of text by taking into account font size, number of lines, and space between lines:
		draw_rect = client_rect;
		draw_rect.left += splash.margin_x;
		draw_rect.right -= splash.margin_x;
		splash.text1_height = DrawText(hdc, aMainText, -1, &draw_rect, DT_CALCRECT | DT_WORDBREAK | DT_EXPANDTABS);
	}
	// else leave the above fields set to the zero defaults.

	if (font_size2 || font_weight2 || aFontName)
		if (   !(splash.hfont2 = CreateFont(-MulDiv(font_size2, pixels_per_point_y, 72), 0, 0, 0
			, font_weight2, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS
			, PROOF_QUALITY, FF_DONTCARE, *aFontName ? aFontName : default_font_name))   )
			// Call it again with default font in case above failed due to non-existent aFontName.
			// Update: I don't think this actually does any good, at least on XP, because it appears
			// that CreateFont() does not fail merely due to a non-existent typeface.  But it is kept
			// in case it ever fails for other reasons:
			if (font_size2 || font_weight2)
				splash.hfont2 = CreateFont(-MulDiv(font_size2, pixels_per_point_y, 72), 0, 0, 0
					, font_weight2, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS
					, PROOF_QUALITY, FF_DONTCARE, default_font_name);
	//else leave it NULL so that hfont_default will be used.

	// The font(s) will be deleted the next time this window is destroyed or recreated,
	// or by the g_script destructor.

	SPLASH_CALC_YPOS  // Calculate the Y position of each control in the window.

	if (splash.height == COORD_UNSPECIFIED)
	{
		// Since the window height was not specified, determine what it should be based on the height
		// of all the controls in the window:
		int subtext_height;
		if (*aSubText)
		{
			SelectObject(hdc, splash.hfont2 ? splash.hfont2 : hfont_default);
			// Calc height of text by taking into account font size, number of lines, and space between lines:
			// Reset unconditionally because the previous DrawText() sometimes alters the rect:
			draw_rect = client_rect;
			draw_rect.left += splash.margin_x;
			draw_rect.right -= splash.margin_x;
			subtext_height = DrawText(hdc, aSubText, -1, &draw_rect, DT_CALCRECT | DT_WORDBREAK);
		}
		else
			subtext_height = 0;
		// For the below: sub_y was previously calc'd to be the top of the subtext control.
		// Also, splash.margin_y is added because the text looks a little better if the window
		// doesn't end immediately beneath it:
		splash.height = subtext_height + sub_y + splash.margin_y;
		client_rect.bottom = splash.height;
	}

	SelectObject(hdc, hfont_old); // Necessary to avoid memory leak.
	if (!DeleteDC(hdc))
		return FAIL;  // Force a failure to detect bugs such as hdc still having a created handle inside.

	// Based on the client area determined above, expand the main_rect to include title bar, borders, etc.
	// If the window has a border or caption this also changes top & left *slightly* from zero.
	RECT main_rect = client_rect;
	AdjustWindowRectEx(&main_rect, style, FALSE, exstyle);
	int main_width = main_rect.right - main_rect.left;  // main.left might be slightly less than zero.
	int main_height = main_rect.bottom - main_rect.top; // main.top might be slightly less than zero.

	RECT work_rect;
	SystemParametersInfo(SPI_GETWORKAREA, 0, &work_rect, 0);  // Get desktop rect excluding task bar.
	int work_width = work_rect.right - work_rect.left;  // Note that "left" won't be zero if task bar is on left!
	int work_height = work_rect.bottom - work_rect.top;  // Note that "top" won't be zero if task bar is on top!

	// Seems best (and easier) to unconditionally restrict window size to the size of the desktop,
	// since most users would probably want that.  This can be overridden by using WinMove afterward.
	if (main_width > work_width)
		main_width = work_width;
	if (main_height > work_height)
		main_height = work_height;

	// Centering doesn't currently handle multi-monitor systems explicitly, since those calculations
	// require API functions that don't exist in Win95/NT (and thus would have to be loaded
	// dynamically to allow the program to launch).  Therefore, windows will likely wind up
	// being centered across the total dimensions of all monitors, which usually results in
	// half being on one monitor and half in the other.  This doesn't seem too terrible and
	// might even be what the user wants in some cases (i.e. for really big windows).
	// See comments above for why work_rect.left and top are added in (they aren't always zero).
	if (xpos == COORD_UNSPECIFIED)
		xpos = work_rect.left + ((work_width - main_width) / 2);  // Don't use splash.width.
	if (ypos == COORD_UNSPECIFIED)
		ypos = work_rect.top + ((work_height - main_height) / 2);  // Don't use splash.width.

	// CREATE Main Splash Window
	// It seems best to make this an unowned window for two reasons:
	// 1) It will get its own task bar icon then, which is usually desirable for cases where
	//    there are several progress/splash windows or the window is monitoring something.
	// 2) The progress/splash window won't prevent the main window from being used (owned windows
	//    prevent their owners from ever becoming active).
	// However, it seems likely that some users would want the above to be configurable,
	// so now there is an option to change this behavior.
	HWND dialog_owner = THREAD_DIALOG_OWNER;  // Resolve macro only once to reduce code size.
	if (!(splash.hwnd = CreateWindowEx(exstyle, WINDOW_CLASS_SPLASH, aTitle, style, xpos, ypos
		, main_width, main_height, owned ? (dialog_owner ? dialog_owner : g_hWnd) : NULL // v1.0.35.01: For flexibility, allow these windows to be owned by GUIs via +OwnDialogs.
		, NULL, g_hInstance, NULL)))
		return FAIL;  // No error msg since so rare.

	if ((style & WS_SYSMENU) || !owned)
	{
		// Setting the small icon puts it in the upper left corner of the dialog window.
		// Setting the big icon makes the dialog show up correctly in the Alt-Tab menu (but big seems to
		// have no effect unless the window is unowned, i.e. it has a button on the task bar).
		LPARAM main_icon = (LPARAM)(g_script.mCustomIcon ? g_script.mCustomIcon
			: LoadIcon(g_hInstance, MAKEINTRESOURCE(IDI_MAIN)));
		if (style & WS_SYSMENU)
			SendMessage(splash.hwnd, WM_SETICON, ICON_SMALL, main_icon);
		if (!owned)
			SendMessage(splash.hwnd, WM_SETICON, ICON_BIG, main_icon);
	}

	// Update client rect in case it was resized due to being too large (above) or in case the OS
	// auto-resized it for some reason.  These updated values are also used by SPLASH_CALC_CTRL_WIDTH
	// to position the static text controls so that text will be centered properly:
	GetClientRect(splash.hwnd, &client_rect);
	splash.height = client_rect.bottom;
	splash.width = client_rect.right;
	int control_width = client_rect.right - (splash.margin_x * 2);

	// CREATE Main label
	if (*aMainText)
	{
		splash.hwnd_text1 = CreateWindowEx(0, "static", aMainText
			, WS_CHILD|WS_VISIBLE|SS_NOPREFIX|(centered_main ? SS_CENTER : SS_LEFT)
			, PROGRESS_MAIN_POS, splash.hwnd, NULL, g_hInstance, NULL);
		SendMessage(splash.hwnd_text1, WM_SETFONT, (WPARAM)(splash.hfont1 ? splash.hfont1 : hfont_default), MAKELPARAM(TRUE, 0));
	}

	if (!aSplashImage && splash.object_height > 0) // Progress window
	{
		// CREATE Progress control (always starts off at its default position as determined by OS/common controls):
		if (splash.hwnd_bar = CreateWindowEx(WS_EX_CLIENTEDGE, PROGRESS_CLASS, NULL, WS_CHILD|WS_VISIBLE|PBS_SMOOTH
			, PROGRESS_BAR_POS, splash.hwnd, NULL, NULL, NULL))
		{
			if (range_min || range_max) // i.e. if both are zero, leave it at the default range, which is 0-100.
			{
				if (range_min > -1 && range_min < 0x10000 && range_max > -1 && range_max < 0x10000)
					// Since the values fall within the bounds for Win95/NT to support, use the old method
					// in case Win95/NT lacks MSIE 3.0:
					SendMessage(splash.hwnd_bar, PBM_SETRANGE, 0, MAKELPARAM(range_min, range_max));
				else
					SendMessage(splash.hwnd_bar, PBM_SETRANGE32, range_min, range_max);
			}


			if (bar_color != CLR_DEFAULT)
			{
				// Remove visual styles so that specified color will be obeyed:
				MySetWindowTheme(splash.hwnd_bar, L"", L"");
				SendMessage(splash.hwnd_bar, PBM_SETBARCOLOR, 0, bar_color); // Set color.
			}
			if (splash.color_bk != CLR_DEFAULT)
				SendMessage(splash.hwnd_bar, PBM_SETBKCOLOR, 0, splash.color_bk); // Set color.
			if (bar_pos_has_been_set) // Note that the window is not yet visible at this stage.
				// This happens when the window doesn't exist and a command such as the following is given:
				// Progress, 50 [, ...].  As of v1.0.25, it also happens via the new 'P' option letter:
				SendMessage(splash.hwnd_bar, PBM_SETPOS, (WPARAM)bar_pos, 0);
			else // Ask the control its starting/default position in case a custom range is in effect.
				bar_pos = (int)SendMessage(splash.hwnd_bar, PBM_GETPOS, 0, 0);
			splash.bar_pos = bar_pos; // Save the current position to avoid redraws when future positions are identical to current.
		}
	}

	// CREATE Sub label
	if (splash.hwnd_text2 = CreateWindowEx(0, "static", aSubText
		, WS_CHILD|WS_VISIBLE|SS_NOPREFIX|(centered_sub ? SS_CENTER : SS_LEFT)
		, PROGRESS_SUB_POS, splash.hwnd, NULL, g_hInstance, NULL))
		SendMessage(splash.hwnd_text2, WM_SETFONT, (WPARAM)(splash.hfont2 ? splash.hfont2 : hfont_default), MAKELPARAM(TRUE, 0));

	// Show it without activating it.  Even with options that allow the window to be activated (such
	// as movable), it seems best to do this to prevent changing the current foreground window, which
	// is usually desirable for progress/splash windows since they should be seen but not be disruptive:
	if (!initially_hidden)
		ShowWindow(splash.hwnd,  SW_SHOWNOACTIVATE);
	return OK;
}



ResultType Line::ToolTip(char *aText, char *aX, char *aY, char *aID)
{
	int window_index = *aID ? ATOI(aID) - 1 : 0;
	if (window_index < 0 || window_index >= MAX_TOOLTIPS)
		return LineError("Max window number is " MAX_TOOLTIPS_STR "." ERR_ABORT, FAIL, aID);
	HWND tip_hwnd = g_hWndToolTip[window_index];

	// Destroy windows except the first (for performance) so that resources/mem are conserved.
	// The first window will be hidden by the TTM_UPDATETIPTEXT message if aText is blank.
	// UPDATE: For simplicity, destroy even the first in this way, because otherwise a script
	// that turns off a non-existent first tooltip window then later turns it on will cause
	// the window to appear in an incorrect position.  Example:
	// ToolTip
	// ToolTip, text, 388, 24
	// Sleep, 1000
	// ToolTip, text, 388, 24
	if (!*aText)
	{
		if (tip_hwnd && IsWindow(tip_hwnd))
			DestroyWindow(tip_hwnd);
		g_hWndToolTip[window_index] = NULL;
		return OK;
	}

	// Use virtual desktop so that tooltip can move onto non-primary monitor in a multi-monitor system:
	RECT dtw;
	GetVirtualDesktopRect(dtw);

	bool one_or_both_coords_unspecified = !*aX || !*aY;
	POINT pt, pt_cursor;
	if (one_or_both_coords_unspecified)
	{
		// Don't call GetCursorPos() unless absolutely needed because it seems to mess
		// up double-click timing, at least on XP.  UPDATE: Is isn't GetCursorPos() that's
		// interfering with double clicks, so it seems it must be the displaying of the ToolTip
		// window itself.
		GetCursorPos(&pt_cursor);
		pt.x = pt_cursor.x + 16;  // Set default spot to be near the mouse cursor.
		pt.y = pt_cursor.y + 16;  // Use 16 to prevent the tooltip from overlapping large cursors.
		// Update: Below is no longer needed due to a better fix further down that handles multi-line tooltips.
		// 20 seems to be about the right amount to prevent it from "warping" to the top of the screen,
		// at least on XP:
		//if (pt.y > dtw.bottom - 20)
		//	pt.y = dtw.bottom - 20;
	}

	RECT rect = {0};
	if ((*aX || *aY) && !(g.CoordMode & COORD_MODE_TOOLTIP)) // Need the rect.
	{
		if (!GetWindowRect(GetForegroundWindow(), &rect))
			return OK;  // Don't bother setting ErrorLevel with this command.
	}
	//else leave all of rect's members initialized to zero.

	// This will also convert from relative to screen coordinates if rect contains non-zero values:
	if (*aX)
		pt.x = ATOI(aX) + rect.left;
	if (*aY)
		pt.y = ATOI(aY) + rect.top;

	TOOLINFO ti = {0};
	ti.cbSize = sizeof(ti) - sizeof(void *); // Fixed for v1.0.36.05: Tooltips fail to work on Win9x and probably NT4/2000 unless the size for the *lpReserved member in _WIN32_WINNT 0x0501 is omitted.
	ti.uFlags = TTF_TRACK;
	ti.lpszText = aText;
	// Note that the ToolTip won't work if ti.hwnd is assigned the HWND from GetDesktopWindow().
	// All of ti's other members are left at NULL/0, including the following:
	//ti.hinst = NULL;
	//ti.uId = 0;
	//ti.rect.left = ti.rect.top = ti.rect.right = ti.rect.bottom = 0;

	// My: This does more harm that good (it causes the cursor to warp from the right side to the left
	// if it gets to close to the right side), so for now, I did a different fix (above) instead:
	//ti.rect.bottom = dtw.bottom;
	//ti.rect.right = dtw.right;
	//ti.rect.top = dtw.top;
	//ti.rect.left = dtw.left;

	// No need to use SendMessageTimeout() since the ToolTip() is owned by our own thread, which
	// (since we're here) we know is not hung or heavily occupied.

	// v1.0.40.12: Added the IsWindow() check below to recreate the tooltip in cases where it was destroyed
	// by external means such as Alt-F4 or WinClose.
	if (!tip_hwnd || !IsWindow(tip_hwnd))
	{
		// This this window has no owner, it won't be automatically destroyed when its owner is.
		// Thus, it will be explicitly by the program's exit function.
		tip_hwnd = g_hWndToolTip[window_index] = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, NULL, TTS_NOPREFIX | TTS_ALWAYSTIP
			, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, NULL, NULL);
		SendMessage(tip_hwnd, TTM_ADDTOOL, 0, (LPARAM)(LPTOOLINFO)&ti);
		// v1.0.21: GetSystemMetrics(SM_CXSCREEN) is used for the maximum width because even on a
		// multi-monitor system, most users would not want a tip window to stretch across multiple monitors:
		SendMessage(tip_hwnd, TTM_SETMAXTIPWIDTH, 0, (LPARAM)GetSystemMetrics(SM_CXSCREEN));
		// Must do these next two when the window is first created, otherwise GetWindowRect() below will retrieve
		// a tooltip window size that is quite a bit taller than it winds up being:
		SendMessage(tip_hwnd, TTM_TRACKPOSITION, 0, (LPARAM)MAKELONG(pt.x, pt.y));
		SendMessage(tip_hwnd, TTM_TRACKACTIVATE, TRUE, (LPARAM)&ti);
	}
	// Bugfix for v1.0.21: The below is now called unconditionally, even if the above newly created the window.
	// If this is not done, the tip window will fail to appear the first time it is invoked, at least when
	// all of the following are true:
	// 1) Windows XP;
	// 2) Common controls v6 (via manifest);
	// 3) "Control Panel >> Display >> Effects >> Use transition >> Fade effect" setting is in effect.
	SendMessage(tip_hwnd, TTM_UPDATETIPTEXT, 0, (LPARAM)&ti);

	RECT ttw = {0};
	GetWindowRect(tip_hwnd, &ttw); // Must be called this late to ensure the tooltip has been created by above.
	int tt_width = ttw.right - ttw.left;
	int tt_height = ttw.bottom - ttw.top;

	// v1.0.21: Revised for multi-monitor support.  I read somewhere that dtw.left can be negative (perhaps
	// if the secondary monitor is to the left of the primary).  So it seems best to assume it is possible:
	if (pt.x + tt_width >= dtw.right)
		pt.x = dtw.right - tt_width - 1;
	if (pt.y + tt_height >= dtw.bottom)
		pt.y = dtw.bottom - tt_height - 1;
	// It seems best not to have each of the below paired with the above.  This is because it allows
	// the flexibility to explicitly move the tooltip above or to the left of the screen.  Such a feat
	// should only be possible if done via explicitly passed-in negative coordinates for aX and/or aY.
	// In other words, it should be impossible for a tooltip window to follow the mouse cursor somewhere
	// off the virtual screen because:
	// 1) The mouse cursor is normally trapped within the bounds of the virtual screen.
	// 2) The tooltip window defaults to appearing South-East of the cursor.  It can only appear
	//    in some other quadrant if jammed against the right or bottom edges of the screen, in which
	//    case it can't be partially above or to the left of the virtual screen unless it's really
	//    huge, which seems very unlikely given that it's limited to the maximum width of the
	//    primary display as set by TTM_SETMAXTIPWIDTH above.
	//else if (pt.x < dtw.left) // Should be impossible for this to happen due to mouse being off the screen.
	//	pt.x = dtw.left;      // But could happen if user explicitly passed in a coord that was too negative.
	//...
	//else if (pt.y < dtw.top)
	//	pt.y = dtw.top;

	if (one_or_both_coords_unspecified)
	{
		// Since Tooltip is being shown at the cursor's coordinates, try to ensure that the above
		// adjustment doesn't result in the cursor being inside the tooltip's window boundaries,
		// since that tends to cause problems such as blocking the tray area (which can make a
		// tootip script impossible to terminate).  Normally, that can only happen in this case
		// (one_or_both_coords_unspecified == true) when the cursor is near the buttom-right
		// corner of the screen (unless the mouse is moving more quickly than the script's
		// ToolTip update-frequency can cope with, but that seems inconsequential since it
		// will adjust when the cursor slows down):
		ttw.left = pt.x;
		ttw.top = pt.y;
		ttw.right = ttw.left + tt_width;
		ttw.bottom = ttw.top + tt_height;
		if (pt_cursor.x >= ttw.left && pt_cursor.x <= ttw.right && pt_cursor.y >= ttw.top && pt_cursor.y <= ttw.bottom)
		{
			// Push the tool tip to the upper-left side, since normally the only way the cursor can
			// be inside its boundaries (when one_or_both_coords_unspecified == true) is when the
			// cursor is near the bottom right corner of the screen.
			pt.x = pt_cursor.x - tt_width - 3;    // Use a small offset since it can't overlap the cursor
			pt.y = pt_cursor.y - tt_height - 3;   // when pushed to the the upper-left side of it.
		}
	}

	SendMessage(tip_hwnd, TTM_TRACKPOSITION, 0, (LPARAM)MAKELONG(pt.x, pt.y));
	// And do a TTM_TRACKACTIVATE even if the tooltip window already existed upon entry to this function,
	// so that in case it was hidden or dismissed while its HWND still exists, it will be shown again:
	SendMessage(tip_hwnd, TTM_TRACKACTIVATE, TRUE, (LPARAM)&ti);
	return OK;
}



ResultType Line::TrayTip(char *aTitle, char *aText, char *aTimeout, char *aOptions)
{
	if (!g_os.IsWin2000orLater()) // Older OSes do not support it, so do nothing.
		return OK;
	NOTIFYICONDATA nic = {0};
	nic.cbSize = sizeof(nic);
	nic.uID = AHK_NOTIFYICON;  // This must match our tray icon's uID or Shell_NotifyIcon() will return failure.
	nic.hWnd = g_hWnd;
	nic.uFlags = NIF_INFO;
	nic.uTimeout = ATOI(aTimeout) * 1000;
	nic.dwInfoFlags = ATOI(aOptions);
	strlcpy(nic.szInfoTitle, aTitle, sizeof(nic.szInfoTitle)); // Empty title omits the title line entirely.
	strlcpy(nic.szInfo, aText, sizeof(nic.szInfo));	// Empty text removes the balloon.
	Shell_NotifyIcon(NIM_MODIFY, &nic);
	return OK; // i.e. never a critical error if it fails.
}



ResultType Line::Transform(char *aCmd, char *aValue1, char *aValue2)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	TransformCmds trans_cmd = ConvertTransformCmd(aCmd);
	// Since command names are validated at load-time, this only happens if the command name
	// was contained in a variable reference.  Since that is very rare, output_var is simply
	// made blank to indicate the problem:
	if (trans_cmd == TRANS_CMD_INVALID)
		return output_var->Assign();

	char buf[32];
	int value32;
	INT64 value64;
	double value_double1, value_double2, multiplier;
	double result_double;
	SymbolType value1_is_pure_numeric, value2_is_pure_numeric;

	#undef DETERMINE_NUMERIC_TYPES
	#define DETERMINE_NUMERIC_TYPES \
		value1_is_pure_numeric = IsPureNumeric(aValue1, true, false, true);\
		value2_is_pure_numeric = IsPureNumeric(aValue2, true, false, true);

	#define EITHER_IS_FLOAT (value1_is_pure_numeric == PURE_FLOAT || value2_is_pure_numeric == PURE_FLOAT)

	// If neither input is float, the result is assigned as an integer (i.e. no decimal places):
	#define ASSIGN_BASED_ON_TYPE \
		DETERMINE_NUMERIC_TYPES \
		if (EITHER_IS_FLOAT) \
			return output_var->Assign(result_double);\
		else\
			return output_var->Assign((INT64)result_double);

	// Have a negative exponent always cause a floating point result:
	#define ASSIGN_BASED_ON_TYPE_POW \
		DETERMINE_NUMERIC_TYPES \
		if (EITHER_IS_FLOAT || value_double2 < 0) \
			return output_var->Assign(result_double);\
		else\
			return output_var->Assign((INT64)result_double);

	#define ASSIGN_BASED_ON_TYPE_SINGLE \
		if (IsPureNumeric(aValue1, true, false, true) == PURE_FLOAT)\
			return output_var->Assign(result_double);\
		else\
			return output_var->Assign((INT64)result_double);

	// If rounding to an integer, ensure the result is stored as an integer:
	#define ASSIGN_BASED_ON_TYPE_SINGLE_ROUND \
		if (IsPureNumeric(aValue1, true, false, true) == PURE_FLOAT && value32 > 0)\
			return output_var->Assign(result_double);\
		else\
			return output_var->Assign((INT64)result_double);

	switch(trans_cmd)
	{
	case TRANS_CMD_ASC:
		if (*aValue1)
			return output_var->Assign((int)(UCHAR)*aValue1); // Cast to UCHAR so that chars above Asc(128) show as positive.
		else
			return output_var->Assign();

	case TRANS_CMD_CHR:
		value32 = ATOI(aValue1);
		if (value32 < 0 || value32 > 255)
			return output_var->Assign();
		else
		{
			*buf = value32;  // Store ASCII value as a single-character string.
			*(buf + 1) = '\0';
			return output_var->Assign(buf);
		}

	case TRANS_CMD_DEREF:
		return Deref(output_var, aValue1);

	case TRANS_CMD_UNICODE:
		int char_count;
		if (output_var->Type() == VAR_CLIPBOARD)
		{
			// Since the output var is the clipboard, the mode is autodetected as the following:
			// Convert aValue1 from UTF-8 to Unicode and put the result onto the clipboard.
			// MSDN: "Windows 95: Under the Microsoft Layer for Unicode, MultiByteToWideChar also
			// supports CP_UTF7 and CP_UTF8."
			// First, get the number of characters needed for the buffer size.  This count includes
			// room for the terminating char:
			if (   !(char_count = MultiByteToWideChar(CP_UTF8, 0, aValue1, -1, NULL, 0))   )
				return output_var->Assign(); // Make output_var (i.e. the clipboard) blank to indicate failure.
			LPVOID clip_buf = g_clip.PrepareForWrite(char_count * sizeof(WCHAR));
			if (!clip_buf)
				return output_var->Assign(); // Make output_var (the clipboard in this case) blank to indicate failure.
			// Perform the conversion:
			if (!MultiByteToWideChar(CP_UTF8, 0, aValue1, -1, (LPWSTR)clip_buf, char_count))
			{
				g_clip.AbortWrite();
				return output_var->Assign(); // Make clipboard blank to indicate failure.
			}
			return g_clip.Commit(CF_UNICODETEXT); // Save as type Unicode. It will display any error that occurs.
		}
		// Otherwise, going in the reverse direction: convert the clipboard contents to UTF-8 and put
		// the result into a normal variable.
		if (!IsClipboardFormatAvailable(CF_UNICODETEXT) || !g_clip.Open()) // Relies on short-circuit boolean order.
			return output_var->Assign(); // Make the (non-clipboard) output_var blank to indicate failure.
		if (   !(g_clip.mClipMemNow = GetClipboardData(CF_UNICODETEXT)) // Relies on short-circuit boolean order.
			|| !(g_clip.mClipMemNowLocked = (char *)GlobalLock(g_clip.mClipMemNow))
			|| !(char_count = WideCharToMultiByte(CP_UTF8, 0, (LPCWSTR)g_clip.mClipMemNowLocked, -1, NULL, 0, NULL, NULL))   )
		{
			// Above finds out how large the contents will be when converted to UTF-8.
			// In this case, it failed to determine the count, perhaps due to Win95 lacking Unicode layer, etc.
			g_clip.Close();
			return output_var->Assign(); // Make the (non-clipboard) output_var blank to indicate failure.
		}
		// Othewise, it found the count.  Set up the output variable, enlarging it if needed:
		if (output_var->Assign(NULL, char_count - 1) != OK) // Don't combine this with the above or below it can return FAIL.
		{
			g_clip.Close();
			return FAIL;  // It already displayed the error.
		}
		// Perform the conversion:
		char_count = WideCharToMultiByte(CP_UTF8, 0, (LPCWSTR)g_clip.mClipMemNowLocked, -1, output_var->Contents(), char_count, NULL, NULL);
		g_clip.Close(); // Close the clipboard and free the memory.
		output_var->Close();  // In case it's the clipboard, though currently it can't be since that would auto-detect as the reverse direction.
		if (!char_count)
			return output_var->Assign(); // Make non-clipboard output_var blank to indicate failure.
		return OK;

	case TRANS_CMD_HTML:
	{
		// These are the encoding-neutral translations for ASC 128 through 255 as shown by Dreamweaver.
		// It's possible that using just the &#number convention (e.g. &#128 through &#255;) would be
		// more appropriate for some users, but that mode can be added in the future if it is ever
		// needed (by passing a mode setting for aValue2):
		// ÄÅÇÉÑÖÜáàâäãåçéèêëíìîïñóòôöõúùûü†°¢£§•¶ß®©™´¨≠ÆØ∞±≤≥¥µ∂∑∏π∫ªºΩæø
		// ¿¡¬√ƒ≈∆«»… ÀÃÕŒœ–—“”‘’÷◊ÿŸ⁄€‹›ﬁﬂ‡·‚„‰ÂÊÁËÈÍÎÏÌÓÔÒÚÛÙıˆ˜¯˘˙˚¸˝˛ˇ
		static const char *sHtml[128] = { // v1.0.40.02: Removed leading '&' and trailing ';' to reduce code size.
			  "euro", "#129", "sbquo", "fnof", "bdquo", "hellip", "dagger", "Dagger"
			, "circ", "permil", "Scaron", "lsaquo", "OElig", "#141", "#381", "#143"
			, "#144", "lsquo", "rsquo", "ldquo", "rdquo", "bull", "ndash", "mdash"
			, "tilde", "trade", "scaron", "rsaquo", "oelig", "#157", "#382", "Yuml"
			, "nbsp", "iexcl", "cent", "pound", "curren", "yen", "brvbar", "sect"
			, "uml", "copy", "ordf", "laquo", "not", "shy", "reg", "macr"
			, "deg", "plusmn", "sup2", "sup3", "acute", "micro", "para", "middot"
			, "cedil", "sup1", "ordm", "raquo", "frac14", "frac12", "frac34", "iquest"
			, "Agrave", "Aacute", "Acirc", "Atilde", "Auml", "Aring", "AElig", "Ccedil"
			, "Egrave", "Eacute", "Ecirc", "Euml", "Igrave", "Iacute", "Icirc", "Iuml"
			, "ETH", "Ntilde", "Ograve", "Oacute", "Ocirc", "Otilde", "Ouml", "times"
			, "Oslash", "Ugrave", "Uacute", "Ucirc", "Uuml", "Yacute", "THORN", "szlig"
			, "agrave", "aacute", "acirc", "atilde", "auml", "aring", "aelig", "ccedil"
			, "egrave", "eacute", "ecirc", "euml", "igrave", "iacute", "icirc", "iuml"
			, "eth", "ntilde", "ograve", "oacute", "ocirc", "otilde", "ouml", "divide"
			, "oslash", "ugrave", "uacute", "ucirc", "uuml", "yacute", "thorn", "yuml"
		};

		// Determine how long the result string will be so that the output variable can be expanded
		// to handle it:
		VarSizeType length;
		UCHAR *ucp;
		for (length = 0, ucp = (UCHAR *)aValue1; *ucp; ++ucp)
		{
			switch(*ucp)
			{
			case '\"':  // &quot;
				length += 6;
				break;
			case '&': // &amp;
			case '\n': // <br>\n
				length += 5;
			case '<': // &lt;
			case '>': // &gt;
				length += 4;
			default:
				if (*ucp > 127)
					length += (VarSizeType)strlen(sHtml[*ucp - 128]) + 2; // +2 for the leading '&' and the trailing ';'.
				else
					++length;
			}
		}

		// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
		// this call will set up the clipboard for writing:
		if (output_var->Assign(NULL, length) != OK)
			return FAIL;  // It already displayed the error.
		char *contents = output_var->Contents();  // For performance and tracking.

		// Translate the text to HTML:
		for (ucp = (UCHAR *)aValue1; *ucp; ++ucp)
		{
			switch(*ucp)
			{
			case '\"':  // &quot;
				strcpy(contents, "&quot;");
				contents += 6;
				break;
			case '&': // &amp;
				strcpy(contents, "&amp;");
				contents += 5;
				break;
			case '\n': // <br>\n
				strcpy(contents, "<br>\n");
				contents += 5;
				break;
			case '<': // &lt;
				strcpy(contents, "&lt;");
				contents += 4;
				break;
			case '>': // &gt;
				strcpy(contents, "&gt;");
				contents += 4;
				break;
			default:
				if (*ucp > 127)
				{
					*contents++ = '&'; // v1.0.40.02
					strcpy(contents, sHtml[*ucp - 128]);
					contents += strlen(contents); // Added as a fix in v1.0.41 (broken in v1.0.40.02).
					*contents++ = ';'; // v1.0.40.02
				}
				else
					*contents++ = *ucp;
			}
		}
		*contents = '\0';  // Terminate the string.
		return output_var->Close();  // In case it's the clipboard.
	}

	case TRANS_CMD_MOD:
		if (   !(value_double2 = ATOF(aValue2))   ) // Divide by zero, set it to be blank to indicate the problem.
			return output_var->Assign();
		// Otherwise:
		result_double = qmathFmod(ATOF(aValue1), value_double2);
		ASSIGN_BASED_ON_TYPE

	case TRANS_CMD_POW:
		// The code here should be kept in sync with the behavior of the POWER operator (**)
		// in ExpandExpression.
		// Currently, a negative aValue1 isn't supported.  The reason for this is that since
		// fractional exponents are supported (e.g. 0.5, which results in the square root),
		// there would have to be some extra detection to ensure that a negative aValue1 is
		// never used with fractional exponent (since the sqrt of a negative is undefined).
		// In addition, qmathPow() doesn't support negatives, returning an unexpectedly large
		// value or -1.#IND00 instead.
		value_double1 = ATOF(aValue1);
		value_double2 = ATOF(aValue2);
		// Zero raised to a negative power is undefined, similar to division-by-zero, and thus treated as a failure.
		if (value_double1 < 0 || (value_double1 == 0.0 && value_double2 < 0))
			return output_var->Assign();  // Return a consistent result (blank) rather than something that varies.
		result_double = qmathPow(value_double1, value_double2);
		ASSIGN_BASED_ON_TYPE_POW

	case TRANS_CMD_EXP:
		return output_var->Assign(qmathExp(ATOF(aValue1)));

	case TRANS_CMD_SQRT:
		value_double1 = ATOF(aValue1);
		if (value_double1 < 0)
			return output_var->Assign();
		return output_var->Assign(qmathSqrt(value_double1));

	case TRANS_CMD_LOG:
		value_double1 = ATOF(aValue1);
		if (value_double1 < 0)
			return output_var->Assign();
		return output_var->Assign(qmathLog10(ATOF(aValue1)));

	case TRANS_CMD_LN:
		value_double1 = ATOF(aValue1);
		if (value_double1 < 0)
			return output_var->Assign();
		return output_var->Assign(qmathLog(ATOF(aValue1)));

	case TRANS_CMD_ROUND:
		// In the future, a string conversion algorithm might be better to avoid the loss
		// of 64-bit integer precision that it currently caused by the use of doubles in
		// the calculation:
		value32 = ATOI(aValue2);
		multiplier = *aValue2 ? qmathPow(10, value32) : 1;
		value_double1 = ATOF(aValue1);
		if (value_double1 >= 0.0)
			result_double = qmathFloor(value_double1 * multiplier + 0.5) / multiplier;
		else
			result_double = qmathCeil(value_double1 * multiplier - 0.5) / multiplier;
		ASSIGN_BASED_ON_TYPE_SINGLE_ROUND

	case TRANS_CMD_CEIL:
	case TRANS_CMD_FLOOR:
		// The code here is similar to that in BIF_FloorCeil(), so maintain them together.
		result_double = ATOF(aValue1);
		result_double = (trans_cmd == TRANS_CMD_FLOOR) ? qmathFloor(result_double) : qmathCeil(result_double);
		return output_var->Assign((__int64)(result_double + (result_double > 0 ? 0.2 : -0.2))); // Fixed for v1.0.40.05: See comments in BIF_FloorCeil() for details.

	case TRANS_CMD_ABS:
	{
		// Seems better to convert as string to avoid loss of 64-bit integer precision
		// that would be caused by conversion to double.  I think this will work even
		// for negative hex numbers that are close to the 64-bit limit since they too have
		// a minus sign when generated by the script (e.g. -0x1).
		//result_double = qmathFabs(ATOF(aValue1));
		//ASSIGN_BASED_ON_TYPE_SINGLE
		char *cp = omit_leading_whitespace(aValue1); // i.e. caller doesn't have to have ltrimmed it.
		if (*cp == '-')
			return output_var->Assign(cp + 1);  // Omit the first minus sign (simple conversion only).
		// Otherwise, no minus sign, so just omit the leading whitespace for consistency:
		return output_var->Assign(cp);
	}

	case TRANS_CMD_SIN:
		return output_var->Assign(qmathSin(ATOF(aValue1)));

	case TRANS_CMD_COS:
		return output_var->Assign(qmathCos(ATOF(aValue1)));

	case TRANS_CMD_TAN:
		return output_var->Assign(qmathTan(ATOF(aValue1)));

	case TRANS_CMD_ASIN:
		value_double1 = ATOF(aValue1);
		if (value_double1 > 1 || value_double1 < -1)
			return output_var->Assign(); // ASin and ACos aren't defined for other values.
		return output_var->Assign(qmathAsin(ATOF(aValue1)));

	case TRANS_CMD_ACOS:
		value_double1 = ATOF(aValue1);
		if (value_double1 > 1 || value_double1 < -1)
			return output_var->Assign(); // ASin and ACos aren't defined for other values.
		return output_var->Assign(qmathAcos(ATOF(aValue1)));

	case TRANS_CMD_ATAN:
		return output_var->Assign(qmathAtan(ATOF(aValue1)));

	// For all of the below bitwise operations:
	// Seems better to convert to signed rather than unsigned so that signed values can
	// be supported.  i.e. it seems better to trade one bit in capacity in order to support
	// negative numbers.  Another reason is that commands such as IfEquals use ATOI64 (signed),
	// so if we were to produce unsigned 64 bit values here, they would be somewhat incompatible
	// with other script operations.
	case TRANS_CMD_BITAND:
		return output_var->Assign(ATOI64(aValue1) & ATOI64(aValue2));

	case TRANS_CMD_BITOR:
		return output_var->Assign(ATOI64(aValue1) | ATOI64(aValue2));

	case TRANS_CMD_BITXOR:
		return output_var->Assign(ATOI64(aValue1) ^ ATOI64(aValue2));

	case TRANS_CMD_BITNOT:
		value64 = ATOI64(aValue1);
		if (value64 < 0 || value64 > UINT_MAX)
			// Treat it as a 64-bit signed value, since no other aspects of the program
			// (e.g. IfEqual) will recognize an unsigned 64 bit number.
			return output_var->Assign(~value64);
		else
			// Treat it as a 32-bit unsigned value when inverting and assigning.  This is
			// because assigning it as a signed value would "convert" it into a 64-bit
			// value, which in turn is caused by the fact that the script sees all negative
			// numbers as 64-bit values (e.g. -1 is 0xFFFFFFFFFFFFFFFF).
			return output_var->Assign(~(DWORD)value64);

	case TRANS_CMD_BITSHIFTLEFT:  // Equivalent to multiplying by 2^value2
		return output_var->Assign(ATOI64(aValue1) << ATOI(aValue2));

	case TRANS_CMD_BITSHIFTRIGHT:  // Equivalent to dividing (integer) by 2^value2
		return output_var->Assign(ATOI64(aValue1) >> ATOI(aValue2));
	}

	return FAIL;  // Never executed (increases maintainability and avoids compiler warning).
}



ResultType Line::Input(char *aOptions, char *aEndKeys, char *aMatchList)
// The aEndKeys string must be modifiable (not constant), since for performance reasons,
// it's allowed to be temporarily altered by this function.
// OVERVIEW:
// Although a script can have many concurrent quasi-threads, there can only be one input
// at a time.  Thus, if an input is ongoing and a new thread starts, and it begins its
// own input, that input should terminate the prior input prior to beginning the new one.
// In a "worst case" scenario, each interrupted quasi-thread could have its own
// input, which is in turn terminated by the thread that interrupts it.  Every time
// this function returns, it must be sure to set g_input.status to INPUT_OFF beforehand.
// This signals the quasi-threads beneath, when they finally return, that their input
// was terminated due to a new input that took precedence.
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
	{
		// No output variable, which due to load-time validation means there are no other args either.
		// This means that the user is specifically canceling the prior input (if any).  Thus, our
		// ErrorLevel here is set to 1 or 0, but the prior input's ErrorLevel will be set to "NewInput"
		// when its quasi-thread is resumed:
		bool prior_input_is_being_terminated = (g_input.status == INPUT_IN_PROGRESS);
		g_input.status = INPUT_OFF;
		return g_ErrorLevel->Assign(prior_input_is_being_terminated ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);
		// Above: It's considered an "error" of sorts when there is no prior input to terminate.
	}

	// Set default in case of early return (we want these to be in effect even if
	// FAIL is returned for our thread, since the underlying thread that had the
	// active input prior to us didn't fail and it it needs to know how its input
	// was terminated):
	g_input.status = INPUT_OFF;

	//////////////////////////////////////////////
	// Set up sparse arrays according to aEndKeys:
	//////////////////////////////////////////////
	UCHAR end_vk[VK_ARRAY_COUNT] = {0};  // A sparse array that indicates which VKs terminate the input.
	UCHAR end_sc[SC_ARRAY_COUNT] = {0};  // A sparse array that indicates which SCs terminate the input.

	char single_char_string[2];
	vk_type vk = 0;
	sc_type sc = 0;
	modLR_type modifiersLR;

	for (; *aEndKeys; ++aEndKeys) // This a modified version of the processing loop used in SendKeys().
	{
		switch (*aEndKeys)
		{
		case '}': break;  // Important that these be ignored.
		case '{':
		{
			char *end_pos = strchr(aEndKeys + 1, '}');
			if (!end_pos)
				break;  // do nothing, just ignore it and continue.
			size_t key_text_length = end_pos - aEndKeys - 1;
			if (!key_text_length)
			{
				if (end_pos[1] == '}')
				{
					// The literal string "{}}" has been encountered, which is interpreted as a single "}".
					++end_pos;
					key_text_length = 1;
				}
				else // Empty braces {} were encountered.
					break;  // do nothing: let it proceed to the }, which will then be ignored.
			}

			*end_pos = '\0';  // temporarily terminate the string here.

			if (vk = TextToVK(aEndKeys + 1, NULL, true))
				end_vk[vk] = END_KEY_ENABLED;
			else
				if (sc = TextToSC(aEndKeys + 1))
					end_sc[sc] = END_KEY_ENABLED;

			*end_pos = '}';  // undo the temporary termination

			aEndKeys = end_pos;  // In prep for aEndKeys++ at the bottom of the loop.
			break;
		}
		default:
			single_char_string[0] = *aEndKeys;
			single_char_string[1] = '\0';
			modifiersLR = 0;  // Init prior to below.
			if (vk = TextToVK(single_char_string, &modifiersLR, true))
			{
				end_vk[vk] |= END_KEY_ENABLED; // Use of |= is essential for cases such as ";:".
				// Insist the shift key be down to form genuinely different symbols --
				// namely punctuation marks -- but not for alphabetic chars.  In the
				// future, an option can be added to the Options param to treat
				// end chars as case sensitive (if there is any demand for that):
				if (!IsCharAlpha(*single_char_string))
				{
					// Now we know it's not alphabetic, and it's not a key whose name
					// is longer than one char such as a function key or numpad number.
					// That leaves mostly just the number keys (top row) and all
					// punctuation chars, which are the ones that we want to be
					// distinguished between shifted and unshifted:
					if (modifiersLR & (MOD_LSHIFT | MOD_RSHIFT))
						end_vk[vk] |= END_KEY_WITH_SHIFT;
					else
						end_vk[vk] |= END_KEY_WITHOUT_SHIFT;
				}
			}
		} // switch()
	} // for()

	/////////////////////////////////////////////////
	// Parse aMatchList into an array of key phrases:
	/////////////////////////////////////////////////
	char **realloc_temp;  // Needed since realloc returns NULL on failure but leaves original block allocated.
	g_input.MatchCount = 0;  // Set default.
	if (*aMatchList)
	{
		// If needed, create the array of pointers that points into MatchBuf to each match phrase:
		if (!g_input.match)
		{
			if (   !(g_input.match = (char **)malloc(INPUT_ARRAY_BLOCK_SIZE * sizeof(char *)))   )
				return LineError(ERR_OUTOFMEM);  // Short msg. since so rare.
			g_input.MatchCountMax = INPUT_ARRAY_BLOCK_SIZE;
		}
		// If needed, create or enlarge the buffer that contains all the match phrases:
		size_t aMatchList_length = strlen(aMatchList);
		size_t space_needed = aMatchList_length + 1;  // +1 for the final zero terminator.
		if (space_needed > g_input.MatchBufSize)
		{
			g_input.MatchBufSize = (UINT)(space_needed > 4096 ? space_needed : 4096);
			if (g_input.MatchBuf) // free the old one since it's too small.
				free(g_input.MatchBuf);
			if (   !(g_input.MatchBuf = (char *)malloc(g_input.MatchBufSize))   )
			{
				g_input.MatchBufSize = 0;
				return LineError(ERR_OUTOFMEM);  // Short msg. since so rare.
			}
		}
		// Copy aMatchList into the match buffer:
		char *source, *dest;
		for (source = aMatchList, dest = g_input.match[g_input.MatchCount] = g_input.MatchBuf
			; *source; ++source)
		{
			if (*source != ',') // Not a comma, so just copy it over.
			{
				*dest++ = *source;
				continue;
			}
			// Otherwise: it's a comma, which becomes the terminator of the previous key phrase unless
			// it's a double comma, in which case it's considered to be part of the previous phrase
			// rather than the next.
			if (*(source + 1) == ',') // double comma
			{
				*dest++ = *source;
				++source;  // Omit the second comma of the pair, i.e. each pair becomes a single literal comma.
				continue;
			}
			// Otherwise, this is a delimiting comma.
			*dest = '\0';
			// If the previous item is blank -- which I think can only happen now if the MatchList
			// begins with an orphaned comma (since two adjacent commas resolve to one literal comma)
			// -- don't add it to the match list:
			if (*g_input.match[g_input.MatchCount])
			{
				++g_input.MatchCount;
				g_input.match[g_input.MatchCount] = ++dest;
				*dest = '\0';  // Init to prevent crash on ophaned comma such as "btw,otoh,"
			}
			if (*(source + 1)) // There is a next element.
			{
				if (g_input.MatchCount >= g_input.MatchCountMax) // Rarely needed, so just realloc() to expand.
				{
					// Expand the array by one block:
					if (   !(realloc_temp = (char **)realloc(g_input.match  // Must use a temp variable.
						, (g_input.MatchCountMax + INPUT_ARRAY_BLOCK_SIZE) * sizeof(char *)))   )
						return LineError(ERR_OUTOFMEM);  // Short msg. since so rare.
					g_input.match = realloc_temp;
					g_input.MatchCountMax += INPUT_ARRAY_BLOCK_SIZE;
				}
			}
		} // for()
		*dest = '\0';  // Terminate the last item.
		// This check is necessary for only a single isolated case: When the match list
		// consists of nothing except a single comma.  See above comment for details:
		if (*g_input.match[g_input.MatchCount]) // i.e. omit empty strings from the match list.
			++g_input.MatchCount;
	}

	// Notes about the below macro:
	// In case the Input timer has already put a WM_TIMER msg in our queue before we killed it,
	// clean out the queue now to avoid any chance that such a WM_TIMER message will take effect
	// later when it would be unexpected and might interfere with this input.  To avoid an
	// unnecessary call to PeekMessage(), which might in turn yield our timeslice to other
	// processes if the CPU is under load (which might be undesirable if this input is
	// time-critical, such as in a game), call GetQueueStatus() to see if there are any timer
	// messages in the queue.  I believe that GetQueueStatus(), unlike PeekMessage(), does not
	// have the nasty/undocumented side-effect of yielding our timeslice, but Google and MSDN
	// are completely devoid of any confirming info on this:
	#define KILL_AND_PURGE_INPUT_TIMER \
	if (g_InputTimerExists)\
	{\
		KILL_INPUT_TIMER \
		if (HIWORD(GetQueueStatus(QS_TIMER)) & QS_TIMER)\
			MsgSleep(-1);\
	}

	// Be sure to get rid of the timer if it exists due to a prior, ongoing input.
	// It seems best to do this only after signaling the hook to start the input
	// so that it's MsgSleep(-1), if it launches a new hotkey or timed subroutine,
	// will be less likely to interrupt us during our setup of the input, i.e.
	// it seems best that we put the input in progress prior to allowing any
	// interruption.  UPDATE: Must do this before changing to INPUT_IN_PROGRESS
	// because otherwise the purging of the timer message might call InputTimeout(),
	// which in turn would set the status immediately to INPUT_TIMED_OUT:
	KILL_AND_PURGE_INPUT_TIMER

	//////////////////////////////////////////////////////////////
	// Initialize buffers and state variables for use by the hook:
	//////////////////////////////////////////////////////////////
	// Set the defaults that will be in effect unless overridden by an item in aOptions:
	g_input.BackspaceIsUndo = true;
	g_input.CaseSensitive = false;
	g_input.IgnoreAHKInput = false;
	g_input.TranscribeModifiedKeys = false;
	g_input.Visible = false;
	g_input.FindAnywhere = false;
	int timeout = 0;  // Set default.
	char input_buf[INPUT_BUFFER_SIZE] = ""; // Will contain the actual input from the user.
	g_input.buffer = input_buf;
	g_input.BufferLength = 0;
	g_input.BufferLengthMax = INPUT_BUFFER_SIZE - 1;

	for (char *cp = aOptions; *cp; ++cp)
	{
		switch(toupper(*cp))
		{
		case 'B':
			g_input.BackspaceIsUndo = false;
			break;
		case 'C':
			g_input.CaseSensitive = true;
			break;
		case 'I':
			g_input.IgnoreAHKInput = true;
			break;
		case 'M':
			g_input.TranscribeModifiedKeys = true;
			break;
		case 'L':
			// Use atoi() vs. ATOI() to avoid interpreting something like 0x01C as hex
			// when in fact the C was meant to be an option letter:
			g_input.BufferLengthMax = atoi(cp + 1);
			if (g_input.BufferLengthMax > INPUT_BUFFER_SIZE - 1)
				g_input.BufferLengthMax = INPUT_BUFFER_SIZE - 1;
			break;
		case 'T':
			// Although ATOF() supports hex, it's been documented in the help file that hex should
			// not be used (see comment above) so if someone does it anyway, some option letters
			// might be misinterpreted:
			timeout = (int)(ATOF(cp + 1) * 1000);
			break;
		case 'V':
			g_input.Visible = true;
			break;
		case '*':
			g_input.FindAnywhere = true;
			break;
		}
	}
	// Point the global addresses to our memory areas on the stack:
	g_input.EndVK = end_vk;
	g_input.EndSC = end_sc;
	g_input.status = INPUT_IN_PROGRESS; // Signal the hook to start the input.
	if (!g_KeybdHook) // Install the hook (if needed) upon first use of this feature.
		Hotkey::InstallKeybdHook();

	// A timer is used rather than monitoring the elapsed time here directly because
	// this script's quasi-thread might be interrupted by a Timer or Hotkey subroutine,
	// which (if it takes a long time) would result in our Input not obeying its timeout.
	// By using an actual timer, the TimerProc() will run when the timer expires regardless
	// of which quasi-thread is active, and it will end our input on schedule:
	if (timeout > 0)
		SET_INPUT_TIMER(timeout < 10 ? 10 : timeout)

	//////////////////////////////////////////////////////////////////
	// Wait for one of the following to terminate our input:
	// 1) The hook (due a match in aEndKeys or aMatchList);
	// 2) A thread that interrupts us with a new Input of its own;
	// 3) The timer we put in effect for our timeout (if we have one).
	//////////////////////////////////////////////////////////////////
	for (;;)
	{
		// Rather than monitoring the timeout here, just wait for the incoming WM_TIMER message
		// to take effect as a TimerProc() call during the MsgSleep():
		MsgSleep();
		if (g_input.status != INPUT_IN_PROGRESS)
			break;
	}

	switch(g_input.status)
	{
	case INPUT_TIMED_OUT:
		g_ErrorLevel->Assign("Timeout"); // Lowercase to match IfMsgBox's string, which might help compiler string pooling.
		break;
	case INPUT_TERMINATED_BY_MATCH:
		g_ErrorLevel->Assign("Match");
		break;
	case INPUT_TERMINATED_BY_ENDKEY:
	{
		char key_name[128] = "EndKey:";
		if (g_input.EndingRequiredShift)
		{
			// Since the only way a shift key can be required in our case is if it's a key whose name
			// is a single char (such as a shifted punctuation mark), use a diff. method to look up the
			// key name based on fact that the shift key was down to terminate the input.  We also know
			// that the key is an EndingVK because there's no way for the shift key to have been
			// required by a scan code based on the logic (above) that builds the end_key arrays.
			// MSDN: "Typically, ToAscii performs the translation based on the virtual-key code.
			// In some cases, however, bit 15 of the uScanCode parameter may be used to distinguish
			// between a key press and a key release. The scan code is used for translating ALT+
			// number key combinations.
			BYTE state[256] = {0};
			state[VK_SHIFT] |= 0x80; // Indicate that the neutral shift key is down for conversion purposes.
			int count = ToAscii(g_input.EndingVK, vk_to_sc(g_input.EndingVK), (PBYTE)&state
				, (LPWORD)(key_name + 7), g_MenuIsVisible ? 1 : 0);
			*(key_name + 7 + count) = '\0';  // Terminate the string.
		}
		else
			g_input.EndedBySC ? SCtoKeyName(g_input.EndingSC, key_name + 7, sizeof(key_name) - 7)
				: VKtoKeyName(g_input.EndingVK, g_input.EndingSC, key_name + 7, sizeof(key_name) - 7);
		g_ErrorLevel->Assign(key_name);
		break;
	}
	case INPUT_LIMIT_REACHED:
		g_ErrorLevel->Assign("Max");
		break;
	default: // Our input was terminated due to a new input in a quasi-thread that interrupted ours.
		g_ErrorLevel->Assign("NewInput");
		break;
	}

	g_input.status = INPUT_OFF;  // See OVERVIEW above for why this must be set prior to returning.

	// In case it ended for reason other than a timeout, in which case the timer is still on:
	KILL_AND_PURGE_INPUT_TIMER

	// Seems ok to assign after the kill/purge above since input_buf is our own stack variable
	// and its contents shouldn't be affected even if KILL_AND_PURGE_INPUT_TIMER's MsgSleep()
	// results in a new thread being created that starts a new Input:
	return output_var->Assign(input_buf);
}



ResultType Line::PerformShowWindow(ActionTypeType aActionType, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
{
	// By design, the WinShow command must always unhide a hidden window, even if the user has
	// specified that hidden windows should not be detected.  So set this now so that
	// DetermineTargetWindow() will make its calls in the right mode:
	bool need_restore = (aActionType == ACT_WINSHOW && !g.DetectHiddenWindows);
	if (need_restore)
		g.DetectHiddenWindows = true;
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (need_restore)
		g.DetectHiddenWindows = false;
	if (!target_window)
		return OK;

	// WinGroup's EnumParentActUponAll() is quite similar to the following, so the two should be
	// maintained together.

	int nCmdShow = SW_INVALID; // Set default.

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
	case ACT_WINMINIMIZE:
		if (IsWindowHung(target_window))
		{
			if (g_os.IsWin2000orLater())
				nCmdShow = SW_FORCEMINIMIZE;
			//else it's not Win2k or later, so don't attempt to minimize hung windows because I
			// have an 80% expectation (i.e. untested) that our thread would hang because the
			// call to ShowWindow() would never return.  I have confirmed that SW_MINIMIZE can
			// lock up our thread on WinXP, which is why we revert to SW_FORCEMINIMIZE above.
		}
		else
			nCmdShow = SW_MINIMIZE;
		break;
	case ACT_WINMAXIMIZE: if (!IsWindowHung(target_window)) nCmdShow = SW_MAXIMIZE; break;
	case ACT_WINRESTORE:  if (!IsWindowHung(target_window)) nCmdShow = SW_RESTORE;  break;
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
		// queue, instead opting for more aggresive measures.  Thus, it seems best
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
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		return OK;
	RECT rect;
	if (!GetWindowRect(target_window, &rect))
		return OK;  // Can't set errorlevel, see above.
	MoveWindow(target_window
		, *aX && stricmp(aX, "default") ? ATOI(aX) : rect.left  // X-position
		, *aY && stricmp(aY, "default") ? ATOI(aY) : rect.top   // Y-position
		, *aWidth && stricmp(aWidth, "default") ? ATOI(aWidth) : rect.right - rect.left
		, *aHeight && stricmp(aHeight, "default") ? ATOI(aHeight) : rect.bottom - rect.top
		, TRUE);  // Do repaint.
	DoWinDelay;
	return OK;
}



ResultType Line::ControlSend(char *aControl, char *aKeysToSend, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText, bool aSendRaw)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		return OK;
	HWND control_window = stricmp(aControl, "ahk_parent") ? ControlExist(target_window, aControl)
		: target_window;
	if (!control_window)
		return OK;
	SendKeys(aKeysToSend, aSendRaw, control_window);
	// But don't do WinDelay because KeyDelay should have been in effect for the above.
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
}



ResultType Line::ControlClick(vk_type aVK, int aClickCount, char *aOptions, char *aControl
	, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		return OK;

	// Set the defaults that will be in effect unless overridden by options:
	KeyEventTypes event_type = KEYDOWNANDUP;
	bool position_mode = false;
	// These default coords can be overridden either by aOptions or aControl's X/Y mode:
	POINT click = {COORD_UNSPECIFIED, COORD_UNSPECIFIED};

	for (char *cp = aOptions; *cp; ++cp)
	{
		switch(toupper(*cp))
		{
		case 'D':
			event_type = KEYDOWN;
			break;
		case 'U':
			event_type = KEYUP;
			break;
		case 'P':
			if (!strnicmp(cp, "Pos", 3))
			{
				cp += 2;  // Add 2 vs. 3 to skip over the rest of the letters in this option word.
				position_mode = true;
			}
			break;
		// For the below:
		// Use atoi() vs. ATOI() to avoid interpreting something like 0x01D as hex
		// when in fact the D was meant to be an option letter:
		case 'X':
			click.x = atoi(cp + 1); // Will be overridden later below if it turns out that position_mode is in effect.
			break;
		case 'Y':
			click.y = atoi(cp + 1); // Will be overridden later below if it turns out that position_mode is in effect.
			break;
		}
	}

	HWND control_window = position_mode ? NULL : ControlExist(target_window, aControl);
	if (!control_window) // Even if position_mode is false, the below is still attempted, as documented.
	{
		// New section for v1.0.24.  But only after the above fails to find a control do we consider
		// whether aControl contains X and Y coordinates.  That way, if a control class happens to be
		// named something like "X1 Y1", it will still be found by giving precedence to class names.
		point_and_hwnd_type pah = {0};
		// Parse the X an Y coordinates in a strict way to reduce ambiguity with control names and also
		// to keep the code simple.
		char *cp = omit_leading_whitespace(aControl);
		if (toupper(*cp) != 'X')
			return OK; // Let ErrorLevel tell the story.
		++cp;
		if (!*cp)
			return OK;
		pah.pt.x = ATOI(cp);
		if (   !(cp = StrChrAny(cp, " \t"))   ) // Find next space or tab (there must be one for it to be considered valid).
			return OK;
		cp = omit_leading_whitespace(cp + 1);
		if (!*cp || toupper(*cp) != 'Y')
			return OK;
		++cp;
		if (!*cp)
			return OK;
		pah.pt.y = ATOI(cp);
		// The passed-in coordinates are always relative to target_window's upper left corner because offering
		// an option for absolute/screen coordinates doesn't seem useful.
		RECT rect;
		GetWindowRect(target_window, &rect);
		pah.pt.x += rect.left; // Convert to screen coordinates.
		pah.pt.y += rect.top;
		EnumChildWindows(target_window, EnumChildFindPoint, (LPARAM)&pah); // Find topmost control containing point.
		// If no control is at this point, try posting the mouse event message(s) directly to the
		// parent window to increase the flexibility of this feature:
		control_window = pah.hwnd_found ? pah.hwnd_found : target_window;
		// Convert click's target coordinates to be relative to the client area of the control or
		// parent window because that is the format required by messages such as WM_LBUTTONDOWN
		// used later below:
		click = pah.pt;
		ScreenToClient(control_window, &click);
	}

	// This is done this late because it seems better to set an ErrorLevel of 1 whenever the target
	// window or control isn't found, or any other error condition occurs above:
	if (aClickCount < 1)
		// Allow this to simply "do nothing", because it increases flexibility
		// in the case where the number of clicks is a dereferenced script variable
		// that may sometimes (by intent) resolve to zero or negative:
		return g_ErrorLevel->Assign(ERRORLEVEL_NONE);

	RECT rect;
	if (click.x == COORD_UNSPECIFIED || click.y == COORD_UNSPECIFIED)
	{
		// The following idea is from AutoIt3. It states: "Get the dimensions of the control so we can click
		// the centre of it" (maybe safer and more natural than 0,0).
		// My: In addition, this is probably better for some large controls (e.g. SysListView32) because
		// clicking at 0,0 might activate a part of the control that is not even visible:
		if (!GetWindowRect(control_window, &rect))
			return OK;  // Let ErrorLevel tell the story.
		if (click.x == COORD_UNSPECIFIED)
			click.x = (rect.right - rect.left) / 2;
		if (click.y == COORD_UNSPECIFIED)
			click.y = (rect.bottom - rect.top) / 2;
	}
	LPARAM lparam = MAKELPARAM(click.x, click.y);

	UINT msg_down, msg_up;
	WPARAM wparam;
	bool vk_is_wheel = aVK == VK_WHEEL_UP || aVK == VK_WHEEL_DOWN;

	if (vk_is_wheel)
	{
		wparam = (aClickCount * ((aVK == VK_WHEEL_UP) ? WHEEL_DELTA : -WHEEL_DELTA)) << 16;  // High order word contains the delta.
		// Make the event more accurate by having the state of the keys reflected in the event.
		// The logical state (not physical state) of the modifier keys is used so that something
		// like this is supported:
		// Send, {ShiftDown}
		// MouseClick, WheelUp
		// Send, {ShiftUp}
		// In addition, if the mouse hook is installed, use its logical mouse button state so that
		// something like this is supported:
		// MouseClick, left, , , , , D  ; Hold down the left mouse button
		// MouseClick, WheelUp
		// MouseClick, left, , , , , U  ; Release the left mouse button.
		// UPDATE: Since the other ControlClick types (such as leftclick) do not reflect these
		// modifiers -- and we want to keep it that way, at least by default, for compatibility
		// reasons -- it seems best for consistency not to do them for WheelUp/Down either.
		// A script option can be added in the future to obey the state of the modifiers:
		//mod_type mod = GetModifierState();
		//if (mod & MOD_SHIFT)
		//	wparam |= MK_SHIFT;
		//if (mod & MOD_CONTROL)
		//	wparam |= MK_CONTROL;
        //if (g_MouseHook)
		//	wparam |= g_mouse_buttons_logical;
	}
	else
	{
		switch (aVK)
		{
			case VK_LBUTTON:  msg_down = WM_LBUTTONDOWN; msg_up = WM_LBUTTONUP; wparam = MK_LBUTTON; break;
			case VK_RBUTTON:  msg_down = WM_RBUTTONDOWN; msg_up = WM_RBUTTONUP; wparam = MK_RBUTTON; break;
			case VK_MBUTTON:  msg_down = WM_MBUTTONDOWN; msg_up = WM_MBUTTONUP; wparam = MK_MBUTTON; break;
			case VK_XBUTTON1: msg_down = WM_XBUTTONDOWN; msg_up = WM_XBUTTONUP; wparam = MK_XBUTTON1; break;
			case VK_XBUTTON2: msg_down = WM_XBUTTONDOWN; msg_up = WM_XBUTTONUP; wparam = MK_XBUTTON2; break;
			default: return OK; // Just do nothing since this should realistically never happen.
		}
	}

	// SetActiveWindow() requires ATTACH_THREAD_INPUT to succeed.  Even though the MSDN docs state
	// that SetActiveWindow() has no effect unless the parent window is foreground, Jon insists
	// that SetActiveWindow() resolved some problems for some users.  In any case, it seems best
	// to do this in case the window really is foreground, in which case MSDN indicates that
	// it will help for certain types of dialogs.
	ATTACH_THREAD_INPUT
	SetActiveWindow(target_window);

	if (vk_is_wheel)
	{
		PostMessage(control_window, WM_MOUSEWHEEL, wparam, lparam);
		DoControlDelay;
	}
	else
	{
		for (int i = 0; i < aClickCount; ++i)
		{
			if (event_type != KEYUP) // It's either down-only or up-and-down so always to the down-event.
			{
				PostMessage(control_window, msg_down, wparam, lparam);
				// Seems best to do this one too, which is what AutoIt3 does also.  User can always reduce
				// ControlDelay to 0 or -1.  Update: Jon says this delay might be causing it to fail in
				// some cases.  Upon reflection, it seems best not to do this anyway because PostMessage()
				// should queue up the message for the app correctly even if it's busy.  Update: But I
				// think the timestamp is available on every posted message, so if some apps check for
				// inhumanly fast clicks (to weed out transients with partial clicks of the mouse, or
				// to detect artificial input), the click might not work.  So it might be better after
				// all to do the delay until it's proven to be problematic (Jon implies that he has
				// no proof yet).  IF THIS IS EVER DISABLED, be sure to do the ControlDelay anyway
				// if event_type == KEYDOWN:
				DoControlDelay;
			}
			if (event_type != KEYDOWN) // It's either up-only or up-and-down so always to the up-event.
			{
				PostMessage(control_window, msg_up, 0, lparam);
				DoControlDelay;
			}
		}
	}

	DETACH_THREAD_INPUT

	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
}



ResultType Line::ControlMove(char *aControl, char *aX, char *aY, char *aWidth, char *aHeight
	, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	HWND control_window = ControlExist(target_window, aControl);
	if (!control_window)
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);

	POINT point;
	point.x = *aX ? ATOI(aX) : COORD_UNSPECIFIED;
	point.y = *aY ? ATOI(aY) : COORD_UNSPECIFIED;

	// First convert the user's given coordinates -- which by default are relative to the window's
	// upper left corner -- to screen coordinates:
	if (point.x != COORD_UNSPECIFIED || point.y != COORD_UNSPECIFIED)
	{
		RECT rect;
		if (!GetWindowRect(target_window, &rect))
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		if (point.x != COORD_UNSPECIFIED)
			point.x += rect.left;
		if (point.y != COORD_UNSPECIFIED)
			point.y += rect.top;
	}

	// If either coordinate is unspecified, put the control's current screen coordinate(s)
	// into point:
	RECT control_rect;
	if (!GetWindowRect(control_window, &control_rect))
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	if (point.x == COORD_UNSPECIFIED)
		point.x = control_rect.left;
	if (point.y == COORD_UNSPECIFIED)
		point.y = control_rect.top;

	// Use the immediate parent since controls can themselves have child controls:
	HWND immediate_parent = GetParent(control_window);
	if (!immediate_parent)
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);

	// Convert from absolute screen coordinates to coordinates used with MoveWindow(),
	// which are relative to control_window's parent's client area:
	if (!ScreenToClient(immediate_parent, &point))
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);

	MoveWindow(control_window
		, point.x
		, point.y
		, *aWidth ? ATOI(aWidth) : control_rect.right - control_rect.left
		, *aHeight ? ATOI(aHeight) : control_rect.bottom - control_rect.top
		, TRUE);  // Do repaint.

	DoControlDelay
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
}



ResultType Line::ControlGetPos(char *aControl, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	Var *output_var_x = ResolveVarOfArg(0);  // Ok if NULL.
	Var *output_var_y = ResolveVarOfArg(1);  // Ok if NULL.
	Var *output_var_width = ResolveVarOfArg(2);  // Ok if NULL.
	Var *output_var_height = ResolveVarOfArg(3);  // Ok if NULL.

	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	HWND control_window = target_window ? ControlExist(target_window, aControl) : NULL;
	if (!control_window)
	{
		if (output_var_x)
			output_var_x->Assign();
		if (output_var_y)
			output_var_y->Assign();
		if (output_var_width)
			output_var_width->Assign();
		if (output_var_height)
			output_var_height->Assign();
		return OK;
	}

	RECT parent_rect, child_rect;
	// Realistically never fails since DetermineTargetWindow() and ControlExist() should always yield
	// valid window handles:
	GetWindowRect(target_window, &parent_rect);
	GetWindowRect(control_window, &child_rect);

	if (output_var_x && !output_var_x->Assign(child_rect.left - parent_rect.left))
		return FAIL;
	if (output_var_y && !output_var_y->Assign(child_rect.top - parent_rect.top))
		return FAIL;
	if (output_var_width && !output_var_width->Assign(child_rect.right - child_rect.left))
		return FAIL;
	if (output_var_height && !output_var_height->Assign(child_rect.bottom - child_rect.top))
		return FAIL;

	return OK;
}



ResultType Line::ControlGetFocus(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	output_var->Assign();  // Set default: blank for the output variable.
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		return OK;  // Let ErrorLevel and the blank output variable tell the story.

	// Unlike many of the other Control commands, this one requires AttachThreadInput().
	ATTACH_THREAD_INPUT

	class_and_hwnd_type cah;
	cah.hwnd = GetFocus();  // Do this now that our thread is attached to the target window's.

	// Very important to detach any threads whose inputs were attached above,
	// prior to returning, otherwise the next attempt to attach thread inputs
	// for these particular windows may result in a hung thread or other
	// undesirable effect:
	DETACH_THREAD_INPUT

	if (!cah.hwnd)
		return OK;  // Let ErrorLevel and the blank output variable tell the story.

	char class_name[WINDOW_CLASS_SIZE];
	cah.class_name = class_name;
	if (!GetClassName(cah.hwnd, class_name, sizeof(class_name) - 5)) // -5 to allow room for sequence number.
		return OK;  // Let ErrorLevel and the blank output variable tell the story.
	
	cah.class_count = 0;  // Init for the below.
	cah.is_found = false; // Same.
	EnumChildWindows(target_window, EnumChildFindSeqNum, (LPARAM)&cah);
	if (!cah.is_found)
		return OK;  // Let ErrorLevel and the blank output variable tell the story.
	// Append the class sequence number onto the class name set the output param to be that value:
	snprintfcat(class_name, sizeof(class_name), "%d", cah.class_count);
	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return output_var->Assign(class_name);
}



BOOL CALLBACK EnumChildFindSeqNum(HWND aWnd, LPARAM lParam)
{
	class_and_hwnd_type &cah = *((class_and_hwnd_type *)lParam);  // For performance and convenience.
	char class_name[WINDOW_CLASS_SIZE];
	if (!GetClassName(aWnd, class_name, sizeof(class_name)))
		return TRUE;  // Continue the enumeration.
	if (!strcmp(class_name, cah.class_name)) // Class names match.
	{
		++cah.class_count;
		if (aWnd == cah.hwnd)  // The caller-specified window has been found.
		{
			cah.is_found = true;
			return FALSE;
		}
	}
	return TRUE; // Continue enumeration until a match is found or there aren't any windows remaining.
}



ResultType Line::ControlFocus(char *aControl, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		return OK;
	HWND control_window = ControlExist(target_window, aControl);
	if (!control_window)
		return OK;

	// Unlike many of the other Control commands, this one requires AttachThreadInput()
	// to have any realistic chance of success (though sometimes it may work by pure
	// chance even without it):
	ATTACH_THREAD_INPUT

	if (SetFocus(control_window))
	{
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		DoControlDelay;
	}

	// Very important to detach any threads whose inputs were attached above,
	// prior to returning, otherwise the next attempt to attach thread inputs
	// for these particular windows may result in a hung thread or other
	// undesirable effect:
	DETACH_THREAD_INPUT

	return OK;
}



ResultType Line::ControlSetText(char *aControl, char *aNewText, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
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
	DoControlDelay;
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
}



ResultType Line::ControlGetText(char *aControl, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR);  // Set default.
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	HWND control_window = target_window ? ControlExist(target_window, aControl) : NULL;
	// Even if control_window is NULL, we want to continue on so that the output
	// param is set to be the empty string, which is the proper thing to do
	// rather than leaving whatever was in there before.

	// Handle the output parameter.  This section is similar to that in
	// PerformAssign().  Note: Using GetWindowTextTimeout() vs. GetWindowText()
	// because it is able to get text from more types of controls (e.g. large edit controls):
	VarSizeType space_needed = control_window ? GetWindowTextTimeout(control_window) + 1 : 1; // 1 for terminator.
	if (space_needed > g_MaxVarCapacity) // Allow the command to succeed by truncating the text.
		space_needed = g_MaxVarCapacity;

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
		if (   !(output_var->Length() = (VarSizeType)GetWindowTextTimeout(control_window
			, output_var->Contents(), space_needed))   ) // There was no text to get or GetWindowTextTimeout() failed.
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



ResultType Line::ControlGetListView(Var &aOutputVar, HWND aHwnd, char *aOptions)
// Called by ControlGet() below.  It has ensured that aHwnd is a valid handle to a ListView.
// It has also initialized g_ErrorLevel to be ERRORLEVEL_ERROR, which will be overridden
// if we succeed here.
{
	aOutputVar.Assign(); // Init to blank in case of early return.  Caller has already initialized g_ErrorLevel for us.

	// GET ROW COUNT
	LRESULT row_count;
	if (!SendMessageTimeout(aHwnd, LVM_GETITEMCOUNT, 0, 0, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&row_count)) // Timed out or failed.
		return OK;  // Let ErrorLevel tell the story.

	// GET COLUMN COUNT
	// Through testing, could probably get to a level of 90% certainty that a ListView for which
	// InsertColumn() was never called (or was called only once) might lack a header control if the LV is
	// created in List/Icon view-mode and/or with LVS_NOCOLUMNHEADER. The problem is that 90% doesn't
	// seem to be enough to justify elimination of the code for "undetermined column count" mode.  If it
	// ever does become a certainty, the following could be changed:
	// 1) The extra code for "undetermined" mode rather than simply forcing col_count to be 1.
	// 2) Probably should be kept for compatibility: -1 being returned when undetermined "col count".
	//
	// The following approach might be the only simple yet reliable way to get the column count (sending
	// LVM_GETITEM until it returns false doesn't work because it apparently returns true even for
	// nonexistent subitems -- the same is reported to happen with LVM_GETCOLUMN and such, though I seem
	// to remember that LVM_SETCOLUMN fails on non-existent columns -- but calling that on a ListView
	// that isn't in Report view has been known to traumatize the control).
	// Fix for v1.0.37.01: It appears that the header doesn't always exist.  For example, when an
	// Explorer window opens and is *initially* in icon or list mode vs. details/tiles mode, testing
	// shows that there is no header control.  Testing also shows that there is exactly one column
	// in such cases but only for Explorer and other things that avoid creating the invisible columns.
	// For example, a script can create a ListView in Icon-mode and give it retrievable column data for
	// columns beyond the first.  Thus, having the undetermined-col-count mode preserves flexibility
	// by allowing individual columns beyond the first to be retrieved.  On a related note, testing shows
	// that attempts to explicitly retrieve columns (i.e. fields/subitems) other than the first in the
	// case of Explorer's Icon/List view modes behave the same as fetching the first column (i.e. Col3
	// would retrieve the same text as specifying Col1 or not having the Col option at all).
	// Obsolete because not always true: Testing shows that a ListView always has a header control
	// (at least on XP), even if you can't see it (such as when the view is Icon/Tile or when -Hdr has
	// been specified in the options).
	HWND header_control;
	LRESULT col_count = -1;  // Fix for v1.0.37.01: Use -1 to indicate "undetermined col count".
	if (SendMessageTimeout(aHwnd, LVM_GETHEADER, 0, 0, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&header_control)
		&& header_control) // Relies on short-circuit boolean order.
		SendMessageTimeout(header_control, HDM_GETITEMCOUNT, 0, 0, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&col_count);
		// Return value is not checked because if it fails, col_count is left at its default of -1 set above.
		// In fact, if any of the above conditions made it impossible to determine col_count, col_count stays
		// at -1 to indicate "undetermined".

	// PARSE OPTIONS (a simple vs. strict method is used to reduce code size)
	bool get_count = strcasestr(aOptions, "Count");
	bool include_selected_only = strcasestr(aOptions, "Selected"); // Explicit "ed" to reserve "Select" for possible future use.
	bool include_focused_only = strcasestr(aOptions, "Focused");  // Same.
	char *col_option = strcasestr(aOptions, "Col"); // Also used for mode "Count Col"
	int requested_col = col_option ? ATOI(col_option + 3) - 1 : -1;
	// If the above yields a negative col number for any reason, it's ok because below will just ignore it.
	if (col_count > -1 && requested_col > -1 && requested_col >= col_count) // Specified column does not exist.
		return OK;  // Let ErrorLevel tell the story.

	// IF THE "COUNT" OPTION IS PRESENT, FULLY HANDLE THAT AND RETURN
	if (get_count)
	{
		int result; // Must be signed to support writing a col count of -1 to aOutputVar.
		if (include_focused_only) // Listed first so that it takes precedence over include_selected_only.
		{
			if (!SendMessageTimeout(aHwnd, LVM_GETNEXTITEM, -1, LVNI_FOCUSED, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&result)) // Timed out or failed.
				return OK;  // Let ErrorLevel tell the story.
			++result; // i.e. Set it to 0 if not found, or the 1-based row-number otherwise.
		}
		else if (include_selected_only)
		{
			if (!SendMessageTimeout(aHwnd, LVM_GETSELECTEDCOUNT, 0, 0, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&result)) // Timed out or failed.
				return OK;  // Let ErrorLevel tell the story.
		}
		else if (col_option) // "Count Col" returns the number of columns.
			result = (int)col_count;
		else // Total row count.
			result = (int)row_count;
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		return aOutputVar.Assign(result);
	}

	// FINAL CHECKS
	if (row_count < 1 || !col_count) // But don't return when col_count == -1 (i.e. always make the attempt when col count is undetermined).
		return g_ErrorLevel->Assign(ERRORLEVEL_NONE);  // No text in the control, so indicate success.

	// ALLOCATE INTERPROCESS MEMORY FOR TEXT RETRIEVAL
	HANDLE handle;
	LPVOID p_remote_lvi; // Not of type LPLVITEM to help catch bugs where p_remote_lvi->member is wrongly accessed here in our process.
	if (   !(p_remote_lvi = AllocInterProcMem(handle, LV_REMOTE_BUF_SIZE + sizeof(LVITEM), aHwnd))   ) // Allocate the right type of memory (depending on OS type).
		return OK;  // Let ErrorLevel tell the story.
	bool is_win9x = g_os.IsWin9x(); // Resolve once for possible slight perf./code size benefit.

	// PREPARE LVI STRUCT MEMBERS FOR TEXT RETRIEVAL
	LVITEM lvi_for_nt; // Only used for NT/2k/XP method.
	LVITEM &local_lvi = is_win9x ? *(LPLVITEM)p_remote_lvi : lvi_for_nt; // Local is the same as remote for Win9x.
	// Subtract 1 because of that nagging doubt about size vs. length. Some MSDN examples subtract one,
	// such as TabCtrl_GetItem()'s cchTextMax:
	local_lvi.cchTextMax = LV_REMOTE_BUF_SIZE - 1; // Note that LVM_GETITEM doesn't update this member to reflect the new length.
	local_lvi.pszText = (char *)p_remote_lvi + sizeof(LVITEM); // The next buffer is the memory area adjacent to, but after the struct.

	LRESULT i, next, length, total_length;
	bool is_selective = include_focused_only || include_selected_only;
	bool single_col_mode = (requested_col > -1 || col_count == -1); // Get only one column in these cases.

	// ESTIMATE THE AMOUNT OF MEMORY NEEDED TO STORE ALL THE TEXT
	// It's important to note that a ListView might legitimately have a collection of rows whose
	// fields are all empty.  Since it is difficult to know whether the control is truly owner-drawn
	// (checking its style might not be enough?), there is no way to distinguish this condition
	// from one where the control's text can't be retrieved due to being owner-drawn.  In any case,
	// this all-empty-field behavior simplifies the code and will be documented in the help file.
	for (i = 0, next = -1, total_length = 0; i < row_count; ++i) // For each row:
	{
		if (is_selective)
		{
			// Fix for v1.0.37.01: Prevent an infinite loop that might occur if the target control no longer
			// exists (perhaps having been closed in the middle of the operation) or is permanently hung.
			// If GetLastError() were to return zero after the below, it would mean the function timed out.
			// However, rather than checking and retrying, it seems better to abort the operation because:
			// 1) Timeout should be quite rare.
			// 2) Reduces code size.
			// 3) Having a retry really should be accompanied by SLEEP_WITHOUT_INTERRUPTION because all this
			//    time our thread would not pumping messages (and worse, if the keyboard/mouse hooks are installed,
			//    mouse/key lag would occur).
			if (!SendMessageTimeout(aHwnd, LVM_GETNEXTITEM, next, include_focused_only ? LVNI_FOCUSED : LVNI_SELECTED
				, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&next) // Timed out or failed.
				|| next == -1) // No next item.  Relies on short-circuit boolean order.
				break; // End of estimation phase (if estimate is too small, the text retrieval below will truncate it).
		}
		else
			next = i;
		for (local_lvi.iSubItem = (requested_col > -1) ? requested_col : 0 // iSubItem is which field to fetch. If it's zero, the item vs. subitem will be fetched.
			; col_count == -1 || local_lvi.iSubItem < col_count // If column count is undetermined (-1), always make the attempt.
			; ++local_lvi.iSubItem) // For each column:
		{
			if ((is_win9x || WriteProcessMemory(handle, p_remote_lvi, &local_lvi, sizeof(LVITEM), NULL)) // Relies on short-circuit boolean order.
				&& SendMessageTimeout(aHwnd, LVM_GETITEMTEXT, next, (LPARAM)p_remote_lvi, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&length))
				total_length += length;
			//else timed out or failed, don't include the length in the estimate.  Instead, the
			// text-fetching routine below will ensure the text doesn't overflow the var capacity.
			if (single_col_mode)
				break;
		}
	}
	// Add to total_length enough room for one linefeed per row, and one tab after each column
	// except the last (formula verified correct, though it's inflated by 1 for safety). "i" contains the
	// actual number of rows that will be transcribed, which might be less than row_count if is_selective==true.
	total_length += i * (single_col_mode ? 1 : col_count);

	// SET UP THE OUTPUT VARIABLE, ENLARGING IT IF NECESSARY
	// If the aOutputVar is of type VAR_CLIPBOARD, this call will set up the clipboard for writing:
	aOutputVar.Assign(NULL, (VarSizeType)total_length); // Since failure is extremely rare, continue onward using the available capacity.
	char *contents = aOutputVar.Contents();
	LRESULT capacity = (int)aOutputVar.Capacity(); // LRESULT avoids signed vs. unsigned compiler warnings.
	if (capacity > 0) // For maintainability, avoid going negative.
		--capacity; // Adjust to exclude the zero terminator, which simplifies things below.

	// RETRIEVE THE TEXT FROM THE REMOTE LISTVIEW
	// Start total_length at zero in case actual size is greater than estimate, in which case only a partial set of text along with its '\t' and '\n' chars will be written.
	for (i = 0, next = -1, total_length = 0; i < row_count; ++i) // For each row:
	{
		if (is_selective)
		{
			// Fix for v1.0.37.01: Prevent an infinite loop (for details, see comments in the estimation phase above).
			if (!SendMessageTimeout(aHwnd, LVM_GETNEXTITEM, next, include_focused_only ? LVNI_FOCUSED : LVNI_SELECTED
				, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&next) // Timed out or failed.
				|| next == -1) // No next item.
				break; // See comment above for why unconditional break vs. continue.
		}
		else // Retrieve every row, so the "next" row becomes the "i" index.
			next = i;
		// Insert a linefeed before each row except the first:
		if (i && total_length < capacity) // If we're at capacity, it will exit the loops when the next field is read.
		{
			*contents++ = '\n';
			++total_length;
		}

		// iSubItem is which field to fetch. If it's zero, the item vs. subitem will be fetched:
		for (local_lvi.iSubItem = (requested_col > -1) ? requested_col : 0
			; col_count == -1 || local_lvi.iSubItem < col_count // If column count is undetermined (-1), always make the attempt.
			; ++local_lvi.iSubItem) // For each column:
		{
			// Insert a tab before each column except the first and except when in single-column mode:
			if (!single_col_mode && local_lvi.iSubItem && total_length < capacity)  // If we're at capacity, it will exit the loops when the next field is read.
			{
				*contents++ = '\t';
				++total_length;
			}

			if (!(is_win9x || WriteProcessMemory(handle, p_remote_lvi, &local_lvi, sizeof(LVITEM), NULL)) // Relies on short-circuit boolean order.
				|| !SendMessageTimeout(aHwnd, LVM_GETITEMTEXT, next, (LPARAM)p_remote_lvi, SMTO_ABORTIFHUNG, 2000, (PDWORD_PTR)&length))
				continue; // Timed out or failed. It seems more useful to continue getting text rather than aborting the operation.

			// Otherwise, the message was successfully sent.
			if (length > 0)
			{
				if (total_length + length > capacity)
					goto break_both; // "goto" for simplicity and code size reduction.
				// Otherwise:
				// READ THE TEXT FROM THE REMOTE PROCESS
				// Although MSDN has the following comment about LVM_GETITEM, it is not present for
				// LVM_GETITEMTEXT. Therefore, to improve performance (by avoiding a second call to
				// ReadProcessMemory) and to reduce code size, we'll take them at their word until
				// proven otherwise.  Here is the MSDN comment about LVM_GETITEM: "Applications
				// should not assume that the text will necessarily be placed in the specified
				// buffer. The control may instead change the pszText member of the structure
				// to point to the new text, rather than place it in the buffer."
				if (is_win9x)
				{
					memcpy(contents, local_lvi.pszText, length); // Usually benches a little faster than strcpy().
					contents += length; // Point it to the position where the next char will be written.
					total_length += length; // Recalculate length in case its different than the estimate (for any reason).
				}
				else
				{
					if (ReadProcessMemory(handle, local_lvi.pszText, contents, length, NULL)) // local_lvi.pszText == p_remote_lvi->pszText
					{
						contents += length; // Point it to the position where the next char will be written.
						total_length += length; // Recalculate length in case its different than the estimate (for any reason).
					}
					//else it failed; but even so, continue on to put in a tab (if called for).
				}
			}
			//else length is zero; but even so, continue on to put in a tab (if called for).
			if (single_col_mode)
				break;
		} // for() each column
	} // for() each row

break_both:
	if (contents) // Might be NULL if Assign() failed and thus var has zero capacity.
		*contents = '\0'; // Final termination.  Above has reserved room for for this one byte.

	// CLEAN UP
	FreeInterProcMem(handle, p_remote_lvi);
	aOutputVar.Close(); // In case it's the clipboard.
	aOutputVar.Length() = (VarSizeType)total_length; // Update to actual vs. estimated length.
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE);  // Indicate success.
}



ResultType Line::StatusBarGetText(char *aPart, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	// Note: ErrorLevel is handled by StatusBarUtil(), below.
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	HWND control_window = target_window ? ControlExist(target_window, "msctls_statusbar321") : NULL;
	// Call this even if control_window is NULL because in that case, it will set the output var to
	// be blank for us:
	return StatusBarUtil(output_var, control_window, ATOI(aPart)); // It will handle any zero part# for us.
}



ResultType Line::StatusBarWait(char *aTextToWaitFor, char *aSeconds, char *aPart, char *aTitle, char *aText
	, char *aInterval, char *aExcludeTitle, char *aExcludeText)
{
	// Note: ErrorLevel is handled by StatusBarUtil(), below.
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	// Make a copy of any memory areas that are volatile (due to Deref buf being overwritten
	// if a new hotkey subroutine is launched while we are waiting) but whose contents we
	// need to refer to while we are waiting:
	char text_to_wait_for[4096];
	strlcpy(text_to_wait_for, aTextToWaitFor, sizeof(text_to_wait_for));
	HWND control_window = target_window ? ControlExist(target_window, "msctls_statusbar321") : NULL;
	return StatusBarUtil(NULL, control_window, ATOI(aPart) // It will handle a NULL control_window or zero part# for us.
		, text_to_wait_for, *aSeconds ? (int)(ATOF(aSeconds)*1000) : -1 // Blank->indefinite.  0 means 500ms.
		, ATOI(aInterval));
}



ResultType Line::ScriptPostSendMessage(bool aUseSend)
// Arg list:
// sArgDeref[0]: Msg number
// sArgDeref[1]: wParam
// sArgDeref[2]: lParam
// sArgDeref[3]: Control
// sArgDeref[4]: WinTitle
// sArgDeref[5]: WinText
// sArgDeref[6]: ExcludeTitle
// sArgDeref[7]: ExcludeText
{
	HWND target_window, control_window;
	if (   !(target_window = DetermineTargetWindow(sArgDeref[4], sArgDeref[5], sArgDeref[6], sArgDeref[7]))
		|| !(control_window = *sArgDeref[3] ? ControlExist(target_window, sArgDeref[3]) : target_window)   ) // Relies on short-circuit boolean order.
		return g_ErrorLevel->Assign(aUseSend ? "FAIL" : ERRORLEVEL_ERROR); // Need a special value to distinguish this from numeric reply-values.
	UINT msg = ATOU(sArgDeref[0]);
	// UPDATE: Note that ATOU(), in both past and current versions, supports negative numbers too.
	// For example, ATOU("-1") has always produced 0xFFFFFFFF.
	// Use ATOU() to support unsigned (i.e. UINT, LPARAM, and WPARAM are all 32-bit unsigned values).
	// ATOU() also supports hex strings in the script, such as 0xFF, which is why it's commonly
	// used in functions such as this.  v1.0.40.05: Support the passing of a literal (quoted) string
	// by checking whether the original/raw arg's first character is '"'.  The avoids the need to
	// put the string into a variable and then pass something like &MyVar.
	WPARAM wparam = (mArgc > 1 && mArg[1].text[0] == '"') ? (WPARAM)sArgDeref[1] : ATOU(sArgDeref[1]);
	LPARAM lparam = (mArgc > 2 && mArg[2].text[0] == '"') ? (LPARAM)sArgDeref[2] : ATOU(sArgDeref[2]);
	if (aUseSend)
	{
		DWORD dwResult;
		// Timeout increased from 2000 to 5000 in v1.0.27:
		if (!SendMessageTimeout(control_window, msg, wparam, lparam, SMTO_ABORTIFHUNG, 5000, &dwResult))
			return g_ErrorLevel->Assign("FAIL"); // Need a special value to distinguish this from numeric reply-values.
		return g_ErrorLevel->Assign(dwResult); // UINT seems best most of the time?
	}
	else // Post vs. Send
		return g_ErrorLevel->Assign(PostMessage(control_window, msg, wparam, lparam)
			? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);
	// By design (since this is a power user feature), no ControlDelay is done here.
}



ResultType Line::ScriptProcess(char *aCmd, char *aProcess, char *aParam3)
{
	ProcessCmds process_cmd = ConvertProcessCmd(aCmd);
	// Runtime error is rare since it is caught at load-time unless it's in a var. ref.
	if (process_cmd == PROCESS_CMD_INVALID)
		return LineError(ERR_PARAM1_INVALID ERR_ABORT, FAIL, aCmd);

	HANDLE hProcess;
	DWORD pid, priority;
	BOOL result;

	switch (process_cmd)
	{
	case PROCESS_CMD_EXIST:
		return g_ErrorLevel->Assign(*aProcess ? ProcessExist(aProcess) : GetCurrentProcessId()); // The discovered PID or zero if none.

	case PROCESS_CMD_CLOSE:
		if (pid = ProcessExist(aProcess))  // Assign
		{
			if (hProcess = OpenProcess(PROCESS_ALL_ACCESS, FALSE, pid))
			{
				result = TerminateProcess(hProcess, 0);
				CloseHandle(hProcess);
				return g_ErrorLevel->Assign(result ? pid : 0); // Indicate success or failure.
			}
		}
		// Since above didn't return, yield a PID of 0 to indicate failure.
		return g_ErrorLevel->Assign("0");

	case PROCESS_CMD_PRIORITY:
		switch (toupper(*aParam3))
		{
		case 'L': priority = IDLE_PRIORITY_CLASS; break;
		case 'B': priority = BELOW_NORMAL_PRIORITY_CLASS; break;
		case 'N': priority = NORMAL_PRIORITY_CLASS; break;
		case 'A': priority = ABOVE_NORMAL_PRIORITY_CLASS; break;
		case 'H': priority = HIGH_PRIORITY_CLASS; break;
		case 'R': priority = REALTIME_PRIORITY_CLASS; break;
		default:
			return g_ErrorLevel->Assign("0");  // 0 indicates failure in this case (i.e. a PID of zero).
		}
		if (pid = *aProcess ? ProcessExist(aProcess) : GetCurrentProcessId())  // Assign
		{
			if (hProcess = OpenProcess(PROCESS_SET_INFORMATION, FALSE, pid)) // Assign
			{
				// If OS doesn't support "above/below normal", seems best to default to normal rather than high/low,
				// since "above/below normal" aren't that dramatically different from normal:
				if (!g_os.IsWin2000orLater() && (priority == BELOW_NORMAL_PRIORITY_CLASS || priority == ABOVE_NORMAL_PRIORITY_CLASS))
					priority = NORMAL_PRIORITY_CLASS;
				result = SetPriorityClass(hProcess, priority);
				CloseHandle(hProcess);
				return g_ErrorLevel->Assign(result ? pid : 0); // Indicate success or failure.
			}
		}
		// Otherwise, return a PID of 0 to indicate failure.
		return g_ErrorLevel->Assign("0");

	case PROCESS_CMD_WAIT:
	case PROCESS_CMD_WAITCLOSE:
	{
		// This section is similar to that used for WINWAIT and RUNWAIT:
		bool wait_indefinitely;
		int sleep_duration;
		DWORD start_time;
		if (*aParam3) // the param containing the timeout value isn't blank.
		{
			wait_indefinitely = false;
			sleep_duration = (int)(ATOF(aParam3) * 1000); // Can be zero.
			start_time = GetTickCount();
		}
		else
		{
			wait_indefinitely = true;
			sleep_duration = 0; // Just to catch any bugs.
		}
		for (;;)
		{ // Always do the first iteration so that at least one check is done.
			pid = ProcessExist(aProcess);
			if (process_cmd == PROCESS_CMD_WAIT)
			{
				if (pid)
					return g_ErrorLevel->Assign(pid);
			}
			else // PROCESS_CMD_WAITCLOSE
			{
				// Since PID cannot always be determined (i.e. if process never existed, there was
				// no need to wait for it to close), for consistency, return 0 on success.
				if (!pid)
					return g_ErrorLevel->Assign("0");
			}
			// Must cast to int or any negative result will be lost due to DWORD type:
			if (wait_indefinitely || (int)(sleep_duration - (GetTickCount() - start_time)) > SLEEP_INTERVAL_HALF)
				MsgSleep(100);  // For performance reasons, don't check as often as the WinWait family does.
			else // Done waiting.
				return g_ErrorLevel->Assign(pid);
				// Above assigns 0 if "Process Wait" times out; or the PID of the process that still exists
				// if "Process WaitClose" times out.
		} // for()
	} // case
	} // switch()

	return FAIL;  // Should never be executed; just here to catch bugs.
}



ResultType WinSetRegion(HWND aWnd, char *aPoints)
// Caller has initialized g_ErrorLevel to ERRORLEVEL_ERROR for us.
{
	if (!*aPoints) // Attempt to restore the window's normal/correct region.
	{
		// Fix for v1.0.31.07: The old method used the following, but apparently it's not the correct
		// way to restore a window's proper/normal region because when such a window is later maximized,
		// it retains its incorrect/smaller region:
		//if (GetWindowRect(aWnd, &rect))
		//{
		//	// Adjust the rect to keep the same size but have its upper-left corner at 0,0:
		//	rect.right -= rect.left;
		//	rect.bottom -= rect.top;
		//	rect.left = 0;
		//	rect.top = 0;
		//	if (hrgn = CreateRectRgnIndirect(&rect)) // Assign
		//	{
		//		// Presumably, the system deletes the former region when upon a successful call to SetWindowRgn().
		//		if (SetWindowRgn(aWnd, hrgn, TRUE))
		//			return g_ErrorLevel->Assign(ERRORLEVEL_NONE);
		//		// Otherwise, get rid of it since it didn't take effect:
		//		DeleteObject(hrgn);
		//	}
		//}
		//// Since above didn't return:
		//return OK; // Let ErrorLevel tell the story.

		// It's undocumented by MSDN, but apparently setting the Window's region to NULL restores it
		// to proper working order:
		if (SetWindowRgn(aWnd, NULL, TRUE))
			return g_ErrorLevel->Assign(ERRORLEVEL_NONE);
		return OK; // Let ErrorLevel tell the story.
	}

	#define MAX_REGION_POINTS 2000  // 2000 requires 16 KB of stack space.
	POINT pt[MAX_REGION_POINTS];
	int pt_count;
	char *cp;

	// Set defaults prior to parsing options in case any options are absent:
	int width = COORD_UNSPECIFIED;
	int height = COORD_UNSPECIFIED;
	int rr_width = COORD_UNSPECIFIED; // These two are for the rounded-rectangle method.
	int rr_height = COORD_UNSPECIFIED;
	bool use_ellipse = false;

	int fill_mode = ALTERNATE;
	// Concerning polygon regions: ALTERNATE is used by default (somewhat arbitrarily, but it seems to be the
	// more typical default).
	// MSDN: "In general, the modes [ALTERNATE vs. WINDING] differ only in cases where a complex,
	// overlapping polygon must be filled (for example, a five-sided polygon that forms a five-pointed
	// star with a pentagon in the center). In such cases, ALTERNATE mode fills every other enclosed
	// region within the polygon (that is, the points of the star), but WINDING mode fills all regions
	// (that is, the points and the pentagon)."

	for (pt_count = 0, cp = aPoints; *(cp = omit_leading_whitespace(cp));)
	{
		// To allow the MAX to be increased in the future with less chance of breaking existing scripts, consider this an error.
		if (pt_count >= MAX_REGION_POINTS)
			return OK; // Let ErrorLevel tell the story.

		if (isdigit(*cp) || *cp == '-' || *cp == '+') // v1.0.38.02: Recognize leading minus/plus sign so that the X-coord is just as tolerant as the Y.
		{
			// Assume it's a pair of X/Y coordinates.  It's done this way rather than using X and Y
			// as option letters because:
			// 1) The script is more readable when there are multiple coordinates (for polygon).
			// 2) It enforces the fact that each X must have a Y and that X must always come before Y
			//    (which simplifies and reduces the size of the code).
			pt[pt_count].x = ATOI(cp);
			// For the delimiter, dash is more readable than pipe, even though it overlaps with "minus sign".
			// "x" is not used to avoid detecting "x" inside hex numbers.
			#define REGION_DELIMITER '-'
			if (   !(cp = strchr(cp + 1, REGION_DELIMITER))   ) // v1.0.38.02: cp + 1 to omit any leading minus sign.
				return OK; // Let ErrorLevel tell the story.
			pt[pt_count].y = ATOI(++cp);  // Increment cp by only 1 to support negative Y-coord.
			++pt_count; // Move on to the next element of the pt array.
		}
		else
		{
			++cp;
			switch(toupper(cp[-1]))
			{
			case 'E':
				use_ellipse = true;
				break;
			case 'R':
				if (!*cp || *cp == ' ') // Use 30x30 default.
				{
					rr_width = 30;
					rr_height = 30;
				}
				else
				{
					rr_width = ATOI(cp);
					if (cp = strchr(cp, REGION_DELIMITER)) // Assign
						rr_height = ATOI(++cp);
					else // Avoid problems with going beyond the end of the string.
						return OK; // Let ErrorLevel tell the story.
				}
				break;
			case 'W':
				if (!strnicmp(cp, "ind", 3)) // [W]ind.
					fill_mode = WINDING;
				else
					width = ATOI(cp);
				break;
			case 'H':
				height = ATOI(cp);
				break;
			default: // For simplicity and to reserve other letters for future use, unknown options result in failure.
				return OK; // Let ErrorLevel tell the story.
			} // switch()
		} // else

		if (   !(cp = strchr(cp, ' '))   ) // No more items.
			break;
	}

	if (!pt_count)
		return OK; // Let ErrorLevel tell the story.

	bool width_and_height_were_both_specified = !(width == COORD_UNSPECIFIED || height == COORD_UNSPECIFIED);
	if (width_and_height_were_both_specified)
	{
		width += pt[0].x;   // Make width become the right side of the rect.
		height += pt[0].y;  // Make height become the bottom.
	}

	HRGN hrgn;
	if (use_ellipse) // Ellipse.
	{
		if (!width_and_height_were_both_specified || !(hrgn = CreateEllipticRgn(pt[0].x, pt[0].y, width, height)))
			return OK; // Let ErrorLevel tell the story.
	}
	else if (rr_width != COORD_UNSPECIFIED) // Rounded rectangle.
	{
		if (!width_and_height_were_both_specified || !(hrgn = CreateRoundRectRgn(pt[0].x, pt[0].y, width, height, rr_width, rr_height)))
			return OK; // Let ErrorLevel tell the story.
	}
	else if (width_and_height_were_both_specified) // Rectangle.
	{
		if (!(hrgn = CreateRectRgn(pt[0].x, pt[0].y, width, height)))
			return OK; // Let ErrorLevel tell the story.
	}
	else // Polygon
	{
		if (   !(hrgn = CreatePolygonRgn(pt, pt_count, fill_mode))   )
			return OK;
	}

	// Since above didn't return, hrgn is now a non-NULL region ready to be assigned to the window.

	// Presumably, the system deletes the window's former region upon a successful call to SetWindowRgn():
	if (!SetWindowRgn(aWnd, hrgn, TRUE))
	{
		DeleteObject(hrgn);
		return OK; // Let ErrorLevel tell the story.
	}
	//else don't delete hrgn since the system has taken ownership of it.

	// Since above didn't return, indicate success.
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE);
}


						
ResultType Line::WinSet(char *aAttrib, char *aValue, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
{
	WinSetAttributes attrib = ConvertWinSetAttribute(aAttrib);
	if (attrib == WINSET_INVALID)
		return LineError(ERR_PARAM1_INVALID, FAIL, aAttrib);

	// Set default ErrorLevel for any commands that change ErrorLevel.
	// The default should be ERRORLEVEL_ERROR so that that value will be returned
	// by default when the target window doesn't exist:
	if (attrib == WINSET_STYLE || attrib == WINSET_EXSTYLE || attrib == WINSET_REGION)
		g_ErrorLevel->Assign(ERRORLEVEL_ERROR);

	// Since this is a macro, avoid repeating it for every case of the switch():
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		return OK; // Let ErrorLevel (for commands that set it above) tell the story.

	int value;
	DWORD exstyle;

	switch (attrib)
	{
	case WINSET_ALWAYSONTOP:
	{
		if (   !(exstyle = GetWindowLong(target_window, GWL_EXSTYLE))   )
			return OK;
		HWND topmost_or_not;
		switch(ConvertOnOffToggle(aValue))
		{
		case TOGGLED_ON: topmost_or_not = HWND_TOPMOST; break;
		case TOGGLED_OFF: topmost_or_not = HWND_NOTOPMOST; break;
		case NEUTRAL: // parameter was blank, so it defaults to TOGGLE.
		case TOGGLE: topmost_or_not = (exstyle & WS_EX_TOPMOST) ? HWND_NOTOPMOST : HWND_TOPMOST; break;
		default: return OK;
		}
		// SetWindowLong() didn't seem to work, at least not on some windows.  But this does.
		// As of v1.0.25.14, SWP_NOACTIVATE is also specified, though its absence does not actually
		// seem to activate the window, at least on XP (perhaps due to anti-focus-stealing measure
		// in Win98/2000 and beyond).  Or perhaps its something to do with the presence of
		// topmost_or_not (HWND_TOPMOST/HWND_NOTOPMOST), which might always avoid activating the
		// window.
		SetWindowPos(target_window, topmost_or_not, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE);
		break;
	}

	// Note that WINSET_TOP is not offered as an option since testing reveals it has no effect on
	// top level (parent) windows, perhaps due to the anti focus-stealing measures in the OS.
	case WINSET_BOTTOM:
		// Note: SWP_NOACTIVATE must be specified otherwise the target window often/always fails to go
		// to the bottom:
		SetWindowPos(target_window, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE);
		break;
	case WINSET_TOP:
		// Note: SWP_NOACTIVATE must be specified otherwise the target window often/always fails to go
		// to the bottom:
		SetWindowPos(target_window, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE);
		break;

	case WINSET_TRANSPARENT:
	case WINSET_TRANSCOLOR:
	{
		// IMPORTANT (when considering future enhancements to these commands): Unlike
		// SetLayeredWindowAttributes(), which works on Windows 2000, GetLayeredWindowAttributes()
		// is supported only on XP or later.

		// It appears that turning on WS_EX_LAYERED in an attempt to retrieve the window's
		// former transparency setting does not work.  The OS probably does not store the
		// former transparency level (i.e. it forgets it the moment the WS_EX_LAYERED exstyle
		// is turned off).  This is true even if the following are done after the SetWindowLong():
		//MySetLayeredWindowAttributes(target_window, 0, 0, 0)
		// or:
		//if (MyGetLayeredWindowAttributes(target_window, &color, &alpha, &flags))
		//	MySetLayeredWindowAttributes(target_window, color, alpha, flags);
		// The above is why there is currently no "on" or "toggle" sub-command, just "Off".

		// Must fetch the below at runtime, otherwise the program can't even be launched on Win9x/NT.
		// Also, since the color of an HBRUSH can't be easily determined (since it can be a pattern and
		// since there seem to be no easy API calls to discover the colors of pixels in an HBRUSH),
		// the following is not yet implemented: Use window's own class background color (via
		// GetClassLong) if aValue is entirely blank.
		typedef BOOL (WINAPI *MySetLayeredWindowAttributesType)(HWND, COLORREF, BYTE, DWORD);
		static MySetLayeredWindowAttributesType MySetLayeredWindowAttributes = (MySetLayeredWindowAttributesType)
			GetProcAddress(GetModuleHandle("user32"), "SetLayeredWindowAttributes");
		if (!MySetLayeredWindowAttributes || !(exstyle = GetWindowLong(target_window, GWL_EXSTYLE)))
			return OK;  // Do nothing on OSes that don't support it.
		if (!stricmp(aValue, "Off"))
			// One user reported that turning off the attribute helps window's scrolling performance.
			SetWindowLong(target_window, GWL_EXSTYLE, exstyle & ~WS_EX_LAYERED);
		else
		{
			if (attrib == WINSET_TRANSPARENT)
			{
				// Update to the below for v1.0.23: WS_EX_LAYERED can now be removed via the above:
				// NOTE: It seems best never to remove the WS_EX_LAYERED attribute, even if the value is 255
				// (non-transparent), since the window might have had that attribute previously and may need
				// it to function properly.  For example, an app may support making its own windows transparent
				// but might not expect to have to turn WS_EX_LAYERED back on if we turned it off.  One drawback
				// of this is a quote from somewhere that might or might not be accurate: "To make this window
				// completely opaque again, remove the WS_EX_LAYERED bit by calling SetWindowLong and then ask
				// the window to repaint. Removing the bit is desired to let the system know that it can free up
				// some memory associated with layering and redirection."
				value = ATOI(aValue);
				// A little debatable, but this behavior seems best, at least in some cases:
				if (value < 0)
					value = 0;
				else if (value > 255)
					value = 255;
				SetWindowLong(target_window, GWL_EXSTYLE, exstyle | WS_EX_LAYERED);
				MySetLayeredWindowAttributes(target_window, 0, value, LWA_ALPHA);
			}
			else // attrib == WINSET_TRANSCOLOR
			{
				// The reason WINSET_TRANSCOLOR accepts both the color and an optional transparency settings
				// is that calling SetLayeredWindowAttributes() with only the LWA_COLORKEY flag causes the
				// window to lose its current transparency setting in favor of the transparent color.  This
				// is true even though the LWA_ALPHA flag was not specified, which seems odd and is a little
				// disappointing, but that's the way it is on XP at least.
				char aValue_copy[256];
				strlcpy(aValue_copy, aValue, sizeof(aValue_copy)); // Make a modifiable copy.
				char *space_pos = StrChrAny(aValue_copy, " \t"); // Space or tab.
				if (space_pos)
				{
					*space_pos = '\0';
					++space_pos;  // Point it to the second substring.
				}
				COLORREF color = ColorNameToBGR(aValue_copy);
				if (color == CLR_NONE) // A matching color name was not found, so assume it's in hex format.
					// It seems strtol() automatically handles the optional leading "0x" if present:
					color = rgb_to_bgr(strtol(aValue_copy, NULL, 16));
				DWORD flags;
				if (   space_pos && *(space_pos = omit_leading_whitespace(space_pos))   ) // Relies on short-circuit boolean.
				{
					value = ATOI(space_pos);  // To keep it simple, don't bother with 0 to 255 range validation in this case.
					flags = LWA_COLORKEY|LWA_ALPHA;  // i.e. set both the trans-color and the transparency level.
				}
				else // No translucency value is present, only a trans-color.
				{
					value = 0;  // Init to avoid possible compiler warning.
					flags = LWA_COLORKEY;
				}
				SetWindowLong(target_window, GWL_EXSTYLE, exstyle | WS_EX_LAYERED);
				MySetLayeredWindowAttributes(target_window, color, value, flags);
			}
		}
		break;
	}

	case WINSET_STYLE:
	case WINSET_EXSTYLE:
	{
		if (!*aValue)
			return OK; // Seems best not to treat an explicit blank as zero. Let ErrorLevel tell the story.
		int style_index = (attrib == WINSET_STYLE) ? GWL_STYLE : GWL_EXSTYLE;
		DWORD new_style, orig_style = GetWindowLong(target_window, style_index);
		if (!strchr("+-^", *aValue))  // | and & are used instead of +/- to allow +/- to have their native function.
			new_style = ATOU(aValue); // No prefix, so this new style will entirely replace the current style.
		else
		{
			++aValue; // Won't work combined with next line, due to next line being a macro that uses the arg twice.
			DWORD style_change = ATOU(aValue);
			// +/-/^ are used instead of |&^ because the latter is confusing, namely that
			// "&" really means &=~style, etc.
			switch(aValue[-1])
			{
			case '+': new_style = orig_style | style_change; break;
			case '-': new_style = orig_style & ~style_change; break;
			case '^': new_style = orig_style ^ style_change; break;
			}
		}
		SetLastError(0); // Prior to SetWindowLong(), as recommended by MSDN.
		if (SetWindowLong(target_window, style_index, new_style) || !GetLastError()) // This is the precise way to detect success according to MSDN.
		{
			// Even if it indicated success, sometimes it failed anyway.  Find out for sure:
			if (GetWindowLong(target_window, style_index) != orig_style) // Even a partial change counts as a success.
			{
				// SetWindowPos is also necessary, otherwise the frame thickness entirely around the window
				// does not get updated (just parts of it):
				SetWindowPos(target_window, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
				// Since SetWindowPos() probably doesn't know that the style changed, below is probably necessary
				// too, at least in some cases:
				InvalidateRect(target_window, NULL, TRUE); // Quite a few styles require this to become visibly manifest.
				return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
			}
		}
		return OK; // Since above didn't return, it's a failure.  Let ErrorLevel tell the story.
	}

	case WINSET_ENABLE:
	case WINSET_DISABLE: // These are separate sub-commands from WINSET_STYLE because merely changing the WS_DISABLED style is usually not as effective as calling EnableWindow().
		EnableWindow(target_window, attrib == WINSET_ENABLE);
		return OK;

	case WINSET_REGION:
		return WinSetRegion(target_window, aValue);

	case WINSET_REDRAW:
		// Seems best to always have the last param be TRUE, for now, so that aValue can be
		// reserved for future use such as invalidating only part of a window, etc. Also, it
		// seems best not to call UpdateWindow(), which forces the window to immediately
		// process a WM_PAINT message, since that might not be desirable as a default (maybe
		// an option someday).  Other future options might include alternate methods of
		// getting a window to redraw, such as:
		// SendMessage(mHwnd, WM_NCPAINT, 1, 0);
		// RedrawWindow(mHwnd, NULL, NULL, RDW_INVALIDATE|RDW_FRAME|RDW_UPDATENOW);
		// SetWindowPos(mHwnd, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
		// GetClientRect(mControl[mDefaultButtonIndex].hwnd, &client_rect);
		// InvalidateRect(mControl[mDefaultButtonIndex].hwnd, &client_rect, TRUE);
		InvalidateRect(target_window, NULL, TRUE);
		break;

	} // switch()
	return OK;
}



ResultType Line::WinSetTitle(char *aTitle, char *aText, char *aNewTitle, char *aExcludeTitle, char *aExcludeText)
// Like AutoIt2, this function and others like it always return OK, even if the target window doesn't
// exist or there action doesn't actually succeed.
{
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
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
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	// Even if target_window is NULL, we want to continue on so that the output
	// param is set to be the empty string, which is the proper thing to do
	// rather than leaving whatever was in there before.

	// Handle the output parameter.  See the comments in ACT_CONTROLGETTEXT for details.
	VarSizeType space_needed = target_window ? GetWindowTextLength(target_window) + 1 : 1; // 1 for terminator.
	if (output_var->Assign(NULL, space_needed - 1) != OK)
		return FAIL;  // It already displayed the error.
	if (target_window)
	{
		// Update length using the actual length, rather than the estimate provided by GetWindowTextLength():
		output_var->Length() = (VarSizeType)GetWindowText(target_window, output_var->Contents(), space_needed);
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



ResultType Line::WinGetClass(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	if (!target_window)
		return output_var->Assign();
	char class_name[WINDOW_CLASS_SIZE];
	if (!GetClassName(target_window, class_name, sizeof(class_name)))
		return output_var->Assign();
	return output_var->Assign(class_name);
}



ResultType WinGetList(Var *output_var, WinGetCmds aCmd, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
// Helper function for WinGet() to avoid having a WindowSearch object on its stack (since that object
// normally isn't needed).  Caller has ensured that output_var isn't NULL.
{
	WindowSearch ws;
	ws.mFindLastMatch = true; // Must set mFindLastMatch to get all matches rather than just the first.
	ws.mArrayStart = (aCmd == WINGET_CMD_LIST) ? output_var : NULL; // Provide the position in the var list of where the array-element vars will be.
	// If aTitle is ahk_id nnnn, the Enum() below will be inefficient.  However, ahk_id is almost unheard of
	// in this context because it makes little sense, so no extra code is added to make that case efficient.
	if (ws.SetCriteria(g, aTitle, aText, aExcludeTitle, aExcludeText)) // These criteria allow the possibilty of a match.
		EnumWindows(EnumParentFind, (LPARAM)&ws);
	//else leave ws.mFoundCount set to zero (by the constructor).
	return output_var->Assign(ws.mFoundCount);
}



ResultType Line::WinGet(char *aCmd, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	Var *output_var = ResolveVarOfArg(0);  // This is done even for WINGET_CMD_LIST.
	if (!output_var)
		return FAIL;

	WinGetCmds cmd = ConvertWinGetCmd(aCmd);
	// Since command names are validated at load-time, this only happens if the command name
	// was contained in a variable reference.  But for simplicity of design here, return
	// failure in this case (unlike other functions similar to this one):
	if (cmd == WINGET_CMD_INVALID)
		return LineError(ERR_PARAM2_INVALID ERR_ABORT, FAIL, aCmd);

	bool target_window_determined = true;  // Set default.
	HWND target_window;
	IF_USE_FOREGROUND_WINDOW(g.DetectHiddenWindows, aTitle, aText, aExcludeTitle, aExcludeText)
	else if (!(*aTitle || *aText || *aExcludeTitle || *aExcludeText)
		&& !(cmd == WINGET_CMD_LIST || cmd == WINGET_CMD_COUNT)) // v1.0.30.02/v1.0.30.03: Have "list"/"count" get all windows on the system when there are no parameters.
		target_window = GetValidLastUsedWindow(g);
	else
		target_window_determined = false;  // A different method is required.

	// Used with WINGET_CMD_LIST to create an array (if needed).  Make it longer than Max var name
	// so that FindOrAddVar() will be able to spot and report var names that are too long:
	char var_name[MAX_VAR_NAME_LENGTH + 20], buf[32];
	Var *array_item;

	switch(cmd)
	{
	case WINGET_CMD_ID:
	case WINGET_CMD_IDLAST:
		if (!target_window_determined)
			target_window = WinExist(g, aTitle, aText, aExcludeTitle, aExcludeText, cmd == WINGET_CMD_IDLAST);
		if (target_window)
			return output_var->AssignHWND(target_window);
		else
			return output_var->Assign();

	case WINGET_CMD_PID:
	case WINGET_CMD_PROCESSNAME:
		if (!target_window_determined)
			target_window = WinExist(g, aTitle, aText, aExcludeTitle, aExcludeText);
		if (target_window)
		{
			DWORD pid;
			GetWindowThreadProcessId(target_window, &pid);
			if (cmd == WINGET_CMD_PID)
				return output_var->Assign(pid);
			// Otherwise, get the full path and name of the executable that owns this window.
			_ultoa(pid, buf, 10);
			char process_name[MAX_PATH];
			if (ProcessExist(buf, process_name))
				return output_var->Assign(process_name);
		}
		// If above didn't return:
		return output_var->Assign();

	case WINGET_CMD_COUNT:
	case WINGET_CMD_LIST:
		// LIST retrieves a list of HWNDs for the windows that match the given criteria and stores them in
		// an array.  The number of items in the array is stored in the base array name (unlike
		// StringSplit, which stores them in array element #0).  This is done for performance reasons
		// (so that element #0 doesn't have to be looked up at runtime), but mostly because of the
		// complexity of resolving a parameter than can be either an output-var or an array name at
		// load-time -- namely that if param #1 were allowed to be an array name, there is ambiguity
		// about where the name of the array is actually stored depending on whether param#1 was literal
		// text or a deref.  So it's easier and performs better just to do it this way, even though it
		// breaks from the StringSplit tradition:
		if (target_window_determined)
		{
			if (!target_window)
				return output_var->Assign("0"); // 0 windows found
			if (cmd == WINGET_CMD_LIST)
			{
				// Otherwise, since the target window has been determined, we know that it is
				// the only window to be put into the array:
				snprintf(var_name, sizeof(var_name), "%s1", output_var->mName);
				if (   !(array_item = g_script.FindOrAddVar(var_name, 0, output_var->IsLocal()
					? ALWAYS_USE_LOCAL : ALWAYS_USE_GLOBAL))   )  // Find or create element #1.
					return FAIL;  // It will have already displayed the error.
				if (!array_item->AssignHWND(target_window))
					return FAIL;
			}
			return output_var->Assign("1");  // 1 window found
		}
		// Otherwise, the target window(s) have not yet been determined and a special method
		// is required to gather them.
		return WinGetList(output_var, cmd, aTitle, aText, aExcludeTitle, aExcludeText); // Outsourced to avoid having a WindowSearch object on this function's stack.

	case WINGET_CMD_MINMAX:
		if (!target_window_determined)
			target_window = WinExist(g, aTitle, aText, aExcludeTitle, aExcludeText);
		// Testing shows that it's not possible for a minimized window to also be maximized (under
		// the theory that upon restoration, it *would* be maximized.  This is unfortunate if there
		// is no other way to determine what the restoration size and maxmized state will be for a
		// minimized window.
		if (target_window)
			return output_var->Assign(IsZoomed(target_window) ? 1 : (IsIconic(target_window) ? -1 : 0));
		else
			return output_var->Assign();

	case WINGET_CMD_CONTROLLIST:
		if (!target_window_determined)
			target_window = WinExist(g, aTitle, aText, aExcludeTitle, aExcludeText);
		return target_window ? WinGetControlList(output_var, target_window) : output_var->Assign();

	case WINGET_CMD_STYLE:
	case WINGET_CMD_EXSTYLE:
		if (!target_window_determined)
			target_window = WinExist(g, aTitle, aText, aExcludeTitle, aExcludeText);
		if (!target_window)
			return output_var->Assign();
		sprintf(buf, "0x%08X", GetWindowLong(target_window, cmd == WINGET_CMD_STYLE ? GWL_STYLE : GWL_EXSTYLE));
		return output_var->Assign(buf);

	case WINGET_CMD_TRANSPARENT:
	case WINGET_CMD_TRANSCOLOR:
		if (!target_window_determined)
			target_window = WinExist(g, aTitle, aText, aExcludeTitle, aExcludeText);
		if (!target_window)
			return output_var->Assign();
		typedef BOOL (WINAPI *MyGetLayeredWindowAttributesType)(HWND, COLORREF*, BYTE*, DWORD*);
		static MyGetLayeredWindowAttributesType MyGetLayeredWindowAttributes = (MyGetLayeredWindowAttributesType)
			GetProcAddress(GetModuleHandle("user32"), "GetLayeredWindowAttributes");
		COLORREF color;
		BYTE alpha;
		DWORD flags;
		// IMPORTANT (when considering future enhancements to these commands): Unlike
		// SetLayeredWindowAttributes(), which works on Windows 2000, GetLayeredWindowAttributes()
		// is supported only on XP or later.
		if (!MyGetLayeredWindowAttributes || !(MyGetLayeredWindowAttributes(target_window, &color, &alpha, &flags)))
			return output_var->Assign();
		if (cmd == WINGET_CMD_TRANSPARENT)
			return (flags & LWA_ALPHA) ? output_var->Assign((DWORD)alpha) : output_var->Assign();
		else // WINGET_CMD_TRANSCOLOR
		{
			if (flags & LWA_COLORKEY)
			{
				// Store in hex format to aid in debugging scripts.  Also, the color is always
				// stored in RGB format, since that's what WinSet uses:
				sprintf(buf, "0x%06X", bgr_to_rgb(color));
				return output_var->Assign(buf);
			}
			else // This window does not have a transparent color (or it's not accessible to us, perhaps for reasons described at MSDN GetLayeredWindowAttributes()).
				return output_var->Assign();
		}
	}

	return FAIL;  // Never executed (increases maintainability and avoids compiler warning).
}



ResultType Line::WinGetControlList(Var *aOutputVar, HWND aTargetWindow)
// Caller must ensure that aOutputVar and aTargetWindow are non-NULL and valid.
// Every control is fetched rather than just a list of distinct class names (possibly with a
// second script array containing the quantity of each class) because it's conceivable that the
// z-order of the controls will be useful information to some script authors.
// A delimited list is used rather than the array technique used by "WinGet, OutputVar, List" because:
// 1) It allows the flexibility of searching the list more easily with something like IfInString.
// 2) It seems rather rare that the count of items in the list would be useful info to a script author
//    (the count can be derived with a parsing loop if it's ever needed).
// 3) It saves memory since script arrays are permanent and since each array element would incur
//    the overhead of being a script variable, not to mention that each variable has a minimum
//    capacity (additional overhead) of 64 bytes.
{
	control_list_type cl; // A big struct containing room to store class names and counts for each.
	CL_INIT_CONTROL_LIST(cl)
	cl.target_buf = NULL;  // First pass: Signal it not not to write to the buf, but instead only calculate the length.
	EnumChildWindows(aTargetWindow, EnumChildGetControlList, (LPARAM)&cl);
	if (!cl.total_length) // No controls in the window.
		return aOutputVar->Assign();
	// This adjustment was added because someone reported that max variable capacity was being
	// exceeded in some cases (perhaps custom controls that retrieve large amounts of text
	// from the disk in response to the "get text" message):
	if (cl.total_length >= g_MaxVarCapacity) // Allow the command to succeed by truncating the text.
		cl.total_length = g_MaxVarCapacity - 1;
	// Set up the var, enlarging it if necessary.  If the aOutputVar is of type VAR_CLIPBOARD,
	// this call will set up the clipboard for writing:
	if (aOutputVar->Assign(NULL, (VarSizeType)cl.total_length) != OK)
		return FAIL;  // It already displayed the error.
	// Fetch the text directly into the var.  Also set the length explicitly
	// in case actual size written was off from the esimated size (in case the list of
	// controls changed in the split second between pass #1 and pass #2):
	CL_INIT_CONTROL_LIST(cl)
	cl.target_buf = aOutputVar->Contents();  // Second pass: Write to the buffer.
	cl.capacity = aOutputVar->Capacity(); // Because granted capacity might be a little larger than we asked for.
	EnumChildWindows(aTargetWindow, EnumChildGetControlList, (LPARAM)&cl);
	aOutputVar->Length() = (VarSizeType)cl.total_length;  // In case it wound up being smaller than expected.
	if (!cl.total_length) // Something went wrong, so make sure its terminated just in case.
		*aOutputVar->Contents() = '\0';  // Safe because Assign() gave us a non-constant memory area.
	return aOutputVar->Close();  // In case it's the clipboard.
}



BOOL CALLBACK EnumChildGetControlList(HWND aWnd, LPARAM lParam)
{
	// Note: IsWindowVisible(aWnd) is not checked because although Window Spy does not reveal
	// hidden controls if the mouse happens to be hovering over one, it does include them in its
	// sequence numbering (which is a relieve, since results are probably much more consistent
	// then, esp. for apps that hide and unhide controls in response to actions on other controls).
	char class_name[WINDOW_CLASS_SIZE + 5];  // +5 to allow room for the sequence number to be appended later below.
	int class_name_length = GetClassName(aWnd, class_name, WINDOW_CLASS_SIZE); // Don't include the +5 in this.
	if (!class_name_length)
		return TRUE; // Probably very rare. Continue enumeration since Window Spy doesn't even check for failure.
	// It has been verified that GetClassName()'s returned length does not count the terminator.

	control_list_type &cl = *((control_list_type *)lParam);  // For performance and convenience.

	// Check if this class already exists in the class array:
	int class_index;
	for (class_index = 0; class_index < cl.total_classes; ++class_index)
		if (!stricmp(cl.class_name[class_index], class_name))
			break;
	if (class_index < cl.total_classes) // Match found.
	{
		++cl.class_count[class_index]; // Increment the number of controls of this class that have been found so far.
		if (cl.class_count[class_index] > 99999) // Sanity check; prevents buffer overflow or number truncation in class_name.
			return TRUE;  // Continue the enumeration.
	}
	else // No match found, so create new entry if there's room.
	{
		if (cl.total_classes == CL_MAX_CLASSES // No pointers left.
			|| CL_CLASS_BUF_SIZE - (cl.buf_free_spot - cl.class_buf) - 1 < class_name_length) // Insuff. room in buf.
			return TRUE; // Very rare. Continue the enumeration so that class names already found can be collected.
		// Otherwise:
		cl.class_name[class_index] = cl.buf_free_spot;  // Set this pointer to its place in the buffer.
		strcpy(cl.class_name[class_index], class_name); // Copy the string into this place.
		cl.buf_free_spot += class_name_length + 1;  // +1 because every string in the buf needs its own terminator.
		cl.class_count[class_index] = 1;  // Indicate that the quantity of this class so far is 1.
		++cl.total_classes;
	}

	_itoa(cl.class_count[class_index], class_name + class_name_length, 10); // Append the seq. number to class_name.
	class_name_length = (int)strlen(class_name);  // Update the length.
	int extra_length;
	if (cl.is_first_iteration)
	{
		extra_length = 0; // All items except the first are preceded by a delimiting LF.
		cl.is_first_iteration = false;
	}
	else
		extra_length = 1;
	if (cl.target_buf)
	{
		if ((int)(cl.capacity - cl.total_length - extra_length - 1) < class_name_length)
			// No room in target_buf (i.e. don't write a partial item to the buffer).
			return TRUE;  // Rare: it should only happen if size in pass #2 differed from that calc'd in pass #1.
		if (extra_length)
		{
			cl.target_buf[cl.total_length] = '\n'; // Replace previous item's terminator with newline.
			cl.total_length += extra_length;
		}
		strcpy(cl.target_buf + cl.total_length, class_name); // Write class name + seq. number.
		cl.total_length += class_name_length;
	}
	else // Caller only wanted the total length calculated.
		cl.total_length += class_name_length + extra_length;

	return TRUE; // Continue enumeration through all the windows.
}



ResultType Line::WinGetText(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
	// Even if target_window is NULL, we want to continue on so that the output
	// variables are set to be the empty string, which is the proper thing to do
	// rather than leaving whatever was in there before:
	if (!target_window)
		return output_var->Assign(); // Tell it not to free the memory by not calling with "".

	length_and_buf_type sab;
	sab.buf = NULL; // Tell it just to calculate the length this time around.
	sab.total_length = 0; // Init
	sab.capacity = 0;     //
	EnumChildWindows(target_window, EnumChildGetText, (LPARAM)&sab);

	if (!sab.total_length) // No text in window.
	{
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		return output_var->Assign(); // Tell it not to free the memory by omitting all params.
	}
	// This adjustment was added because someone reported that max variable capacity was being
	// exceeded in some cases (perhaps custom controls that retrieve large amounts of text
	// from the disk in response to the "get text" message):
	if (sab.total_length >= g_MaxVarCapacity)    // Allow the command to succeed by truncating the text.
		sab.total_length = g_MaxVarCapacity - 1; // And this length will be used to limit the retrieval capacity below.

	// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
	// this call will set up the clipboard for writing:
	if (output_var->Assign(NULL, (VarSizeType)sab.total_length) != OK)
		return FAIL;  // It already displayed the error.

	// Fetch the text directly into the var.  Also set the length explicitly
	// in case actual size written was different than the esimated size (since
	// GetWindowTextLength() can return more space that will actually be required
	// in certain circumstances, see MSDN):
	sab.buf = output_var->Contents();
	sab.total_length = 0; // Init
	// Note: The capacity member below exists because granted capacity might be a little larger than we asked for,
	// which allows the actual text fetched to be larger than the length estimate retrieved by the first pass
	// (which generally shouldn't happen since MSDN docs say that the actual length can be less, but never greater,
	// than the estimate length):
	sab.capacity = output_var->Capacity(); // Capacity includes the zero terminator, i.e. it's the size of the memory area.
	EnumChildWindows(target_window, EnumChildGetText, (LPARAM)&sab); // Get the text.

	// Length is set explicitly below in case it wound up being smaller than expected/estimated.
	// MSDN says that can happen generally, and also specifically because: "ANSI applications may have
	// the string in the buffer reduced in size (to a minimum of half that of the wParam value) due to
	// conversion from ANSI to Unicode."
	output_var->Length() = (VarSizeType)sab.total_length;
	if (sab.total_length)
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	else // Something went wrong, so make sure we set to empty string.
		*sab.buf = '\0';  // Safe because Assign() gave us a non-constant memory area.
	return output_var->Close();  // In case it's the clipboard.
}



BOOL CALLBACK EnumChildGetText(HWND aWnd, LPARAM lParam)
{
	if (!g.DetectHiddenText && !IsWindowVisible(aWnd))
		return TRUE;  // This child/control is hidden and user doesn't want it considered, so skip it.
	length_and_buf_type &lab = *((length_and_buf_type *)lParam);  // For performance and convenience.
	int length;
	if (lab.buf)
		length = GetWindowTextTimeout(aWnd, lab.buf + lab.total_length
			, (int)(lab.capacity - lab.total_length)); // Not +1.  Verified correct because WM_GETTEXT accepts size of buffer, not its length.
	else
		length = GetWindowTextTimeout(aWnd);
	lab.total_length += length;
	if (length)
	{
		if (lab.buf)
		{
			if (lab.capacity - lab.total_length > 2) // Must be >2 due to zero terminator.
			{
				strcpy(lab.buf + lab.total_length, "\r\n"); // Something to delimit each control's text.
				lab.total_length += 2;
			}
			// else don't increment total_length
		}
		else
			lab.total_length += 2; // Since buf is NULL, accumulate the size that *would* be needed.
	}
	return TRUE; // Continue enumeration through all the child windows of this parent.
}



ResultType Line::WinGetPos(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	Var *output_var_x = ResolveVarOfArg(0);  // Ok if NULL.
	Var *output_var_y = ResolveVarOfArg(1);  // Ok if NULL.
	Var *output_var_width = ResolveVarOfArg(2);  // Ok if NULL.
	Var *output_var_height = ResolveVarOfArg(3);  // Ok if NULL.

	HWND target_window = DetermineTargetWindow(aTitle, aText, aExcludeTitle, aExcludeText);
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
			if (!output_var_x->Assign(rect.left))  // X position
				result = FAIL;
		}
		else
			if (!output_var_x->Assign(""))
				result = FAIL;
	if (output_var_y)
		if (target_window)
		{
			if (!output_var_y->Assign(rect.top))  // Y position
				result = FAIL;
		}
		else
			if (!output_var_y->Assign(""))
				result = FAIL;
	if (output_var_width) // else user didn't want this value saved to an output param
		if (target_window)
		{
			if (!output_var_width->Assign(rect.right - rect.left))  // Width
				result = FAIL;
		}
		else
			if (!output_var_width->Assign("")) // Set it to be empty to signal the user that the window wasn't found.
				result = FAIL;
	if (output_var_height)
		if (target_window)
		{
			if (!output_var_height->Assign(rect.bottom - rect.top))  // Height
				result = FAIL;
		}
		else
			if (!output_var_height->Assign(""))
				result = FAIL;

	return result;
}



ResultType Line::SysGet(char *aCmd, char *aValue)
// Thanks to Gregory F. Hogg of Hogg's Software for providing sample code on which this function
// is based.
{
	// For simplicity and array look-up performance, this is done even for sub-commands that output to an array:
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;

	SysGetCmds cmd = ConvertSysGetCmd(aCmd);
	// Since command names are validated at load-time, this only happens if the command name
	// was contained in a variable reference.  But for simplicity of design here, return
	// failure in this case (unlike other functions similar to this one):
	if (cmd == SYSGET_CMD_INVALID)
		return LineError(ERR_PARAM2_INVALID ERR_ABORT, FAIL, aCmd);

	MonitorInfoPackage mip = {0};  // Improves maintainability to initialize unconditionally, here.
	mip.monitor_info_ex.cbSize = sizeof(MONITORINFOEX); // Also improves maintainability.

	// EnumDisplayMonitors() must be dynamically loaded; otherwise, the app won't launch at all on Win95/NT.
	typedef BOOL (WINAPI* EnumDisplayMonitorsType)(HDC, LPCRECT, MONITORENUMPROC, LPARAM);
	static EnumDisplayMonitorsType MyEnumDisplayMonitors = (EnumDisplayMonitorsType)
		GetProcAddress(GetModuleHandle("user32"), "EnumDisplayMonitors");

	switch(cmd)
	{
	case SYSGET_CMD_METRICS: // In this case, aCmd is the value itself.
		return output_var->Assign(GetSystemMetrics(ATOI(aCmd)));  // Input and output are both signed integers.

	// For the next few cases, I'm not sure if it is possible to have zero monitors.  Obviously it's possible
	// to not have a monitor turned on or not connected at all.  But it seems likely that these various API
	// functions will provide a "default monitor" in the absence of a physical monitor connected to the
	// system.  To be safe, all of the below will assume that zero is possible, at least on some OSes or
	// under some conditions.  However, on Win95/NT, "1" is assumed since there is probably no way to tell
	// for sure if there are zero monitors except via GetSystemMetrics(SM_CMONITORS), which is a different
	// animal as described below.
	case SYSGET_CMD_MONITORCOUNT:
		// Don't use GetSystemMetrics(SM_CMONITORS) because of this:
		// MSDN: "GetSystemMetrics(SM_CMONITORS) counts only display monitors. This is different from
		// EnumDisplayMonitors, which enumerates display monitors and also non-display pseudo-monitors."
		if (!MyEnumDisplayMonitors) // Since system only supports 1 monitor, the first must be primary.
			return output_var->Assign(1); // Assign as 1 vs. "1" to use hexidecimal display if that is in effect.
		mip.monitor_number_to_find = COUNT_ALL_MONITORS;
		MyEnumDisplayMonitors(NULL, NULL, EnumMonitorProc, (LPARAM)&mip);
		return output_var->Assign(mip.count); // Will assign zero if the API ever returns a legitimate zero.

	// Even if the first monitor to be retrieved by the EnumProc is always the primary (which is doubtful
	// since there's no mention of this in the MSDN docs) it seems best to have this sub-cmd in case that
	// policy ever changes:
	case SYSGET_CMD_MONITORPRIMARY:
		if (!MyEnumDisplayMonitors) // Since system only supports 1 monitor, the first must be primary.
			return output_var->Assign(1); // Assign as 1 vs. "1" to use hexidecimal display if that is in effect.
		// The mip struct's values have already initalized correctly for the below:
		MyEnumDisplayMonitors(NULL, NULL, EnumMonitorProc, (LPARAM)&mip);
		return output_var->Assign(mip.count); // Will assign zero if the API ever returns a legitimate zero.

	case SYSGET_CMD_MONITORAREA:
	case SYSGET_CMD_MONITORWORKAREA:
		Var *output_var_left, *output_var_top, *output_var_right, *output_var_bottom;
		// Make it longer than max var name so that FindOrAddVar() will be able to spot and report
		// var names that are too long:
		char var_name[MAX_VAR_NAME_LENGTH + 20];
		// To help performance (in case the linked list of variables is huge), tell FindOrAddVar where
		// to start the search.  Use the base array name rather than the preceding element because,
		// for example, Array19 is alphabetially less than Array2, so we can't rely on the
		// numerical ordering:
		int always_use;
		always_use = output_var->IsLocal() ? ALWAYS_USE_LOCAL : ALWAYS_USE_GLOBAL;
		snprintf(var_name, sizeof(var_name), "%sLeft", output_var->mName);
		if (   !(output_var_left = g_script.FindOrAddVar(var_name, 0, always_use))   )
			return FAIL;  // It already reported the error.
		snprintf(var_name, sizeof(var_name), "%sTop", output_var->mName);
		if (   !(output_var_top = g_script.FindOrAddVar(var_name, 0, always_use))   )
			return FAIL;
		snprintf(var_name, sizeof(var_name), "%sRight", output_var->mName);
		if (   !(output_var_right = g_script.FindOrAddVar(var_name, 0, always_use))   )
			return FAIL;
		snprintf(var_name, sizeof(var_name), "%sBottom", output_var->mName);
		if (   !(output_var_bottom = g_script.FindOrAddVar(var_name, 0, always_use))   )
			return FAIL;

		RECT monitor_rect;
		if (MyEnumDisplayMonitors)
		{
			mip.monitor_number_to_find = ATOI(aValue);  // If this returns 0, it will default to the primary monitor.
			MyEnumDisplayMonitors(NULL, NULL, EnumMonitorProc, (LPARAM)&mip);
			if (!mip.count || (mip.monitor_number_to_find && mip.monitor_number_to_find != mip.count))
			{
				// With the exception of the caller having specified a non-existent monitor number, all of
				// the ways the above can happen are probably impossible in practice.  Make all the variables
				// blank vs. zero to indicate the problem.
				output_var_left->Assign();
				output_var_top->Assign();
				output_var_right->Assign();
				output_var_bottom->Assign();
				return OK;
			}
			// Otherwise:
			monitor_rect = (cmd == SYSGET_CMD_MONITORAREA) ? mip.monitor_info_ex.rcMonitor : mip.monitor_info_ex.rcWork;
		}
		else // Win95/NT: Since system only supports 1 monitor, the first must be primary.
		{
			if (cmd == SYSGET_CMD_MONITORAREA)
			{
				monitor_rect.left = 0;
				monitor_rect.top = 0;
				monitor_rect.right = GetSystemMetrics(SM_CXSCREEN);
				monitor_rect.bottom = GetSystemMetrics(SM_CYSCREEN);
			}
			else // Work area
				SystemParametersInfo(SPI_GETWORKAREA, 0, &monitor_rect, 0);  // Get desktop rect excluding task bar.
		}
		output_var_left->Assign(monitor_rect.left);
		output_var_top->Assign(monitor_rect.top);
		output_var_right->Assign(monitor_rect.right);
		output_var_bottom->Assign(monitor_rect.bottom);
		return OK;

	case SYSGET_CMD_MONITORNAME:
		if (MyEnumDisplayMonitors)
		{
			mip.monitor_number_to_find = ATOI(aValue);  // If this returns 0, it will default to the primary monitor.
			MyEnumDisplayMonitors(NULL, NULL, EnumMonitorProc, (LPARAM)&mip);
			if (!mip.count || (mip.monitor_number_to_find && mip.monitor_number_to_find != mip.count))
				// With the exception of the caller having specified a non-existent monitor number, all of
				// the ways the above can happen are probably impossible in practice.  Make the variable
				// blank to indicate the problem:
				return output_var->Assign();
			else
				return output_var->Assign(mip.monitor_info_ex.szDevice);
		}
		else // Win95/NT: There is probably no way to find out the name of the monitor.
			return output_var->Assign();
	} // switch()

	return FAIL;  // Never executed (increases maintainability and avoids compiler warning).
}



BOOL CALLBACK EnumMonitorProc(HMONITOR hMonitor, HDC hdcMonitor, LPRECT lprcMonitor, LPARAM lParam)
{
	MonitorInfoPackage &mip = *((MonitorInfoPackage *)lParam);  // For performance and convenience.
	if (mip.monitor_number_to_find == COUNT_ALL_MONITORS)
	{
		++mip.count;
		return TRUE;  // Enumerate all monitors so that they can be counted.
	}
	// GetMonitorInfo() must be dynamically loaded; otherwise, the app won't launch at all on Win95/NT.
	typedef BOOL (WINAPI* GetMonitorInfoType)(HMONITOR, LPMONITORINFO);
	static GetMonitorInfoType MyGetMonitorInfo = (GetMonitorInfoType)
		GetProcAddress(GetModuleHandle("user32"), "GetMonitorInfoA");
	if (!MyGetMonitorInfo) // Shouldn't normally happen since caller wouldn't have called us if OS is Win95/NT. 
		return FALSE;
	if (!MyGetMonitorInfo(hMonitor, &mip.monitor_info_ex)) // Failed.  Probably very rare.
		return FALSE; // Due to the complexity of needing to stop at the correct monitor number, do not continue.
		// In the unlikely event that the above fails when the caller wanted us to find the primary
		// monitor, the caller will think the primary is the previously found monitor (if any).
		// This is just documented here as a known limitation since this combination of circumstances
		// is probably impossible.
	++mip.count; // So that caller can detect failure, increment only now that failure conditions have been checked.
	if (mip.monitor_number_to_find) // Caller gave a specific monitor number, so don't search for the primary monitor.
	{
		if (mip.count == mip.monitor_number_to_find) // Since the desired monitor has been found, must not continue.
			return FALSE;
	}
	else // Caller wants the primary monitor found.
		// MSDN docs are unclear that MONITORINFOF_PRIMARY is a bitwise value, but the name "dwFlags" implies so:
		if (mip.monitor_info_ex.dwFlags & MONITORINFOF_PRIMARY)
			return FALSE;  // Primary has been found and "count" contains its number. Must not continue the enumeration.
			// Above assumes that it is impossible to not have a primary monitor in a system that has at least
			// one monitor.  MSDN certainly implies this through multiple references to the primary monitor.
	// Otherwise, continue the enumeration:
	return TRUE;
}



LPCOLORREF getbits(HBITMAP ahImage, HDC hdc, LONG &aWidth, LONG &aHeight, bool &aIs16Bit, int aMinColorDepth = 8)
// Helper function used by PixelSearch below.
// Returns an array of pixels to the caller, which it must free when done.  Returns NULL on failure,
// in which case the contents of the output parameters is indeterminate.
{
	HDC tdc = CreateCompatibleDC(hdc);
	if (!tdc)
		return NULL;

	// From this point on, "goto end" will assume tdc is non-NULL, but that the below
	// might still be NULL.  Therefore, all of the following must be initialized so that the "end"
	// label can detect them:
	HGDIOBJ tdc_orig_select = NULL;
	LPCOLORREF image_pixel = NULL;
	bool success = false;

	// Confirmed:
	// Needs extra memory to prevent buffer overflow due to: "A bottom-up DIB is specified by setting
	// the height to a positive number, while a top-down DIB is specified by setting the height to a
	// negative number. THE BITMAP COLOR TABLE WILL BE APPENDED to the BITMAPINFO structure."
	// Maybe this applies only to negative height, in which case the second call to GetDIBits()
	// below uses one.
	struct BITMAPINFO3
	{
		BITMAPINFOHEADER    bmiHeader;
		RGBQUAD             bmiColors[260];  // v1.0.40.10: 260 vs. 3 to allow room for color table when color depth is 8-bit or less.
	} bmi;

	bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	bmi.bmiHeader.biBitCount = 0; // i.e. "query bitmap attributes" only.
	if (!GetDIBits(tdc, ahImage, 0, 0, NULL, (LPBITMAPINFO)&bmi, DIB_RGB_COLORS)
		|| bmi.bmiHeader.biBitCount < aMinColorDepth) // Relies on short-circuit boolean order.
		goto end;

	// Set output parameters for caller:
	aIs16Bit = (bmi.bmiHeader.biBitCount == 16);
	aWidth = bmi.bmiHeader.biWidth;
	aHeight = bmi.bmiHeader.biHeight;

	int image_pixel_count = aWidth * aHeight;
	if (   !(image_pixel = (LPCOLORREF)malloc(image_pixel_count * sizeof(COLORREF)))   )
		goto end;

	// v1.0.40.10: To preserve compatibility with callers who check for transparency in icons, don't do any
	// of the extra color table handling for 1-bpp images.  Update: For code simplification, support only
	// 8-bpp images.  If ever support lower color depths, use something like "bmi.bmiHeader.biBitCount > 1
	// && bmi.bmiHeader.biBitCount < 9";
	bool is_8bit = (bmi.bmiHeader.biBitCount == 8);
	if (!is_8bit)
		bmi.bmiHeader.biBitCount = 32;
	bmi.bmiHeader.biHeight = -bmi.bmiHeader.biHeight; // Storing a negative inside the bmiHeader struct is a signal for GetDIBits().

	// Must be done only after GetDIBits() because: "The bitmap identified by the hbmp parameter
	// must not be selected into a device context when the application calls GetDIBits()."
	// (Although testing shows it works anyway, perhaps because GetDIBits() above is being
	// called in its informational mode only).
	// Note that this seems to return NULL sometimes even though everything still works.
	// Perhaps that is normal.
	tdc_orig_select = SelectObject(tdc, ahImage); // Returns NULL when we're called the second time?

	// Appparently there is no need to specify DIB_PAL_COLORS below when color depth is 8-bit because
	// DIB_RGB_COLORS also retrieves the color indices.
	if (   !(GetDIBits(tdc, ahImage, 0, aHeight, image_pixel, (LPBITMAPINFO)&bmi, DIB_RGB_COLORS))   )
		goto end;

	if (is_8bit) // This section added in v1.0.40.10.
	{
		// Convert the color indicies to RGB colors by going through the array in reverse order.
		// Reverse order allows an in-place conversion of each 8-bit color index to its corresponding
		// 32-bit RGB color.
		LPDWORD palette = (LPDWORD)_alloca(256 * sizeof(PALETTEENTRY));
		GetSystemPaletteEntries(tdc, 0, 256, (LPPALETTEENTRY)palette); // Even if failure can realistically happen, consequences of using uninitialized palette seem acceptable.
		// Above: GetSystemPaletteEntries() is the only approach that provided the correct palette.
		// The following other approaches didn't give the right one:
		// GetDIBits(): The palette it stores in bmi.bmiColors seems completely wrong.
		// GetPaletteEntries()+GetCurrentObject(hdc, OBJ_PAL): Returned only 20 entries rather than the expected 256.
		// GetDIBColorTable(): I think same as above or maybe it returns 0.

		// The following section is necessary because apparently each new row in the region starts on
		// a DWORD boundary.  So if the number of pixels in each row isn't an exact multiple of 4, there
		// are between 1 and 3 zero-bytes at the end of each row.
		int remainder = aWidth % 4;
		int empty_bytes_at_end_of_each_row = remainder ? (4 - remainder) : 0;

		// Start at the last RGB slot and the last color index slot:
		BYTE *byte = (BYTE *)image_pixel + image_pixel_count - 1 + (aHeight * empty_bytes_at_end_of_each_row); // Pointer to 8-bit color indices.
		DWORD *pixel = image_pixel + image_pixel_count - 1; // Pointer to 32-bit RGB entries.

		int row, col;
		for (row = 0; row < aHeight; ++row) // For each row.
		{
			byte -= empty_bytes_at_end_of_each_row;
			for (col = 0; col < aWidth; ++col) // For each column.
				*pixel-- = rgb_to_bgr(palette[*byte--]); // Caller always wants RGB vs. BGR format.
		}
	}
	
	// Since above didn't "goto end", indicate success:
	success = true;

end:
	if (tdc_orig_select) // i.e. the original call to SelectObject() didn't fail.
		SelectObject(tdc, tdc_orig_select); // Probably necessary to prevent memory leak.
	DeleteDC(tdc);
	if (!success && image_pixel)
	{
		free(image_pixel);
		image_pixel = NULL;
	}
	return image_pixel;
}



ResultType Line::PixelSearch(int aLeft, int aTop, int aRight, int aBottom, COLORREF aColorBGR
	, int aVariation, char *aOptions)
// Caller has ensured that aColor is in BGR format unless caller passed true for aUseRGB, in which case
// it's in RGB format.
// Author: The fast-mode PixelSearch was created by Aurelian Maga.
{
	// For maintainability, get options and RGB/BGR conversion out of the way early.
	bool fast_mode = strcasestr(aOptions, "Fast");
	COLORREF aColorRGB;
	if (strcasestr(aOptions, "RGB")) // aColorBGR currently contains an RGB value.
	{
		aColorRGB = aColorBGR;
		aColorBGR = rgb_to_bgr(aColorBGR);
	}
	else
		aColorRGB = rgb_to_bgr(aColorBGR); // rgb_to_bgr() also converts in the reverse direction, i.e. bgr_to_rgb().

	// Many of the following sections are similar to those in ImageSearch(), so they should be
	// maintained together.

	Var *output_var_x = ResolveVarOfArg(0);  // Ok if NULL.
	Var *output_var_y = ResolveVarOfArg(1);  // Ok if NULL.

	g_ErrorLevel->Assign(ERRORLEVEL_ERROR2); // Set default ErrorLevel.  2 means error other than "color not found".
	if (output_var_x)
		output_var_x->Assign();  // Init to empty string regardless of whether we succeed here.
	if (output_var_y)
		output_var_y->Assign();  // Same.

	RECT rect = {0};
	if (!(g.CoordMode & COORD_MODE_PIXEL)) // Using relative vs. screen coordinates.
	{
		if (!GetWindowRect(GetForegroundWindow(), &rect))
			return OK;  // Let ErrorLevel tell the story.
		aLeft   += rect.left;
		aTop    += rect.top;
		aRight  += rect.left;  // Add left vs. right because we're adjusting based on the position of the window.
		aBottom += rect.top;   // Same.
	}

	if (aVariation < 0)
		aVariation = 0;
	if (aVariation > 255)
		aVariation = 255;

	// Allow colors to vary within the spectrum of intensity, rather than having them
	// wrap around (which doesn't seem to make much sense).  For example, if the user specified
	// a variation of 5, but the red component of aColorBGR is only 0x01, we don't want red_low to go
	// below zero, which would cause it to wrap around to a very intense red color:
	COLORREF pixel; // Used much further down.
	BYTE red, green, blue; // Used much further down.
	BYTE search_red, search_green, search_blue;
	BYTE red_low, green_low, blue_low, red_high, green_high, blue_high;
	if (aVariation > 0)
	{
		search_red = GetRValue(aColorBGR);
		search_green = GetGValue(aColorBGR);
		search_blue = GetBValue(aColorBGR);
	}
	//else leave uninitialized since they won't be used.

	HDC hdc = GetDC(NULL);
	if (!hdc)
		return OK;  // Let ErrorLevel tell the story.

	bool found = false; // Must init here for use by "goto fast_end" and for use by both fast and slow modes.

	if (fast_mode)
	{
		// From this point on, "goto fast_end" will assume hdc is non-NULL but that the below might still be NULL.
		// Therefore, all of the following must be initialized so that the "fast_end" label can detect them:
		HDC sdc = NULL;
		HBITMAP hbitmap_screen = NULL;
		LPCOLORREF screen_pixel = NULL;
		HGDIOBJ sdc_orig_select = NULL;

		// Some explanation for the method below is contained in this quote from the newsgroups:
		// "you shouldn't really be getting the current bitmap from the GetDC DC. This might
		// have weird effects like returning the entire screen or not working. Create yourself
		// a memory DC first of the correct size. Then BitBlt into it and then GetDIBits on
		// that instead. This way, the provider of the DC (the video driver) can make sure that
		// the correct pixels are copied across."

		// Create an empty bitmap to hold all the pixels currently visible on the screen (within the search area):
		int search_width = aRight - aLeft + 1;
		int search_height = aBottom - aTop + 1;
		if (   !(sdc = CreateCompatibleDC(hdc)) || !(hbitmap_screen = CreateCompatibleBitmap(hdc, search_width, search_height))   )
			goto fast_end;

		if (   !(sdc_orig_select = SelectObject(sdc, hbitmap_screen))   )
			goto fast_end;

		// Copy the pixels in the search-area of the screen into the DC to be searched:
		if (   !(BitBlt(sdc, 0, 0, search_width, search_height, hdc, aLeft, aTop, SRCCOPY))   )
			goto fast_end;

		LONG screen_width, screen_height;
		bool screen_is_16bit;
		if (   !(screen_pixel = getbits(hbitmap_screen, sdc, screen_width, screen_height, screen_is_16bit))   )
			goto fast_end;

		// Concerning 0xF8F8F8F8: "On 16bit and 15 bit color the first 5 bits in each byte are valid
		// (in 16bit there is an extra bit but i forgot for which color). And this will explain the
		// second problem [in the test script], since GetPixel even in 16bit will return some "valid"
		// data in the last 3bits of each byte."
		register i;
		LONG screen_pixel_count = screen_width * screen_height;
		if (screen_is_16bit)
			for (i = 0; i < screen_pixel_count; ++i)
				screen_pixel[i] &= 0xF8F8F8F8;

		if (aVariation < 1) // Caller wants an exact match on one particular color.
		{
			if (screen_is_16bit)
				aColorRGB &= 0xF8F8F8F8;
			for (i = 0; i < screen_pixel_count; ++i)
			{
				// Note that screen pixels sometimes have a non-zero high-order byte.  That's why
				// bit-and with 0x00FFFFFF is done.  Otherwise, Redish/orangish colors are not properly
				// found:
				if ((screen_pixel[i] & 0x00FFFFFF) == aColorRGB)
				{
					found = true;
					break;
				}
			}
		}
		else
		{
			// It seems more appropriate to do the 16-bit conversion prior to SET_COLOR_RANGE,
			// rather than applying 0xF8 to each of the high/low values individually.
			if (screen_is_16bit)
			{
				search_red &= 0xF8;
				search_green &= 0xF8;
				search_blue &= 0xF8;
			}

#define SET_COLOR_RANGE \
{\
	red_low = (aVariation > search_red) ? 0 : search_red - aVariation;\
	green_low = (aVariation > search_green) ? 0 : search_green - aVariation;\
	blue_low = (aVariation > search_blue) ? 0 : search_blue - aVariation;\
	red_high = (aVariation > 0xFF - search_red) ? 0xFF : search_red + aVariation;\
	green_high = (aVariation > 0xFF - search_green) ? 0xFF : search_green + aVariation;\
	blue_high = (aVariation > 0xFF - search_blue) ? 0xFF : search_blue + aVariation;\
}
			
			SET_COLOR_RANGE

			for (i = 0; i < screen_pixel_count; ++i)
			{
				// Note that screen pixels sometimes have a non-zero high-order byte.  But it doesn't
				// matter with the below approach, since that byte is not checked in the comparison.
				pixel = screen_pixel[i];
				// Because pixel is in RGB vs. BGR format, red is retrieved with GetBValue() and blue
				// is retrieved with GetRValue().
				red = GetBValue(pixel);
				green = GetGValue(pixel);
				blue = GetRValue(pixel);
				if (red >= red_low && red <= red_high
					&& green >= green_low && green <= green_high
					&& blue >= blue_low && blue <= blue_high)
				{
					found = true;
					break;
				}
			}
		}
		if (!found) // Must override ErrorLevel to its new value prior to the label below.
			g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // "1" indicates search completed okay, but didn't find it.

fast_end:
		// If found==false when execution reaches here, ErrorLevel is already set to the right value, so just
		// clean up then return.
		ReleaseDC(NULL, hdc);
		if (sdc)
		{
			if (sdc_orig_select) // i.e. the original call to SelectObject() didn't fail.
				SelectObject(sdc, sdc_orig_select); // Probably necessary to prevent memory leak.
			DeleteDC(sdc);
		}
		if (hbitmap_screen)
			DeleteObject(hbitmap_screen);
		if (screen_pixel)
			free(screen_pixel);

		if (!found) // Let ErrorLevel, which is either "1" or "2" as set earlier, tell the story.
			return OK;

		// Otherwise, success.  Calculate xpos and ypos of where the match was found and adjust
		// coords to make them relative to the position of the target window (rect will contain
		// zeroes if this doesn't need to be done):
		if (output_var_x && !output_var_x->Assign((aLeft + i%screen_width) - rect.left))
			return FAIL;
		if (output_var_y && !output_var_y->Assign((aTop + i/screen_width) - rect.top))
			return FAIL;

		return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	}

	// Otherwise (since above didn't return): fast_mode==false
	// This old/slower method is kept  because fast mode will break older scripts that rely on
	// which match is found if there is more than one match (since fast mode searches the
	// pixels in a different order (horizontally rather than verically, I believe).
	// In addition, there is doubt that the fast mode works in all the screen color depths, games,
	// and other circumstances that the slow mode is known to work in.

	// If the caller gives us inverted X or Y coordinates, conduct the search in reverse order.
	// This feature was requested; it was put into effect for v1.0.25.06.
	bool right_to_left = aLeft > aRight;
	bool bottom_to_top = aTop > aBottom;
	register xpos, ypos;

	if (aVariation > 0)
		SET_COLOR_RANGE

	for (xpos = aLeft  // It starts at aLeft even if right_to_left is true.
		; (right_to_left ? (xpos >= aRight) : (xpos <= aRight)) // Verified correct.
		; xpos += right_to_left ? -1 : 1)
	{
		for (ypos = aTop  // It starts at aTop even if bottom_to_top is true.
			; bottom_to_top ? (ypos >= aBottom) : (ypos <= aBottom) // Verified correct.
			; ypos += bottom_to_top ? -1 : 1)
		{
			pixel = GetPixel(hdc, xpos, ypos); // Returns a BGR value, not RGB.
			if (aVariation < 1)  // User wanted an exact match.
			{
				if (pixel == aColorBGR)
				{
					found = true;
					break;
				}
			}
			else  // User specified that some variation in each of the RGB components is allowable.
			{
				red = GetRValue(pixel);
				green = GetGValue(pixel);
				blue = GetBValue(pixel);
				if (red >= red_low && red <= red_high && green >= green_low && green <= green_high
					&& blue >= blue_low && blue <= blue_high)
				{
					found = true;
					break;
				}
			}
		}
		// Check this here rather than in the outer loop's top line because otherwise the loop's
		// increment would make xpos too big by 1:
		if (found)
			break;
	}

	ReleaseDC(NULL, hdc);

	if (!found)
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // This value indicates "color not found".

	// Otherwise, this pixel matches one of the specified color(s).
	// Adjust coords to make them relative to the position of the target window
	// (rect will contain zeroes if this doesn't need to be done):
	if (output_var_x && !output_var_x->Assign(xpos - rect.left))
		return FAIL;
	if (output_var_y && !output_var_y->Assign(ypos - rect.top))
		return FAIL;
	// Since above didn't return:
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
}



ResultType Line::ImageSearch(int aLeft, int aTop, int aRight, int aBottom, char *aImageFile)
// Author: ImageSearch was created by Aurelian Maga.
{
	// Many of the following sections are similar to those in PixelSearch(), so they should be
	// maintained together.
	Var *output_var_x = ResolveVarOfArg(0);  // Ok if NULL.
	Var *output_var_y = ResolveVarOfArg(1);  // Ok if NULL.

	// Set default results, both ErrorLevel and output variables, in case of early return:
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR2);  // 2 means error other than "image not found".
	if (output_var_x)
		output_var_x->Assign();  // Init to empty string regardless of whether we succeed here.
	if (output_var_y)
		output_var_y->Assign(); // Same.

	RECT rect = {0}; // Set default (for CoordMode == "screen").
	if (!(g.CoordMode & COORD_MODE_PIXEL)) // Using relative vs. screen coordinates.
	{
		if (!GetWindowRect(GetForegroundWindow(), &rect))
			return OK; // Let ErrorLevel tell the story.
		aLeft   += rect.left;
		aTop    += rect.top;
		aRight  += rect.left;  // Add left vs. right because we're adjusting based on the position of the window.
		aBottom += rect.top;   // Same.
	}

	// Options are done as asterisk+option to permit future expansion.
	// Set defaults to be possibly overridden by any specified options:
	int aVariation = 0;  // This is named aVariation vs. variation for use with the SET_COLOR_RANGE macro.
	COLORREF trans_color = CLR_NONE; // The default must be a value that can't occur naturally in an image.
	int icon_index = -1; // This is the value that should be passed to LoadPicture() when there is no preference about which icon to use (in a multi-icon file).
	int width = 0, height = 0;
	// For icons, override the default to be 16x16 because that is what is sought 99% of the time.
	// This new default can be overridden by explicitly specifying w0 h0:
	char *cp = strrchr(aImageFile, '.');
	if (cp)
	{
		++cp;
		if (!(stricmp(cp, "ico") && stricmp(cp, "exe") && stricmp(cp, "dll")))
			width = GetSystemMetrics(SM_CXSMICON), height = GetSystemMetrics(SM_CYSMICON);
	}

	char color_name[32], *dp;
	cp = omit_leading_whitespace(aImageFile); // But don't alter aImageFile yet in case it contains literal whitespace we want to retain.
	while (*cp == '*')
	{
		++cp;
		switch (toupper(*cp))
		{
		case 'W': width = ATOI(cp + 1); break;
		case 'H': height = ATOI(cp + 1); break;
		default:
			if (!strnicmp(cp, "Icon", 4))
			{
				cp += 4;  // Now it's the character after the word.
				icon_index = ATOI(cp) - 1; // LoadPicture() correctly handles any negative value.
			}
			else if (!strnicmp(cp, "Trans", 5))
			{
				cp += 5;  // Now it's the character after the word.
				// Isolate the color name/number for ColorNameToBGR():
				strlcpy(color_name, cp, sizeof(color_name));
				if (dp = StrChrAny(color_name, " \t")) // Find space or tab, if any.
					*dp = '\0';
				trans_color = ColorNameToBGR(color_name);
				if (trans_color == CLR_NONE) // A matching color name was not found, so assume it's in hex format.
					// It seems strtol() automatically handles the optional leading "0x" if present:
					trans_color = rgb_to_bgr(strtol(color_name, NULL, 16));
					// if cp did not contain something hex-numeric, black (0x00) will be assumed,
					// which seems okay given how rare such a problem would be.
			}
			else // Assume it's a number since that's the only other asterisk-option.
			{
				aVariation = ATOI(cp); // Seems okay to support hex via ATOI because the space after the number is documented as being mandatory.
				if (aVariation < 0)
					aVariation = 0;
				if (aVariation > 255)
					aVariation = 255;
				// Note: because it's possible for filenames to start with a space (even though Explorer itself
				// won't let you create them that way), allow exactly one space between end of option and the
				// filename itself:
			}
		} // switch()
		if (   !(cp = StrChrAny(cp, " \t"))   ) // Find the first space or tab after the option.
			return OK; // Bad option/format.  Let ErrorLevel tell the story.
		// Now it's the space or tab (if there is one) after the option letter.  Advance by exactly one character
		// because only one space or tab is considered the delimiter.  Any others are considered to be part of the
		// filename (though some or all OSes might simply ignore them or tolerate them as first-try match criteria).
		aImageFile = ++cp; // This should now point to another asterisk or the filename itself.
		// Above also serves to reset the filename to omit the option string whenever at least one asterisk-option is present.
		cp = omit_leading_whitespace(cp); // This is done to make it more tolerant of having more than one space/tab between options.
	}

	// Update: Transparency is now supported in icons by using the icon's mask.  In addition, an attempt
	// is made to support transparency in GIF, PNG, and possibly TIF files via the *Trans option, which
	// assumes that one color in the image is transparent.  In GIFs not loaded via GDIPlus, the transparent
	// color might always been seen as pure white, but when GDIPlus is used, it's probably always black
	// like it is in PNG -- however, this will not relied upon, at least not until confirmed.
	// OLDER/OBSOLETE comment kept for background:
	// For now, images that can't be loaded as bitmaps (icons and cursors) are not supported because most
	// icons have a transparent background or color present, which the image search routine here is
	// probably not equipped to handle (since the transparent color, when shown, typically reveals the
	// color of whatever is behind it; thus screen pixel color won't match image's pixel color).
	// So currently, only BMP and GIF seem to work reliably, though some of the other GDIPlus-supported
	// formats might work too.
	int image_type;
	HBITMAP hbitmap_image = LoadPicture(aImageFile, width, height, image_type, icon_index, false);
	// The comment marked OBSOLETE below is no longer true because the elimination of the high-byte via
	// 0x00FFFFFF seems to have fixed it.  But "true" is still not passed because that should increase
	// consistency when GIF/BMP/ICO files are used by a script on both Win9x and other OSs (since the
	// same loading method would be used via "false" for these formats across all OSes).
	// OBSOLETE: Must not pass "true" with the above because that causes bitmaps and gifs to be not found
	// by the search.  In other words, nothing works.  Obsolete comment: Pass "true" so that an attempt
	// will be made to load icons as bitmaps if GDIPlus is available.
	if (!hbitmap_image)
		return OK; // Let ErrorLevel tell the story.

	HDC hdc = GetDC(NULL);
	if (!hdc)
	{
		DeleteObject(hbitmap_image);
		return OK; // Let ErrorLevel tell the story.
	}

	// From this point on, "goto end" will assume hdc and hbitmap_image are non-NULL, but that the below
	// might still be NULL.  Therefore, all of the following must be initialized so that the "end"
	// label can detect them:
	HDC sdc = NULL;
	HBITMAP hbitmap_screen = NULL;
	LPCOLORREF image_pixel = NULL, screen_pixel = NULL, image_mask = NULL;
	HGDIOBJ sdc_orig_select = NULL;
	bool found = false; // Must init here for use by "goto end".
    
	bool image_is_16bit;
	LONG image_width, image_height;

	if (image_type == IMAGE_ICON)
	{
		// Must be done prior to IconToBitmap() since it deletes (HICON)hbitmap_image:
		ICONINFO ii;
		if (GetIconInfo((HICON)hbitmap_image, &ii))
		{
			// If the icon is monochrome (black and white), ii.hbmMask will contain twice as many pixels as
			// are actually in the icon.  But since the top half of the pixels are the AND-mask, it seems
			// okay to get all the pixels given the rarity of monochrome icons.  This scenario should be
			// handled properly because: 1) the variables image_height and image_width will be overridden
			// further below with the correct icon dimensions; 2) Only the first half of the pixels within
			// the image_mask array will actually be referenced by the transparency checker in the loops,
			// and that first half is the AND-mask, which is the transparency part that is needed.  The
			// second half, the XOR part, is not needed and thus ignored.  Also note that if width/height
			// required the icon to be scaled, LoadPicture() has already done that directly to the icon,
			// so ii.hbmMask should already be scaled to match the size of the bitmap created later below.
			image_mask = getbits(ii.hbmMask, hdc, image_width, image_height, image_is_16bit, 1);
			DeleteObject(ii.hbmColor); // DeleteObject() probably handles NULL okay since few MSDN/other examples ever check for NULL.
			DeleteObject(ii.hbmMask);
		}
		if (   !(hbitmap_image = IconToBitmap((HICON)hbitmap_image, true))   )
			return OK; // Let ErrorLevel tell the story.
	}

	if (   !(image_pixel = getbits(hbitmap_image, hdc, image_width, image_height, image_is_16bit))   )
		goto end;

	// Create an empty bitmap to hold all the pixels currently visible on the screen that lie within the search area:
	int search_width = aRight - aLeft + 1;
	int search_height = aBottom - aTop + 1;
	if (   !(sdc = CreateCompatibleDC(hdc)) || !(hbitmap_screen = CreateCompatibleBitmap(hdc, search_width, search_height))   )
		goto end;

	if (   !(sdc_orig_select = SelectObject(sdc, hbitmap_screen))   )
		goto end;

	// Copy the pixels in the search-area of the screen into the DC to be searched:
	if (   !(BitBlt(sdc, 0, 0, search_width, search_height, hdc, aLeft, aTop, SRCCOPY))   )
		goto end;

	LONG screen_width, screen_height;
	bool screen_is_16bit;
	if (   !(screen_pixel = getbits(hbitmap_screen, sdc, screen_width, screen_height, screen_is_16bit))   )
		goto end;

	LONG image_pixel_count = image_width * image_height;
	LONG screen_pixel_count = screen_width * screen_height;
	int i, j, k, x, y; // Declaring as "register" makes no performance difference with current compiler, so let the compiler choose which should be registers.

	// If either is 16-bit, convert *both* to the 16-bit-compatible 32-bit format:
	if (image_is_16bit || screen_is_16bit)
	{
		if (trans_color != CLR_NONE)
			trans_color &= 0x00F8F8F8; // Convert indicated trans-color to be compatible with the conversion below.
		for (i = 0; i < screen_pixel_count; ++i)
			screen_pixel[i] &= 0x00F8F8F8; // Highest order byte must be masked to zero for consistency with use of 0x00FFFFFF below.
		for (i = 0; i < image_pixel_count; ++i)
			image_pixel[i] &= 0x00F8F8F8;  // Same.
	}

	// Search the specified region for the first occurrence of the image:
	if (aVariation < 1) // Caller wants an exact match.
	{
		// Concerning the following use of 0x00FFFFFF, the use of 0x00F8F8F8 above is related (both have high order byte 00).
		// The following needs to be done only when shades-of-variation mode isn't in effect because
		// shades-of-variation mode ignores the high-order byte due to its use of macros such as GetRValue().
		// This transformation incurs about a 15% performance decrease (percentage is fairly constant since
		// it is proportional to the search-region size, which tends to be much larger than the search-image and
		// is therefore the primary determination of how long the loops take). But it definitely helps find images
		// more successfully in some cases.  For example, if a PNG file is displayed in a GUI window, this
		// transformation allows certain bitmap search-images to be found via variation==0 when they otherwise
		// would require variation==1 (possibly the variation==1 success is just a side-effect of it
		// ignoring the high-order byte -- maybe a much higher variation would be needed if the high
		// order byte were also subject to the same shades-of-variation analysis as the other three bytes [RGB]).
		for (i = 0; i < screen_pixel_count; ++i)
			screen_pixel[i] &= 0x00FFFFFF;
		for (i = 0; i < image_pixel_count; ++i)
			image_pixel[i] &= 0x00FFFFFF;

		for (i = 0; i < screen_pixel_count; ++i)
		{
			// Unlike the variation-loop, the following one uses a first-pixel optimization to boost performance
			// by about 10% because it's only 3 extra comparisons and exact-match mode is probably used more often.
			// Before even checking whether the other adjacent pixels in the region match the image, ensure
			// the image does not extend past the right or bottom edges of the current part of the search region.
			// This is done for performance but more importantly to prevent partial matches at the edges of the
			// search region from being considered complete matches.
			// The following check is ordered for short-circuit performance.  In addition, image_mask, if
			// non-NULL, is used to determine which pixels are transparent within the image and thus should
			// match any color on the screen.
			if ((screen_pixel[i] == image_pixel[0] // A screen pixel has been found that matches the image's first pixel.
				|| image_mask && image_mask[0]     // Or: It's an icon's transparent pixel, which matches any color.
				|| image_pixel[0] == trans_color)  // This should be okay even if trans_color==CLR_NONE, since CLR none should never occur naturally in the image.
				&& image_height <= screen_height - i/screen_width // Image is short enough to fit in the remaining rows of the search region.
				&& image_width <= screen_width - i%screen_width)  // Image is narrow enough not to exceed the right-side boundary of the search region.
			{
				// Check if this candidate region -- which is a subset of the search region whose height and width
				// matches that of the image -- is a pixel-for-pixel match of the image.
				for (found = true, x = 0, y = 0, j = 0, k = i; j < image_pixel_count; ++j)
				{
					if (!(found = (screen_pixel[k] == image_pixel[j] // At least one pixel doesn't match, so this candidate is discarded.
						|| image_mask && image_mask[j]      // Or: It's an icon's transparent pixel, which matches any color.
						|| image_pixel[j] == trans_color))) // This should be okay even if trans_color==CLR_NONE, since CLR none should never occur naturally in the image.
						break;
					if (++x < image_width) // We're still within the same row of the image, so just move on to the next screen pixel.
						++k;
					else // We're starting a new row of the image.
					{
						x = 0; // Return to the leftmost column of the image.
						++y;   // Move one row downward in the image.
						// Move to the next row within the current-candiate region (not the entire search region).
						// This is done by moving vertically downward from "i" (which is the upper-left pixel of the
						// current-candidate region) by "y" rows.
						k = i + y*screen_width; // Verified correct.
					}
				}
				if (found) // Complete match found.
					break;
			}
		}
	}
	else // Allow colors to vary by aVariation shades; i.e. approximate match is okay.
	{
		// The following section is part of the first-pixel-check optimization that improves performance by
		// 15% or more depending on where and whether a match is found.  This section and one the follows
		// later is commented out to reduce code size.
		// Set high/low range for the first pixel of the image since it is the pixel most often checked
		// (i.e. for performance).
		//BYTE search_red1 = GetBValue(image_pixel[0]);  // Because it's RGB vs. BGR, the B value is fetched, not R (though it doesn't matter as long as everything is internally consistent here).
		//BYTE search_green1 = GetGValue(image_pixel[0]);
		//BYTE search_blue1 = GetRValue(image_pixel[0]); // Same comment as above.
		//BYTE red_low1 = (aVariation > search_red1) ? 0 : search_red1 - aVariation;
		//BYTE green_low1 = (aVariation > search_green1) ? 0 : search_green1 - aVariation;
		//BYTE blue_low1 = (aVariation > search_blue1) ? 0 : search_blue1 - aVariation;
		//BYTE red_high1 = (aVariation > 0xFF - search_red1) ? 0xFF : search_red1 + aVariation;
		//BYTE green_high1 = (aVariation > 0xFF - search_green1) ? 0xFF : search_green1 + aVariation;
		//BYTE blue_high1 = (aVariation > 0xFF - search_blue1) ? 0xFF : search_blue1 + aVariation;
		// Above relies on the fact that the 16-bit conversion higher above was already done because like
		// in PixelSearch, it seems more appropriate to do the 16-bit conversion prior to setting the range
		// of high and low colors (vs. than applying 0xF8 to each of the high/low values individually).

		BYTE red, green, blue;
		BYTE search_red, search_green, search_blue;
		BYTE red_low, green_low, blue_low, red_high, green_high, blue_high;

		// The following loop is very similar to its counterpart above that finds an exact match, so maintain
		// them together and see above for more detailed comments about it.
		for (i = 0; i < screen_pixel_count; ++i)
		{
			// The following is commented out to trade code size reduction for performance (see comment above).
			//red = GetBValue(screen_pixel[i]);   // Because it's RGB vs. BGR, the B value is fetched, not R (though it doesn't matter as long as everything is internally consistent here).
			//green = GetGValue(screen_pixel[i]);
			//blue = GetRValue(screen_pixel[i]);
			//if ((red >= red_low1 && red <= red_high1
			//	&& green >= green_low1 && green <= green_high1
			//	&& blue >= blue_low1 && blue <= blue_high1 // All three color components are a match, so this screen pixel matches the image's first pixel.
			//		|| image_mask && image_mask[0]         // Or: It's an icon's transparent pixel, which matches any color.
			//		|| image_pixel[0] == trans_color)      // This should be okay even if trans_color==CLR_NONE, since CLR none should never occur naturally in the image.
			//	&& image_height <= screen_height - i/screen_width // Image is short enough to fit in the remaining rows of the search region.
			//	&& image_width <= screen_width - i%screen_width)  // Image is narrow enough not to exceed the right-side boundary of the search region.
			
			// Instead of the above, only this abbreviated check is done:
			if (image_height <= screen_height - i/screen_width    // Image is short enough to fit in the remaining rows of the search region.
				&& image_width <= screen_width - i%screen_width)  // Image is narrow enough not to exceed the right-side boundary of the search region.
			{
				// Since the first pixel is a match, check the other pixels.
				for (found = true, x = 0, y = 0, j = 0, k = i; j < image_pixel_count; ++j)
				{
   					search_red = GetBValue(image_pixel[j]);
	   				search_green = GetGValue(image_pixel[j]);
		   			search_blue = GetRValue(image_pixel[j]);
					SET_COLOR_RANGE
   					red = GetBValue(screen_pixel[k]);
	   				green = GetGValue(screen_pixel[k]);
		   			blue = GetRValue(screen_pixel[k]);

					if (!(found = red >= red_low && red <= red_high
						&& green >= green_low && green <= green_high
                        && blue >= blue_low && blue <= blue_high
							|| image_mask && image_mask[j]     // Or: It's an icon's transparent pixel, which matches any color.
							|| image_pixel[j] == trans_color)) // This should be okay even if trans_color==CLR_NONE, since CLR none should never occur naturally in the image.
						break; // At least one pixel doesn't match, so this candidate is discarded.
					if (++x < image_width) // We're still within the same row of the image, so just move on to the next screen pixel.
						++k;
					else // We're starting a new row of the image.
					{
						x = 0; // Return to the leftmost column of the image.
						++y;   // Move one row downward in the image.
						k = i + y*screen_width; // Verified correct.
					}
				}
				if (found) // Complete match found.
					break;
			}
		}
	}

	if (!found) // Must override ErrorLevel to its new value prior to the label below.
		g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // "1" indicates search completed okay, but didn't find it.

end:
	// If found==false when execution reaches here, ErrorLevel is already set to the right value, so just
	// clean up then return.
	ReleaseDC(NULL, hdc);
	DeleteObject(hbitmap_image);
	if (sdc)
	{
		if (sdc_orig_select) // i.e. the original call to SelectObject() didn't fail.
			SelectObject(sdc, sdc_orig_select); // Probably necessary to prevent memory leak.
		DeleteDC(sdc);
	}
	if (hbitmap_screen)
		DeleteObject(hbitmap_screen);
	if (image_pixel)
		free(image_pixel);
	if (image_mask)
		free(image_mask);
	if (screen_pixel)
		free(screen_pixel);

	if (!found) // Let ErrorLevel, which is either "1" or "2" as set earlier, tell the story.
		return OK;

	// Otherwise, success.  Calculate xpos and ypos of where the match was found and adjust
	// coords to make them relative to the position of the target window (rect will contain
	// zeroes if this doesn't need to be done):
	if (output_var_x && !output_var_x->Assign((aLeft + i%screen_width) - rect.left))
		return FAIL;
	if (output_var_y && !output_var_y->Assign((aTop + i/screen_width) - rect.top))
		return FAIL;

	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
}



/////////////////
// Main Window //
/////////////////

LRESULT CALLBACK MainWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
	int i;
	DWORD dwTemp;

	// Detect Explorer crashes so that tray icon can be recreated.  I think this only works on Win98
	// and beyond, since the feature was never properly implemented in Win95:
	static UINT WM_TASKBARCREATED = RegisterWindowMessage("TaskbarCreated");

	// See GuiWindowProc() for details about this first section:
	LRESULT msg_reply;
	if (g_MsgMonitorCount && !g.CalledByIsDialogMessageOrDispatch // Count is checked here to avoid function-call overhead.
		&& MsgMonitor(hWnd, iMsg, wParam, lParam, NULL, msg_reply))
		return msg_reply; // MsgMonitor has returned "true", indicating that this message should be omitted from further processing.
	g.CalledByIsDialogMessageOrDispatch = false; // v1.0.40.01.

	TRANSLATE_AHK_MSG(iMsg, wParam)
	
	switch (iMsg)
	{
	case WM_COMMAND:
		if (HandleMenuItem(hWnd, LOWORD(wParam), -1)) // It was handled fully. -1 flags it as a non-GUI menu item such as a tray menu or popup menu.
			return 0; // If an application processes this message, it should return zero.
		break; // Otherwise, let DefWindowProc() try to handle it (this actually seems to happen normally sometimes).

	case AHK_NOTIFYICON:  // Tray icon clicked on.
	{
        switch(lParam)
        {
// Don't allow the main window to be opened this way by a compiled EXE, since it will display
// the lines most recently executed, and many people who compile scripts don't want their users
// to see the contents of the script:
		case WM_LBUTTONDOWN:
			if (g_script.mTrayMenu->mClickCount != 1) // Activating tray menu's default item requires double-click.
				break; // Let default proc handle it (since that's what used to happen, it seems safest).
			//else fall through to the next case.
		case WM_LBUTTONDBLCLK:
			if (g_script.mTrayMenu->mDefault)
				POST_AHK_USER_MENU(hWnd, g_script.mTrayMenu->mDefault->mMenuID, -1) // -1 flags it as a non-GUI menu item.
#ifdef AUTOHOTKEYSC
			else if (g_script.mTrayMenu->mIncludeStandardItems && g_AllowMainWindow)
				ShowMainWindow();
			// else do nothing.
#else
			else if (g_script.mTrayMenu->mIncludeStandardItems)
				ShowMainWindow();
			// else do nothing.
#endif
			return 0;
		case WM_RBUTTONUP:
			// v1.0.30.03:
			// Opening the menu upon UP vs. DOWN solves at least one set of problems: The fact that
			// when the right mouse button is remapped as shown in the example below, it prevents
			// the left button from being able to select a menu item from the tray menu.  It might
			// solve other problems also, and it seems fairly common for other apps to open the
			// menu upon UP rather than down.  Even Explorer's own context menus are like this.
			// The following example is trivial and serves only to illustrate the problem caused
			// by the old open-tray-on-mouse-down method:
			//MButton::Send {RButton down}
			//MButton up::Send {RButton up}
			g_script.mTrayMenu->Display(false);
			return 0;
		} // Inner switch()
		break;
	} // case AHK_NOTIFYICON

	case AHK_DIALOG:  // User defined msg sent from our functions MsgBox() or FileSelectFile().
	{
		// Always call this to close the clipboard if it was open (e.g. due to a script
		// line such as "MsgBox, %clipboard%" that got us here).  Seems better just to
		// do this rather than incurring the delay and overhead of a MsgSleep() call:
		CLOSE_CLIPBOARD_IF_OPEN;
		
		// Ensure that the app's top-most window (the modal dialog) is the system's
		// foreground window.  This doesn't use FindWindow() since it can hang in rare
		// cases.  And GetActiveWindow, GetTopWindow, GetWindow, etc. don't seem appropriate.
		// So EnumWindows is probably the way to do it:
		HWND top_box = FindOurTopDialog();
		if (top_box)
		{

			// v1.0.33: The following is probably reliable since the AHK_DIALOG should
			// be in front of any messages that would launch an interrupting thread.  In other
			// words, the "g" struct should still be the one that owns this MsgBox/dialog window.
			g.DialogHWND = top_box; // This is used to work around an AHK_TIMEOUT issue in which a MsgBox that has only an OK button fails to deliver the Timeout indicator to the script.

			SetForegroundWindowEx(top_box);

			// Setting the big icon makes AutoHotkey dialogs more distinct in the Alt-tab menu.
			// Unfortunately, it seems that setting the big icon also indirectly sets the small
			// icon, or more precisely, that the dialog simply scales the large icon whenever
			// a small one isn't available.  This results in the FileSelectFile dialog's title
			// being initially messed up (at least on WinXP) and also puts an unwanted icon in
			// the title bar of each MsgBox.  So for now it's disabled:
			//LPARAM main_icon = (LPARAM)LoadIcon(g_hInstance, MAKEINTRESOURCE(IDI_MAIN));
			//SendMessage(top_box, WM_SETICON, ICON_BIG, main_icon);
			//SendMessage(top_box, WM_SETICON, ICON_SMALL, 0);  // Tried this to get rid of it, but it didn't help.
			// But don't set the small one, because it reduces the area available for title text
			// without adding any significant benefit:
			//SendMessage(top_box, WM_SETICON, ICON_SMALL, main_icon);

			UINT timeout = (UINT)lParam;  // Caller has ensured that this is non-negative.
			if (timeout)
				// Caller told us to establish a timeout for this modal dialog (currently always MessageBox).
				// In addition to any other reasons, the first param of the below must not be NULL because
				// that would cause the 2nd param to be ignored.  We want the 2nd param to be the actual
				// ID assigned to this timer.
				SetTimer(top_box, g_nMessageBoxes, (UINT)timeout, MsgBoxTimeout);
		}
		// else: if !top_box: no error reporting currently.
		return 0;
	}

	case AHK_USER_MENU:
		// Search for AHK_USER_MENU in GuiWindowProc() for comments about why this is done:
		PostMessage(hWnd, iMsg, wParam, lParam);
		MsgSleep(-1);
		return 0;

	case WM_HOTKEY: // As a result of this app having previously called RegisterHotkey().
	case AHK_HOOK_HOTKEY:  // Sent from this app's keyboard or mouse hook.
	case AHK_HOTSTRING: // Added for v1.0.36.02 so that hotstrings work even while an InputBox or other non-standard msg pump is running.
		// If the following facts are ever confirmed, there would be no need to post the message in cases where
		// the MsgSleep() won't be done:
		// 1) The mere fact that any of the above messages has been received here in MainWindowProc means that a
		//    message pump other than our own main one is running (i.e. it is the closest pump on the call stack).
		//    This is because our main message pump would never have dispatched the types of messages above because
		//    it is designed to fully handle then discard them.
		// 2) All of these types of non-main message pumps would discard a message with a NULL hwnd.
		//
		// One source of confusion is that there are quite a few different types of message pumps that might
		// be running:
		// - InputBox/MsgBox, or other dialog
		// - Popup menu (tray menu, popup menu from Menu command, or context menu of an Edit/MonthCal, including
		//   our main window's edit control g_hWndEdit).
		// - Probably others, such as ListView marquee-drag, that should be listed here as they are
		//   remembered/discovered.
		//
		// Due to maintainability and the uncertainty over backward compatibility (see comments above), the
		// following message is posted even when INTERRUPTIBLE==false.
		// Post it with a NULL hwnd (update: also for backward compatibility) to avoid any chance that our
		// message pump will dispatch it back to us.  We want these events to always be handled there,
		// where almost all new quasi-threads get launched.  Update: Even if it were safe in terms of
		// backward compatibility to change NULL to gHwnd, testing shows it causes problems when a hotkey
		// is pressed while one of the script's menus is displayed (at least a menu bar).  For example:
		// *LCtrl::Send {Blind}{Ctrl up}{Alt down}
		// *LCtrl up::Send {Blind}{Alt up}
		PostMessage(NULL, iMsg, wParam, lParam);
		if (INTERRUPTIBLE)
			MsgSleep(-1);
		//else let the other pump discard this hotkey event since in most cases it would do more harm than good
		// (see comments above for why the message is posted even when it is 90% certain it will be discarded
		// in all cases where MsgSleep isn't done).
		return 0;

	case WM_TIMER:
		// MSDN: "When you specify a TimerProc callback function, the default window procedure calls
		// the callback function when it processes WM_TIMER. Therefore, you need to dispatch messages
		// in the calling thread, even when you use TimerProc instead of processing WM_TIMER."
		// MSDN CONTRADICTION: "You can process the message by providing a WM_TIMER case in the window
		// procedure. Otherwise, DispatchMessage will call the TimerProc callback function specified in
		// the call to the SetTimer function used to install the timer."
		// In light of the above, it seems best to let the default proc handle this message if it
		// has a non-NULL lparam:
		if (lParam)
			break;
		// Otherwise, it's the main timer, which is the means by which joystick hotkeys and script timers
		// created via the SetTimer script command continue to execute even while a dialog's message pump
		// is running.  Even if the script is NOT INTERRUPTIBLE (which generally isn't possible, since
		// the mere fact that we're here means that a dialog's message pump dispatched a message to us
		// [since our msg pump would not dispatch this type of msg], which in turn means that the script
		// should be interruptible due to DIALOG_PREP), call MsgSleep() anyway so that joystick
		// hotkeys will be polled.  If any such hotkeys are "newly down" right now, those events queued
		// will be buffered/queued for later, when the script becomes interruptible again.  Also, don't
		// call CheckScriptTimers() or PollJoysticks() directly from here.  See comments at the top of
		// those functions for why.
		// This is an older comment, but I think it might still apply, which is why MsgSleep() is not
		// called when a popup menu or a window's main menu is visible.  We don't really want to run the
		// script's timed subroutines or monitor joystick hotkeys while a menu is displayed anyway:
		// Do not call MsgSleep() while a popup menu is visible because that causes long delays
		// sometime when the user is trying to select a menu (the user's click is ignored and the menu
		// stays visible).  I think this is because MsgSleep()'s PeekMessage() intercepts the user's
		// clicks and is unable to route them to TrackPopupMenuEx()'s message loop, which is the only
		// place they can be properly processed.  UPDATE: This also needs to be done when the MAIN MENU
		// is visible, because testing shows that that menu would otherwise become sluggish too, perhaps
		// more rarely, when timers are running.
		// Other background info:
		// Checking g_MenuIsVisible here prevents timed subroutines from running while the tray menu
		// or main menu is in use.  This is documented behavior, and is desirable most of the time
		// anyway.  But not to do this would produce strange effects because any timed subroutine
		// that took a long time to run might keep us away from the "menu loop", which would result
		// in the menu becoming temporarily unresponsive while the user is in it (and probably other
		// undesired effects).
		if (!g_MenuIsVisible)
			MsgSleep(-1);
		return 0;

	case WM_SYSCOMMAND:
		if ((wParam == SC_CLOSE || wParam == SC_MINIMIZE) && hWnd == g_hWnd) // i.e. behave this way only for main window.
		{
			// The user has either clicked the window's "X" button, chosen "Close"
			// from the system (upper-left icon) menu, or pressed Alt-F4.  In all
			// these cases, we want to hide the window rather than actually closing
			// it.  If the user really wishes to exit the program, a File->Exit
			// menu option may be available, or use the Tray Icon, or launch another
			// instance which will close the previous, etc.  UPDATE: SC_MINIMIZE is
			// now handled this way also so that owned windows (such as Splash and
			// Progress) won't be hidden when the main window is hidden.
			ShowWindow(g_hWnd, SW_HIDE);
			return 0;
		}
		break;

	case WM_CLOSE:
		if (hWnd == g_hWnd) // i.e. not the SplashText window or anything other than the main.
		{
			// Receiving this msg is fairly unusual since SC_CLOSE is intercepted and redefined above.
			// However, it does happen if an external app is asking us to close, such as another
			// instance of this same script during the Reload command.  So treat it in a way similar
			// to the user having chosen Exit from the menu.
			//
			// Leave it up to ExitApp() to decide whether to terminate based upon whether
			// there is an OnExit subroutine, whether that subroutine is already running at
			// the time a new WM_CLOSE is received, etc.  It's also its responsibility to call
			// DestroyWindow() upon termination so that the WM_DESTROY message winds up being
			// received and process in this function (which is probably necessary for a clean
			// termination of the app and all its windows):
			g_script.ExitApp(EXIT_WM_CLOSE);
			return 0;  // Verified correct.
		}
		// Otherwise, some window of ours other than our main window was destroyed
		// (perhaps the splash window).  Let DefWindowProc() handle it:
		break;

	case WM_ENDSESSION: // MSDN: "A window receives this message through its WindowProc function."
		if (wParam) // The session is being ended.
			g_script.ExitApp((lParam & ENDSESSION_LOGOFF) ? EXIT_LOGOFF : EXIT_SHUTDOWN);
		//else a prior WM_QUERYENDSESSION was aborted; i.e. the session really isn't ending.
		return 0;  // Verified correct.

	case AHK_EXIT_BY_RELOAD:
		g_script.ExitApp(EXIT_RELOAD);
		return 0; // Whether ExitApp() terminates depends on whether there's an OnExit subroutine and what it does.

	case AHK_EXIT_BY_SINGLEINSTANCE:
		g_script.ExitApp(EXIT_SINGLEINSTANCE);
		return 0; // Whether ExitApp() terminates depends on whether there's an OnExit subroutine and what it does.

	case WM_DESTROY:
		if (hWnd == g_hWnd) // i.e. not the SplashText window or anything other than the main.
		{
			if (!g_DestroyWindowCalled)
				// This is done because I believe it's possible for a WM_DESTROY message to be received
				// even though we didn't call DestroyWindow() ourselves (e.g. via DefWindowProc() receiving
				// and acting upon a WM_CLOSE or us calling DestroyWindow() directly) -- perhaps the window
				// is being forcibly closed or something else abnormal happened.  Make a best effort to run
				// the OnExit subroutine, if present, even without a main window (testing on an earlier
				// versions shows that most commands work fine without the window). Pass the empty string
				// to tell it to terminate after running the OnExit subroutine:
				g_script.ExitApp(EXIT_DESTROY, "");
			// Do not do PostQuitMessage() here because we don't know the proper exit code.
			// MSDN: "The exit value returned to the system must be the wParam parameter of
			// the WM_QUIT message."
			// If we're here, it means our thread called DestroyWindow() directly or indirectly
			// (currently, it's only called directly).  By returning, our thread should resume
			// execution at the statement after DestroyWindow() in whichever caller called that:
			return 0;  // "If an application processes this message, it should return zero."
		}
		// Otherwise, some window of ours other than our main window was destroyed
		// (perhaps the splash window).  Let DefWindowProc() handle it:
		break;

	case WM_CREATE:
		// MSDN: If an application processes this message, it should return zero to continue
		// creation of the window. If the application returns 1, the window is destroyed and
		// the CreateWindowEx or CreateWindow function returns a NULL handle.
		return 0;

	case WM_ERASEBKGND:
	case WM_CTLCOLORSTATIC:
	case WM_PAINT:
	case WM_SIZE:
	{
		if (iMsg == WM_SIZE)
		{
			if (hWnd == g_hWnd)
			{
				if (wParam == SIZE_MINIMIZED)
					// Minimizing the main window hides it.
					ShowWindow(g_hWnd, SW_HIDE);
				else
					MoveWindow(g_hWndEdit, 0, 0, LOWORD(lParam), HIWORD(lParam), TRUE);
				return 0; // The correct return value for this msg.
			}
			if (hWnd == g_hWndSplash || wParam == SIZE_MINIMIZED)
				break;  // Let DefWindowProc() handle it for splash window and Progress windows.
		}
		else
			if (hWnd == g_hWnd || hWnd == g_hWndSplash)
				break; // Let DWP handle it.

		for (i = 0; i < MAX_SPLASHIMAGE_WINDOWS; ++i)
			if (g_SplashImage[i].hwnd == hWnd)
				break;
		bool is_splashimage = (i < MAX_SPLASHIMAGE_WINDOWS);
		if (!is_splashimage)
		{
			for (i = 0; i < MAX_PROGRESS_WINDOWS; ++i)
				if (g_Progress[i].hwnd == hWnd)
					break;
			if (i == MAX_PROGRESS_WINDOWS) // It's not a progress window either.
				// Let DefWindowProc() handle it (should probably never happen since currently the only
				// other type of window is SplashText, which never receive this msg?)
				break;
		}

		SplashType &splash = is_splashimage ? g_SplashImage[i] : g_Progress[i];
		RECT client_rect;

		switch (iMsg)
		{
		case WM_SIZE:
		{
			// Allow any width/height to be specified so that the window can be "rolled up" to its title bar.
			int new_width = LOWORD(lParam);
			int new_height = HIWORD(lParam);
			if (new_width != splash.width || new_height != splash.height)
			{
				GetClientRect(splash.hwnd, &client_rect);
				int control_width = client_rect.right - (splash.margin_x * 2);
				SPLASH_CALC_YPOS
				// The Y offset for each control should match those used in Splash():
				if (new_width != splash.width)
				{
					if (splash.hwnd_text1) // This control doesn't exist if the main text was originally blank.
						MoveWindow(splash.hwnd_text1, PROGRESS_MAIN_POS, FALSE);
					if (splash.hwnd_bar)
						MoveWindow(splash.hwnd_bar, PROGRESS_BAR_POS, FALSE);
					splash.width = new_width;
				}
				// Move the window EVEN IF new_height == splash.height because otherwise the text won't
				// get re-centered when only the width of the window changes:
				MoveWindow(splash.hwnd_text2, PROGRESS_SUB_POS, FALSE); // Negative height seems handled okay.
				// Specifying true for "repaint" in MoveWindow() is not always enough refresh the text correctly,
				// so this is done instead:
				InvalidateRect(splash.hwnd, &client_rect, TRUE);
				// If the user resizes the window, have that size retained (remembered) until the script
				// explicitly changes it or the script destroys the window.
				splash.height = new_height;
			}
			return 0;  // i.e. completely handled here.
		}
		case WM_CTLCOLORSTATIC:
			if (!splash.hbrush && splash.color_text == CLR_DEFAULT) // Let DWP handle it.
				break;
			// Since we're processing this msg and not DWP, must set background color unconditionally,
			// otherwise plain white will likely be used:
			SetBkColor((HDC)wParam, splash.hbrush ? splash.color_bk : GetSysColor(COLOR_BTNFACE));
			if (splash.color_text != CLR_DEFAULT)
				SetTextColor((HDC)wParam, splash.color_text);
			// Always return a real HBRUSH so that Windows knows we altered the HDC for it to use:
			return (LRESULT)(splash.hbrush ? splash.hbrush : GetSysColorBrush(COLOR_BTNFACE));
		case WM_ERASEBKGND:
		{
			if (splash.pic) // And since there is a pic, its object_width/height should already be valid.
			{
				// MSDN: get width and height of picture
				long hm_width, hm_height;
				splash.pic->get_Width(&hm_width);
				splash.pic->get_Height(&hm_height);
				GetClientRect(splash.hwnd, &client_rect);
				// MSDN: display picture using IPicture::Render
				int ypos = splash.margin_y + (splash.text1_height ? (splash.text1_height + splash.margin_y) : 0);
				splash.pic->Render((HDC)wParam, splash.margin_x, ypos, splash.object_width, splash.object_height
					, 0, hm_height, hm_width, -hm_height, &client_rect);
				// Prevent "flashing" by erasing only the part that hasn't already been drawn:
				ExcludeClipRect((HDC)wParam, splash.margin_x, ypos, splash.margin_x + splash.object_width
					, ypos + splash.object_height);
				HRGN hrgn = CreateRectRgn(0, 0, 1, 1);
				GetClipRgn((HDC)wParam, hrgn);
				FillRgn((HDC)wParam, hrgn, splash.hbrush ? splash.hbrush : GetSysColorBrush(COLOR_BTNFACE));
				DeleteObject(hrgn);
				return 1; // "An application should return nonzero if it erases the background."
			}
			// Otherwise, it's a Progress window (or a SplashImage window with no picture):
			if (!splash.hbrush) // Let DWP handle it.
				break;
			RECT clipbox;
			GetClipBox((HDC)wParam, &clipbox);
			FillRect((HDC)wParam, &clipbox, splash.hbrush);
			return 1; // "An application should return nonzero if it erases the background."
		}
		} // switch()
		break; // Let DWP handle it.
	}
		
	case WM_SETFOCUS:
		if (hWnd == g_hWnd)
		{
			SetFocus(g_hWndEdit);  // Always focus the edit window, since it's the only navigable control.
			return 0;
		}
		break;

	case AHK_RETURN_PID:
		// This is obsolete in light of WinGet's support for fetching the PID of any window.
		// But since it's simple, it is retained for backward compatibility.
		// Rajat wanted this so that it's possible to discover the PID based on the title of each
		// script's main window (i.e. if there are multiple scripts running).  Also note that this
		// msg can be sent via TRANSLATE_AHK_MSG() to prevent it from ever being filtered out (and
		// thus delayed) while the script is uninterruptible.  For example:
		// SendMessage, 0x44, 1029,,, %A_ScriptFullPath% - AutoHotkey
		// SendMessage, 1029,,,, %A_ScriptFullPath% - AutoHotkey  ; Same as above but not sent via TRANSLATE.
		return GetCurrentProcessId(); // Don't use ReplyMessage because then our thread can't reply to itself with this answer.

	case WM_DRAWCLIPBOARD:
		if (g_script.mOnClipboardChangeLabel) // In case it's a bogus msg, it's our responsibility to avoid posting the msg if there's no label to launch.
			PostMessage(NULL, AHK_CLIPBOARD_CHANGE, 0, 0); // It's done this way to buffer it when the script is uninterruptible, etc.
		if (g_script.mNextClipboardViewer) // Will be NULL if there are no other windows in the chain.
			SendMessageTimeout(g_script.mNextClipboardViewer, iMsg, wParam, lParam, SMTO_ABORTIFHUNG, 2000, &dwTemp);
		return 0;

	case WM_CHANGECBCHAIN:
		// MSDN: If the next window is closing, repair the chain. 
		if ((HWND)wParam == g_script.mNextClipboardViewer)
			g_script.mNextClipboardViewer = (HWND)lParam;
		// MSDN: Otherwise, pass the message to the next link. 
		else if (g_script.mNextClipboardViewer)
			SendMessageTimeout(g_script.mNextClipboardViewer, iMsg, wParam, lParam, SMTO_ABORTIFHUNG, 2000, &dwTemp);
		return 0;

	HANDLE_MENU_LOOP // Cases for WM_ENTERMENULOOP and WM_EXITMENULOOP.

	default:
		// The following iMsg can't be in the switch() since it's not constant:
		if (iMsg == WM_TASKBARCREATED && !g_NoTrayIcon) // !g_NoTrayIcon --> the tray icon should be always visible.
		{
			g_script.CreateTrayIcon();
			g_script.UpdateTrayIcon(true);  // Force the icon into the correct pause, suspend, or mIconFrozen state.
			// And now pass this iMsg on to DefWindowProc() in case it does anything with it.
		}

	} // switch()

	return DefWindowProc(hWnd, iMsg, wParam, lParam);
}



bool HandleMenuItem(HWND aHwnd, WORD aMenuItemID, WPARAM aGuiIndex)
// See if an item was selected from the tray menu or main menu.  Note that it is possible
// for one of the standard menu items to be triggered from a GUI menu if the menu or one of
// its submenus was modified with the "menu, MenuName, Standard" command.
// Returns true if the message is fully handled here, false otherwise.
{
    char buf_temp[2048];  // For various uses.

	switch (aMenuItemID)
	{
	case ID_TRAY_OPEN:
		ShowMainWindow();
		return true;
	case ID_TRAY_EDITSCRIPT:
	case ID_FILE_EDITSCRIPT:
		g_script.Edit();
		return true;
	case ID_TRAY_RELOADSCRIPT:
	case ID_FILE_RELOADSCRIPT:
		if (!g_script.Reload(false))
			MsgBox("The script could not be reloaded.");
		return true;
	case ID_TRAY_WINDOWSPY:
	case ID_FILE_WINDOWSPY:
		// ActionExec()'s CreateProcess() is currently done in a way that prefers enclosing double quotes:
		*buf_temp = '"';
		// Try GetAHKInstallDir() first so that compiled scripts running on machines that happen
		// to have AHK installed will still be able to fetch the help file:
		if (GetAHKInstallDir(buf_temp + 1))
			snprintfcat(buf_temp, sizeof(buf_temp), "\\AU3_Spy.exe\"");
		else
			// Mostly this ELSE is here in case AHK isn't installed (i.e. the user just
			// copied the files over without running the installer).  But also:
			// Even if this is the self-contained version (AUTOHOTKEYSC), attempt to launch anyway in
			// case the user has put a copy of WindowSpy in the same dir with the compiled script:
			// ActionExec()'s CreateProcess() is currently done in a way that prefers enclosing double quotes:
			snprintfcat(buf_temp, sizeof(buf_temp), "%sAU3_Spy.exe\"", g_script.mOurEXEDir);
		if (!g_script.ActionExec(buf_temp, "", NULL, false))
			MsgBox(buf_temp, 0, "Could not launch Window Spy:");
		return true;
	case ID_TRAY_HELP:
	case ID_HELP_USERMANUAL:
		// ActionExec()'s CreateProcess() is currently done in a way that prefers enclosing double quotes:
		*buf_temp = '"';
		// Try GetAHKInstallDir() first so that compiled scripts running on machines that happen
		// to have AHK installed will still be able to fetch the help file:
		if (GetAHKInstallDir(buf_temp + 1))
			snprintfcat(buf_temp, sizeof(buf_temp), "\\AutoHotkey.chm\"");
		else
			// Even if this is the self-contained version (AUTOHOTKEYSC), attempt to launch anyway in
			// case the user has put a copy of the help file in the same dir with the compiled script:
			// Also, for this one I saw it report failure once on Win98SE even though the help file did
			// wind up getting launched.  Couldn't repeat it.  So in reponse to that try explicit "hh.exe":
			snprintfcat(buf_temp, sizeof(buf_temp), "%sAutoHotkey.chm\"", g_script.mOurEXEDir);
		if (!g_script.ActionExec("hh.exe", buf_temp, NULL, false))
		{
			// Try it without the hh.exe in case .CHM is associate with some other application
			// in some OSes:
			if (!g_script.ActionExec(buf_temp, "", NULL, false)) // Use "" vs. NULL to specify that there are no params at all.
				MsgBox(buf_temp, 0, "Could not launch help file:");
		}
		return true;
	case ID_TRAY_SUSPEND:
	case ID_FILE_SUSPEND:
		Line::ToggleSuspendState();
		return true;
	case ID_TRAY_PAUSE:
	case ID_FILE_PAUSE:
		if (g_nThreads > 0) // v1.0.37.06: Pausing the "idle thread" (which is not included in the thread count) is an easy means by which to disable all timers.
		{
			if (g.IsPaused)
				--g_nPausedThreads;
			else
				++g_nPausedThreads;
		}
		else // Toggle the pause state of the idle thread.
			g_IdleIsPaused = !g_IdleIsPaused;
		g.IsPaused = !g.IsPaused;
		g_script.UpdateTrayIcon();
		CheckMenuItem(GetMenu(g_hWnd), ID_FILE_PAUSE, g.IsPaused ? MF_CHECKED : MF_UNCHECKED);
		return true;
	case ID_TRAY_EXIT:
	case ID_FILE_EXIT:
		g_script.ExitApp(EXIT_MENU);  // More reliable than PostQuitMessage(), which has been known to fail in rare cases.
		return true; // If there is an OnExit subroutine, the above might not actually exit.
	case ID_VIEW_LINES:
		ShowMainWindow(MAIN_MODE_LINES);
		return true;
	case ID_VIEW_VARIABLES:
		ShowMainWindow(MAIN_MODE_VARS);
		return true;
	case ID_VIEW_HOTKEYS:
		ShowMainWindow(MAIN_MODE_HOTKEYS);
		return true;
	case ID_VIEW_KEYHISTORY:
		ShowMainWindow(MAIN_MODE_KEYHISTORY);
		return true;
	case ID_VIEW_REFRESH:
		ShowMainWindow(MAIN_MODE_REFRESH);
		return true;
	case ID_HELP_WEBSITE:
		if (!g_script.ActionExec("http://www.autohotkey.com", "", NULL, false))
			MsgBox("Could not open URL http://www.autohotkey.com in default browser.");
		return true;
	default:
		// See if this command ID is one of the user's custom menu items.  Due to the possibility
		// that some items have been deleted from the menu, can't rely on comparing
		// aMenuItemID to g_script.mMenuItemCount in any way.  Just look up the ID to make sure
		// there really is a menu item for it:
		if (!g_script.FindMenuItemByID(aMenuItemID)) // Do nothing, let caller try to handle it some other way.
			return false;
		// It seems best to treat the selection of a custom menu item in a way similar
		// to how hotkeys are handled by the hook. See comments near the definition of
		// POST_AHK_USER_MENU for more details.
		POST_AHK_USER_MENU(aHwnd, aMenuItemID, aGuiIndex) // Send the menu's cmd ID and the window index (index is safer than pointer, since pointer might get deleted).
		// Try to maintain a list here of all the ways the script can be uninterruptible
		// at this moment in time, and whether that uninterruptibility should be overridden here:
		// 1) YES: g_MenuIsVisible is true (which in turn means that the script is marked
		//    uninterruptible to prevent timed subroutines from running and possibly
		//    interfering with menu navigation): Seems impossible because apparently 
		//    the WM_RBUTTONDOWN must first be returned from before we're called directly
		//    with the WM_COMMAND message corresponding to the menu item chosen by the user.
		//    In other words, g_MenuIsVisible will be false and the script thus will
		//    not be uninterruptible, at least not solely for that reason.
		// 2) YES: A new hotkey or timed subroutine was just launched and it's still in its
		//    grace period.  In this case, ExecUntil()'s call of PeekMessage() every 10ms
		//    or so will catch the item we just posted.  But it seems okay to interrupt
		//    here directly in most such cases.  InitNewThread(): Newly launched
		//    timed subroutine or hotkey subroutine.
		// 3) YES: Script is engaged in an uninterruptible activity such as SendKeys().  In this
		//    case, since the user has managed to get the tray menu open, it's probably
		//    best to process the menu item with the same priority as if any other menu
		//    item had been selected, interrupting even a critical operation since that's
		//    probably what the user would want.  SLEEP_WITHOUT_INTERRUPTION: SendKeys,
		//    Mouse input, Clipboard open, SetForegroundWindowEx().
		// 4) YES: AutoExecSection(): Since its grace period is only 100ms, doesn't seem to be
		//    a problem.  In any case, the timer would fire and briefly interrupt the menu
		//    subroutine we're trying to launch here even if a menu item were somehow
		//    activated in the first 100ms.
		//
		// IN LIGHT OF THE ABOVE, it seems best not to do the below.  In addition, the msg
		// filtering done by MsgSleep when the script is uninterruptible now excludes the
		// AHK_USER_MENU message (i.e. that message is always retrieved and acted upon,
		// even when the script is uninterruptible):
		//if (!INTERRUPTIBLE)
		//	return true;  // Leave the message buffered until later.
		// Now call the main loop to handle the message we just posted (and any others):
		return true;
	} // switch()
	return false;  // Indicate that the message was NOT handled.
}



ResultType ShowMainWindow(MainWindowModes aMode, bool aRestricted)
// Always returns OK for caller convenience.
{
	// v1.0.30.05: Increased from 32 KB to 64 KB, which is the maximum size of an Edit
	// in Win9x:
	char buf_temp[65534] = "";  // Formerly 32767.
	bool jump_to_bottom = false;  // Set default behavior for edit control.
	static MainWindowModes current_mode = MAIN_MODE_NO_CHANGE;

#ifdef AUTOHOTKEYSC
	// If we were called from a restricted place, such as via the Tray Menu or the Main Menu,
	// don't allow potentially sensitive info such as script lines and variables to be shown.
	// This is done so that scripts can be compiled more securely, making it difficult for anyone
	// to use ListLines to see the author's source code.  Rather than make exceptions for things
	// like KeyHistory, it seems best to forbit all information reporting except in cases where
	// existing info in the main window -- which must have gotten their via an allowed command
	// such as ListLines encountered in the script -- is being refreshed.  This is because in
	// that case, the script author has given de facto permission for that loophole (and it's
	// a pretty small one, not easy to exploit):
	if (aRestricted && !g_AllowMainWindow && (current_mode == MAIN_MODE_NO_CHANGE || aMode != MAIN_MODE_REFRESH))
	{
		SendMessage(g_hWndEdit, WM_SETTEXT, 0, (LPARAM)
			"Script info will not be shown because the \"Menu, Tray, MainWindow\"\r\n"
			"command option was not enabled in the original script.");
		return OK;
	}
#endif

	// If the window is empty, caller wants us to default it to showing the most recently
	// executed script lines:
	if (current_mode == MAIN_MODE_NO_CHANGE && (aMode == MAIN_MODE_NO_CHANGE || aMode == MAIN_MODE_REFRESH))
		aMode = MAIN_MODE_LINES;

	switch (aMode)
	{
	// case MAIN_MODE_NO_CHANGE: do nothing
	case MAIN_MODE_LINES:
		Line::LogToText(buf_temp, sizeof(buf_temp));
		jump_to_bottom = true;
		break;
	case MAIN_MODE_VARS:
		g_script.ListVars(buf_temp, sizeof(buf_temp));
		break;
	case MAIN_MODE_HOTKEYS:
		Hotkey::ListHotkeys(buf_temp, sizeof(buf_temp));
		break;
	case MAIN_MODE_KEYHISTORY:
		g_script.ListKeyHistory(buf_temp, sizeof(buf_temp));
		break;
	case MAIN_MODE_REFRESH:
		// Rather than do a recursive call to self, which might stress the stack if the script is heavily recursed:
		switch (current_mode)
		{
		case MAIN_MODE_LINES:
			Line::LogToText(buf_temp, sizeof(buf_temp));
			jump_to_bottom = true;
			break;
		case MAIN_MODE_VARS:
			g_script.ListVars(buf_temp, sizeof(buf_temp));
			break;
		case MAIN_MODE_HOTKEYS:
			Hotkey::ListHotkeys(buf_temp, sizeof(buf_temp));
			break;
		case MAIN_MODE_KEYHISTORY:
			g_script.ListKeyHistory(buf_temp, sizeof(buf_temp));
			// Special mode for when user refreshes, so that new keys can be seen without having
			// to scroll down again:
			jump_to_bottom = true;
			break;
		}
		break;
	}

	if (aMode != MAIN_MODE_REFRESH && aMode != MAIN_MODE_NO_CHANGE)
		current_mode = aMode;

	// Update the text before displaying the window, since it might be a little less disruptive
	// and might also be quicker if the window is hidden or non-foreground.
	// Unlike SetWindowText(), this method seems to expand tab characters:
	if (aMode != MAIN_MODE_NO_CHANGE)
		SendMessage(g_hWndEdit, WM_SETTEXT, 0, (LPARAM)buf_temp);

	if (!IsWindowVisible(g_hWnd))
	{
		ShowWindow(g_hWnd, SW_SHOW);
		if (IsIconic(g_hWnd)) // This happens whenver the window was last hidden via the minimize button.
			ShowWindow(g_hWnd, SW_RESTORE);
	}
	if (g_hWnd != GetForegroundWindow())
		if (!SetForegroundWindow(g_hWnd))
			SetForegroundWindowEx(g_hWnd);  // Only as a last resort, since it uses AttachThreadInput()

	if (jump_to_bottom)
	{
		SendMessage(g_hWndEdit, EM_LINESCROLL , 0, 999999);
		//SendMessage(g_hWndEdit, EM_SETSEL, -1, -1);
		//SendMessage(g_hWndEdit, EM_SCROLLCARET, 0, 0);
	}
	return OK;
}



ResultType GetAHKInstallDir(char *aBuf)
// Caller must ensure that aBuf is at least MAX_PATH in capacity.
{
	return RegReadString(HKEY_LOCAL_MACHINE, "SOFTWARE\\AutoHotkey", "InstallDir", aBuf, MAX_PATH);
}



//////////////
// InputBox //
//////////////

ResultType InputBox(Var *aOutputVar, char *aTitle, char *aText, bool aHideInput, int aWidth, int aHeight
	, int aX, int aY, double aTimeout, char *aDefault)
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
	if (!*aTitle)
		// If available, the script's filename seems a much better title in case the user has
		// more than one script running:
		aTitle = (g_script.mFileName && *g_script.mFileName) ? g_script.mFileName : NAME_PV;
	// Limit the size of what we were given to prevent unreasonably huge strings from
	// possibly causing a failure in CreateDialog().  This copying method is always done because:
	// Make a copy of all string parameters, using the stack, because they may reside in the deref buffer
	// and other commands (such as those in timed/hotkey subroutines) maybe overwrite the deref buffer.
	// This is not strictly necessary since InputBoxProc() is called immediately and makes instantaneous
	// and one-time use of these strings (not needing them after that), but it feels safer:
	char title[DIALOG_TITLE_SIZE];
	char text[4096];  // Size was increased in light of the fact that dialog can be made larger now.
	char default_string[4096];
	strlcpy(title, aTitle, sizeof(title));
	strlcpy(text, aText, sizeof(text));
	strlcpy(default_string, aDefault, sizeof(default_string));
	g_InputBox[g_nInputBoxes].title = title;
	g_InputBox[g_nInputBoxes].text = text;
	g_InputBox[g_nInputBoxes].default_string = default_string;

	if (aTimeout > 2147483) // This is approximately the max number of seconds that SetTimer() can handle.
		aTimeout = 2147483;
	if (aTimeout < 0) // But it can be equal to zero to indicate no timeout at all.
		aTimeout = 0.1;  // A value that might cue the user that something is wrong.
	g_InputBox[g_nInputBoxes].timeout = (DWORD)(aTimeout * 1000);  // Convert to ms

	// Allow 0 width or height (hides the window):
	g_InputBox[g_nInputBoxes].width = aWidth != INPUTBOX_DEFAULT && aWidth < 0 ? 0 : aWidth;
	g_InputBox[g_nInputBoxes].height = aHeight != INPUTBOX_DEFAULT && aHeight < 0 ? 0 : aHeight;
	g_InputBox[g_nInputBoxes].xpos = aX;  // But seems okay to allow these to be negative, even if absolute coords.
	g_InputBox[g_nInputBoxes].ypos = aY;
	g_InputBox[g_nInputBoxes].output_var = aOutputVar;
	g_InputBox[g_nInputBoxes].password_char = aHideInput ? '*' : '\0';

	// At this point, we know a dialog will be displayed.  See macro's comments for details:
	DIALOG_PREP

	// Specify NULL as the owner window since we want to be able to have the main window in the foreground even
	// if there are InputBox windows.  Update: A GUI window can now be the parent if thread has that setting.
	++g_nInputBoxes;
	int result = (int)DialogBox(g_hInstance, MAKEINTRESOURCE(IDD_INPUTBOX), THREAD_DIALOG_OWNER, InputBoxProc);
	--g_nInputBoxes;

	DIALOG_END

	// See the comments in InputBoxProc() for why ErrorLevel is set here rather than there.
	switch(result)
	{
	case AHK_TIMEOUT:
		// In this case the TimerProc already set the output variable to be what the user
		// entered.  Set ErrorLevel (in this case) even for AutoIt2 scripts since the script
		// is explicitly using a new feature:
		return g_ErrorLevel->Assign("2");
	case IDOK:
	case IDCANCEL:
		// For AutoIt2 (.aut) scripts:
		// If the user pressed the cancel button, InputBoxProc() set the output variable to be blank so
		// that there is a way to detect that the cancel button was pressed.  This is because
		// the InputBox command does not set ErrorLevel for .aut scripts (to maximize backward
		// compatibility), except when the command times out (see the help file for details).
		// For non-AutoIt2 scripts: The output variable is set to whatever the user entered,
		// even if the user pressed the cancel button.  This allows the cancel button to specify
		// that a different operation should be performed on the entered text:
		if (!g_script.mIsAutoIt2)
			return g_ErrorLevel->Assign(result == IDCANCEL ? ERRORLEVEL_ERROR : ERRORLEVEL_NONE);
		// else don't change the value of ErrorLevel at all to retain compatibility.
		break;
	case -1:
		MsgBox("The InputBox window could not be displayed.");
		// No need to set ErrorLevel since this is a runtime error that will kill the current quasi-thread.
		return FAIL;
	case FAIL:
		return FAIL;
	}

	return OK;
}



BOOL CALLBACK InputBoxProc(HWND hWndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
// MSDN:
// Typically, the dialog box procedure should return TRUE if it processed the message,
// and FALSE if it did not. If the dialog box procedure returns FALSE, the dialog
// manager performs the default dialog operation in response to the message.
{
	// See GuiWindowProc() for details about this first part:
	LRESULT msg_reply;
	if (g_MsgMonitorCount && !g.CalledByIsDialogMessageOrDispatch // Count is checked here to avoid function-call overhead.
		&& MsgMonitor(hWndDlg, uMsg, wParam, lParam, NULL, msg_reply))
		return (BOOL)msg_reply; // MsgMonitor has returned "true", indicating that this message should be omitted from further processing.
	g.CalledByIsDialogMessageOrDispatch = false; // v1.0.40.01.

	HWND hControl;

	// Set default array index for g_InputBox[].  Caller has ensured that g_nInputBoxes > 0:
	int target_index = g_nInputBoxes - 1;
	#define CURR_INPUTBOX g_InputBox[target_index]

	switch(uMsg)
	{
	case WM_INITDIALOG:
	{
		// Clipboard may be open if its contents were used to build the text or title
		// of this dialog (e.g. "InputBox, out, %clipboard%").  It's best to do this before
		// anything that might take a relatively long time (e.g. SetForegroundWindowEx()):
		CLOSE_CLIPBOARD_IF_OPEN;

		CURR_INPUTBOX.hwnd = hWndDlg;

		if (CURR_INPUTBOX.password_char)
			SendDlgItemMessage(hWndDlg, IDC_INPUTEDIT, EM_SETPASSWORDCHAR, CURR_INPUTBOX.password_char, 0);

		SetWindowText(hWndDlg, CURR_INPUTBOX.title);
		if (hControl = GetDlgItem(hWndDlg, IDC_INPUTPROMPT))
			SetWindowText(hControl, CURR_INPUTBOX.text);

		// Don't do this check; instead allow the MoveWindow() to occur unconditionally so that
		// the new button positions and such will override those set in the dialog's resource
		// properties:
		//if (CURR_INPUTBOX.width != INPUTBOX_DEFAULT || CURR_INPUTBOX.height != INPUTBOX_DEFAULT
		//	|| CURR_INPUTBOX.xpos != INPUTBOX_DEFAULT || CURR_INPUTBOX.ypos != INPUTBOX_DEFAULT)
		RECT rect;
		GetWindowRect(hWndDlg, &rect);
		int new_width = (CURR_INPUTBOX.width == INPUTBOX_DEFAULT) ? rect.right - rect.left : CURR_INPUTBOX.width;
		int new_height = (CURR_INPUTBOX.height == INPUTBOX_DEFAULT) ? rect.bottom - rect.top : CURR_INPUTBOX.height;

		// If a non-default size was specified, the box will need to be recentered; thus, we can't rely on
		// the dialog's DS_CENTER style in its template.  The exception is when an explicit xpos or ypos is
		// specified, in which case centering is disabled for that dimension.
		int new_xpos, new_ypos;
		if (CURR_INPUTBOX.xpos != INPUTBOX_DEFAULT && CURR_INPUTBOX.ypos != INPUTBOX_DEFAULT)
		{
			new_xpos = CURR_INPUTBOX.xpos;
			new_ypos = CURR_INPUTBOX.ypos;
		}
		else
		{
			POINT pt = CenterWindow(new_width, new_height);
  			if (CURR_INPUTBOX.xpos == INPUTBOX_DEFAULT) // Center horizontally.
				new_xpos = pt.x;
			else
				new_xpos = CURR_INPUTBOX.xpos;
  			if (CURR_INPUTBOX.ypos == INPUTBOX_DEFAULT) // Center vertically.
				new_ypos = pt.y;
			else
				new_ypos = CURR_INPUTBOX.ypos;
		}

		MoveWindow(hWndDlg, new_xpos, new_ypos, new_width, new_height, TRUE);  // Do repaint.
		// This may also needed to make it redraw in some OSes or some conditions:
		GetClientRect(hWndDlg, &rect);  // Not to be confused with GetWindowRect().
		SendMessage(hWndDlg, WM_SIZE, SIZE_RESTORED, rect.right + (rect.bottom<<16));
		
		if (*CURR_INPUTBOX.default_string)
			SetDlgItemText(hWndDlg, IDC_INPUTEDIT, CURR_INPUTBOX.default_string);

		if (hWndDlg != GetForegroundWindow()) // Normally it will be foreground since the template has this property.
			SetForegroundWindowEx(hWndDlg);   // Try to force it to the foreground.

		// Setting the small icon puts it in the upper left corner of the dialog window.
		// Setting the big icon makes the dialog show up correctly in the Alt-Tab menu.
		LPARAM main_icon = (LPARAM)(g_script.mCustomIcon ? g_script.mCustomIcon
			: LoadIcon(g_hInstance, MAKEINTRESOURCE(IDI_MAIN)));
		SendMessage(hWndDlg, WM_SETICON, ICON_SMALL, main_icon);
		SendMessage(hWndDlg, WM_SETICON, ICON_BIG, main_icon);

		// For the timeout, use a timer ID that doesn't conflict with MsgBox's IDs (which are the
		// integers 1 through the max allowed number of msgboxes).  Use +3 vs. +1 for a margin of safety
		// (e.g. in case a few extra MsgBoxes can be created directly by the program and not by
		// the script):
		#define INPUTBOX_TIMER_ID_OFFSET (MAX_MSGBOXES + 3)
		if (CURR_INPUTBOX.timeout)
			SetTimer(hWndDlg, INPUTBOX_TIMER_ID_OFFSET + target_index, CURR_INPUTBOX.timeout, InputBoxTimeout);

		return TRUE; // i.e. let the system set the keyboard focus to the first visible control.
	}

	case WM_SIZE:
	{
		// Adapted from D.Nuttall's InputBox in the AutotIt3 source.

		// don't try moving controls if minimized
		if (wParam == SIZE_MINIMIZED)
			return TRUE;

		int dlg_new_width = LOWORD(lParam);
		int dlg_new_height = HIWORD(lParam);

		int last_ypos = 0, curr_width, curr_height;

		// Changing these might cause weird effects when user resizes the window since the default size and
		// margins is about 5 (as stored in the dialog's resource properties).  UPDATE: That's no longer
		// an issue since the dialog is resized when the dialog is first displayed to make sure everything
		// behaves consistently:
		const int XMargin = 5, YMargin = 5;

		RECT rTmp;

		// start at the bottom - OK button

		HWND hbtOk = GetDlgItem(hWndDlg, IDOK);
		if (hbtOk != NULL)
		{
			// how big is the control?
			GetWindowRect(hbtOk, &rTmp);
			if (rTmp.left > rTmp.right)
				swap(rTmp.left, rTmp.right);
			if (rTmp.top > rTmp.bottom)
				swap(rTmp.top, rTmp.bottom);
			curr_width = rTmp.right - rTmp.left;
			curr_height = rTmp.bottom - rTmp.top;
			last_ypos = dlg_new_height - YMargin - curr_height;
			// where to put the control?
			MoveWindow(hbtOk, dlg_new_width/4+(XMargin-curr_width)/2, last_ypos, curr_width, curr_height, FALSE);
		}

		// Cancel Button
		HWND hbtCancel = GetDlgItem(hWndDlg, IDCANCEL);
		if (hbtCancel != NULL)
		{
			// how big is the control?
			GetWindowRect(hbtCancel, &rTmp);
			if (rTmp.left > rTmp.right)
				swap(rTmp.left, rTmp.right);
			if (rTmp.top > rTmp.bottom)
				swap(rTmp.top, rTmp.bottom);
			curr_width = rTmp.right - rTmp.left;
			curr_height = rTmp.bottom - rTmp.top;
			// where to put the control?
			MoveWindow(hbtCancel, dlg_new_width*3/4-(XMargin+curr_width)/2, last_ypos, curr_width, curr_height, FALSE);
		}

		// Edit Box
		HWND hedText = GetDlgItem(hWndDlg, IDC_INPUTEDIT);
		if (hedText != NULL)
		{
			// how big is the control?
			GetWindowRect(hedText, &rTmp);
			if (rTmp.left > rTmp.right)
				swap(rTmp.left, rTmp.right);
			if (rTmp.top > rTmp.bottom)
				swap(rTmp.top, rTmp.bottom);
			curr_width = rTmp.right - rTmp.left;
			curr_height = rTmp.bottom - rTmp.top;
			last_ypos -= 5 + curr_height;  // Allows space between the buttons and the edit box.
			// where to put the control?
			MoveWindow(hedText, XMargin, last_ypos, dlg_new_width - XMargin*2
				, curr_height, FALSE);
		}

		// Static Box (Prompt)
		HWND hstPrompt = GetDlgItem(hWndDlg, IDC_INPUTPROMPT);
		if (hstPrompt != NULL)
		{
			last_ypos -= 10;  // Allows space between the edit box and the prompt (static text area).
			// where to put the control?
			MoveWindow(hstPrompt, XMargin, YMargin, dlg_new_width - XMargin*2
				, last_ypos, FALSE);
		}
		InvalidateRect(hWndDlg, NULL, TRUE);	// force window to be redrawn
		return TRUE;  // i.e. completely handled here.
	}

	case WM_COMMAND:
		// In this case, don't use (g_nInputBoxes - 1) as the index because it might
		// not correspond to the g_InputBox[] array element that belongs to hWndDlg.
		// This is because more than one input box can be on the screen at the same time.
		// If the user choses to work with on underneath instead of the most recent one,
		// we would be called with an hWndDlg whose index is less than the most recent
		// one's index (g_nInputBoxes - 1).  Instead, search the array for a match.
		// Work backward because the most recent one(s) are more likely to be a match:
		for (; target_index > -1; --target_index)
			if (g_InputBox[target_index].hwnd == hWndDlg)
				break;
		if (target_index < 0)  // Should never happen if things are designed right.
			return FALSE;
		switch (LOWORD(wParam))
		{
		case IDOK:
		case IDCANCEL:
		{
			WORD return_value = LOWORD(wParam);  // Set default, i.e. IDOK or IDCANCEL
			if (   !(hControl = GetDlgItem(hWndDlg, IDC_INPUTEDIT))   )
				return_value = (WORD)FAIL;
			else
			{
				// For AutoIt2 (.aut) script:
				// If the user presses the cancel button, we set the output variable to be blank so
				// that there is a way to detect that the cancel button was pressed.  This is because
				// the InputBox command does not set ErrorLevel for .aut scripts (to maximize backward
				// compatibility), except in the case of a timeout.
				// For non-AutoIt2 scripts: The output variable is set to whatever the user entered,
				// even if the user pressed the cancel button.  This allows the cancel button to specify
				// that a different operation should be performed on the entered text.
				// NOTE: ErrorLevel must not be set here because it's possible that the user has
				// dismissed a dialog that's underneath another, active dialog, or that's currently
				// suspended due to a timed/hotkey subroutine running on top of it.  In other words,
				// it's only safe to set ErrorLevel when the call to DialogProc() returns in InputBox().
				#define SET_OUTPUT_VAR_TO_BLANK (LOWORD(wParam) == IDCANCEL && g_script.mIsAutoIt2)
				#undef INPUTBOX_VAR
				#define INPUTBOX_VAR (CURR_INPUTBOX.output_var)
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
							// There was no text to get or GetWindowText() failed.
							*INPUTBOX_VAR->Contents() = '\0';  // Safe because Assign() gave us a non-constant memory area.
					}
					if (INPUTBOX_VAR->Close() != OK)  // In case it's the clipboard.
						return_value = (WORD)FAIL;
				}
			}
			// Since the user pressed a button to dismiss the dialog:
			// Kill its timer for performance reasons (might degrade perf. a little since OS has
			// to keep track of it as long as it exists).  InputBoxTimeout() already handles things
			// right even if we don't do this:
			if (CURR_INPUTBOX.timeout) // It has a timer.
				KillTimer(hWndDlg, INPUTBOX_TIMER_ID_OFFSET + target_index);
			EndDialog(hWndDlg, return_value);
			return TRUE;
		} // case
		} // Inner switch()
	} // Outer switch()
	// Otherwise, let the dialog handler do its default action:
	return FALSE;
}



VOID CALLBACK InputBoxTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime)
{
	// First check if the window has already been destroyed.  There are quite a few ways this can
	// happen, and in all of them we want to make sure not to do things such as calling EndDialog()
	// again or updating the output variable.  Reasons:
	// 1) The user has already pressed the OK or Cancel button (the timer isn't killed there because
	//    it relies on us doing this check here).  In this case, EndDialog() has already been called
	//    (with the proper result value) and the script's output variable has already been set.
	// 2) Even if we were to kill the timer when the user presses a button to dismiss the dialog,
	//    this IsWindow() check would still be needed here because TimerProc()'s are called via
	//    WM_TIMER messages, some of which might still be in our msg queue even after the timer
	//    has been killed.  In other words, split second timing issues may cause this TimerProc()
	//    to fire even if the timer were killed when the user dismissed the dialog.
	// UPDATE: For performance reasons, the timer is now killed when the user presses a button,
	// so case #1 is obsolete (but kept here for background/insight).
	if (IsWindow(hWnd))
	{
		// This is the element in the array that corresponds to the InputBox for which
		// this function has just been called.
		int target_index = idEvent - INPUTBOX_TIMER_ID_OFFSET;
		// Even though the dialog has timed out, we still want to write anything the user
		// had a chance to enter into the output var.  This is because it's conceivable that
		// someone might want a short timeout just to enter something quick and let the
		// timeout dismiss the dialog for them (i.e. so that they don't have to press enter
		// or a button:
		HWND hControl = GetDlgItem(hWnd, IDC_INPUTEDIT);
		if (hControl)
		{
			#undef INPUTBOX_VAR
			#define INPUTBOX_VAR (g_InputBox[target_index].output_var)
			VarSizeType space_needed = GetWindowTextLength(hControl) + 1;
			// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
			// this call will set up the clipboard for writing:
			if (INPUTBOX_VAR->Assign(NULL, space_needed - 1) == OK)
			{
				// Write to the variable:
				INPUTBOX_VAR->Length() = (VarSizeType)GetWindowText(hControl
					, INPUTBOX_VAR->Contents(), space_needed);
				if (!INPUTBOX_VAR->Length())
					// There was no text to get or GetWindowText() failed.
					*INPUTBOX_VAR->Contents() = '\0';  // Safe because Assign() gave us a non-constant memory area.
				INPUTBOX_VAR->Close();  // In case it's the clipboard.
			}
		}
		EndDialog(hWnd, AHK_TIMEOUT);
	}
	KillTimer(hWnd, idEvent);
}



VOID CALLBACK DerefTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime)
{
	Line::FreeDerefBufIfLarge(); // It will also kill the timer, if appropriate.
}



///////////////////
// Mouse related //
///////////////////

void Line::DoMouseDelay() // Helper function for the mouse functions below.
{
	if (g.MouseDelay > -1)
	{
		if (g.MouseDelay < 11 || (g.MouseDelay < 25 && g_os.IsWin9x()))
			Sleep(g.MouseDelay);
		else
			SLEEP_WITHOUT_INTERRUPTION(g.MouseDelay)
	}
}



ResultType Line::MouseClickDrag(vk_type aVK, int aX1, int aY1, int aX2, int aY2, int aSpeed, bool aMoveRelative)
{
	// Check if one of the coordinates is missing, which can happen in cases where this was called from
	// a source that didn't already validate it:
	if (   (aX1 == COORD_UNSPECIFIED && aY1 != COORD_UNSPECIFIED) || (aX1 != COORD_UNSPECIFIED && aY1 == COORD_UNSPECIFIED)   )
		return FAIL;
	if (   (aX2 == COORD_UNSPECIFIED && aY2 != COORD_UNSPECIFIED) || (aX2 != COORD_UNSPECIFIED && aY2 == COORD_UNSPECIFIED)   )
		return FAIL;

	// If the drag isn't starting at the mouse's current position, move the mouse to the specified position:
	if (aX1 != COORD_UNSPECIFIED && aY1 != COORD_UNSPECIFIED)
		MouseMove(aX1, aY1, aSpeed, aMoveRelative);

	// I asked Jon, "Have you discovered that insta-drags almost always fail?" and he said
	// "Yeah, it was weird, absolute lack of drag... Don't know if it was my config or what."
	// However, testing reveals "insta-drags" work ok, at least on my system, so leaving them enabled.
	// User can easily increase the speed if there's any problem:
	//if (aSpeed < 2)
	//	aSpeed = 2;

	DWORD event_down, event_up, event_data = 0; // Set default.
	switch (aVK)
	{
	case VK_LBUTTON:
		event_down = MOUSEEVENTF_LEFTDOWN;
		event_up = MOUSEEVENTF_LEFTUP;
		break;
	case VK_RBUTTON:
		event_down = MOUSEEVENTF_RIGHTDOWN;
		event_up = MOUSEEVENTF_RIGHTUP;
		break;
	case VK_MBUTTON:
		event_down = MOUSEEVENTF_MIDDLEDOWN;
		event_up = MOUSEEVENTF_MIDDLEUP;
		break;
	case VK_XBUTTON1:
	case VK_XBUTTON2:
		event_down = MOUSEEVENTF_XDOWN;
		event_up = MOUSEEVENTF_XUP;
		event_data = (aVK == VK_XBUTTON1) ? XBUTTON1 : XBUTTON2;
		break;
	}

	// Always sleep a certain minimum amount of time between events to improve reliability,
	// but allow the user to specify a higher time if desired.  Note that for Win9x,
	// a true Sleep() is done because it is much more accurate than the MsgSleep() method,
	// at least on Win98SE when short sleeps are done.  UPDATE: A true sleep is now done
	// unconditionally if the delay period is small.  This fixes a small issue where if
	// LButton is a hotkey that includes "MouseClick left" somewhere in its subroutine,
	// the script's own main window's title bar buttons for min/max/close would not
	// properly respond to left-clicks.  By contrast, the following is no longer an issue
	// due to the dedicated thread in v1.0.39 (or more likely, due to an older change that
	// causes the tray menu to upon upon RButton-up rather than down):
	// RButton is a hotkey that includes "MouseClick right" somewhere in its subroutine,
	// the user would not be able to correctly open the script's own tray menu via
	// right-click (note that this issue affected only the one script itself, not others).

	// Now that the mouse button has been pushed down, move the mouse to perform the drag:
	MouseEvent(event_down, 0, 0, event_data);
	DoMouseDelay();
	MouseMove(aX2, aY2, aSpeed, aMoveRelative);
	DoMouseDelay();
	MouseEvent(event_up, 0, 0, event_data);

	// It seems best to always do this one too in case the script line that caused
	// us to be called here is followed immediately by another script line which
	// is either another mouse click or something that relies upon this mouse drag
	// having been completed:
	DoMouseDelay();
	return OK;
}



ResultType Line::MouseClick(vk_type aVK, int aX, int aY, int aClickCount, int aSpeed, KeyEventTypes aEventType
	, bool aMoveRelative)
{
	// Check if one of the coordinates is missing, which can happen in cases where this was called from
	// a source that didn't already validate it:
	if (   (aX == COORD_UNSPECIFIED && aY != COORD_UNSPECIFIED) || (aX != COORD_UNSPECIFIED && aY == COORD_UNSPECIFIED)   )
		// This was already validated during load so should never happen
		// unless this function was called directly from somewhere else
		// in the app, rather than by a script line:
		return FAIL;

	if (aClickCount < 1)
		// Allow this to simply "do nothing", because it increases flexibility
		// in the case where the number of clicks is a dereferenced script variable
		// that may sometimes (by intent) resolve to zero or negative:
		return OK;

	// If the click isn't occurring at the mouse's current position, move the mouse to the specified position:
	if (aX != COORD_UNSPECIFIED && aY != COORD_UNSPECIFIED)
		MouseMove(aX, aY, aSpeed, aMoveRelative);

	// For wheel movement, if the user activated this command via a hotkey, and that hotkey
	// has a modifier such as CTRL, the user is probably still holding down the CTRL key
	// at this point.  Therefore, there's some merit to the fact that we should release
	// those modifier keys prior to turning the mouse wheel (since some apps disable the
	// wheel or give it different behavior when the CTRL key is down -- for example, MSIE
	// changes the font size when you use the wheel while CTRL is down).  However, if that
	// were to be done, there would be no way to ever hold down the CTRL key explicitly
	// (via Send, {CtrlDown}) unless the hook were installed.  The same argument could probably
	// be made for mouse button clicks: modifier keys can often affect their behavior.  But
	// changing this function to adjust modifiers for all types of events would probably break
	// some existing scripts.  Maybe it can be a script option in the future.  In the meantime,
	// it seems best not to adjust the modifiers for any mouse events and just document that
	// behavior in the MouseClick command.
	if (aVK == VK_WHEEL_UP)
	{
		MouseEvent(MOUSEEVENTF_WHEEL, 0, 0, aClickCount * WHEEL_DELTA);
		return OK;
	}
	else if (aVK == VK_WHEEL_DOWN)
	{
		MouseEvent(MOUSEEVENTF_WHEEL, 0, 0, -(aClickCount * WHEEL_DELTA));
		return OK;
	}
	// Otherwise:

	// Although not thread-safe, the following static vars seem okay because:
	// 1) This function is currently only called by the main thread.
	// 2) Even if that isn't true, the serialized nature of simulated mouse clicks makes it likely that
	//    the statics will produce the correct behavior anyway.
	// 3) Even if that isn't true, the consequences of incorrect behavior seem minimal in this case.
	static vk_type sWorkaroundVK = 0;
	static LRESULT sWorkaroundHitTest;  // Not initialized because the above will be the sole signal of whether the workaround is in progress.
	DWORD event_down, event_up, event_data = 0; // Set default.

	switch (aVK)
	{
	case VK_LBUTTON:
	case VK_RBUTTON:
		if (aEventType == KEYDOWN || (aEventType == KEYUP && sWorkaroundVK)) // i.e. this is a down-only event or up-only event.
		{
			// v1.0.40.01: The following section corrects misbehavior caused by a thread sending
			// simulated mouse clicks to one of its own windows.  A script consisting only of the
			// following two lines can reproduce this issue:
			// F1::LButton
			// F2::RButton
			// The problems came about from the following sequence of events:
			// 1) Script simulates a left-click-down in the title bar's close, minimize, or maximize button.
			// 2) WM_NCRBUTTONDOWN is sent to the window's window proc, which then passes it on to
			//    DefWindowProc or DefDlgProc, which then apparently enters a loop in which no messages
			//    (or a very limited subset) are pumped.
			// 3) Thus, if the user presses a hotkey while the thread is in this state, that hotkey is
			//    queued/buffered until DefWindowProc/DefDlgProc exits its loop.
			// 4) But the buffered hotkey is the very thing that's supposed to exit the loop via sending a
			//    simulated left-click-up event.
			// 5) Thus, a deadlock occurs.
			// 6) A similar situation arises when a right-click-down is sent to the title bar or sys-menu-icon.
			//
			// The following workaround operates by suppressing qualified click-down events until the
			// corresponding click-up occurs, at which time the click-up is transformed into a down+up if
			// the click-up is still in the same position as the down. It seems preferable to fix this here
			// rather than changing each window proc. to always respond to click-down rather vs. click-up
			// because that would make all of the script's windows behave in a non-standard way, possibly
			// producing side-effects and defeating other programs' attempts to interact with them.
			// (Thanks to Shimanov for this solution.)
			//
			// Remaining known limitations:
			// 1) Title bar buttons are not visibly in a pressed down state when a simulated click-down is sent
			//    to them.
			// 2) A window that should not be activated, such as AlwaysOnTop+Disabled, is activated anyway
			//    by SetForegroundWindowEx().  Not yet fixed due to its rarity and minimal consequences.
			// 3) A related problem for which no solution has been discovered (and perhaps it's too obscure
			//    an issue to justify any added code size): If a remapping such as "F1::LButton" is in effect,
			//    pressing and releasing F1 while the cursor is over a script window's title bar will cause the
			//    window to move slightly the next time the mouse is moved.
			// 4) Clicking one of the script's window's title bar with a key/button that has been remapped to
			//    become the left mouse button sometimes causes the button to get stuck down from the window's
			//    point of view.  The reasons are related to those in #1 above.  In both #1 and #2, the workaround
			//    is not at fault because it's not in effect then.  Instead, the issue is that DefWindowProc enters
			//    a non-msg-pumping loop while it waits for the user to drag-move the window.  If instead the user
			//    releases the button without dragging, the loop exits on its own after a 500ms delay or so.
			// 5) Obscure behavior cause by keyboard's auto-repeat feature: Use a key that's been remapped to
			//    become the left mouse button to click and hold the minimize button of one of the script's windows.
			//    Drag to the left.  The window starts moving.  This is caused by the fact that the down-click is
			//    suppressed, thus the remap's hotkey subroutine thinks the mouse button is down, thus its
			//    auto-repeat suppression doesn't work and it sends another click.
			POINT point;
			GetCursorPos(&point); // Assuming success seems harmless.
			// Despite what MSDN says, WindowFromPoint() appears to fetch a non-NULL value even when the
			// mouse is hovering over a disabled control (at least on XP).
			HWND child_under_cursor, parent_under_cursor;
			if (   (child_under_cursor = WindowFromPoint(point))
				&& (parent_under_cursor = GetNonChildParent(child_under_cursor)) // WM_NCHITTEST below probably requires parent vs. child.
				&& GetWindowThreadProcessId(parent_under_cursor, NULL) == g_MainThreadID   ) // It's one of our thread's windows.
			{
				LRESULT hit_test = SendMessage(parent_under_cursor, WM_NCHITTEST, 0, MAKELPARAM(point.x, point.y));
				if (   aVK == VK_LBUTTON && (hit_test == HTCLOSE || hit_test == HTMAXBUTTON // Title bar buttons: Close, Maximize.
						|| hit_test == HTMINBUTTON || hit_test == HTHELP) // Title bar buttons: Minimize, Help.
					|| aVK == VK_RBUTTON && (hit_test == HTCAPTION || hit_test == HTSYSMENU)   )
				{
					if (aEventType == KEYDOWN)
					{
						sWorkaroundVK = aVK;
						sWorkaroundHitTest = hit_test;
						SetForegroundWindowEx(parent_under_cursor); // Try to reproduce customary behavior.
						// For simplicity, aClickCount>1 is ignored and DoMouseDelay() is not done.
						return OK;
					}
					else // KEYUP
					{
						if (sWorkaroundHitTest == hit_test) // To weed out cases where user clicked down on a button then released somewhere other than the button.
							aEventType = KEYDOWNANDUP; // Translate this click-up into down+up to make up for the fact that the down was previously suppressed.
						//else let the click-up occur in case it does something or user wants it.
					}
				}
			} // Work-around for sending mouse clicks to one of our thread's own windows.
		}
		// sWorkaroundVK is reset later below.

		// Since above didn't return, the work-around isn't in effect and normal click(s) will be sent:
		if (aVK == VK_LBUTTON)
		{
			event_down = MOUSEEVENTF_LEFTDOWN;
			event_up = MOUSEEVENTF_LEFTUP;
		}
		else // aVK == VK_RBUTTON
		{
			event_down = MOUSEEVENTF_RIGHTDOWN;
			event_up = MOUSEEVENTF_RIGHTUP;
		}
		break;
	case VK_MBUTTON:
		event_down = MOUSEEVENTF_MIDDLEDOWN;
		event_up = MOUSEEVENTF_MIDDLEUP;
		break;
	case VK_XBUTTON1:
	case VK_XBUTTON2:
		event_down = MOUSEEVENTF_XDOWN;
		event_up = MOUSEEVENTF_XUP;
		event_data = (aVK == VK_XBUTTON1) ? XBUTTON1 : XBUTTON2;
		break;
	} // switch()

	for (int i = 0; i < aClickCount; ++i)
	{
		// The below calls to MouseEvent() do not specify coordinates because such are only
		// needed if we were to include MOUSEEVENTF_MOVE in the dwFlags parameter, which
		// we don't since we've already moved the mouse (above) if that was needed.
		if (aEventType != KEYUP) // It's either KEYDOWN or KEYDOWNANDUP.
		{
			MouseEvent(event_down, 0, 0, event_data);
			// It seems best to always Sleep a certain minimum time between events
			// because the click-down event may cause the target app to do something which
			// changes the context or nature of the click-up event.  AutoIt3 has also been
			// revised to do this. v1.0.40.02: Avoid doing the Sleep between the down and up
			// events when the workaround is in effect because any MouseDelay greater than 10
			// would cause DoMouseDelay() to pump messages, which would defeat the workaround:
			if (!sWorkaroundVK)
				DoMouseDelay();
		}
		if (aEventType != KEYDOWN) // It's either KEYUP or KEYDOWNANDUP.
		{
			MouseEvent(event_up, 0, 0, event_data);
			// It seems best to always do this one too in case the script line that caused
			// us to be called here is followed immediately by another script line which
			// is either another mouse click or something that relies upon the mouse click
			// having been completed:
			DoMouseDelay();
		}
	} // for()

	sWorkaroundVK = 0; // Reset this indicator in all cases except those for which above already returned.
	return OK;
}



ResultType Line::MouseGetPos(bool aSimpleMode)
// Returns OK or FAIL.
{
	// Caller should already have ensured that at least one of these will be non-NULL.
	// The only time this isn't true is for dynamically-built variable names.  In that
	// case, we don't worry about it if it's NULL, since the user will already have been
	// warned:
	Var *output_var_x = ResolveVarOfArg(0);  // Ok if NULL.
	Var *output_var_y = ResolveVarOfArg(1);  // Ok if NULL.
	Var *output_var_parent = ResolveVarOfArg(2);  // Ok if NULL.
	Var *output_var_child = ResolveVarOfArg(3);  // Ok if NULL.

	POINT point;
	GetCursorPos(&point);  // Realistically, can't fail?

	RECT rect = {0};  // ensure it's initialized for later calculations.
	if (!(g.CoordMode & COORD_MODE_MOUSE)) // Using relative vs. absolute coordinates.
	{
		HWND fore_win = GetForegroundWindow();
		GetWindowRect(fore_win, &rect);  // If this call fails, above default values will be used.
	}

	if (output_var_x) // else the user didn't want the X coordinate, just the Y.
		if (!output_var_x->Assign(point.x - rect.left))
			return FAIL;
	if (output_var_y) // else the user didn't want the Y coordinate, just the X.
		if (!output_var_y->Assign(point.y - rect.top))
			return FAIL;

	if (!output_var_parent && !output_var_child)
		return OK;

	// This is the child window.  Despite what MSDN says, WindowFromPoint() appears to fetch
	// a non-NULL value even when the mouse is hovering over a disabled control (at least on XP).
	HWND child_under_cursor = WindowFromPoint(point);
	if (!child_under_cursor)
	{
		if (output_var_parent)
			output_var_parent->Assign();
		if (output_var_child)
			output_var_child->Assign();
		return OK;
	}

	HWND parent_under_cursor = GetNonChildParent(child_under_cursor);  // Find the first ancestor that isn't a child.
	if (output_var_parent)
	{
		// Testing reveals that an invisible parent window never obscures another window beneath it as seen by
		// WindowFromPoint().  In other words, the below never happens, so there's no point in having it as a
		// documented feature:
		//if (!g.DetectHiddenWindows && !IsWindowVisible(parent_under_cursor))
		//	return output_var_parent->Assign();
		if (!output_var_parent->AssignHWND(parent_under_cursor))
			return FAIL;
	}

	if (!output_var_child)
		return OK;

	// Doing it this way overcomes the limitations of WindowFromPoint() and ChildWindowFromPoint()
	// and also better matches the control that Window Spy would think is under the cursor:
	if (!aSimpleMode)
	{
		point_and_hwnd_type pah = {0};
		pah.pt = point;
		EnumChildWindows(parent_under_cursor, EnumChildFindPoint, (LPARAM)&pah); // Find topmost control containing point.
		if (pah.hwnd_found)
			child_under_cursor = pah.hwnd_found;
	}
	//else as of v1.0.25.10, leave child_under_cursor set the the value retrieved earlier from WindowFromPoint().
	// This allows MDI child windows to be reported correctly; i.e. that the window on top of the others
	// is reported rather than the one at the top of the z-order (the z-order of MDI child windows,
	// although probably constant, is not useful for determine which one is one top of the others).

	if (parent_under_cursor == child_under_cursor) // if there's no control per se, make it blank.
		return output_var_child->Assign();

	class_and_hwnd_type cah;
	cah.hwnd = child_under_cursor;  // This is the specific control we need to find the sequence number of.
	char class_name[WINDOW_CLASS_SIZE];
	cah.class_name = class_name;
	if (!GetClassName(cah.hwnd, class_name, sizeof(class_name) - 5))  // -5 to allow room for sequence number.
		return output_var_child->Assign();
	cah.class_count = 0;  // Init for the below.
	cah.is_found = false; // Same.
	EnumChildWindows(parent_under_cursor, EnumChildFindSeqNum, (LPARAM)&cah); // Find this control's seq. number.
	if (!cah.is_found)
		return output_var_child->Assign();  
	// Append the class sequence number onto the class name and set the output param to be that value:
	snprintfcat(class_name, sizeof(class_name), "%d", cah.class_count);
	return output_var_child->Assign(class_name);
}



BOOL CALLBACK EnumChildFindPoint(HWND aWnd, LPARAM lParam)
// This is called by more than one caller.  It finds the most appropriate child window that contains
// the specified point (the point should be in screen coordinates).
{
	point_and_hwnd_type &pah = *((point_and_hwnd_type *)lParam);  // For performance and convenience.
	if (!IsWindowVisible(aWnd)) // Omit hidden controls, like Window Spy does.
		return TRUE;
	RECT rect;
	if (!GetWindowRect(aWnd, &rect))
		return TRUE;
	// The given point must be inside aWnd's bounds.  Then, if there is no hwnd found yet or if aWnd
	// is entirely contained within the previously found hwnd, update to a "better" found window like
	// Window Spy.  This overcomes the limitations of WindowFromPoint() and ChildWindowFromPoint():
	if (pah.pt.x >= rect.left && pah.pt.x <= rect.right && pah.pt.y >= rect.top && pah.pt.y <= rect.bottom)
	{
		// If the window's center is closer to the given point, break the tie and have it take
		// precedence.  This solves the problem where a particular control from a set of overlapping
		// controls is chosen arbitrarily (based on Z-order) rather than based on something the
		// user would find more intuitive (the control whose center is closest to the mouse):
		double center_x = rect.left + (double)(rect.right - rect.left) / 2;
		double center_y = rect.top + (double)(rect.bottom - rect.top) / 2;
		// Taking the absolute value first is not necessary because it seems that qmathHypot()
		// takes the square root of the sum of the squares, which handles negatives correctly:
		double distance = qmathHypot(pah.pt.x - center_x, pah.pt.y - center_y);
		//double distance = qmathSqrt(qmathPow(pah.pt.x - center_x, 2) + qmathPow(pah.pt.y - center_y, 2));
		bool update_it = !pah.hwnd_found;
		if (!update_it)
		{
			// If the new window's rect is entirely contained within the old found-window's rect, update
			// even if the distance is greater.  Conversely, if the new window's rect entirely encloses
			// the old window's rect, do not update even if the distance is less:
			if (rect.left >= pah.rect_found.left && rect.right <= pah.rect_found.right
				&& rect.top >= pah.rect_found.top && rect.bottom <= pah.rect_found.bottom)
				update_it = true; // New is entirely enclosed by old: update to the New.
			else if (   distance < pah.distance &&
				(pah.rect_found.left < rect.left || pah.rect_found.right > rect.right
					|| pah.rect_found.top < rect.top || pah.rect_found.bottom > rect.bottom)   )
				update_it = true; // New doesn't entirely enclose old and new's center is closer to the point.
		}
		if (update_it)
		{
			pah.hwnd_found = aWnd;
			pah.rect_found = rect; // And at least one caller uses this returned rect.
			pah.distance = distance;
		}
	}
	return TRUE; // Continue enumeration all the way through.
}



///////////////////////////////
// Related to other commands //
///////////////////////////////

ResultType Line::FormatTime(char *aYYYYMMDD, char *aFormat)
// The compressed code size of this function is about 1 KB (2 KB uncompressed), which compares
// favorably to using setlocale()+strftime(), which together are about 8 KB of compressed code
// (setlocale() seems to be needed to put the user's or system's locale into effect for strftime()).
// setlocale() weighs in at about 6.5 KB compressed (14 KB uncompressed).
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;

	#define FT_MAX_INPUT_CHARS 2000  // In preparation for future use of TCHARs, since GetDateFormat() uses char-count not size.
	// Input/format length is restricted since it must be translated and expanded into a new format
	// string that uses single quotes around non-alphanumeric characters such as punctuation:
	if (strlen(aFormat) > FT_MAX_INPUT_CHARS)
		return output_var->Assign();

	// Worst case expansion: .d.d.d.d. (9 chars) --> '.'d'.'d'.'d'.'d'.' (19 chars)
	// Buffer below is sized to a little more than twice as big as the largest allowed format,
	// which avoids having to constantly check for buffer overflow while translating aFormat
	// into format_buf:
	#define FT_MAX_OUTPUT_CHARS (2*FT_MAX_INPUT_CHARS + 10)
	char format_buf[FT_MAX_OUTPUT_CHARS + 1];
	char output_buf[FT_MAX_OUTPUT_CHARS + 1]; // The size of this is somewhat arbitrary, but buffer overflow is checked so it's safe.

	char yyyymmdd[256] = ""; // Large enough to hold date/time and any options that follow it (note that D and T options can appear multiple times).

	SYSTEMTIME st;
	char *options = NULL;

	if (!*aYYYYMMDD) // Use current local time by default.
		GetLocalTime(&st);
	else
	{
		strlcpy(yyyymmdd, omit_leading_whitespace(aYYYYMMDD), sizeof(yyyymmdd)); // Make a modifiable copy.
		if (*yyyymmdd < '0' || *yyyymmdd > '9') // First character isn't a digit, therefore...
		{
			// ... options are present without date (since yyyymmdd [if present] must come before options).
			options = yyyymmdd;
			GetLocalTime(&st);  // Use current local time by default.
		}
		else // Since the string starts with a digit, rules say it must be a YYYYMMDD string, possibly followed by options.
		{
			// Find first space or tab because YYYYMMDD portion might contain only the leading part of date/timestamp.
			if (options = StrChrAny(yyyymmdd, " \t")) // Find space or tab.
			{
				*options = '\0'; // Terminate yyyymmdd at the end of the YYYYMMDDHH24MISS string.
				options = omit_leading_whitespace(++options); // Point options to the right place (can be empty string).
			}
			//else leave options set to NULL to indicate that there are none.

			// Pass "false" for validatation so that times can still be reported even if the year
			// is prior to 1601.  If the time and/or date is invalid, GetTimeFormat() and GetDateFormat()
			// will refuse to produce anything, which is documented behavior:
			YYYYMMDDToSystemTime(yyyymmdd, st, false);
		}
	}
	
	// Set defaults.  Some can be overridden by options (if there are any options).
	LCID lcid = LOCALE_USER_DEFAULT;
	DWORD date_flags = 0, time_flags = 0;
	bool date_flags_specified = false, time_flags_specified = false, reverse_date_time = false;
	#define FT_FORMAT_NONE 0
	#define FT_FORMAT_TIME 1
	#define FT_FORMAT_DATE 2
	int format_type1 = FT_FORMAT_NONE;
	char *format2_marker = NULL; // Will hold the location of the first char of the second format (if present).
	bool do_null_format2 = false;  // Will be changed to true if a default date *and* time should be used.

	if (options) // Parse options.
	{
		char *option_end, orig_char;
		for (char *next_option = options; *next_option; next_option = omit_leading_whitespace(option_end))
		{
			// Find the end of this option item:
			if (   !(option_end = StrChrAny(next_option, " \t"))   )  // Space or tab.
				option_end = next_option + strlen(next_option); // Set to position of zero terminator instead.

			// Permanently terminate in between options to help eliminate ambiguity for words contained
			// inside other words, and increase confidence in decimal and hexadecimal conversion.
			orig_char = *option_end;
			*option_end = '\0';

			++next_option;
			switch (toupper(next_option[-1]))
			{
			case 'D':
				date_flags_specified = true;
				date_flags |= ATOU(next_option); // ATOU() for unsigned.
				break;
			case 'T':
				time_flags_specified = true;
				time_flags |= ATOU(next_option); // ATOU() for unsigned.
				break;
			case 'R':
				reverse_date_time = true;
				break;
			case 'L':
				lcid = !stricmp(next_option, "Sys") ? LOCALE_SYSTEM_DEFAULT : (LCID)ATOU(next_option);
				break;
			// If not one of the above, such as zero terminator or a number, just ignore it.
			}

			*option_end = orig_char; // Undo the temporary termination so that loop's omit_leading() will work.
		} // for() each item in option list
	} // Parse options.

	if (!*aFormat)
	{
		aFormat = NULL; // Tell GetDateFormat() and GetTimeFormat() to use default for the specified locale.
		if (!date_flags_specified) // No preference was given, so use long (which seems generally more useful).
			date_flags |= DATE_LONGDATE;
		if (!time_flags_specified)
			time_flags |= TIME_NOSECONDS;  // Seems more desirable/typical to default to no seconds.
		// Put the time first by default, though this is debatable (Metapad does it and I like it).
		format_type1 = reverse_date_time ? FT_FORMAT_DATE : FT_FORMAT_TIME;
		do_null_format2 = true;
	}
	else // aFormat is non-blank.
	{
		// Omit whitespace only for consideration of special keywords.  Whitespace is later kept for
		// a normal format string such as %A_Space%MM/dd/yy:
		char *candidate = omit_leading_whitespace(aFormat);
		if (!stricmp(candidate, "YWeek"))
		{
			GetISOWeekNumber(output_buf, st.wYear, GetYDay(st.wMonth, st.wDay, IS_LEAP_YEAR(st.wYear)), st.wDayOfWeek);
			return output_var->Assign(output_buf);
		}
		if (!stricmp(candidate, "YDay") || !stricmp(candidate, "YDay0"))
		{
			int yday = GetYDay(st.wMonth, st.wDay, IS_LEAP_YEAR(st.wYear));
			if (!stricmp(candidate, "YDay"))
				return output_var->Assign(yday); // Assign with no leading zeroes, also will be in hex format if that format is in effect.
			// Otherwise:
			sprintf(output_buf, "%03d", yday);
			return output_var->Assign(output_buf);
		}
		if (!stricmp(candidate, "WDay"))
			return output_var->Assign(st.wDayOfWeek + 1);  // Convert to 1-based for compatibility with A_WDay.

		// Since above didn't return, check for those that require a call to GetTimeFormat/GetDateFormat
		// further below:
		if (!stricmp(candidate, "ShortDate"))
		{
			aFormat = NULL;
			date_flags |= DATE_SHORTDATE;
			date_flags &= ~(DATE_LONGDATE | DATE_YEARMONTH); // If present, these would prevent it from working.
		}
		else if (!stricmp(candidate, "LongDate"))
		{
			aFormat = NULL;
			date_flags |= DATE_LONGDATE;
			date_flags &= ~(DATE_SHORTDATE | DATE_YEARMONTH); // If present, these would prevent it from working.
		}
		else if (!stricmp(candidate, "YearMonth"))
		{
			aFormat = NULL;
			date_flags |= DATE_YEARMONTH;
			date_flags &= ~(DATE_SHORTDATE | DATE_LONGDATE); // If present, these would prevent it from working.
		}
		else if (!stricmp(candidate, "Time"))
		{
			format_type1 = FT_FORMAT_TIME;
			aFormat = NULL;
			if (!time_flags_specified)
				time_flags |= TIME_NOSECONDS;  // Seems more desirable/typical to default to no seconds.
		}
		else // Assume normal format string.
		{
			char *cp = aFormat, *dp = format_buf;   // Initialize source and destination pointers.
			bool inside_their_quotes = false; // Whether we are inside a single-quoted string in the source.
			bool inside_our_quotes = false;   // Whether we are inside a single-quoted string of our own making in dest.
			for (; *cp; ++cp) // Transcribe aFormat into format_buf and also check for which comes first: date or time.
			{
				if (*cp == '\'') // Note that '''' (four consecutive quotes) is a single literal quote, which this logic handles okay.
				{
					if (inside_our_quotes)
					{
						// Upon encountering their quotes while we're still in ours, merge theirs with ours and
						// remark it as theirs.  This is done to avoid having two back-to-back quoted sections,
						// which would result in an unwanted literal single quote.  Example:
						// 'Some string'':' (the two quotes in the middle would be seen as a literal quote).
						inside_our_quotes = false;
						inside_their_quotes = true;
						continue;
					}
					if (inside_their_quotes)
					{
						// If next char needs to be quoted, don't close out this quote section because that
						// would introduce two consecutive quotes, which would be interpreted as a single
						// literal quote if its enclosed by two outer single quotes.  Instead convert this
						// quoted section over to "ours":
						if (cp[1] && !IsCharAlphaNumeric(cp[1]) && cp[1] != '\'') // Also consider single quotes to be theirs due to this example: dddd:''''y
							inside_our_quotes = true;
							// And don't do "*dp++ = *cp"
						else // there's no next-char or it's alpha-numeric, so it doesn't need to be inside quotes.
							*dp++ = *cp; // Close out their quoted section.
					}
					else // They're starting a new quoted section, so just transcribe this single quote as-is.
						*dp++ = *cp;
					inside_their_quotes = !inside_their_quotes; // Must be done after the above.
					continue;
				}
				// Otherwise, it's not a single quote.
				if (inside_their_quotes) // *cp is inside a single-quoted string, so it can be part of format/picture
					*dp++ = *cp; // Transcribe as-is.
				else
				{
					if (IsCharAlphaNumeric(*cp))
					{
						if (inside_our_quotes)
						{
							*dp++ = '\''; // Close out the previous quoted section, since this char should not be a part of it.
							inside_our_quotes = false;
						}
						if (strchr("dMyg", *cp)) // A format unique to Date is present.
						{
							if (!format_type1)
								format_type1 = FT_FORMAT_DATE;
							else if (format_type1 == FT_FORMAT_TIME && !format2_marker) // type2 should only be set if different than type1.
							{
								*dp++ = '\0';  // Terminate the first section and (below) indicate that there's a second.
								format2_marker = dp;  // Point it to the location in format_buf where the split should occur.
							}
						}
						else if (strchr("hHmst", *cp)) // A format unique to Time is present.
						{
							if (!format_type1)
								format_type1 = FT_FORMAT_TIME;
							else if (format_type1 == FT_FORMAT_DATE && !format2_marker) // type2 should only be set if different than type1.
							{
								*dp++ = '\0';  // Terminate the first section and (below) indicate that there's a second.
								format2_marker = dp;  // Point it to the location in format_buf where the split should occur.
							}
						}
						// For consistency, transcribe all AlphaNumeric chars not inside single quotes as-is
						// (numbers are transcribed in case they are ever used as part of pic/format).
						*dp++ = *cp;
					}
					else // Not alphanumeric, so enclose this and any other non-alphanumeric characters in single quotes.
					{
						if (!inside_our_quotes)
						{
							*dp++ = '\''; // Create a new quoted section of our own, since this char should be inside quotes to be understood.
							inside_our_quotes = true;
						}
						*dp++ = *cp;  // Add this character between the quotes, since it's of the right "type".
					}
				}
			} // for()
			if (inside_our_quotes)
				*dp++ = '\'';  // Close out our quotes.
			*dp = '\0'; // Final terminator.
			aFormat = format_buf; // Point it to the freshly translated format string, for use below.
		} // aFormat contains normal format/pic string.
	} // aFormat isn't blank.

	// If there are no date or time formats present, still do the transcription so that
	// any quoted strings and whatnot are resolved.  This increases runtime flexibility.
	// The below is also relied upon by "LongDate" and "ShortDate" above:
	if (!format_type1)
		format_type1 = FT_FORMAT_DATE;

	// MSDN: Time: "The function checks each of the time values to determine that it is within the
	// appropriate range of values. If any of the time values are outside the correct range, the
	// function fails, and sets the last-error to ERROR_INVALID_PARAMETER. 
	// Dates: "...year, month, day, and day of week. If the day of the week is incorrect, the
	// function uses the correct value, and returns no error. If any of the other date values
	// are outside the correct range, the function fails, and sets the last-error to ERROR_INVALID_PARAMETER.

	if (format_type1 == FT_FORMAT_DATE) // DATE comes first.
	{
		if (!GetDateFormat(lcid, date_flags, &st, aFormat, output_buf, FT_MAX_OUTPUT_CHARS))
			*output_buf = '\0';  // Ensure it's still the empty string, then try to continue to get the second half (if there is one).
	}
	else // TIME comes first.
		if (!GetTimeFormat(lcid, time_flags, &st, aFormat, output_buf, FT_MAX_OUTPUT_CHARS))
			*output_buf = '\0';  // Ensure it's still the empty string, then try to continue to get the second half (if there is one).

	if (format2_marker || do_null_format2) // There is also a second format present.
	{
		size_t output_buf_length = strlen(output_buf);
		char *output_buf_marker = output_buf + output_buf_length;
		char *format2;
		if (do_null_format2)
		{
			format2 = NULL;
			*output_buf_marker++ = ' '; // Provide a space between time and date.
			++output_buf_length;
		}
		else
			format2 = format2_marker;

		int buf_remaining_size = (int)(FT_MAX_OUTPUT_CHARS - output_buf_length);
		int result;

		if (format_type1 == FT_FORMAT_DATE) // DATE came first, so this one is TIME.
			result = GetTimeFormat(lcid, time_flags, &st, format2, output_buf_marker, buf_remaining_size);
		else
			result = GetDateFormat(lcid, date_flags, &st, format2, output_buf_marker, buf_remaining_size);
		if (!result)
			output_buf[output_buf_length] = '\0'; // Ensure the first part is still terminated and just return that rather than nothing.
	}

	return output_var->Assign(output_buf);
}



ResultType Line::PerformAssign()
// Returns OK or FAIL.  Caller has ensured that none of this line's derefs is a function-call.
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	output_var = output_var->ResolveAlias(); // Resolve alias now to detect "source_is_being_appended_to_target" and perhaps other things.
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
	bool source_is_being_appended_to_target = false; // v1.0.25
	if (output_var->Type() != VAR_CLIPBOARD && mArgc > 1)
	{
		// It has a second arg, which in this case is the value to be assigned to the var.
		// Examine any derefs that the second arg has to see if output_var is mentioned.
		// Also, calls to script functions aren't possible within these derefs because
		// our caller has ensured there are no expressions, and thus no function calls,
		// inside this line.
		for (DerefType *deref = mArg[1].deref; deref && deref->marker; ++deref)
		{
			if (deref->is_function) // Silent failure, for rare cases such ACT_ASSIGNEXPR calling us due to something like Clipboard:=SavedClipboard + fn(x)
				return FAIL;
			if (source_is_being_appended_to_target)
			{
				// Check if target is mentioned more than once in source, e.g. Var = %Var%Some Text%Var%
				// would be disqualified for the "fast append" method because %Var% occurs more than once.
				if (deref->var->ResolveAlias() == output_var) // deref->is_function was checked above just in case.
				{
					source_is_being_appended_to_target = false;
					break;
				}
			}
			else
			{
				if (deref->var->ResolveAlias() == output_var) // deref->is_function was checked above just in case.
				{
					target_is_involved_in_source = true;
					// The below disqualifies both of the following cases from the simple-append mode:
					// Var = %OtherVar%%Var%   ; Var is not the first item as required.
					// Var = LiteralText%Var%  ; Same.
					if (deref->marker == mArg[1].text)
						source_is_being_appended_to_target = true;
						// And continue the loop to ensure that Var is not referenced more than once,
						// e.g. Var = %Var%%Var% would be disqualified.
					else
						break;
				}
			}
		}
	}

	VarSizeType space_needed;
	bool assign_clipboardall = false, assign_binary_var = false;
	Var *source_var;

	if (mArgc > 1)
	{
		source_var = NULL;
		if (mArg[1].type == ARG_TYPE_INPUT_VAR) // This can only happen when called from ACT_ASSIGNEXPR, in which case it's safe to look at sArgVar.
			source_var = sArgVar[1]; // Should be non-null, but okay even if it isn't.
		else // Can't use sArgVar here because ExecUntil() never calls ExpandArgs() for ACT_ASSIGN.
			if (ArgHasDeref(2)) // There is at least one deref in Arg #2.
				// For simplicity, we don't check that it's the only deref, nor whether it has any literal text
				// around it, since those things aren't supported anyway.
				source_var = mArg[1].deref[0].var; // Caller has ensured none of this line's derefs is a function-call.
		if (source_var)
		{
			assign_clipboardall = source_var->Type() == VAR_CLIPBOARDALL;
			assign_binary_var = source_var->IsBinaryClip();
		}
	}

	LPVOID binary_contents;
	HGLOBAL hglobal;
	LPVOID hglobal_locked;
	UINT format;
	SIZE_T size;
	VarSizeType added_size;

	if (assign_clipboardall)
	{
		// The caller is performing the special mode "Var = %ClipboardAll%".
		if (output_var->Type() == VAR_CLIPBOARD) // Seems pointless (nor is the below equipped to handle it), so make this have no effect.
			return OK;
		if (!g_clip.Open())
			return LineError(CANT_OPEN_CLIPBOARD_READ);
		// Calculate the size needed:
		// EnumClipboardFormats() retrieves all formats, including synthesized formats that don't
		// actually exist on the clipboard but are instead constructed on demand.  Unfortunately,
		// there doesn't appear to be any way to reliably determine which formats are real and
		// which are synthesized (if there were such a way, a large memory savings could be
		// realized by omitting the synthesized formats from the saved version). One thing that
		// is certain is that the "real" format(s) come first and the synthesized ones afterward.
		// However, that's not quite enough because although it is recommended that apps store
		// the primary/preferred format first, the OS does not enforce this.  For example, testing
		// shows that the apps do not have to store CF_UNICODETEXT prior to storing CF_TEXT,
		// in which case the clipboard might have inaccurate CF_TEXT as the first element and
		// more accurate/complete (non-synthesized) CF_UNICODETEXT stored as the next.
		// In spite of the above, the below seems likely to be accurate 99% or more of the time,
		// which seems worth it given the large savings of memory that are achieved, especially
		// for large quantities of text or large images. Confidence is further raised by the
		// fact that MSDN says there's no advantage/reason for an app to place multiple formats
		// onto the clipboard if those formats are available through synthesis.
		// And since CF_TEXT always(?) yields synthetic CF_OEMTEXT and CF_UNICODETEXT, and
		// probably (but less certainly) vice versa: if CF_TEXT is listed first, it might certainly
		// mean that the other two do not need to be stored.  There is some slight doubt about this
		// in a situation where an app explicitly put CF_TEXT onto the clipboard and then followed
		// it with CF_UNICODETEXT that isn't synthesized, nor does it match what would have been
		// synthesized. However, that seems extremely unlikely (it would be much more likely for
		// an app to store CF_UNICODETEXT *first* followed by custom/non-synthesized CF_TEXT, but
		// even that might be unheard of in practice).  So for now -- since there is no documentation
		// to be found about this anywhere -- it seems best to omit some of the most common
		// synthesized formats:
		// CF_TEXT is the first of three text formats to appear: Omit CF_OEMTEXT and CF_UNICODETEXT.
		//    (but not vice versa since those are less certain to be synthesized)
		//    (above avoids using four times the amount of memory that would otherwise be required)
		//    UPDATE: Only the first text format is included now, since MSDN says there is no
		//    advantage/reason to having multiple non-synthesized text formats on the clipboard.
		// CF_DIB: Always omit this if CF_DIBV5 is available (which must be present on Win2k+, at least
		// as a synthesized format, whenever CF_DIB is present?) This policy seems likely to avoid
		// the issue where CF_DIB occurs first yet CF_DIBV5 that comes later is *not* synthesized,
		// perhaps simply because the app stored DIB prior to DIBV5 by mistake (though there is
		// nothing mandatory, so maybe it's not really a mistake). Note: CF_DIBV5 supports alpha
		// channel / transparency, and perhaps other things, and it is likely that when synthesized,
		// no information of the original CF_DIB is lost. Thus, when CF_DIBV5 is placed back onto
		// the clipboard, any app that needs CF_DIB will have it synthesized back to the original
		// data (hopefully). It's debatable whether to do it that way or store whichever comes first
		// under the theory that an app would never store both formats on the clipboard since MSDN
		// says: "If the system provides an automatic type conversion for a particular clipboard format,
		// there is no advantage to placing the conversion format(s) on the clipboard."
		bool format_is_text;
		UINT dib_format_to_omit = 0, meta_format_to_omit = 0, text_format_to_include = 0;
		// Start space_needed off at 4 to allow room for guaranteed final termination of output_var's contents.
		// The termination must be of the same size as format because a single-byte terminator would
		// be read in as a format of 0x00?????? where ?????? is an access violation beyond the buffer.
		for (space_needed = sizeof(format), format = 0; format = EnumClipboardFormats(format);)
		{
			// No point in calling GetLastError() since it would never be executed because the loop's
			// condition breaks on zero return value.
			format_is_text = (format == CF_TEXT || format == CF_OEMTEXT || format == CF_UNICODETEXT);
			if ((format_is_text && text_format_to_include) // The first text format has already been found and included, so exclude all other text formats.
				|| format == dib_format_to_omit) // ... or this format was marked excluded by a prior iteration.
				continue;
			// GetClipboardData() causes Task Manager to report a (sometimes large) increase in
			// memory utilization for the script, which is odd since it persists even after the
			// clipboard is closed.  However, when something new is put onto the clipboard by the
			// the user or any app, that memory seems to get freed automatically.  Also, 
			// GetClipboardData(49356) fails in MS Visual C++ when the copied text is greater than
			// about 200 KB (but GetLastError() returns ERROR_SUCCESS).  When pasting large sections
			// of colorized text into MS Word, it can't get the colorized text either (just the plain
			// text). Because of this example, it seems likely it can fail in other places or under
			// other circumstances, perhaps by design of the app. Therefore, be tolerant of failures
			// because partially saving the clipboard seems much better than aborting the operation.
			if (hglobal = GetClipboardData(format))
			{
				space_needed += (VarSizeType)(sizeof(format) + sizeof(size) + GlobalSize(hglobal)); // The total amount of storage space required for this item.
				if (format_is_text) // If this is true, then text_format_to_include must be 0 since above didn't "continue".
					text_format_to_include = format;
				if (!dib_format_to_omit)
				{
					if (format == CF_DIB)
						dib_format_to_omit = CF_DIBV5;
					else if (format == CF_DIBV5)
						dib_format_to_omit = CF_DIB;
				}
				if (!meta_format_to_omit) // Checked for the same reasons as dib_format_to_omit.
				{
					if (format == CF_ENHMETAFILE)
						meta_format_to_omit = CF_METAFILEPICT;
					else if (format == CF_METAFILEPICT)
						meta_format_to_omit = CF_ENHMETAFILE;
				}
			}
			//else omit this format from consideration.
		}

		if (space_needed == sizeof(format)) // This works because even a single empty format requires space beyond sizeof(format) for storing its format+size.
		{
			g_clip.Close();
			return output_var->Assign(); // Nothing on the clipboard, so just make output_var blank.
		}

		// Resize the output variable, if needed:
		if (!output_var->Assign(NULL, space_needed - 1))
		{
			g_clip.Close();
			return FAIL; // Above should have already reported the error.
		}

		// Retrieve and store all the clipboard formats.  Because failures of GetClipboardData() are now
		// tolerated, it seems safest to recalculate the actual size (actual_space_needed) of the data
		// in case it varies from that found in the estimation phase.  This is especially necessary in
		// case GlobalLock() ever fails, since that isn't even attempted during the estimation phase.
		// Otherwise, the variable's mLength member would be set to something too high (the estimate),
		// which might cause problems elsewhere.
		binary_contents = output_var->Contents();
		VarSizeType capacity = output_var->Capacity(); // Note that this is the granted capacity, which might be a little larger than requested.
		VarSizeType actual_space_used;
		for (actual_space_used = sizeof(format), format = 0; format = EnumClipboardFormats(format);)
		{
			// No point in calling GetLastError() since it would never be executed because the loop's
			// condition breaks on zero return value.
			if ((format == CF_TEXT || format == CF_OEMTEXT || format == CF_UNICODETEXT) && format != text_format_to_include
				|| format == dib_format_to_omit || format == meta_format_to_omit)
				continue;
			// Although the GlobalSize() documentation implies that a valid HGLOBAL should not be zero in
			// size, it does happen, at least in MS Word and for CF_BITMAP.  Therefore, in order to save
			// the clipboard as accurately as possible, also save formats whose size is zero.  Note that
			// GlobalLock() fails to work on hglobals of size zero, so don't do it for them.
			if ((hglobal = GetClipboardData(format)) // This and the next line rely on short-circuit boolean order.
				&& (!(size = GlobalSize(hglobal)) || (hglobal_locked = GlobalLock(hglobal)))) // Size of zero or lock succeeded: Include this format.
			{
				// Any changes made to how things are stored here should also be made to the size-estimation
				// phase so that space_needed matches what is done here:
				added_size = (VarSizeType)(sizeof(format) + sizeof(size) + size);
				actual_space_used += added_size;
				if (actual_space_used > capacity) // Tolerate incorrect estimate by omitting formats that won't fit.
					actual_space_used -= added_size;
				else
				{
					*(UINT *)binary_contents = format;
					binary_contents = (char *)binary_contents + sizeof(format);
					*(SIZE_T *)binary_contents = size;
					binary_contents = (char *)binary_contents + sizeof(size);
					if (size)
					{
						memcpy(binary_contents, hglobal_locked, size);
						binary_contents = (char *)binary_contents + size;
					}
					//else hglobal_locked is not valid, so don't reference it or unlock it.
				}
				if (size)
					GlobalUnlock(hglobal); // hglobal not hglobal_locked.
			}
		}
		g_clip.Close();
		*(UINT *)binary_contents = 0; // Final termination (must be UINT, see above).
		output_var->Length() = actual_space_used - 1; // Omit the final zero-byte from the length in case any other routines assume that exactly one zero exists at the end of var's length.
		return output_var->Close(true); // Pass "true" to make it binary-clipboard type.
	}

	if (assign_binary_var)
	{
		// Caller wants a variable with binary contents assigned (copied) to another variable (usually VAR_CLIPBOARD).
		binary_contents = source_var->Contents();
		VarSizeType source_length = source_var->Length();
		if (output_var->Type() != VAR_CLIPBOARD) // Copy a binary variable to another variable that isn't the clipboard.
		{
			if (!output_var->Assign(NULL, source_length))
				return FAIL; // Above should have already reported the error.
			memcpy(output_var->Contents(), binary_contents, source_length + 1);  // Add 1 not sizeof(format).
			output_var->Length() = source_length;
			return output_var->Close(true); // Pass "true" to make it binary-clipboard type.
		}

		// Since above didn't return, a variable containing binary clipboard data is being copied back onto
		// the clipboard.
		if (!g_clip.Open())
			return LineError(CANT_OPEN_CLIPBOARD_WRITE);
		EmptyClipboard(); // Failure is not checked for since it's probably impossible under these conditions.

		// In case the variable contents are incomplete or corrupted (such as having been read in from a
		// bad file with FileRead), prevent reading beyond the end of the variable:
		LPVOID next, binary_contents_max = (char *)binary_contents + source_length + 1; // The last acessible byte, which should be the last byte of the (UINT)0 terminator.

		while ((next = (char *)binary_contents + sizeof(format)) <= binary_contents_max
			&& (format = *(UINT *)binary_contents)) // Get the format.  Relies on short-circuit boolean order.
		{
			binary_contents = next;
			if ((next = (char *)binary_contents + sizeof(size)) > binary_contents_max)
				break;
			size = *(UINT *)binary_contents; // Get the size of this format's data.
			binary_contents = next;
			if ((next = (char *)binary_contents + size) > binary_contents_max)
				break;
	        if (   !(hglobal = GlobalAlloc(GMEM_MOVEABLE, size))   ) // size==0 is okay.
			{
				g_clip.Close();
				return LineError(ERR_OUTOFMEM, FAIL); // Short msg since so rare.
			}
			if (size) // i.e. Don't try to lock memory of size zero.  It won't work and it's not needed.
			{
				if (   !(hglobal_locked = GlobalLock(hglobal))   )
				{
					GlobalFree(hglobal);
					g_clip.Close();
					return LineError("GlobalLock", FAIL); // Short msg since so rare.
				}
				memcpy(hglobal_locked, binary_contents, size);
				GlobalUnlock(hglobal);
				binary_contents = next;
			}
			//else hglobal is just an empty format, but store it for completeness/accuracy (e.g. CF_BITMAP).
			SetClipboardData(format, hglobal); // The system now owns hglobal.
		}
		return g_clip.Close();
	}

	// Otherwise (since above didn't return):
	// Note: It might be possible to improve performance in the case where
	// the target variable is large enough to accommodate the new source data
	// by moving memory around inside it.  For example, Var1 = xxxxxVar1
	// could be handled by moving the memory in Var1 to make room to insert
	// the literal string.  In addition to being quicker than the ExpandArgs()
	// method, this approach would avoid the possibility of needing to expand the
	// deref buffer just to handle the operation.  However, if that is ever done,
	// be sure to check that output_var is mentioned only once in the list of derefs.
	// For example, something like this would probably be much easier to
	// implement by using ExpandArgs(): Var1 = xxxx %Var1% %Var2% %Var1% xxxx.
	// So the main thing to be possibly later improved here is the case where
	// output_var is mentioned only once in the deref list (which as of v1.0.25,
	// has been partially done via the concatenation improvement, e.g. Var = %Var%Text).
	Var *arg_var[MAX_ARGS];
	if (target_is_involved_in_source && !source_is_being_appended_to_target)
	{
		if (ExpandArgs() != OK)
			return FAIL;
		// ARG2 now contains the dereferenced (literal) contents of the text we want to assign.
		space_needed = (VarSizeType)strlen(ARG2) + 1;  // +1 for the zero terminator.
	}
	else
	{
		space_needed = GetExpandedArgSize(false, arg_var); // There's at most one arg to expand in this case.
		if (space_needed == VARSIZE_ERROR)
			return FAIL;  // It will have already displayed the error.
	}

	// Now above has ensured that space_needed is at least 1 (it should not be zero because even
	// the empty string uses up 1 char for its zero terminator).  The below relies upon this fact.

	if (space_needed <= 1) // Variable is being assigned the empty string (or a deref that resolves to it).
		return output_var->Assign("");  // If the var is of large capacity, this will also free its memory.

	if (source_is_being_appended_to_target)
	{
		if (space_needed > output_var->Capacity())
		{
			// Since expanding the size of output_var while preserving its existing contents would
			// likely be a slow operation, revert to the normal method rather than the fast-append
			// mode.  Expand the args then continue on normally to the below.
			if (ExpandArgs(space_needed, arg_var) != OK) // In this case, both params were previously calculated by GetExpandedArgSize().
				return FAIL;
		}
		else // there's enough capacity in output_var to accept the text to be appended.
			target_is_involved_in_source = false;  // Tell the below not to consider expanding the args.
	}

	if (target_is_involved_in_source)
		// It was already dereferenced above, so use ARG2, which points to the
		// derefed contents of ARG2 (i.e. the data to be assigned).
		return output_var->Assign(ARG2, VARSIZE_MAX, g.AutoTrim); // Pass VARSIZE_MAX to have it recalculate, since space_needed might be a conservative estimate larger than the actual length+1.

	// Otherwise:
	// If we're here, output_var->Type() must be clipboard or normal because otherwise
	// the validation during load would have prevented the script from loading:

	// First set everything up for the operation.  If output_var is the clipboard, this
	// will prepare the clipboard for writing.  Update: If source is being appended
	// to target using the simple method, we know output_var isn't the clipboard because the
	// logic at the top of this function ensures that.
	if (!source_is_being_appended_to_target && output_var->Assign(NULL, space_needed - 1) != OK)
		return FAIL;
	// Expand Arg2 directly into the var.  Also set the length explicitly
	// in case actual size written was off from the esimated size, perhaps
	// due to a failure or size discrepancy between the deref size-estimate
	// and the actual deref itself.  Note: If output_var is the clipboard,
	// it's probably okay if the below actually writes less than the size of
	// the mem that has already been allocated for the new clipboard contents
	// That might happen due to a failure or size discrepancy between the
	// deref size-estimate and the actual deref itself:
	char *contents = output_var->Contents();
	// This knows not to copy the first var-ref onto itself (for when source_is_being_appended_to_target is true).
	// In addition, to reach this point, arg_var[0]'s value will already have been determined (possibly NULL)
	// by GetExpandedArgSize():
	char *one_beyond_contents_end = ExpandArg(contents, 1, arg_var[0]);
	if (!one_beyond_contents_end)
		return FAIL;  // ExpandArg() will have already displayed the error.
	// Set the length explicitly rather than using space_needed because GetExpandedArgSize()
	// sometimes returns a larger size than is actually needed (e.g. for ScriptGetCursor()):
	size_t length = one_beyond_contents_end - contents - 1;
	// v1.0.25: Passing the precalculated length to trim() greatly improves performance,
	// especially for concat loops involving things like Var = %Var%String:
	output_var->Length() = (VarSizeType)(g.AutoTrim ? trim(contents, length) : length);
	return output_var->Close();  // i.e. Consider this function to be always successful unless this fails.
}



ResultType Line::StringSplit(char *aArrayName, char *aInputString, char *aDelimiterList, char *aOmitList)
{
	// Make it longer than Max so that FindOrAddVar() will be able to spot and report var names
	// that are too long, either because the base-name is too long, or the name becomes too long
	// as a result of appending the array index number:
	char var_name[MAX_VAR_NAME_LENGTH + 20];
	snprintf(var_name, sizeof(var_name), "%s0", aArrayName);
	// ALWAYS_PREFER_LOCAL below allows any existing local variable that matches array0's name
	// (e.g. Array0) to be given preference over creating a new global variable if the function's
	// mode is to assume globals:
	Var *array0 = g_script.FindOrAddVar(var_name, 0, ALWAYS_PREFER_LOCAL);
	if (!array0)
		return FAIL;  // It will have already displayed the error.
	int always_use = array0->IsLocal() ? ALWAYS_USE_LOCAL : ALWAYS_USE_GLOBAL;

	if (!*aInputString) // The input variable is blank, thus there will be zero elements.
		return array0->Assign("0");  // Store the count in the 0th element.

	DWORD next_element_number;
	Var *next_element;

	if (*aDelimiterList) // The user provided a list of delimiters, so process the input variable normally.
	{
		char *contents_of_next_element, *delimiter, *new_starting_pos;
		size_t element_length;
		for (contents_of_next_element = aInputString, next_element_number = 1; ; ++next_element_number)
		{
			snprintf(var_name, sizeof(var_name), "%s%u", aArrayName, next_element_number);

			// To help performance (in case the linked list of variables is huge), tell it where
			// to start the search.  Use element #0 rather than the preceding element because,
			// for example, Array19 is alphabetially less than Array2, so we can't rely on the
			// numerical ordering:
			if (   !(next_element = g_script.FindOrAddVar(var_name, 0, always_use))   )
				return FAIL;  // It will have already displayed the error.

			if (delimiter = StrChrAny(contents_of_next_element, aDelimiterList)) // A delimiter was found.
			{
				element_length = delimiter - contents_of_next_element;
				if (*aOmitList && element_length > 0)
				{
					contents_of_next_element = omit_leading_any(contents_of_next_element, aOmitList, element_length);
					element_length = delimiter - contents_of_next_element; // Update in case above changed it.
					if (element_length)
						element_length = omit_trailing_any(contents_of_next_element, aOmitList, delimiter - 1);
				}
				// If there are no chars to the left of the delim, or if they were all in the list of omitted
				// chars, the variable will be assigned the empty string:
				if (!next_element->Assign(contents_of_next_element, (VarSizeType)element_length))
					return FAIL;
				contents_of_next_element = delimiter + 1;  // Omit the delimiter since it's never included in contents.
			}
			else // the entire length of contents_of_next_element is what will be stored
			{
				element_length = strlen(contents_of_next_element);
				if (*aOmitList && element_length > 0)
				{
					new_starting_pos = omit_leading_any(contents_of_next_element, aOmitList, element_length);
					element_length -= (new_starting_pos - contents_of_next_element); // Update in case above changed it.
					contents_of_next_element = new_starting_pos;
					if (element_length)
						// If this is true, the string must contain at least one char that isn't in the list
						// of omitted chars, otherwise omit_leading_any() would have already omitted them:
						element_length = omit_trailing_any(contents_of_next_element, aOmitList
							, contents_of_next_element + element_length - 1);
				}
				// If there are no chars to the left of the delim, or if they were all in the list of omitted
				// chars, the variable will be assigned the empty string:
				if (!next_element->Assign(contents_of_next_element, (VarSizeType)element_length))
					return FAIL;
				// This is the only way out of the loop other than critical errors:
				return array0->Assign(next_element_number); // Store the count of how many items were stored in the array.
			}
		}
	}

	// Otherwise aDelimiterList is empty, so store each char of aInputString in its own array element.
	char *cp, *dp;
	for (cp = aInputString, next_element_number = 1; *cp; ++cp)
	{
		for (dp = aOmitList; *dp; ++dp)
			if (*cp == *dp) // This char is a member of the omitted list, thus it is not included in the output array.
				break;
		if (*dp) // Omitted.
			continue;
		snprintf(var_name, sizeof(var_name), "%s%u", aArrayName, next_element_number);
		if (   !(next_element = g_script.FindOrAddVar(var_name, 0, always_use))   )
			return FAIL;  // It will have already displayed the error.
		if (!next_element->Assign(cp, 1))
			return FAIL;
		++next_element_number; // Only increment this if above didn't "continue".
	}
	return array0->Assign(next_element_number - 1); // Store the count of how many items were stored in the array.
}



ResultType Line::SplitPath(char *aFileSpec)
{
	Var *output_var_name = ResolveVarOfArg(1);  // i.e. Param #2. Ok if NULL.
	Var *output_var_dir = ResolveVarOfArg(2);  // Ok if NULL.
	Var *output_var_ext = ResolveVarOfArg(3);  // Ok if NULL.
	Var *output_var_name_no_ext = ResolveVarOfArg(4);  // Ok if NULL.
	Var *output_var_drive = ResolveVarOfArg(5);  // Ok if NULL.

	// For URLs, "drive" is defined as the server name, e.g. http://somedomain.com
	char *name = "", *name_delimiter = NULL, *drive_end = NULL; // Set defaults to improve maintainability.
	char *drive = omit_leading_whitespace(aFileSpec); // i.e. whitespace is considered for everything except the drive letter or server name, so that a pathless filename can have leading whitespace.
	char *colon_double_slash = strstr(aFileSpec, "://");

	if (colon_double_slash) // This is a URL such as ftp://... or http://...
	{
		if (   !(drive_end = strchr(colon_double_slash + 3, '/'))   )
		{
			if (   !(drive_end = strchr(colon_double_slash + 3, '\\'))   ) // Try backslash so that things like file://C:\Folder\File.txt are supported.
				drive_end = colon_double_slash + strlen(colon_double_slash); // Set it to the position of the zero terminator instead.
				// And because there is no filename, leave name and name_delimiter set to their defaults.
			//else there is a backslash, e.g. file://C:\Folder\File.txt, so treat that backslash as the end of the drive name.
		}
		name_delimiter = drive_end; // Set default, to be possibly overridden below.
		// Above has set drive_end to one of the following:
		// 1) The slash that occurs to the right of the doubleslash in a URL.
		// 2) The backslash that occurs to the right of the doubleslash in a URL.
		// 3) The zero terminator if there is no slash or backslash to the right of the doubleslash.
		if (*drive_end) // A slash or backslash exists to the right of the server name.
		{
			if (*(drive_end + 1))
			{
				// Find the rightmost slash.  At this stage, this is known to find the correct slash.
				// In the case of a file at the root of a domain such as http://domain.com/root_file.htm,
				// the directory consists of only the domain name, e.g. http://domain.com.  This is because
				// the directory always the "drive letter" by design, since that is more often what the
				// caller wants.  A script can use StringReplace to remove the drive/server portion from
				// the directory, if desired.
				name_delimiter = strrchr(aFileSpec, '/');
				if (name_delimiter == colon_double_slash + 2) // To reach this point, it must have a backslash, something like file://c:\folder\file.txt
					name_delimiter = strrchr(aFileSpec, '\\'); // Will always be found.
				name = name_delimiter + 1; // This will be the empty string for something like http://domain.com/dir/
			}
			//else something like http://domain.com/, so leave name and name_delimiter set to their defaults.
		}
		//else something like http://domain.com, so leave name and name_delimiter set to their defaults.
	}
	else // It's not a URL, just a file specification such as c:\my folder\my file.txt, or \\server01\folder\file.txt
	{
		// Differences between _splitpath() and the method used here:
		// _splitpath() doesn't include drive in output_var_dir, it includes a trailing
		// backslash, it includes the . in the extension, it considers ":" to be a filename.
		// _splitpath(pathname, drive, dir, file, ext);
		//char sdrive[16], sdir[MAX_PATH], sname[MAX_PATH], sext[MAX_PATH];
		//_splitpath(aFileSpec, sdrive, sdir, sname, sext);
		//if (output_var_name_no_ext)
		//	output_var_name_no_ext->Assign(sname);
		//strcat(sname, sext);
		//if (output_var_name)
		//	output_var_name->Assign(sname);
		//if (output_var_dir)
		//	output_var_dir->Assign(sdir);
		//if (output_var_ext)
		//	output_var_ext->Assign(sext);
		//if (output_var_drive)
		//	output_var_drive->Assign(sdrive);
		//return OK;

		// Don't use _splitpath() since it supposedly doesn't handle UNC paths correctly,
		// and anyway we need more info than it provides.  Also note that it is possible
		// for a file to begin with space(s) or a dot (if created programmatically), so
		// don't trim or omit leading space unless it's known to be an absolute path.

		// Note that "C:Some File.txt" is a valid filename in some contexts, which the below
		// tries to take into account.  However, there will be no way for this command to
		// return a path that differentiates between "C:Some File.txt" and "C:\Some File.txt"
		// since the first backslash is not included with the returned path, even if it's
		// the root directory (i.e. "C:" is returned in both cases).  The "C:Filename"
		// convention is pretty rare, and anyway this trait can be detected via something like
		// IfInString, Filespec, :, IfNotInString, Filespec, :\, MsgBox Drive with no absolute path.

		// UNCs are detected with this approach so that double sets of backslashes -- which sometimes
		// occur by accident in "built filespecs" and are tolerated by the OS -- are not falsely
		// detected as UNCs.
		if (*drive == '\\' && *(drive + 1) == '\\') // Relies on short-circuit evaluation order.
		{
			if (   !(drive_end = strchr(drive + 2, '\\'))   )
				drive_end = drive + strlen(drive); // Set it to the position of the zero terminator instead.
		}
		else if (*(drive + 1) == ':') // It's an absolute path.
			// Assign letter and colon for consistency with server naming convention above.
			// i.e. so that server name and drive can be used without having to worry about
			// whether it needs a colon added or not.
			drive_end = drive + 2;
		else
			// It's debatable, but it seems best to return a blank drive if a aFileSpec is a relative path.
			// rather than trying to use GetFullPathName() on a potentially non-existent file/dir.
			// _splitpath() doesn't fetch the drive letter of relative paths either.  This also reports
			// a blank drive for something like file://C:\My Folder\My File.txt, which seems too rarely
			// to justify a special mode.
			drive_end = drive = ""; // This is necessary to allow Assign() to work correctly later below, since it interprets a length of zero as "use string's entire length".

		if (   !(name_delimiter = strrchr(aFileSpec, '\\'))   ) // No backslash.
			if (   !(name_delimiter = strrchr(aFileSpec, ':'))   ) // No colon.
				name_delimiter = NULL; // Indicate that there is no directory.

		name = name_delimiter ? name_delimiter + 1 : aFileSpec; // If no delimiter, name is the entire string.
	}

	// The above has now set the following variables:
	// name: As an empty string or the actual name of the file, including extension.
	// name_delimiter: As NULL if there is no directory, otherwise, the end of the directory's name.
	// drive: As the start of the drive/server name, e.g. C:, \\Workstation01, http://domain.com, etc.
	// drive_end: As the position after the drive's last character, either a zero terminator, slash, or backslash.

	if (output_var_name && !output_var_name->Assign(name))
		return FAIL;

	if (output_var_dir)
	{
		if (!name_delimiter)
			output_var_dir->Assign(); // Shouldn't fail.
		else if (*name_delimiter == '\\' || *name_delimiter == '/')
		{
			if (!output_var_dir->Assign(aFileSpec, (VarSizeType)(name_delimiter - aFileSpec)))
				return FAIL;
		}
		else // *name_delimiter == ':', e.g. "C:Some File.txt".  If aFileSpec starts with just ":",
			 // the dir returned here will also start with just ":" since that's rare & illegal anyway.
			if (!output_var_dir->Assign(aFileSpec, (VarSizeType)(name_delimiter - aFileSpec + 1)))
				return FAIL;
	}

	char *ext_dot = strrchr(name, '.');
	if (output_var_ext)
	{
		// Note that the OS doesn't allow filenames to end in a period.
		if (!ext_dot)
			output_var_ext->Assign();
		else
			if (!output_var_ext->Assign(ext_dot + 1)) // Can be empty string if filename ends in just a dot.
				return FAIL;
	}

	if (output_var_name_no_ext && !output_var_name_no_ext->Assign(name, (VarSizeType)(ext_dot ? ext_dot - name : strlen(name))))
		return FAIL;

	if (output_var_drive && !output_var_drive->Assign(drive, (VarSizeType)(drive_end - drive)))
		return FAIL;

	return OK;
}



int SortWithOptions(const void *a1, const void *a2)
// Decided to just have one sort function since there are so many permutations.  The performance
// will be a little bit worse, but it seems simpler to implement and maintain.
// This function's input parameters are pointers to the elements of the array.  Snce those elements
// are themselves pointers, the input parameters are therefore pointers to pointers (handles).
{
	char *sort_item1 = *(char **)a1;
	char *sort_item2 = *(char **)a2;
	if (g_SortColumnOffset > 0)
	{
		// Adjust each string (even for numerical sort) to be the right column position,
		// or the position of its zero terminator if the column offset goes beyond its length:
		size_t length = strlen(sort_item1);
		sort_item1 += (size_t)g_SortColumnOffset > length ? length : g_SortColumnOffset;
		length = strlen(sort_item2);
		sort_item2 += (size_t)g_SortColumnOffset > length ? length : g_SortColumnOffset;
	}
	if (g_SortNumeric) // Takes precedence over g_SortCaseSensitive
	{
		// For now, assume both are numbers.  If one of them isn't, it will be sorted as a zero.
		// Thus, all non-numeric items should wind up in a sequential, unsorted group.
		// Resolve only once since parts of the ATOF() macro are inline:
		double item1_minus_2 = ATOF(sort_item1) - ATOF(sort_item2);
		if (!item1_minus_2) // Exactly equal.
			return 0;
		// Otherwise, it's either greater or less than zero:
		int result = (item1_minus_2 > 0.0) ? 1 : -1;
		return g_SortReverse ? -result : result;
	}
	// Otherwise, it's a non-numeric sort.
	if (g_SortReverse)
		return g_SortCaseSensitive ? strcmp(sort_item2, sort_item1) : stricmp(sort_item2, sort_item1);
	else
		return g_SortCaseSensitive ? strcmp(sort_item1, sort_item2) : stricmp(sort_item1, sort_item2);
}



int SortByNakedFilename(const void *a1, const void *a2)
// See comments in prior function for details.
{
	char *sort_item1 = *(char **)a1;
	char *sort_item2 = *(char **)a2;
	char *cp;
	if (cp = strrchr(sort_item1, '\\'))  // Assign
		sort_item1 = cp + 1;
	if (cp = strrchr(sort_item2, '\\'))  // Assign
		sort_item2 = cp + 1;
	if (g_SortReverse)
		return g_SortCaseSensitive ? strcmp(sort_item2, sort_item1) : stricmp(sort_item2, sort_item1);
	else
		return g_SortCaseSensitive ? strcmp(sort_item1, sort_item2) : stricmp(sort_item1, sort_item2);
}



struct sort_rand_type
{
	char *cp; // This must be the first member of the struct, otherwise the array trickery in PerformSort will fail.
	// Below must be the same size in bytes as the above, which is why it's maintained as a union with
	// a char* rather than a plain int (though currently they would be the same size anyway).
	union
	{
		char *noname;
		int rand;
	};
};

int SortRandom(const void *a1, const void *a2)
// See comments in prior functions for details.
{
	return ((sort_rand_type *)a1)->rand - ((sort_rand_type *)a2)->rand;
}



ResultType Line::PerformSort(char *aContents, char *aOptions)
// Caller must ensure that aContents is modifiable (ArgMustBeDereferenced() currently ensures this).
// It seems best to treat ACT_SORT's var to be an input vs. output var because if
// it's an environment variable or the clipboard, the input variable handler will
// automatically resolve it to be ARG1 (i.e. put its contents into the deref buf).
// This is especially necessary if the clipboard contains files, in which case
// output_var->Get(), not Contents(),  must be used to resolve the filenames into text.
// And on average, using the deref buffer for this operation will not be wasteful in
// terms of expanding it unnecessarily, because usually the contents will have been
// already (or will soon be) in the deref buffer as a result of commands before
// or after the Sort command in the script.
{
	if (!*aContents) // Variable is empty, nothing to sort.
		return OK;

	Var *output_var = ResolveVarOfArg(0); // The input var (ARG1) is also the output var in this case.
	if (!output_var)
		return FAIL;

	// Do nothing for reserved variables, since most of them are read-only and besides, none
	// of them (realistically) should ever need sorting:
	if (VAR_IS_RESERVED(output_var))
		return OK;

	// Resolve options.  Set defaults first:
	char delimiter = '\n';
	g_SortCaseSensitive = false;
	g_SortNumeric = false;
	g_SortReverse = false;
	g_SortColumnOffset = 0;
	bool allow_last_item_to_be_blank = false, terminate_last_item_with_delimiter = false;
	bool sort_by_naked_filename = false;
	bool sort_random = false;
	bool omit_dupes = false;
	char *cp;

	for (cp = aOptions; *cp; ++cp)
	{
		switch(toupper(*cp))
		{
		case 'C':
			g_SortCaseSensitive = true;
			break;
		case 'D':
			if (!cp[1]) // Avoids out-of-bounds when the loop's own ++cp is done.
				break;
			++cp;
			if (*cp)
				delimiter = *cp;
			break;
		case 'N':
			g_SortNumeric = true;
			break;
		case 'P':
			// Use atoi() vs. ATOI() to avoid interpreting something like 0x01C as hex
			// when in fact the C was meant to be an option letter:
			g_SortColumnOffset = atoi(cp + 1);
			if (g_SortColumnOffset < 1)
				g_SortColumnOffset = 1;
			--g_SortColumnOffset;  // Convert to zero-based.
			break;
		case 'R':
			if (!strnicmp(cp, "Random", 6))
			{
				sort_random = true;
				cp += 5; // Point it to the last char so that the loop's ++cp will point to the character after it.
			}
			else
				g_SortReverse = true;
			break;
		case 'U':  // Unique.
			omit_dupes = true;
			break;
		case 'Z':
			// By setting this to true, the final item in the list, if it ends in a delimiter,
			// is considered to be followed by a blank item:
			allow_last_item_to_be_blank = true;
			break;
		case '\\':
			sort_by_naked_filename = true;
		}
	}

	// size_t helps performance and should be plenty of capacity for many years of advancement.
	// In addition, things like realloc() can't accept anything larger than size_t anyway,
	// so there's no point making this 64-bit until size_t itself becomes 64-bit:
	size_t item_count;

	// Explicitly calculate the length in case it's the clipboard or an environment var.
	// (in which case Length() does not contain the current length).  While calculating
	// the length, also check how many delimiters are present:
	for (item_count = 1, cp = aContents; *cp; ++cp)  // Start at 1 since item_count is delimiter_count+1
		if (*cp == delimiter)
			++item_count;
	size_t aContents_length = cp - aContents;

	// Last item is a delimiter, which means the last item is technically a blank item.
	// However, if the options specify not to allow that, don't count that blank item as
	// an item:
	if (!allow_last_item_to_be_blank && cp > aContents && *(cp - 1) == delimiter)
	{
		terminate_last_item_with_delimiter = true; // So don't consider it to *be* an item.
		--item_count;
	}

	if (item_count == 1) // 1 item is already sorted
		// Put the exact contents back into the output_var, which is necessary in case
		// the variable was an environment variable or the clipboard-containing-files,
		// since in those cases we want the behavior to be consistent regardless of
		// whether there's only 1 item or more than one:
		// Clipboard-contains-files: The single file should be translated into its
		// text equivalent.  Var was an environment variable: the corresponding script
		// variable should be assigned the contents, so it will basically "take over"
		// for the environment variable.
		return output_var->Assign(aContents);

	// Create the array of pointers that points into aContents to each delimited item.
	// Use item_count + 1 to allow space for the last (blank) item in case
	// allow_last_item_to_be_blank is false:
	int unit_size = sort_random ? 2 : 1;
	size_t item_size = unit_size * sizeof(char *);
	char **item = (char **)malloc((item_count + 1) * item_size);
	if (!item)
		return LineError(ERR_OUTOFMEM);  // Short msg. since so rare.

	// If sort_random is in effect, the above has created an array twice the normal size.
	// This allows the random numbers to be interleaved inside the array as though it
	// were an array consisting of sort_rand_type (which it actually is when viewed that way).
	// Because of this, the array should be accessed through pointer addition rather than
	// indexing via [].

	// Scan aContents and do the following:
	// 1) Replace each delimiter with a terminator so that the individual items can be seen
	//    as real strings by the SortWithOptions() and when copying the sorted results back
	//    into output_vav.  It is safe change aContents in this way because
	//    ArgMustBeDereferenced() has ensured that those contents are in the deref buffer.
	// 2) Store a marker/pointer to each item (string) in aContents so that we know where
	//    each item begins for sorting and recopying purposes.
	char **item_curr = item; // i.e. Don't use [] indexing for the reason in the paragraph previous to above.
	for (item_count = 0, cp = *item_curr = aContents; *cp; ++cp)
	{
		if (*cp == delimiter)  // Each delimiter char becomes the terminator of the previous key phrase.
		{
			*cp = '\0';  // Terminate the item that appears before this delimiter.
			++item_count;
			if (sort_random)
				*(item_curr + 1) = (char *)genrand_int31(); // i.e. the randoms are in the odd fields, the pointers in the even.
				// For the above:
				// I don't know the exact reasons, but using genrand_int31() is much more random than
				// using genrand_int32() in this case.  Perhaps it is some kind of statistical/cyclical
				// anomaly in the random number generator.  Or perhaps it's something to do with integer
				// underflow/overflow in SortRandom().  In any case, the problem can be proven via the
				// following script, which shows a sharply non-random distribution when genrand_int32()
				// is used:
				//count = 0
				//Loop 10000
				//{
				//	var = 1`n2`n3`n4`n5`n
				//	Sort, Var, Random
				//	StringLeft, Var1, Var, 1
				//	if Var1 = 5  ; Change this value to 1 to see the opposite problem.
				//		count += 1
				//}
				//Msgbox %count%
				//
				// I e-mailed the author about this sometime around/prior to 12/1/04 but never got a response.
			item_curr += unit_size; // i.e. Don't use [] indexing for the reason described above.
			*item_curr = cp + 1; // Make a pointer to the next item's place in aContents.
		}
	}
	// Add the last item to the count only if it wasn't disqualified earlier:
	if (!terminate_last_item_with_delimiter)
	{
		++item_count;
		if (sort_random) // Provide a random number for the last item.
			*(item_curr + 1) = (char *)genrand_int31(); // i.e. the randoms are in the odd fields, the pointers in the even.
	}
	else // Since the final item is not included in the count, point item_curr to the one before the last, for use below.
		item_curr -= unit_size;
	char *original_last_item = *item_curr;  // The location of the last item before it gets sorted.

	// Now aContents has been divided up based on delimiter.  Sort the array of pointers
	// so that they indicate the correct ordering to copy aContents into output_var:
	if (sort_random) // Takes precedence over all other options.
		qsort((void *)item, item_count, item_size, SortRandom);
	else
		qsort((void *)item, item_count, item_size, sort_by_naked_filename ? SortByNakedFilename : SortWithOptions);

	// Copy the sorted pointers back into output_var, which might not already be sized correctly
	// if it's the clipboard or it was an environment variable when it came in as the input.
	// If output_var is the clipboard, this call will set up the clipboard for writing:
	if (output_var->Assign(NULL, (VarSizeType)aContents_length) != OK) // Might fail due to clipboard problem.
		return FAIL;

	// Set default in case original last item is still the last item, or if last item was omitted due to being a dupe:
	char *pos_of_original_last_item_in_dest = NULL;
	size_t i, item_count_less_1 = item_count - 1;
	DWORD omit_dupe_count = 0;
	bool keep_this_item;
	char *source, *dest;
	char *item_prev = NULL;

	// Copy the sorted result back into output_var.  Do all except the last item, since the last
	// item gets special treatment depending on the options that were specified.  The call to
	// output_var->Contents() below should never fail due to the above having prepped it:
	item_curr = item; // i.e. Don't use [] indexing for the reason described higher above (same applies to item += unit_size below).
	for (dest = output_var->Contents(), i = 0; i < item_count; ++i, item_curr += unit_size)
	{
		keep_this_item = true;  // Set default.
		if (omit_dupes && item_prev)
		{
			// Update to the comment below: Exact dupes will still be removed when sort_by_naked_filename
			// or g_SortColumnOffset is in effect because duplicate lines would still be adjacent to
			// each other even in these modes.  There doesn't appear to be any exceptions, even if
			// some items in the list are sorted as blanks due to being shorter than the specified 
			// g_SortColumnOffset.
			// As documented, special dupe-checking modes are not offered when sort_by_naked_filename
			// is in effect, or g_SortColumnOffset is greater than 1.  That's because the need for such
			// a thing seems too rare (and the result too strange) to justify the extra code size.
			// However, adjacent dupes are still removed when any of the above modes are in effect,
			// or when the "random" mode is in effect.  This might have some usefulness; for example,
			// if a list of songs is sorted in random order, but it contains "favorite" songs listed twice,
			// the dupe-removal feature would remove duplicate songs if they happen to be sorted
			// to lie adjacent to each other, which would be useful to prevent the same song from
			// playing twice in a row.
			if (g_SortNumeric && !g_SortColumnOffset)
				// if g_SortColumnOffset is zero, fall back to the normal dupe checking in case its
				// ever useful to anyone.  This is done because numbers in an offset column are not supported
				// since the extra code size doensn't seem justified given the rarity of the need.
				keep_this_item = (ATOF(*item_curr) != ATOF(item_prev));
			else
				keep_this_item = g_SortCaseSensitive ? strcmp(*item_curr, item_prev) : stricmp(*item_curr, item_prev);
				// Permutations of sorting case sensitive vs. eliminating duplicates based on case sensitivity:
				// 1) Sort is not case sens, but dupes are: Won't work because sort didn't necessarily put
				//    same-case dupes adjacent to each other.
				// 2) Converse: probably not reliable because there could be unrelated items in between
				//    two strings that are duplicates but weren't sorted adjacently due to their case.
				// 3) Both are case sensitive: seems okay
				// 4) Both are not case sensitive: seems okay
				// In light of the above, using the g_SortCaseSensitive flag to control the behavior of
				// both sorting and dupe-removal seems best.
		}
		if (keep_this_item)
		{
			if (*item_curr == original_last_item && i < item_count_less_1) // i.e. If last item is still last, don't update the below.
				pos_of_original_last_item_in_dest = dest;
			for (source = *item_curr; *source;)
				*dest++ = *source++;
			// If we're at the last item and the original list's last item had a terminating delimiter
			// and the specified options said to treat it not as a delimiter but as a final char of sorts,
			// include it after the item that is now last so that the overall layout is the same:
			if (i < item_count_less_1 || terminate_last_item_with_delimiter)
				*dest++ = delimiter;  // Put each item's delimiter back in so that format is the same as the original.
			item_prev = *item_curr; // Since the item just processed above isn't a dupe, save this item to compare against the next item.
		}
		else // This item is a duplicate of the previous item.
			++omit_dupe_count;
			// But don't change the value of item_prev.
	}

	*dest = '\0';  // Terminate the variable's contents.

	// Check if special handling is needed due to the following situation:
	// Delimiter is LF but the contents are lines delimited by CRLF, not just LF
	// and the original/unsorted list's last item was not terminated by an
	// "allowed delimiter".  The symptoms of this are that after the sort, the
	// last item will end in \r when it should end in no delimiter at all.
	// This happens pretty often, such as when the clipboard contains files.
	// In the future, an option letter can be added to turn off this workaround,
	// but it seems unlikely that anyone would ever want to:
	if (delimiter == '\n' && !terminate_last_item_with_delimiter && *(dest - 1) == '\r')
	{
		if (pos_of_original_last_item_in_dest)
		{
			if (cp = strchr(pos_of_original_last_item_in_dest, delimiter))  // Assign
			{
				// Remove the offending '\r':
				--dest;
				*dest = '\0';
				// And insert it where it belongs:
				memmove(cp + 1, cp, dest - cp + 1);  // memmove() allows source & dest to overlap.
				*cp = '\r';
			}
		}
		else if (omit_dupe_count)
		{
			// The item that was originally last was probably omitted from the sorted list.
			// But in that case, it seems best to remove the final \r completely, because
			// it doesn't belong as the final character, nor is there a place to move it
			// to because the line it belongs with doesn't exist in the string.
			--dest;
			*dest = '\0';
		}
		//else do nothing
	}

	free(item); // Free the memory used for the sort.

	if (omit_dupes)
	{
		if (omit_dupe_count) // Update the length to actual whenever at least one dupe was omitted.
			output_var->Length() = (VarSizeType)strlen(output_var->Contents());
		g_ErrorLevel->Assign(omit_dupe_count); // ErrorLevel is set only when dupe-mode is in effect.
	}
	//else it is not necessary to set output_var->Length() here because its length hasn't changed
	// since it was originally set by the above call "output_var->Assign(NULL..."
	return output_var->Close();  // Close in case it's the clipboard.
}



ResultType Line::GetKeyJoyState(char *aKeyName, char *aOption)
// Keep this in sync with FUNC_GETKEYSTATE.
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	JoyControls joy;
	int joystick_id;
	vk_type vk = TextToVK(aKeyName);
    if (!vk)
	{
		if (   !(joy = (JoyControls)ConvertJoy(aKeyName, &joystick_id))   )
			return output_var->Assign("");
		// Since the above didn't return, joy contains a valid joystick button/control ID.
		// Caller needs a token with a buffer of at least this size:
		char buf[MAX_FORMATTED_NUMBER_LENGTH + 1];
		ExprTokenType token;
		// The following must be set for ScriptGetJoyState():
		token.symbol = SYM_STRING;
		token.marker = buf;
		ScriptGetJoyState(joy, joystick_id, token, false);
		ExprTokenToVar(token, *output_var); // Write the result based on whether the token is a string or number.
		// Always returns OK since ScriptGetJoyState() returns FAIL and sets output_var to be blank if
		// the result is indeterminate or there was a problem reading the joystick.  We don't want
		// such a failure to be considered a "critical failure" that will exit the current quasi-thread.
		return OK;
	}
	// Otherwise: There is a virtual key (not a joystick control).
	KeyStateTypes key_state_type;
	switch (toupper(*aOption))
	{
	case 'T': key_state_type = KEYSTATE_TOGGLE; break; // Whether a toggleable key such as CapsLock is currently turned on.
	case 'P': key_state_type = KEYSTATE_PHYSICAL; break; // Physical state of key.
	default: key_state_type = KEYSTATE_LOGICAL;
	}
	return output_var->Assign(ScriptGetKeyState(vk, key_state_type) ? "D" : "U");
}



ResultType Line::DriveSpace(char *aPath, bool aGetFreeSpace)
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

	char buf[MAX_PATH + 1];  // +1 to allow appending of backslash.
	strlcpy(buf, aPath, sizeof(buf));
	size_t length = strlen(buf);
	if (buf[length - 1] != '\\') // Trailing backslash is present, which some of the API calls below don't like.
	{
		if (length + 1 >= sizeof(buf)) // No room to fix it.
			return OK; // Let ErrorLevel tell the story.
		buf[length++] = '\\';
		buf[length] = '\0';
	}

	SetErrorMode(SEM_FAILCRITICALERRORS); // If target drive is a floppy, this avoids a dialog prompting to insert a disk.

	// The program won't launch at all on Win95a (original Win95) unless the function address is resolved
	// at runtime:
	typedef BOOL (WINAPI *GetDiskFreeSpaceExType)(LPCTSTR, PULARGE_INTEGER, PULARGE_INTEGER, PULARGE_INTEGER);
	static GetDiskFreeSpaceExType MyGetDiskFreeSpaceEx =
		(GetDiskFreeSpaceExType)GetProcAddress(GetModuleHandle("kernel32"), "GetDiskFreeSpaceExA");

	// MSDN: "The GetDiskFreeSpaceEx function returns correct values for all volumes, including those
	// that are greater than 2 gigabytes."
	__int64 free_space;
	if (MyGetDiskFreeSpaceEx)  // Function is available (unpatched Win95 and WinNT might not have it).
	{
		ULARGE_INTEGER total, free, used;
		if (!MyGetDiskFreeSpaceEx(buf, &free, &total, &used))
			return OK; // Let ErrorLevel tell the story.
		// Casting this way allows sizes of up to 2,097,152 gigabytes:
		free_space = (__int64)((unsigned __int64)(aGetFreeSpace ? free.QuadPart : total.QuadPart)
			/ (1024*1024));
	}
	else // For unpatched versions of Win95/NT4, fall back to compatibility mode.
	{
		DWORD sectors_per_cluster, bytes_per_sector, free_clusters, total_clusters;
		if (!GetDiskFreeSpace(buf, &sectors_per_cluster, &bytes_per_sector, &free_clusters, &total_clusters))
			return OK; // Let ErrorLevel tell the story.
		free_space = (__int64)((unsigned __int64)((aGetFreeSpace ? free_clusters : total_clusters)
			* sectors_per_cluster * bytes_per_sector) / (1024*1024));
	}

	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return output_var->Assign(free_space);
}



ResultType Line::Drive(char *aCmd, char *aValue, char *aValue2) // aValue not aValue1, for use with a shared macro.
{
	DriveCmds drive_cmd = ConvertDriveCmd(aCmd);

	char path[MAX_PATH + 1];  // +1 to allow room for trailing backslash in case it needs to be added.
	size_t path_length;

	// Notes about the below macro:
	// - It adds a backslash to the contents of the path variable because certain API calls or OS versions
	//   might require it.
	// - It is used by both Drive() and DriveGet().
	// - Leave space for the backslash in case its needed.
	#define DRIVE_SET_PATH \
		strlcpy(path, aValue, sizeof(path) - 1);\
		path_length = strlen(path);\
		if (path_length && path[path_length - 1] != '\\')\
			path[path_length++] = '\\';

	switch(drive_cmd)
	{
	case DRIVE_CMD_INVALID:
		// Since command names are validated at load-time, this only happens if the command name
		// was contained in a variable reference.  Since that is very rare, just set ErrorLevel
		// and return:
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);

	case DRIVE_CMD_LOCK:
	case DRIVE_CMD_UNLOCK:
		return g_ErrorLevel->Assign(DriveLock(*aValue, drive_cmd == DRIVE_CMD_LOCK) ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR); // Indicate success or failure.

	case DRIVE_CMD_EJECT:
		// Don't do DRIVE_SET_PATH in this case since trailing backslash might prevent it from
		// working on some OSes.
		// It seems best not to do the below check since:
		// 1) aValue usually lacks a trailing backslash so that it will work correctly with "open c: type cdaudio".
		//    That lack might prevent DriveGetType() from working on some OSes.
		// 2) It's conceivable that tray eject/retract might work on certain types of drives even though
		//    they aren't of type DRIVE_CDROM.
		// 3) One or both of the calls to mciSendString() will simply fail if the drive isn't of the right type.
		//if (GetDriveType(aValue) != DRIVE_CDROM) // Testing reveals that the below method does not work on Network CD/DVD drives.
		//	return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		char mci_string[256];
		MCIERROR error;
		// Note: The following comment is obsolete because research of MSDN indicates that there is no way
		// not to wait when the tray must be physically opened or closed, at least on Windows XP.  Omitting
		// the word "wait" from both "close cd wait" and "set cd door open/closed wait" does not help, nor
		// does replacing wait with the word notify in "set cdaudio door open/closed wait".
		// The word "wait" is always specified with these operations to ensure consistent behavior across
		// all OSes (on the off-chance that the absence of "wait" really avoids waiting on Win9x or future
		// OSes, or perhaps under certain conditions or for certain types of drives).  See above comment
		// for details.
		if (!*aValue) // When drive is omitted, operate upon default CD/DVD drive.
		{
			snprintf(mci_string, sizeof(mci_string), "set cdaudio door %s wait", ATOI(aValue2) == 1 ? "closed" : "open");
			error = mciSendString(mci_string, NULL, 0, NULL); // Open or close the tray.
			return g_ErrorLevel->Assign(error ? ERRORLEVEL_ERROR : ERRORLEVEL_NONE); // Indicate success or failure.
		}
		snprintf(mci_string, sizeof(mci_string), "open %s type cdaudio alias cd wait shareable", aValue);
		if (mciSendString(mci_string, NULL, 0, NULL)) // Error.
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		snprintf(mci_string, sizeof(mci_string), "set cd door %s wait", ATOI(aValue2) == 1 ? "closed" : "open");
		error = mciSendString(mci_string, NULL, 0, NULL); // Open or close the tray.
		mciSendString("close cd wait", NULL, 0, NULL);
		return g_ErrorLevel->Assign(error ? ERRORLEVEL_ERROR : ERRORLEVEL_NONE); // Indicate success or failure.

	case DRIVE_CMD_LABEL: // Note that is is possible and allowed for the new label to be blank.
		DRIVE_SET_PATH
		SetErrorMode(SEM_FAILCRITICALERRORS);  // So that a floppy drive doesn't prompt for a disk
		return g_ErrorLevel->Assign(SetVolumeLabel(path, aValue2) ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);

	} // switch()

	return FAIL;  // Should never be executed.  Helps catch bugs.
}



ResultType Line::DriveLock(char aDriveLetter, bool aLockIt)
{
	HANDLE hdevice;
	DWORD unused;
	BOOL result;

	if (g_os.IsWin9x())
	{
		// blisteringhot@hotmail.com has confirmed that the code below works on Win98 with an IDE CD Drive:
		// System:  Win98 IDE CdRom (my ejecter is CloseTray)
		// I get a blue screen when I try to eject after using the test script.
		// "eject request to drive in use"
		// It asks me to Ok or Esc, Ok is default.
		//	-probably a bit scary for a novice.
		// BUT its locked alright!"

		// Use the Windows 9x method.  The code below is based on an example posted by Microsoft.
		// Note: The presence of the code below does not add a detectible amount to the EXE size
		// (probably because it's mostly defines and data types).
		#pragma pack(1)
		typedef struct _DIOC_REGISTERS
		{
			DWORD reg_EBX;
			DWORD reg_EDX;
			DWORD reg_ECX;
			DWORD reg_EAX;
			DWORD reg_EDI;
			DWORD reg_ESI;
			DWORD reg_Flags;
		} DIOC_REGISTERS, *PDIOC_REGISTERS;
		typedef struct _PARAMBLOCK
		{
			BYTE Operation;
			BYTE NumLocks;
		} PARAMBLOCK, *PPARAMBLOCK;
		#pragma pack()

		// MS: Prepare for lock or unlock IOCTL call
		#define CARRY_FLAG 0x1
		#define VWIN32_DIOC_DOS_IOCTL 1
		#define LOCK_MEDIA   0
		#define UNLOCK_MEDIA 1
		#define STATUS_LOCK  2
		PARAMBLOCK pb = {0};
		pb.Operation = aLockIt ? LOCK_MEDIA : UNLOCK_MEDIA;
		
		DIOC_REGISTERS regs = {0};
		regs.reg_EAX = 0x440D;
		regs.reg_EBX = toupper(aDriveLetter) - 'A' + 1; // Convert to drive index. 0 = default, 1 = A, 2 = B, 3 = C
		regs.reg_ECX = 0x0848; // MS: Lock/unlock media
		regs.reg_EDX = (DWORD)&pb;
		
		// MS: Open VWIN32
		hdevice = CreateFile("\\\\.\\vwin32", 0, 0, NULL, 0, FILE_FLAG_DELETE_ON_CLOSE, NULL);
		if (hdevice == INVALID_HANDLE_VALUE)
			return FAIL;
		
		// MS: Call VWIN32
		result = DeviceIoControl(hdevice, VWIN32_DIOC_DOS_IOCTL, &regs, sizeof(regs), &regs, sizeof(regs), &unused, 0);
		if (result)
			result = !(regs.reg_Flags & CARRY_FLAG);
	}
	else // NT4/2k/XP or later
	{
		// The calls below cannot work on Win9x (as documented by MSDN's PREVENT_MEDIA_REMOVAL).
		// Don't even attempt them on Win9x because they might blow up.
		char filename[64];
		sprintf(filename, "\\\\.\\%c:", aDriveLetter);
		// FILE_READ_ATTRIBUTES is not enough; it yields "Access Denied" error.  So apparently all or part
		// of the sub-attributes in GENERIC_READ are needed.  An MSDN example implies that GENERIC_WRITE is
		// only needed for GetDriveType() == DRIVE_REMOVABLE drives, and maybe not even those when all we
		// want to do is lock/unlock the drive (that example did quite a bit more).  In any case, research
		// indicates that all CD/DVD drives are ever considered DRIVE_CDROM, not DRIVE_REMOVABLE.
		// Due to this and the unlikelihood that GENERIC_WRITE is ever needed anyway, GetDriveType() is
		// not called for the purpose of conditionally adding the GENERIC_WRITE attribute.
		hdevice = CreateFile(filename, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
		if (hdevice == INVALID_HANDLE_VALUE)
			return FAIL;
		PREVENT_MEDIA_REMOVAL pmr;
		pmr.PreventMediaRemoval = aLockIt;
		result = DeviceIoControl(hdevice, IOCTL_STORAGE_MEDIA_REMOVAL, &pmr, sizeof(PREVENT_MEDIA_REMOVAL)
			, NULL, 0, &unused, NULL);
	}
	CloseHandle(hdevice);
	return result ? OK : FAIL;
}



ResultType Line::DriveGet(char *aCmd, char *aValue)
{
	DriveGetCmds drive_get_cmd = ConvertDriveGetCmd(aCmd);
	if (drive_get_cmd == DRIVEGET_CMD_CAPACITY)
		return DriveSpace(aValue, false);

	char path[MAX_PATH + 1];  // +1 to allow room for trailing backslash in case it needs to be added.
	size_t path_length;

	if (drive_get_cmd == DRIVEGET_CMD_SETLABEL) // The is retained for backward compatibility even though the Drive cmd is normally used.
	{
		DRIVE_SET_PATH
		SetErrorMode(SEM_FAILCRITICALERRORS); // If drive is a floppy, prevents pop-up dialog prompting to insert disk.
		char *new_label = omit_leading_whitespace(aCmd + 9);  // Example: SetLabel:MyLabel
		return g_ErrorLevel->Assign(SetVolumeLabel(path, new_label) ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);
	}

	g_ErrorLevel->Assign(ERRORLEVEL_ERROR);  // Set default since there are many points of return.

	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;

	switch(drive_get_cmd)
	{

	case DRIVEGET_CMD_INVALID:
		// Since command names are validated at load-time, this only happens if the command name
		// was contained in a variable reference.  Since that is very rare, just set ErrorLevel
		// and return:
		return output_var->Assign();  // Let ErrorLevel tell the story.

	case DRIVEGET_CMD_LIST:
	{
		UINT drive_type;
		#define ALL_DRIVE_TYPES 256
		if (!*aValue) drive_type = ALL_DRIVE_TYPES;
		else if (!stricmp(aValue, "CDRom")) drive_type = DRIVE_CDROM;
		else if (!stricmp(aValue, "Removable")) drive_type = DRIVE_REMOVABLE;
		else if (!stricmp(aValue, "Fixed")) drive_type = DRIVE_FIXED;
		else if (!stricmp(aValue, "Network")) drive_type = DRIVE_REMOTE;
		else if (!stricmp(aValue, "Ramdisk")) drive_type = DRIVE_RAMDISK;
		else if (!stricmp(aValue, "Unknown")) drive_type = DRIVE_UNKNOWN;
		else // Let ErrorLevel tell the story.
			return OK;

		char found_drives[32];  // Need room for all 26 possible drive letters.
		int found_drives_count;
		UCHAR letter;
		char buf[128], *buf_ptr;

		SetErrorMode(SEM_FAILCRITICALERRORS); // If drive is a floppy, prevents pop-up dialog prompting to insert disk.

		for (found_drives_count = 0, letter = 'A'; letter <= 'Z'; ++letter)
		{
			buf_ptr = buf;
			*buf_ptr++ = letter;
			*buf_ptr++ = ':';
			*buf_ptr++ = '\\';
			*buf_ptr = '\0';
			UINT this_type = GetDriveType(buf);
			if (this_type == drive_type || (drive_type == ALL_DRIVE_TYPES && this_type != DRIVE_NO_ROOT_DIR))
				found_drives[found_drives_count++] = letter;  // Store just the drive letters.
		}
		found_drives[found_drives_count] = '\0';  // Terminate the string of found drive letters.
		output_var->Assign(found_drives);
		if (!*found_drives)
			return OK;  // Seems best to flag zero drives in the system as default ErrorLevel of "1".
		break;
	}

	case DRIVEGET_CMD_FILESYSTEM:
	case DRIVEGET_CMD_LABEL:
	case DRIVEGET_CMD_SERIAL:
	{
		char volume_name[256];
		char file_system[256];
		DRIVE_SET_PATH
		SetErrorMode(SEM_FAILCRITICALERRORS); // If drive is a floppy, prevents pop-up dialog prompting to insert disk.
		DWORD serial_number, max_component_length, file_system_flags;
		if (!GetVolumeInformation(path, volume_name, sizeof(volume_name) - 1, &serial_number, &max_component_length
			, &file_system_flags, file_system, sizeof(file_system) - 1))
			return output_var->Assign(); // Let ErrorLevel tell the story.
		switch(drive_get_cmd)
		{
		case DRIVEGET_CMD_FILESYSTEM: output_var->Assign(file_system); break;
		case DRIVEGET_CMD_LABEL: output_var->Assign(volume_name); break;
		case DRIVEGET_CMD_SERIAL: output_var->Assign(serial_number); break;
		}
		break;
	}

	case DRIVEGET_CMD_TYPE:
	{
		DRIVE_SET_PATH
		SetErrorMode(SEM_FAILCRITICALERRORS); // If drive is a floppy, prevents pop-up dialog prompting to insert disk.
		switch (GetDriveType(path))
		{
		case DRIVE_UNKNOWN:   output_var->Assign("Unknown"); break;
		case DRIVE_REMOVABLE: output_var->Assign("Removable"); break;
		case DRIVE_FIXED:     output_var->Assign("Fixed"); break;
		case DRIVE_REMOTE:    output_var->Assign("Network"); break;
		case DRIVE_CDROM:     output_var->Assign("CDROM"); break;
		case DRIVE_RAMDISK:   output_var->Assign("RAMDisk"); break;
		default: // DRIVE_NO_ROOT_DIR
			return output_var->Assign();  // Let ErrorLevel tell the story.
		}
		break;
	}

	case DRIVEGET_CMD_STATUS:
	{
		DRIVE_SET_PATH
		SetErrorMode(SEM_FAILCRITICALERRORS); // If drive is a floppy, prevents pop-up dialog prompting to insert disk.
		DWORD sectors_per_cluster, bytes_per_sector, free_clusters, total_clusters;
		switch (GetDiskFreeSpace(path, &sectors_per_cluster, &bytes_per_sector, &free_clusters, &total_clusters)
			? ERROR_SUCCESS : GetLastError())
		{
		case ERROR_SUCCESS:        output_var->Assign("Ready"); break;
		case ERROR_PATH_NOT_FOUND: output_var->Assign("Invalid"); break;
		case ERROR_NOT_READY:      output_var->Assign("NotReady"); break;
		case ERROR_WRITE_PROTECT:  output_var->Assign("ReadOnly"); break;
		default:                   output_var->Assign("Unknown");
		}
		break;
	}

	case DRIVEGET_CMD_STATUSCD:
		// Don't do DRIVE_SET_PATH in this case since trailing backslash might prevent it from
		// working on some OSes.
		// It seems best not to do the below check since:
		// 1) aValue usually lacks a trailing backslash so that it will work correctly with "open c: type cdaudio".
		//    That lack might prevent DriveGetType() from working on some OSes.
		// 2) It's conceivable that tray eject/retract might work on certain types of drives even though
		//    they aren't of type DRIVE_CDROM.
		// 3) One or both of the calls to mciSendString() will simply fail if the drive isn't of the right type.
		//if (GetDriveType(aValue) != DRIVE_CDROM) // Testing reveals that the below method does not work on Network CD/DVD drives.
		//	return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		char mci_string[256], status[128];
		// Note that there is apparently no way to determine via mciSendString() whether the tray is ejected
		// or not, since "open" is returned even when the tray is closed but there is no media.
		if (!*aValue) // When drive is omitted, operate upon default CD/DVD drive.
		{
			if (mciSendString("status cdaudio mode", status, sizeof(status), NULL))
				return output_var->Assign(); // Let ErrorLevel tell the story.
		}
		else // Operate upon a specific drive letter.
		{
			snprintf(mci_string, sizeof(mci_string), "open %s type cdaudio alias cd wait shareable", aValue);
			if (mciSendString(mci_string, NULL, 0, NULL)) // Error.
				return output_var->Assign(); // Let ErrorLevel tell the story.
			MCIERROR error = mciSendString("status cd mode", status, sizeof(status), NULL);
			mciSendString("close cd wait", NULL, 0, NULL);
			if (error)
				return output_var->Assign(); // Let ErrorLevel tell the story.
		}
		// Otherwise, success:
		output_var->Assign(status);
		break;

	} // switch()

	// Note that ControlDelay is not done for the Get type commands, because it seems unnecessary.
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
}



ResultType Line::SoundSetGet(char *aSetting, DWORD aComponentType, int aComponentInstance
	, DWORD aControlType, UINT aMixerID)
// If the caller specifies NULL for aSetting, the mode will be "Get".  Otherwise, it will be "Set".
{
	#define SOUND_MODE_IS_SET aSetting // Boolean: i.e. if it's not NULL, the mode is "SET".
	double setting_percent;
	Var *output_var;
	if (SOUND_MODE_IS_SET)
	{
		output_var = NULL; // To help catch bugs.
		setting_percent = ATOF(aSetting);
		if (setting_percent < -100)
			setting_percent = -100;
		else if (setting_percent > 100)
			setting_percent = 100;
	}
	else // The mode is GET.
	{
		if (   !(output_var = ResolveVarOfArg(0))   )
			return FAIL;  // Don't bother setting ErrorLevel if there's a critical error like this.
		output_var->Assign(); // Init to empty string regardless of whether we succeed here.
	}

	// Rare, since load-time validation would have caught problems unless the params were variable references.
	// Text values for ErrorLevels should be kept below 64 characters in length so that the variable doesn't
	// have to be expanded with a different memory allocation method:
	if (aControlType == MIXERCONTROL_CONTROLTYPE_INVALID || aComponentType == MIXERLINE_COMPONENTTYPE_DST_UNDEFINED)
		return g_ErrorLevel->Assign("Invalid Control Type or Component Type");
	
	// Open the specified mixer ID:
	HMIXER hMixer;
    if (mixerOpen(&hMixer, aMixerID, 0, 0, 0) != MMSYSERR_NOERROR)
		return g_ErrorLevel->Assign("Can't Open Specified Mixer");

	// Find out how many destinations are available on this mixer (should always be at least one):
	int dest_count;
	MIXERCAPS mxcaps;
	if (mixerGetDevCaps((UINT_PTR)hMixer, &mxcaps, sizeof(mxcaps)) == MMSYSERR_NOERROR)
		dest_count = mxcaps.cDestinations;
	else
		dest_count = 1;  // Assume it has one so that we can try to proceed anyway.

	// Find specified line (aComponentType + aComponentInstance):
	MIXERLINE ml = {0};
    ml.cbStruct = sizeof(ml);
	if (aComponentInstance == 1)  // Just get the first line of this type, the easy way.
	{
		ml.dwComponentType = aComponentType;
		if (mixerGetLineInfo((HMIXEROBJ)hMixer, &ml, MIXER_GETLINEINFOF_COMPONENTTYPE) != MMSYSERR_NOERROR)
		{
			mixerClose(hMixer);
			return g_ErrorLevel->Assign("Mixer Doesn't Support This Component Type");
		}
	}
	else
	{
		// Search through each source of each destination, looking for the indicated instance
		// number for the indicated component type:
		int source_count;
		bool found = false;
		for (int d = 0, found_instance = 0; d < dest_count && !found; ++d) // For each destination of this mixer.
		{
			ml.dwDestination = d;
			if (mixerGetLineInfo((HMIXEROBJ)hMixer, &ml, MIXER_GETLINEINFOF_DESTINATION) != MMSYSERR_NOERROR)
				// Keep trying in case the others can be retrieved.
				continue;
			source_count = ml.cConnections;  // Make a copy of this value so that the struct can be reused.
			for (int s = 0; s < source_count && !found; ++s) // For each source of this destination.
			{
				ml.dwDestination = d; // Set it again in case it was changed.
				ml.dwSource = s;
				if (mixerGetLineInfo((HMIXEROBJ)hMixer, &ml, MIXER_GETLINEINFOF_SOURCE) != MMSYSERR_NOERROR)
					// Keep trying in case the others can be retrieved.
					continue;
				// This line can be used to show a soundcard's component types (match them against mmsystem.h):
				//MsgBox(ml.dwComponentType);
				if (ml.dwComponentType == aComponentType)
				{
					++found_instance;
					if (found_instance == aComponentInstance)
						found = true;
				}
			} // inner for()
		} // outer for()
		if (!found)
		{
			mixerClose(hMixer);
			return g_ErrorLevel->Assign("Mixer Doesn't Have That Many of That Component Type");
		}
	}

	// Find the mixer control (aControlType) for the above component:
    MIXERCONTROL mc; // MSDN: "No initialization of the buffer pointed to by [pamxctrl below] is required"
    MIXERLINECONTROLS mlc;
	mlc.cbStruct = sizeof(mlc);
	mlc.pamxctrl = &mc;
	mlc.cbmxctrl = sizeof(mc);
	mlc.dwLineID = ml.dwLineID;
	mlc.dwControlType = aControlType;
	mlc.cControls = 1;
	if (mixerGetLineControls((HMIXEROBJ)hMixer, &mlc, MIXER_GETLINECONTROLSF_ONEBYTYPE) != MMSYSERR_NOERROR)
	{
		mixerClose(hMixer);
		return g_ErrorLevel->Assign("Component Doesn't Support This Control Type");
	}

	// Does user want to adjust the current setting by a certain amount?
	// For v1.0.25, the first char of RAW_ARG is also checked in case this is an expression intended
	// to be a positive offset, such as +(var + 10)
	bool adjust_current_setting = aSetting && (*aSetting == '-' || *aSetting == '+' || *RAW_ARG1 == '+');

	// These are used in more than once place, so always initialize them here:
	MIXERCONTROLDETAILS mcd = {0};
    MIXERCONTROLDETAILS_UNSIGNED mcdMeter;
	mcd.cbStruct = sizeof(MIXERCONTROLDETAILS);
	mcd.dwControlID = mc.dwControlID;
	mcd.cChannels = 1; // MSDN: "when an application needs to get and set all channels as if they were uniform"
	mcd.paDetails = &mcdMeter;
	mcd.cbDetails = sizeof(mcdMeter);

	// Get the current setting of the control, if necessary:
	if (!SOUND_MODE_IS_SET || adjust_current_setting)
	{
		if (mixerGetControlDetails((HMIXEROBJ)hMixer, &mcd, MIXER_GETCONTROLDETAILSF_VALUE) != MMSYSERR_NOERROR)
		{
			mixerClose(hMixer);
			return g_ErrorLevel->Assign("Can't Get Current Setting");
		}
	}

	bool control_type_is_boolean;
	switch (aControlType)
	{
	case MIXERCONTROL_CONTROLTYPE_ONOFF:
	case MIXERCONTROL_CONTROLTYPE_MUTE:
	case MIXERCONTROL_CONTROLTYPE_MONO:
	case MIXERCONTROL_CONTROLTYPE_LOUDNESS:
	case MIXERCONTROL_CONTROLTYPE_STEREOENH:
	case MIXERCONTROL_CONTROLTYPE_BASS_BOOST:
		control_type_is_boolean = true;
		break;
	default: // For all others, assume the control can have more than just ON/OFF as its allowed states.
		control_type_is_boolean = false;
	}

	if (SOUND_MODE_IS_SET)
	{
		if (control_type_is_boolean)
		{
			if (adjust_current_setting) // The user wants this toggleable control to be toggled to its opposite state:
				mcdMeter.dwValue = (mcdMeter.dwValue > mc.Bounds.dwMinimum) ? mc.Bounds.dwMinimum : mc.Bounds.dwMaximum;
			else // Set the value according to whether the user gave us a setting that is greater than zero:
				mcdMeter.dwValue = (setting_percent > 0.0) ? mc.Bounds.dwMaximum : mc.Bounds.dwMinimum;
		}
		else // For all others, assume the control can have more than just ON/OFF as its allowed states.
		{
			// Make this an __int64 vs. DWORD to avoid underflow (so that a setting_percent of -100
			// is supported whenenver the difference between Min and Max is large, such as MAXDWORD):
			__int64 specified_vol = (__int64)((mc.Bounds.dwMaximum - mc.Bounds.dwMinimum) * (setting_percent / 100.0));
			if (adjust_current_setting)
			{
				// Make it a big int so that overflow/underflow can be detected:
				__int64 vol_new = mcdMeter.dwValue + specified_vol;
				if (vol_new < mc.Bounds.dwMinimum)
					vol_new = mc.Bounds.dwMinimum;
				else if (vol_new > mc.Bounds.dwMaximum)
					vol_new = mc.Bounds.dwMaximum;
				mcdMeter.dwValue = (DWORD)vol_new;
			}
			else
				mcdMeter.dwValue = (DWORD)specified_vol; // Due to the above, it's known to be positive in this case.
		}

		MMRESULT result = mixerSetControlDetails((HMIXEROBJ)hMixer, &mcd, MIXER_GETCONTROLDETAILSF_VALUE);
		mixerClose(hMixer);
		return g_ErrorLevel->Assign(result == MMSYSERR_NOERROR ? ERRORLEVEL_NONE : "Can't Change Setting");
	}

	// Otherwise, the mode is "Get":
	mixerClose(hMixer);
	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.

	if (control_type_is_boolean)
		return output_var->Assign(mcdMeter.dwValue ? "On" : "Off");
	else // For all others, assume the control can have more than just ON/OFF as its allowed states.
		// The MSDN docs imply that values fetched via the above method do not distinguish between
		// left and right volume levels, unlike waveOutGetVolume():
		return output_var->Assign(   ((double)100 * (mcdMeter.dwValue - (DWORD)mc.Bounds.dwMinimum))
			/ (mc.Bounds.dwMaximum - mc.Bounds.dwMinimum)   );
}



ResultType Line::SoundGetWaveVolume(HWAVEOUT aDeviceID)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	output_var->Assign(); // Init to empty string regardless of whether we succeed here.

	DWORD current_vol;
	if (waveOutGetVolume(aDeviceID, &current_vol) != MMSYSERR_NOERROR)
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);

	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.

	// Return only the left channel volume level (in case right is different, or device is mono vs. stereo).
	// MSDN: "If a device does not support both left and right volume control, the low-order word
	// of the specified location contains the mono volume level.
	return output_var->Assign((double)(LOWORD(current_vol) * 100) / 0xFFFF);
}



ResultType Line::SoundSetWaveVolume(char *aVolume, HWAVEOUT aDeviceID)
{
	double volume = ATOF(aVolume);
	if (volume < -100)
		volume = -100;
	else if (volume > 100)
		volume = 100;

	// Make this an int vs. WORD so that negative values are supported (e.g. adjust volume by -10%).
	int specified_vol_per_channel = (int)(0xFFFF * (volume / 100));
	DWORD vol_new;

	// For v1.0.25, the first char of RAW_ARG is also checked in case this is an expression intended
	// to be a positive offset, such as +(var + 10)
	if (*aVolume == '-' || *aVolume == '+' || *RAW_ARG1 == '+') // User wants to adjust the current level by a certain amount.
	{
		DWORD current_vol;
		if (waveOutGetVolume(aDeviceID, &current_vol) != MMSYSERR_NOERROR)
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		// Adjust left & right independently so that we at least attempt to retain the user's
		// balance setting (if overflow or underflow occurs, the relative balance might be lost):
		int vol_left = LOWORD(current_vol); // Make them ints so that overflow/underflow can be detected.
		int vol_right = HIWORD(current_vol);
		vol_left += specified_vol_per_channel;
		vol_right += specified_vol_per_channel;
		// Handle underflow or overflow:
		if (vol_left < 0)
			vol_left = 0;
		else if (vol_left > 0xFFFF)
			vol_left = 0xFFFF;
		if (vol_right < 0)
			vol_right = 0;
		else if (vol_right > 0xFFFF)
			vol_right = 0xFFFF;
		vol_new = MAKELONG((WORD)vol_left, (WORD)vol_right);  // Left is low-order, right is high-order.
	}
	else // User wants the volume level set to an absolute level (i.e. ignore its current level).
		vol_new = MAKELONG((WORD)specified_vol_per_channel, (WORD)specified_vol_per_channel);

	if (waveOutSetVolume(aDeviceID, vol_new) == MMSYSERR_NOERROR)
		return g_ErrorLevel->Assign(ERRORLEVEL_NONE);
	else
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
}



ResultType Line::SoundPlay(char *aFilespec, bool aSleepUntilDone)
{
	char *cp = omit_leading_whitespace(aFilespec);
	if (*cp == '*')
		return g_ErrorLevel->Assign(MessageBeep(ATOU(cp + 1)) ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);
		// ATOU() returns 0xFFFFFFFF for -1, which is relied upon to support the -1 sound.
	// See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/multimed/htm/_win32_play.asp
	// for some documentation mciSendString() and related.
	char buf[MAX_PATH * 2]; // Allow room for filename and commands.
	mciSendString("status " SOUNDPLAY_ALIAS " mode", buf, sizeof(buf), NULL);
	if (*buf) // "playing" or "stopped" (so close it before trying to re-open with a new aFilespec).
		mciSendString("close " SOUNDPLAY_ALIAS, NULL, 0, NULL);
	snprintf(buf, sizeof(buf), "open \"%s\" alias " SOUNDPLAY_ALIAS, aFilespec);
	if (mciSendString(buf, NULL, 0, NULL)) // Failure.
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);  // Let ErrorLevel tell the story.
	g_SoundWasPlayed = true;  // For use by Script's destructor.
	if (mciSendString("play " SOUNDPLAY_ALIAS, NULL, 0, NULL)) // Failure.
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);  // Let ErrorLevel tell the story.
	// Otherwise, the sound is now playing.
	g_ErrorLevel->Assign(ERRORLEVEL_NONE);
	if (!aSleepUntilDone)
		return OK;
	// Otherwise, caller wants us to wait until the file is done playing.  To allow our app to remain
	// responsive during this time, use a loop that checks our message queue:
	// Older method: "mciSendString("play " SOUNDPLAY_ALIAS " wait", NULL, 0, NULL)"
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



ResultType Line::FileSelectFile(char *aOptions, char *aWorkingDir, char *aGreeting, char *aFilter)
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
	
	// Large in case more than one file is allowed to be selected.
	// The call to GetOpenFileName() may fail if the first character of the buffer isn't NULL
	// because it then thinks the buffer contains the default filename, which if it's uninitialized
	// may be a string that's too long.
	char file_buf[65535] = ""; // Set default.

	char working_dir[MAX_PATH];
	if (!aWorkingDir || !*aWorkingDir)
		*working_dir = '\0';
	else
	{
		strlcpy(working_dir, aWorkingDir, sizeof(working_dir));
		DWORD attr = GetFileAttributes(working_dir);
		if (attr == 0xFFFFFFFF || !(attr & FILE_ATTRIBUTE_DIRECTORY))
		{
			// Above condition indicates it's either an existing file or an invalid
			// folder/filename (one that doesn't currently exist).  In light of this,
			// it seems best to assume it's a file because the user may want to
			// provide a default SAVE filename, and it would be normal for such
			// a file not to already exist.
			char *last_backslash = strrchr(working_dir, '\\');
			if (last_backslash)
			{
				strlcpy(file_buf, last_backslash + 1, sizeof(file_buf)); // Set the default filename.
				*last_backslash = '\0'; // Make the working directory just the file's path.
			}
			else // the entire working_dir string is the default file.
			{
				strlcpy(file_buf, working_dir, sizeof(file_buf));
				*working_dir = '\0';  // This signals it to use the default directory.
			}
		}
		// else it is a directory, so just leave working_dir set as it was initially.
	}

	char greeting[1024];
	if (aGreeting && *aGreeting)
		strlcpy(greeting, aGreeting, sizeof(greeting));
	else
		// Use a more specific title so that the dialogs of different scripts can be distinguished
		// from one another, which may help script automation in rare cases:
		snprintf(greeting, sizeof(greeting), "Select File - %s", g_script.mFileName);

	// The filter must be terminated by two NULL characters.  One is explicit, the other automatic:
	char filter[1024] = "", pattern[1024] = "";  // Set default.
	if (*aFilter)
	{
		char *pattern_start = strchr(aFilter, '(');
		if (pattern_start)
		{
			// Make pattern a separate string because we want to remove any spaces from it.
			// For example, if the user specified Documents (*.txt; *.doc), the space after
			// the semicolon should be removed for the pattern string itself but not from
			// the displayed version of the pattern:
			strlcpy(pattern, ++pattern_start, sizeof(pattern));
			char *pattern_end = strrchr(pattern, ')'); // strrchr() in case there are other literal parentheses.
			if (pattern_end)
				*pattern_end = '\0';  // If parentheses are empty, this will set pattern to be the empty string.
			else // no closing paren, so set to empty string as an indicator:
				*pattern = '\0';

		}
		else // No open-paren, so assume the entire string is the filter.
			strlcpy(pattern, aFilter, sizeof(pattern));
		if (*pattern)
		{
			// Remove any spaces present in the pattern, such as a space after every semicolon
			// that separates the allowed file extensions.  The API docs specify that there
			// should be no spaces in the pattern itself, even though it's okay if they exist
			// in the displayed name of the file-type:
			StrReplaceAll(pattern, " ", "", true);
			// Also include the All Files (*.*) filter, since there doesn't seem to be much
			// point to making this an option.  This is because the user could always type
			// *.* and press ENTER in the filename field and achieve the same result:
			snprintf(filter, sizeof(filter), "%s%c%s%cAll Files (*.*)%c*.*%c"
				, aFilter, '\0', pattern, '\0', '\0', '\0'); // The final '\0' double-terminates by virtue of the fact that snprintf() itself provides a final terminator.
		}
		else
			*filter = '\0';  // It will use a standard default below.
	}

	OPENFILENAME ofn = {0};
	// OPENFILENAME_SIZE_VERSION_400 must be used for 9x/NT otherwise the dialog will not appear!
	// MSDN: "In an application that is compiled with WINVER and _WIN32_WINNT >= 0x0500, use
	// OPENFILENAME_SIZE_VERSION_400 for this member.  Windows 2000/XP: Use sizeof(OPENFILENAME)
	// for this parameter."
	ofn.lStructSize = g_os.IsWin2000orLater() ? sizeof(OPENFILENAME) : OPENFILENAME_SIZE_VERSION_400;
	ofn.hwndOwner = THREAD_DIALOG_OWNER; // Can be NULL, which is used instead of main window since no need to have main window forced into the background for this.
	ofn.lpstrTitle = greeting;
	ofn.lpstrFilter = *filter ? filter : "All Files (*.*)\0*.*\0Text Documents (*.txt)\0*.txt\0";
	ofn.lpstrFile = file_buf;
	ofn.nMaxFile = sizeof(file_buf) - 1; // -1 to be extra safe.
	// Specifying NULL will make it default to the last used directory (at least in Win2k):
	ofn.lpstrInitialDir = *working_dir ? working_dir : NULL;

	// Note that the OFN_NOCHANGEDIR flag is ineffective in some cases, so we'll use a custom
	// workaround instead.  MSDN: "Windows NT 4.0/2000/XP: This flag is ineffective for GetOpenFileName."
	// In addition, it does not prevent the CWD from changing while the user navigates from folder to
	// folder in the dialog, except perhaps on Win9x.

	// For v1.0.25.05, the new "M" letter is used for a new multi-select method since the old multi-select
	// is faulty in the following ways:
	// 1) If the user selects a single file in a multi-select dialog, the result is inconsistent: it
	//    contains the full path and name of that single file rather than the folder followed by the
	//    single file name as most users would expect.  To make matters worse, it includes a linefeed
	//    after that full path in name, which makes it difficult for a script to determine whether
	//    only a single file was selected.
	// 2) The last item in the list is terminated by a linefeed, which is not as easily used with a
	//    parsing loop as shown in example in the help file.
	bool always_use_save_dialog = false; // Set default.
	bool new_multi_select_method = false; // Set default.
	switch (toupper(*aOptions))
	{
	case 'M':  // Multi-select.
		++aOptions;
		new_multi_select_method = true;
		break;
	case 'S': // Have a "Save" button rather than an "Open" button.
		++aOptions;
		always_use_save_dialog = true;
		break;
	}

	int options = ATOI(aOptions);
	ofn.Flags = OFN_HIDEREADONLY | OFN_EXPLORER | OFN_NODEREFERENCELINKS;
	if (options & 0x10)
		ofn.Flags |= OFN_OVERWRITEPROMPT;
	if (options & 0x08)
		ofn.Flags |= OFN_CREATEPROMPT;
	if (new_multi_select_method || (options & 0x04))
		ofn.Flags |= OFN_ALLOWMULTISELECT;
	if (options & 0x02)
		ofn.Flags |= OFN_PATHMUSTEXIST;
	if (options & 0x01)
		ofn.Flags |= OFN_FILEMUSTEXIST;

	// At this point, we know a dialog will be displayed.  See macro's comments for details:
	DIALOG_PREP
	POST_AHK_DIALOG(0) // Do this only after the above. Must pass 0 for timeout in this case.

	++g_nFileDialogs;
	// Below: OFN_CREATEPROMPT doesn't seem to work with GetSaveFileName(), so always
	// use GetOpenFileName() in that case:
	BOOL result = (always_use_save_dialog || ((ofn.Flags & OFN_OVERWRITEPROMPT) && !(ofn.Flags & OFN_CREATEPROMPT)))
		? GetSaveFileName(&ofn) : GetOpenFileName(&ofn);
	--g_nFileDialogs;

	DIALOG_END

	// Both GetOpenFileName() and GetSaveFileName() change the working directory as a side-effect
	// of their operation.  The below is not a 100% workaround for the problem because even while
	// a new quasi-thread is running (having interrupted this one while the dialog is still
	// displayed), the dialog is still functional, and as a result, the dialog changes the
	// working directory every time the user navigates to a new folder.
	// This is only needed when the user pressed OK, since the dialog auto-restores the
	// working directory if CANCEL is pressed or the window was closed by means other than OK.
	// UPDATE: No, it's needed for CANCEL too because GetSaveFileName/GetOpenFileName will restore
	// the working dir to the wrong dir if the user changed it (via SetWorkingDir) while the
	// dialog was displayed.
	// Restore the original working directory so that any threads suspended beneath this one,
	// and any newly launched ones if there aren't any suspended threads, will have the directory
	// that the user expects.  NOTE: It's possible for g_WorkingDir to have changed via the
	// SetWorkingDir command while the dialog was displayed (e.g. a newly launched quasi-thread):
	if (*g_WorkingDir)
		SetCurrentDirectory(g_WorkingDir);

	if (!result) // User pressed CANCEL vs. OK to dismiss the dialog or there was a problem displaying it.
		// It seems best to clear the variable in these cases, since this is a scripting
		// language where performance is not the primary goal.  So do that and return OK,
		// but leave ErrorLevel set to ERRORLEVEL_ERROR.
		return output_var->Assign(); // Tell it not to free the memory by not calling with "".
	else
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate that the user pressed OK vs. CANCEL.

	if (ofn.Flags & OFN_ALLOWMULTISELECT)
	{
		char *cp;
		if (new_multi_select_method) // v1.0.25.05+ method.
		{
			// If the first terminator in file_buf is also the last, the user selected only
			// a single file:
			size_t length = strlen(file_buf);
			if (!file_buf[length + 1]) // The list contains only a single file (full path and name).
			{
				// v1.0.25.05: To make the result of selecting one file the same as selecting multiple files
				// -- and thus easier to work with in a script -- convert the result into the multi-file
				// format (folder as first item and naked filename as second):
				if (cp = strrchr(file_buf, '\\'))
				{
					*cp = '\n';
					// If the folder is the root folder, add a backslash so that selecting a single
					// file yields the same reported folder as selecting multiple files.  One reason
					// for doing it this way is that SetCurrentDirectory() requires a backslash after
					// a root folder to succeed.  This allows a script to use SetWorkingDir to change
					// to the selected folder before operating on each of the selected/naked filenames.
					if (cp - file_buf == 2 && *(cp - 1) == ':') // e.g. "C:"
					{
						memmove(cp + 1, cp, strlen(cp) + 1); // Make room to insert backslash (since only one file was selcted, the buf is large enough).
						*cp = '\\';
					}
				}
			}
			else // More than one file was selected.
			{
				// Use the same method as the old multi-select format except don't provide a
				// linefeed after the final item.  That final linefeed would make parsing via
				// a parsing loop more complex because a parsing loop would see a blank item
				// at the end of the list:
				for (cp = file_buf;;)
				{
					for (; *cp; ++cp); // Find the next terminator.
					if (!cp[1]) // This is the last file because it's double-terminated, so we're done.
						break;
					*cp = '\n'; // Replace zero-delimiter with a visible/printable delimiter, for the user.
				}
			}
		}
		else  // Old multi-select method is in effect (kept for backward compatibility).
		{
			// Replace all the zero terminators with a delimiter, except the one for the last file
			// (the last file should be followed by two sequential zero terminators).
			// Use a delimiter that can't be confused with a real character inside a filename, i.e.
			// not a comma.  We only have room for one without getting into the complexity of having
			// to expand the string, so \r\n is disqualified for now.
			for (cp = file_buf;;)
			{
				for (; *cp; ++cp); // Find the next terminator.
				*cp = '\n'; // Replace zero-delimiter with a visible/printable delimiter, for the user.
				if (!cp[1]) // This is the last file because it's double-terminated, so we're done.
					break;
			}
		}
	}
	return output_var->Assign(file_buf);
}



ResultType Line::FileCreateDir(char *aDirSpec)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	if (!aDirSpec || !*aDirSpec)
		return OK;  // Return OK because g_ErrorLevel tells the story.

	DWORD attr = GetFileAttributes(aDirSpec);
	if (attr != 0xFFFFFFFF)  // aDirSpec already exists.
	{
		if (attr & FILE_ATTRIBUTE_DIRECTORY)
			g_ErrorLevel->Assign(ERRORLEVEL_NONE);  // Indicate success since it already exists as a dir.
		// else leave as failure, since aDirSpec exists as a file, not a dir.
		return OK;
	}

	// If it has a backslash, make sure all its parent directories exist before we attempt
	// to create this directory:
	char *last_backslash = strrchr(aDirSpec, '\\');
	if (last_backslash)
	{
		char parent_dir[MAX_PATH];
		if (strlen(aDirSpec) >= sizeof(parent_dir)) // avoid overflow
			return OK; // Let ErrorLevel tell the story.
		strlcpy(parent_dir, aDirSpec, last_backslash - aDirSpec + 1); // Omits the last backslash.
		FileCreateDir(parent_dir); // Recursively create all needed ancestor directories.
		if (*g_ErrorLevel->Contents() == *ERRORLEVEL_ERROR)
			return OK; // Let ERRORLEVEL_ERROR tell the story.
	}

	// The above has recursively created all parent directories of aDirSpec if needed.
	// Now we can create aDirSpec.  Be sure to explicitly set g_ErrorLevel since it's value
	// is now indeterminate due to action above:
	return g_ErrorLevel->Assign(CreateDirectory(aDirSpec, NULL) ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);
}



ResultType Line::FileRead(char *aFilespec)
// Returns OK or FAIL.  Will almost always return OK because if an error occurs,
// the script's ErrorLevel variable will be set accordingly.  However, if some
// kind of unexpected and more serious error occurs, such as variable-out-of-memory,
// that will cause FAIL to be returned.
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;

	// Init output var to be blank as an additional indicator of failure (or empty file).
	// Caller must check ErrorLevel to distinguish between an empty file and an error.
	output_var->Assign();
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.

	// Set default options:
	bool translate_crlf_to_lf = false;
	bool is_binary_clipboard = false;

	// It's done as asterisk+option letter to permit future expansion.  A plain asterisk such as used
	// by the FileAppend command would create ambiguity if there was ever an effort to add other asterisk-
	// prefixed options later:
	char *cp = omit_leading_whitespace(aFilespec); // omit leading whitespace only temporarily in case aFilespec contains literal whitespace we want to retain.
	if (*cp == '*')
	{
		// Currently, all the options are mutually exclusive, so only one is checked for.
		++cp;
		switch (toupper(*cp))
		{
		case 'C': // Clipboard (binary).
			is_binary_clipboard = true;
			break;
		case 'T': // Text mode.
            translate_crlf_to_lf = true;
			break;
		}
		// Note: because it's possible for filenames to start with a space (even though Explorer itself
		// won't let you create them that way), allow exactly one space between end of option and the
		// filename itself:
		aFilespec = cp;  // aFilespec is now the option letter after the asterisk, or empty string if there was none.
		if (*aFilespec)
		{
			++aFilespec;
			// Now it's the space or tab (if there is one) after the option letter.  It seems best for
			// future expansion to assume that this is a space or tab even if it's really the start of
			// the filename.  For example, in the future, multiletter optios might be wanted, in which
			// case allowing the omission of the space or tab between *t and the start of the filename
			// would cause the following to be ambiguous:
			// FileRead, OutputVar, *delimC:\File.txt
			// (assuming *delim would accept an optional arg immediately following it).
			// Enforcing this format also simplifies parsing in the future, if there are ever multiple options.
			// It also conforms to the precedent/behavior of GuiControl when it accepts picture sizing options
			// such as *w/h and *x/y
			if (*aFilespec)
				++aFilespec; // And now it's the start of the filename.  This behavior is as documented in the help file.
		}
	}

	// It seems more flexible to allow other processes to read and write the file while we're reading it.
	// For example, this allows the file to be appended to during the read operation, which could be
	// desirable, especially it's a very large log file that would take a long time to read.
	// MSDN: "To enable other processes to share the object while your process has it open, use a combination
	// of one or more of [FILE_SHARE_READ, FILE_SHARE_WRITE]."
	HANDLE hfile = CreateFile(aFilespec, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING
		, FILE_FLAG_SEQUENTIAL_SCAN, NULL); // MSDN says that FILE_FLAG_SEQUENTIAL_SCAN will often improve performance in this case.
	if (hfile == INVALID_HANDLE_VALUE)
		return OK; // Let ErrorLevel tell the story.

	if (is_binary_clipboard && output_var->Type() == VAR_CLIPBOARD)
		return ReadClipboardFromFile(hfile);
	// Otherwise, if is_binary_clipboard, load it directly into a normal variable.  The data in the
	// clipboard file should already have the (UINT)0 as its ending terminator.

	// The program is currently compiled with a 2GB address limit, so loading files larger than that
	// would probably fail or perhaps crash the program.  Therefore, just putting a basic 1.0 GB sanity
	// limit on the file here, for now.  Note: a variable can still be operated upon without necessarily
	// using the deref buffer, since that buffer only comes into play when there is no other way to
	// manipulate the variable.  In other words, the deref buffer won't necessarily grow to be the same
	// size as the file, which if it happened for a 1GB file would exceed the 2GB address limit.
	// That is why a smaller limit such as 800 MB seems too severe:
	unsigned __int64 bytes_to_read = GetFileSize64(hfile);
	if (bytes_to_read > 1024*1024*1024) // Also note that bytes_to_read==ULLONG_MAX means GetFileSize64() failed.
	{
		CloseHandle(hfile);
		return OK; // Let ErrorLevel tell the story.
	}

	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Set default for this point forward to be "success".

	if (!bytes_to_read)
	{
		CloseHandle(hfile);
		return OK; // And ErrorLevel will indicate success (a zero length file results in empty output_var).
	}

	// Set up the var, enlarging it if necessary.  If the output_var is of type VAR_CLIPBOARD,
	// this call will set up the clipboard for writing:
	if (output_var->Assign(NULL, (VarSizeType)bytes_to_read) != OK) // Probably due to "out of memory".
	{
		CloseHandle(hfile);
		return FAIL;  // It already displayed the error. ErrorLevel doesn't matter now because the current quasi-thread will be aborted.
	}
	char *output_buf = output_var->Contents();

	DWORD bytes_actually_read;
	BOOL result = ReadFile(hfile, output_buf, (DWORD)bytes_to_read, &bytes_actually_read, NULL);
	CloseHandle(hfile);

	// Upon result==success, bytes_actually_read is not checked against bytes_to_read because it
	// shouldn't be different (result should have set to failure if there was a read error).
	// If it ever is different, a partial read is considered a success since ReadFile() told us
	// that nothing bad happened.

	if (result)
	{
		output_buf[bytes_actually_read] = '\0';  // Ensure text is terminated where indicated.
		// Since a larger string is being replaced with a smaller, there's a good chance the 2 GB
		// address limit will not be exceeded by StrReplaceAll even if the file is close to the
		// 1 GB limit as described above:
		if (translate_crlf_to_lf)
			StrReplaceAll(output_buf, "\r\n", "\n", false); // Safe only because larger string is being replaced with smaller.
		output_var->Length() = is_binary_clipboard ? (bytes_actually_read - 1) // Length excludes the very last byte of the (UINT)0 terminator.
			: (VarSizeType)strlen(output_buf); // For non-binary, explicitly calculate the "usable" length in case any binary zeros were read.
	}
	else
	{
		// ReadFile() failed.  Since MSDN does not document what is in the buffer at this stage,
		// or whether what's in it is even null-terminated, or even whether bytes_to_read contains
		// a valid value, it seems best to abort the entire operation rather than try to put partial
		// file contents into output_var.  ErrorLevel will indicate the failure.
		// Since ReadFile() failed, to avoid complications or side-effects in functions such as Var::Close(),
		// avoid storing a potentially non-terminated string in the variable.
		*output_buf = '\0';
		output_var->Length() = 0;
		g_ErrorLevel->Assign(ERRORLEVEL_ERROR);  // Override the success default that was set in the middle of this function.
	}

	// ErrorLevel, as set in various places above, indicates success or failure.
	return output_var->Close(is_binary_clipboard);
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
	__int64 line_number = ATOI64(aLineNumber);
	if (line_number < 1)
		return OK;  // Return OK because g_ErrorLevel tells the story.
	FILE *fp = fopen(aFilespec, "r");
	if (!fp)
		return OK;  // Return OK because g_ErrorLevel tells the story.

	// Remember that once the first call to MsgSleep() is done, a new hotkey subroutine
	// may fire and suspend what we're doing here.  Such a subroutine might also overwrite
	// the values our params, some of which may be in the deref buffer.  So be sure not
	// to refer to those strings once MsgSleep() has been done, below.  Alternatively,
	// a copy of such params can be made using our own stack space.

	LONG_OPERATION_INIT

	char buf[READ_FILE_LINE_SIZE];
	for (__int64 i = 0; i < line_number; ++i)
	{
		if (fgets(buf, sizeof(buf) - 1, fp) == NULL) // end-of-file or error
		{
			fclose(fp);
			return OK;  // Return OK because g_ErrorLevel tells the story.
		}
		LONG_OPERATION_UPDATE
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
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
}



ResultType Line::FileAppend(char *aFilespec, char *aBuf, LoopReadFileStruct *aCurrentReadFile)
{
	// The below is avoided because want to allow "nothing" to be written to a file in case the
	// user is doing this to reset it's timestamp (or create an empty file).
	//if (!aBuf || !*aBuf)
	//	return g_ErrorLevel->Assign(ERRORLEVEL_NONE);

	if (aCurrentReadFile) // It always takes precedence over aFilespec.
		aFilespec = aCurrentReadFile->mWriteFileName;
	if (!*aFilespec) // Nothing to write to (caller relies on this check).
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);

	FILE *fp = aCurrentReadFile ? aCurrentReadFile->mWriteFile : NULL;
	bool file_was_already_open = fp;

	bool open_as_binary = (*aFilespec == '*');
	if (open_as_binary && !*(aFilespec + 1)) // Naked "*" means write to stdout.
		// Avoid puts() in case it bloats the code in some compilers. i.e. fputs() is already used,
		// so using it again here shouldn't bloat it:
		return g_ErrorLevel->Assign(fputs(aBuf, stdout) ? ERRORLEVEL_ERROR : ERRORLEVEL_NONE); // fputs() returns 0 on success.
		
	if (open_as_binary)
		// Do not do this because it's possible for filenames to start with a space
		// (even though Explorer itself won't let you create them that way):
		//write_filespec = omit_leading_whitespace(write_filespec + 1);
		// Instead just do this:
		++aFilespec;
	else if (!file_was_already_open) // As of 1.0.25, auto-detect binary if that mode wasn't explicitly specified.
	{
		// sArgVar is used for two reasons:
		// 1) It properly resolves dynamic variables, such as "FileAppend, %VarContainingTheStringClipboardAll%, File".
		// 2) It resolves them only once at a prior stage, rather than having to do them again here
		//    (which helps performance).
		if (mArgc > 0 && sArgVar[0])
		{
			if (sArgVar[0]->Type() == VAR_CLIPBOARDALL)
				return WriteClipboardToFile(aFilespec);
			else if (sArgVar[0]->IsBinaryClip())
			{
				// Since there is at least one deref in Arg #1 and the first deref is binary clipboard,
				// assume this operation's only purpose is to write binary data from that deref to a file.
				// This is because that's the only purpose that seems useful and that's currently supported.
				// In addition, the file is always overwritten in this mode, since appending clipboard data
				// to an existing clipboard file would not work due to:
				// 1) Duplicate clipboard formats not making sense (i.e. two CF_TEXT formats would cause the
				//    first to be overwritten by the second when restoring to clipboard).
				// 2) There is a 4-byte zero terminator at the end of the file.
				if (   !(fp = fopen(aFilespec, "wb"))   ) // Overwrite.
					return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
				g_ErrorLevel->Assign(fwrite(sArgVar[0]->Contents(), sArgVar[0]->Length() + 1, 1, fp)
					? ERRORLEVEL_NONE : ERRORLEVEL_ERROR); // In this case, fwrite() will return 1 on success, 0 on failure.
				fclose(fp);
				return OK;
			}
		}
		// Auto-detection avoids the need to have to translate \r\n to \n when reading
		// a file via the FileRead command.  This seems extremely unlikely to break any
		// existing scripts because the intentional use of \r\r\n in a text file (two
		// consecutive carriage returns) -- which would happen if \r\n were written in
		// text mode -- is so rare as to be close to non-existent.  If this behavior
		// is ever specifically needed, the script can explicitly places some \r\r\n's
		// in the file and then write it as binary mode.
		open_as_binary = strstr(aBuf, "\r\n");
		// Due to "else if", the above will not turn off binary mode if binary was explicitly specified.
		// That is useful to write Unix style text files whose lines end in solitary linefeeds.
	}

	// Check if the file needes to be opened.  As of 1.0.25, this is done here rather than
	// at the time the loop first begins so that:
	// 1) Binary mode can be auto-detected if the first block of text appended to the file
	//    contains any \r\n's.
	// 2) To avoid opening the file if the file-reading loop has zero iterations (i.e. it's
	//    opened only upon first actual use to help performance and avoid changing the
	//    file-modification time when no actual text will be appended).
	if (!file_was_already_open)
	{
		// Open the output file (if one was specified).  Unlike the input file, this is not
		// a critical error if it fails.  We want it to be non-critical so that FileAppend
		// commands in the body of the loop will set ErrorLevel to indicate the problem:
		if (   !(fp = fopen(aFilespec, open_as_binary ? "ab" : "a"))   )
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		if (aCurrentReadFile)
			aCurrentReadFile->mWriteFile = fp;
	}

	// Write to the file:
	g_ErrorLevel->Assign(fputs(aBuf, fp) ? ERRORLEVEL_ERROR : ERRORLEVEL_NONE); // fputs() returns 0 on success.

	if (!aCurrentReadFile)
		fclose(fp);
	// else it's the caller's responsibility, or it's caller's, to close it.

	return OK;
}



ResultType Line::WriteClipboardToFile(char *aFilespec)
// Returns OK or FAIL.  If OK, it sets ErrorLevel to the appropriate result.
// If the clipboard is empty, a zero length file will be written, which seems best for its consistency.
{
	// This method used is very similar to that used in PerformAssign(), so see that section
	// for a large quantity of comments.

	if (!g_clip.Open())
		return LineError(CANT_OPEN_CLIPBOARD_READ); // Make this a critical/stop error since it's unusual and something the user would probably want to know about.

	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default.

	HANDLE hfile = CreateFile(aFilespec, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, 0, NULL); // Overwrite. Unsharable (since reading the file while it is being written would probably produce bad data in this case).
	if (hfile == INVALID_HANDLE_VALUE)
		return g_clip.Close(); // Let ErrorLevel tell the story.

	UINT format;
	HGLOBAL hglobal;
	LPVOID hglobal_locked;
	SIZE_T size;
	DWORD bytes_written;
	BOOL result;
	bool format_is_text, format_is_dib, format_is_meta;
	bool text_was_already_written = false, dib_was_already_written = false, meta_was_already_written = false;

	for (format = 0; format = EnumClipboardFormats(format);)
	{
		format_is_text = (format == CF_TEXT || format == CF_OEMTEXT || format == CF_UNICODETEXT);
		format_is_dib = (format == CF_DIB || format == CF_DIBV5);
		format_is_meta = (format == CF_ENHMETAFILE || format == CF_METAFILEPICT);

		// Only write one Text and one Dib format, omitting the others to save space.  See
		// similar section in PerformAssign() for details:
		if (format_is_text && text_was_already_written
			|| format_is_dib && dib_was_already_written
			|| format_is_meta && meta_was_already_written)
			continue;

		if (format_is_text)
			text_was_already_written = true;
		else if (format_is_dib)
			dib_was_already_written = true;
		else if (format_is_meta)
			meta_was_already_written = true;

		if ((hglobal = GetClipboardData(format)) // Relies on short-circuit boolean order:
			&& (!(size = GlobalSize(hglobal)) || (hglobal_locked = GlobalLock(hglobal)))) // Size of zero or lock succeeded: Include this format.
		{
			if (!WriteFile(hfile, &format, sizeof(format), &bytes_written, NULL)
				|| !WriteFile(hfile, &size, sizeof(size), &bytes_written, NULL))
			{
				if (size)
					GlobalUnlock(hglobal); // hglobal not hglobal_locked.
				break; // File might be in an incomplete state now, but that's okay because the reading process checks for that.
			}

			if (size)
			{
				result = WriteFile(hfile, hglobal_locked, (DWORD)size, &bytes_written, NULL);
				GlobalUnlock(hglobal); // hglobal not hglobal_locked.
				if (!result)
					break; // File might be in an incomplete state now, but that's okay because the reading process checks for that.
			}
			//else hglobal_locked is not valid, so don't reference it or unlock it. In other words, 0 bytes are written for this format.
		}
	}

	g_clip.Close();

	if (!format) // Since the loop was not terminated as a result of a failed WriteFile(), write the 4-byte terminator (otherwise, omit it to avoid further corrupting the file).
		if (WriteFile(hfile, &format, sizeof(format), &bytes_written, NULL))
			g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success (otherwise, leave it set to failure).

	CloseHandle(hfile);
	return OK; // Let ErrorLevel, set above, tell the story.
}



ResultType Line::ReadClipboardFromFile(HANDLE hfile)
// Returns OK or FAIL.  If OK, ErrorLevel is overridden from the callers ERRORLEVEL_ERROR setting to
// ERRORLEVEL_SUCCESS, if appropriate.  This function also closes hfile before returning.
// The method used here is very similar to that used in PerformAssign(), so see that section
// for a large quantity of comments.
{
	if (!g_clip.Open())
	{
		CloseHandle(hfile);
		return LineError(CANT_OPEN_CLIPBOARD_WRITE); // Make this a critical/stop error since it's unusual and something the user would probably want to know about.
	}
	EmptyClipboard(); // Failure is not checked for since it's probably impossible under these conditions.

	UINT format;
	HGLOBAL hglobal;
	LPVOID hglobal_locked;
	SIZE_T size;
	DWORD bytes_read;

    if (!ReadFile(hfile, &format, sizeof(format), &bytes_read, NULL) || bytes_read < sizeof(format))
	{
		g_clip.Close();
		CloseHandle(hfile);
		return OK; // Let ErrorLevel, set by the caller, tell the story.
	}

	while (format)
	{
		if (!ReadFile(hfile, &size, sizeof(size), &bytes_read, NULL) || bytes_read < sizeof(size))
			break; // Leave what's already on the clipboard intact since it might be better than nothing.

	    if (   !(hglobal = GlobalAlloc(GMEM_MOVEABLE, size))   ) // size==0 is okay.
		{
			g_clip.Close();
			CloseHandle(hfile);
			return LineError(ERR_OUTOFMEM, FAIL); // Short msg since so rare.
		}

		if (size) // i.e. Don't try to lock memory of size zero.  It won't work and it's not needed.
		{
			if (   !(hglobal_locked = GlobalLock(hglobal))   )
			{
				GlobalFree(hglobal);
				g_clip.Close();
				CloseHandle(hfile);
				return LineError("GlobalLock", FAIL); // Short msg since so rare.
			}
			if (!ReadFile(hfile, hglobal_locked, (DWORD)size, &bytes_read, NULL) || bytes_read < size)
			{
				GlobalUnlock(hglobal);
				GlobalFree(hglobal); // Seems best not to do SetClipboardData for incomplete format (especially without zeroing the unused portion of global_locked).
				break; // Leave what's already on the clipboard intact since it might be better than nothing.
			}
			GlobalUnlock(hglobal);
		}
		//else hglobal is just an empty format, but store it for completeness/accuracy (e.g. CF_BITMAP).

		SetClipboardData(format, hglobal); // The system now owns hglobal.

		if (!ReadFile(hfile, &format, sizeof(format), &bytes_read, NULL) || bytes_read < sizeof(format))
			break;
	}

	g_clip.Close();
	CloseHandle(hfile);

	if (format) // The loop ended as a result of a file error.
		return OK; // Let ErrorLevel, set above, tell the story.
	else // Indicate success.
		return g_ErrorLevel->Assign(ERRORLEVEL_NONE);
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
	// Testing shows that the ANSI version of FindFirstFile() will not accept a path+pattern longer
	// than 256 or so, even if the pattern would match files whose names are short enough to be legal.
	// Therefore, as of v1.0.25, there is also a hard limit of MAX_PATH on all these variables.
	// MSDN confirms this in a vague way: "In the ANSI version of FindFirstFile(), [plpFileName] is
	// limited to MAX_PATH characters."
	char file_path[MAX_PATH], target_filespec[MAX_PATH];
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

	LONG_OPERATION_INIT

	WIN32_FIND_DATA current_file;
	HANDLE file_search = FindFirstFile(aFilePattern, &current_file);
	bool file_found = (file_search != INVALID_HANDLE_VALUE);
	int failure_count = 0;

	for (; file_found; file_found = FindNextFile(file_search, &current_file))
	{
		LONG_OPERATION_UPDATE
		if (current_file.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) // skip any matching directories.
			continue;
		snprintf(target_filespec, sizeof(target_filespec), "%s%s", file_path, current_file.cFileName);
		if (!DeleteFile(target_filespec))
			++failure_count;
	}

	if (file_search != INVALID_HANDLE_VALUE) // In case the loop had zero iterations.
		FindClose(file_search);
	return g_ErrorLevel->Assign(failure_count); // i.e. indicate success if there were no failures.
}



ResultType Line::FileInstall(char *aSource, char *aDest, char *aFlag)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	bool allow_overwrite = (ATOI(aFlag) == 1);
#ifdef AUTOHOTKEYSC
	if (!allow_overwrite && Util_DoesFileExist(aDest))
		return OK; // Let ErrorLevel tell the story.
	HS_EXEArc_Read oRead;
	// AutoIt3: Open the archive in this compiled exe.
	// Jon gave me some details about why a password isn't needed: "The code in those libararies will
	// only allow files to be extracted from the exe is is bound to (i.e the script that it was
	// compiled with).  There are various checks and CRCs to make sure that it can't be used to read
	// the files from any other exe that is passed."
	if (oRead.Open(g_script.mFileSpec, "") != HS_EXEARC_E_OK)
	{
		MsgBox(ERR_EXE_CORRUPTED, 0, g_script.mFileSpec); // Usually caused by virus corruption. Probably impossible since it was previously opened successfully to run the main script.
		return OK; // Let ErrorLevel tell the story.
	}
	// aSource should be the same as the "file id" used to originally compress the file
	// when it was compiled into an EXE.  So this should seek for the right file:
	if ( oRead.FileExtract(aSource, aDest) != HS_EXEARC_E_OK)
	{
		oRead.Close();
		MsgBox(aSource, 0, "Could not extract file:");
		return OK; // Let ErrorLevel tell the story.
	}
	oRead.Close();
	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
#else
	// v1.0.35.11: Must search in A_ScriptDir by default because that's where ahk2exe will search by default.
	// The old behavior was to search in A_WorkingDir, which seems pointless because ahk2exe would never
	// be able to use that value if the script changes it while running.
	SetCurrentDirectory(g_script.mFileDir);
	if (CopyFile(aSource, aDest, !allow_overwrite))
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	SetCurrentDirectory(g_WorkingDir); // Restore to proper value.
#endif
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
{
	if (!aCalledRecursively)  // i.e. Only need to do this if we're not called by ourself:
	{
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

	// Testing shows that the ANSI version of FindFirstFile() will not accept a path+pattern longer
	// than 256 or so, even if the pattern would match files whose names are short enough to be legal.
	// Therefore, as of v1.0.25, there is also a hard limit of MAX_PATH on all these variables.
	// MSDN confirms this in a vague way: "In the ANSI version of FindFirstFile(), [plpFileName] is
	// limited to MAX_PATH characters."
	char file_pattern[MAX_PATH], file_path[MAX_PATH], target_filespec[MAX_PATH];
	strlcpy(file_pattern, aFilePattern, sizeof(file_pattern));

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

	LONG_OPERATION_INIT

	WIN32_FIND_DATA current_file;
	HANDLE file_search = FindFirstFile(file_pattern, &current_file);
	bool file_found = (file_search != INVALID_HANDLE_VALUE);
	char *cp;
	int failure_count = 0;
	enum attrib_modes {ATTRIB_MODE_NONE, ATTRIB_MODE_ADD, ATTRIB_MODE_REMOVE, ATTRIB_MODE_TOGGLE};
	attrib_modes mode = ATTRIB_MODE_NONE;

	for (; file_found; file_found = FindNextFile(file_search, &current_file))
	{
		LONG_OPERATION_UPDATE
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
		// Testing shows that the ANSI version of FindFirstFile() will not accept a path+pattern longer
		// than 256 or so, even if the pattern would match files whose names are short enough to be legal.
		// Therefore, as of v1.0.25, there is also a hard limit of MAX_PATH on all these variables.
		// MSDN confirms this in a vague way: "In the ANSI version of FindFirstFile(), [plpFileName] is
		// limited to MAX_PATH characters."
		char all_file_pattern[MAX_PATH];
		snprintf(all_file_pattern, sizeof(all_file_pattern), "%s*.*", file_path);
		file_search = FindFirstFile(all_file_pattern, &current_file);
		file_found = (file_search != INVALID_HANDLE_VALUE);
		for (; file_found; file_found = FindNextFile(file_search, &current_file))
		{
			LONG_OPERATION_UPDATE
			if (!(current_file.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
				|| !strcmp(current_file.cFileName, "..") || !strcmp(current_file.cFileName, "."))
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
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default.
	output_var->Assign(); // Init to be blank, in case of failure.

	if (!aFilespec || !*aFilespec)
		return OK;  // Let ErrorLevel indicate an error, since this is probably not what the user intended.

	// Don't use CreateFile() & FileGetSize() size they will fail to work on a file that's in use.
	// Research indicates that this method has no disadvantages compared to the other method.
	WIN32_FIND_DATA found_file;
	HANDLE file_search = FindFirstFile(aFilespec, &found_file);
	if (file_search == INVALID_HANDLE_VALUE)
		return OK;  // Let ErrorLevel Tell the story.
	FindClose(file_search);

	FILETIME local_file_time;
	switch (toupper(aWhichTime))
	{
	case 'C': // File's creation time.
		FileTimeToLocalFileTime(&found_file.ftCreationTime, &local_file_time);
		break;
	case 'A': // File's last access time.
		FileTimeToLocalFileTime(&found_file.ftLastAccessTime, &local_file_time);
		break;
	default:  // 'M', unspecified, or some other value.  Use the file's modification time.
		FileTimeToLocalFileTime(&found_file.ftLastWriteTime, &local_file_time);
	}

    g_ErrorLevel->Assign(ERRORLEVEL_NONE);  // Indicate success.
	char local_file_time_string[128];
	return output_var->Assign(FileTimeToYYYYMMDD(local_file_time_string, local_file_time));
}



int Line::FileSetTime(char *aYYYYMMDD, char *aFilePattern, char aWhichTime
	, FileLoopModeType aOperateOnFolders, bool aDoRecurse, bool aCalledRecursively)
// Returns the number of files and folders that could not be changed due to an error.
// Current limitation: It will not recurse into subfolders unless their names also match
// aFilePattern.
{
	if (!aCalledRecursively)  // i.e. Only need to do this if we're not called by ourself:
	{
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
	char file_pattern[MAX_PATH];
	strlcpy(file_pattern, aFilePattern, sizeof(file_pattern));

	FILETIME ft, ftUTC;
	if (*yyyymmdd)
	{
		// Convert the arg into the time struct as local (non-UTC) time:
		if (!YYYYMMDDToFileTime(yyyymmdd, ft))
			return 0;  // Let ErrorLevel tell the story.
		// Convert from local to UTC:
		if (!LocalFileTimeToFileTime(&ft, &ftUTC))
			return 0;  // Let ErrorLevel tell the story.
	}
	else // User wants to use the current time (i.e. now) as the new timestamp.
		GetSystemTimeAsFileTime(&ftUTC);

	// This following section is very similar to that in FileSetAttrib and FileDelete:
	char file_path[MAX_PATH], target_filespec[MAX_PATH];
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

	LONG_OPERATION_INIT

	WIN32_FIND_DATA current_file;
	HANDLE file_search = FindFirstFile(aFilePattern, &current_file);
	bool file_found = (file_search != INVALID_HANDLE_VALUE);
	int failure_count = 0;
	HANDLE hFile;

	for (; file_found; file_found = FindNextFile(file_search, &current_file))
	{
		LONG_OPERATION_UPDATE
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

		// Open existing file.  Uses CreateFile() rather than OpenFile for an expectation
		// of greater compatibility for all files, and folder support too.
		// FILE_FLAG_NO_BUFFERING might improve performance because all we're doing is
		// changing one of the file's attributes.  FILE_FLAG_BACKUP_SEMANTICS must be
		// used, otherwise changing the time of a directory under NT and beyond will
		// not succeed.  Win95 (not sure about Win98/Me) does not support this, but it
		// should be harmless to specify it even if the OS is Win95:
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
		// Testing shows that the ANSI version of FindFirstFile() will not accept a path+pattern longer
		// than 256 or so, even if the pattern would match files whose names are short enough to be legal.
		// Therefore, as of v1.0.25, there is also a hard limit of MAX_PATH on all these variables.
		// MSDN confirms this in a vague way: "In the ANSI version of FindFirstFile(), [plpFileName] is
		// limited to MAX_PATH characters."
		char all_file_pattern[MAX_PATH];
		snprintf(all_file_pattern, sizeof(all_file_pattern), "%s*.*", file_path);
		file_search = FindFirstFile(all_file_pattern, &current_file);
		file_found = (file_search != INVALID_HANDLE_VALUE);
		for (; file_found; file_found = FindNextFile(file_search, &current_file))
		{
			LONG_OPERATION_UPDATE
			if (!(current_file.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
				|| !strcmp(current_file.cFileName, "..") || !strcmp(current_file.cFileName, "."))
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
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default
	output_var->Assign(); // Init to be blank, in case of failure.

	if (!aFilespec || !*aFilespec)
		return OK;  // Let ErrorLevel indicate an error, since this is probably not what the user intended.

	// Don't use CreateFile() & FileGetSize() size they will fail to work on a file that's in use.
	// Research indicates that this method has no disadvantages compared to the other method.
	WIN32_FIND_DATA found_file;
	HANDLE file_search = FindFirstFile(aFilespec, &found_file);
	if (file_search == INVALID_HANDLE_VALUE)
		return OK;  // Let ErrorLevel Tell the story.
	FindClose(file_search);

	unsigned __int64 size = (found_file.nFileSizeHigh * (unsigned __int64)MAXDWORD) + found_file.nFileSizeLow;

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
	return output_var->Assign((__int64)(size > ULLONG_MAX ? -1 : size)); // i.e. don't allow it to wrap around.
	// The below comment is obsolete in light of the switch to 64-bit integers.  But it might
	// be good to keep for background:
	// Currently, the above is basically subject to a 2 gig limit, I believe, after which the
	// size will appear to be negative.  Beyond a 4 gig limit, the value will probably wrap around
	// to zero and start counting from there as file sizes grow beyond 4 gig (UPDATE: The size
	// is now set to -1 [the maximum DWORD when expressed as a signed int] whenever >4 gig).
	// There's not much sense in putting values larger than 2 gig into the var as a text string
	// containing a positive number because such a var could never be properly handled by anything
	// that compares to it (e.g. IfGreater) or does math on it (e.g. EnvAdd), since those operations
	// use ATOI() to convert the string.  So as a future enhancement (unless the whole program is
	// revamped to use 64bit ints or something) might add an optional param to the end to indicate
	// size should be returned in K(ilobyte) or M(egabyte).  However, this is sorta bad too since
	// adding a param can break existing scripts which use filenames containing commas (delimiters)
	// with this command.  Therefore, I think I'll just add the K/M param now.
	// Also, the above assigns an int because unsigned ints should never be stored in script
	// variables.  This is because an unsigned variable larger than INT_MAX would not be properly
	// converted by ATOI(), which is current standard method for variables to be auto-converted
	// from text back to a number whenever that is needed.
}



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
		// The hook is currently needed to support keeping these keys AlwaysOn or AlwaysOff, though
		// there may be better ways to do it (such as registering them as a hotkey, but
		// that may introduce quite a bit of complexity):
		if (!g_KeybdHook)
			Hotkey::InstallKeybdHook();
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

HWND Line::DetermineTargetWindow(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	HWND target_window; // A variable of this name is used by the macros below.
	IF_USE_FOREGROUND_WINDOW(g.DetectHiddenWindows, aTitle, aText, aExcludeTitle, aExcludeText)
	else if (*aTitle || *aText || *aExcludeTitle || *aExcludeText)
		target_window = WinExist(g, aTitle, aText, aExcludeTitle, aExcludeText);
	else // Use the "last found" window.
		target_window = GetValidLastUsedWindow(g);
	return target_window;
}



#ifndef AUTOHOTKEYSC
int Line::ConvertEscapeChar(char *aFilespec, char aOldChar, char aNewChar, bool aFromAutoIt2)
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
	char new_filespec[MAX_PATH + 10];  // +10 in case StrReplace below would otherwise overflow the buffer.
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
			next_char = cp[1];
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
		// Before the line is written, also do some conversions if the source file is known to
		// be an AutoIt2 script:
		if (aFromAutoIt2)
		{
			// This will not fix all possible uses of A_ScriptDir, just those that are dereferences.
			// For example, this would not be fixed: StringLen, length, a_ScriptDir
			StrReplaceAllSafe(buf, sizeof(buf), "%A_ScriptDir%", "%A_ScriptDir%\\", false);
			// Later can add some other, similar conversions here.
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
	if (aMaxCharsToRead < 1) return 0;
	if (feof(fp)) return -1; // Previous call to this function probably already read the last line.
	if (fgets(aBuf, aMaxCharsToRead, fp) == NULL) // end-of-file or error
	{
		*aBuf = '\0';  // Reset since on error, contents added by fgets() are indeterminate.
		return -1;
	}
	return strlen(aBuf);
}
#endif  // The functions above are not needed by the self-contained version.



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



Line *Line::GetJumpTarget(bool aIsDereferenced)
{
	char *target_label = aIsDereferenced ? ARG1 : RAW_ARG1;
	Label *label = g_script.FindLabel(target_label);
	if (!label)
	{
		if (aIsDereferenced)
			LineError(ERR_NO_LABEL ERR_ABORT, FAIL, target_label);
		else
			LineError(ERR_NO_LABEL, FAIL, target_label);
		return NULL;
	}
	if (!aIsDereferenced)
		mRelatedLine = label->mJumpToLine; // The script loader has ensured that label->mJumpToLine isn't NULL.
	// else don't update it, because that would permanently resolve the jump target, and we want it to stay dynamic.
	// Seems best to do this even for GOSUBs even though it's a bit weird:
	return IsJumpValid(label->mJumpToLine) ? label->mJumpToLine : NULL;
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
	return LineError("A Goto/Gosub must not jump into a block that doesn't enclose it."); // Omit GroupActivate from the error msg since that is rare enough to justify the increase in common-case clarify.
	// Above currently doesn't attempt to detect runtime vs. load-time for the purpose of appending
	// ERR_ABORT.
}



////////////////////////
// BUILT-IN FUNCTIONS //
////////////////////////

// Interface for DynaCall():
#define  DC_MICROSOFT           0x0000      // Default
#define  DC_BORLAND             0x0001      // Borland compat
#define  DC_CALL_CDECL          0x0010      // __cdecl
#define  DC_CALL_STD            0x0020      // __stdcall
#define  DC_RETVAL_MATH4        0x0100      // Return value in ST
#define  DC_RETVAL_MATH8        0x0200      // Return value in ST

#define  DC_CALL_STD_BO         (DC_CALL_STD | DC_BORLAND)
#define  DC_CALL_STD_MS         (DC_CALL_STD | DC_MICROSOFT)
#define  DC_CALL_STD_M8         (DC_CALL_STD | DC_RETVAL_MATH8)

union DYNARESULT                // Various result types
{      
    int     Int;                // Generic four-byte type
    long    Long;               // Four-byte long
    void   *Pointer;            // 32-bit pointer
    float   Float;              // Four byte real
    double  Double;             // 8-byte real
    __int64 Int64;              // big int (64-bit)
};

struct DYNAPARM
{
    union
	{
		int value_int; // Args whose width is less than 32-bit are also put in here because they are right justified within a 32-bit block on the stack.
		float value_float;
		__int64 value_int64;
		double value_double;
		char *str;
    };
	// Might help reduce struct size to keep other members last and adjacent to each other (due to
	// 8-byte alignment caused by the presence of double and __int64 members in the union above).
	DllArgTypes type;
	bool passed_by_address;
	bool is_unsigned; // Allows return value and output parameters to be interpreted as unsigned vs. signed.
};



DYNARESULT DynaCall(int aFlags, void *aFunction, DYNAPARM aParam[], int aParamCount, DWORD &aException
	, void *aRet, int aRetSize)
// Based on the code by Ton Plooy <tonp@xs4all.nl>.
// Call the specified function with the given parameters. Build a proper stack and take care of correct
// return value processing.
{
	aException = 0;  // Set default output parameter for caller.

	// Declaring all variables early should help minimize stack interference of C code with asm.
	DWORD *our_stack;
    int param_size;
	DWORD stack_dword, our_stack_size = 0; // Both might have to be DWORD for _asm.
	BYTE *cp;
    DYNARESULT Res = {0}; // This struct is to be returned to caller by value.
    DWORD esp_start, esp_end, dwEAX, dwEDX;
	int i, esp_delta; // Declare this here rather than later to prevent C code from interfering with esp.

	// Reserve enough space on the stack to handle the worst case of our args (which is currently a
	// maximum of 8 bytes per arg). This avoids any chance that compiler-generated code will use
	// the stack in a way that disrupts our insertion of args onto the stack.
	DWORD reserved_stack_size = aParamCount * 8;
	_asm
	{
		mov our_stack, esp  // our_stack is the location where we will write our args (bypassing "push").
		sub esp, reserved_stack_size  // The stack grows downward, so this "allocates" space on the stack.
	}

	// "Push" args onto the portion of the stack reserved above. Every argument is aligned on a 4-byte boundary.
	// We start at the rightmost argument (i.e. reverse order).
	for (i = aParamCount - 1; i > -1; --i)
	{
		DYNAPARM &this_param = aParam[i]; // For performance and convenience.
		// Push the arg or its address onto the portion of the stack that was reserved for our use above.
		if (this_param.passed_by_address)
		{
			stack_dword = (DWORD)&this_param.value_int; // Any union member would work.
			--our_stack;              // ESP = ESP - 4
			*our_stack = stack_dword; // SS:[ESP] = stack_dword
			our_stack_size += 4;      // Keep track of how many bytes are on our reserved portion of the stack.
		}
		else // this_param's value is contained directly inside the union.
		{
			param_size = (this_param.type == DLL_ARG_INT64 || this_param.type == DLL_ARG_DOUBLE) ? 8 : 4;
			our_stack_size += param_size; // Must be done before our_stack_size is decremented below.  Keep track of how many bytes are on our reserved portion of the stack.
			cp = (BYTE *)&this_param.value_int + param_size - 4; // Start at the right side of the arg and work leftward.
			while (param_size > 0)
			{
				stack_dword = *(DWORD *)cp;  // Get first four bytes
				cp -= 4;                     // Next part of argument
				--our_stack;                 // ESP = ESP - 4
				*our_stack = stack_dword;    // SS:[ESP] = stack_dword
				param_size -= 4;
			}
		}
    }

	if ((aRet != NULL) && ((aFlags & DC_BORLAND) || (aRetSize > 8)))
	{
		// Return value isn't passed through registers, memory copy
		// is performed instead. Pass the pointer as hidden arg.
		our_stack_size += 4;       // Add stack size
		--our_stack;               // ESP = ESP - 4
		*our_stack = (DWORD)aRet;  // SS:[ESP] = pMem
	}

	// Call the function.
	__try // Checked code bloat of __try{} and it doesn't appear to add any size.
	{
		_asm
		{
			add esp, reserved_stack_size // Restore to original position
			mov esp_start, esp      // For detecting whether a DC_CALL_STD function was sent too many or too few args.
			sub esp, our_stack_size // Adjust ESP to indicate that the args have already been pushed onto the stack.
			call [aFunction]        // Stack is now properly built, we can call the function
		}
	}
	__except(EXCEPTION_EXECUTE_HANDLER)
	{
		aException = GetExceptionCode(); // aException is an output parameter for our caller.
	}

	// Even if an exception occurred (perhaps due to the callee having been passed a bad pointer),
	// attempt to restore the stack to prevent making things even worse.
	_asm
	{
		mov esp_end, esp        // See below.
		// For DC_CALL_STD functions (since they pop their own arguments off the stack):
		// Since the stack grows downward in memory, if the value of esp after the call is less than
		// that before the call's args were pushed onto the stack, there are still items left over on
		// the stack, meaning that too many args (or an arg too large) were passed to the callee.
		// Conversely, if esp is now greater that it should be, too many args were popped off the
		// stack by the callee, meaning that too few args were provided to it.  In either case,
		// and even for CDECL, the following line restores esp to what it was before we pushed the
		// function's args onto the stack, which in the case of DC_CALL_STD helps prevent crashes
		// due too too many or to few args having been passed.
		mov esp, esp_start      // See above.
		mov dwEAX, eax          // Save eax/edx registers
		mov dwEDX, edx
	}

	// Possibly adjust stack and read return values.
	// The following is commented out because the stack (esp) is restored above, for both CDECL and STD.
	//if (aFlags & DC_CALL_CDECL)
	//	_asm add esp, our_stack_size    // CDECL requires us to restore the stack after the call.
	if (aFlags & DC_RETVAL_MATH4)
		_asm fstp dword ptr [Res]
	else if (aFlags & DC_RETVAL_MATH8)
		_asm fstp qword ptr [Res]
	else if (!aRet)
	{
		_asm
		{
			mov  eax, [dwEAX]
			mov  DWORD PTR [Res], eax
			mov  edx, [dwEDX]
			mov  DWORD PTR [Res + 4], edx
		}
	}
	else if (((aFlags & DC_BORLAND) == 0) && (aRetSize <= 8))
	{
		// Microsoft optimized less than 8-bytes structure passing
        _asm
		{
			mov ecx, DWORD PTR [aRet]
			mov eax, [dwEAX]
			mov DWORD PTR [ecx], eax
			mov edx, [dwEDX]
			mov DWORD PTR [ecx + 4], edx
		}
	}

	char buf[32];
	esp_delta = esp_start - esp_end; // Positive number means too many args were passed, negative means too few.
	if (esp_delta && (aFlags & DC_CALL_STD))
	{
		*buf = 'A'; // The 'A' prefix indicates the call was made, but with too many or too few args.
		_itoa(esp_delta, buf + 1, 10);
		g_ErrorLevel->Assign(buf); // Assign buf not _itoa()'s return value, which is the wrong location.
	}
	// Too many or too few args takes precedence over reporting the exception because it's more informative.
	// In other words, any exception was likely caused by the fact that there were too many or too few.
	else if (aException)
	{
		// It's a little easier to recongize the common error codes when they're in hex format.
		buf[0] = '0';
		buf[1] = 'x';
		_ultoa(aException, buf + 2, 16);
		g_ErrorLevel->Assign(buf); // Positive ErrorLevel numbers are reserved for exception codes.
	}
	else
		g_ErrorLevel->Assign(ERRORLEVEL_NONE);

	return Res;
}



void ConvertDllArgType(char *aBuf[], DYNAPARM &aDynaParam)
// Helper function for DllCall().  Updates aDynaParam's type and other attributes.
// Caller has ensured that aBuf contains exactly two strings (though the second can be NULL).
{
	char buf[32], *type_string;
	int i;

	// Up to two iterations are done to cover the following cases:
	// No second type because there was no SYM_VAR to get it from:
	//	blank means int
	//	invalid is err
	// (for the below, note that 2nd can't be blank because var name can't be blank, and the first case above would have caught it if 2nd is NULL)
	// 1Blank, 2Invalid: blank (but ensure is_unsigned and passed_by_address get reset)
	// 1Blank, 2Valid: 2
	// 1Valid, 2Invalid: 1 (second iteration would never have run, so no danger of it having erroneously reset is_unsigned/passed_by_address)
	// 1Valid, 2Valid: 1 (same comment)
	// 1Invalid, 2Invalid: invalid
	// 1Invalid, 2Valid: 2

	for (i = 0, type_string = aBuf[0]; i < 2 && type_string; type_string = aBuf[++i])
	{
		if (toupper(*type_string) == 'U') // Unsigned
		{
			aDynaParam.is_unsigned = true;
			++type_string; // Omit the 'U' prefix from further consideration.
		}
		else
			aDynaParam.is_unsigned = false;

		strlcpy(buf, type_string, sizeof(buf)); // Make a modifiable copy for easier parsing below.

		// v1.0.30.02: The addition of 'P' allows the quotes to be omitted around a pointer type.
		// However, the current detection below relies upon the fact that not of the types currently
		// contain the letter P anywhere in them, so it would have to be altered if that ever changes.
		char *cp = StrChrAny(buf, "*pP"); // Asterisk or the letter P.
		if (cp)
		{
			aDynaParam.passed_by_address = true;
			// Remove trailing options so that stricmp() can be used below.
			// Allow optional space in front of asterisk (seems okay even for 'P').
			if (cp > buf && IS_SPACE_OR_TAB(cp[-1]))
			{
				cp = omit_trailing_whitespace(buf, cp - 1);
				cp[1] = '\0'; // Terminate at the leftmost whitespace to remove all whitespace and the suffix.
			}
			else
				*cp = '\0'; // Terminate at the suffix to remove it.
		}
		else
			aDynaParam.passed_by_address = false;

		if (!*buf)
		{
			// The following also serves to set the default in case this is the first iteration.
			// Set default but perform second iteration in case the second type string isn't NULL.
			// In other words, if the second type string is explicitly valid rather than blank,
			// it should override the following default:
			aDynaParam.type = DLL_ARG_INT;  // Assume int.  This is relied upon at least for having a return type such as a naked "CDecl".
			continue; // OK to do this regardless of whether this is the first or second iteration.
		}
		else if (!stricmp(buf, "Str"))     aDynaParam.type = DLL_ARG_STR; // The few most common types are kept up top for performance.
		else if (!stricmp(buf, "Int"))     aDynaParam.type = DLL_ARG_INT;
		else if (!stricmp(buf, "Short"))   aDynaParam.type = DLL_ARG_SHORT;
		else if (!stricmp(buf, "Char"))    aDynaParam.type = DLL_ARG_CHAR;
		else if (!stricmp(buf, "Int64"))   aDynaParam.type = DLL_ARG_INT64;
		else if (!stricmp(buf, "Float"))   aDynaParam.type = DLL_ARG_FLOAT;
		else if (!stricmp(buf, "Double"))  aDynaParam.type = DLL_ARG_DOUBLE;
		// Unnecessary: else if (!stricmp(buf, "None"))    aDynaParam.type = DLL_ARG_NONE;
		// Otherwise, it's blank or an unknown type, leave it set at the default.
		else
		{
			if (i > 0) // Second iteration.
			{
				// Reset flags to go with any blank value we're falling back to from the first iteration
				// (in case our iteration changed the flags based on bogus contents of the second type_string):
				aDynaParam.passed_by_address = false;
				aDynaParam.is_unsigned = false;
			}
			else // First iteration, so aDynaParam.type's value will be set by the second.
				continue;
		}
		// Since above didn't "continue", the type is explicitly valid so "return" to ensure that
		// the second iteration doesn't run (in case this is the first iteration):
		return;
	}
}



void BIF_DllCall(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// Stores a number or a SYM_STRING result in aResultToken.
// Sets ErrorLevel to the error code appropriate to any problem that occurred.
// Caller has set up aParam to be viewable as a left-to-right array of params rather than a stack.
// It has also ensured that the array has exactly aParamCount items in it.
// Author: Marcus Sonntag (Ultra)
{
	// Set default result in case of early return; a blank value:
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = "";

	// Check that the mandatory first parameter (DLL+Function) is present and valid.
	if (aParamCount < 1 || !IS_OPERAND(aParam[0]->symbol) || IS_NUMERIC(aParam[0]->symbol)) // Relies on short-circuit order.
	{
		g_ErrorLevel->Assign("-1"); // Stage 1 error: Too few params or blank/invalid first param.
		return;
	}

	// Determine the type of return value.
	DYNAPARM return_attrib = {0}; // Will hold the type and other attributes of the function's return value.
	int dll_call_mode = DC_CALL_STD; // Set default.  Can be overridden to DC_CALL_CDECL and flags can be OR'd into it.
	if (aParamCount % 2) // Odd number of parameters indicates the return type has been omitted, so assume BOOL/INT.
		return_attrib.type = DLL_ARG_INT;
	else
	{
		// Check validity of this arg's return type:
		ExprTokenType &token = *aParam[aParamCount - 1];
		if (!IS_OPERAND(token.symbol) // Haven't found a way to produce this situation yet, but safe to assume it's possible.
			|| IS_NUMERIC(token.symbol)) // The return type should be a string, not something purely numeric.
		{
			g_ErrorLevel->Assign("-2"); // Stage 2 error: Invalid return type or arg type.
			return;
		}
		char *return_type_string[2];
		if (token.symbol == SYM_VAR)
		{
			return_type_string[0] = token.var->Contents();
			return_type_string[1] = token.var->mName; // v1.0.33.01: Improve convenience by falling back to the variable's name if the contents are not appropriate.
		}
		else
		{
			return_type_string[0] = token.marker;
			return_type_string[1] = NULL;
		}
		if (!strnicmp(return_type_string[0], "CDecl", 5)) // Alternate calling convention.
		{
			dll_call_mode = DC_CALL_CDECL;
			return_type_string[0] = omit_leading_whitespace(return_type_string[0] + 5);
		}
		// This next part is a little iffy because if a legitimate return type is contained in a variable
		// that happens to be named Cdecl, Cdecl will be put into effect regardless of what's in the variable.
		// But the convenience of being able to omit the quotes around Cdecl seems to outweigh the extreme
		// rarity of such a thing happening.
		else if (return_type_string[1] && !strnicmp(return_type_string[1], "CDecl", 5)) // Alternate calling convention.
		{
			dll_call_mode = DC_CALL_CDECL;
			return_type_string[1] = NULL; // Must be NULL since return_type_string[1] is the variable's name, by definition, so it can't have any spaces in it, and thus no space delimited items after "Cdecl".
		}
		ConvertDllArgType(return_type_string, return_attrib);
		if (return_attrib.type == DLL_ARG_INVALID)
		{
			g_ErrorLevel->Assign("-2"); // Stage 2 error: Invalid return type or arg type.
			return;
		}
		--aParamCount;  // Remove the last parameter from further consideration.
		if (!return_attrib.passed_by_address) // i.e. the special return flags below are not needed when an address is being returned.
		{
			if (return_attrib.type == DLL_ARG_DOUBLE)
				dll_call_mode |= DC_RETVAL_MATH8;
			else if (return_attrib.type == DLL_ARG_FLOAT)
				dll_call_mode |= DC_RETVAL_MATH4;
		}
	}

	// Using stack memory, create an array of dll args large enough to hold the actual number of args present.
	int arg_count = aParamCount/2; // Might provide one extra due to first/last params, which is inconsequential.
	DYNAPARM *dyna_param = arg_count ? (DYNAPARM *)_alloca(arg_count * sizeof(DYNAPARM)) : NULL;
	// Above: _alloca() has been checked for code-bloat and it doesn't appear to be an issue.
	// Above: Fix for v1.0.36.07: According to MSDN, on failure, this implementation of _alloca() generates a
	// stack overflow exception rather than returning a NULL value.  Therefore, NULL is no longer checked,
	// nor is an exception block used since stack overflow in this case should be exceptionally rare (if it
	// does happen, it would probably mean the script or the program has a design flaw somewhere, such as
	// infinite recursion).

	char *arg_type_string[2], *arg_as_string;
	int i;

	// Above has already ensured that after the first parameter, there are either zero additional parameters
	// or an even number of them.  In other words, each arg type will have an arg value to go with it.
	// It has also verified that the dyna_param array is large enough to hold all of the args.
	for (arg_count = 0, i = 1; i < aParamCount; ++arg_count, i += 2)  // Same loop as used later below, so maintain them together.
	{
		// Check validity of this arg's type and contents:
		if (!(IS_OPERAND(aParam[i]->symbol) && IS_OPERAND(aParam[i + 1]->symbol)) // Haven't found a way to produce this situation yet, but safe to assume it's possible.
			|| IS_NUMERIC(aParam[i]->symbol)) // The arg type should be a string, not something purely numeric.
		{
			g_ErrorLevel->Assign("-2"); // Stage 2 error: Invalid return type or arg type.
			return;
		}
		// Otherwise, this arg's type is a string as it should be, so retrieve it:
		if (aParam[i]->symbol == SYM_VAR)
		{
			arg_type_string[0] = aParam[i]->var->Contents();
			arg_type_string[1] = aParam[i]->var->mName;
			// v1.0.33.01: arg_type_string2 improves convenience by falling back to the variable's name
			// if the contents are not appropriate.  In other words, both Int and "Int" are treated the same.
			// It's done this way to allow the variable named "Int" to actually contain some other legitimate
			// type-name such as "Str".
		}
		else
		{
			arg_type_string[0] = aParam[i]->marker;
			arg_type_string[1] = NULL;
		}

		ExprTokenType &this_param = *aParam[i + 1];  // This and the next are resolved for performance and convenience.
		DYNAPARM &this_dyna_param = dyna_param[arg_count];

		// If the arg's contents is a string, resolve it once here to simplify things that reference it later.
		// NOTE: aResultToken.buf is not used here to resolve a number to a string because although it would
		// add a little flexibility by allowing up to one string parameter to be a numeric expression, it
		// seems to add more complexity and confusion than its worth to other sections further below:
		if (IS_NUMERIC(this_param.symbol))
			arg_as_string = NULL;
		else
			arg_as_string = (this_param.symbol == SYM_VAR) ? this_param.var->Contents() : this_param.marker;

		// Store the each arg into a dyna_param struct, using its arg type to determine how.
		ConvertDllArgType(arg_type_string, this_dyna_param);
		switch (this_dyna_param.type)
		{
		case DLL_ARG_STR:
			if (arg_as_string)
				this_dyna_param.str = arg_as_string;
			else
			{
				// For now, string args must be real strings rather than floats or ints.  An alternative
				// to this would be to convert it to number using persistent memory from the caller (which
				// is necessary because our own stack memory should not be passed to any function since
				// that might cause it to return a pointer to stack memory, or update an output-parameter
				// to be stack memory, which would be invalid memory upon return to the caller).
				// The complexity of this doesn't seem worth the rarity of the need, so this will be
				// documented in the help file.
				g_ErrorLevel->Assign("-2"); // Stage 2 error: Invalid return type or arg type.
				return;
			}
			break;

		case DLL_ARG_INT:
		case DLL_ARG_SHORT:
		case DLL_ARG_CHAR:
		case DLL_ARG_INT64:
			if (arg_as_string)
			{
				// Support for unsigned values that are 32 bits wide or less is done via ATOI64() since
				// it should be able to handle both signed and unsigned values.  However, unsigned 64-bit
				// values probably require ATOU64(), which will prevent something like -1 from being seen
				// as the largest unsigned 64-bit int, but more importantly there are some other issues
				// with unsigned 64-bit numbers: The script internals use 64-bit signed values everywhere,
				// so unsigned values can only be partially supported for incoming parameters, but probably
				// not for outgoing parameters (values the function changed) or the return value.  Those
				// should probably be written back out to the script as negatives so that other parts of
				// the script, such as expressions, can see them as signed values.  In other words, if the
				// script somehow gets a 64-bit unsigned value into a variable, and that value is larger
				// that LLONG_MAX (i.e. too large for ATOI64 to handle), ATOU64() will be able to resolve
				// it, but any output parameter should be written back out as a negative if it exceeds
				// LLONG_MAX (return values can be written out as unsigned since the script can specify
				// signed to avoid this, since they don't need the incoming detection for ATOU()).
				if (this_dyna_param.is_unsigned && this_dyna_param.type == DLL_ARG_INT64)
					this_dyna_param.value_int64 = (__int64)ATOU64(arg_as_string); // Cast should not prevent called function from seeing it as an undamaged unsigned number.
				else
					this_dyna_param.value_int64 = ATOI64(arg_as_string);
			}
			else if (this_param.symbol == SYM_INTEGER)
				this_dyna_param.value_int64 = this_param.value_int64;
			else
				this_dyna_param.value_int64 = (__int64)this_param.value_double;

			// Values less than or equal to 32-bits wide always get copied into a single 32-bit value
			// because they should be right justified within it for insertion onto the call stack.
			if (this_dyna_param.type != DLL_ARG_INT64) // Shift the 32-bit value into the high-order DWORD of the 64-bit value for later use by DynaCall().
				this_dyna_param.value_int = (int)this_dyna_param.value_int64; // Force a failure if compiler generates code for this that corrupts the union (since the same method is used for the more obscure float vs. double below).
			break;

		case DLL_ARG_FLOAT:
		case DLL_ARG_DOUBLE:
			// This currently doesn't validate that this_dyna_param.is_unsigned==false, since it seems
			// too rare and mostly harmless to worry about something like "Ufloat" having been specified.
			if (arg_as_string)
				this_dyna_param.value_double = (double)ATOF(arg_as_string);
			else if (this_param.symbol == SYM_INTEGER)
				this_dyna_param.value_double = (double)this_param.value_int64;
			else
				this_dyna_param.value_double = this_param.value_double;

			if (this_dyna_param.type == DLL_ARG_FLOAT)
				this_dyna_param.value_float = (float)this_dyna_param.value_double;
			break;

		default: // DLL_ARG_INVALID or a bug due to an unhandled type.
			g_ErrorLevel->Assign("-2"); // Stage 2 error: Invalid return type or arg type.
			return;
		}
	}
    
	void *function = NULL;
	HMODULE hmodule, hmodule_to_free = NULL;
	char param1_buf[MAX_PATH*2], *function_name, *dll_name; // Uses MAX_PATH*2 to hold worst-case PATH plus function name.

	// Define the standard libraries here. If they reside in %SYSTEMROOT%\system32 it is not
	// necessary to specify the full path (it wouldn't make sense anyway).
	static HMODULE sStdModule[] = {GetModuleHandle("user32"), GetModuleHandle("kernel32")
		, GetModuleHandle("comctl32"), GetModuleHandle("gdi32")}; // user32 is listed first for performance.
	static int sStdModule_count = sizeof(sStdModule) / sizeof(HMODULE);

	// Make a modifiable copy of param1 so that the DLL name and function name can be parsed out easily:
	strlcpy(param1_buf, aParam[0]->symbol == SYM_VAR ? aParam[0]->var->Contents() : aParam[0]->marker, sizeof(param1_buf) - 1); // -1 to reserve space for the "A" suffix later below.
	if (   !(function_name = strrchr(param1_buf, '\\'))   ) // No DLL name specified, so a search among standard defaults will be done.
	{
		dll_name = NULL;
		function_name = param1_buf;

		// Since no DLL was specified, search for the specified function among the standard modules.
		for (i = 0; i < sStdModule_count; ++i)
			if (   sStdModule[i] && (function = (void *)GetProcAddress(sStdModule[i], function_name))   )
				break;
		if (!function)
		{
			// Since the absence of the "A" suffix (e.g. MessageBoxA) is so common, try it that way
			// but only here with the standard libraries since the risk of ambiguity (calling the wrong
			// function) seems unacceptably high in a custom DLL.  For example, a custom DLL might have
			// function called "AA" but not one called "A".
			strcat(function_name, "A"); // 1 byte of memory was already reserved above for the 'A'.
			for (i = 0; i < sStdModule_count; ++i)
				if (   sStdModule[i] && (function = (void *)GetProcAddress(sStdModule[i], function_name))   )
					break;
		}
	}
	else // DLL file name is explicitly present.
	{
		dll_name = param1_buf;
		*function_name = '\0';  // Terminate dll_name to split it off from function_name.
		++function_name; // Set it to the character after the last backslash.

		// Get module handle. This will work when DLL is already loaded and might improve performance if
		// LoadLibrary is a high-overhead call even when the library already being loaded.  If
		// GetModuleHandle() fails, fall back to LoadLibrary().
		if (   !(hmodule = GetModuleHandle(dll_name))    )
			if (   !(hmodule = hmodule_to_free = LoadLibrary(dll_name))   )
			{
				g_ErrorLevel->Assign("-3"); // Stage 3 error: DLL couldn't be loaded.
				return;
			}
		if (   !(function = (void *)GetProcAddress(hmodule, function_name))   )
		{
			// v1.0.34: If it's one of the standard libraries, try the "A" suffix.
			for (i = 0; i < sStdModule_count; ++i)
				if (hmodule == sStdModule[i]) // Match found.
				{
					strcat(function_name, "A"); // 1 byte of memory was already reserved above for the 'A'.
					function = (void *)GetProcAddress(hmodule, function_name);
					break;
				}
		}
	}

	if (!function)
	{
		g_ErrorLevel->Assign("-4"); // Stage 4 error: Function could not be found in the DLL(s).
		goto end;
	}

	////////////////////////
	// Call the DLL function
	////////////////////////
	DWORD exception_occurred; // Must not be named "exception_code" to avoid interfering with MSVC macros.
	DYNARESULT return_value;  // Doing assignment as separate step avoids compiler warning about "goto end" skipping it.
	return_value = DynaCall(dll_call_mode, function, dyna_param, arg_count, exception_occurred, NULL, 0);
	// The above has also set g_ErrorLevel appropriately.

	if (exception_occurred)
	{
		// If the called function generated an exception, I think it's impossible for the return value
		// to be valid/meaningful since it the function never returned properly.  Confirmation of this
		// would be good, but in the meantime it seems best to make the return value an empty string as
		// an indicator that the call failed (in addition to ErrorLevel).
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = "";
		// But continue on to write out any output parameters because the called function might have
		// had a chance to update them before aborting.
	}
	else // The call was successful.  Interpret and store the return value.
	{
		// If the return value is passed by address, dereference it here.
		if (return_attrib.passed_by_address)
		{
			return_attrib.passed_by_address = false; // Because the address is about to be dereferenced/resolved.

			switch(return_attrib.type)
			{
			case DLL_ARG_STR:  // Even strings can be passed by address, which is equivalent to "char **".
			case DLL_ARG_INT:
			case DLL_ARG_SHORT:
			case DLL_ARG_CHAR:
			case DLL_ARG_FLOAT:
				// All the above are stored in four bytes, so a straight dereference will copy the value
				// over unchanged, even if it's a float.
				return_value.Int = *(int *)return_value.Pointer;
				break;

			case DLL_ARG_INT64:
			case DLL_ARG_DOUBLE:
				// Same as above but for eight bytes:
				return_value.Int64 = *(__int64 *)return_value.Pointer;
				break;
			}
		}

		switch(return_attrib.type)
		{
		case DLL_ARG_STR:
			// The contents of the string returned from the function must not reside in our stack memory since
			// that will vanish when we return to our caller.  As long as every string that went into the
			// function isn't on our stack (which is the case), there should be no way for what comes out to be
			// on the stack either.
			aResultToken.symbol = SYM_STRING;
			aResultToken.marker = (char *)(return_value.Pointer ? return_value.Pointer : "");
			// Above: Fix for v1.0.33.01: Don't allow marker to be set to NULL, which prevents crash
			// with something like the following, which in this case probably happens because the inner
			// call produces a non-numeric string, which "int" then sees as zero, which CharLower() then
			// sees as NULL, which causes CharLower to return NULL rather than a real string:
			//result := DllCall("CharLower", "int", DllCall("CharUpper", "str", MyVar, "str"), "str")
			break;
		case DLL_ARG_INT: // If the function has a void return value (formerly DLL_ARG_NONE), the value assigned here is undefined and inconsequential since the script should be designed to ignore it.
			aResultToken.symbol = SYM_INTEGER;
			if (return_attrib.is_unsigned)
				aResultToken.value_int64 = (UINT)return_value.Int; // Preserve unsigned nature upon promotion to signed 64-bit.
			else // Signed.
				aResultToken.value_int64 = return_value.Int;
			break;
		case DLL_ARG_SHORT:
			aResultToken.symbol = SYM_INTEGER;
			if (return_attrib.is_unsigned)
				aResultToken.value_int64 = return_value.Int & 0x0000FFFF; // This also forces the value into the unsigned domain of a signed int.
			else // Signed.
				aResultToken.value_int64 = (SHORT)(WORD)return_value.Int; // These casts properly preserve negatives.
			break;
		case DLL_ARG_CHAR:
			aResultToken.symbol = SYM_INTEGER;
			if (return_attrib.is_unsigned)
				aResultToken.value_int64 = return_value.Int & 0x000000FF; // This also forces the value into the unsigned domain of a signed int.
			else // Signed.
				aResultToken.value_int64 = (char)(BYTE)return_value.Int; // These casts properly preserve negatives.
			break;
		case DLL_ARG_INT64:
			// Even for unsigned 64-bit values, it seems best both for simplicity and consistency to write
			// them back out to the script as signed values because script internals are not currently
			// equipped to handle unsigned 64-bit values.  This has been documented.
			aResultToken.symbol = SYM_INTEGER;
			aResultToken.value_int64 = return_value.Int64;
			break;
		case DLL_ARG_FLOAT:
			aResultToken.symbol = SYM_FLOAT;
			aResultToken.value_double = return_value.Float;
			break;
		case DLL_ARG_DOUBLE:
			aResultToken.symbol = SYM_FLOAT; // There is no SYM_DOUBLE since all floats are stored as doubles.
			aResultToken.value_double = return_value.Double;
			break;
		default: // Should never be reached unless there's a bug.
			aResultToken.symbol = SYM_STRING;
			aResultToken.marker = "";
		} // switch(return_attrib.type)
	} // Storing the return value when no exception occurred.

	// Store any output parameters back into the input variables.  This allows a function to change the
	// contents of a variable for the following arg types: String and Pointer to <various number types>.
	for (arg_count = 0, i = 1; i < aParamCount; ++arg_count, i += 2) // Same loop as used above, so maintain them together.
	{
		ExprTokenType &this_param = *aParam[i + 1];  // This and the next are resolved for performance and convenience.
		DYNAPARM &this_dyna_param = dyna_param[arg_count];

		if (this_param.symbol != SYM_VAR) // Output parameters are copied back only if its counterpart parameter is a naked variable.
			continue;
		Var &output_var = *this_param.var; // For performance and convenience.
		if (this_dyna_param.type == DLL_ARG_STR) // The function might have altered Contents(), so update Length().
		{
			char *contents = output_var.Contents();
			VarSizeType capacity = output_var.Capacity();
			// Since the performance cost is low, ensure the string is terminated at the limit of its
			// capacity (helps prevent crashes if DLL function didn't do its job and terminate the string,
			// or when a function is called that deliberately doesn't terminate the string, such as
			// RtlMoveMemory()).
			if (capacity)
				contents[capacity - 1] = '\0';
			output_var.Length() = (VarSizeType)strlen(contents);
			continue;
		}

		// Since above didn't "continue", this arg wasn't passed as a string.  Of the remaining types, only
		// those passed by address can possibly be output parameters, so skip the rest:
		if (!this_dyna_param.passed_by_address)
			continue;

		switch (this_dyna_param.type)
		{
		// case DLL_ARG_STR:  Already handled above.
		case DLL_ARG_INT:
			if (this_dyna_param.is_unsigned)
				output_var.Assign((DWORD)this_dyna_param.value_int);
			else // Signed.
				output_var.Assign(this_dyna_param.value_int);
			break;
		case DLL_ARG_SHORT:
			if (this_dyna_param.is_unsigned) // Force omission of the high-order word in case it is non-zero from a parameter that was originally and erroneously larger than a short.
				output_var.Assign(this_dyna_param.value_int & 0x0000FFFF); // This also forces the value into the unsigned domain of a signed int.
			else // Signed.
				output_var.Assign((int)(SHORT)(WORD)this_dyna_param.value_int); // These casts properly preserve negatives.
			break;
		case DLL_ARG_CHAR:
			if (this_dyna_param.is_unsigned) // Force omission of the high-order word in case it is non-zero from a parameter that was originally and erroneously larger than a short.
				output_var.Assign(this_dyna_param.value_int & 0x000000FF); // This also forces the value into the unsigned domain of a signed int.
			else // Signed.
				output_var.Assign((int)(char)(BYTE)this_dyna_param.value_int); // These casts properly preserve negatives.
			break;
		case DLL_ARG_INT64: // Unsigned and signed are both written as signed for the reasons described elsewhere above.
			output_var.Assign(this_dyna_param.value_int64);
			break;
		case DLL_ARG_FLOAT:
			output_var.Assign(this_dyna_param.value_float);
			break;
		case DLL_ARG_DOUBLE:
			output_var.Assign(this_dyna_param.value_double);
			break;
		}
	}

end:
	if (hmodule_to_free)
		FreeLibrary(hmodule_to_free);
}



void BIF_StrLen(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	// Loadtime validation has ensured that there's exactly one actual parameter.
	if (aParam[0]->symbol == SYM_VAR && aParam[0]->var->IsBinaryClip())
		aResultToken.value_int64 = aParam[0]->var->Length() + 1;
	else
	{
		// Result will always be an integer.
		// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.
		char *cp;
		if (   !(cp = ExprTokenToString(*aParam[0], aResultToken.buf))   ) // Allow StrLen(numeric_expr) for flexibility.
			aResultToken.value_int64 = 0; // Not an operand.  Haven't found a way to produce this situation yet, but safe to assume it's possible.
		else
			aResultToken.value_int64 = strlen(cp);
	}
}



void BIF_Asc(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	// Result will always be an integer (this simplifies scripts that work with binary zeros since an
	// empy string yields zero).
	// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.
	char *cp;
	if (   !(cp = ExprTokenToString(*aParam[0], aResultToken.buf))   ) // Not an operand.  Haven't found a way to produce this situation yet, but safe to assume it's possible.
		aResultToken.value_int64 = -1; // Store out-of-bounds value as a flag.
	else
		aResultToken.value_int64 = (UCHAR)*cp;
}



void BIF_Chr(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	int param1 = (int)ExprTokenToInt64(*aParam[0]); // Convert to INT vs. UINT so that negatives can be detected.
	char *cp = aResultToken.buf; // If necessary, it will be moved to a persistent memory location by our caller.
	if (param1 < 0 || param1 > 255)
		*cp = '\0'; // Empty string indicates both Chr(0) and an out-of-bounds param1.
	else
	{
		cp[0] = param1;
		cp[1] = '\0';
	}
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = cp;
}



void BIF_IsLabel(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// For performance and code-size reasons, this function does not currently return what
// type of label it is (hotstring, hotkey, or generic).  To preserve the option to do
// this in the future, it has been documented that the function returns non-zero rather
// than "true".  However, if performance is an issue (since scripts that use IsLabel are
// often performance sensitive), it might be better to add a second parameter that tells
// IsLabel to look up the type of label, and return it as a number or letter.
{
	char *label_name;
	if (   !(label_name = ExprTokenToString(*aParam[0], aResultToken.buf))   ) // Not an operand.  Haven't found a way to produce this situation yet, but safe to assume it's possible.
		aResultToken.value_int64 = 0; // Indicate false, to conform to boolean return type.
	else
		aResultToken.value_int64 = g_script.FindLabel(label_name) ? 1 : 0;
}



void BIF_InStr(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	// Load-time validation has already ensured that at least two actual params are present.
	char needle_buf[MAX_FORMATTED_NUMBER_LENGTH + 1];
	char *haystack = ExprTokenToString(*aParam[0], aResultToken.buf);
	char *needle = ExprTokenToString(*aParam[1], needle_buf);
	if (!haystack || !needle) // Not an operand.  Haven't found a way to produce this situation yet, but safe to assume it's possible.
	{
		aResultToken.value_int64 = -1; // Store out-of-bounds value as a flag.
		return;
	}

	// Result type will always be an integer:
	// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.
	bool case_sensitive = aParamCount >= 3 && ExprTokenToInt64(*aParam[2]);
	char *found_pos;
	__int64 offset = 0; // Set default.

	if (aParamCount >= 4) // There is a starting position present.
	{
		offset = ExprTokenToInt64(*aParam[3]) - 1; // +3 to get the fourth arg.
		if (offset == -1) // Special mode to search from the right side.  Other negative values are reserved for possible future use as offsets from the right side.
		{
			found_pos = strrstr(haystack, needle, case_sensitive, 1);
			aResultToken.value_int64 = found_pos ? (found_pos - haystack + 1) : 0;  // +1 to convert to 1-based, since 0 indicates "not found".
			return;
		}
		// Otherwise, offset is less than -1 or >= 0.
		// Since InStr("", "") yields 1, it seems consistent for InStr("Red", "", 4) to yield
		// 4 rather than 0.  The below takes this into account:
		if (offset < 0 || offset > strlen(haystack))
		{
			aResultToken.value_int64 = 0; // Match never found when offset is beyond length of string.
			return;
		}
	}
	// Since above didn't return:
	haystack += offset; // Above has verified that this won't exceed the length of haystack.
	found_pos = case_sensitive ? strstr(haystack, needle) : strcasestr(haystack, needle);
	aResultToken.value_int64 = found_pos ? (found_pos - haystack + offset + 1) : 0;  // +1 
}



void BIF_GetKeyState(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	char *key_name, key_name_buf[MAX_FORMATTED_NUMBER_LENGTH + 1];
	if (   !(key_name = ExprTokenToString(*aParam[0], key_name_buf))   ) // Not an operand.  Haven't found a way to produce this situation yet, but safe to assume it's possible.
	{
		aResultToken.value_int64 = -1; // Store out-of-bounds value as a flag.
		return;
	}
	// Keep this in sync with GetKeyJoyState().
	// See GetKeyJoyState() for more comments about the following lines.
	JoyControls joy;
	int joystick_id;
	vk_type vk = TextToVK(key_name);
	if (!vk)
	{
		aResultToken.symbol = SYM_STRING; // ScriptGetJoyState() also requires that this be initialized.
		if (   !(joy = (JoyControls)ConvertJoy(key_name, &joystick_id))   )
			aResultToken.marker = "";
		else
		{
			// The following must be set for ScriptGetJoyState():
			aResultToken.marker = aResultToken.buf; // If necessary, it will be moved to a persistent memory location by our caller.
			ScriptGetJoyState(joy, joystick_id, aResultToken, true);
		}
		return;
	}
	// Since above didn't return: There is a virtual key (not a joystick control).
	char *mode, mode_buf[MAX_FORMATTED_NUMBER_LENGTH + 1];
	if (aParamCount > 1)
	{
		if (   !(mode = ExprTokenToString(*aParam[1], mode_buf))   ) // Not an operand.  Haven't found a way to produce this situation yet, but safe to assume it's possible.
		{
			aResultToken.value_int64 = -1; // Store out-of-bounds value as a flag.
			return;
		}
	}
	else
		mode = "";
	KeyStateTypes key_state_type;
	switch (toupper(*mode)) // Second parameter.
	{
	case 'T': key_state_type = KEYSTATE_TOGGLE; break; // Whether a toggleable key such as CapsLock is currently turned on.
	case 'P': key_state_type = KEYSTATE_PHYSICAL; break; // Physical state of key.
	default: key_state_type = KEYSTATE_LOGICAL;
	}
	// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.
	aResultToken.value_int64 = ScriptGetKeyState(vk, key_state_type); // 1 for down and 0 for up.
}



void BIF_VarSetCapacity(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// Returns: The variable's new capacity.
// Parameters:
// 1: Target variable (unquoted).
// 2: Requested capacity.
// 3: Byte-value to fill the variable with (e.g. 0 to have the same effect as ZeroMemory).
{
	// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.
	aResultToken.value_int64 = 0; // Set default. In spite of being ambiguous with the result of Free(), 0 seems a little better than -1 since it indicates "no capacity" and is also equal to "false" for easy use in expressions.
	if (aParam[0]->symbol == SYM_VAR)
	{
		Var &var = *aParam[0]->var; // For performance and convenience.
		if (var.Type() == VAR_NORMAL) // Don't allow resizing/reporting of built-in variables since Assign() isn't designed to work on them, and resizing the clipboard's memory area would serve no purpose and isn't intended for this anyway.
		{
			if (aParamCount > 1) // Second parameter is present.
			{
				VarSizeType new_capacity = (VarSizeType)ExprTokenToInt64(*aParam[1]);
				if (new_capacity)
				{
					var.Assign(NULL, new_capacity, false, true); // This also destroys the variables contents.
					VarSizeType capacity;
					if (aParamCount > 2 && (capacity = var.Capacity()) > 1) // Third parameter is present and var has enough capacity to make FillMemory() meaningful.
					{
						--capacity; // Convert to script-POV capacity. To avoid underflow, do this only now that Capacity() is known not to be zero.
						// The following uses capacity-1 because the last byte of a variable should always
						// be left as a binary zero to avoid crashes and problems due to unterminated strings.
						// In other words, a variable's usable capacity from the script's POV is always one
						// less than its actual capacity:
						BYTE fill_byte = (BYTE)ExprTokenToInt64(*aParam[2]); // For simplicity, only numeric characters are supported, not something like "a" to mean the character 'a'.
						char *contents = var.Contents();
						FillMemory(contents, capacity, fill_byte); // Last byte of variable is always left as a binary zero.
						contents[capacity] = '\0'; // Must terminate because nothing else is explicitly reponsible for doing it.
						var.Length() = fill_byte ? capacity : 0; // Length is same as capacity unless fill_byte is zero.
					}
					else
						// By design, Assign() has already set the length of the variable to reflect new_capacity.
						// This is not what is wanted in this case since it should be truly empty.
						var.Length() = 0;
				}
				else // ALLOC_SIMPLE, due to its nature, will not actually be freed, which is documented.
					var.Free();
			}
			//else the var is not altered; instead, the current capacity is reported, which seems more intuitive/useful than having it do a Free().
			aResultToken.value_int64 = var.Capacity(); // Don't subtract 1 here in lieu doing it below (avoids underflow).
			if (aResultToken.value_int64)
				--aResultToken.value_int64; // Omit the room for the zero terminator since script capacity is defined as length vs. size.
		}
	}
}



void BIF_FileExist(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	aResultToken.symbol = SYM_STRING;
	char *filename, filename_buf[MAX_FORMATTED_NUMBER_LENGTH + 1];
	if (   !(filename = ExprTokenToString(*aParam[0], filename_buf))   )
	{
		// Not an operand.  Haven't found a way to produce this situation yet, but safe to assume it's possible.
		aResultToken.marker = "";
		return;
	}
	aResultToken.marker = aResultToken.buf; // If necessary, it will be moved to a persistent memory location by our caller.
	DWORD attr;
	if (DoesFilePatternExist(filename, &attr))
	{
		// Yield the attributes of the first matching file.  If not match, yield an empty string.
		// This relies upon the fact that a file's attributes are never legitimately zero, which
		// seems true but in case it ever isn't, this forces a non-empty string be used.
		if (!attr) // Seems impossible, but the use of 0xFFFFFFFF vs. 0 as "invalid" indicates otherwise.
		{
			// File exists but has no attributes!  Use a placeholder so that any expression will
			// see the result as "true" (i.e. because the file does exist):
			aResultToken.marker[0] = 'X'; // Some arbirary letter so that it's seen as "true"; letter mustn't be reserved by RASHNDOCT.
			aResultToken.marker[1] = '\0';
		}
		else
			FileAttribToStr(aResultToken.marker, attr);
	}
	else // Empty string is the indicator of "not found" (seems more consistent than using an integer 0, since caller might rely on it being SYM_STRING).
		*aResultToken.marker = '\0';
}



void BIF_WinExistActive(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	char *bif_name = aResultToken.marker;  // Save this early for maintainability (it is the name of the function, provided by the caller).
	aResultToken.symbol = SYM_STRING; // Returns a string to preserve hex format.
	char *param[4], param_buf[4][MAX_FORMATTED_NUMBER_LENGTH + 1];
	for (int j = 0; j < 4; ++j) // For each formal parameter, including optional ones.
	{
		if (j >= aParamCount) // No actual to go with it (should be possible only if the parameter is optional or has a default value).
		{
			param[j] = "";
			continue;
		}
		// Otherwise, assign actual parameter's value to the formal parameter.
		// The stack can contain both generic and specific operands.  Specific operands were
		// evaluated by a previous iteration of this section.  Generic ones were pushed as-is
		// onto the stack by a previous iteration.
		if (   !(param[j] = ExprTokenToString(*aParam[j], param_buf[j]))   )
		{
			// Not an operand.  Haven't found a way to produce this situation yet, but safe to assume it's possible.
			aResultToken.marker = "";
			return;
		}
	}

	// Should be called the same was as ACT_IFWINEXIST and ACT_IFWINACTIVE:
	HWND found_hwnd = (toupper(bif_name[3]) == 'E') // Win[E]xist.
		? WinExist(g, param[0], param[1], param[2], param[3], false, true)
		: WinActive(g, param[0], param[1], param[2], param[3], true);
	aResultToken.marker = aResultToken.buf;
	aResultToken.marker[0] = '0';
	aResultToken.marker[1] = 'x';
	_ui64toa((unsigned __int64)found_hwnd, aResultToken.marker + 2, 16); // If necessary, it will be moved to a persistent memory location by our caller.
}



void BIF_Round(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	// See TRANS_CMD_ROUND for details.
	int param2;
	double multiplier;
	if (aParamCount > 1)
	{
		param2 = (int)ExprTokenToInt64(*aParam[1]);
		multiplier = qmathPow(10, param2);
	}
	else // Omitting the parameter is the same as explicitly specifying 0 for it.
	{
		param2 = 0;
		multiplier = 1;
	}
	double value = ExprTokenToDouble(*aParam[0]);
	if (value >= 0.0)
		aResultToken.value_double = qmathFloor(value * multiplier + 0.5) / multiplier;
	else
		aResultToken.value_double = qmathCeil(value * multiplier - 0.5) / multiplier;
	// If incoming value is an integer, it seems best for flexibility to convert it to a
	// floating point number whenever the second param is >0.  That way, it can be used
	// to "cast" integers into floats.  Conversely, it seems best to yield an integer
	// whenever the second param is <=0 or omitted.
	if (param2 > 0)
		aResultToken.symbol = SYM_FLOAT; // aResultToken.value_double already contains the result.
	else
		// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.
		aResultToken.value_int64 = (__int64)aResultToken.value_double;
}



void BIF_FloorCeil(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// Probably saves little code size to merge extremely short/fast functions, hence FloorCeil.
// Floor() rounds down to the nearest integer; that is, to the integer that lies to the left on the
// number line (this is not the same as truncation because Floor(-1.2) is -2, not -1).
// Ceil() rounds up to the nearest integer; that is, to the integer that lies to the right on the number line.
{
	// The code here is similar to that in TRANS_CMD_FLOOR/CEIL, so maintain them together.
	// The qmath routines are used because Floor() and Ceil() are deceptively difficult to implement in a way
	// that gives the correct result in all permutations of the following:
	// 1) Negative vs. positive input.
	// 2) Whether or not the input is already an integer.
	// Therefore, do not change this without conduction a thorough test.
	double x = ExprTokenToDouble(*aParam[0]);
	x = (toupper(aResultToken.marker[0]) == 'F') ? qmathFloor(x) : qmathCeil(x);
	// Fix for v1.0.40.05: For some inputs, qmathCeil/Floor yield a number slightly to the left of the target
	// integer, while for others they yield one slightly to the right.  For example, Ceil(62/61) and Floor(-4/3)
	// yield a double that would give an incorrect answer if it were simply truncated to an integer via
	// type casting.  The below seems to fix this without breaking the answers for other inputs (which is
	// surprisingly harder than it seemed).
	aResultToken.value_int64 = (__int64)(x + (x > 0 ? 0.2 : -0.2));
	// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.
}



void BIF_Mod(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	// Load-time validation has already ensured there are exactly two parameters.
	// "Cast" each operand to Int64/Double depending on whether it has a decimal point.
	if (!ExprTokenToDoubleOrInt(*aParam[0]) || !ExprTokenToDoubleOrInt(*aParam[1])) // Non-operand or non-numeric string.
	{
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = "";
		return;
	}
	if (aParam[0]->symbol == SYM_INTEGER && aParam[1]->symbol == SYM_INTEGER)
	{
		if (!aParam[1]->value_int64) // Divide by zero.
		{
			aResultToken.symbol = SYM_STRING;
			aResultToken.marker = "";
		}
		else
			// For performance, % is used vs. qmath for integers.
			// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.
			aResultToken.value_int64 = aParam[0]->value_int64 % aParam[1]->value_int64;
	}
	else // At least one is a floating point number.
	{
		double dividend = ExprTokenToDouble(*aParam[0]);
		double divisor = ExprTokenToDouble(*aParam[1]);
		if (divisor == 0.0) // Divide by zero.
		{
			aResultToken.symbol = SYM_STRING;
			aResultToken.marker = "";
		}
		else
		{
			aResultToken.symbol = SYM_FLOAT;
			aResultToken.value_double = qmathFmod(dividend, divisor);
		}
	}
}



void BIF_Abs(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	// Unlike TRANS_CMD_ABS, which removes the minus sign from the string if it has one,
	// this is done in a more traditional way.  It's hard to imagine needing the minus
	// sign removal method here since a negative hex literal such as -0xFF seems too rare
	// to worry about.  One additional reason not to remove minus signs from strings is
	// that it might produce inconsistent results depending on whether the operand is
	// generic (SYM_OPERAND) and numeric.  In other words, abs() shouldn't treat a
	// sub-expression differently than a numeric literal.
	aResultToken = *aParam[0]; // Structure/union copy.
	// v1.0.40.06: ExprTokenToDoubleOrInt() and here has been fixed to set proper result to be empty string
	// when the incoming parameter is non-numeric.
	if (!ExprTokenToDoubleOrInt(aResultToken)) // "Cast" token to Int64/Double depending on whether it has a decimal point.
		// Non-operand or non-numeric string. ExprTokenToDoubleOrInt() has already set the token to be an
		// empty string for us.
		return;
	if (aResultToken.symbol == SYM_INTEGER)
	{
		// The following method is used instead of __abs64() to allow linking against the multi-threaded
		// DLLs (vs. libs) if that option is ever used (such as for a minimum size AutoHotkeySC.bin file).
		// It might be somewhat faster than __abs64() anyway, unless __abs64() is a macro or inline or something.
		if (aResultToken.value_int64 < 0)
			aResultToken.value_int64 = -aResultToken.value_int64;
	}
	else // Must be SYM_FLOAT due to the conversion above.
		aResultToken.value_double = qmathFabs(aResultToken.value_double);
}



void BIF_Sin(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	aResultToken.symbol = SYM_FLOAT;
	aResultToken.value_double = qmathSin(ExprTokenToDouble(*aParam[0]));
}



void BIF_Cos(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	aResultToken.symbol = SYM_FLOAT;
	aResultToken.value_double = qmathCos(ExprTokenToDouble(*aParam[0]));
}



void BIF_Tan(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	aResultToken.symbol = SYM_FLOAT;
	aResultToken.value_double = qmathTan(ExprTokenToDouble(*aParam[0]));
}



void BIF_ASinACos(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	double value = ExprTokenToDouble(*aParam[0]);
	if (value > 1 || value < -1) // ASin and ACos aren't defined for other values.
	{
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = "";
	}
	else
	{
		aResultToken.symbol = SYM_FLOAT;
		// Below: marker contains either "ASin" or "ACos"
		aResultToken.value_double = (toupper(aResultToken.marker[1]) == 'S') ? qmathAsin(value) : qmathAcos(value);
	}
}



void BIF_ATan(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	aResultToken.symbol = SYM_FLOAT;
	aResultToken.value_double = qmathAtan(ExprTokenToDouble(*aParam[0]));
}



void BIF_Exp(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	aResultToken.symbol = SYM_FLOAT;
	aResultToken.value_double = qmathExp(ExprTokenToDouble(*aParam[0]));
}



void BIF_SqrtLogLn(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	double value = ExprTokenToDouble(*aParam[0]);
	if (value < 0) // Result is undefined in these cases, so make blank to indicate.
	{
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker = "";
	}
	else
	{
		aResultToken.symbol = SYM_FLOAT;
		switch (toupper(aResultToken.marker[1]))
		{
		case 'Q': // S[q]rt
			aResultToken.value_double = qmathSqrt(value);
			break;
		case 'O': // L[o]g
			aResultToken.value_double = qmathLog10(value);
			break;
		default: // L[n]
			aResultToken.value_double = qmathLog(value);
		}
	}
}



void BIF_OnMessage(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// Returns: An empty string on failure or the name of a function (depends on mode) on success.
// Parameters:
// 1: Message number to monitor.
// 2: Name of the function that will monitor the message.
// 3: (FUTURE): A flex-list of space-delimited option words/letters.
{
	char *buf = aResultToken.buf; // Must be saved early since below overwrites the union (better maintainability too).
	// Set default result in case of early return; a blank value:
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = "";

	// Load-time validation has ensured there's at least one parameters for use here:
	UINT specified_msg = (UINT)ExprTokenToInt64(*aParam[0]); // Parameter #1

	Func *func = NULL;           // Set defaults.
	bool mode_is_delete = false; //
	if (aParamCount > 1) // Parameter #2 is present.
	{
		char *func_name = ExprTokenToString(*aParam[1], buf); // Resolve parameter #2.
		if (*func_name)
		{
			if (   !(func = g_script.FindFunc(func_name))   )
				return; // Yield the default return value set earlier.
			// If too many formal parameters or any are ByRef/optional, indicate failure.
			// This helps catch bugs in scripts that are assigning the wrong function to a monitor.
			// It also preserves additional parameters for possible future use (i.e. avoids breaking
			// existing scripts if more formal parameters are supported in a future version).
			if (func->mIsBuiltIn || func->mParamCount > 4 || func->mMinParams < func->mParamCount) // Too many params, or some are optional.
				return; // Yield the default return value set earlier.
			for (int i = 0; i < func->mParamCount; ++i) // Check if any formal parameters are ByRef.
				if (func->mParam[i].var->IsByRef())
					return; // Yield the default return value set earlier.
		}
		else // Explicitly blank function name ("") means delete this item.  By contrast, an omitted second parameter means "give me current function of this message".
			mode_is_delete = true;
	}

	// If this is the first use of the g_MsgMonitor array, create it now rather than later to reduce code size
	// and help the maintainability of sections further below. The following relies on short-circuit boolean order:
	if (!g_MsgMonitor && !(g_MsgMonitor = (MsgMonitorStruct *)malloc(sizeof(MsgMonitorStruct) * MAX_MSG_MONITORS)))
		return; // Yield the default return value set earlier.

	// Check if this message already exists in the array:
	int msg_index;
	for (msg_index = 0; msg_index < g_MsgMonitorCount; ++msg_index)
		if (g_MsgMonitor[msg_index].msg == specified_msg)
			break;
	bool item_already_exists = (msg_index < g_MsgMonitorCount);
	MsgMonitorStruct &monitor = g_MsgMonitor[msg_index == MAX_MSG_MONITORS ? 0 : msg_index]; // The 0th item is just a placeholder.

	if (item_already_exists)
	{
		// In all cases, yield the OLD function's name as the return value:
		strcpy(buf, monitor.func->mName); // Caller has ensured that buf large enough to support max function name.
		aResultToken.marker = buf;
		if (mode_is_delete)
		{
			// The msg-monitor is deleted from the array for two reasons:
			// 1) It improves performance because every incoming message for the app now needs to be compared
			//    to one less filter. If the count will now be zero, performance is improved even more because
			//    the overhead of the call to MsgMonitor() is completely avoided for every incoming message.
			// 2) It conserves space in the array in a situation where the script creates hundreds of
			//    msg-monitors and then later deletes them, then later creates hundreds of filters for entirely
			//    different message numbers.
			// The main disadvantage to deleting message filters from the array is that the deletion might
			// occur while the monitor is currently running, which requires more complex handling within
			// MsgMonitor() (see its comments for details).
			--g_MsgMonitorCount;  // Must be done prior to the below.
			if (msg_index < g_MsgMonitorCount) // An element other than the last is being removed. Shift the array to cover/delete it.
				MoveMemory(g_MsgMonitor+msg_index, g_MsgMonitor+msg_index+1, sizeof(MsgMonitorStruct)*(g_MsgMonitorCount-msg_index));
			return;
		}
		if (aParamCount < 2) // Single-parameter mode: Report existing item's function name.
			return; // Everything was already set up above to yield the proper return value.
		// Otherwise, an existing item is being assigned a new function (the function has already
		// been verified valid above). Continue on to update this item's attributes.
	}
	else // This message doesn't exist in array yet.
	{
		if (mode_is_delete || aParamCount < 2) // Delete or report function-name of a non-existent item.
			return; // Yield the default return value set earlier (an empty string).
		// Since above didn't return, the message is to be added as a new element. The above already
		// verified that func is not NULL.
		if (msg_index == MAX_MSG_MONITORS) // No room in array.
			return; // Indicate failure by yielding the default return value set earlier.
		// Otherwise, the message is to be added, so increment the total:
		++g_MsgMonitorCount;
		strcpy(buf, func->mName); // Yield the NEW name as an indicator of success. Caller has ensured that buf large enough to support max function name.
		aResultToken.marker = buf;
		// Continue on to the update-or-create logic below.
	}

	// Since above didn't return, above has ensured that msg_index is the index of the existing or new
	// MsgMonitorStruct in the array.  In addition, it has set the proper return value for us.
	// Regardless of whether this is an update or creation, update all the struct attributes:
	monitor.msg = specified_msg;
	monitor.func = func;
	if (!item_already_exists) // Reset label_is_running only for new items since existing items might currently be running.
		monitor.label_is_running = false;
}



void BIF_LV_GetNextOrCount(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// Returns: The index of the found item, or 0 on failure.
// Parameters:
// 1: Starting index (one-based when it comes in).  If absent, search starts at the top.
// 2: Options string.
// 3: (FUTURE): Possible for use with LV_FindItem (though I think it can only search item text, not subitem text).
{
	bool mode_is_count = toupper(aResultToken.marker[6] == 'C'); // Marker contains the function name. LV_Get[C]ount.
	char *buf = aResultToken.buf; // Must be saved early since below overwrites the union (better maintainability too).
	aResultToken.value_int64 = 0;
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Window doesn't exist.
	// Control doesn't exist (i.e. no ListView in window).
	// Item not found in ListView.
	if (!g_gui[g.GuiDefaultWindowIndex])
		return;
	GuiType &gui = *g_gui[g.GuiDefaultWindowIndex]; // Always operate on thread's default window to simplify the syntax.
	if (!gui.mCurrentListView)
		return;
	HWND control_hwnd = gui.mCurrentListView->hwnd;

	char *options;
	if (mode_is_count)
	{
		options = (aParamCount > 0) ? omit_leading_whitespace(ExprTokenToString(*aParam[0], buf)) : "";
		if (*options)
		{
			if (toupper(*options) == 'S')
				aResultToken.value_int64 = SendMessage(control_hwnd, LVM_GETSELECTEDCOUNT, 0, 0);
			else if (!strnicmp(options, "Col", 3)) // "Col" or "Column". Don't allow "C" by itself, so that "Checked" can be added in the future.
				aResultToken.value_int64 = gui.mCurrentListView->union_lv_attrib->col_count;
			//else some unsupported value, leave aResultToken.value_int64 set to zero to indicate failure.
		}
		else
			aResultToken.value_int64 = SendMessage(control_hwnd, LVM_GETITEMCOUNT, 0, 0);
		return;
	}
	// Since above didn't return, this is GetNext() mode.

	int index = (int)((aParamCount > 0) ? ExprTokenToInt64(*aParam[0]) - 1 : -1); // -1 to convert to zero-based.
	// For flexibility, allow index to be less than -1 to avoid first-iteration complications in script loops
	// (such as when deleting rows, which shifts the row index upward, require the search to resume at
	// the previously found index rather than the row after it).  However, reset it to -1 to ensure
	// proper return values from the API in the "find checked item" mode used below.
	if (index < -1)
		index = -1;  // Signal it to start at the top.

	if (aParamCount < 2)
		options = "";
	else
		if (   !(options = ExprTokenToString(*aParam[1], buf))   ) // Not an operand.  Haven't found a way to produce this situation yet, but safe to assume it's possible.
			return; // Due to rarity of this condition, keep 0/false value as the result.

	// For performance, decided to always find next selected item when the "C" option hasn't been specified,
	// even when the checkboxes style is in effect.  Otherwise, would have to fetch and check checkbox style
	// bit for each call, which would slow down this heavily-called function.

	char first_char = toupper(*omit_leading_whitespace(options));
	// To retain compatibility in the future, also allow "Check(ed)" and "Focus(ed)" since any word that
	// starts with C or F is already supported.

	switch(first_char)
	{
	case '\0': // Listed first for performance.
	case 'F':
		aResultToken.value_int64 = ListView_GetNextItem(control_hwnd, index
			, first_char ? LVNI_FOCUSED : LVNI_SELECTED) + 1; // +1 to convert to 1-based.
		break;
	case 'C': // Checkbox: Find checked items. For performance assume that the control really has checkboxes.
		int item_count = ListView_GetItemCount(control_hwnd);
		for (int i = index + 1; i < item_count; ++i) // Start at index+1 to omit the first item from the search (for consistency with the other mode above).
			if (ListView_GetCheckState(control_hwnd, i)) // Item's box is checked.
			{
				aResultToken.value_int64 = i + 1; // +1 to convert from zero-based to one-based.
				return;
			}
		// Since above didn't return, no match found.  The 0/false value previously set as the default is retained.
		break;
	}
}



void BIF_LV_GetText(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: Output variable (doing it this way allows success/fail return value to more closely mirror the API and
//    simplifies the code since there is currently no easy means of passing back large strings to our caller.
// 2: Row index (one-based when it comes in).
// 3: Column index (one-based when it comes in).
{
	aResultToken.value_int64 = 0;
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Window doesn't exist.
	// Control doesn't exist (i.e. no ListView in window).
	// Item not found in ListView.
	// And others.
	if (!g_gui[g.GuiDefaultWindowIndex])
		return;
	GuiType &gui = *g_gui[g.GuiDefaultWindowIndex]; // Always operate on thread's default window to simplify the syntax.
	if (!gui.mCurrentListView)
		return;
	// Caller has ensured there is at least two parameters:
	if (aParam[0]->symbol != SYM_VAR) // No output variable.  Supporting a NULL for the purpose of checking for the existence of a cell seems too rarely needed.
		return;

	// Caller has ensured there is at least two parameters.
	int row_index = (int)ExprTokenToInt64(*aParam[1]) - 1; // -1 to convert to zero-based.
	// If parameter 3 is omitted, default to the first column (index 0):
	int col_index = (aParamCount > 2) ? (int)ExprTokenToInt64(*aParam[2]) - 1 : 0; // -1 to convert to zero-based.
	if (row_index < -1 || col_index < 0) // row_index==-1 is reserved to mean "get column heading's text".
		return;

	Var &output_var = *aParam[0]->var; // It was already ensured higher above that symbol==SYM_VAR.
	char buf[LV_TEXT_BUF_SIZE];

	if (row_index == -1) // Special mode to get column's text.
	{
		LVCOLUMN lvc;
		lvc.cchTextMax = LV_TEXT_BUF_SIZE - 1;  // See notes below about -1.
		lvc.pszText = buf;
		lvc.mask = LVCF_TEXT;
		if (aResultToken.value_int64 = SendMessage(gui.mCurrentListView->hwnd, LVM_GETCOLUMN, col_index, (LPARAM)&lvc)) // Assign.
			output_var.Assign(lvc.pszText); // See notes below about why pszText is used instead of buf (might apply to this too).
		else // On failure, it seems best to also clear the output var for better consistency and in case the script doesn't check the return value.
			output_var.Assign();
	}
	else // Get row's indicated item or subitem text.
	{
		LVITEM lvi;
		// Subtract 1 because of that nagging doubt about size vs. length. Some MSDN examples subtract one, such as
		// TabCtrl_GetItem()'s cchTextMax:
		lvi.cchTextMax = LV_TEXT_BUF_SIZE - 1; // Note that LVM_GETITEM doesn't update this member to reflect the new length.
		lvi.pszText = buf;
		lvi.mask = LVIF_TEXT;
		lvi.iItem = row_index;
		lvi.iSubItem = col_index; // Which field to fetch.  If it's zero, the item vs. subitem will be fetched.
		// Unlike LVM_GETITEMTEXT, LVM_GETITEM indicates success or failure, which seems more useful/preferable
		// as a return value since a text length of zero would be ambiguous: could be an empty field or a failure.
		if (aResultToken.value_int64 = SendMessage(gui.mCurrentListView->hwnd, LVM_GETITEM, 0, (LPARAM)&lvi)) // Assign
			// Must use lvi.pszText vs. buf because MSDN says: "Applications should not assume that the text will
			// necessarily be placed in the specified buffer. The control may instead change the pszText member
			// of the structure to point to the new text rather than place it in the buffer."
			output_var.Assign(lvi.pszText);
		else // On failure, it seems best to also clear the output var for better consistency and in case the script doesn't check the return value.
			output_var.Assign();
	}
}



void BIF_LV_AddInsertModify(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: For Add(), this is the options.  For Insert/Modify, it's the row index (one-based when it comes in).
// 2: For Add(), this is the first field's text.  For Insert/Modify, it's the options.
// 3 and beyond: Additional field text.
// In Add/Insert mode, if there are no text fields present, a blank for is appended/inserted.
{
	char mode = toupper(aResultToken.marker[3]); // Marker contains the function name. e.g. LV_[I]nsert.
	char *buf = aResultToken.buf; // Must be saved early since below overwrites the union (better maintainability too).
	aResultToken.value_int64 = 0;
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Window doesn't exist.
	// Control doesn't exist (i.e. no ListView in window).
	// And others as shown below.

	int index;
	if (mode == 'A') // For Add mode (A), use INT_MAX as a signal to append the item rather than inserting it.
	{
		index = INT_MAX;
		mode = 'I'; // Add has now been set up to be the same as insert, so change the mode to simplify other things.
	}
	else // Insert or Modify: the target row-index is their first parameter, which load-time has ensured is present.
	{
		index = (int)ExprTokenToInt64(*aParam[0]) - 1; // -1 to convert to zero-based.
		if (index < -1 || (mode != 'M' && index < 0)) // Allow -1 to mean "all rows" when in modify mode.
			return;
		++aParam;  // Remove the first parameter from further consideration to make Insert/Modify symmetric with Add.
		--aParamCount;
	}

	if (!g_gui[g.GuiDefaultWindowIndex])
		return;
	GuiType &gui = *g_gui[g.GuiDefaultWindowIndex]; // Always operate on thread's default window to simplify the syntax.
	if (!gui.mCurrentListView)
		return;
	GuiControlType &control = *gui.mCurrentListView;

	char *options;
	if (aParamCount > 0)
	{
		if (   !(options = ExprTokenToString(*aParam[0], buf))   ) // Not an operand.  Haven't found a way to produce this situation yet, but safe to assume it's possible.
			return; // Due to rarity of this condition, keep 0/false value as the result.
	}
	else  // No options parameter is present.
		options = "";

	bool is_checked = false;  // Checkmark.
	int col_start_index = 0;
	LVITEM lvi;
	lvi.mask = LVIF_STATE; // LVIF_STATE: state member is valid, but only to the extent that corresponding bits are set in stateMask (the rest will be ignored).
	lvi.stateMask = 0;
	lvi.state = 0;

	// Parse list of space-delimited options:
	char *next_option, *option_end, orig_char;
	bool adding; // Whether this option is beeing added (+) or removed (-).

	for (next_option = options; *next_option; next_option = omit_leading_whitespace(option_end))
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

		if (!strnicmp(next_option, "Select", 6)) // Could further allow "ed" suffix by checking for that inside, but "Selected" is getting long so it doesn't seem something many would want to use.
		{
			next_option += 6;
			// If it's Select0, invert the mode to become "no select". This allows a boolean variable
			// to be more easily applied, such as this expression: "Check" . VarContainingState
			if (*next_option && !ATOI(next_option))
				adding = !adding;
			// Another reason for not having "Select" imply "Focus" by default is that it would probably
			// reduce performance when selecting all or a large number of rows.
			// Because a row might or might not have focus, the script may wish to retain its current
			// focused state.  For this reason, "select" does not imply "focus", which allows the
			// LVIS_FOCUSED bit to be omitted from the stateMask, which in turn retains the current
			// focus-state of the row rather than disrupting it.
			lvi.stateMask |= LVIS_SELECTED;
			if (adding)
				lvi.state |= LVIS_SELECTED;
			//else removing, so the presence of LVIS_SELECTED in the stateMask above will cause it to be de-selected.
		}
		else if (!strnicmp(next_option, "Focus", 5))
		{
			next_option += 5;
			if (*next_option && !ATOI(next_option)) // If it's Focus0, invert the mode to become "no focus".
				adding = !adding;
			lvi.stateMask |= LVIS_FOCUSED;
			if (adding)
				lvi.state |= LVIS_FOCUSED;
			//else removing, so the presence of LVIS_FOCUSED in the stateMask above will cause it to be de-focused.
		}
		else if (!strnicmp(next_option, "Check", 5))
		{
			// The rationale for not checking for an optional "ed" suffix here and incrementing next_option by 2
			// is that: 1) It would be inconsistent with the lack of support for "selected" (see reason above);
			// 2) Checkboxes in a ListView are fairly rarely used, so code size reduction might be more important.
			next_option += 5;
			if (*next_option && !ATOI(next_option)) // If it's Check0, invert the mode to become "unchecked".
				adding = !adding;
			lvi.stateMask |= LVIS_STATEIMAGEMASK;
			lvi.state |= INDEXTOSTATEIMAGEMASK(adding ? 2 : 1); // The #1 image is "unchecked" and the #2 is "checked".
			is_checked = adding;
		}
		else if (!strnicmp(next_option, "Col", 3))
		{
			if (adding)
			{
				col_start_index = ATOI(next_option + 3) - 1; // The ability start start at a column other than 1 (i.e. subitem vs. item).
				if (col_start_index < 0)
					col_start_index = 0;
			}
		}
		else if (!strnicmp(next_option, "Icon", 4))
		{
			// Testing shows that there is no way to avoid having an item icon in report view if the
			// ListView has an associated small-icon ImageList (well, perhaps you could have it show
			// a blank square by specifying an invalid icon index, but that doesn't seem useful).
			// If LVIF_IMAGE is entirely omitted when adding and item/row, the item will take on the
			// first icon in the list.  This is probably by design because the control wants to make
			// each item look consistent by indenting its first field by a certain amount for the icon.
			if (adding)
			{
				lvi.mask |= LVIF_IMAGE;
				lvi.iImage = ATOI(next_option + 4) - 1;  // -1 to convert to zero-based.
			}
			//else removal of icon currently not supported (see comment above), so do nothing in order
			// to reserve "-Icon" in case a future way can be found to do it.
		}

		// If the item was not handled by the above, ignore it because it is unknown.
		*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.
	}

	// More maintainable and performs better to have a separate struct for subitems vs. items.
	LVITEM lvi_sub;
	// Ensure mask is pure to avoid giving it any excuse to fail due to the fact that
	// "You cannot set the state or lParam members for subitems."
	lvi_sub.mask = LVIF_TEXT;

	int i, j, rows_to_change;
	if (index == -1) // Modify all rows (above has ensured that this is only happens in modify-mode).
	{
		rows_to_change = ListView_GetItemCount(control.hwnd);
		lvi.iItem = 0;
	}
	else // Modify or insert a single row.  Set it up for the loop to perform exactly one iteration.
	{
		rows_to_change = 1;
		lvi.iItem = index; // Which row to operate upon.  This can be a huge number such as 999999 if the caller wanted to append vs. insert.
	}
	lvi.iSubItem = 0;  // Always zero to operate upon the item vs. sub-item (subitems have their own LVITEM struct).
	aResultToken.value_int64 = 1; // Set default from this point forward to be true/success. It will be overridden in insert mode to be the index of the new row.

	for (j = 0; j < rows_to_change; ++j, ++lvi.iItem) // ++lvi.iItem because if the loop has more than one iteration, by definition it is modifying all rows starting at 0.
	{
		if (aParamCount > 1 && col_start_index == 0) // 2nd parameter: item's text (first field) is present, so include that when setting the item.
		{
			if (lvi.pszText = ExprTokenToString(*aParam[1], buf)) // Fairly low-overhead, so called every iteration for simplicity (so that buf can be used for both items and subitems).
				lvi.mask |= LVIF_TEXT;
			//else not an operand, so don't add the mask.  Haven't found a way to produce this situation yet, but safe to assume it's possible.
		}
		if (mode == 'I') // Insert or Add.
		{
			// Note that ListView_InsertItem() will append vs. insert if the index is too large, in which case
			// it returns the items new index (which will be the last item in the list unless the control has
			// auto-sort style).
			// Below uses +1 to convert from zero-based to 1-based.  This also converts a failure result of -1 to 0.
			if (   !(aResultToken.value_int64 = ListView_InsertItem(control.hwnd, &lvi) + 1)   )
				return; // Since item can't be inserted, no reason to try attaching any subitems to it.
			// Update iItem with the actual index assigned to the item, which might be different than the
			// specified index if the control has an auto-sort style in effect.  This new iItem value
			// is used for ListView_SetCheckState() and for the attaching of any subitems to this item.
			lvi_sub.iItem = (int)aResultToken.value_int64 - 1;  // -1 to convert back to zero-based.
			// For add/insert (but not modify), testing shows that checkmark must be added only after
			// the item has been inserted rather than provided in the lvi.state/stateMask fields.
			// MSDN confirms this by saying "When an item is added with [LVS_EX_CHECKBOXES],
			// it will always be set to the unchecked state [ignoring any value placed in bits
			// 12 through 15 of the state member]."
			if (is_checked)
				ListView_SetCheckState(control.hwnd, lvi_sub.iItem, TRUE); // TRUE = Check the row's checkbox.
				// Note that 95/NT4 systems that lack comctl32.dll 4.70+ distributed with MSIE 3.x
				// do not support LVS_EX_CHECKBOXES, so the above will have no effect for them.
		}
		else // Modify.
		{
			// Rather than trying to detect if anything was actually changed, this is called
			// unconditionally to simplify the code.
			// By design (to help catch script bugs), a failure here does not revert to append mode.
			if (!ListView_SetItem(control.hwnd, &lvi)) // Returns TRUE/FALSE.
				aResultToken.value_int64 = 0; // Indicate partial failure, but attempt to continue in case it failed for reason other than "row doesn't exist".
			lvi_sub.iItem = lvi.iItem; // In preparation for modifying any subitems that need it.
		}

		// For each remainining parameter, assign its text to a subitem.  
		for (lvi_sub.iSubItem = (col_start_index > 1) ? col_start_index : 1 // Start at the first subitem unless we were told to start at or after the third column.
			// "i" starts at 2 (the third parameter) unless col_start_index is greater than 0, in which case
			// it starts at 1 (the second parameter) because that parameter has not yet been assigned to anything:
			, i = 2 - (col_start_index > 0)
			; i < aParamCount
			; ++i, ++lvi_sub.iSubItem)
			if (lvi_sub.pszText = ExprTokenToString(*aParam[i], buf)) // Done every time through the outer loop since it's not high-overhead, and for code simplicity.
				if (!ListView_SetItem(control.hwnd, &lvi_sub) && mode != 'I') // Relies on short-circuit. Seems best to avoid loss of item's index in insert mode, since failure here should be rare.
					aResultToken.value_int64 = 0; // Indicate partial failure, but attempt to continue in case it failed for reason other than "row doesn't exist".
			// else not an operand, but it's simplest just to try to continue.
	} // outer for()

	// When the control has no rows, work around the fact that LVM_SETITEMCOUNT delivers less than 20%
	// of its full benefit unless done after the first row is added (at least on XP SP1).  A non-zero
	// row_count_hint tells us that this message should be sent after the row has been inserted/appended:
	if (control.union_lv_attrib->row_count_hint > 0 && mode == 'I')
	{
		SendMessage(control.hwnd, LVM_SETITEMCOUNT, control.union_lv_attrib->row_count_hint, 0); // Last parameter should be 0 for LVS_OWNERDATA (verified if you look at the definition of ListView_SetItemCount macro).
		control.union_lv_attrib->row_count_hint = 0; // Reset so that it only gets set once per request.
	}
}



void BIF_LV_Delete(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: Row index (one-based when it comes in).
{
	aResultToken.value_int64 = 0;
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Window doesn't exist.
	// Control doesn't exist (i.e. no ListView in window).
	// And others as shown below.

	if (!g_gui[g.GuiDefaultWindowIndex])
		return;
	GuiType &gui = *g_gui[g.GuiDefaultWindowIndex]; // Always operate on thread's default window to simplify the syntax.
	if (!gui.mCurrentListView)
		return;
	GuiControlType &control = *gui.mCurrentListView;

	if (aParamCount < 1)
	{
		aResultToken.value_int64 = SendMessage(control.hwnd, LVM_DELETEALLITEMS, 0, 0); // Returns TRUE/FALSE.
		return;
	}

	// Since above didn't return, there is a first paramter present.
	int index = (int)ExprTokenToInt64(*aParam[0]) - 1; // -1 to convert to zero-based.
	if (index > -1)
		aResultToken.value_int64 = SendMessage(control.hwnd, LVM_DELETEITEM, index, 0); // Returns TRUE/FALSE.
}



void BIF_LV_InsertModifyDeleteCol(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: Column index (one-based when it comes in).
// 2: String of options
// 3: New text of column
// There are also some special modes when only zero or one parameter is present, see below.
{
	char mode = toupper(aResultToken.marker[3]); // Marker contains the function name. LV_[I]nsertCol.
	char *buf = aResultToken.buf; // Must be saved early since below overwrites the union (better maintainability too).
	aResultToken.value_int64 = 0;
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Window doesn't exist.
	// Control doesn't exist (i.e. no ListView in window).
	// Column not found in ListView.

	if (!g_gui[g.GuiDefaultWindowIndex])
		return;
	GuiType &gui = *g_gui[g.GuiDefaultWindowIndex]; // Always operate on thread's default window to simplify the syntax.
	if (!gui.mCurrentListView)
		return;
	GuiControlType &control = *gui.mCurrentListView;
	lv_attrib_type &lv_attrib = *control.union_lv_attrib;

	int index;
	if (aParamCount > 0)
		index = (int)ExprTokenToInt64(*aParam[0]) - 1; // -1 to convert to zero-based.
	else // Zero parameters.  Load-time validation has ensured that the 'D' (delete) mode cannot have zero params.
	{
		if (mode == 'M')
		{
			if (GuiType::ControlGetListViewMode(control.hwnd) != LVS_REPORT)
				return; // And leave aResultToken.value_int64 at 0 to indicate failure.
			// Otherwise:
			aResultToken.value_int64 = 1; // Always successful (for consistency), regardless of what happens below.
			// v1.0.36.03: Don't attempt to auto-size the columns while the view is not report-view because
			// that causes any subsequent switch to the "list" view to be corrupted (invisible icons and items):
			for (int i = 0; ; ++i) // Don't limit it to lv_attrib.col_count in case script added extra columns via direct API calls.
				if (!ListView_SetColumnWidth(control.hwnd, i, LVSCW_AUTOSIZE)) // Failure means last column has already been processed.
					break; // Break vs. return in case the loop has zero iterations due to zero columns (not currently possible, but helps maintainability).
			return;
		}
		// Since above didn't return, mode must be 'I' (insert).
		index = lv_attrib.col_count; // When no insertion index was specified, append to the end of the list.
	}

	// Do this prior to checking if index is in bounds so that it can support columns beyond LV_MAX_COLUMNS:
	if (mode == 'D') // Delete a column.  In this mode, index parameter was made mandatory via load-time validation.
	{
		if (aResultToken.value_int64 = ListView_DeleteColumn(control.hwnd, index))  // Returns TRUE/FALSE.
		{
			// It's important to note that when the user slides columns around via drag and drop, the
			// column index as seen by the script is not changed.  This is fortunate because otherwise,
			// the lv_attrib.col array would get out of sync with the column indices.  Testing shows that
			// all of the following operations respect the original column index, regardless of where the
			// user may have moved the column physically: InsertCol, DeleteCol, ModifyCol.  Insert and Delete
			// shifts the indices of those columns that *originally* lay to the right of the affected column.
			if (lv_attrib.col_count > 0) // Avoid going negative, which would otherwise happen if script previously added columns by calling the API directly.
				--lv_attrib.col_count; // Must be done prior to the below.
			if (index < lv_attrib.col_count) // When a column other than the last was removed, adjust the array so that it stays in sync with actual columns.
				MoveMemory(lv_attrib.col+index, lv_attrib.col+index+1, sizeof(lv_col_type)*(lv_attrib.col_count-index));
		}
		return;
	}
	// Do this prior to checking if index is in bounds so that it can support columns beyond LV_MAX_COLUMNS:
	if (mode == 'M' && aParamCount < 2) // A single parameter is a special modify-mode to auto-size that column.
	{
		// v1.0.36.03: Don't attempt to auto-size the columns while the view is not report-view because
		// that causes any subsequent switch to the "list" view to be corrupted (invisible icons and items):
		if (GuiType::ControlGetListViewMode(control.hwnd) == LVS_REPORT)
			aResultToken.value_int64 = ListView_SetColumnWidth(control.hwnd, index, LVSCW_AUTOSIZE);
		//else leave aResultToken.value_int64 set to 0.
		return;
	}
	if (mode == 'I')
	{
		if (lv_attrib.col_count >= LV_MAX_COLUMNS) // No room to insert or append.
			return;
		if (index >= lv_attrib.col_count) // For convenience, fall back to "append" when index too large.
			index = lv_attrib.col_count;
	}
	//else do nothing so that modification and deletion of columns that were added via script's
	// direct calls to the API can sort-of work (it's documented in the help file that it's not supported,
	// since col-attrib array can get out of sync with actual columns that way).

	if (index < 0 || index >= LV_MAX_COLUMNS) // For simplicity, do nothing else if index out of bounds.
		return; // Avoid array under/overflow below.

	// It's done the following way so that when in insert-mode, if the column fails to be inserted, don't
	// have to remove the inserted array element from the lv_attrib.col array:
	lv_col_type temp_col = {0}; // Init unconditionally even though only needed for mode=='I'.
	lv_col_type &col = (mode == 'I') ? temp_col : lv_attrib.col[index]; // Done only after index has been confirmed in-bounds.

	// In addition to other reasons, must convert any numeric value to a string so that an isolated width is
	// recognized, e.g. LV_SetCol(1, old_width + 10):
	char *options;
	if (aParamCount > 1) // Second parameter is present.
	{
		if (   !(options = ExprTokenToString(*aParam[1], buf))   ) // Not an operand.  Haven't found a way to produce this situation yet, but safe to assume it's possible.
			return; // Due to rarity of this condition, keep 0/false value as the result.
	}
	else
		options = "";

	LVCOLUMN lvc;
	lvc.mask = LVCF_FMT;
	if (mode == 'M') // Fetch the current format so that it's possible to leave parts of it unaltered.
		ListView_GetColumn(control.hwnd, index, &lvc);
	else // Mode is "insert".
		lvc.fmt = 0;

	// Init defaults prior to parsing options:
	bool sort_now = false;
	int do_auto_size = (mode == 'I') ? LVSCW_AUTOSIZE_USEHEADER : 0;  // Default to auto-size for new columns.
	char sort_now_direction = 'A'; // Ascending.
	int new_justify = lvc.fmt & LVCFMT_JUSTIFYMASK; // Simplifies the handling of the justification bitfield.
	//lvc.iSubItem = 0; // Not necessary if the LVCF_SUBITEM mask-bit is absent.

	// Parse list of space-delimited options:
	char *next_option, *option_end, orig_char;
	bool adding; // Whether this option is beeing added (+) or removed (-).

	for (next_option = options; *next_option; next_option = omit_leading_whitespace(option_end))
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

		// For simplicity, the value of "adding" is ignored for this and the other number/alignment options.
		if (!stricmp(next_option, "Integer"))
		{
			// For simplicity, changing the col.type dynamically (since it's so rarely needed)
			// does not try to set up col.is_now_sorted_ascending so that the next click on the column
			// puts it into default starting order (which is ascending unless the Desc flag was originally
			// present).
			col.type = LV_COL_INTEGER;
			new_justify = LVCFMT_RIGHT;
		}
		else if (!stricmp(next_option, "Float"))
		{
			col.type = LV_COL_FLOAT;
			new_justify = LVCFMT_RIGHT;
		}
		else if (!stricmp(next_option, "Text")) // Seems more approp. name than "Str" or "String"
			// Since "Text" is so general, it seems to leave existing alignment (Center/Right) as it is.
			col.type = LV_COL_TEXT;

		// The following can exist by themselves or in conjunction with the above.  They can also occur
		// *after* one of the above words so that alignment can be used to override the default for the type;
		// e.g. "Integer Left" to have left-aligned integers.
		else if (!stricmp(next_option, "Right"))
			new_justify = adding ? LVCFMT_RIGHT : LVCFMT_LEFT;
		else if (!stricmp(next_option, "Center"))
			new_justify = adding ? LVCFMT_CENTER : LVCFMT_LEFT;
		else if (!stricmp(next_option, "Left")) // Supported so that existing right/center column can be changed back to left.
			new_justify = LVCFMT_LEFT; // The value of "adding" seems inconsequential so is ignored.

		else if (!stricmp(next_option, "Uni")) // Unidirectional sort (clicking the column will not invert to the opposite direction).
			col.unidirectional = adding;
		else if (!stricmp(next_option, "Desc")) // Make descending order the default order (applies to uni and first click of col for non-uni).
			col.prefer_descending = adding; // So that the next click will toggle to the opposite direction.
		else if (!stricmp(next_option, "Case"))
			col.case_sensitive = adding;

		else if (!strnicmp(next_option, "Sort", 4)) // This is done as an option vs. LV_SortCol/LV_Sort so that the column's options can be changed simultaneously with a "sort now" to refresh.
		{
			// Defer the sort until after all options have been parsed and applied.
			sort_now = true;
			if (!stricmp(next_option + 4, "Desc"))
				sort_now_direction = 'D'; // Descending.
		}
		else if (!stricmp(next_option, "NoSort")) // Called "NoSort" so that there's a way to enable and disable the setting via +/-.
			col.sort_disabled = adding;

		else if (!strnicmp(next_option, "Auto", 4)) // No separate failure result is reported for this item.
			// In case the mode is "insert", defer auto-width of column until col exists.
			do_auto_size = stricmp(next_option + 4, "Hdr") ? LVSCW_AUTOSIZE : LVSCW_AUTOSIZE_USEHEADER;

		else if (!strnicmp(next_option, "Icon", 4))
		{
			next_option += 4;
			if (!stricmp(next_option, "Right"))
			{
				if (adding)
					lvc.fmt |= LVCFMT_BITMAP_ON_RIGHT;
				else
					lvc.fmt &= ~LVCFMT_BITMAP_ON_RIGHT;
			}
			else // Assume its an icon number or the removal of the icon via -Icon.
			{
				if (adding)
				{
					lvc.mask |= LVCF_IMAGE;
					lvc.fmt |= LVCFMT_IMAGE; // Flag this column as displaying an image.
					lvc.iImage = ATOI(next_option) - 1; // -1 to convert to zero based.
				}
				else
					lvc.fmt &= ~LVCFMT_IMAGE; // Flag this column as NOT displaying an image.
			}
		}

		else // Handle things that are more general than the above, such as single letter options and pure numbers.
		{
			// Width does not have a W prefix to permit a naked expression to be used as the entirely of
			// options.  For example: LV_SetCol(1, old_width + 10)
			// v1.0.37: Fixed to allow floating point (although ATOI below will convert it to integer).
			if (IsPureNumeric(next_option, true, false, true)) // Above has already verified that *next_option can't be whitespace.
			{
				do_auto_size = 0; // Turn off any auto-sizing that may have been put into effect by default (such as for insertion).
				lvc.mask |= LVCF_WIDTH;
				lvc.cx = ATOI(next_option);
			}
		}

		// If the item was not handled by the above, ignore it because it is unknown.
		*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.
	}

	// Apply any changed justification/alignment to the fmt bit field:
	lvc.fmt = (lvc.fmt & ~LVCFMT_JUSTIFYMASK) | new_justify;

	if (aParamCount > 2) // Parameter #3 (text) is present.
	{
		if (lvc.pszText = ExprTokenToString(*aParam[2], buf)) // Assign.
			lvc.mask |= LVCF_TEXT;
		//else not an operand, so don't add the mask.  Haven't found a way to produce this situation yet, but safe to assume it's possible.
	}

	if (mode == 'M') // Modify vs. Insert (Delete was already returned from, higher above).
		// For code simplicity, this is called unconditionally even if nothing internal the control's column
		// needs updating.  This seems justified given how rarely columns are modified.
		aResultToken.value_int64 = ListView_SetColumn(control.hwnd, index, &lvc); // Returns TRUE/FALSE.
	else // Insert
	{
		// It's important to note that when the user slides columns around via drag and drop, the
		// column index as seen by the script is not changed.  This is fortunate because otherwise,
		// the lv_attrib.col array would get out of sync with the column indices.  Testing shows that
		// all of the following operations respect the original column index, regardless of where the
		// user may have moved the column physically: InsertCol, DeleteCol, ModifyCol.  Insert and Delete
		// shifts the indices of those columns that *originally* lay to the right of the affected column.
		// Doesn't seem to do anything -- not even with respect to inserting a new first column with it's
		// unusual behavior of inheriting the previously column's contents -- so it's disabled for now.
		// Testing shows that it also does not seem to cause a new column to inherit the indicated subitem's
		// text, even when iSubItem is set to index + 1 vs. index:
		//lvc.mask |= LVCF_SUBITEM;
		//lvc.iSubItem = index;
		// Testing shows that the following serve to set the column's physical/display position in the
		// heading to iOrder without affecting the specified index.  This concept is very similar to
		// when the user drags and drops a column heading to a new position: it's index doesn't change,
		// only it's displayed position:
		//lvc.mask |= LVCF_ORDER;
		//lvc.iOrder = index + 1;
		if (   !(aResultToken.value_int64 = ListView_InsertColumn(control.hwnd, index, &lvc) + 1)   ) // +1 to convert the new index to 1-based.
			return; // Since column could not be inserted, return so that below, sort-now, etc. are not done.
		index = (int)aResultToken.value_int64 - 1; // Update in case some other index was assigned. -1 to convert back to zero-based.
		if (index < lv_attrib.col_count) // Since col is not being appended to the end, make room in the array to insert this column.
			MoveMemory(lv_attrib.col+index+1, lv_attrib.col+index, sizeof(lv_col_type)*(lv_attrib.col_count-index));
			// Above: Shift columns to the right by one.
		lv_attrib.col[index] = col; // Copy temp struct's members to the correct element in the array.
		// The above is done even when index==0 because "col" may contain attributes set via the Options
		// parameter.  Therefore, for code simplicity and rarity of real-world need, no attempt is made
		// to make the following idea work:
		// When index==0, retain the existing attributes due to the unique behavior of inserting a new first
		// column: The new first column inherit's the old column's values (fields), so it seems best to also have it
		// inherit the old column's attributs.
		++lv_attrib.col_count; // New column successfully added.  Must be done only after the MoveMemory() above.
	}

	// Auto-size is done only at this late a stage, in case column was just created above.
	// Note that ListView_SetColumn() apparently does not support LVSCW_AUTOSIZE_USEHEADER for it's "cx" member.
	if (do_auto_size && GuiType::ControlGetListViewMode(control.hwnd) == LVS_REPORT)
		ListView_SetColumnWidth(control.hwnd, index, do_auto_size); // aResultToken.value_int64 was previous set to the more important result above.
	//else v1.0.36.03: Don't attempt to auto-size the columns while the view is not report-view because
	// that causes any subsequent switch to the "list" view to be corrupted (invisible icons and items).

	if (sort_now)
		GuiType::LV_Sort(control, index, false, sort_now_direction);
}



void BIF_LV_SetImageList(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// Returns (MSDN): "handle to the image list previously associated with the control if successful; NULL otherwise."
// Parameters:
// 1: HIMAGELIST obtained from somewhere such as IL_Create().
// 2: Optional: Type of list.
{
	aResultToken.value_int64 = 0;
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Window doesn't exist.
	// Control doesn't exist (i.e. no ListView in window).
	// Column not found in ListView.
	if (!g_gui[g.GuiDefaultWindowIndex])
		return;
	GuiType &gui = *g_gui[g.GuiDefaultWindowIndex]; // Always operate on thread's default window to simplify the syntax.
	if (!gui.mCurrentListView)
		return;
	// Caller has ensured that there is at least one incoming parameter:
	HIMAGELIST himl = (HIMAGELIST)ExprTokenToInt64(*aParam[0]);
	int list_type;
	if (aParamCount > 1)
		list_type = (int)ExprTokenToInt64(*aParam[1]);
	else // Auto-detect large vs. small icons based on the actual icon size in the image list.
	{
		int cx, cy;
		ImageList_GetIconSize(himl, &cx, &cy);
		list_type = (cx > GetSystemMetrics(SM_CXSMICON)) ? LVSIL_NORMAL : LVSIL_SMALL;
	}
	aResultToken.value_int64 = (__int64)ListView_SetImageList(gui.mCurrentListView->hwnd, himl, list_type);
}



void BIF_IL_Create(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// Returns: Handle to the new image list, or 0 on failure.
// Parameters:
// 1: Initial image count (ImageList_Create() ignores values <=0, so no need for error checking).
// 2: Grow count (testing shows it can grow multiple times, even when this is set <=0, so it's apparently only a performance aid)
// 3: Width of each image (overloaded to mean small icon size when omitted or false, large icon size otherwise).
// 4: Future: Height of each image [if this param is present and >0, it would mean param 3 is not being used in its TRUE/FALSE mode)
// 5: Future: Flags/Color depth
{
	// So that param3 can be reserved as a future "specified width" param, to go along with "specified height"
	// after it, only when the parameter is both present and numerically zero are large icons used.  Otherwise,
	// small icons are used.
	int param3 = aParamCount > 2 ? (int)ExprTokenToInt64(*aParam[2]) : 0;
	aResultToken.value_int64 = (__int64)ImageList_Create(GetSystemMetrics(param3 ? SM_CXICON : SM_CXSMICON)
		, GetSystemMetrics(param3 ? SM_CYICON : SM_CYSMICON)
		, ILC_MASK | ILC_COLOR32  // ILC_COLOR32 or at least something higher than ILC_COLOR is necessary to support true-color icons.
		, aParamCount > 0 ? (int)ExprTokenToInt64(*aParam[0]) : 2    // cInitial. 2 seems a better default than one, since it might be common to have only two icons in the list.
		, aParamCount > 1 ? (int)ExprTokenToInt64(*aParam[1]) : 5);  // cGrow.  Somewhat arbitrary default.
}



void BIF_IL_Destroy(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: HIMAGELIST obtained from somewhere such as IL_Create().
{
	// Load-time validation has ensured there is at least one parameter.
	// Returns nonzero if successful, or zero otherwise, so force it to conform to TRUE/FALSE for
	// better consistency with other functions:
	aResultToken.value_int64 = ImageList_Destroy((HIMAGELIST)ExprTokenToInt64(*aParam[0])) ? 1 : 0;
}



void BIF_IL_Add(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
// Returns: the one-based index of the newly added icon, or zero on failure.
// Parameters:
// 1: HIMAGELIST: Handle of an existing ImageList.
// 2: Filename from which to load the icon or bitmap.
// 3: Icon number within the filename (or mask color for non-icon images).
// 4: The mere presence of this parameter indicates that param #3 is mask RGB-color vs. icon number.
//    This param's value should be "true" to resize the image to fit the image-list's size or false
//    to divide up the image into a series of separate images based on its width.
//    (this parameter could be overloaded to be the filename containing the mask image, or perhaps an HBITMAP
//    provided directly by the script)
// 5: Future: can be the scaling height to go along with an overload of #4 as the width.  However,
//    since all images in an image list are of the same size, the use of this would be limited to
//    only those times when the imagelist would be scaled prior to dividing it into separate images.
// The parameters above (at least #4) can be overloaded in the future calling ImageList_GetImageInfo() to determine
// whether the imagelist has a mask.
{
	char *buf = aResultToken.buf; // Must be saved early since below overwrites the union (better maintainability too).
	aResultToken.value_int64 = 0; // Set default in case of early return.
	HIMAGELIST himl = (HIMAGELIST)ExprTokenToInt64(*aParam[0]); // Load-time validation has ensured there is a first parameter.
	if (!himl)
		return;

	// Load-time validation has ensured there are at least two parameters:
	char *filespec;
	if (   !(filespec = ExprTokenToString(*aParam[1], buf))   ) // Not an operand.  Haven't found a way to produce this situation yet, but safe to assume it's possible.
		return; // Due to rarity of this condition, keep 0/false value as the result.

	int param3 = (aParamCount > 2) ? (int)ExprTokenToInt64(*aParam[2]) : 0; // Default to zero so that -1 later below will get passed to LoadPicture().

	int icon_number, width = 0, height = 0; // Zero width/height causes image to be loaded at its actual width/height.
	if (aParamCount > 3) // Presence of fourth parameter switches mode to be "load a non-icon image".
	{
		icon_number = -1;
		if (ExprTokenToInt64(*aParam[3])) // A value of True indicates that the image should be scaled to fit the imagelist's image size.
			ImageList_GetIconSize(himl, &width, &height); // Determine the width/height to which it should be scaled.
		//else retain defaults of zero for width/height, which loads the image at actual size, which in turn
		// lets ImageList_AddMasked() divide it up into separate images based on its width.
	}
	else
		icon_number = param3 - 1; // Param3 can be explicitly 1, which results in a zero to force the ExtractIcon method.

	int image_type;
	HBITMAP hbitmap = LoadPicture(filespec, width, height, image_type
		, icon_number // -1 means "unspecified", which allows the use of LoadImage vs. ExtractIcon.
		, false); // Defaulting to "false" for "use GDIplus" provides more consistent appearance across multiple OSes.
	if (!hbitmap)
		return;

	if (image_type == IMAGE_BITMAP) // In this mode, param3 is always assumed to be an RGB color.
	{
		// Return the index of the new image or 0 on failure.
		aResultToken.value_int64 = ImageList_AddMasked(himl, hbitmap, rgb_to_bgr((int)param3)) + 1; // +1 to convert to one-based.
		DeleteObject(hbitmap);
	}
	else // ICON or CURSOR.
	{
		// Return the index of the new image or 0 on failure.
		aResultToken.value_int64 = ImageList_AddIcon(himl, (HICON)hbitmap) + 1; // +1 to convert to one-based.
		DestroyIcon((HICON)hbitmap); // Works on cursors too.  See notes in LoadPicture().
	}
}




////////////////////////////////////////////////////////
// HELPER FUNCTIONS FOR TOKENS AND BUILT-IN FUNCTIONS //
////////////////////////////////////////////////////////

__int64 ExprTokenToInt64(ExprTokenType &aToken)
// Converts the contents of aToken to an int.
{
	// Some callers, such as those that cast our return value to UINT, rely on the use of 64-bit to preserve
	// unsigned values and also wrap any signed values into the unsigned domain.
	switch (aToken.symbol)
	{
		case SYM_INTEGER: return (int)aToken.value_int64;
		case SYM_FLOAT: return (int)aToken.value_double;
		case SYM_VAR: return (int)ATOI(aToken.var->Contents());
		default: // SYM_STRING or SYM_OPERAND
			return ATOI(aToken.marker);
	}
}



double ExprTokenToDouble(ExprTokenType &aToken)
// Converts the contents of aToken to a double.
{
	switch (aToken.symbol)
	{
		case SYM_INTEGER: return (double)aToken.value_int64;
		case SYM_FLOAT: return aToken.value_double;
		case SYM_VAR: return ATOF(aToken.var->Contents());
		default: // SYM_STRING or SYM_OPERAND
			return ATOF(aToken.marker);
	}
}



char *ExprTokenToString(ExprTokenType &aToken, char *aBuf)
// Returns NULL on failure.  Otherwise, it returns either aBuf (if aBuf was needed for the conversion)
// or the token's own string.  Caller has ensured that aBuf is at least MAX_FORMATTED_NUMBER_LENGTH+1 in size.
{
	switch (aToken.symbol)
	{
	case SYM_STRING:
	case SYM_OPERAND:
		return aToken.marker;
	case SYM_VAR:
		return aToken.var->Contents();
	case SYM_INTEGER:
		return ITOA64(aToken.value_int64, aBuf);
	case SYM_FLOAT:
		snprintf(aBuf, MAX_FORMATTED_NUMBER_LENGTH + 1, g.FormatFloat, aToken.value_double);
		return aBuf;
	default: // Not an operand.
		return NULL;
	}
}



ResultType ExprTokenToVar(ExprTokenType &aToken, Var &aOutputVar)
// Writes aToken's value into aOutputVar.
// Returns FAIL if aToken isn't an operand or the assignment failed.  Returns OK on success.
// Currently only supports SYM_VAR if the variable is a normal variable, not a built-in or env. var.
{
	switch (aToken.symbol)
	{
	case SYM_STRING:
	case SYM_OPERAND:
		return aOutputVar.Assign(aToken.marker);
	case SYM_VAR:
		return aOutputVar.Assign(aToken.var->Contents());
	case SYM_INTEGER:
		return aOutputVar.Assign(aToken.value_int64);
	case SYM_FLOAT:
		return aOutputVar.Assign(aToken.value_double);
	default: // Not an operand.
		return FAIL;
	}
}



ResultType ExprTokenToDoubleOrInt(ExprTokenType &aToken)
// Converts aToken's contents to a numeric value, either int or float (whichever is more appropriate).
// Returns FAIL when aToken isn't an operand or is but contains a string that isn't purely numeric.
{
	char *str;
	switch (aToken.symbol)
	{
		case SYM_INTEGER:
		case SYM_FLOAT:
			return OK;
		case SYM_VAR:
			str = aToken.var->Contents();
			break;
		case SYM_STRING:   // v1.0.40.06: Fixed to be listed explicitly so that "default" case can return failure.
		case SYM_OPERAND:
			str = aToken.marker;
			break;
		default:  // Not an operand. Haven't found a way to produce this situation yet, but safe to assume it's possible.
			return FAIL;
	}
	// Since above didn't return, interpret "str" as a number.
	switch (aToken.symbol = IsPureNumeric(str, true, false, true))
	{
	case PURE_INTEGER:
		aToken.value_int64 = ATOI64(str);
		break;
	case PURE_FLOAT:
		aToken.value_double = ATOF(str);
		break;
	default: // Not a pure number.
		aToken.marker = ""; // For completeness.  Some callers such as BIF_Abs() rely on this being done.
		return FAIL;
	}
	return OK; // Since above didn't return, indicate success.
}



int ConvertJoy(char *aBuf, int *aJoystickID, bool aAllowOnlyButtons)
// The caller TextToKey() currently relies on the fact that when aAllowOnlyButtons==true, a value
// that can fit in a sc_type (USHORT) is returned, which is true since the joystick buttons
// are very small numbers (JOYCTRL_1==12).
{
	if (aJoystickID)
		*aJoystickID = 0;  // Set default output value for the caller.
	if (!aBuf || !*aBuf) return JOYCTRL_INVALID;
	char *aBuf_orig = aBuf;
	for (; *aBuf >= '0' && *aBuf <= '9'; ++aBuf); // self-contained loop to find the first non-digit.
	if (aBuf > aBuf_orig) // The string starts with a number.
	{
		int joystick_id = ATOI(aBuf_orig) - 1;
		if (joystick_id < 0 || joystick_id >= MAX_JOYSTICKS)
			return JOYCTRL_INVALID;
		if (aJoystickID)
			*aJoystickID = joystick_id;  // Use ATOI vs. atoi even though hex isn't supported yet.
	}

	if (!strnicmp(aBuf, "Joy", 3))
	{
		if (IsPureNumeric(aBuf + 3, false, false))
		{
			int offset = ATOI(aBuf + 3);
			if (offset < 1 || offset > MAX_JOY_BUTTONS)
				return JOYCTRL_INVALID;
			return JOYCTRL_1 + offset - 1;
		}
	}
	if (aAllowOnlyButtons)
		return JOYCTRL_INVALID;

	// Otherwise:
	if (!stricmp(aBuf, "JoyX")) return JOYCTRL_XPOS;
	if (!stricmp(aBuf, "JoyY")) return JOYCTRL_YPOS;
	if (!stricmp(aBuf, "JoyZ")) return JOYCTRL_ZPOS;
	if (!stricmp(aBuf, "JoyR")) return JOYCTRL_RPOS;
	if (!stricmp(aBuf, "JoyU")) return JOYCTRL_UPOS;
	if (!stricmp(aBuf, "JoyV")) return JOYCTRL_VPOS;
	if (!stricmp(aBuf, "JoyPOV")) return JOYCTRL_POV;
	if (!stricmp(aBuf, "JoyName")) return JOYCTRL_NAME;
	if (!stricmp(aBuf, "JoyButtons")) return JOYCTRL_BUTTONS;
	if (!stricmp(aBuf, "JoyAxes")) return JOYCTRL_AXES;
	if (!stricmp(aBuf, "JoyInfo")) return JOYCTRL_INFO;
	return JOYCTRL_INVALID;
}



bool ScriptGetKeyState(vk_type aVK, KeyStateTypes aKeyStateType)
// Returns true if "down", false if "up".
{
    if (!aVK) // Assume "up" if indeterminate.
		return false;

	switch (aKeyStateType)
	{
	case KEYSTATE_TOGGLE: // Whether a toggleable key such as CapsLock is currently turned on.
		// Under Win9x, at least certain versions and for certain hardware, this
		// doesn't seem to be always accurate, especially when the key has just
		// been toggled and the user hasn't pressed any other key since then.
		// I tried using GetKeyboardState() instead, but it produces the same
		// result.  Therefore, I've documented this as a limitation in the help file.
		// In addition, this was attempted but it didn't seem to help:
		//if (g_os.IsWin9x())
		//{
		//	DWORD fore_thread = GetWindowThreadProcessId(GetForegroundWindow(), NULL);
		//	bool is_attached_my_to_fore = false;
		//	if (fore_thread && fore_thread != g_MainThreadID)
		//		is_attached_my_to_fore = AttachThreadInput(g_MainThreadID, fore_thread, TRUE) != 0;
		//	output_var->Assign(IsKeyToggledOn(aVK) ? "D" : "U");
		//	if (is_attached_my_to_fore)
		//		AttachThreadInput(g_MainThreadID, fore_thread, FALSE);
		//	return OK;
		//}
		//else
		return IsKeyToggledOn(aVK); // This also works for the INSERT key, but only on XP (and possibly Win2k).
	case KEYSTATE_PHYSICAL: // Physical state of key.
		if (IsMouseVK(aVK)) // mouse button
		{
			if (g_MouseHook) // mouse hook is installed, so use it's tracking of physical state.
				return g_PhysicalKeyState[aVK] & STATE_DOWN;
			else // Even for Win9x/NT, it seems slightly better to call this rather than IsKeyDown9xNT():
				return IsPhysicallyDown(aVK);
		}
		else // keyboard
		{
			if (g_KeybdHook)
			{
				// Since the hook is installed, use its value rather than that from
				// GetAsyncKeyState(), which doesn't seem to return the physical state
				// as expected/advertised, least under WinXP.
				// But first, correct the hook modifier state if it needs it.  See comments
				// in GetModifierLRState() for why this is needed:
				if (KeyToModifiersLR(aVK))     // It's a modifier.
					GetModifierLRState(true); // Correct hook's physical state if needed.
				return g_PhysicalKeyState[aVK] & STATE_DOWN;
			}
			else // Even for Win9x/NT, it seems slightly better to call this rather than IsKeyDown9xNT():
				return IsPhysicallyDown(aVK);
		}
	} // switch()

	// Otherwise, use the default state-type: KEYSTATE_LOGICAL
	if (g_os.IsWin9x() || g_os.IsWinNT4())
		return IsKeyDown9xNT(aVK); // This seems more likely to be reliable.
	else
		// On XP/2K at least, a key can be physically down even if it isn't logically down,
		// which is why the below specifically calls IsKeyDown2kXP() rather than some more
		// comprehensive method such as consulting the physical key state as tracked by the hook:
		return IsKeyDown2kXP(aVK);
		// Known limitation: For some reason, both the above and IsKeyDown9xNT() will indicate
		// that the CONTROL key is up whenever RButton is down, at least if the mouse hook is
		// installed without the keyboard hook.  No known explanation.
}



double ScriptGetJoyState(JoyControls aJoy, int aJoystickID, ExprTokenType &aToken, bool aUseBoolForUpDown)
// Caller must ensure that aToken.marker is a buffer large enough to handle the longest thing put into
// it here, which is currently jc.szPname (size=32). Caller has set aToken.symbol to be SYM_STRING.
// For buttons: Returns 0 if "up", non-zero if down.
// For axes and other controls: Returns a number indicating that controls position or status.
// If there was a problem determining the position/state, aToken is made blank and zero is returned.
// Also returns zero in cases where a non-numerical result is requested, such as the joystick name.
// In those cases, caller should use aToken.marker as the result.
{
	// Set default in case of early return.
	*aToken.marker = '\0'; // Blank vs. string "0" serves as an indication of failure.

	if (!aJoy) // Currently never called this way.
		return 0; // And leave aToken set to blank.

	bool aJoy_is_button = IS_JOYSTICK_BUTTON(aJoy);

	JOYCAPS jc;
	if (!aJoy_is_button && aJoy != JOYCTRL_POV)
	{
		// Get the joystick's range of motion so that we can report position as a percentage.
		if (joyGetDevCaps(aJoystickID, &jc, sizeof(JOYCAPS)) != JOYERR_NOERROR)
			ZeroMemory(&jc, sizeof(jc));  // Zero it on failure, for use of the zeroes later below.
	}

	// Fetch this struct's info only if needed:
	JOYINFOEX jie;
	if (aJoy != JOYCTRL_NAME && aJoy != JOYCTRL_BUTTONS && aJoy != JOYCTRL_AXES && aJoy != JOYCTRL_INFO)
	{
		jie.dwSize = sizeof(JOYINFOEX);
		jie.dwFlags = JOY_RETURNALL;
		if (joyGetPosEx(aJoystickID, &jie) != JOYERR_NOERROR)
			return 0; // And leave aToken set to blank.
		if (aJoy_is_button)
		{
			bool is_down = ((jie.dwButtons >> (aJoy - JOYCTRL_1)) & (DWORD)0x01);
			if (aUseBoolForUpDown) // i.e. Down==true and Up==false
			{
				aToken.symbol = SYM_INTEGER; // Override default type.
				aToken.value_int64 = is_down; // Forced to be 1 or 0 above, since it's "bool".
			}
			else
			{
				aToken.marker[0] = is_down ? 'D' : 'U';
				aToken.marker[1] = '\0';
			}
			return is_down;
		}
	}

	// Otherwise:
	UINT range;
	char *buf_ptr;
	double result_double;  // Not initialized to help catch bugs.

	switch(aJoy)
	{
	case JOYCTRL_XPOS:
		range = (jc.wXmax > jc.wXmin) ? jc.wXmax - jc.wXmin : 0;
		result_double = range ? 100 * (double)jie.dwXpos / range : jie.dwXpos;
		break;
	case JOYCTRL_YPOS:
		range = (jc.wYmax > jc.wYmin) ? jc.wYmax - jc.wYmin : 0;
		result_double = range ? 100 * (double)jie.dwYpos / range : jie.dwYpos;
		break;
	case JOYCTRL_ZPOS:
		range = (jc.wZmax > jc.wZmin) ? jc.wZmax - jc.wZmin : 0;
		result_double = range ? 100 * (double)jie.dwZpos / range : jie.dwZpos;
		break;
	case JOYCTRL_RPOS:  // Rudder or 4th axis.
		range = (jc.wRmax > jc.wRmin) ? jc.wRmax - jc.wRmin : 0;
		result_double = range ? 100 * (double)jie.dwRpos / range : jie.dwRpos;
		break;
	case JOYCTRL_UPOS:  // 5th axis.
		range = (jc.wUmax > jc.wUmin) ? jc.wUmax - jc.wUmin : 0;
		result_double = range ? 100 * (double)jie.dwUpos / range : jie.dwUpos;
		break;
	case JOYCTRL_VPOS:  // 6th axis.
		range = (jc.wVmax > jc.wVmin) ? jc.wVmax - jc.wVmin : 0;
		result_double = range ? 100 * (double)jie.dwVpos / range : jie.dwVpos;
		break;

	case JOYCTRL_POV:  // Need to explicitly compare against JOY_POVCENTERED because it's a WORD not a DWORD.
		if (jie.dwPOV == JOY_POVCENTERED)
		{
			// Retain default SYM_STRING type.
			strcpy(aToken.marker, "-1"); // Assign as string to ensure its written exactly as "-1". Documented behavior.
			return -1;
		}
		else
		{
			aToken.symbol = SYM_INTEGER; // Override default type.
			aToken.value_int64 = jie.dwPOV;
			return jie.dwPOV;
		}
		// No break since above always returns.

	case JOYCTRL_NAME:
		strcpy(aToken.marker, jc.szPname);
		return 0;  // Returns zero in cases where a non-numerical result is obtained.

	case JOYCTRL_BUTTONS:
		aToken.symbol = SYM_INTEGER; // Override default type.
		aToken.value_int64 = jc.wNumButtons;
		return jc.wNumButtons;  // wMaxButtons is the *driver's* max supported buttons.

	case JOYCTRL_AXES:
		aToken.symbol = SYM_INTEGER; // Override default type.
		aToken.value_int64 = jc.wNumAxes; // wMaxAxes is the *driver's* max supported axes.
		return jc.wNumAxes;

	case JOYCTRL_INFO:
		buf_ptr = aToken.marker;
		if (jc.wCaps & JOYCAPS_HASZ)
			*buf_ptr++ = 'Z';
		if (jc.wCaps & JOYCAPS_HASR)
			*buf_ptr++ = 'R';
		if (jc.wCaps & JOYCAPS_HASU)
			*buf_ptr++ = 'U';
		if (jc.wCaps & JOYCAPS_HASV)
			*buf_ptr++ = 'V';
		if (jc.wCaps & JOYCAPS_HASPOV)
		{
			*buf_ptr++ = 'P';
			if (jc.wCaps & JOYCAPS_POV4DIR)
				*buf_ptr++ = 'D';
			if (jc.wCaps & JOYCAPS_POVCTS)
				*buf_ptr++ = 'C';
		}
		*buf_ptr = '\0'; // Final termination.
		return 0;  // Returns zero in cases where a non-numerical result is obtained.
	} // switch()

	// If above didn't return, the result should now be in result_double.
	aToken.symbol = SYM_FLOAT; // Override default type.
	aToken.value_double = result_double;
	return result_double;
}
