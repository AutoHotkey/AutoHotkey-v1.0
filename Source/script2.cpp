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
#include <wininet.h> // For URLDownloadToFile().
#include <olectl.h> // for OleLoadPicture()
#include "qmath.h" // Used by Transform() [math.h incurs 2k larger code size just for ceil() & floor()]
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


ResultType Line::Splash(char *aOptions, char *aSubText, char *aMainText, char *aTitle, char *aFontName
	, char *aImageFile, bool aSplashImage)
// Parts of this have been adapted from the AutoIt3 source.
{
	int window_index = 0;  // Set the default window to operate upon (the first).
	char *options, *image_filename = aImageFile;  // Set default.
	bool turn_off = false;
	int percent = -1;  // Must be set to -1 in the case of SplashImage.

	if (aSplashImage)
	{
		options = aOptions;
		if (*aImageFile)
		{
			char *colon_pos = strchr(aImageFile, ':');
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
						char *blank_str = omit_leading_whitespace(image_filename);
						if (!*blank_str)
							image_filename = blank_str;
						window_index = ATOI(window_number_str) - 1;
						if (window_index < 0 || window_index >= MAX_SPLASHIMAGE_WINDOWS)
							return LineError("The window number must be between 1 and " MAX_SPLASHIMAGE_WINDOWS_STR
								"." ERR_ABORT, FAIL, aOptions);
					}
				}
			}
			turn_off = !stricmp(image_filename, "Off");
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
				return LineError("The window number must be between 1 and " MAX_PROGRESS_WINDOWS_STR "." ERR_ABORT
					, FAIL, aOptions);
			++options;
		}
		turn_off = !stricmp(options, "Off");
		// Allow floats at runtime for flexibility (i.e. in case aOptions was in a variable reference).
		// But still use ATOI for the conversion:
		if (!turn_off && IsPureNumeric(options, false, false, true))
			percent = ATOI(options);
		//else leave it set to the default.
	}

	SplashType &splash = aSplashImage ? g_SplashImage[window_index] : g_Progress[window_index];

	// In case it's possible for the window to get destroyed by other means (WinClose?).
	// Do this only after the above options were set so that the each window's settings
	// will be remembered until such time as "Command, Off" is used:
	if (splash.hwnd && !IsWindow(splash.hwnd))
		splash.hwnd = NULL;

	if (!turn_off && splash.hwnd && !*image_filename && (percent >= 0 || !*options)) // The "modify existing window" mode is in effect.
	{
		// If there is an existing window, just update its percentage (progress bar position) and text.
		// If not, do nothing since we don't have the original text of the window to recreate it.
		// Since this is our thread's window, it shouldn't be necessary to use SendMessageTimeout()
		// since the window cannot be hung if our thread is active.  Also, setting a text item
		// from non-blank to blank is not supported so that elements can be omitted from an update
		// command without changing the text that's in the window.  The script can specify %a_space%
		// to explicitly make an element blank.
		if (!aSplashImage && percent >= 0 && splash.percent != percent)
		{
			splash.percent = percent;
			if (splash.hwnd_bar)
				SendMessage(splash.hwnd_bar, PBM_SETPOS, (WPARAM)percent, 0);
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
	int style = WS_DISABLED|WS_POPUP|WS_CAPTION;  // WS_CAPTION implies WS_BORDER
	int xstyle = WS_EX_TOPMOST;
	int xpos = COORD_UNSPECIFIED;
	int ypos = COORD_UNSPECIFIED;
	int font_size1 = 0; // 0 is the flag to "use default size".
	int font_size2 = 0;
	int font_weight1 = FW_DONTCARE;  // Flag later logic to use default.
	int font_weight2 = FW_DONTCARE;  // Flag later logic to use default.
	COLORREF bar_color = CLR_DEFAULT;
	splash.color_bk = CLR_DEFAULT;
	splash.color_text = CLR_DEFAULT;
	splash.height = COORD_UNSPECIFIED; // au3's default is 100 (client area)
	if (aSplashImage)
	{
		#define SPLASH_DEFAULT_WIDTH 300   // au3's default
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

	for (char *cp = options; *cp; ++cp)
	{
		switch(toupper(*cp))
		{
		case 'A':  // Non-Always-on-top.  Synonymous with A0 in early versions.
			// Decided against this enforcement.  In the enhancement mentioned below is ever done (unlikely),
			// it seems that A1 can turn always-on-top on and A0 or A by itself can turn it off:
			//if (*(cp + 1) == '0') // The zero is required to allow for future enhancement: modify attrib. of existing window.
			xstyle &= ~WS_EX_TOPMOST;
			break;
		case 'B': // Borderless and/or Titleless
			style &= ~WS_CAPTION;
			if (*(cp + 1) == '1')
				style |= WS_BORDER;
			else if (*(cp + 1) == '2')
				style |= WS_DLGFRAME;
			break;
		case 'C': // Centered
			if (!*(cp + 1)) // Avoids out-of-bounds when the loop's own ++cp is done.
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
				if (color == CLR_DEFAULT) // A matching color name was not found, so assume it's in hex format.
				{
					if (strlen(color_str) > 6)
						color_str[6] = '\0';  // Shorten to exactly 6 chars, which happens if no space/tab delimiter is present.
					color = rgb_to_bgr(strtol(color_str, NULL, 16));
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
				centered_main = (*(cp + 1) != '0');
			}
		case 'F':
			if (!*(cp + 1)) // Avoids out-of-bounds when the loop's own ++cp is done.
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
			if (*(cp + 1) == '1')
				style |= WS_SIZEBOX;
			if (*(cp + 1) == '2')
				style |= WS_SIZEBOX|WS_MINIMIZEBOX|WS_MAXIMIZEBOX|WS_SYSMENU;
			break;
		case 'T': // Give it a task bar button by making it a non-owned window.
			owned = false;
			break;
		// For options such as W, X and Y:
		// Use atoi() vs. ATOI() to avoid interpreting something like 0x01B as hex when in fact
		// the B was meant to be an option letter:
		case 'W':
			if (!*(cp + 1)) // Avoids out-of-bounds when the loop's own ++cp is done.
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
			// Allow any width/height to be specified so that the window can be "rolled up" to its title bar:
			splash.height = atoi(cp + 1);
			break;
		case 'X':
			xpos = atoi(cp + 1);
			break;
		case 'Y':
			ypos = atoi(cp + 1);
			break;
		case 'Z':
			if (!*(cp + 1)) // Avoids out-of-bounds when the loop's own ++cp is done.
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
		}
	}

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
	if (   splash.object_height <= 0 && splash.object_height != COORD_UNSPECIFIED
		&& splash.object_width <= 0 && splash.object_width != COORD_UNSPECIFIED
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
		// does not affect the reported default font size.  Thus, the default is probably 8 for most/all
		// XP systems and probably other OSes as well.
		// By specifying PROOF_QUALITY the nearest matching font size should be chosen, which should avoid
		// any scaling artifacts that might be caused if default_gui_font_height is not 8.
		if (   !(splash.hfont1 = CreateFont(font_size1 ? -MulDiv(font_size1, pixels_per_point_y, 72) : (int)(1.25 * default_gui_font_height)
			, 0, 0, 0, font_weight1 ? font_weight1 : FW_SEMIBOLD, 0, 0, 0, DEFAULT_CHARSET, OUT_TT_PRECIS
			, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, FF_DONTCARE, *aFontName ? aFontName : default_font_name))   )
			// Call it again with default font in case above failed due to non-existent aFontName.
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
	AdjustWindowRectEx(&main_rect, style, FALSE, xstyle);
	int main_width = main_rect.right - main_rect.left;  // main.left might be slightly less than zero.
	int main_height = main_rect.bottom - main_rect.top; // main.top might be slightly less than zero.

	RECT desk_rect;
	SystemParametersInfo(SPI_GETWORKAREA, 0, &desk_rect, 0);  // Get Desktop rect excluding task bar.

	// Seems best (and easier) to unconditionally restrict window size to the size of the desktop,
	// since most users would probably want that.  This can be overridden by using WinMove afterward.
	if (main_width > desk_rect.right)
		main_width = desk_rect.right;
	if (main_height > desk_rect.bottom)
		main_height = desk_rect.bottom;

	if (xpos == COORD_UNSPECIFIED)
		xpos = (desk_rect.right - main_width) / 2;  // Don't use splash.width.
	if (ypos == COORD_UNSPECIFIED)
		ypos = (desk_rect.bottom - main_height) / 2; // Don't use splash.height

	// CREATE Main Splash Window
	// It seems best to make this an unowned window for two reasons:
	// 1) It will get its own task bar icon then, which is usually desirable for cases where
	//    there are several progress/splash windows or the window is monitoring something.
	// 2) The progress/splash window won't prevent the main window from being used (owned windows
	//    prevent their owners from ever becoming active).
	// However, it seems likely that some users would want the above to be configurable,
	// so now there is an option to change this behavior.
	if (!(splash.hwnd = CreateWindowEx(xstyle, WINDOW_CLASS_SPLASH, aTitle, style, xpos, ypos
		, main_width, main_height, owned ? g_hWnd : NULL, NULL, g_hInstance, NULL)))
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
		// CREATE Progress control (always starts off at its default percentage of zero):
		splash.hwnd_bar = CreateWindowEx(WS_EX_CLIENTEDGE, "msctls_progress32", NULL, WS_CHILD|WS_VISIBLE|PBS_SMOOTH
			, PROGRESS_BAR_POS, splash.hwnd, NULL, NULL, NULL);
		SendMessage(splash.hwnd_bar, PBM_SETRANGE, 0, MAKELONG(0, 100));
		SendMessage(splash.hwnd_bar, PBM_SETSTEP, 1, 0); // set some characteristics
		if (bar_color != CLR_DEFAULT)
			SendMessage(splash.hwnd_bar, PBM_SETBARCOLOR, 0, bar_color); // Set color.
		if (splash.color_bk != CLR_DEFAULT)
			SendMessage(splash.hwnd_bar, PBM_SETBKCOLOR, 0, splash.color_bk); // Set color.
		if (percent > 0)
		{
			// This happens when the window doesn't exist and a command such as the following is given:
			// Progress, 50 [, ...]
			splash.percent = percent;
			SendMessage(splash.hwnd_bar, PBM_SETPOS, (WPARAM)percent, 0);
		}
	}

	// CREATE Sub label
	splash.hwnd_text2 = CreateWindowEx(0, "static", aSubText
		, WS_CHILD|WS_VISIBLE|SS_NOPREFIX|(centered_sub ? SS_CENTER : SS_LEFT)
		, PROGRESS_SUB_POS, splash.hwnd, NULL, g_hInstance, NULL);
	SendMessage(splash.hwnd_text2, WM_SETFONT, (WPARAM)(splash.hfont2 ? splash.hfont2 : hfont_default), MAKELPARAM(TRUE, 0));

	// Show it without activating it.  Even with options that allow the window to be activated (such
	// as movable), it seems best to do this to prevent changing the current foreground window, which
	// is usually desirable for progress/splash windows since they should be seen but not be disruptive:
	ShowWindow(splash.hwnd,  SW_SHOWNOACTIVATE);
	return OK;
}



ResultType Line::ToolTip(char *aText, char *aX, char *aY, char *aID)
// Adapted from the AutoIt3 source.
// au3: Creates a tooltip with the specified text at any location on the screen.
// The window isn't created until it's first needed, so no resources are used until then.
// Also, the window is destroyed in AutoIt_Script's destructor so no resource leaks occur.
{
	int window_index = *aID ? ATOI(aID) - 1 : 0;
	if (window_index < 0 || window_index >= MAX_TOOLTIPS)
		return LineError("The window number must be between 1 and " MAX_TOOLTIPS_STR "." ERR_ABORT, FAIL, aID);
	HWND tip_hwnd = g_hWndToolTip[window_index];

	// Destroy windows except the first (for performance) so that resources/mem are conserved.
	// The first window will be hidden by the TTM_UPDATETIPTEXT message if aText is blank.
	if (window_index > 0 && !*aText)
	{
		if (tip_hwnd && IsWindow(tip_hwnd))
			DestroyWindow(tip_hwnd);
		g_hWndToolTip[window_index] = NULL;
		return OK;
	}

	// Set default values for the tip as the current mouse position.
	// UPDATE: Don't call GetCursorPos() unless absolutely needed because it seems to mess
	// up double-click timing, at least on XP.  UPDATE #2: Is isn't GetCursorPos() that's
	// interfering with double clicks, so it seems it must be the displaying of the ToolTip
	// window itself.

	RECT dtw;
	GetWindowRect(GetDesktopWindow(), &dtw);

	bool one_or_both_coords_unspecified = !*aX || !*aY;
	POINT pt, pt_cursor;
	if (one_or_both_coords_unspecified)
	{
		GetCursorPos(&pt_cursor);
		pt.x = pt_cursor.x + 16;  // Set default spot to be near the mouse cursor.
		pt.y = pt_cursor.y + 16;  // Use 16 to prevent the tooltip from overlapping large cursors.
		// Update: Below is no longer needed due to a better fix further down that handles multi-line tooltips.
		// 20 seems to be about the right amount to prevent it from "warping" to the top of the screen,
		// at least on XP:
		//if (pt.y > dtw.bottom - 20)
		//	pt.y = dtw.bottom - 20;
	}

	RECT rect;
	if ((*aX || *aY) && !(g.CoordMode & COORD_MODE_TOOLTIP)) // Need the rect.
	{
		if (!GetWindowRect(GetForegroundWindow(), &rect))
			return OK;  // Don't bother setting ErrorLevel with this command.
	}
	else
		rect.bottom = rect.left = rect.right = rect.top = 0;

	// This will also convert from relative to screen coordinates if rect contains non-zero values:
	if (*aX)
		pt.x = ATOI(aX) + rect.left;
	if (*aY)
		pt.y = ATOI(aY) + rect.top;

	TOOLINFO ti;
	ti.cbSize	= sizeof(ti);
	ti.uFlags	= TTF_TRACK;
	ti.hwnd		= NULL;  // Doesn't work: GetDesktopWindow()
	ti.hinst	= NULL;
	ti.uId		= 0;
	ti.lpszText	= aText;
	ti.rect.left = ti.rect.top = ti.rect.right = ti.rect.bottom = 0;

	// My: This does more harm that good (it causes the cursor to warp from the right side to the left
	// if it gets to close to the right side), so for now, I did a different fix (above) instead:
	//ti.rect.bottom = dtw.bottom;  // Just this first line was the au3 fix, not the others below.
	//ti.rect.right = dtw.right;
	//ti.rect.top = dtw.top;
	//ti.rect.left = dtw.left;

	DWORD dwResult;
	if (!tip_hwnd)
	{
		// This this window has no owner, it won't be automatically destroyed when its owner is.
		// Thus, it should be destroyed upon program termination.
		tip_hwnd = g_hWndToolTip[window_index] = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, NULL, TTS_NOPREFIX | TTS_ALWAYSTIP
			, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, NULL, NULL);
		SendMessageTimeout(tip_hwnd, TTM_ADDTOOL, 0, (LPARAM)(LPTOOLINFO)&ti, SMTO_ABORTIFHUNG, 2000, &dwResult);
		SendMessageTimeout(tip_hwnd, TTM_SETMAXTIPWIDTH, 0, (LPARAM)dtw.right, SMTO_ABORTIFHUNG, 2000, &dwResult);
		// Must do these next two when the window is first created, otherwise GetWindowRect() below will retrieve
		// a tooltip window size that is quite a bit taller than it winds up being:
		SendMessageTimeout(tip_hwnd, TTM_TRACKPOSITION, 0, (LPARAM)MAKELONG(pt.x, pt.y), SMTO_ABORTIFHUNG, 2000, &dwResult);
		SendMessageTimeout(tip_hwnd, TTM_TRACKACTIVATE, TRUE, (LPARAM)&ti, SMTO_ABORTIFHUNG, 2000, &dwResult);
	}
	else
		SendMessageTimeout(tip_hwnd, TTM_UPDATETIPTEXT, 0, (LPARAM)&ti, SMTO_ABORTIFHUNG, 2000, &dwResult);


	RECT ttw = {0};
	GetWindowRect(tip_hwnd, &ttw); // Must be called this late to ensure the tooltip has been created by above.
	int tt_width = ttw.right - ttw.left;
	int tt_height = ttw.bottom - ttw.top;
	if (pt.x + tt_width >= dtw.right)
		pt.x = dtw.right - tt_width - 1;
	if (pt.y + tt_height >= dtw.bottom)
		pt.y = dtw.bottom - tt_height - 1;

	if (one_or_both_coords_unspecified)
	{
		// Since Tooltip is being shown at the cursor's coordinates, try to ensure that the above
		// adjustment doesn't result in the cursor being inside the tooltip's window boundaries,
		// since that tends to cause problems such as blocking the tray area (which can make a
		// tootip script impossible to terminate).  Normally, that can only happen in this case
		// (one_or_both_coords_unspecified == true) when the cursor is near the buttom-right
		// corner of the screen (unless the mouse is moving more quickly than the script's
		// tool-tip update-frequency can cope with, but that seems inconsequential since it
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

	SendMessageTimeout(tip_hwnd, TTM_TRACKPOSITION, 0, (LPARAM)MAKELONG(pt.x, pt.y), SMTO_ABORTIFHUNG, 2000, &dwResult);
	// And do a TTM_TRACKACTIVATE even if the tooltip window already existed upon entry to this function,
	// so that in case it was hidden or dismissed while its HWND still exists, it will be shown again:
	SendMessageTimeout(tip_hwnd, TTM_TRACKACTIVATE, TRUE, (LPARAM)&ti, SMTO_ABORTIFHUNG, 2000, &dwResult);
	return OK;
}



ResultType Line::TrayTip(char *aTitle, char *aText, char *aTimeout, char *aOptions)
{
	if (!g_os.IsWin2000orLater()) // Older OSes not supported, so by design it does nothing.
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
	return OK;
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
	pure_numeric_type value1_is_pure_numeric, value2_is_pure_numeric;

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

	case TRANS_CMD_HTML:
	{
		// These are the encoding-neutral translations for ASC 128 through 255 as shown by Dreamweaver.
		// It's possible that using just the &#number convention (e.g. &#128 through &#255;) would be
		// more appropriate for some users, but that mode can be added in the future if it is ever
		// needed (by passing a mode setting for aValue2):
		// ÄÅÇÉÑÖÜáàâäãåçéèêëíìîïñóòôöõúùûü†°¢£§•¶ß®©™´¨≠ÆØ∞±≤≥¥µ∂∑∏π∫ªºΩæø
		// ¿¡¬√ƒ≈∆«»… ÀÃÕŒœ–—“”‘’÷◊ÿŸ⁄€‹›ﬁﬂ‡·‚„‰ÂÊÁËÈÍÎÏÌÓÔÒÚÛÙıˆ˜¯˘˙˚¸˝˛ˇ
		static const char *html[128] = {
			  "&euro;", "&#129;", "&sbquo;", "&fnof;", "&bdquo;", "&hellip;", "&dagger;", "&Dagger;"
			, "&circ;", "&permil;", "&Scaron;", "&lsaquo;", "&OElig;", "&#141;", "&#381;", "&#143;"
			, "&#144;", "&lsquo;", "&rsquo;", "&ldquo;", "&rdquo;", "&bull;", "&ndash;", "&mdash;"
			, "&tilde;", "&trade;", "&scaron;", "&rsaquo;", "&oelig;", "&#157;", "&#382;", "&Yuml;"
			, "&nbsp;", "&iexcl;", "&cent;", "&pound;", "&curren;", "&yen;", "&brvbar;", "&sect;"
			, "&uml;", "&copy;", "&ordf;", "&laquo;", "&not;", "&shy;", "&reg;", "&macr;"
			, "&deg;", "&plusmn;", "&sup2;", "&sup3;", "&acute;", "&micro;", "&para;", "&middot;"
			, "&cedil;", "&sup1;", "&ordm;", "&raquo;", "&frac14;", "&frac12;", "&frac34;", "&iquest;"
			, "&Agrave;", "&Aacute;", "&Acirc;", "&Atilde;", "&Auml;", "&Aring;", "&AElig;", "&Ccedil;"
			, "&Egrave;", "&Eacute;", "&Ecirc;", "&Euml;", "&Igrave;", "&Iacute;", "&Icirc;", "&Iuml;"
			, "&ETH;", "&Ntilde;", "&Ograve;", "&Oacute;", "&Ocirc;", "&Otilde;", "&Ouml;", "&times;"
			, "&Oslash;", "&Ugrave;", "&Uacute;", "&Ucirc;", "&Uuml;", "&Yacute;", "&THORN;", "&szlig;"
			, "&agrave;", "&aacute;", "&acirc;", "&atilde;", "&auml;", "&aring;", "&aelig;", "&ccedil;"
			, "&egrave;", "&eacute;", "&ecirc;", "&euml;", "&igrave;", "&iacute;", "&icirc;", "&iuml;"
			, "&eth;", "&ntilde;", "&ograve;", "&oacute;", "&ocirc;", "&otilde;", "&ouml;", "&divide;"
			, "&oslash;", "&ugrave;", "&uacute;", "&ucirc;", "&uuml;", "&yacute;", "&thorn;", "&yuml;"
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
					length += (VarSizeType)strlen(html[*ucp - 128]);
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
					for (char *dp = (char *)html[*ucp - 128]; *dp; ++dp)
						*contents++ = *dp;
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
		// Currently, a negative aValue1 isn't supported (AutoIt3 doesn't support them either).
		// The reason for this is that since fractional exponents are supported (e.g. 0.5, which
		// results in the square root), there would have to be some extra detection to ensure
		// that a negative aValue1 is never used with fractional exponent (since the sqrt of
		// a negative is undefined).  In addition, qmathPow() doesn't support negatives, returning
		// an unexpectedly large value or -1.#IND00 instead.
		value_double1 = ATOF(aValue1);
		if (value_double1 < 0)
			return output_var->Assign();  // Return a consistent result (blank) rather than something that varies.
		value_double2 = ATOF(aValue2);
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
		// Adapted from the AutoIt3 source.
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

	// These next two might be improved by to avoid loss of 64-bit integer precision
	// by using a string conversion algorithm rather than converting to double:
	case TRANS_CMD_CEIL:
		return output_var->Assign((INT64)qmathCeil(ATOF(aValue1)));

	case TRANS_CMD_FLOOR:
		return output_var->Assign((INT64)qmathFloor(ATOF(aValue1)));

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
		else // Treat it as a 32-bit unsigned value when inverting and assigning:
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
	mod_type modifiers;

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
			modifiers = 0;  // Init prior to below.
			if (vk = TextToVK(single_char_string, &modifiers, true))
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
					if (modifiers & MOD_SHIFT)
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
	char **realloc_temp;  // Needed since realloc returns NULL on failure but leaves original block allocated!
	g_input.MatchCount = 0;  // Set default.
	if (*aMatchList)
	{
		// If needed, create the array of pointers that points into MatchBuf to each match phrase:
		if (!g_input.match)
		{
			if (   !(g_input.match = (char **)malloc(INPUT_ARRAY_BLOCK_SIZE * sizeof(char *)))   )
				return LineError("Out of mem #1");  // Short msg. since so rare.
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
				return LineError("Out of mem #2");  // Short msg. since so rare.
			}
		}
		// Copy aMatchList into the match buffer:
		char *source, *dest;
		for (source = aMatchList, dest = g_input.match[g_input.MatchCount] = g_input.MatchBuf; *source; ++source, ++dest)
		{
			if (*source == ',')  // Each comma becomes the terminator of the previous key phrase.
			{
				*dest = '\0';
				if (*g_input.match[g_input.MatchCount]) // i.e. omit empty strings from the match list.
					++g_input.MatchCount;
				if (*(source + 1)) // There is a next element.
				{
					if (g_input.MatchCount >= g_input.MatchCountMax) // Rarely needed, so just realloc() to expand.
					{
						// Expand the array by one block:
						if (   !(realloc_temp = (char **)realloc(g_input.match  // Must use a temp variable.
							, (g_input.MatchCountMax + INPUT_ARRAY_BLOCK_SIZE) * sizeof(char *)))   )
							return LineError("Out of mem #3");  // Short msg. since so rare.
						g_input.match = realloc_temp;
						g_input.MatchCountMax += INPUT_ARRAY_BLOCK_SIZE;
					}
					g_input.match[g_input.MatchCount] = dest + 1;
				}
			}
			else // Not a comma, so just copy it over.
				*dest = *source;
		}
		*dest = '\0';  // Terminate the last item.
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
		g_ErrorLevel->Assign("Timeout");
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
			int count = ToAscii(g_input.EndingVK, g_vk_to_sc[g_input.EndingVK].a, (PBYTE)&state
				, (LPWORD)(key_name + 7), g_MenuIsVisible ? 1 : 0);
			*(key_name + 7 + count) = '\0';  // Terminate the string.
		}
		else
			g_input.EndedBySC ? SCToKeyName(g_input.EndingSC, key_name + 7, sizeof(key_name) - 7)
				: VKToKeyName(g_input.EndingVK, g_input.EndingSC, key_name + 7, sizeof(key_name) - 7);
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
	if (!GetWindowRect(target_window, &rect))
		return OK;  // Can't set errorlevel, see above.
	MoveWindow(target_window
		, *aX && stricmp(aX, "default") ? ATOI(aX) : rect.left  // X-position
		, *aY && stricmp(aY, "default") ? ATOI(aY) : rect.top   // Y-position
		, *aWidth && stricmp(aWidth, "default") ? ATOI(aWidth) : rect.right - rect.left
		, *aHeight && stricmp(aHeight, "default") ? ATOI(aHeight) : rect.bottom - rect.top
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
		target_menu_pos = (menu_param[i][menu_param_length - 1] == '&') ? ATOI(menu_param[i]) - 1 : -1;
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
	, char *aExcludeTitle, char *aExcludeText, bool aSendRaw)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return OK;
	HWND control_window = stricmp(aControl, "ahk_parent") ? ControlExist(target_window, aControl)
		: target_window;
	if (!control_window)
		return OK;
	SendKeys(aKeysToSend, aSendRaw, control_window);
	// But don't do WinDelay because KeyDelay should have been in effect for the above.
	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return OK;
}



ResultType Line::ControlClick(vk_type aVK, int aClickCount, char *aOptions, char *aControl
	, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return OK;
	HWND control_window = ControlExist(target_window, aControl);
	if (!control_window)
		return OK;

	if (aClickCount <= 0)
		// Allow this to simply "do nothing", because it increases flexibility
		// in the case where the number of clicks is a dereferenced script variable
		// that may sometimes (by intent) resolve to zero or negative:
		return g_ErrorLevel->Assign(ERRORLEVEL_NONE);

	// Set the defaults that will be in effect unless overridden by options:
	KeyEventTypes event_type = KEYDOWNANDUP;
	int xpos = COORD_UNSPECIFIED;
	int ypos = COORD_UNSPECIFIED;

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
		// For the below:
		// Use atoi() vs. ATOI() to avoid interpreting something like 0x01D as hex
		// when in fact the D was meant to be an option letter:
		case 'X':
			xpos = atoi(cp + 1);
			break;
		case 'Y':
			ypos = atoi(cp + 1);
			break;
		}
	}

	RECT rect;
	if (xpos == COORD_UNSPECIFIED || ypos == COORD_UNSPECIFIED)
	{
		// AutoIt3: Get the dimensions of the control so we can click the centre of it
		// (maybe safer and more natural than 0,0).
		// My: In addition, this is probably better for some large controls (e.g. SysListView32) because
		// clicking at 0,0 might activate a part of the control that is not even visible:
		if (!GetWindowRect(control_window, &rect))
			return OK;  // Let ErrorLevel tell the story.
		if (xpos == COORD_UNSPECIFIED)
			xpos = (rect.right - rect.left) / 2;
		if (ypos == COORD_UNSPECIFIED)
			ypos = (rect.bottom - rect.top) / 2;
	}
	LPARAM lparam = MAKELPARAM(xpos, ypos);

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
	// AutoIt3: See BM_CLICK documentation, applies to this too
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

	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return OK;
}



ResultType Line::ControlMove(char *aControl, char *aX, char *aY, char *aWidth, char *aHeight
	, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	DETERMINE_TARGET_WINDOW
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

	DETERMINE_TARGET_WINDOW
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
	// Realistically never fails since DETERMINE_TARGET_WINDOW and ControlExist should always yield
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
	DETERMINE_TARGET_WINDOW
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

	char class_name[1024];
	cah.class_name = class_name;
	if (!GetClassName(cah.hwnd, class_name, sizeof(class_name) - 3)) // -3 to allow room for sequence number.
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
	#define CAH ((class_and_hwnd_type *)lParam)
	char class_name[1024];
	if (!GetClassName(aWnd, class_name, sizeof(class_name)))
		return TRUE;  // Continue the enumeration.
	if (!strcmp(class_name, CAH->class_name)) // Class names match.
	{
		++CAH->class_count;
		if (aWnd == CAH->hwnd)  // The caller-specified window has been found.
		{
			CAH->is_found = true;
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
	DoControlDelay;
	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return OK;
}



ResultType Line::ControlGetText(char *aControl, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR);  // Set default.
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



ResultType Line::Control(char *aCmd, char *aValue, char *aControl, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
// This function has been adapted from the AutoIt3 source.
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR);  // Set default since there are many points of return.
	ControlCmds control_cmd = ConvertControlCmd(aCmd);
	// Since command names are validated at load-time, this only happens if the command name
	// was contained in a variable reference.  Since that is very rare, just set ErrorLevel
	// and return:
	if (control_cmd == CONTROL_CMD_INVALID)
		return OK;  // Let ErrorLevel tell the story.

	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return OK;  // Let ErrorLevel tell the story.
	HWND control_window = ControlExist(target_window, aControl);
	if (!control_window)
		return OK;  // Let ErrorLevel tell the story.

	HWND immediate_parent;  // Possibly not the same as target_window since controls can themselves have children.
	int control_id, control_index;
	DWORD dwResult, new_button_state;
	UINT msg, x_msg, y_msg;
	RECT rect;
	LPARAM lparam;
	vk_type vk;
	int key_count;

	switch(control_cmd)
	{
	case CONTROL_CMD_CHECK: // au3: Must be a Button
	case CONTROL_CMD_UNCHECK:
	{ // Need braces for ATTACH_THREAD_INPUT macro.
		new_button_state = (control_cmd == CONTROL_CMD_CHECK) ? BST_CHECKED : BST_UNCHECKED;
		if (!SendMessageTimeout(control_window, BM_GETCHECK, 0, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return OK;  // Let ErrorLevel tell the story.
		if (dwResult == new_button_state) // It's already in the right state, so don't press it.
			break;
		// MSDN docs for BM_CLICK (and au3 author says it applies to this situation also):
		// "If the button is in a dialog box and the dialog box is not active, the BM_CLICK message
		// might fail. To ensure success in this situation, call the SetActiveWindow function to activate
		// the dialog box before sending the BM_CLICK message to the button."
		ATTACH_THREAD_INPUT
		SetActiveWindow(target_window);
		if (!GetWindowRect(control_window, &rect))	// au3: Code to primary click the centre of the control
			rect.bottom = rect.left = rect.right = rect.top = 0;
		lparam = MAKELPARAM((rect.right - rect.left) / 2, (rect.bottom - rect.top) / 2);
		PostMessage(control_window, WM_LBUTTONDOWN, MK_LBUTTON, lparam);
		PostMessage(control_window, WM_LBUTTONUP, 0, lparam);
		DETACH_THREAD_INPUT
		break;
	}

	case CONTROL_CMD_ENABLE:
		EnableWindow(control_window, TRUE);
		break;

	case CONTROL_CMD_DISABLE:
		EnableWindow(control_window, FALSE);
		break;

	case CONTROL_CMD_SHOW:
		ShowWindow(control_window, SW_SHOWNOACTIVATE); // SW_SHOWNOACTIVATE is what au3 uses.
		break;

	case CONTROL_CMD_HIDE:
		ShowWindow(control_window, SW_HIDE);
		break;

	case CONTROL_CMD_SHOWDROPDOWN:
	case CONTROL_CMD_HIDEDROPDOWN:
		// CB_SHOWDROPDOWN: Although the return value (dwResult) is always TRUE, SendMessageTimeout()
		// will return failure if it times out:
		if (!SendMessageTimeout(control_window, CB_SHOWDROPDOWN
			, (WPARAM)(control_cmd == CONTROL_CMD_SHOWDROPDOWN ? TRUE : FALSE)
			, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return OK;  // Let ErrorLevel tell the story.
		break;

	case CONTROL_CMD_TABLEFT:
	case CONTROL_CMD_TABRIGHT: // must be a Tab Control
		key_count = *aValue ? ATOI(aValue) : 1;
		vk = (control_cmd == CONTROL_CMD_TABLEFT) ? VK_LEFT : VK_RIGHT;
		lparam = (LPARAM)(g_vk_to_sc[vk].a << 16);
		for (int i = 0; i < key_count; ++i)
		{
			// DoControlDelay isn't done for every iteration because it seems likely that
			// the Sleep(0) will take care of things.
			PostMessage(control_window, WM_KEYDOWN, vk, lparam | 0x00000001);
			SLEEP_WITHOUT_INTERRUPTION(0); // Au3 uses a Sleep(0).
			PostMessage(control_window, WM_KEYUP, vk, lparam | 0xC0000001);
		}
		break;

	case CONTROL_CMD_ADD:
		if (!strnicmp(aControl, "Combo", 5))
			msg = CB_ADDSTRING;
		else if (!strnicmp(aControl, "List", 4))
			msg = LB_ADDSTRING;
		else
			return OK;  // Must be ComboBox or ListBox.  Let ErrorLevel tell the story.
		if (!SendMessageTimeout(control_window, msg, 0, (LPARAM)aValue, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return OK;  // Let ErrorLevel tell the story.
		if (dwResult == CB_ERR || dwResult == CB_ERRSPACE) // General error or insufficient space to store it.
			// CB_ERR == LB_ERR
			return OK;  // Let ErrorLevel tell the story.
		break;

	case CONTROL_CMD_DELETE:
		if (!*aValue)
			return OK;
		control_index = ATOI(aValue) - 1;
		if (control_index < 0)
			return OK;
		if (!strnicmp(aControl, "Combo", 5))
			msg = CB_DELETESTRING;
		else if (!strnicmp(aControl, "List", 4))
			msg = LB_DELETESTRING;
		else
			return OK;  // Must be ComboBox or ListBox.  Let ErrorLevel tell the story.
		if (!SendMessageTimeout(control_window, msg, (WPARAM)control_index, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return OK;  // Let ErrorLevel tell the story.
		if (dwResult == CB_ERR)  // CB_ERR == LB_ERR
			return OK;  // Let ErrorLevel tell the story.
		break;

	case CONTROL_CMD_CHOOSE:
		if (!*aValue)
			return OK;
		control_index = ATOI(aValue) - 1;
		if (control_index < 0)
			return OK;  // Let ErrorLevel tell the story.
		if (!strnicmp(aControl, "Combo", 5))
		{
			msg = CB_SETCURSEL;
			x_msg = CBN_SELCHANGE;
			y_msg = CBN_SELENDOK;
		}
		else if (!strnicmp(aControl, "List" ,4))
		{
			msg = LB_SETCURSEL;
			x_msg = LBN_SELCHANGE;
			y_msg = LBN_DBLCLK;
		}
		else
			return OK;  // Must be ComboBox or ListBox.  Let ErrorLevel tell the story.
		if (!SendMessageTimeout(control_window, msg, control_index, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return OK;  // Let ErrorLevel tell the story.
		if (dwResult == CB_ERR)  // CB_ERR == LB_ERR
			return OK;
		if (   !(immediate_parent = GetParent(control_window))   )
			return OK;
		if (   !(control_id = GetDlgCtrlID(control_window))   )
			return OK;
		if (!SendMessageTimeout(immediate_parent, WM_COMMAND, (WPARAM)MAKELONG(control_id, x_msg)
			, (LPARAM)control_window, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return OK;
		if (!SendMessageTimeout(immediate_parent, WM_COMMAND, (WPARAM)MAKELONG(control_id, y_msg)
			, (LPARAM)control_window, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return OK;
		// Otherwise break and do the end-function processing.
		break;

	case CONTROL_CMD_CHOOSESTRING:
		if (!strnicmp(aControl, "ComboBox",8))
		{
			msg = CB_SELECTSTRING;
			x_msg = CBN_SELCHANGE;
			y_msg = CBN_SELENDOK;
		}
		else if (!strnicmp(aControl, "ListBox", 7))
		{
			msg = LB_SELECTSTRING;
			x_msg = LBN_SELCHANGE;
			y_msg = LBN_DBLCLK;
		}
		else
			return OK;  // Must be ComboBox or ListBox.  Let ErrorLevel tell the story.
		if (!SendMessageTimeout(control_window, msg, 1, (LPARAM)aValue, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return OK;  // Let ErrorLevel tell the story.
		if (dwResult == CB_ERR)  // CB_ERR == LB_ERR
			return OK;
		if (   !(immediate_parent = GetParent(control_window))   )
			return OK;
		if (   !(control_id = GetDlgCtrlID(control_window))   )
			return OK;
		if (!SendMessageTimeout(immediate_parent, WM_COMMAND, (WPARAM)MAKELONG(control_id, x_msg)
			, (LPARAM)control_window, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return OK;
		if (!SendMessageTimeout(immediate_parent, WM_COMMAND, (WPARAM)MAKELONG(control_id, y_msg)
			, (LPARAM)control_window, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return OK;
		// Otherwise break and do the end-function processing.
		break;

	case CONTROL_CMD_EDITPASTE:
		if (!SendMessageTimeout(control_window, EM_REPLACESEL, TRUE, (LPARAM)aValue, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return OK;  // Let ErrorLevel tell the story.
		// Note: dwResult is not used by EM_REPLACESEL since it doesn't return a value.
		break;
	} // switch()

	DoControlDelay;  // Seems safest to do this for all of these commands.
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
}



ResultType Line::ControlGet(char *aCmd, char *aValue, char *aControl, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
// This function has been adapted from the AutoIt3 source.
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR);  // Set default since there are many points of return.
	ControlGetCmds control_cmd = ConvertControlGetCmd(aCmd);
	// Since command names are validated at load-time, this only happens if the command name
	// was contained in a variable reference.  Since that is very rare, just set ErrorLevel
	// and return:
	if (control_cmd == CONTROLGET_CMD_INVALID)
		return output_var->Assign();  // Let ErrorLevel tell the story.

	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return output_var->Assign();  // Let ErrorLevel tell the story.
	HWND control_window = ControlExist(target_window, aControl);
	if (!control_window)
		return output_var->Assign();  // Let ErrorLevel tell the story.

	DWORD dwResult, index, length, start, end;
	UINT msg, x_msg, y_msg;
	int control_index;
	char *dyn_buf, buf[32768];  // 32768 is the size Au3 uses for GETLINE and such.

	switch(control_cmd)
	{
	case CONTROLGET_CMD_CHECKED: //Must be a Button
		if (!SendMessageTimeout(control_window, BM_GETCHECK, 0, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return output_var->Assign();
		output_var->Assign(dwResult == BST_CHECKED ? "1" : "0");
		break;

	case CONTROLGET_CMD_ENABLED:
		output_var->Assign(IsWindowEnabled(control_window) ? "1" : "0");
		break;

	case CONTROLGET_CMD_VISIBLE:
		output_var->Assign(IsWindowVisible(control_window) ? "1" : "0");
		break;

	case CONTROLGET_CMD_TAB: // must be a Tab Control
		if (!SendMessageTimeout(control_window, TCM_GETCURSEL, 0, 0, SMTO_ABORTIFHUNG, 2000, &index))
			return output_var->Assign();
		if (index == -1)
			return output_var->Assign();
		output_var->Assign(index + 1);
		break;

	case CONTROLGET_CMD_FINDSTRING:
		if (!strnicmp(aControl, "Combo", 5))
			msg = CB_FINDSTRINGEXACT;
		else if (!strnicmp(aControl, "List", 4))
			msg = LB_FINDSTRINGEXACT;
		else // Must be ComboBox or ListBox
			return output_var->Assign();  // Let ErrorLevel tell the story.
		if (!SendMessageTimeout(control_window, msg, 1, (LPARAM)aValue, SMTO_ABORTIFHUNG, 2000, &index))
			return output_var->Assign();
		if (index == CB_ERR)  // CB_ERR == LB_ERR
			return output_var->Assign();
		output_var->Assign(index + 1);
		break;

	case CONTROLGET_CMD_CHOICE:
		if (!strnicmp(aControl, "ComboBox", 8))
		{
			msg = CB_GETCURSEL;
			x_msg = CB_GETLBTEXTLEN;
			y_msg = CB_GETLBTEXT;
		}
		else if (!strnicmp(aControl, "ListBox" ,7))
		{
			msg = LB_GETCURSEL;
			x_msg = LB_GETTEXTLEN;
			y_msg = LB_GETTEXT;
		}
		else // Must be ComboBox or ListBox
			return output_var->Assign();  // Let ErrorLevel tell the story.
		if (!SendMessageTimeout(control_window, msg, 0, 0, SMTO_ABORTIFHUNG, 2000, &index))
			return output_var->Assign();
		if (index == CB_ERR)  // CB_ERR == LB_ERR
			return output_var->Assign();
		if (!SendMessageTimeout(control_window, x_msg, (WPARAM)index, 0, SMTO_ABORTIFHUNG, 2000, &length))
			return output_var->Assign();
		if (length == CB_ERR)  // CB_ERR == LB_ERR
			return output_var->Assign();
		++length;
		if (   !(dyn_buf = (char *)calloc(256 + length, 1))   )
			return output_var->Assign();
		if (!SendMessageTimeout(control_window, y_msg, (WPARAM)index, (LPARAM)dyn_buf, SMTO_ABORTIFHUNG, 2000, &length))
		{
			free(dyn_buf);
			return output_var->Assign();
		}
		if (length == CB_ERR)  // CB_ERR == LB_ERR
		{
			free(dyn_buf);
			return output_var->Assign();
		}
		output_var->Assign(dyn_buf);
		free(dyn_buf);
		break;

	case CONTROLGET_CMD_LINECOUNT:  //Must be an Edit
		// MSDN: "If the control has no text, the return value is 1. The return value will never be less than 1."
		if (!SendMessageTimeout(control_window, EM_GETLINECOUNT, 0, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return output_var->Assign();
		output_var->Assign(dwResult);
		break;

	case CONTROLGET_CMD_CURRENTLINE:
		if (!SendMessageTimeout(control_window, EM_LINEFROMCHAR, -1, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return output_var->Assign();
		output_var->Assign(dwResult + 1);
		break;

	case CONTROLGET_CMD_CURRENTCOL:
	{
		if (!SendMessageTimeout(control_window, EM_GETSEL, (WPARAM)&start, (LPARAM)&end, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return output_var->Assign();
		// The dwResult from the above is not useful and is not checked.
		DWORD line_number;
		if (!SendMessageTimeout(control_window, EM_LINEFROMCHAR, (WPARAM)start, 0, SMTO_ABORTIFHUNG, 2000, &line_number))
			return output_var->Assign();
		if (!line_number) // Since we're on line zero, the column number is simply start+1.
		{
			output_var->Assign(start + 1);  // +1 to convert from zero based.
			break;
		}
		// Au3: Decrement the character index until the row changes.  Difference between this
		// char index and original is the column:
		DWORD start_orig = start;  // Au3: the character index
		for (;;)
		{
			if (!SendMessageTimeout(control_window, EM_LINEFROMCHAR, (WPARAM)start, 0, SMTO_ABORTIFHUNG, 2000, &dwResult))
				return output_var->Assign();
			if (dwResult != line_number)
				break;
			--start;
		}
		output_var->Assign((int)(start_orig - start));
		break;
	}

	case CONTROLGET_CMD_LINE:
		if (!*aValue)
			return output_var->Assign();
		control_index = ATOI(aValue) - 1;
		if (control_index < 0)
			return output_var->Assign();  // Let ErrorLevel tell the story.
		*((LPINT)buf) = sizeof(buf);  // EM_GETLINE requires first word of string to be set to its size.
		if (!SendMessageTimeout(control_window, EM_GETLINE, (WPARAM)control_index, (LPARAM)buf, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return output_var->Assign();
		if (!dwResult) // due to the specified line number being greater than the number of lines in the edit control.
			return output_var->Assign();
		buf[dwResult] = '\0'; // Ensure terminated since the API might not do it in some cases.
		output_var->Assign(buf);
		break;

	case CONTROLGET_CMD_SELECTED: //Must be an Edit
		// Note: The RichEdit controls of certain apps such as Metapad don't return the right selection
		// with this technique.  Au3 has the same problem with them, so for now it's just documented here
		// as a limitation.
		if (!SendMessageTimeout(control_window, EM_GETSEL, (WPARAM)&start,(LPARAM)&end, SMTO_ABORTIFHUNG, 2000, &dwResult))
			return output_var->Assign();
		// The above sets start to be the zero-based position of the start of the selection (similar for end).
		// If there is no selection, start and end will be equal, at least in the edit controls I tried it with.
		// The dwResult from the above is not useful and is not checked.
		if (start == end) // Unlike Au3, it seems best to consider a blank selection to be a non-error.
		{
			output_var->Assign();
			break;
		}
		if (!SendMessageTimeout(control_window, WM_GETTEXTLENGTH, 0, 0, SMTO_ABORTIFHUNG, 2000, &length))
			return output_var->Assign();
		if (!length)
			// Since the above didn't return for start == end, this is an error because
			// we have a selection of non-zero length, but no text to go with it!
			return output_var->Assign();
		if (   !(dyn_buf = (char *)calloc(256 + length, 1))   )
			return output_var->Assign();
		if (!SendMessageTimeout(control_window, WM_GETTEXT, (WPARAM)(length + 1), (LPARAM)dyn_buf, SMTO_ABORTIFHUNG, 2000, &length))
		{
			free(dyn_buf);
			return output_var->Assign();
		}
		if (!length || end > length)
		{
			// The first check above is reveals a problem (ErrorLevel = 1) since the length
			// is unexpectedly zero (above implied it shouldn't be).  The second check is also
			// a problem because the end of the selection should not be beyond length of text
			// that was retrieved.
			free(dyn_buf);
			return output_var->Assign();
		}
		dyn_buf[end] = '\0'; // Terminate the string at the end of the selection.
		output_var->Assign(dyn_buf + start);
		free(dyn_buf);
		break;
	}

	// Note that ControlDelay is not done for the Get type commands, because it seems unnecessary.
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
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
	StatusBarUtil(output_var, control_window, ATOI(aPart)); // It will handle any zero part# for us.
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
	StatusBarUtil(NULL, control_window, ATOI(aPart) // It will handle a NULL control_window or zero part# for us.
		, text_to_wait_for, *aSeconds ? (int)(ATOF(aSeconds)*1000) : -1 // Blank->indefinite.  0 means 500ms.
		, ATOI(aInterval));
	return OK; // Even if it fails, seems best to return OK so that subroutine can continue.
}



ResultType Line::ScriptPostMessage(char *aMsg, char *awParam, char *alParam, char *aControl
	, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	HWND control_window = *aControl ? ControlExist(target_window, aControl) : target_window;
	if (!control_window)
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	// Use ATOI64 to support unsigned (i.e. UINT, LPARAM, and WPARAM are all 32-bit unsigned values).
	// ATOI64 also supports hex strings in the script, such as 0xFF, which is why it's commonly
	// used in functions such as this:
	g_ErrorLevel->Assign(PostMessage(control_window, (UINT)ATOI64(aMsg), (LPARAM)ATOI64(awParam)
		, (WPARAM)ATOI64(alParam)) ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);
	// By design (since this is a power user feature), no ControlDelay is done here.
	return OK;
}



ResultType Line::ScriptSendMessage(char *aMsg, char *awParam, char *alParam, char *aControl
	, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return g_ErrorLevel->Assign("FAIL"); // Need a special value to distinguish this from numeric reply-values.
	HWND control_window = *aControl ? ControlExist(target_window, aControl) : target_window;
	if (!control_window)
		return g_ErrorLevel->Assign("FAIL");
	// Use ATOI64 to support unsigned (i.e. UINT, LPARAM, and WPARAM are all 32-bit unsigned values).
	// ATOI64 also supports hex strings in the script, such as 0xFF, which is why it's commonly
	// used in functions such as this:
	DWORD dwResult;
	if (!SendMessageTimeout(control_window, (UINT)ATOI64(aMsg), (WPARAM)ATOI64(awParam), (LPARAM)ATOI64(alParam)
		, SMTO_ABORTIFHUNG, 2000, &dwResult))
		return g_ErrorLevel->Assign("FAIL"); // Need a special value to distinguish this from numeric reply-values.
	// By design (since this is a power user feature), no ControlDelay is done here.
	return g_ErrorLevel->Assign(dwResult); // UINT seems best most of the time?
}



ResultType Line::WinSet(char *aAttrib, char *aValue, char *aTitle, char *aText
	, char *aExcludeTitle, char *aExcludeText)
{
	WinSetAttributes attrib = ConvertWinSetAttribute(aAttrib);
	if (attrib == WINSET_INVALID)
		return LineError(ERR_WINSET, FAIL, aAttrib);

	// Since this is a macro, avoid repeating it for every case of the switch():
	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return OK;

	int value;
	int target_window_long;

	switch (attrib)
	{
	case WINSET_ALWAYSONTOP:
	{
		if (   !(target_window_long = GetWindowLong(target_window, GWL_EXSTYLE))   )
			return OK;
		HWND topmost_or_not;
		switch(ConvertOnOffToggle(aValue))
		{
		case TOGGLED_ON: topmost_or_not = HWND_TOPMOST; break;
		case TOGGLED_OFF: topmost_or_not = HWND_NOTOPMOST; break;
		case NEUTRAL: // parameter was blank, so it defaults to TOGGLE.
		case TOGGLE: topmost_or_not = (target_window_long & WS_EX_TOPMOST) ? HWND_NOTOPMOST : HWND_TOPMOST; break;
		default: return OK;
		}
		// SetWindowLong() didn't seem to work, at least not on some windows.  But this does:
		SetWindowPos(target_window, topmost_or_not, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
		break;
	}

	case WINSET_TRANSPARENT:
		if (!g_os.IsWin2000orLater())
			return OK;  // Do nothing on OSes that don't support it.
		value = ATOI(aValue);
		// A little debatable, but this behavior seems best, at least in some cases:
		if (value < 0)
			value = 0;
		else if (value > 255)
			value = 255;
		// Must fetch it at runtime, otherwise the program can't even be launched on Win9x/NT:
		typedef BOOL (WINAPI *MySetLayeredWindowAttributesType)(HWND, COLORREF, BYTE, DWORD);
		static MySetLayeredWindowAttributesType MySetLayeredWindowAttributes = (MySetLayeredWindowAttributesType)
			GetProcAddress(GetModuleHandle("User32.dll"), "SetLayeredWindowAttributes");
		if (MySetLayeredWindowAttributes)
		{
			// NOTE: It seems best never to remove the WS_EX_LAYERED attribute, even if the value is 255
			// (non-transparent), since the window might have had that attribute previously and may need
			// it to function properly.  For example, an app may support making its own windows transparent
			// but might not expect to have to turn WS_EX_LAYERED back on if we turned it off.  One drawback
			// of this is a quote from somewhere that might or might not be accurate: "To make this window
			// completely opaque again, remove the WS_EX_LAYERED bit by calling SetWindowLong and then ask
			// the window to repaint. Removing the bit is desired to let the system know that it can free up
			// some memory associated with layering and redirection."
			SetWindowLong(target_window, GWL_EXSTYLE, GetWindowLong(target_window, GWL_EXSTYLE) | WS_EX_LAYERED);
			MySetLayeredWindowAttributes(target_window, 0, value, LWA_ALPHA);
		}
		break;
	}
	return OK;
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



ResultType Line::WinGetClass(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText)
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	DETERMINE_TARGET_WINDOW
	if (!target_window)
		return output_var->Assign();
	char class_name[1024];
	if (!GetClassName(target_window, class_name, sizeof(class_name)))
		return output_var->Assign();
	return output_var->Assign(class_name);
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
		return LineError(ERR_WINGET ERR_ABORT, FAIL, aCmd);

	bool target_window_determined = true;  // Set default.
	HWND target_window;
	IF_USE_FOREGROUND_WINDOW(aTitle, aText, aExcludeTitle, aExcludeText)
	else if (!*aTitle && !*aText && !*aExcludeTitle && !*aExcludeText)
		target_window = g_ValidLastUsedWindow;
	else
		target_window_determined = false;  // A different method is required.

	// Used with WINGET_CMD_LIST to create an array (if needed).  Make it longer than Max var name
	// so that FindOrAddVar() will be able to spot and report var names that are too long:
	char var_name[MAX_VAR_NAME_LENGTH + 20];
	Var *array_item;
	WindowInfoPackage wip;

	switch(cmd)
	{
	case WINGET_CMD_ID:
	case WINGET_CMD_IDLAST:
		if (target_window_determined)
		{
			if (target_window)
				return output_var->AssignHWND(target_window);
			else
				return output_var->Assign();
		}
		// Otherwise:
		if (   !(target_window = WinExist(aTitle, aText, aExcludeTitle, aExcludeText, cmd == WINGET_CMD_IDLAST))   )
			return output_var->Assign();
		return output_var->AssignHWND(target_window);

	case WINGET_CMD_COUNT:
		if (target_window_determined) // target_window (if non-NULL) represents exactly 1 window in this case.
			return output_var->Assign(target_window ? "1" : "0");
		// Otherwise, have WinExist() get the count for us:
		return output_var->Assign((DWORD)WinExist(aTitle, aText, aExcludeTitle, aExcludeText, true, false, NULL, 0, true));

	case WINGET_CMD_LIST:
		// Retrieves a list of HWNDs for the windows that match the given criteria and stores them in
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
			// Otherwise, since the target window has been determined, we know that it is
			// the only window to be put into the array:
			snprintf(var_name, sizeof(var_name), "%s1", output_var->mName);
			if (   !(array_item = g_script.FindOrAddVar(var_name))   )  // Find or create element #1.
				return FAIL;  // It will have already displayed the error.
			if (!array_item->AssignHWND(target_window))
				return FAIL;
			return output_var->Assign("1");  // 1 window found
		}
		// Otherwise, the target window(s) have not yet been determined and a special method
		// is required to gather them.
		wip.find_last_match = true;
		wip.array_start = output_var;  // Give it the position in the var list of where the array will be.
		strlcpy(wip.title, aTitle, sizeof(wip.title));
		strlcpy(wip.text, aText, sizeof(wip.text));
		strlcpy(wip.exclude_title, aExcludeTitle, sizeof(wip.exclude_title));
		strlcpy(wip.exclude_text, aExcludeText, sizeof(wip.exclude_text));
		EnumWindows(EnumParentFind, (LPARAM)&wip);
		return output_var->Assign(wip.match_count);
	}

	return FAIL;  // Never executed (increases maintainability and avoids compiler warning).
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



ResultType Line::PixelSearch(int aLeft, int aTop, int aRight, int aBottom, int aColor, int aVariation)
{
	Var *output_var_x = ResolveVarOfArg(0);  // Ok if NULL.
	Var *output_var_y = ResolveVarOfArg(1);  // Ok if NULL.

	g_ErrorLevel->Assign(ERRORLEVEL_ERROR2); // Set default ErrorLevel.  2 means error other than "color not found".
	if (output_var_x)
		output_var_x->Assign();  // Init to empty string regardless of whether we succeed here.
	if (output_var_y)
		output_var_y->Assign(); // Same.

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
				// Adjust coords to make them relative to the position of the target window
				// (rect will contain zero values if this doesn't need to be done):
				xpos -= rect.left;
				ypos -= rect.top;
				if (output_var_x && !output_var_x->Assign(xpos))
					result = FAIL;
				if (output_var_y && !output_var_y->Assign(ypos))
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



//ResultType Line::ImageSearch(int aLeft, int aTop, int aRight, int aBottom, char *aImageFile)
//{
//	Var *output_var_x = ResolveVarOfArg(0);  // Ok if NULL.
//	Var *output_var_y = ResolveVarOfArg(1);  // Ok if NULL.
//
//	g_ErrorLevel->Assign(ERRORLEVEL_ERROR2); // Set default ErrorLevel.  2 means error other than "color not found".
//	if (output_var_x)
//		output_var_x->Assign();  // Init to empty string regardless of whether we succeed here.
//	if (output_var_y)
//		output_var_y->Assign(); // Same.
//
//	// Always adjust coords to reflect the position of the foreground window because AutoHotkey
//	// doesn't yet support AutoIt3's absolute-screen-coords mode:
//	RECT rect;
//	HWND cw=GetForegroundWindow();
//	if (!GetWindowRect(GetForegroundWindow(), &rect))
//		return OK;  // Let ErrorLevel tell the story.
//	aLeft   += rect.left;
//	aTop    += rect.top;
//	aRight  += rect.left;  // Add left vs. right because we're adjusting based on the position of the window.
//	aBottom += rect.top;   // Same.
//
//	HDC hdc = GetDC(NULL);
//	if (!hdc)
//		return OK;  // Let ErrorLevel tell the story.
//
//	int aWidth,aHeight;
//	aWidth = aRight - aLeft +1;
//	aHeight = aBottom - aTop +1;
//
//	int xpos, ypos;
//	ResultType result = OK;
//	
//	bool is16bits = false;
//
//	int i,j,k,x,y,p;
//
//	HBITMAP tmpl;
//	BITMAP bm;
//	long width;	
//	long height;
//	HDC tdc;
//	tmpl=(HBITMAP)::LoadImage(NULL, aImageFile, IMAGE_BITMAP, 0, 0,LR_LOADFROMFILE | LR_CREATEDIBSECTION);
//
//	GetObject(tmpl, sizeof(bm), &bm);
//	width = bm.bmWidth ;	
//	height = bm.bmHeight;
//	
//	//put the image in the HDC context
//	tdc = ::CreateCompatibleDC(hdc);
//	HGDIOBJ oldtdc=::SelectObject(tdc,tmpl);
//	
//
//	BITMAPINFO bmInfo;
//	::memset(&bmInfo,0,sizeof(BITMAPINFO));
//	bmInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
//	int err=::GetDIBits(tdc,tmpl,0,0,0,&bmInfo,DIB_RGB_COLORS);
//	if(err==0){
//			ReleaseDC(NULL, hdc);	
//			ReleaseDC(NULL,tdc);
//			DeleteObject(tmpl);
//			return FAIL;
//	}
//
//
//	if(bmInfo.bmiHeader.biBitCount<16){
//			ReleaseDC(NULL, hdc);	
//			ReleaseDC(NULL, tdc);
//			DeleteObject(tmpl);
//			return FAIL;
//	}
//
//	is16bits = (bmInfo.bmiHeader.biBitCount == 16);
//
//
//	int szb = bmInfo.bmiHeader.biWidth*bmInfo.bmiHeader.biHeight;
//	DWORD* bits=new DWORD[szb];              
//	::memset(bits,0xAA,szb*4);
//	bmInfo.bmiHeader.biBitCount = 32;
//	bmInfo.bmiHeader.biHeight = -bmInfo.bmiHeader.biHeight;
//
//	err=::GetDIBits(tdc,tmpl,0,-bmInfo.bmiHeader.biHeight,bits,&bmInfo,DIB_RGB_COLORS);
//	bmInfo.bmiHeader.biHeight = -bmInfo.bmiHeader.biHeight;
//	if(err==0){
//			ReleaseDC(NULL, hdc);	
//			ReleaseDC(NULL,tdc);
//			if(bits)delete bits;
//			DeleteObject(tmpl);
//			return FAIL;
//	}
//
//	//get the screen
//
//	HDC sdc = ::CreateCompatibleDC(hdc);
//	HBITMAP scr = ::CreateCompatibleBitmap(hdc,aWidth,aHeight);
//
//	HGDIOBJ old = ::SelectObject(sdc,scr);
//
//
//	::BitBlt(sdc,0,0,aWidth,aHeight,hdc,aLeft,aTop,SRCCOPY);
//	BITMAPINFO bmScr;
//	::memset(&bmScr,0,sizeof(BITMAPINFO));
//	bmScr.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
//	err=::GetDIBits(sdc,scr,0,0,0,&bmScr,DIB_RGB_COLORS);
//	if(err==0){
//				ReleaseDC(NULL, hdc);	
//				ReleaseDC(NULL,sdc);
//				ReleaseDC(NULL,tdc);
//				if(bits)delete bits;
//				DeleteObject(scr);
//				DeleteObject(tmpl);
//				return FAIL;
//	}
//	int szs = bmScr.bmiHeader.biWidth*bmScr.bmiHeader.biHeight;
//	DWORD* bScr=new DWORD[szs];              
//	::memset(bScr,0xAA,szs*4);
//	bmScr.bmiHeader.biHeight = -bmScr.bmiHeader.biHeight;
//
//	is16bits = (bmScr.bmiHeader.biBitCount == 16);
//
//
//	bmScr.bmiHeader.biBitCount = 32;
//
//	err=::GetDIBits(sdc,scr,0,-bmScr.bmiHeader.biHeight,bScr,&bmScr,DIB_RGB_COLORS);
//	bmScr.bmiHeader.biHeight = -bmScr.bmiHeader.biHeight;
//	if(err==0){
//				ReleaseDC(NULL, hdc);
//				ReleaseDC(NULL,sdc);
//				ReleaseDC(NULL,tdc);
//				if(bits)delete bits;
//				if(bScr)delete bScr;
//				DeleteObject(scr);
//				DeleteObject(tmpl);
//				return FAIL;
//	}
//
//	if(is16bits){
//		for(i=0;i<szs;i++)bScr[i]&=0xF8F8F8F8;
//		for(i=0;i<szb;i++)bits[i]&=0xF8F8F8F8;
//	}
//
//	for (i=0;i <szs; i++)
//	{
//		
//			if (bScr[i]==bits[0]) // This pixel matches one of the specified color(s).
//			{
//				bool found = true;
//
//				for(y=0,j=0,k=i;y<bmInfo.bmiHeader.biHeight && k<szs && found;y++,k+=bmScr.bmiHeader.biWidth){
//					for(x=0,p=k;x<bmInfo.bmiHeader.biWidth && p<szs && found ;x++,j++,p++){
//						if(bScr[p]==bits[j]){
//							found = true;
//						}else{
//							found = false;
//						}
//					}
//				}
//				if(found){
//					ReleaseDC(NULL, hdc);
//					
//					ReleaseDC(NULL,sdc);
//					ReleaseDC(NULL,tdc);
//					if(bits)delete bits;
//					if(bScr)delete bScr;
//					DeleteObject(scr);
//					DeleteObject(tmpl);
//
//					//calculate xpos,ypos;	
//					xpos = aLeft + i%bmScr.bmiHeader.biWidth;
//					ypos = aTop + i/bmScr.bmiHeader.biWidth;
//
//					if (output_var_x)
//						// Adjust coords to make them relative to the position of the target window:
//						if (!output_var_x->Assign(xpos - rect.left))
//							result = FAIL;
//						if (output_var_y)
//							if (!output_var_y->Assign(ypos - rect.top))
//								result = FAIL;
//						if (result == OK)
//							g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
//					return result;
//				}
//			}
//	}
//	
//	// If the above didn't return, the pixel wasn't found in the specified region.
//	// So leave ErrorLevel set to "error" to indicate that:
//
//	ReleaseDC(NULL, hdc);
//					
//	ReleaseDC(NULL,sdc);
//	ReleaseDC(NULL,tdc);
//	if(bits)delete bits;
//	if(bScr)delete bScr;
//	DeleteObject(scr);
//	DeleteObject(tmpl);
//	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // This value indicates "color not found".
//	return OK;
//}



ResultType Line::PixelGetColor(int aX, int aY)
// This has been adapted from the AutoIt3 source.
{
	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	output_var->Assign(); // Init to empty string regardless of whether we succeed here.

	if (!(g.CoordMode & COORD_MODE_PIXEL)) // Using relative vs. screen coordinates.
	{
		// Convert from relative to absolute (screen) coordinates:
		RECT rect;
		if (!GetWindowRect(GetForegroundWindow(), &rect))
			return OK;  // Let ErrorLevel tell the story.
		aX += rect.left;
		aY += rect.top;
	}

	HDC hdc = GetDC(NULL);
	if (!hdc)
		return OK;  // Let ErrorLevel tell the story.
	// Assign the value as an 32-bit int since I believe that's how Window Spy reports color values:
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
    char buf_temp[2048];  // For various uses.
	int i;

	TRANSLATE_AHK_MSG(iMsg, wParam)
	
	switch (iMsg)
	{
	case WM_COMMAND: // If an application processes this message, it should return zero.
		// See if an item was selected from the tray menu or main menu:
		switch (LOWORD(wParam))
		{
		case ID_TRAY_OPEN:
			ShowMainWindow();
			return 0;
		case ID_TRAY_EDITSCRIPT:
		case ID_FILE_EDITSCRIPT:
			g_script.Edit();
			return 0;
		case ID_TRAY_RELOADSCRIPT:
		case ID_FILE_RELOADSCRIPT:
			if (!g_script.Reload(false))
				MsgBox("The script could not be reloaded." PLEASE_REPORT);
			return 0;
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
			return 0;
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
			return 0;
		case ID_TRAY_SUSPEND:
		case ID_FILE_SUSPEND:
			Line::ToggleSuspendState();
			return 0;
		case ID_TRAY_PAUSE:
		case ID_FILE_PAUSE:
			if (!g.IsPaused && g_nThreads < 1)
			{
				MsgBox("The script cannot be paused while it is doing nothing.  If you wish to prevent new"
					" hotkey subroutines from running, use Suspend instead.");
				// i.e. we don't want idle scripts to ever be in a paused state.
				return 0;
			}
			if (g.IsPaused)
				--g_nPausedThreads;
			else
				++g_nPausedThreads;
			g.IsPaused = !g.IsPaused;
			g_script.UpdateTrayIcon();
			CheckMenuItem(GetMenu(g_hWnd), ID_FILE_PAUSE, g.IsPaused ? MF_CHECKED : MF_UNCHECKED);
			return 0;
		case ID_TRAY_EXIT:
		case ID_FILE_EXIT:
			g_script.ExitApp(EXIT_MENU);  // More reliable than PostQuitMessage(), which has been known to fail in rare cases.
			return 0; // If there is an OnExit subroutine, the above might not actually exit.
		case ID_VIEW_LINES:
			ShowMainWindow(MAIN_MODE_LINES);
			return 0;
		case ID_VIEW_VARIABLES:
			ShowMainWindow(MAIN_MODE_VARS);
			return 0;
		case ID_VIEW_HOTKEYS:
			ShowMainWindow(MAIN_MODE_HOTKEYS);
			return 0;
		case ID_VIEW_KEYHISTORY:
			ShowMainWindow(MAIN_MODE_KEYHISTORY);
			return 0;
		case ID_VIEW_REFRESH:
			ShowMainWindow(MAIN_MODE_REFRESH);
			return 0;
		case ID_HELP_WEBSITE:
			if (!g_script.ActionExec("http://www.autohotkey.com", "", NULL, false))
				MsgBox("Could not open URL http://www.autohotkey.com in default browser.");
			return 0;
		case ID_HELP_EMAIL:
			if (!g_script.ActionExec("mailto:support@autohotkey.com", "", NULL, false))
				MsgBox("Could not open URL mailto:support@autohotkey.com in default e-mail client.");
			return 0;
		default:
		{
			// See if this command ID is one of the user's custom menu items.  Due to the possibility
			// that some items have been deleted from the menu, can't rely on comparing
			// LOWORD(wParam) to g_script.mMenuItemCount in any way.  Just look up the ID to make sure
			// there really is a menu item for it:
			if (g_script.FindMenuItemByID(LOWORD(wParam)))
			{
				// It seems best to treat the selection of a custom menu item in a way similar
				// to how hotkeys are handled by the hook.
				// Post it to the thread, just in case the OS tries to be "helpful" and
				// directly call the WindowProc (i.e. this function) rather than actually
				// posting the message.  We don't want to be called, we want the main loop
				// to handle this message:
				#define HANDLE_USER_MENU(menu_id) \
				{\
					POST_AHK_USER_MENU(menu_id) \
					MsgSleep(-1);\
				}
				HANDLE_USER_MENU(LOWORD(wParam))  // Send the menu's cmd ID.
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
				//    here directly in most such cases.  INIT_NEW_THREAD: Newly launched
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
				//	return 0;  // Leave the message buffered until later.
				// Now call the main loop to handle the message we just posted (and any others):
				return 0;
			}
			// else do nothing, let DefWindowProc() try to handle it.
		}
		} // Inner switch()
		break;

	case AHK_NOTIFYICON:  // Tray icon clicked on.
	{
        switch(lParam)
        {
// Don't allow the main window to be opened this way by a compiled EXE, since it will display
// the lines most recently executed, and many people who compile scripts don't want their users
// to see the contents of the script:
		case WM_LBUTTONDBLCLK:
			if (g_script.mTrayMenu->mDefault)
				HANDLE_USER_MENU(g_script.mTrayMenu->mDefault->mMenuID)
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
		case WM_RBUTTONDOWN:
			g_script.mTrayMenu->Display(false);
			return 0;
		} // Inner switch()
		break;
	} // case AHK_NOTIFYICON

	case WM_ENTERMENULOOP:
		// One of the menus in the menu bar has been displayed, and the we know the user is is still in
		// the menu bar, even moving to different menus and/or menu items, until WM_EXITMENULOOP is received.
		// Note: It seems that when window's menu bar is being displayed/navigated by the user, our thread
		// is tied up in a message loop other than our own.  In other words, it's very similar to the
		// TrackPopupMenuEx() call used to handle the tray menu, which is why g_MenuIsVisible can be used
		// for both types of menus to indicate to MainWindowProc() that timed subroutines should not be
		// checked or allowed to launch during such times:
		g_MenuIsVisible = MENU_VISIBLE_MAIN;
		break; // Let DefWindowProc() handle it from here.

	case WM_EXITMENULOOP:
		g_MenuIsVisible = MENU_VISIBLE_NONE;
		break; // Let DefWindowProc() handle it from here.

	case AHK_DIALOG:  // User defined msg sent from MsgBox() or FileSelectFile().
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

	case WM_HOTKEY: // As a result of this app having previously called RegisterHotkey().
	case AHK_HOOK_HOTKEY:  // Sent from this app's keyboard or mouse hook.
	{
		// Post it to the thread, just in case the OS tries to be "helpful" and
		// directly call the WindowProc (i.e. this function) rather than actually
		// posting the message.  We don't want to be called, we want the main loop
		// to handle this message:
		PostThreadMessage(GetCurrentThreadId(), iMsg, wParam, lParam);
		// The message is posted unconditionally (above), but the processing of it will be deferred
		// if hotkeys are not allowed to be activated right now:
		if (!INTERRUPTIBLE)
		{
			// Used to prevent runaway hotkeys, or too many happening due to key-repeat feature.
			// It can also be used to prevent a call to MsgSleep() from accepting new hotkeys
			// in cases where the caller's activity might be interferred with by the launch
			// of a new hotkey subroutine, such as reading or writing to the clipboard.
			// Also turn on the timer so that if a dialog happens to be displayed currently,
			// the hotkey we just posted won't be delayed until after the dialog is gone,
			// i.e we want this hotkey to fire the moment the script becomes interruptible again.
			// See comments in the WM_TIMER case of this function for more explanation.
			SET_MAIN_TIMER
			return 0;
		}
		if (g_MenuIsVisible == MENU_VISIBLE_POPUP || (iMsg == AHK_HOOK_HOTKEY && lParam && g_hWnd == GetForegroundWindow()))
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
			// this AutoHotkey instance's own main window is active (foreground).  UPDATE:
			// Testing shows that it is not necessary to do this when the MAIN MENU is displayed,
			// so it's safer not to do it in that case:
			SendMessage(hWnd, WM_CANCELMODE, 0, 0);
			// The menu is now gone because the above should have called this function
			// recursively to close the it.  Now, rather than continuing in this
			// recursion layer, it seems best to return to the caller so that the menu
			// will be destroyed and g_MenuIsVisible updated.  After that is done,
			// the next call to MsgSleep() should notice the hotkey we posted above and
			// act upon it.
			// The above section has been tested and seems to work as expected.
			// UPDATE: Below doesn't work if there's a MsgBox() window displayed because
			// the caller to which we return is the MsgBox's msg pump, and that pump
			// ignores any messages for our thread so they just sit there.  So instead
			// of returning, call MsgSleep() without resetting the value of
			// g_MenuIsVisible (so that it can use it).  When MsgSleep() returns,
			// we will return to our caller, which in this case should be TrackPopupMenuEx's
			// msg pump.  That pump should immediately return also since we've already
			// closed the menu.  And we will let it set the value of g_MenuIsVisible
			// to "none" at that time rather than doing it here or in IsCycleComplete().
			// In keeping with the above, don't return:
			//return 0;
		}
		// Now call the main loop to handle the message we just posted (and any others):
		MsgSleep(-1);
		return 0; // Not sure if this is the correct return value.  It probably doesn't matter.
	}

	case WM_TIMER:
		// Since we're here, the receipt of this msg indicates that g_script.mTimerEnabledCount > 0,
		// so it performs a little better to call it directly vs. CHECK_SCRIPT_TIMERS_IF_NEEDED.
		// UPDATE: The WM_HOTKEY case of this function now also turns on the main timer in some
		// cases so that hotkeys that are pressed while the script is both displaying a dialog
		// and is uninterruptible will not get stuck in the buffer until the dialog is dismissed.
		// Therefore, just do a quick msg check, which will also handle the script timers for us:
		//CheckScriptTimers();
		// More explanation:
		// At first I was concerned that the moment ExecUntil() or anyone does a Sleep(10+) (due to
		// SetBatchLines or whatever), the timer would be disabled by IsCycleComplete() upon return.
		// But realistically, I don't think the following conditions can ever be true simultaneously?:
		// 1) The timer is not enabled.
		// 2) The script is not interruptible.
		// 3) A dialog is displayed (it's msg pump is routing hotkey presses to MainWindowProc())
		// I don't think it's possible (or realistically likely) for all of the above to be true, since
		// for example the script can only be uninterruptible while it's idle-waiting-for-dialog
		// (which is the only time hotkey presses get caught by its msg pump rather than ours, and thus
		// get routed to MainWindowProc() rather than get caught directly by MsgSleep()) if there are
		// timed subroutines, in which case the timer is already running anyway.  But that's just it:
		// when there are timed subroutines, hotkeys would get blocked if they happened to be pressed at
		// an instant when the script is uninterruptible during a short timed subroutine (and maybe even
		// a short hotkey subroutine) that never has a chance to do any MsgSleep()'s (and this would be
		// even more likely if the user increased the uninterruptible time via a setting).  That is why
		// MsgSleep() is called in lieu of a direct call to CheckScriptTimers().  So the issue is that
		// it's unlikely that the SET_MAIN_TIMER statement in the WM_HOTKEY case will ever actually turn
		// on the timer since it will already be on in nearly all such cases (i.e. so it's really just a
		// safety check).  And the overhead of calling MsgSleep(-1) vs. CheckScriptTimers() here isn't
		// too much of a worry since scripts don't spend most of their time waiting for dialogs.
		// UPDATE: Do not call MsgSleep() while the tray menu is visible because that causes long delays
		// sometime when the user is trying to select a menu (the user's click is ignored and the menu
		// stays visible).  I think this is because MsgSleep()'s PeekMessage() intercepts the user's
		// clicks and is unable to route them to TrackPopupMenuEx()'s message loop, which is the only
		// place they can be properly processed.  UPDATE: This also needs to be done when the MAIN MENU
		// is visible, because testing shows that that menu would otherwise become sluggish too, perhaps
		// more rarely, when timers are running.  UPDATE: It seems pointless to call MsgSleep() (and harmful
		// in the case where g_MenuIsVisible is true) if the script is not interruptible, so it's now called
		// only when interrupts are possible (this also covers g_MenuIsVisible).
		// Other background info:
		// Checking g_MenuIsVisible here prevents timed subroutines from running while the tray menu
		// or main menu is in use.  This is documented behavior, and is desirable most of the time
		// anyway.  But not to do this would produce strange effects because any timed subroutine
		// that took a long time to run might keep us away from the "menu loop", which would result
		// in the menu becoming temporarily unreponsive while the user is in it (and probably other
		// undesired effects).
		// Also, this allows MainWindowProc() to close the popup menu upon receive of any hotkey,
		// which is probably a good idea since most hotkeys change the foreground window and if that
		// happens, the menu cannot be dismissed (ever?) except by selecting one of the items in the
		// menu (which is often undesirable).
		if (INTERRUPTIBLE)
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
	case WM_ENDSESSION:
	case AHK_EXIT_BY_RELOAD:
	case AHK_EXIT_BY_SINGLEINSTANCE:
		if (hWnd == g_hWnd) // i.e. not the SplashText window or anything other than the main.
		{
			// Receiving this msg is fairly unusual since SC_CLOSE is intercepted and redefined above.
			// However, it does happen if an external app is asking us to close, such as another
			// instance of this same script during the Reload command.  So treat it in a way similar
			// to the user having chosen Exit from the menu.

			// Leave it up to ExitApp() to decide whether to terminate based upon whether
			// there is an OnExit subroutine, whether that subroutine is already running at
			// the time a new WM_CLOSE is received, etc.  It's also its responsibility to call
			// DestroyWindow() upon termination so that the WM_DESTROY message winds up being
			// received and process in this function (which is probably necessary for a clean
			// termination of the app and all its windows):
			switch (iMsg)
			{
			case WM_CLOSE:
				g_script.ExitApp(EXIT_WM_CLOSE);
				break;
			case WM_ENDSESSION: // MSDN: "A window receives this message through its WindowProc function."
				if (wParam) // the session is being ended (otherwise, a prior WM_QUERYENDSESSION was aborted).
					g_script.ExitApp(lParam & ENDSESSION_LOGOFF ? EXIT_LOGOFF : EXIT_SHUTDOWN);
				break;
			case AHK_EXIT_BY_RELOAD:
				g_script.ExitApp(EXIT_RELOAD);
				break;
			case AHK_EXIT_BY_SINGLEINSTANCE:
				g_script.ExitApp(EXIT_SINGLEINSTANCE);
				break;
			}
			return 0;  // Verified correct.
		}
		// Otherwise, some window of ours other than our main window was destroyed
		// (perhaps the splash window).  Let DefWindowProc() handle it:
		break;

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
			return 0;  // Verified correct.
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
			return splash.hbrush ? (LRESULT)splash.hbrush : (LRESULT)GetSysColorBrush(COLOR_BTNFACE);
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
		// Rajat wanted this so that it's possible to discover the PID based on the title of each
		// script's main window (i.e. if there are multiple scripts running).  Note that
		// ReplyMessage() has no effect if our own thread sent this message to us.  In other words,
		// if a script asks itself what its PID is, the answer will be 0.  Also note that this
		// msg uses TRANSLATE_AHK_MSG() to prevent it from ever being filtered out (and thus
		// delayed) while the script is uninterruptible.
		ReplyMessage(GetCurrentProcessId());
		return 0;

	} // end main switch

	// Detect Explorer crashes so that tray icon can be recreated.  I think this only works on Win98
	// and beyond, since the feature was never properly implemented in Win95:
	static UINT WM_TASKBARCREATED = RegisterWindowMessage("TaskbarCreated");
	if (iMsg == WM_TASKBARCREATED && !g_NoTrayIcon)
	{
		g_script.CreateTrayIcon();
		g_script.UpdateTrayIcon(true);  // Force the icon into the correct pause/suspend state.
		// And now pass this iMsg on to DefWindowProc() in case it does anything with it.
	}

	return DefWindowProc(hWnd, iMsg, wParam, lParam);
}



ResultType ShowMainWindow(MainWindowModes aMode, bool aRestricted)
// Always returns OK for caller convenience.
{
	// 32767 might be the limit for an edit control, at least under Win95.
	// Later maybe use the EM_REPLACESEL method to avoid this limit if the
	// OS is WinNT/2k/XP:
	char buf_temp[32767] = "";
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
	*aBuf = '\0';  // Init in case of failure.  Some callers may rely on this.
	HKEY hRegKey;
	if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\\AutoHotkey", 0, KEY_READ, &hRegKey) != ERROR_SUCCESS)
		return FAIL;
	DWORD aBuf_size = MAX_PATH;
	if (RegQueryValueEx(hRegKey, "InstallDir", NULL, NULL, (LPBYTE)aBuf, &aBuf_size) != ERROR_SUCCESS)
	{
		RegCloseKey(hRegKey);
		return FAIL;
	}
	RegCloseKey(hRegKey);
	return OK;
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

	if (aTimeout > 2147483) // This is approximately the max number of seconds that SetTimer can handle.
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

	// Specify NULL as the owner window since we want to be able to have the main window in the foreground
	// even if there are InputBox windows:
	++g_nInputBoxes;
	int result = (int)DialogBox(g_hInstance, MAKEINTRESOURCE(IDD_INPUTBOX), NULL, InputBoxProc);
	--g_nInputBoxes;

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

		// If a non-default size was specified, the box will need to be recentered.  The exception is when
		// an explicit xpos or ypos is specified, in which case centering is disabled for that dimension.
		int new_xpos, new_ypos;
		if (CURR_INPUTBOX.xpos != INPUTBOX_DEFAULT && CURR_INPUTBOX.ypos != INPUTBOX_DEFAULT)
		{
			new_xpos = CURR_INPUTBOX.xpos;
			new_ypos = CURR_INPUTBOX.ypos;
		}
		else
		{
			SystemParametersInfo(SPI_GETWORKAREA, 0, &rect, 0);  // Get Desktop rect excluding task bar.
  			if (CURR_INPUTBOX.xpos == INPUTBOX_DEFAULT) // Center horizontally.
				new_xpos = (rect.right - new_width) / 2;
			else
				new_xpos = CURR_INPUTBOX.xpos;
  			if (CURR_INPUTBOX.ypos == INPUTBOX_DEFAULT) // Center vertically.
				new_ypos = (rect.bottom - new_height) / 2;
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



///////////////////
// Mouse related //
///////////////////

ResultType Line::MouseClickDrag(vk_type aVK, int aX1, int aY1, int aX2, int aY2, int aSpeed, bool aMoveRelative)
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
		MouseMove(aX1, aY1, aSpeed, aMoveRelative);

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
	// but allow the user to specify a higher time if desired.  Note that for Win9x,
	// a true Sleep() is done because it is much more accurate than the MsgSleep() method,
	// at least on Win98SE when short sleeps are done:
	#define MOUSE_SLEEP \
		if (g.MouseDelay >= 0)\
		{\
			if (g.MouseDelay < 25 && g_os.IsWin9x())\
				Sleep(g.MouseDelay);\
			else\
				SLEEP_WITHOUT_INTERRUPTION(g.MouseDelay)\
		}

	// Do the drag operation
	switch (aVK)
	{
	case VK_LBUTTON:
		MouseEvent(MOUSEEVENTF_LEFTDOWN);
		MOUSE_SLEEP;
		MouseMove(aX2, aY2, aSpeed, aMoveRelative);
		MOUSE_SLEEP;
		MouseEvent(MOUSEEVENTF_LEFTUP);
		break;
	case VK_RBUTTON:
		MouseEvent(MOUSEEVENTF_RIGHTDOWN);
		MOUSE_SLEEP;
		MouseMove(aX2, aY2, aSpeed, aMoveRelative);
		MOUSE_SLEEP;
		MouseEvent(MOUSEEVENTF_RIGHTUP);
		break;
	case VK_MBUTTON:
		MouseEvent(MOUSEEVENTF_MIDDLEDOWN);
		MOUSE_SLEEP;
		MouseMove(aX2, aY2, aSpeed, aMoveRelative);
		MOUSE_SLEEP;
		MouseEvent(MOUSEEVENTF_MIDDLEUP);
		break;
	case VK_XBUTTON1:
		MouseEvent(MOUSEEVENTF_XDOWN, 0, 0, XBUTTON1);
		MOUSE_SLEEP;
		MouseMove(aX2, aY2, aSpeed, aMoveRelative);
		MOUSE_SLEEP;
		MouseEvent(MOUSEEVENTF_XUP, 0, 0, XBUTTON1);
		break;
	case VK_XBUTTON2:
		MouseEvent(MOUSEEVENTF_XDOWN, 0, 0, XBUTTON2);
		MOUSE_SLEEP;
		MouseMove(aX2, aY2, aSpeed, aMoveRelative);
		MOUSE_SLEEP;
		MouseEvent(MOUSEEVENTF_XUP, 0, 0, XBUTTON2);
		break;
	}
	// It seems best to always do this one too in case the script line that caused
	// us to be called here is followed immediately by another script line which
	// is either another mouse click or something that relies upon this mouse drag
	// having been completed:
	MOUSE_SLEEP;
	return OK;
}



ResultType Line::MouseClick(vk_type aVK, int aX, int aY, int aClickCount, int aSpeed, char aEventType
	, bool aMoveRelative)
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
		// that may sometimes (by intent) resolve to zero or negative:
		return OK;

	// The chars 'U' (up) and 'D' (down), if specified, will restrict the clicks
	// to being only DOWN or UP (so that the mouse button can be held down, for
	// example):
	if (aEventType)
		aEventType = toupper(aEventType);

	// Do we need to move the mouse?
	if (aX != COORD_UNSPECIFIED && aY != COORD_UNSPECIFIED) // Otherwise don't bother.
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
		MOUSE_SLEEP;
		return OK;
	}
	else if (aVK == VK_WHEEL_DOWN)
	{
		MouseEvent(MOUSEEVENTF_WHEEL, 0, 0, -(aClickCount * WHEEL_DELTA));
		MOUSE_SLEEP;
		return OK;
	}

	// Otherwise:
	for (int i = 0; i < aClickCount; ++i)
	{
		// Note: It seems best to always Sleep a certain minimum time between events
		// because the click-down event may cause the target app to do something which
		// changes the context or nature of the click-up event.  AutoIt3 has also been
		// revised to do this.
		switch (aVK)
		{
		// The below calls to MouseEvent() do not specify coordinates because such are only
		// needed if we were to include MOUSEEVENTF_MOVE in the dwFlags parameter, which
		// we don't since we've already moved the mouse (above) if that was needed:
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
		case VK_XBUTTON1:
			if (aEventType != 'U')
			{
				MouseEvent(MOUSEEVENTF_XDOWN, 0, 0, XBUTTON1);
				MOUSE_SLEEP;
			}
			if (aEventType != 'D')
			{
				MouseEvent(MOUSEEVENTF_XUP, 0, 0, XBUTTON1);
				MOUSE_SLEEP;
			}
			break;
		case VK_XBUTTON2:
			if (aEventType != 'U')
			{
				MouseEvent(MOUSEEVENTF_XDOWN, 0, 0, XBUTTON2);
				MOUSE_SLEEP;
			}
			if (aEventType != 'D')
			{
				MouseEvent(MOUSEEVENTF_XUP, 0, 0, XBUTTON2);
				MOUSE_SLEEP;
			}
			break;
		} // switch()
	} // for()

	return OK;
}



void Line::MouseMove(int aX, int aY, int aSpeed, bool aMoveRelative)
// Note: This is based on code in the AutoIt3 source.
{
	POINT ptCur;
	int xCur, yCur;
	int delta;
	const int	nMinSpeed = 32;
	RECT rect;

	if (aSpeed < 0) // This can happen during script's runtime due to MouseClick's speed being a var containing a neg.
		aSpeed = 0;  // 0 is the fastest.
	else
		if (aSpeed > MAX_MOUSE_SPEED)
			aSpeed = MAX_MOUSE_SPEED;

	if (aMoveRelative)  // We're moving the mouse cursor relative to its current position.
	{
		GetCursorPos(&ptCur);
		aX += ptCur.x;
		aY += ptCur.y;
	}
	else
	{
		if (!(g.CoordMode & COORD_MODE_MOUSE) && GetWindowRect(GetForegroundWindow(), &rect))
		{
			// Relative vs. screen coords.
			aX += rect.left;
			aY += rect.top;
		}
	}

	// AutoIt3: Get size of desktop
	if (!GetWindowRect(GetDesktopWindow(), &rect)) // Might fail if there is no desktop (e.g. user not logged in).
		rect.bottom = rect.left = rect.right = rect.top = 0;  // Arbitrary defaults.

	// AutoIt3: Convert our coords to MOUSEEVENTF_ABSOLUTE coords
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
	point_and_hwnd_type pah = {0};
	pah.pt = point;
	EnumChildWindows(parent_under_cursor, EnumChildFindPoint, (LPARAM)&pah); // Find topmost control containing point.
	if (pah.hwnd_found)
		child_under_cursor = pah.hwnd_found;

	if (parent_under_cursor == child_under_cursor) // if there's no control per se, make it blank.
		return output_var_child->Assign();

	class_and_hwnd_type cah;
	cah.hwnd = child_under_cursor;  // This is the specific control we need to find the sequence number of.
	char class_name[1024];
	cah.class_name = class_name;
	if (!GetClassName(cah.hwnd, class_name, sizeof(class_name) - 3))  // -3 to allow room for sequence number.
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
{
	#define PAH ((point_and_hwnd_type *)lParam)
	if (!IsWindowVisible(aWnd)) // Omit hidden controls, like Window Spy does.
		return TRUE;
	RECT rect;
	if (!GetWindowRect(aWnd, &rect))
		return TRUE;
	// The given point must be inside aWnd's bounds.  Then, if there is no hwnd found yet or if aWnd
	// is entirely contained within the previously found hwnd, update to a "better" found window like
	// Window Spy.  This overcomes the limitations of WindowFromPoint() and ChildWindowFromPoint():
	if (PAH->pt.x >= rect.left && PAH->pt.x <= rect.right && PAH->pt.y >= rect.top && PAH->pt.y <= rect.bottom)
	{
		// If the window's center is closer to the given point, break the tie and have it take
		// precedence.  This solves the problem where a particular control from a set of overlapping
		// controls is chosen arbitrarily (based on Z-order) rather than based on something the
		// user would find more intuitive (the control whose center is closest to the mouse):
		double center_x = rect.left + (double)(rect.right - rect.left) / 2;
		double center_y = rect.top + (double)(rect.bottom - rect.top) / 2;
		// Taking the absolute value first is not necessary because it seems that qmathHypot()
		// takes the square root of the sum of the squares, which handles negatives correctly:
		double distance = qmathHypot(PAH->pt.x - center_x, PAH->pt.y - center_y);
		//double distance = qmathSqrt(qmathPow(PAH->pt.x - center_x, 2) + qmathPow(PAH->pt.y - center_y, 2));
		bool update_it = !PAH->hwnd_found;
		if (!update_it)
		{
			// If the new window's rect is entirely contained within the old found-window's rect, update
			// even if the distance is greater.  Conversely, if the new window's rect entirely encloses
			// the old window's rect, do not update even if the distance is less:
			if (rect.left >= PAH->rect_found.left && rect.right <= PAH->rect_found.right
				&& rect.top >= PAH->rect_found.top && rect.bottom <= PAH->rect_found.bottom)
				update_it = true; // New is entirely enclosed by old: update to the New.
			else if (   distance < PAH->distance &&
				(PAH->rect_found.left < rect.left || PAH->rect_found.right > rect.right
					|| PAH->rect_found.top < rect.top || PAH->rect_found.bottom > rect.bottom)   )
				update_it = true; // New doesn't entirely enclose old and new's center is closer to the point.
		}
		if (update_it)
		{
			PAH->hwnd_found = aWnd;
			PAH->rect_found = rect;
			PAH->distance = distance;
		}
	}
	return TRUE; // Continue enumeration all the way through.
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
	{
		space_needed = GetExpandedArgSize(false); // There's at most one arg to expand in this case.
		if (space_needed == VARSIZE_ERROR)
			return FAIL;  // It will have already displayed the error.
	}

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
	char *one_beyond_contents_end = ExpandArg(output_var->Contents(), 1);
	if (!one_beyond_contents_end)
		return FAIL;  // ExpandArg() will have already displayed the error.
	// Set the length explicitly rather than using space_needed because GetExpandedArgSize()
	// sometimes returns a larger size than is actually needed (e.g. for ScriptGetCursor()):
	output_var->Length() = (VarSizeType)(one_beyond_contents_end - output_var->Contents() - 1);
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

	// If we are running NT/2k/XP, make sure we have rights to shutdown
	if (g_os.IsWinNT()) // NT/2k/XP/2003 and family
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



ResultType Line::StringSplit(char *aArrayName, char *aInputString, char *aDelimiterList, char *aOmitList)
{
	// Make it longer than Max so that FindOrAddVar() will be able to spot and report var names
	// that are too long:
	char var_name[MAX_VAR_NAME_LENGTH + 20];
	snprintf(var_name, sizeof(var_name), "%s0", aArrayName);
	Var *array0 = g_script.FindOrAddVar(var_name);
	if (!array0)
		return FAIL;  // It will have already displayed the error.

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
			if (   !(next_element = g_script.FindOrAddVar(var_name, VAR_NAME_LENGTH_DEFAULT, array0))   )
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
		if (   !(next_element = g_script.FindOrAddVar(var_name, VAR_NAME_LENGTH_DEFAULT, array0))   )
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
	char *name_delimiter = strrchr(aFileSpec, '\\');
	if (!name_delimiter)
		if (   !(name_delimiter = strrchr(aFileSpec, ':'))   )
			name_delimiter = NULL;

	char *name = name_delimiter ? name_delimiter + 1 : aFileSpec; // If no delimiter, name is the entire string.
	char *ext_dot = strrchr(name, '.');

	if (output_var_name && !output_var_name->Assign(name))
		return FAIL;

	if (output_var_dir)
	{
		if (!name_delimiter)
			output_var_dir->Assign(); // Shouldn't fail.
		else if (*name_delimiter == '\\')
		{
			if (!output_var_dir->Assign(aFileSpec, (VarSizeType)(name_delimiter - aFileSpec)))
				return FAIL;
		}
		else // *name_delimiter == ':', e.g. "C:Some File.txt".  If aFileSpec starts with just ":",
			 // the dir returned here will also start with just ":" since that's rare & illegal anyway.
			if (!output_var_dir->Assign(aFileSpec, (VarSizeType)(name_delimiter - aFileSpec + 1)))
				return FAIL;
	}

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

	if (output_var_drive)
	{
		char *drive = omit_leading_whitespace(aFileSpec);
		if (!*drive)
			output_var_drive->Assign();
		else
		{
			// UNCs are detected with this approach so that double sets of backslashes -- which sometimes
			// occur by accident in "built filespecs" and are tolerated by the OS -- are not falsely
			// detected as UNCs.
			if (*drive == '\\' && *(drive + 1) == '\\')
			{
				char *drive_end = strchr(drive + 2, '\\');
				if (drive_end)
					output_var_drive->Assign(drive, (VarSizeType)(drive_end - drive));
				else
					output_var_drive->Assign(drive); // Assume the entire string is the server name.
			}
			else if (*(drive + 1) == ':') // It's an absolute path.
				// Assign letter and colon for consistency with server naming convention above.
				// i.e. so that server name and drive can be used without having to worry about
				// whether it needs a colon added or not.
				output_var_drive->Assign(drive, 2);
			else // It's debatable, but it seems best to return a blank drive if a aFileSpec is a relative path.
				 // rather than trying to use GetFullPathName() on a potentially non-existent file/dir.
				 // _splitpath() doesn't fetch the drive letter of relative paths either.
				output_var_drive->Assign();
		}
	}

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
		// Resolve since parts of the ATOI() macro are inline:
		__int64 item1_int = ATOI64(sort_item1);
		__int64 item2_int = ATOI64(sort_item2);
		return (int)(g_SortReverse ? item2_int - item1_int : item1_int - item2_int);
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
	char *cp;

	for (cp = aOptions; *cp; ++cp)
	{
		switch(toupper(*cp))
		{
		case 'C':
			g_SortCaseSensitive = true;
			break;
		case 'D':
			if (!*(cp + 1)) // Avoids out-of-bounds when the loop's own ++cp is done.
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
			g_SortReverse = true;
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
	char **item = (char **)malloc((item_count + 1) * sizeof(char *));
	if (!item)
		return LineError("Out of mem");  // Short msg. since so rare.

	// Scan aContents and do the following:
	// 1) Replace each delimiter with a terminator so that the individual items can be seen
	//    as real strings by the SortWithOptions() and when copying the sorted results back
	//    into output_vav.  It is safe change aContents in this way because
	//    ArgMustBeDereferenced() has ensured that those contents are in the deref buffer.
	// 2) Store a marker/pointer to each item (string) in aContents so that we know where
	//    each item begins for sorting and recopying purposes.
	for (item_count = 0, cp = item[0] = aContents; *cp; ++cp)
	{
		if (*cp == delimiter)  // Each delimiter char becomes the terminator of the previous key phrase.
		{
			*cp = '\0';  // Terminate the item that appears before this delimiter.
			item[++item_count] = cp + 1; // Make a pointer to the next item's place in aContents.
		}
	}
	// Add the last item to the count only if it wasn't disqualified earlier:
	if (!terminate_last_item_with_delimiter)
		++item_count;
	char *original_last_item = item[item_count - 1];  // The location of the last item before it gets sorted.

	// Now aContents has been divided up based on delimiter.  Sort the array of pointers
	// so that they indicate the correct ordering to copy aContents into output_var:
	qsort((void *)item, item_count, sizeof(item[0]), sort_by_naked_filename ? SortByNakedFilename : SortWithOptions);

	// Copy the sorted pointers back into output_var, which might not already be sized correctly
	// if it's the clipboard or it was an environment variable when it came in as the input.
	// If output_var is the clipboard, this call will set up the clipboard for writing:
	if (output_var->Assign(NULL, (VarSizeType)aContents_length) != OK) // Might fail due to clipboard problem.
		return FAIL;

	char *source, *dest;
	char *pos_of_original_last_item_in_dest = NULL;  // Set default in case original last item is still the last item.
	size_t i, item_count_less_1 = item_count - 1;

	// Copy the sorted result back into output_var.  Do all except the last item, since the last
	// item gets special treatment depending on the options that were specified.  The call to
	// output_var->Contents() below should never fail due to the above having prepped it:
	for (dest = output_var->Contents(), i = 0; i < item_count_less_1; ++i)
	{
		if (item[i] == original_last_item)
			pos_of_original_last_item_in_dest = dest;
		for (source = item[i]; *source;)
			*dest++ = *source++;
		*dest++ = delimiter;  // Put each item's delimiter back in so that format is the same as the original.
	}

	// Copy the last item:
	for (source = item[item_count_less_1]; *source;)
		*dest++ = *source++;
	// If the original list's last item had a terminating delimiter and the specified options said
	// to treat it not as a delimiter but as a final char of sorts, include it after the item that
	// is now last so that the overall layout is the same:
	if (terminate_last_item_with_delimiter)
		*dest++ = delimiter;
	*dest = '\0';  // Terminate the variable's contents.

	// Check if special handling is needed due to the following situation:
	// Delimiter is LF but the contents are lines delimited by CRLF, not just LF
	// and the original/unsorted list's last item was not terminated by an
	// "allowed delimiter".  The symptoms of this are that after the sort, the
	// last item will end in \r when it should end in no delimiter at all.
	// This happens pretty often, such as when the clipboard contains files.
	// In the future, an option letter can be added to turn off this workaround,
	// but it seems unlikely that anyone would ever want to:
	if (delimiter == '\n' && !terminate_last_item_with_delimiter && *(dest - 1) == '\r'
		&& pos_of_original_last_item_in_dest)
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

	// Free the memory used for the sort:
	free(item);

	// It is not necessary to set output_var->Length() here because its length hasn't changed
	// since it was originally set by the above call "output_var->Assign(NULL..."
	return output_var->Close();  // Close in case it's the clipboard.
}



ResultType Line::ScriptGetKeyState(char *aKeyName, char *aOption)
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
	}
	else // There is a virtual key (not a joystick control).
	{
		switch (toupper(*aOption))
		{
		case 'T': // Whether a toggleable key such as CapsLock is currently turned on.
			// Under Win9x, at least certain versions and for certain hardware, this
			// doesn't seem to be always accurate, especially when the key has just
			// been toggled and the user hasn't pressed any other key since then.
			// I tried using GetKeyboardState() instead, but it produces the same
			// result.  Therefore, I've documented this as a limitation in the help file.
			// In addition, this was attempted but it didn't seem to help:
			//if (g_os.IsWin9x())
			//{
			//	DWORD my_thread  = GetCurrentThreadId();
			//	DWORD fore_thread = GetWindowThreadProcessId(GetForegroundWindow(), NULL);
			//	bool is_attached_my_to_fore = false;
			//	if (fore_thread && fore_thread != my_thread)
			//		is_attached_my_to_fore = AttachThreadInput(my_thread, fore_thread, TRUE) != 0;
			//	output_var->Assign(IsKeyToggledOn(vk) ? "D" : "U");
			//	if (is_attached_my_to_fore)
			//		AttachThreadInput(my_thread, fore_thread, FALSE);
			//	return OK;
			//}
			//else
			return output_var->Assign(IsKeyToggledOn(vk) ? "D" : "U");
		case 'P': // Physical state of key.
			if (VK_IS_MOUSE(vk)) // mouse button
			{
				if (g_MouseHook) // mouse hook is installed, so use it's tracking of physical state.
					return output_var->Assign(g_PhysicalKeyState[vk] & STATE_DOWN ? "D" : "U");
				else // Even for Win9x/NT, it seems slightly better to call this rather than IsKeyDown9xNT():
					return output_var->Assign(IsPhysicallyDown(vk) ? "D" : "U");
			}
			else // keyboard
			{
				if (g_KeybdHook)
					// Since the hook is installed, use its value rather than that from
					// GetAsyncKeyState(), which doesn't seem to return the physical state
					// as expected/advertised, least under WinXP:
					return output_var->Assign(g_PhysicalKeyState[vk] & STATE_DOWN ? "D" : "U");
				else // Even for Win9x/NT, it seems slightly better to call this rather than IsKeyDown9xNT():
					return output_var->Assign(IsPhysicallyDown(vk) ? "D" : "U");
			}
		default: // Logical state of key.
			if (g_os.IsWin9x() || g_os.IsWinNT4())
				return output_var->Assign(IsKeyDown9xNT(vk) ? "D" : "U"); // This seems more likely to be reliable.
			else
				// On XP/2K at least, a key can be physically down even if it isn't logically down,
				// which is why the below specifically calls IsKeyDown2kXP() rather than some more
				// comprehensive method such as consulting the physical key state as tracked by the hook:
				return output_var->Assign(IsKeyDown2kXP(vk) ? "D" : "U");
		}
	}

	// Since the above didn't return, joy contains a valid joystick button/control ID:
	bool joy_is_button = (joy >= JOYCTRL_1 && joy <= JOYCTRL_BUTTON_MAX);

	JOYCAPS jc;
	if (!joy_is_button && joy != JOYCTRL_POV)
	{
		// Get the joystick's range of motion so that we can report it as a percentage.
		if (joyGetDevCaps(joystick_id, &jc, sizeof(JOYCAPS)) != JOYERR_NOERROR)
			ZeroMemory(&jc, sizeof(jc));
	}

	JOYINFOEX jie;

	if (joy != JOYCTRL_NAME && joy != JOYCTRL_BUTTONS && joy != JOYCTRL_AXES && joy != JOYCTRL_INFO)
	{
		jie.dwSize = sizeof(JOYINFOEX);
		jie.dwFlags = JOY_RETURNALL;
		if (joyGetPosEx(joystick_id, &jie) != JOYERR_NOERROR)
			return output_var->Assign("");
		if (joy_is_button)
			return output_var->Assign(((jie.dwButtons >> (joy - JOYCTRL_1)) & (DWORD)0x01) ? "D" : "U");
	}

	// Otherwise:
	UINT range;
	char buf[128], *buf_ptr;

	switch(joy)
	{
	case JOYCTRL_XPOS:
		range = (jc.wXmax > jc.wXmin) ? jc.wXmax - jc.wXmin : 0;
		return output_var->Assign(range ? 100 * (double)jie.dwXpos / range : jie.dwXpos);
	case JOYCTRL_YPOS:
		range = (jc.wYmax > jc.wYmin) ? jc.wYmax - jc.wYmin : 0;
		return output_var->Assign(range ? 100 * (double)jie.dwYpos / range : jie.dwYpos);
	case JOYCTRL_ZPOS:
		range = (jc.wZmax > jc.wZmin) ? jc.wZmax - jc.wZmin : 0;
		return output_var->Assign(range ? 100 * (double)jie.dwZpos / range : jie.dwZpos);
	case JOYCTRL_RPOS:  // Rudder or 4th axis.
		range = (jc.wRmax > jc.wRmin) ? jc.wRmax - jc.wRmin : 0;
		return output_var->Assign(range ? 100 * (double)jie.dwRpos / range : jie.dwRpos);
	case JOYCTRL_UPOS:  // 5th axis.
		range = (jc.wUmax > jc.wUmin) ? jc.wUmax - jc.wUmin : 0;
		return output_var->Assign(range ? 100 * (double)jie.dwUpos / range : jie.dwUpos);
	case JOYCTRL_VPOS:  // 6th axis.
		range = (jc.wVmax > jc.wVmin) ? jc.wVmax - jc.wVmin : 0;
		return output_var->Assign(range ? 100 * (double)jie.dwVpos / range : jie.dwVpos);
	case JOYCTRL_POV:  // Need to explicitly compare against JOY_POVCENTERED because it's a WORD not a DWORD.
		if (jie.dwPOV == JOY_POVCENTERED)
			return output_var->Assign("-1"); // Documented behavior.
		else
			return output_var->Assign(jie.dwPOV);
	case JOYCTRL_NAME:
		return output_var->Assign(jc.szPname);
	case JOYCTRL_BUTTONS:
		return output_var->Assign((DWORD)jc.wNumButtons);  // wMaxButtons is the *driver's* max supported buttons.
	case JOYCTRL_AXES:
		return output_var->Assign((DWORD)jc.wNumAxes);  // wMaxAxes is the *driver's* max supported axes.
	case JOYCTRL_INFO:
		buf_ptr = buf;
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
		return output_var->Assign(buf);
	}

	return output_var->Assign(); // Should never be executed.
}



ResultType Line::DriveSpace(char *aPath, bool aGetFreeSpace)
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
	// The program won't launch at all on Win95a (original Win95) unless the function address is resolved
	// at runtime:
	typedef BOOL (WINAPI *GetDiskFreeSpaceExType)(LPCTSTR, PULARGE_INTEGER, PULARGE_INTEGER, PULARGE_INTEGER);
	static GetDiskFreeSpaceExType MyGetDiskFreeSpaceEx =
		(GetDiskFreeSpaceExType)GetProcAddress(GetModuleHandle("kernel32.dll"), "GetDiskFreeSpaceExA");

	// MSDN: "The GetDiskFreeSpaceEx function returns correct values for all volumes, including those
	// that are greater than 2 gigabytes."
	if (MyGetDiskFreeSpaceEx)  // Function is available (unpatched Win95 and WinNT might not have it).
	{
		ULARGE_INTEGER uiTotal, uiFree, uiUsed;
		if (!MyGetDiskFreeSpaceEx(buf, &uiFree, &uiTotal, &uiUsed))
			return OK; // Let ErrorLevel tell the story.
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		// Casting this way limits us to 2,097,152 gigabytes in size:
		return output_var->Assign(   (__int64)((unsigned __int64)(aGetFreeSpace ? uiFree.QuadPart : uiTotal.QuadPart)
			/ (1024*1024))   );
	}
	else
	{
		DWORD dwSectPerClust, dwBytesPerSect, dwFreeClusters, dwTotalClusters;
		if (!GetDiskFreeSpace(buf, &dwSectPerClust, &dwBytesPerSect, &dwFreeClusters, &dwTotalClusters))
			return OK; // Let ErrorLevel tell the story.
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		return output_var->Assign(   (__int64)((unsigned __int64)((aGetFreeSpace ? dwFreeClusters : dwTotalClusters)
			* dwSectPerClust * dwBytesPerSect) / (1024*1024))   );
	}
#endif	

	return OK;
}



ResultType Line::DriveGet(char *aCmd, char *aValue)
// This function has been adapted from the AutoIt3 source.
{
	DriveCmds drive_cmd = ConvertDriveCmd(aCmd);
	if (drive_cmd == DRIVE_CMD_CAPACITY)
		return DriveSpace(aValue, false);

	char path[MAX_PATH + 1];  // +1 to allow room for trailing backslash in case it needs to be added.
	size_t path_length;

	// Notes about the below macro:
	// Leave space for the backslash in case its needed:
	// au3: attempt to fix the parameter passed (backslash may be needed in some OSes).
	#define DRIVE_SET_PATH \
		strlcpy(path, aValue, sizeof(path) - 1);\
		path_length = strlen(path);\
		if (path_length && path[path_length - 1] != '\\')\
			path[path_length++] = '\\';

	if (drive_cmd == DRIVE_CMD_SETLABEL)
	{
		DRIVE_SET_PATH
		SetErrorMode(SEM_FAILCRITICALERRORS);		// So a:\ does not ask for disk
		char *new_label = omit_leading_whitespace(aCmd + 9);  // Example: SetLabel:MyLabel
		return g_ErrorLevel->Assign(SetVolumeLabel(path, new_label) ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);
	}

	g_ErrorLevel->Assign(ERRORLEVEL_ERROR);  // Set default since there are many points of return.

	Var *output_var = ResolveVarOfArg(0);
	if (!output_var)
		return FAIL;

	switch(drive_cmd)
	{

	case DRIVE_CMD_INVALID:
		// Since command names are validated at load-time, this only happens if the command name
		// was contained in a variable reference.  Since that is very rare, just set ErrorLevel
		// and return:
		return output_var->Assign();  // Let ErrorLevel tell the story.

	case DRIVE_CMD_LIST:
	{
		UINT uiFlag, uiTemp;

		if (!*aValue) uiFlag = 99;
		else if (!stricmp(aValue, "CDRom")) uiFlag = DRIVE_CDROM;
		else if (!stricmp(aValue, "Removable")) uiFlag = DRIVE_REMOVABLE;
		else if (!stricmp(aValue, "Fixed")) uiFlag = DRIVE_FIXED;
		else if (!stricmp(aValue, "Network")) uiFlag = DRIVE_REMOTE;
		else if (!stricmp(aValue, "Ramdisk")) uiFlag = DRIVE_RAMDISK;
		else if (!stricmp(aValue, "Unknown")) uiFlag = DRIVE_UNKNOWN;
		else // Let ErrorLevel tell the story.
			return OK;

		char found_drives[32];  // Need room for all 26 possible drive letters.
		int found_drives_count;
		UCHAR letter;
		char buf[128], *buf_ptr;

		SetErrorMode(SEM_FAILCRITICALERRORS);		// So a:\ does not ask for disk

		for (found_drives_count = 0, letter = 'A'; letter <= 'Z'; ++letter)
		{
			buf_ptr = buf;
			*buf_ptr++ = letter;
			*buf_ptr++ = ':';
			*buf_ptr++ = '\\';
			*buf_ptr = '\0';
			uiTemp = GetDriveType(buf);
			if (uiTemp == uiFlag || (uiFlag == 99 && uiTemp != DRIVE_NO_ROOT_DIR))
				found_drives[found_drives_count++] = letter;  // Store just the drive letters.
		}
		found_drives[found_drives_count] = '\0';  // Terminate the string of found drive letters.
		output_var->Assign(found_drives);
		if (!*found_drives)
			return OK;  // Seems best to flag zero drives in the system as default ErrorLevel of "1".
		break;
	}

	case DRIVE_CMD_FILESYSTEM:
	case DRIVE_CMD_LABEL:
	case DRIVE_CMD_SERIAL:
	{
		char label[256];
		char file_system[256];
		DRIVE_SET_PATH
		SetErrorMode(SEM_FAILCRITICALERRORS);		// So a:\ does not ask for disk
		DWORD dwVolumeSerial, dwMaxCL, dwFSFlags;
		if (!GetVolumeInformation(path, label, sizeof(label) - 1, &dwVolumeSerial, &dwMaxCL
			, &dwFSFlags, file_system, sizeof(file_system) - 1))
			return output_var->Assign(); // Let ErrorLevel tell the story.
		switch(drive_cmd)
		{
		case DRIVE_CMD_FILESYSTEM: output_var->Assign(file_system); break;
		case DRIVE_CMD_LABEL: output_var->Assign(label); break;
		case DRIVE_CMD_SERIAL: output_var->Assign(dwVolumeSerial); break;
		}
		break;
	}

	case DRIVE_CMD_TYPE:
	{
		DRIVE_SET_PATH
		SetErrorMode(SEM_FAILCRITICALERRORS);	// So a:\ does not ask for disk
		UINT uiType = GetDriveType(path);
		switch (uiType)
		{
		case DRIVE_UNKNOWN:
			output_var->Assign("Unknown");
			break;
		case DRIVE_REMOVABLE:
			output_var->Assign("Removable");
			break;
		case DRIVE_FIXED:
			output_var->Assign("Fixed");
			break;
		case DRIVE_REMOTE:
			output_var->Assign("Network");
			break;
		case DRIVE_CDROM:
			output_var->Assign("CDROM");
			break;
		case DRIVE_RAMDISK:
			output_var->Assign("RAMDisk");
			break;
		default: // DRIVE_NO_ROOT_DIR
			return output_var->Assign();  // Let ErrorLevel tell the story.
		}
		break;
	}

	case DRIVE_CMD_STATUS:
	{
		DRIVE_SET_PATH
		DWORD dwSectPerClust, dwBytesPerSect, dwFreeClusters, dwTotalClusters;
		DWORD last_error = ERROR_SUCCESS;
		SetErrorMode(SEM_FAILCRITICALERRORS);		// So a:\ does not ask for disk
		if (   !(GetDiskFreeSpace(path, &dwSectPerClust, &dwBytesPerSect, &dwFreeClusters, &dwTotalClusters))   )
			last_error = GetLastError();
		switch (last_error)
		{
		case ERROR_SUCCESS:
			output_var->Assign("Ready");
			break;
		case ERROR_PATH_NOT_FOUND:
			output_var->Assign("Invalid");
			break;
		case ERROR_NOT_READY:
			output_var->Assign("NotReady");
			break;
		case ERROR_WRITE_PROTECT:
			output_var->Assign("ReadOnly");
			break;
		default:
			output_var->Assign("Unknown");
		}
		break;
	}
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
		for (int d = 0, found_instance = 0; d < dest_count && !found; ++d)
		{
			ml.dwDestination = d;
			if (mixerGetLineInfo((HMIXEROBJ)hMixer, &ml, MIXER_GETLINEINFOF_DESTINATION) != MMSYSERR_NOERROR)
				// Keep trying in case the others can be retrieved.
				continue;
			source_count = ml.cConnections;  // Make a copy of this value so that the struct can be reused.
			for (int s = 0; s < source_count && !found; ++s)
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
    MIXERCONTROL mc = {0};
    mc.cbStruct = sizeof(mc);
    MIXERLINECONTROLS mlc = {0};
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

	// Does user want to adjust the current setting by a certain amount?:
	bool adjust_current_setting = aSetting && (*aSetting == '-' || *aSetting == '+');

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

	if (SOUND_MODE_IS_SET)
	{
		switch (aControlType)
		{
		case MIXERCONTROL_CONTROLTYPE_ONOFF:
		case MIXERCONTROL_CONTROLTYPE_MUTE:
		case MIXERCONTROL_CONTROLTYPE_MONO:
		case MIXERCONTROL_CONTROLTYPE_LOUDNESS:
		case MIXERCONTROL_CONTROLTYPE_STEREOENH:
		case MIXERCONTROL_CONTROLTYPE_BASS_BOOST:
			if (adjust_current_setting) // The user wants this toggleable control to be toggled to its opposite state:
				mcdMeter.dwValue = (mcdMeter.dwValue > mc.Bounds.dwMinimum) ? mc.Bounds.dwMinimum : mc.Bounds.dwMaximum;
			else // Set the value according to whether the user gave us a setting that is greater than zero:
				mcdMeter.dwValue = (setting_percent > 0.0) ? mc.Bounds.dwMaximum : mc.Bounds.dwMinimum;
			break;
		default: // For all others, assume the control can have more than just ON/OFF as its allowed states.
		{
			// Make this an __int64 vs. DWORD to avoid underflow (so that a setting_percent of -100
			// is supported whenenver the difference between Min and Max is large, such as MAXDWORD):
			__int64 specified_vol = (__int64)((mc.Bounds.dwMaximum - mc.Bounds.dwMinimum) * (setting_percent / 100.0));
			if (adjust_current_setting)
			{
				// Make it a big int so that overflow/underflow can be detected:
				__int64 vol_new = mcdMeter.dwValue + specified_vol;
				if (vol_new < mc.Bounds.dwMinimum) vol_new = mc.Bounds.dwMinimum;
				else if (vol_new > mc.Bounds.dwMaximum) vol_new = mc.Bounds.dwMaximum;
				mcdMeter.dwValue = (DWORD)vol_new;
			}
			else
				mcdMeter.dwValue = (DWORD)specified_vol; // Due to the above, it's known to be positive in this case.
		}
		} // switch()

		MMRESULT result = mixerSetControlDetails((HMIXEROBJ)hMixer, &mcd, MIXER_GETCONTROLDETAILSF_VALUE);
		mixerClose(hMixer);
		return g_ErrorLevel->Assign(result == MMSYSERR_NOERROR ? ERRORLEVEL_NONE : "Can't Change Setting");
	}

	// Otherwise, the mode is "Get":
	mixerClose(hMixer);
	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.

	switch (aControlType)
	{
	case MIXERCONTROL_CONTROLTYPE_ONOFF:
	case MIXERCONTROL_CONTROLTYPE_MUTE:
	case MIXERCONTROL_CONTROLTYPE_MONO:
	case MIXERCONTROL_CONTROLTYPE_LOUDNESS:
	case MIXERCONTROL_CONTROLTYPE_STEREOENH:
	case MIXERCONTROL_CONTROLTYPE_BASS_BOOST:
		return output_var->Assign(mcdMeter.dwValue ? "On" : "Off");
	default: // For all others, assume the control can have more than just ON/OFF as its allowed states.
		// The MSDN docs imply that values fetched via the above method do not distinguish between
		// left and right volume levels, unlike waveOutGetVolume():
		return output_var->Assign(   ((double)100 * (mcdMeter.dwValue - (DWORD)mc.Bounds.dwMinimum))
			/ (mc.Bounds.dwMaximum - mc.Bounds.dwMinimum)   );
	}
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

	if (*aVolume == '-' || *aVolume == '+') // User wants to adjust the current level by a certain amount.
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
		if (vol_left < 0) vol_left = 0;
		else if (vol_left > 0xFFFF) vol_left = 0xFFFF;
		if (vol_right < 0) vol_right = 0;
		else if (vol_right > 0xFFFF) vol_right = 0xFFFF;
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
	// Adapted from the AutoIt3 source.
	// See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/multimed/htm/_win32_play.asp
	// for some documentation mciSendString() and related.
	char buf[MAX_PATH * 2];
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



ResultType Line::URLDownloadToFile(char *aURL, char *aFilespec)
// This has been adapted from the AutoIt3 source.
{
typedef HINTERNET (WINAPI *MyInternetOpen)(LPCTSTR, DWORD, LPCTSTR, LPCTSTR, DWORD dwFlags);
typedef HINTERNET (WINAPI *MyInternetOpenUrl)(HINTERNET hInternet, LPCTSTR, LPCTSTR, DWORD, DWORD, LPDWORD);
typedef BOOL (WINAPI *MyInternetCloseHandle)(HINTERNET);
typedef BOOL (WINAPI *MyInternetReadFile)(HINTERNET, LPVOID, DWORD, LPDWORD);

#ifndef INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY
	#define INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY 4
#endif

	HINTERNET				hInet;
	HINTERNET				hFile;
	BYTE					bufData[8192];
	DWORD					dwBytesRead;
	FILE					*fptr;
	MyInternetOpen			lpfnInternetOpen;
	MyInternetOpenUrl		lpfnInternetOpenUrl;
	MyInternetCloseHandle	lpfnInternetCloseHandle;
	MyInternetReadFile		lpfnInternetReadFile;
	HINSTANCE				hinstLib;

	// Check that we have IE3 and access to wininet.dll
	hinstLib = LoadLibrary("wininet.dll");
	if (hinstLib == NULL)
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);

	// Get the address of all the functions we require
 	lpfnInternetOpen		= (MyInternetOpen)GetProcAddress(hinstLib, "InternetOpenA");
	lpfnInternetOpenUrl		= (MyInternetOpenUrl)GetProcAddress(hinstLib, "InternetOpenUrlA");
	lpfnInternetCloseHandle	= (MyInternetCloseHandle)GetProcAddress(hinstLib, "InternetCloseHandle");
	lpfnInternetReadFile	= (MyInternetReadFile)GetProcAddress(hinstLib, "InternetReadFile");

	// Open the internet session
	hInet = lpfnInternetOpen(NULL, INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY, NULL, NULL, 0);
	if (hInet == NULL)
	{
		FreeLibrary(hinstLib);					// Free the DLL module.
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	}

	// Open the required URL
	hFile = lpfnInternetOpenUrl(hInet, aURL, NULL, 0, 0, 0);
	if (hFile == NULL)
	{
		lpfnInternetCloseHandle(hInet);
		FreeLibrary(hinstLib);					// Free the DLL module.
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	}

	// Open our output file
	fptr = fopen(aFilespec, "wb");	// Open in binary write/destroy mode
	if (fptr == NULL)
	{
		lpfnInternetCloseHandle(hFile);
		lpfnInternetCloseHandle(hInet);
		FreeLibrary(hinstLib);					// Free the DLL module.
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	}

	LONG_OPERATION_INIT_FOR_URL  // Added for AutoHotkey

	// Read the file
	dwBytesRead = 1;
	while (dwBytesRead)
	{
		if ( lpfnInternetReadFile(hFile, (LPVOID)bufData, bytes_to_read, &dwBytesRead) == FALSE )
		{
			lpfnInternetCloseHandle(hFile);
			lpfnInternetCloseHandle(hInet);
			FreeLibrary(hinstLib);				// Free the DLL module.
			fclose(fptr);
			DeleteFile(aFilespec);				// Output is trashed - delete
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
		}

		LONG_OPERATION_UPDATE // Added for AutoHotkey

		if (dwBytesRead)
			fwrite(bufData, dwBytesRead, 1, fptr);
	}

	// Close internet session
	lpfnInternetCloseHandle(hFile);
	lpfnInternetCloseHandle(hInet);
	FreeLibrary(hinstLib);					// Free the DLL module.

	// Close output file
	fclose(fptr);

	return g_ErrorLevel->Assign(ERRORLEVEL_NONE);  // Indicate success.
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

	char working_dir[MAX_PATH * 2];
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
	if (aFilter && *aFilter)
	{
		char *pattern_start = strchr(aFilter, '(');
		if (pattern_start)
		{
			// Make pattern a separate string because we want to remove any spaces from it.
			// For example, if the user specified Documents (*.txt; *.doc), the space after
			// the semicolon should be removed for the pattern string itself but not from
			// the displayed version of the pattern:
			strlcpy(pattern, ++pattern_start, sizeof(pattern));
			char *pattern_end = strrchr(pattern, ')'); // strrchr() in case there are other literal parens.
			if (pattern_end)
				*pattern_end = '\0';  // If parens are empty, this will set pattern to be the empty string.
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
			StrReplace(pattern, " ", "");
			// Also include the All Files (*.*) filter, since there doesn't seem to be much
			// point to making this an option.  This is because the user could always type
			// *.* and press ENTER in the filename field and achieve the same result:
			snprintf(filter, sizeof(filter), "%s%c%s%c" "All Files (*.*)%c*.*%c"
				, aFilter, '\0', pattern, '\0'
				, '\0', '\0'); // double-terminate.
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
	ofn.hwndOwner = NULL; // i.e. no need to have main window forced into the background for this.
	ofn.lpstrTitle = greeting;
	ofn.lpstrFilter = *filter ? filter : "All Files (*.*)\0*.*\0Text Documents (*.txt)\0*.txt\0";
	ofn.lpstrFile = file_buf;
	ofn.nMaxFile = sizeof(file_buf) - 1; // -1 to be extra safe, like AutoIt3.
	// Specifying NULL will make it default to the last used directory (at least in Win2k):
	ofn.lpstrInitialDir = *working_dir ? working_dir : NULL;

	// Note that the OFN_NOCHANGEDIR flag is ineffective in some cases, so we'll use a custom
	// workaround instead.  MSDN: "Windows NT 4.0/2000/XP: This flag is ineffective for GetOpenFileName."
	// In addition, it does not prevent the CWD from changing while the user navigates from folder to
	// folder in the dialog, except perhaps on Win9x.

	bool always_use_save_dialog;
	if (toupper(*aOptions) == 'S')
	{
		always_use_save_dialog = true;
		++aOptions;
	}
	else
		always_use_save_dialog = false;

	int options = ATOI(aOptions);
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

	POST_AHK_DIALOG(0) // Must pass 0 for timeout in this case.

	++g_nFileDialogs;
	// Below: OFN_CREATEPROMPT doesn't seem to work with GetSaveFileName(), so always
	// use GetOpenFileName() in that case:
	BOOL result = (always_use_save_dialog || ((ofn.Flags & OFN_OVERWRITEPROMPT) && !(ofn.Flags & OFN_CREATEPROMPT)))
		? GetSaveFileName(&ofn) : GetOpenFileName(&ofn);
	--g_nFileDialogs;

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



ResultType Line::FileSelectFolder(char *aRootDir, DWORD aOptions, char *aGreeting)
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
	browseInfo.ulFlags = 0x0040 | ((aOptions & FSF_ALLOW_CREATE) ? 0 : 0x200) | ((aOptions & (DWORD)FSF_EDITBOX) ? BIF_EDITBOX : 0);

	char Result[2048];
	browseInfo.pszDisplayName = Result;  // This will hold the user's choice.

	POST_AHK_DIALOG(0) // Must pass 0 for timeout in this case.

	++g_nFolderDialogs;
	LPITEMIDLIST lpItemIDList = SHBrowseForFolder(&browseInfo);  // Spawn Dialog
	--g_nFolderDialogs;

	if (!lpItemIDList)
		return OK;  // Let ErrorLevel tell the story.

	*Result = '\0';  // Reuse this var, this time to old the result of the below:
	SHGetPathFromIDList(lpItemIDList, Result);
	pMalloc->Free(lpItemIDList);
	pMalloc->Release();

	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return output_var->Assign(Result);
}



ResultType Line::FileCreateShortcut(char *aTargetFile, char *aShortcutFile, char *aWorkingDir, char *aArgs
	, char *aDescription, char *aIconFile, char *aHotkey)
// Adapted from the AutoIt3 source.
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	CoInitialize(NULL);
	IShellLink *psl;

	if (SUCCEEDED(CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (LPVOID *)&psl)))
	{
		psl->SetPath(aTargetFile);
		if (aWorkingDir && *aWorkingDir)
			psl->SetWorkingDirectory(aWorkingDir);
		if (aArgs && *aArgs)
			psl->SetArguments(aArgs);
		if (aDescription && *aDescription)
			psl->SetDescription(aDescription);
		if (aIconFile && *aIconFile)
			psl->SetIconLocation(aIconFile, 0);
		if (aHotkey && *aHotkey)
		{
			// If badly formatted, it's not a critical error, just continue.
			// Currently, only shortcuts with a CTRL+ALT are supported.
			// AutoIt3 note: Make sure that CTRL+ALT is selected (otherwise invalid)
			vk_type vk = TextToVK(aHotkey);
			if (vk)
				// Vk in low 8 bits, mods in high 8:
				psl->SetHotkey(   (WORD)vk | ((WORD)(HOTKEYF_CONTROL | HOTKEYF_ALT) << 8)   );
		}

		IPersistFile *ppf;
		WORD wsz[MAX_PATH];
		if(SUCCEEDED(psl->QueryInterface(IID_IPersistFile,(LPVOID *)&ppf)))
		{
			MultiByteToWideChar(CP_ACP, 0, aShortcutFile, -1, (LPWSTR)wsz, MAX_PATH);
			if (SUCCEEDED(ppf->Save((LPCWSTR)wsz, TRUE)))
				g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
			ppf->Release();
		}
		psl->Release();
	}

	CoUninitialize();
	return OK; // ErrorLevel indicates whether or not it succeeded.
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
		if (strlen(aDirSpec) >= sizeof(parent_dir)) // avoid overflow
			return OK; // Let ErrorLevel tell the story.
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
	__int64 line_number = ATOI64(aLineNumber);
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
	g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
	return OK;
}



ResultType Line::FileAppend(char *aFilespec, char *aBuf, FILE *aTargetFileAlreadyOpen)
{
	if (!aTargetFileAlreadyOpen && (!aFilespec || !*aFilespec)) // Nothing to write to (caller relies on this check).
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);

	// Don't do this because want to allow "nothing" to be written to a file in case the
	// user is doing this to reset it's timestamp:
	//if (!aBuf || !*aBuf)
	//	return g_ErrorLevel->Assign(ERRORLEVEL_NONE);

	FILE *fp;
	if (aTargetFileAlreadyOpen)
		fp = aTargetFileAlreadyOpen;
	else
	{
		bool open_as_binary = false;
		if (*aFilespec == '*')
		{
			open_as_binary = true;
			// Do not do this because I think it's possible for filenames to start with a space
			// (even though Explorer itself won't let you create them that way):
			//aFilespec = omit_leading_whitespace(aFilespec + 1);
			// Instead just do this:
			++aFilespec;
		}
		if (   !(fp = fopen(aFilespec, open_as_binary ? "ab" : "a"))   )
			return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	}
	g_ErrorLevel->Assign(fputs(aBuf, fp) ? ERRORLEVEL_ERROR : ERRORLEVEL_NONE); // fputs() returns 0 on success.
	if (!aTargetFileAlreadyOpen)
		fclose(fp);
	// else it's the caller's responsibility, or it's caller's, to close it.
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
	g_ErrorLevel->Assign(failure_count); // i.e. indicate success if there were no failures.
	return OK;
}



ResultType Line::FileRecycle(char *aFilePattern)
// Adapted from the AutoIt3 source.
{
	if (!aFilePattern || !*aFilePattern)
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);  // Since this is probably not what the user intended.

	SHFILEOPSTRUCT FileOp;
	char szFileTemp[_MAX_PATH+2];

	// au3: Get the fullpathname - required for UNDO to work
	Util_GetFullPathName(aFilePattern, szFileTemp);

	// au3: We must also make it a double nulled string *sigh*
	szFileTemp[strlen(szFileTemp)+1] = '\0';	

	// au3: set to known values - Corrects crash
	FileOp.hNameMappings = NULL;
	FileOp.lpszProgressTitle = NULL;
	FileOp.fAnyOperationsAborted = FALSE;
	FileOp.hwnd = NULL;
	FileOp.pTo = NULL;

	FileOp.pFrom = szFileTemp;
	FileOp.wFunc = FO_DELETE;
	FileOp.fFlags = FOF_SILENT | FOF_ALLOWUNDO | FOF_NOCONFIRMATION;

	// SHFileOperation() returns 0 on success:
	return g_ErrorLevel->Assign(SHFileOperation(&FileOp) ? ERRORLEVEL_ERROR : ERRORLEVEL_NONE);
}



ResultType Line::FileRecycleEmpty(char *aDriveLetter)
// Adapted from the AutoIt3 source.
{
	HINSTANCE hinstLib = LoadLibrary("shell32.dll");
	if (!hinstLib)
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	// au3: Get the address of all the functions we require
	typedef HRESULT (WINAPI *MySHEmptyRecycleBin)(HWND, LPCTSTR, DWORD);
 	MySHEmptyRecycleBin lpfnEmpty = (MySHEmptyRecycleBin)GetProcAddress(hinstLib, "SHEmptyRecycleBinA");
	if (!lpfnEmpty)
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	const char *szPath = *aDriveLetter ? aDriveLetter : NULL;
	if (lpfnEmpty(NULL, szPath, SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND) != S_OK)
		return g_ErrorLevel->Assign(ERRORLEVEL_ERROR);
	return g_ErrorLevel->Assign(ERRORLEVEL_NONE);
}



ResultType Line::FileInstall(char *aSource, char *aDest, char *aFlag)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	bool allow_overwrite = aFlag && *aFlag == '1';
#ifdef AUTOHOTKEYSC
	HS_EXEArc_Read oRead;
	// AutoIt3: Open the archive in this compiled exe.
	// Jon gave me some details about why a password isn't needed: "The code in those libararies will
	// only allow files to be extracted from the exe is is bound to (i.e the script that it was
	// compiled with).  There are various checks and CRCs to make sure that it can't be used to read
	// the files from any other exe that is passed."
	if (oRead.Open(g_script.mFileSpec, "") != HS_EXEARC_E_OK)
	{
		MsgBox(g_script.mFileSpec, 0, "Could not open this EXE to run the main script:");
		return OK; // Let ErrorLevel tell the story.
	}
	if (!allow_overwrite && Util_DoesFileExist(aDest))
	{
		oRead.Close();
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
	if (CopyFile(aSource, aDest, !allow_overwrite))
		g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
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
		// Now that the base directory of aFilePattern has been processed, recurse into
		// any subfolders to look for the same pattern.  Create a new pattern (don't reuse
		// file_pattern since naked_filename_or_pattern points into it) to be *.* so that
		// we can find all subfolders, not just those matching the original pattern:
		char all_file_pattern[MAX_PATH * 2];  // In case OS supports extra-long filenames.
		snprintf(all_file_pattern, sizeof(all_file_pattern), "%s*.*", file_path);
		file_search = FindFirstFile(all_file_pattern, &current_file);
		file_found = (file_search != INVALID_HANDLE_VALUE);
		for (; file_found; file_found = FindNextFile(file_search, &current_file))
		{
			LONG_OPERATION_UPDATE
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
		// not succeed.  Win95 (not sure about Win98/ME) does not support this, but it
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
		char all_file_pattern[MAX_PATH * 2];
		snprintf(all_file_pattern, sizeof(all_file_pattern), "%s*.*", file_path);
		file_search = FindFirstFile(all_file_pattern, &current_file);
		file_found = (file_search != INVALID_HANDLE_VALUE);
		for (; file_found; file_found = FindNextFile(file_search, &current_file))
		{
			LONG_OPERATION_UPDATE
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
// Util_CopyDir()
///////////////////////////////////////////////////////////////////////////////

bool Line::Util_CopyDir (const char *szInputSource, const char *szInputDest, bool bOverwrite)
{
	SHFILEOPSTRUCT	FileOp;
	char			szSource[_MAX_PATH+2];
	char			szDest[_MAX_PATH+2];

	// Get the fullpathnames and strip trailing \s
	Util_GetFullPathName(szInputSource, szSource);
	Util_GetFullPathName(szInputDest, szDest);

	// Ensure source is a directory
	if (Util_IsDir(szSource) == false)
		return false;							// Nope

	// Does the destination dir exist?
	if (Util_IsDir(szDest))
	{
		if (bOverwrite == false)
			return false;
	}
	else
	{
		// We must create the top level directory
		if (!Util_CreateDir(szDest))
			return false;
	}

	// To work under old versions AND new version of shell32.dll the source must be specifed
	// as "dir\*.*" and the destination directory must already exist... Godamn Microsoft and their APIs...
	strcat(szSource, "\\*.*");

	// We must also make source\dest double nulled strings for the SHFileOp API
	szSource[strlen(szSource)+1] = '\0';	
	szDest[strlen(szDest)+1] = '\0';	

	// Setup the struct
	FileOp.pFrom					= szSource;
	FileOp.pTo						= szDest;
	FileOp.hNameMappings			= NULL;
	FileOp.lpszProgressTitle		= NULL;
	FileOp.fAnyOperationsAborted	= FALSE;
	FileOp.hwnd						= NULL;

	FileOp.wFunc	= FO_COPY;
	FileOp.fFlags	= FOF_SILENT | FOF_NOCONFIRMMKDIR | FOF_NOCONFIRMATION | FOF_NOERRORUI;

	if ( SHFileOperation(&FileOp) ) 
		return false;								

	return true;

} // Util_CopyDir()



///////////////////////////////////////////////////////////////////////////////
// TAKEN FROM THE AUTOIT3 SOURCE
///////////////////////////////////////////////////////////////////////////////
// Util_MoveDir()
///////////////////////////////////////////////////////////////////////////////

bool Line::Util_MoveDir (const char *szInputSource, const char *szInputDest, bool bOverwrite)
{
	SHFILEOPSTRUCT	FileOp;
	char			szSource[_MAX_PATH+2];
	char			szDest[_MAX_PATH+2];

	// Get the fullpathnames and strip trailing \s
	Util_GetFullPathName(szInputSource, szSource);
	Util_GetFullPathName(szInputDest, szDest);

	// Ensure source is a directory
	if (Util_IsDir(szSource) == false)
		return false;							// Nope

	// Does the destination dir exist?
	if (Util_IsDir(szDest))
	{
		if (bOverwrite == false)
			return false;
	}

	// Now, if the source and dest are on different volumes then we must copy rather than move
	// as move in this case only works on some OSes
	if (Util_IsDifferentVolumes(szSource, szDest))
	{
		// Copy and delete (poor man's move)
		if (Util_CopyDir(szSource, szDest, true) == false)
			return false;
		if (Util_RemoveDir(szSource, true) == false)
			return false;
		else
			return true;
	}

	// We must also make source\dest double nulled strings for the SHFileOp API
	szSource[strlen(szSource)+1] = '\0';	
	szDest[strlen(szDest)+1] = '\0';	

	// Setup the struct
	FileOp.pFrom					= szSource;
	FileOp.pTo						= szDest;
	FileOp.hNameMappings			= NULL;
	FileOp.lpszProgressTitle		= NULL;
	FileOp.fAnyOperationsAborted	= FALSE;
	FileOp.hwnd						= NULL;

	FileOp.wFunc	= FO_MOVE;
	FileOp.fFlags	= FOF_SILENT | FOF_NOCONFIRMMKDIR | FOF_NOCONFIRMATION | FOF_NOERRORUI;

	if ( SHFileOperation(&FileOp) ) 
		return false;								
	else
		return true;

} // Util_MoveDir()



///////////////////////////////////////////////////////////////////////////////
// TAKEN FROM THE AUTOIT3 SOURCE
///////////////////////////////////////////////////////////////////////////////
// Util_RemoveDir()
///////////////////////////////////////////////////////////////////////////////

bool Line::Util_RemoveDir (const char *szInputSource, bool bRecurse)
{
	SHFILEOPSTRUCT	FileOp;
	char			szSource[_MAX_PATH+2];

	// Get the fullpathnames and strip trailing \s
	Util_GetFullPathName(szInputSource, szSource);

	// Ensure source is a directory
	if (Util_IsDir(szSource) == false)
		return false;							// Nope

	// If recursion not on just try a standard delete on the directory (the SHFile function WILL
	// delete a directory even if not empty no matter what flags you give it...)
	if (bRecurse == false)
	{
		if (!RemoveDirectory(szSource))
			return false;
		else
			return true;
	}

	// We must also make double nulled strings for the SHFileOp API
	szSource[strlen(szSource)+1] = '\0';	

	// Setup the struct
	FileOp.pFrom					= szSource;
	FileOp.pTo						= NULL;
	FileOp.hNameMappings			= NULL;
	FileOp.lpszProgressTitle		= NULL;
	FileOp.fAnyOperationsAborted	= FALSE;
	FileOp.hwnd						= NULL;

	FileOp.wFunc	= FO_DELETE;
	FileOp.fFlags	= FOF_SILENT | FOF_NOCONFIRMMKDIR | FOF_NOCONFIRMATION | FOF_NOERRORUI;
	
	if ( SHFileOperation(&FileOp) ) 
		return false;								

	return true;

} // Util_RemoveDir()



///////////////////////////////////////////////////////////////////////////////
// TAKEN FROM THE AUTOIT3 SOURCE
///////////////////////////////////////////////////////////////////////////////
// Util_CopyFile()
// (moves files too)
// Returns the number of files that could not be copied or moved due to error.
///////////////////////////////////////////////////////////////////////////////

int Line::Util_CopyFile(const char *szInputSource, const char *szInputDest, bool bOverwrite, bool bMove)
{
	WIN32_FIND_DATA	findData;
	HANDLE			hSearch;
	bool			bLoop;
	bool			bFound = false;				// Not found initially
	BOOL			bRes;

	char			szSource[_MAX_PATH+1];
	char			szDest[_MAX_PATH+1];
	char			szExpandedDest[MAX_PATH+1];
	char			szTempPath[_MAX_PATH+1];

	char			szDrive[_MAX_PATH+1];
	char			szDir[_MAX_PATH+1];
	char			szFile[_MAX_PATH+1];
	char			szExt[_MAX_PATH+1];
	bool			bDiffVol;

	// Get local version of our source/dest with full path names, strip trailing \s
	Util_GetFullPathName(szInputSource, szSource);
	Util_GetFullPathName(szInputDest, szDest);

	// Check if the files are on different volumes (affects how we do a Move operation)
	bDiffVol = Util_IsDifferentVolumes(szSource, szDest);

	// If the source or dest is a directory then add *.* to the end
	if (Util_IsDir(szSource))
		strcat(szSource, "\\*.*");
	if (Util_IsDir(szDest))
		strcat(szDest, "\\*.*");


	// Split source into file and extension (we need this info in the loop below to recontstruct the path)
	_splitpath( szSource, szDrive, szDir, szFile, szExt );

	// Note we now rely on the SOURCE being the contents of szDrive, szDir, szFile, etc.

	// Does the source file exist?
	hSearch = FindFirstFile(szSource, &findData);
	bLoop = true;

	// AutoHotkey:
	int failure_count = 0;
	LONG_OPERATION_INIT

	while (hSearch != INVALID_HANDLE_VALUE && bLoop == true)
	{
		LONG_OPERATION_UPDATE  // AutoHotkey
		bFound = true;							// Found at least one match

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

			// Does the destination exist? - delete it first if it does (unless we not overwriting)
			if ( Util_DoesFileExist(szExpandedDest) )
			{
				if (bOverwrite == false)
				{
					// AutoHotkey:
					++failure_count;
					if (FindNextFile(hSearch, &findData) == FALSE)
						bLoop = false;
					continue;
					//FindClose(hSearch);
					//return false;					// Destination already exists and we not overwriting
				}
				else
					// AutoHotkey:
					if (!DeleteFile(szExpandedDest))
					{
						++failure_count;
						if (FindNextFile(hSearch, &findData) == FALSE)
							bLoop = false;
						continue;
					}
			}
		
			// Move or copy operation?
			if (bMove == true)
			{
				if (bDiffVol == false)
				{
					bRes = MoveFile(szTempPath, szExpandedDest);
				}
				else
				{
					// Do a copy then delete (simulated copy)
					if ( (bRes = CopyFile(szTempPath, szExpandedDest, FALSE)) != FALSE )
						bRes = DeleteFile(szTempPath);
				}
			}
			else
				bRes = CopyFile(szTempPath, szExpandedDest, FALSE);
			
			if (bRes == FALSE)
			{
				// AutoHotkey:
				++failure_count;
				if (FindNextFile(hSearch, &findData) == FALSE)
					bLoop = false;
				continue;
				// FindClose(hSearch);
				// return false;						// Error copying/moving one of the files
			}

		} // End If

		if (FindNextFile(hSearch, &findData) == FALSE)
			bLoop = false;

	} // End while

	FindClose(hSearch);

	// AutoHotkey:
	return failure_count;
	//return bFound;

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
// Util_CreateDir()
// Recursive directory creation function
///////////////////////////////////////////////////////////////////////////////

bool Line::Util_CreateDir(const char *szDirName)
{
	DWORD	dwTemp;
	bool	bRes;
	char	*szTemp = NULL;
	char	*psz_Loc = NULL;

	dwTemp = GetFileAttributes(szDirName);
	if (dwTemp == FILE_ATTRIBUTE_DIRECTORY)
		return true;							// Directory exists, yay!

	if (dwTemp == 0xffffffff) 
	{	// error getting attribute - what was the error?
		if (GetLastError() == ERROR_PATH_NOT_FOUND) 
		{
			// Create path
			szTemp = new char[strlen(szDirName)+1];
			strcpy(szTemp, szDirName);
			psz_Loc = strrchr(szTemp, '\\');	/* find last \ */
			if (psz_Loc == NULL)				// not found
			{
				delete [] szTemp;
				return false;
			}
			else 
			{
				*psz_Loc = '\0';				// remove \ and everything after
				bRes = Util_CreateDir(szTemp);
				delete [] szTemp;
				if (bRes)
				{
					if (CreateDirectory(szDirName, NULL))
						bRes = true;
					else
						bRes = false;
				}

				return bRes;
			}
		} 
		else 
		{
			if (GetLastError() == ERROR_FILE_NOT_FOUND) 
			{
				// Create directory
				if (CreateDirectory(szDirName, NULL))
					return true;
				else
					return false;
			}
		}
	} 
			
	return false;								// Unforeseen error

} // Util_CreateDir()



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



///////////////////////////////////////////////////////////////////////////////
// TAKEN FROM THE AUTOIT3 SOURCE
///////////////////////////////////////////////////////////////////////////////
// Util_IsDir()
// Returns true if the path is a directory
///////////////////////////////////////////////////////////////////////////////

bool Line::Util_IsDir(const char *szPath)
{
	DWORD dwTemp;

	dwTemp = GetFileAttributes(szPath);
	if ( dwTemp != 0xffffffff && (dwTemp & FILE_ATTRIBUTE_DIRECTORY) )
		return true;
	else
		return false;

} // Util_IsDir



///////////////////////////////////////////////////////////////////////////////
// TAKEN FROM THE AUTOIT3 SOURCE
///////////////////////////////////////////////////////////////////////////////
// Util_GetFullPathName()
// Returns the full pathname and strips any trailing \s.  Assumes output
// is _MAX_PATH in size
///////////////////////////////////////////////////////////////////////////////

void Line::Util_GetFullPathName(const char *szIn, char *szOut)
{
	char	*szFilePart;

	GetFullPathName(szIn, _MAX_PATH, szOut, &szFilePart);
	Util_StripTrailingDir(szOut);
} // Util_GetFullPathName()



///////////////////////////////////////////////////////////////////////////////
// TAKEN FROM THE AUTOIT3 SOURCE
///////////////////////////////////////////////////////////////////////////////
// Util_StripTrailingDir()
//
// Makes sure a filename does not have a trailing //
//
///////////////////////////////////////////////////////////////////////////////

void Line::Util_StripTrailingDir(char *szPath)
{
	if (szPath[strlen(szPath)-1] == '\\')
		szPath[strlen(szPath)-1] = '\0';

} // Util_StripTrailingDir



///////////////////////////////////////////////////////////////////////////////
// TAKEN FROM THE AUTOIT3 SOURCE
///////////////////////////////////////////////////////////////////////////////
// Util_IsDifferentVolumes()
// Checks two paths to see if they are on the same volume
///////////////////////////////////////////////////////////////////////////////

bool Line::Util_IsDifferentVolumes(const char *szPath1, const char *szPath2)
{
	char			szP1Drive[_MAX_DRIVE+1];
	char			szP2Drive[_MAX_DRIVE+1];

	char			szDir[_MAX_DIR+1];
	char			szFile[_MAX_FNAME+1];
	char			szExt[_MAX_EXT+1];
	
	char			szP1[_MAX_PATH+1];	
	char			szP2[_MAX_PATH+1];

	// Get full pathnames
	Util_GetFullPathName(szPath1, szP1);
	Util_GetFullPathName(szPath2, szP2);

	// Split the target into bits
	_splitpath( szP1, szP1Drive, szDir, szFile, szExt );
	_splitpath( szP2, szP2Drive, szDir, szFile, szExt );

	if (szP1Drive[0] == '\0' || szP2Drive[0] == '\0')
	{
		// One or both paths is a UNC - assume different volumes
		return true;
	}
	else
	{
		if (stricmp(szP1Drive, szP2Drive))
			return true;
		else
			return false;
	}

} // Util_IsDifferentVolumes()



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
		case ACT_TRANSFORM:
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
		case ACT_CONTROLGET:
		case ACT_STATUSBARGETTEXT:
		case ACT_INPUTBOX:
		case ACT_RANDOM:
		case ACT_INIREAD:
		case ACT_REGREAD:
		case ACT_DRIVESPACEFREE:
		case ACT_DRIVEGET:
		case ACT_SOUNDGET:
		case ACT_SOUNDGETWAVEVOLUME:
		case ACT_FILEREADLINE:
		case ACT_FILEGETATTRIB:
		case ACT_FILEGETTIME:
		case ACT_FILEGETSIZE:
		case ACT_FILEGETVERSION:
		case ACT_FILESELECTFILE:
		case ACT_FILESELECTFOLDER:
		case ACT_MOUSEGETPOS:
		case ACT_WINGETTITLE:
		case ACT_WINGETCLASS:
		case ACT_WINGET:
		case ACT_WINGETTEXT:
		case ACT_WINGETPOS:
		case ACT_CONTROLGETPOS:
		case ACT_PIXELGETCOLOR:
		case ACT_PIXELSEARCH:
		//case ACT_IMAGESEARCH:
		case ACT_INPUT:
			return ARG_TYPE_OUTPUT_VAR;

		case ACT_SORT:
		case ACT_SPLITPATH:
		case ACT_IFINSTRING:
		case ACT_IFNOTINSTRING:
		case ACT_IFEQUAL:
		case ACT_IFNOTEQUAL:
		case ACT_IFGREATER:
		case ACT_IFGREATEROREQUAL:
		case ACT_IFLESS:
		case ACT_IFLESSOREQUAL:
		case ACT_IFIS:
		case ACT_IFISNOT:
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
		case ACT_STRINGSPLIT:
			return ARG_TYPE_INPUT_VAR;

		case ACT_MOUSEGETPOS:
		case ACT_WINGETPOS:
		case ACT_CONTROLGETPOS:
		case ACT_PIXELSEARCH:
		//case ACT_IMAGESEARCH:
		case ACT_SPLITPATH:
			return ARG_TYPE_OUTPUT_VAR;
		}
		break;

	case 2:  // Arg #3
		switch(aActionType)
		{
		case ACT_WINGETPOS:
		case ACT_CONTROLGETPOS:
		case ACT_MOUSEGETPOS:
		case ACT_SPLITPATH:
			return ARG_TYPE_OUTPUT_VAR;
		}
		break;

	case 3:  // Arg #4
		switch(aActionType)
		{
		case ACT_WINGETPOS:
		case ACT_CONTROLGETPOS:
		case ACT_MOUSEGETPOS:
		case ACT_SPLITPATH:
			return ARG_TYPE_OUTPUT_VAR;
		}
		break;

	case 4:  // Arg #5
	case 5:  // Arg #6 
		if (aActionType == ACT_SPLITPATH)
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
	case ACT_CONTROLSENDRAW:
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
	case ACT_WINGETPOS:
	case ACT_CONTROLGETPOS: // But don't bother valididing ACT_SPLITPATH this way since too rare to worry about.
		if (!ARG_HAS_VAR(1) && !ARG_HAS_VAR(2) && !ARG_HAS_VAR(3) && !ARG_HAS_VAR(4))
			return LineError(ERR_MISSING_OUTPUT_VAR);
		return OK;
	case ACT_PIXELSEARCH:
	//case ACT_IMAGESEARCH:
		if (!*RAW_ARG3 || !*RAW_ARG4 || !*RAW_ARG5 || !*RAW_ARG6 || !*RAW_ARG7)
			return LineError("Parameters 3 through 7 must not be blank.");
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



Line *Line::GetJumpTarget(bool aIsDereferenced)
{
	char *target_label = aIsDereferenced ? ARG1 : RAW_ARG1;
	Label *label = g_script.FindLabel(target_label);
	if (!label)
	{
		if (aIsDereferenced)
			LineError("This Goto/Gosub's target label does not exist." ERR_ABORT, FAIL, target_label);
		else
			LineError("This Goto/Gosub's target label does not exist.", FAIL, target_label);
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
	return LineError("A Goto/Gosub/GroupActivate mustn't jump into a block that doesn't enclose it.");
	// Above currently doesn't attempt to detect runtime vs. load-time for the purpose of appending
	// ERR_ABORT.
}
