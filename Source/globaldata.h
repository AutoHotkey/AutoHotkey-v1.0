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

#ifndef globaldata_h
#define globaldata_h

#include "hook.h" // For KeyLogItem and probably other things.
#include "clipboard.h"  // For the global clipboard object
#include "script.h" // For the global script object and g_ErrorLevel
#include "os_version.h" // For the global OS_Version object

extern HWND g_hWnd;  // The main window
extern HWND g_hWndEdit;  // The edit window, child of main.
extern HWND g_hWndSplash;  // The SplashText window.
extern HINSTANCE g_hInstance;
extern modLR_type g_modifiersLR_logical;   // Tracked by hook (if hook is active).
extern modLR_type g_modifiersLR_physical;  // Same as above except it's which modifiers are PHYSICALLY down.
extern modLR_type g_modifiersLR_get;  // From GetKeyState().

extern bool g_PhysicalKeyState[VK_MAX + 1];

// If a SendKeys() operation takes longer than this, hotkey's modifiers won't be pressed back down:
extern int g_HotkeyModifierTimeout;

extern HHOOK g_hhkLowLevelKeybd;
extern HHOOK g_hhkLowLevelMouse;
extern bool g_ForceLaunch;
extern bool g_AllowOnlyOneInstance;
extern bool g_AllowSameLineComments;
extern char g_LastPerformedHotkeyType;
extern bool g_IgnoreHotkeys;
extern bool g_IsSuspended;
#define IGNORE_THIS_HOTKEY(id) (g_IgnoreHotkeys || (g_IsSuspended && Hotkey::GetTypeOfFirstLine(id) != ACT_SUSPEND))
extern int g_nSuspendedSubroutines;

extern int g_MaxHotkeysPerInterval;
extern int g_HotkeyThrottleInterval;

extern bool g_TrayMenuIsVisible;
extern int g_nMessageBoxes;
extern int g_nInputBoxes;
extern int g_nFileDialogs;
extern InputBoxType g_InputBox[MAX_INPUTBOXES];

extern char g_delimiter;
extern char g_DerefChar;
extern char g_EscapeChar;

// Global objects:
extern Script g_script;
extern Var *g_ErrorLevel;
EXTERN_CLIPBOARD;
extern OS_Version g_os;

extern DWORD g_OriginalTimeout;

extern global_struct g, g_default;
// This macro is defined because sometimes g.hWndLastUsed will be out-of-date and the window
// may have been destroyed.  It also returns NULL if the current settings indicate that
// hidden windows should be ignored:
#define g_ValidLastUsedWindow (!g.hWndLastUsed ? NULL\
	: (!IsWindow(g.hWndLastUsed) ? NULL\
	: ((!g.DetectHiddenWindows && !IsWindowVisible(g.hWndLastUsed)) ? NULL : g.hWndLastUsed)))

extern ToggleValueType g_ForceKeybdHook;
extern ToggleValueType g_ForceNumLock;
extern ToggleValueType g_ForceCapsLock;
extern ToggleValueType g_ForceScrollLock;

extern vk2_type g_sc_to_vk[SC_MAX + 1];
extern sc2_type g_vk_to_sc[VK_MAX + 1];

extern Action g_act[];
extern int g_ActionCount;
extern Action g_old_act[];
extern int g_OldActionCount;

extern key_to_vk_type g_key_to_vk[];
extern key_to_sc_type g_key_to_sc[];
extern int g_key_to_vk_count;
extern int g_key_to_sc_count;

extern KeyLogItem g_KeyLog[MAX_LOGGED_KEYS];
extern int g_KeyLogNext;
extern bool g_KeyLogToFile;


inline int GetBatchLines(char *aBuf = NULL)
{
	char buf[2];  // In case some implementations of snprintf() aren't tolerant of a NULL pointer.
	return snprintf(aBuf ? aBuf : buf, aBuf ? 999 : 0, "%u", g.LinesPerCycle);
}

inline int GetOSType(char *aBuf = NULL)
{
	char *type;
	type = g_os.IsWinNT() ? "WIN32_NT" : "WIN32_WINDOWS";
	int length = (int)strlen(type);
	if (!aBuf)
		return length;
	strcpy(aBuf, type);
	return length;
}

inline int GetOSVersion(char *aBuf = NULL)
// Adapted from AutoIt3 source.
{
	char *version;
	if (g_os.IsWinNT())
	{
		if (g_os.IsWinXP())
			version = "WIN_XP";
		else
		{
			if (g_os.IsWin2000())
				version = "WIN_2000";
			else
				version = "WIN_NT4";
		}
	}
	else
	{
		if (g_os.IsWin95())
			version = "WIN_95";
		else
		{
			if (g_os.IsWin98())
				version = "WIN_98";
			else
				version = "WIN_ME";
		}
	}
	int length = (int)strlen(version);
	if (!aBuf)
		return length;
	strcpy(aBuf, version);
	return length;
}

#endif
