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

#include "hook.h" // For KeyHistoryItem and probably other things.
#include "clipboard.h"  // For the global clipboard object
#include "script.h" // For the global script object and g_ErrorLevel
#include "os_version.h" // For the global OS_Version object

extern HINSTANCE g_hInstance;
extern HWND g_hWnd;  // The main window
extern HWND g_hWndEdit;  // The edit window, child of main.
extern HWND g_hWndSplash;  // The SplashText window.
extern HWND g_hWndToolTip;  // The tooltip window.
extern HACCEL g_hAccelTable; // Accelerator table for main menu shortcut keys.

extern modLR_type g_modifiersLR_logical;   // Tracked by hook (if hook is active).
extern modLR_type g_modifiersLR_logical_non_ignored;
extern modLR_type g_modifiersLR_physical;  // Same as above except it's which modifiers are PHYSICALLY down.

#ifdef FUTURE_USE_MOUSE_BUTTONS_LOGICAL
extern WORD g_mouse_buttons_logical; // A bitwise combination of MK_LBUTTON, etc.
#endif

#define STATE_DOWN 0x80
#define STATE_ON 0x01
extern BYTE g_PhysicalKeyState[VK_ARRAY_COUNT];

// If a SendKeys() operation takes longer than this, hotkey's modifiers won't be pressed back down:
extern int g_HotkeyModifierTimeout;

extern HHOOK g_KeybdHook;
extern HHOOK g_MouseHook;
#ifdef HOOK_WARNING
	extern HookType sWhichHookSkipWarning;
#endif
extern bool g_ForceLaunch;
extern bool g_WinActivateForce;
extern SingleInstanceType g_AllowOnlyOneInstance;
extern bool g_persistent;
extern bool g_NoTrayIcon;
#ifdef AUTOHOTKEYSC
	extern bool g_AllowMainWindow;
#endif
extern bool g_AllowSameLineComments;
extern char g_LastPerformedHotkeyType;
extern bool g_AllowInterruption;
extern bool g_AllowInterruptionForSub;
extern bool g_MainTimerExists;
extern bool g_UninterruptibleTimerExists;
extern bool g_AutoExecTimerExists;
extern bool g_InputTimerExists;
extern bool g_SoundWasPlayed;
extern bool g_IsSuspended;
extern int g_nLayersNeedingTimer;
extern int g_nThreads;
extern int g_nPausedThreads;
extern bool g_UnpauseWhenResumed;

// This value is the absolute limit:
#define MAX_THREADS_LIMIT 20
#define MAX_THREADS_DEFAULT 10
extern UCHAR g_MaxThreadsPerHotkey;
extern int g_MaxThreadsTotal;
extern int g_MaxHotkeysPerInterval;
extern int g_HotkeyThrottleInterval;
extern bool g_MaxThreadsBuffer;

extern MenuVisibleType g_MenuIsVisible;
extern int g_nMessageBoxes;
extern int g_nInputBoxes;
extern int g_nFileDialogs;
extern int g_nFolderDialogs;
extern InputBoxType g_InputBox[MAX_INPUTBOXES];

extern char g_delimiter;
extern char g_DerefChar;
extern char g_EscapeChar;

// Global objects:
extern Var *g_ErrorLevel;
extern input_type g_input;
EXTERN_SCRIPT;
EXTERN_CLIPBOARD;
EXTERN_OSVER;

extern int g_IconTray;
extern int g_IconTraySuspend;

extern DWORD g_OriginalTimeout;

EXTERN_G;
extern global_struct g_default;

extern char g_WorkingDir[MAX_PATH];  // Explicit size needed here in .h file for use with sizeof().
extern char *g_WorkingDirOrig;

// This macro is defined because sometimes g.hWndLastUsed will be out-of-date and the window
// may have been destroyed.  It also returns NULL if the current settings indicate that
// hidden windows should be ignored:
#define g_ValidLastUsedWindow (!g.hWndLastUsed ? NULL\
	: (!IsWindow(g.hWndLastUsed) ? NULL\
	: ((!g.DetectHiddenWindows && !IsWindowVisible(g.hWndLastUsed)) ? NULL : g.hWndLastUsed)))

extern bool g_ForceKeybdHook;
extern ToggleValueType g_ForceNumLock;
extern ToggleValueType g_ForceCapsLock;
extern ToggleValueType g_ForceScrollLock;

extern vk2_type g_sc_to_vk[SC_ARRAY_COUNT];
extern sc2_type g_vk_to_sc[VK_ARRAY_COUNT];

extern Action g_act[];
extern int g_ActionCount;
extern Action g_old_act[];
extern int g_OldActionCount;

extern key_to_vk_type g_key_to_vk[];
extern key_to_sc_type g_key_to_sc[];
extern int g_key_to_vk_count;
extern int g_key_to_sc_count;

extern KeyHistoryItem *g_KeyHistory;
extern int g_KeyHistoryNext;
extern DWORD g_HistoryTickNow;
extern DWORD g_HistoryTickPrev;
extern DWORD g_TimeLastInputPhysical;

#ifdef ENABLE_KEY_HISTORY_FILE
extern bool g_KeyHistoryToFile;
#endif


inline VarSizeType GetBatchLines(char *aBuf = NULL)
{
	char buf[256];
	char *target_buf = aBuf ? aBuf : buf;
	if (g.IntervalBeforeRest >= 0) // Have this new method take precedence, if it's in use by the script.
		sprintf(target_buf, "%dms", g.IntervalBeforeRest); // Not snprintf().
	else
		ITOA64(g.LinesPerCycle, target_buf);
	return (VarSizeType)strlen(target_buf);
}

inline VarSizeType GetOSType(char *aBuf = NULL)
{
	char *type = g_os.IsWinNT() ? "WIN32_NT" : "WIN32_WINDOWS";
	if (aBuf)
		strcpy(aBuf, type);
	return (VarSizeType)strlen(type); // Return length of type, not aBuf.
}

inline VarSizeType GetOSVersion(char *aBuf = NULL)
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
	if (aBuf)
		strcpy(aBuf, version);
	return (VarSizeType)strlen(version); // Always return length of version, not aBuf.
}

inline VarSizeType GetIsAdmin(char *aBuf = NULL)
// Adapted from AutoIt3 source.
{
	if (!aBuf)
		return 1;  // The length of the string "1" or "0".
	char result = '0';  // Default.
	if (g_os.IsWin9x())
		result = '1';
	else
	{
		SC_HANDLE h = OpenSCManager(NULL, NULL, SC_MANAGER_LOCK);
		if (h)
		{
			SC_LOCK lock = LockServiceDatabase(h);
			if (lock)
			{
				UnlockServiceDatabase(lock);
				result = '1';
			}
			else
			{
				DWORD lastErr = GetLastError();
				if (lastErr == ERROR_SERVICE_DATABASE_LOCKED)
					result = '1';
			}
			CloseServiceHandle(h);
		}
	}
	aBuf[0] = result;
	aBuf[1] = '\0';
	return 1; // Length of aBuf.
}

#endif
