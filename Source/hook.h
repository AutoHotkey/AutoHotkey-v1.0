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

#ifndef hook_h
#define hook_h

#include "StdAfx.h"  // Pre-compiled headers
#ifndef _MSC_VER  // For non-MS compilers:
	#include <windows.h>
#endif

#include "hotkey.h"

// WM_USER is the lowest number that can be a user-defined message.  Anything above that is also valid.
enum UserMessages {AHK_HOOK_HOTKEY = WM_USER, AHK_HOOK_TEST_MSG, AHK_DIALOG, AHK_NOTIFYICON, AHK_KEYLOG};


// Some reasoning behind the below data structures: Could build a new array for [sc][sc] and [vk][vk]
// (since only two keys are allowed in a ModifierVK/SC combination, only 2 dimensions are needed).
// But this would be a 512x512 array of shorts just for the SC part, which is 512K.  Instead, what we
// do is check whenever a key comes in: if it's a suffix and if a non-standard modifier key of any kind
// is currently down: consider action.  Most of the time, an action be found because the user isn't
// likely to be holding down a ModifierVK/SC, while pressing another key, unless it's modifying that key.
// Nor is he likely to have more than one ModifierVK/SC held down at a time.  It's still somewhat
// inefficient because have to look up the right prefix in a loop.  But most suffixes probably won't
// have more than one ModifierVK/SC anyway, so the lookup will usually find a match on the first
// iteration.

struct vk_hotkey
{
	vk_type vk;
	HotkeyIDType id_with_flags;
};
struct sc_hotkey
{
	sc_type sc;
	HotkeyIDType id_with_flags;
};



// User is likely to use more modifying vks than we do sc's, since sc's are rare:
#define MAX_MODIFIER_VKS_PER_SUFFIX 50
#define MAX_MODIFIER_SCS_PER_SUFFIX 16
// Style reminder: Any POD structs (those without any methods) don't use the "m" prefix
// for member variables because there's no need: the variables are always prefixed by
// the struct that owns them, so there's never any ambiguity:
struct key_type
{
	vk_hotkey ModifierVK[MAX_MODIFIER_VKS_PER_SUFFIX];
	sc_hotkey ModifierSC[MAX_MODIFIER_SCS_PER_SUFFIX];
	UCHAR nModifierVK;
	UCHAR nModifierSC;
//	vk_type toggleable_vk;  // If this key is CAPS/NUM/SCROLL-lock, its virtual key value is stored here.
	ToggleValueType *pForceToggle;  // Pointer to a global variable for toggleable keys only.  NULL for others.
	modLR_type as_modifiersLR; // If this key is a modifier, this will have the corresponding bit(s) for that key.
	bool used_as_prefix;  // whether a given virtual key or scan code is even used by a hotkey.
	bool used_as_suffix;  // whether a given virtual key or scan code is even used by a hotkey.
	bool no_mouse_suppress;  // Whether to omit the normal supression of a mouse hotkey; normally false.
	bool is_down; // this key is currently down.
	bool it_put_alt_down;  // this key resulted in ALT being pushed down (due to alt-tab).
	bool it_put_shift_down;  // this key resulted in SHIFT being pushed down (due to shift-alt-tab).
	bool down_performed_action;  // the last key-down resulted in an action (modifiers matched those of a valid hotkey)
	// The values for "was_just_used" (zero is the inialized default, meaning it wasn't just used):
	char was_just_used; // a non-modifier key of any kind was pressed while this prefix key was down.
	// And these are the values for the above (besides 0):
	#define AS_PREFIX 1
	#define AS_PREFIX_FOR_HOTKEY 2
	bool sc_takes_precedence; // used only by the scan code array: this scan code should take precedence over vk.
};


//-------------------------------------------

#define MAX_LOGGED_KEYS 50
struct KeyLogItem
{
	vk_type vk;
	sc_type sc;
	//LPARAM lParam;
	bool key_up;
	char event_type; // space=none, i=ignored, s=suppressed, h=hotkey, etc.
};


//-------------------------------------------


LRESULT CALLBACK LowLevelKeybdProc(int code, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK LowLevelMouseProc(int code, WPARAM wParam, LPARAM lParam);

HookType RemoveAllHooks();
HookType ChangeHookState(Hotkey *aHK[], int aHK_count, HookType aWhichHook, HookType aWhichHookAlways
, bool aWarnIfHooksAlreadyInstalled, bool aActivateOnlySuspendHotkeys = false);

char *GetHookStatus(char *aBuf, size_t aBufSize);

#endif
