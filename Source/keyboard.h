/*
AutoHotkey

Copyright 2003 Christopher L. Mallett

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
*/

#ifndef keyboard_h
#define keyboard_h

#include "defines.h"

// The max number of keystrokes to Send prior to taking a break to pump messages:
#define MAX_LUMP_KEYS 50

// Maybe define more of these later, perhaps with ifndef (since they should be in the normal header, and probably
// will be eventually):
// ALREADY DEFINED: #define VK_HELP 0x2F
// In case a compiler with a non-updated header file is used:
#ifndef VK_BROWSER_BACK
	#define VK_BROWSER_BACK        0xA6
	#define VK_BROWSER_FORWARD     0xA7
	#define VK_BROWSER_REFRESH     0xA8
	#define VK_BROWSER_STOP        0xA9
	#define VK_BROWSER_SEARCH      0xAA
	#define VK_BROWSER_FAVORITES   0xAB
	#define VK_BROWSER_HOME        0xAC
	#define VK_VOLUME_MUTE         0xAD
	#define VK_VOLUME_DOWN         0xAE
	#define VK_VOLUME_UP           0xAF
	#define VK_MEDIA_NEXT_TRACK    0xB0
	#define VK_MEDIA_PREV_TRACK    0xB1
	#define VK_MEDIA_STOP          0xB2
	#define VK_MEDIA_PLAY_PAUSE    0xB3
	#define VK_LAUNCH_MAIL         0xB4
	#define VK_LAUNCH_MEDIA_SELECT 0xB5
	#define VK_LAUNCH_APP1         0xB6
	#define VK_LAUNCH_APP2         0xB7
#endif

// Create some "fake" virtual keys to simplify sections of the code.
// According to winuser.h, the following ranges (among others)
// are considered "unassigned" rather than "reserved", so should be
// fairly safe to use for the foreseeable future.  0xFF should probably
// be avoided since it's sometimes used as a failure indictor by API
// calls.  And 0x00 should definitely be avoided because it is used
// to indicate failure by many functions that deal with virtual keys:
// 0x88 - 0x8F : unassigned
// 0x97 - 0x9F : unassigned (this range seems less likely to be used)
#define VK_WHEEL_DOWN 0x9E
#define VK_WHEEL_UP 0x9F

// These are the only keys for which another key with the same VK exists.  Therefore, use scan code for these.
// If use VK for some of these (due to them being more likely to be used as hotkeys, thus minimizing the
// use of the keyboard hook), be sure to use SC for its counterpart.
// Always use the compressed version of scancode, i.e. 0x01 for the high-order byte rather than vs. 0xE0.
#define SC_NUMPADENTER 0x11C
#define SC_INSERT 0x152
#define SC_DELETE 0x153
#define SC_HOME 0x147
#define SC_END 0x14F
#define SC_UP 0x148
#define SC_DOWN 0x150
#define SC_LEFT 0x14B
#define SC_RIGHT 0x14D
#define SC_PGUP 0x149
#define SC_PGDN 0x151

// These are the same scan codes as their counterpart except the extended flag is 0 rather than
// 1 (0xE0 uncompressed):
#define SC_ENTER 0x1C
// In addition, the below dual-state numpad keys share the same scan code (but different vk's)
// regardless of the state of numlock:
#define SC_NUMPADDEL 0x53
#define SC_NUMPADINS 0x52
#define SC_NUMPADEND 0x4F
#define SC_NUMPADHOME 0x47
#define SC_NUMPADCLEAR 0x4C
#define SC_NUMPADUP 0x48
#define SC_NUMPADDOWN 0x50
#define SC_NUMPADLEFT 0x4B
#define SC_NUMPADRIGHT 0x4D
#define SC_NUMPADPGUP 0x49
#define SC_NUMPADPGDN 0x51

#define SC_NUMPADDOT SC_NUMPADDEL
#define SC_NUMPAD0 SC_NUMPADINS
#define SC_NUMPAD1 SC_NUMPADEND
#define SC_NUMPAD2 SC_NUMPADDOWN
#define SC_NUMPAD3 SC_NUMPADPGDN
#define SC_NUMPAD4 SC_NUMPADLEFT
#define SC_NUMPAD5 SC_NUMPADCLEAR
#define SC_NUMPAD6 SC_NUMPADRIGHT
#define SC_NUMPAD7 SC_NUMPADHOME
#define SC_NUMPAD8 SC_NUMPADUP
#define SC_NUMPAD9 SC_NUMPADPGUP

// These both have a unique vk and a unique sc (on most keyboards?), but they're listed here because
// MapVirtualKey doesn't support them under Win9x (except maybe NumLock itself):
#define SC_NUMLOCK 0x145
#define SC_NUMPADDIV 0x135
#define SC_NUMPADMULT 0x037
#define SC_NUMPADSUB 0x04A
#define SC_NUMPADADD 0x04E

// Note: A KeyboardProc() (hook) actually receives 0x36 for RSHIFT under both WinXP and Win98se, not 0x136.
// All the below have been verified to be accurate under Win98se and XP (except rctrl and ralt in XP).
#define SC_LCONTROL 0x01D
#define SC_RCONTROL 0x11D
#define SC_LSHIFT 0x02A
#define SC_RSHIFT 0x036
#define SC_LALT 0x038
#define SC_RALT 0x138
#define SC_LWIN 0x15B
#define SC_RWIN 0x15C

typedef UCHAR vk_type;  // Probably better than using UCHAR, since UCHAR might be two bytes.
// Although only need 9 bits for compressed and 16 for uncompressed scan code, use a full 32 bits so that
// mouse messages (WPARAM) can be stored as scan codes.  Formerly USHORT (which is always 16-bit).
typedef UINT sc_type;
typedef UINT mod_type; // Standard windows modifier type.


// The maximum number of virtual keys and scan codes that can ever exist.
// As of WinXP, these are absolute limits, except for scan codes for which there might conceivably
// be more if any non-standard keyboard or keyboard drivers generate scan codes that don't start
// with either 0x00 or 0xE0.  UPDATE: Decided to allow all possible scancodes, rather than just 512,
// since a lookup array for the 16-bit scan code value will only consume 64K of RAM if the element
// size is one char.  UPDATE: decided to go back to 512 scan codes, because WinAPI's KeyboardProc()
// itself can only handle that many (a 9-bit value).  254 is the largest valid vk, according to the
// WinAPI docs (I think 255 is value that is sometimes returned to indicate an invalid vk).  But
// just in case something ever tries to access such arrays using the largest 8-bit value (255), add
// 1 to make it 0xFF, thus ensuring array indexes will always be in-bounds if they are 8-bit values.
#define VK_MAX 0xFF
#define SC_MAX 0x1FF


typedef UCHAR modLR_type; // Only the left-right win/alt/ctrl/shift rather than their generic counterparts.
#define MODLR_MAX 0xFF
#define MOD_LCONTROL 0x01
#define MOD_RCONTROL 0x02
#define MOD_LALT 0x04
#define MOD_RALT 0x08
#define MOD_LSHIFT 0x10
#define MOD_RSHIFT 0x20
#define MOD_LWIN 0x40
#define MOD_RWIN 0x80


struct key_to_vk_type // Map key names to virtual keys.
{
	char *key_name;
	vk_type vk;
};

struct key_to_sc_type // Map key names to scan codes.
{
	char *key_name;
	sc_type sc;
};


// SC_MAX + 1 in case anyone ever tries to reference g_sc_to_vk[SC_MAX] itself:
struct vk2_type
{
	vk_type a;
	vk_type b;
};

struct sc2_type
{
	sc_type a;
	sc_type b;
};


void SendKeys(char *aKeys, modLR_type aModifiersLR = 0, HWND aTargetWindow = NULL);
void SendKey(vk_type aVK, sc_type aSC, mod_type aModifiers, int aRepeatCount, HWND aTargetWindow = NULL);

enum KeyEventTypes {KEYDOWN, KEYUP, KEYDOWNANDUP};
#define KEYIGNORE 0xA1C3D44F
ResultType KeyEvent(KeyEventTypes aEventType, vk_type aVK, sc_type aSC = 0
	, HWND aTargetWindow = NULL, DWORD aExtraInfo = KEYIGNORE);


ToggleValueType ToggleKeyState(vk_type vk, ToggleValueType aToggleValue);

modLR_type SetModifierState(mod_type aModifiersNew, modLR_type aModifiersLRnow);
modLR_type SetModifierLRState(modLR_type modifiersLRnew, modLR_type aModifiersLRnow);
void SetModifierLRStateSpecific(modLR_type aModifiersLR, modLR_type aModifiersLRnow, KeyEventTypes aKeyUp);

modLR_type GetModifierLRState();
modLR_type GetModifierLRStateSimple();

// Doesn't always seem to work as advertised, at least under WinXP:
#define IsPhysicallyDown(vk) (GetAsyncKeyState(vk) & 0x80000000)

modLR_type KeyToModifiersLR(vk_type aVK, sc_type aSC = 0, bool *pIsNeutral = NULL);
modLR_type ConvertModifiers(mod_type aModifiers);
mod_type ConvertModifiersLR(modLR_type aModifiersLR);
char *ModifiersLRToText(modLR_type aModifiersLR, char *aBuf);

//---------------------------------------------------------------------

void init_vk_to_sc();
void init_sc_to_vk();
sc_type TextToSC(char *aText);
vk_type TextToVK(char *aText, mod_type *pModifiers);
int TextToSpecial(char *aText, UINT aTextLength, mod_type *pModifiers);

#endif
