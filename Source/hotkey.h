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

#ifndef hotkey_h
#define hotkey_h

#include "keyboard.h"
#include "script.h"  // For which label (and in turn which line) in the script to jump to.

// Due to control/alt/shift modifiers, quite a lot of hotkey combinations are possible, so support any
// conceivable use.  Note: Increasing this value will increase the memory required (i.e. any arrays
// that use this value):
#define MAX_HOTKEYS 512

// Note: 0xBFFF is the largest ID that can be used with RegisterHotkey().
// But further limit this to 0x3FFF (16,383) so that the two highest order bits
// are reserved for our other uses:
#define HOTKEY_NO_SUPPRESS             0x8000
#define HOTKEY_FUTURE_USE              0x4000
#define HOTKEY_ID_MASK                 0x3FFF
#define HOTKEY_ID_INVALID              HOTKEY_ID_MASK
#define HOTKEY_ID_ALT_TAB              0x3FFE
#define HOTKEY_ID_ALT_TAB_SHIFT        0x3FFD
#define HOTKEY_ID_ALT_TAB_MENU         0x3FFC
#define HOTKEY_ID_ALT_TAB_AND_MENU     0x3FFB
#define HOTKEY_ID_ALT_TAB_MENU_DISMISS 0x3FFA
#define HOTKEY_ID_MAX                  0x3FF9

#define COMPOSITE_DELIMITER " & "

// Smallest workable size: to save mem in some large arrays that use this:
typedef USHORT HotkeyIDType;
typedef HotkeyIDType HookActionType;
enum HotkeyTypes {HK_UNDETERMINED, HK_NORMAL, HK_KEYBD_HOOK, HK_MOUSE_HOOK};


class Hotkey
{
private:
	// These are done as static, rather than having an outer-class to contain all the hotkeys, because
	// the hotkey ID is used as the array index for performance reasons.  Having an outer class implies
	// the potential future use of more than one set of hotkeys, which could still be implemented
	// within static data and methods to retain the indexing/performance method:
	static bool sHotkeysAreActive; // Whether the hotkeys are in effect.
	static HookType sWhichHookNeeded;
	static HookType sWhichHookActive;
	static DWORD sTimePrev;
	static DWORD sTimeNow;
	static Hotkey *shk[MAX_HOTKEYS];
	static HotkeyIDType sNextID;

	HotkeyIDType mID;  // Must be unique for each hotkey of a given thread.
	vk_type mVK; // virtual-key code, e.g. VK_TAB, VK_LWIN, VK_LMENU, VK_APPS, VK_F10.  If zero, use sc below.
	sc_type mSC; // Scan code.  All vk's have a scan code, but not vice versa.
	mod_type mModifiers;  // MOD_ALT, MOD_CONTROL, MOD_SHIFT, MOD_WIN, or some additive or bitwise-or combination of these.
	modLR_type mModifiersLR;  // Left-right centric versions of the above.
	bool mAllowExtraModifiers;  // False if the hotkey should not fire when extraneous modifiers are held down.
	bool mDoSuppress;  // Normally true, but can be overridden by using the hotkey ~ prefix.
	vk_type mModifierVK; // Any other virtual key that must be pressed down in order to activate "vk" itself.
	sc_type mModifierSC; // If mModifierVK is zero, this scan code, if non-zero, will be used as the modifier.
	modLR_type mModifiersConsolidated; // The combination of mModifierVK, mModifierSC, mModifiersLR, modifiers
	HotkeyTypes mType;
	bool mIsRegistered;  // Whether this hotkey has been successfully registered.
	HookActionType mHookAction;
	Label *mJumpToLabel;
	bool mConstructedOK;

	// Something in the compiler is currently preventing me from using HookType in place of UCHAR:
	friend UCHAR HookInit(Hotkey *aHK[], int aHK_count, UCHAR aHooksToActivate);

	ResultType TextInterpret();
	char *TextToModifiers(char *aText);
	int TextToKey(char *aText, bool aIsModifier);
	ResultType Perform() {return mJumpToLabel->mJumpToLine->ExecUntil(UNTIL_RETURN, mModifiersConsolidated);}
	ResultType Register();
	ResultType Unregister();
	static ResultType AllDeactivate();
	static ResultType AllDestruct();

	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}

	// For now, constructor & destructor are private so that only static methods can create new
	// objects.  This allow proper tracking of which OS hotkey IDs have been used.
	Hotkey(HotkeyIDType aID, Label *aJumpToLabel, HookActionType aHookAction);
	~Hotkey() {if (mIsRegistered) Unregister();}
public:
	// Make sHotkeyCount an alias for sNextID.  Make it const to enforce modifying the value in only one way:
	static const HotkeyIDType &sHotkeyCount;

	static void AllDestructAndExit(int exit_code);
	static ResultType AddHotkey(Label *aJumpToLabel, HookActionType aHookAction);
	static ResultType PerformID(HotkeyIDType aHotkeyID);
	static ActionTypeType GetTypeOfFirstLine(HotkeyIDType aHotkeyID)
	{
		// Currently, hotkey_id can't be < 0 due to its type, so we only check if it's too large:
		if (aHotkeyID >= sHotkeyCount) return ACT_INVALID;
		if (shk[aHotkeyID]->mJumpToLabel == NULL) return ACT_INVALID;
		return shk[aHotkeyID]->mJumpToLabel->mJumpToLine->mActionType;
	}
	static int AllActivate();
	static void RequireHook(HookType aWhichHook) {sWhichHookNeeded |= aWhichHook;}
	static HookType HookIsActive() {return sWhichHookActive;} // Returns bitwise values: HOOK_MOUSE, HOOK_KEYBD.

	static char GetType(HotkeyIDType aHotkeyID)
	{
		return(aHotkeyID >= 0 && aHotkeyID < sHotkeyCount) ? shk[aHotkeyID]->mType : -1;
	}

	static Label *GetLabel(HotkeyIDType aHotkeyID)
	{
		return(aHotkeyID >= 0 && aHotkeyID < sHotkeyCount) ? shk[aHotkeyID]->mJumpToLabel : NULL;
	}

	static int FindHotkeyBySC(sc2_type aSC2, mod_type aModifiers, modLR_type aModifiersLR);
	static int FindHotkeyWithThisModifier(vk_type aVK, sc_type aSC);
	static int FindHotkeyContainingModLR(modLR_type aModifiersLR);  //, int hotkey_id_to_omit);

	static char *ListHotkeys(char *aBuf, size_t aBufSize);
	char *ToText(char *aBuf, size_t aBufSize, bool aAppendNewline);
};
#endif
