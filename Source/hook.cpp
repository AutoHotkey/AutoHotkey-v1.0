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

#include "StdAfx.h"  // Pre-compiled headers
#ifndef _MSC_VER  // For non-MS compilers:
	#include <stdlib.h> // for qsort()
#endif

#include "hook.h"
#include "globaldata.h"  // for access to several global vars
#include "hotkey.h" // HookInit() reads directly from static Hotkey class vars.
#include "util.h" // for snprintfcat()

// Declare static variables (global to only this file/module, i.e. no
// external linkage):

// Whether to disguise the next up-event for lwin/rwin to suppress Start Menu.
// These are made global, rather than static inside the hook function, so that
// we can ensure they are initialized by the keyboard init function every
// time it's called (currently it can be only called once):
static bool disguise_next_lwin_up = false;
static bool disguise_next_rwin_up = false;
static bool alt_tab_menu_is_visible = false;

// The prefix key that's currently down (i.e. in effect).
// It's tracked this way, rather than as a count of the number of prefixes currently down, out of
// concern that such a count might accidentally wind up above zero (due to a key-up being missed somehow)
// and never come back down, thus penalizing performance until the program is restarted:
static key_type *pPrefixKey = NULL;

static key_type *kvk = NULL;
static key_type *ksc = NULL;
// Since index zero is a placeholder for the invalid virtual key or scan code, add one to each MAX value
// to compute the number of elements actually needed to accomodate 0 up to and including VK_MAX or SC_MAX:
#define KVK_SIZE (VK_MAX + 1)
#define KSC_SIZE (SC_MAX + 1)

// Less memory overhead (space and performance) to allocate a solid block for multidimensional arrays:
// These store all the valid modifier+suffix combinations (those that result in hotkey actions) except
// those with a ModifierVK/SC.  Doing it this way should cut the CPU overhead caused by having many
// hotkeys handled by the hook, down to a fraction of what it would be otherwise.  Indeed, doing it
// this way makes the performance impact of adding many additional hotkeys of this type exactly zero
// once the program has started up and initialized.  The main alternative is a binary search on an
// array of keyboard-hook hotkeys (similar to how the mouse is done):
static HotkeyIDType *kvkm = NULL;
static HotkeyIDType *kscm = NULL;
// Macros for convenience in accessing the above arrays as multidimensional objects.
// When using them, be sure to consistently access the first index as ModLR (i.e. the rows)
// and the second as VK or SC (i.e. the columns):
#define Kvkm(i,j) kvkm[(i)*(MODLR_MAX + 1) + (j)]
#define Kscm(i,j) kscm[(i)*(MODLR_MAX + 1) + (j)]
#define KVKM_SIZE ((MODLR_MAX + 1)*(VK_MAX + 1))
#define KSCM_SIZE ((MODLR_MAX + 1)*(SC_MAX + 1))



// This next part is done this way because there doesn't seem any way to safely use the
// same HookProc for both the mouse and keybd hook.  MSDN states "nCode [in] Specifies a
// code the hook procedure uses to determine how to process the message. If nCode is less
// than zero, the hook procedure must pass the message to the CallNextHookEx".
// But the problem is that CallNextHookEx() requires the handle to the hook that called it,
// and we wouldn't know whether to send the mouse or the keybd handle.  This is because
// we're not supposed to look at the other args (wParam & lParam) received by the Hook
// when nCode < 0, because those args might have no meaning or even have random values:

// Define the mouse hook function:
#include "hook_include.cpp"
// Define the keybd hook function:
#define INCLUDE_KEYBD_HOOK
#include "hook_include.cpp"



struct hk_sorted_type
{
	int id_with_flags;
	vk_type vk;
	sc_type sc;
	mod_type modifiers;
	modLR_type modifiersLR;
	bool AllowExtraModifiers;
};



int sort_most_general_before_least(const void *a1, const void *a2)
// The only items whose order are important are those with the same suffix.  For a given suffix,
// we want the most general modifiers (e.g. CTRL) to appear closer to the top of the list than
// those with more specific modifiers (e.g. CTRL-ALT).  To make qsort() perform properly, it seems
// best to sort by vk/sc then by generality.
{
	// It's probably not necessary to be so thorough.  For example, if a1 has a vk but a2 has an sc,
	// those two are immediately non-equal.  But I'm worried about consistency: qsort() may get messed
	// up if these same two objects are ever compared, in reverse order, but a different comparison
	// result is returned.  Therefore, we compare rigorously and consistently:
	if (   ((hk_sorted_type *)a1)->vk && ((hk_sorted_type *)a2)->vk   )
		if (   ((hk_sorted_type *)a1)->vk != ((hk_sorted_type *)a2)->vk   )
			return ((hk_sorted_type *)a1)->vk - ((hk_sorted_type *)a2)->vk;
	if (   ((hk_sorted_type *)a1)->sc && ((hk_sorted_type *)a2)->sc   )
		if (   ((hk_sorted_type *)a1)->sc != ((hk_sorted_type *)a2)->sc   )
			return ((hk_sorted_type *)a1)->sc - ((hk_sorted_type *)a2)->sc;
	if (   ((hk_sorted_type *)a1)->vk && !((hk_sorted_type *)a2)->vk   )
		return 1;
	if (   !((hk_sorted_type *)a1)->vk && ((hk_sorted_type *)a2)->vk   )
		return -1;

	// If the above didn't return, we now know that a1 and a2 have the same vk's or sc's.  So
	// we use a tie-breaker to cause the most general keys to appear closer to the top of the
	// list than less general ones.  This should result in a given suffix being grouped together
	// after the sort.  Within each suffix group, the most general modifiers should appear first.

	// This part is basically saying that keys that don't allow extra modifiers can always be processed
	// after all other keys:
	if (((hk_sorted_type *)a1)->AllowExtraModifiers && !((hk_sorted_type *)a2)->AllowExtraModifiers)
		return -1;  // Indicate that a1 is smaller, so that it will go to the top.
	if (!((hk_sorted_type *)a1)->AllowExtraModifiers && ((hk_sorted_type *)a2)->AllowExtraModifiers)
		return 1;

	// However the order of suffixes that don't allow extra modifiers, among themselves, may be important.
	// Thus we don't return a zero if both have AllowExtraModifiers = 0.
	// Example: User defines ^a, but also defines >^a.  What should probably happen is that >^a fores ^a
	// to fire only when <^a occurs.

	mod_type mod_a1_merged = ((hk_sorted_type *)a1)->modifiers;
	mod_type mod_a2_merged = ((hk_sorted_type *)a2)->modifiers;
	if (((hk_sorted_type *)a1)->modifiersLR)
		mod_a1_merged |= ConvertModifiersLR(((hk_sorted_type *)a1)->modifiersLR);
	if (((hk_sorted_type *)a2)->modifiersLR)
		mod_a2_merged |= ConvertModifiersLR(((hk_sorted_type *)a2)->modifiersLR);

	// Check for equality first to avoid a possible infinite loop where two identical sets are subsets of each other:
	if (mod_a1_merged == mod_a1_merged)
	{
		// Here refine it further to handle a case such as ^a and >^a.  We want ^a to be considered
		// more general so that it won't override >^a altogether:
		if (   ((hk_sorted_type *)a1)->modifiersLR && !((hk_sorted_type *)a2)->modifiersLR   )
			return 1;  // Make a1 greater, so that it goes below a2 on the list.
		if (   !((hk_sorted_type *)a1)->modifiersLR && ((hk_sorted_type *)a2)->modifiersLR   )
			return -1;
		// After the above, the only remaining possible-problem case in this block is that
		// a1 and a2 have non-zero modifiersLRs that are different.  e.g. >+^a and +>^a
		// I don't think I want to try to figure out which of those should take precedence,
		// and how they overlap.  Maybe another day.
		return 0;
	}

	mod_type mod_intersect = mod_a1_merged & mod_a2_merged;

	if (mod_a1_merged == mod_intersect)
		// a1's modifiers are containined entirely within a2's, thus a1 is more general and
		// should be considered smaller so that it will go closer to the top of the list:
		return -1;
	if (mod_a2_merged == mod_intersect)
		return 1;

	// Otherwise, since neither is a perfect subset of the other, report that they're equal.
	// More refinement might need to be done here later for modifiers that partially overlap:
	// e.g. At this point is it possible for a1's modifiersLR to be a perfect subset of a2's,
	// or vice versa?
	return 0;
}


void SetModifierAsPrefix(vk_type aVK, sc_type aSC, bool aAlwaysSetAsPrefix = false)
// The caller has already ensured that vk and/or sc is a modifier such as VK_CONTROL.
{
	if (!aVK && !aSC) return;
	if (aVK)
	{
		switch (aVK)
		{
		case VK_MENU:
			// Since the user is configuring both the left and right
			// counterparts of a key to perform a suffix action,
			// it seems best to always consider those keys to be
			// prefixes so that their suffix action will only fire
			// when the key is released.  That way, those keys can
			// still be used as normal modifiers.
			kvk[VK_MENU].used_as_prefix = kvk[VK_LMENU].used_as_prefix = kvk[VK_RMENU].used_as_prefix
				= ksc[SC_LALT].used_as_prefix = ksc[SC_RALT].used_as_prefix = true;
			break;
		case VK_SHIFT:
			kvk[VK_SHIFT].used_as_prefix = kvk[VK_LSHIFT].used_as_prefix = kvk[VK_RSHIFT].used_as_prefix
				= ksc[SC_LSHIFT].used_as_prefix = ksc[SC_RSHIFT].used_as_prefix = true;
			break;
		case VK_CONTROL:
			kvk[VK_CONTROL].used_as_prefix = kvk[VK_LCONTROL].used_as_prefix = kvk[VK_RCONTROL].used_as_prefix
				= ksc[SC_LCONTROL].used_as_prefix = ksc[SC_RCONTROL].used_as_prefix = true;
			break;
		default:  // vk is a left/right modifier key such as VK_LCONTROL or VK_LWIN:
			if (aAlwaysSetAsPrefix)
				kvk[aVK].used_as_prefix = true;
			else
				if (Hotkey::FindHotkeyContainingModLR(kvk[aSC].as_modifiersLR) >= 0)
					kvk[aVK].used_as_prefix = true;
				// else allow its suffix action to fire when key is pressed down,
				// under the fairly safe assumption that the user hasn't configured
				// the opposite key to also be a key-down suffix-action (but even
				// if the user has done this, it's an explicit override of the
				// safety checks here, so probably best to allow it).
		}
		return;
	}
	// Since above didn't return, using scan code instead:
	if (aAlwaysSetAsPrefix)
		ksc[aSC].used_as_prefix = true;
	else
		if (Hotkey::FindHotkeyContainingModLR(ksc[aSC].as_modifiersLR) >= 0)
			ksc[aSC].used_as_prefix = true;
}



HookType HookInit(Hotkey *aHK[], int aHK_count, HookType aHooksToActivate)
// The input params are probably unnecessary because could just
// access directly by using Hotkey::shk[] ... but aHK is a little
// more concise.
// Returns the set of hooks that is currently active.
{
	// Don't allow caller to ask us to deactivate a currently-active hook.
	// The caller should have used HookTerm() for that.  UPDATE: It's okay
	// if caller asks us to install only one of the hooks but the other
	// is already installed.  The convention now is that the caller isn't
	// asking us to turn anything off; it's just specifying what it wants
	// turned on:
	//if (g_hhkLowLevelKeybd && !(aHooksToActivate & HOOK_KEYBD))
	//	return HOOK_FAIL;
	//if (g_hhkLowLevelMouse && !(aHooksToActivate & HOOK_MOUSE))
	//	return HOOK_FAIL;

	HookType hooks_currently_active = 0;
	if (g_hhkLowLevelKeybd)
		hooks_currently_active |= HOOK_KEYBD;
	if (g_hhkLowLevelMouse)
		hooks_currently_active |= HOOK_MOUSE;

	// The below effectively makes it necessary for the caller to call HookTerm()
	// first if it wishes us to reload the hotkey configuration:
	if ((aHooksToActivate & hooks_currently_active) == aHooksToActivate)
		// Bitwise-AND does set intersection.  In this case, the above means that
		// aHooksToActivate is a perfect subset of hooks_currently_active,
		// which means that all the requested hooks are already installed.
		return hooks_currently_active;
	// else at least one of the hooks needs to be activated, and the below relies
	// upon this being true.

	// No longer applicable:
	// But do allow the function to continue even if the hooks specified in aHooksToActivate
	// are already active, so that the data structures can be reinitialized below to reflect
	// changes the caller wants made.  However, support for that has not been fully reviewed
	// or tested.

	// It's valid for this be zero, such as when the user specifies that
	// the Numlock key should be forced on, but nothing else is handled
	// by the hook:
	//if (!aHK_count) return 0;

	// These arrays are dynamically allocated so that memory is conserved in cases when
	// the user doesn't need the hook at all (i.e. just normal registered hotkeys).
	// This is a waste of memory if there are no hook hotkeys, but currently the operation
	// of the hook relies upon these being allocated, even if the arrays are all clean
	// slates with nothing in them.  Presumably, the caller is requesting the keyboard
	// hook with zero hotkeys to support the forcing of Num/Caps/ScrollLock always on
	// or off:
	kvk = ksc = NULL;
	kvkm = kscm = NULL;
	if (NULL != (kvk = new key_type[KVK_SIZE]))
		if (NULL != (ksc = new key_type[KSC_SIZE]))
			if (NULL != (kvkm = new HotkeyIDType[KVKM_SIZE]))
				kscm = new HotkeyIDType[KSCM_SIZE];
	if (kvk == NULL || ksc == NULL || kvkm == NULL || kscm == NULL)
	{
		if (kvk) delete [] kvk;
		if (ksc) delete [] ksc;
		if (kvkm) delete [] kvkm;
		if (kscm) delete [] kscm;
		kvk = ksc = NULL;
		kvkm = kscm = NULL;
		return 0;
	}

	// Very important to initialize in this case.  Don't use sizeof(kvk/ksc) because kvk and ksc
	// are just pointers in the WinNT/2k/XP version, not static arrays as in Win9x version.
	// Also: Not using a constructor and initializer list for this because there may be times
	// in future versions of this (e.g. hook is installed and deinstalled more than once
	// during the lifetime of the process) when we want to explicitly initialize the data
	// in the struct.  Since I don't think you can explicitly call an object's constructor
	// (and even if you could, you'd have loop through every array element), it seems best
	// to do it this way.  This also initializes any pointers to NULL rather than a
	// bitwise-zero for compatibility with any conceivable hardware:
	ZeroMemory(kvk, KVK_SIZE * sizeof(key_type));
	ZeroMemory(ksc, KSC_SIZE * sizeof(key_type));
	int i;
	for (i = 0; i < KVK_SIZE; ++i)
		kvk[i].pForceToggle = NULL;
	for (i = 0; i < KSC_SIZE; ++i)
		ksc[i].pForceToggle = NULL;

	// This attribute is exists for performance reasons (avoids a function call in the hook
	// procedure to determine this value):
	kvk[VK_CONTROL].as_modifiersLR = MOD_LCONTROL | MOD_RCONTROL;
	kvk[VK_LCONTROL].as_modifiersLR = MOD_LCONTROL;
	kvk[VK_RCONTROL].as_modifiersLR = MOD_RCONTROL;
	kvk[VK_MENU].as_modifiersLR = MOD_LALT | MOD_RALT;
	kvk[VK_LMENU].as_modifiersLR = MOD_LALT;
	kvk[VK_RMENU].as_modifiersLR = MOD_RALT;
	kvk[VK_SHIFT].as_modifiersLR = MOD_LSHIFT | MOD_RSHIFT;
	kvk[VK_LSHIFT].as_modifiersLR = MOD_LSHIFT;
	kvk[VK_RSHIFT].as_modifiersLR = MOD_RSHIFT;
	kvk[VK_LWIN].as_modifiersLR = MOD_LWIN;
	kvk[VK_RWIN].as_modifiersLR = MOD_RWIN;

/*
	// Used to assist the performance and readability of the hook procedure:
	kvk[VK_SCROLL].toggleable_vk = VK_SCROLL;
	kvk[VK_CAPITAL].toggleable_vk = VK_CAPITAL;
	kvk[VK_NUMLOCK].toggleable_vk = VK_NUMLOCK;
*/
	// Use the address rather than the value, so that if the global var's value
	// changes during runtime, ours will too:
	kvk[VK_SCROLL].pForceToggle = &g_ForceScrollLock;
	kvk[VK_CAPITAL].pForceToggle = &g_ForceCapsLock;
	kvk[VK_NUMLOCK].pForceToggle = &g_ForceNumLock;

	// This is a bit iffy because it's far from certain that these particular scan codes
	// are really modifier keys on anything but a standard English keyboard.  However,
	// at the very least the Win9x version must rely on something like this because a
	// low-level hook can't be used under Win9x, and a high-level hook doesn't receive
	// the left/right VKs at all (so the scan code must be used to tell them apart).
	// However: it might be possible under Win9x to use MapVirtualKey() or some similar
	// function to verify, at runtime, that the expected scan codes really do map to the
	// expected VK.  If not, perhaps MapVirtualKey() or such can be used to search through
	// every scan code to find out which map to VKs that are modifiers.  Any such keys
	// found can then be initialized similar to below:
	ksc[SC_LCONTROL].as_modifiersLR = MOD_LCONTROL;
	ksc[SC_RCONTROL].as_modifiersLR = MOD_RCONTROL;
	ksc[SC_LALT].as_modifiersLR = MOD_LALT;
	ksc[SC_RALT].as_modifiersLR = MOD_RALT;
	ksc[SC_LSHIFT].as_modifiersLR = MOD_LSHIFT;
	ksc[SC_RSHIFT].as_modifiersLR = MOD_RSHIFT;
	ksc[SC_LWIN].as_modifiersLR = MOD_LWIN;
	ksc[SC_RWIN].as_modifiersLR = MOD_RWIN;

	// These have to be initialized with with element value INVALID.
	// Don't use FillMemory because the array elements are too big (bigger than bytes):
	for (i = 0; i < KVKM_SIZE; ++i)
		kvkm[i] = HOTKEY_ID_INVALID;
	for (i = 0; i < KSCM_SIZE; ++i)
		kscm[i] = HOTKEY_ID_INVALID;

	// Indicate here which scan codes should override their virtual keys:
	for (i = 0; i < g_key_to_sc_count; ++i)
		if (g_key_to_sc[i].sc > 0 && g_key_to_sc[i].sc <= SC_MAX)
			ksc[g_key_to_sc[i].sc].sc_takes_precedence = true;

	hk_sorted_type hk_sorted[MAX_HOTKEYS];
	ZeroMemory(hk_sorted, sizeof(hk_sorted));
	int hk_sorted_count = 0;
	key_type *pThisKey = NULL;
	for (i = 0; i < aHK_count; ++i)
	{
		if (aHK[i]->mType != HK_KEYBD_HOOK && aHK[i]->mType != HK_MOUSE_HOOK)
			continue;

		// Rule out the possibility of obnoxious values right away, preventing array-out-of bounds, etc.:
		if ((!aHK[i]->mVK && !aHK[i]->mSC) || aHK[i]->mVK > VK_MAX || aHK[i]->mSC > SC_MAX)
			continue;

		if (!aHK[i]->mVK)
			// scan codes don't need something like the switch stmt below because they can't be neutral.
			// In other words, there's no scan code equivalent for something like VK_CONTROL.
			// In addition, SC_LCONTROL, for example, doesn't also need to change the kvk array
			// for VK_LCONTROL because the hook knows to give the scan code precedence, and thus
			// look it up only in the ksc array in that case.
			pThisKey = ksc + aHK[i]->mSC;
		else
		{
			pThisKey = kvk + aHK[i]->mVK;
			// Keys that have a neutral as well as a left/right counterpart must be
			// fully initialized since the hook can receive the left, the right, or
			// the neutral (neutral only if another app calls KeyEvent(), probably).
			// There are several other switch stmts in this function like the below
			// that serve a similar purpose.  The alternative to doing all these
			// switch stmts is to always translate left/right vk's (whose sc's don't
			// take precedence) in the KeyboardProc() itself.  But that would add
			// the overhead of a switch stmt to *every* keypress ever made on the
			// system, so it seems better to set up everything correctly here since
			// this init section is only done once.
			// Note: These switch stmts probably aren't needed under Win9x since I think
			// it might be impossible for them to receive something like VK_LCONTROL,
			// except *possibly* if keybd_event() is explicitly called with VK_LCONTROL
			// and (might want to verify that -- if true, might want to keep the switches
			// even for Win9x for safety and in case Win9x ever gets overhauled and
			// improved in some future era, or in case Win9x is running in an emulator
			// that expands its capabilities.
			switch (aHK[i]->mVK)
			{
			case VK_MENU:
				// It's not strictly necessary to init all of these, since the
				// hook currently never handles VK_RMENU, for example, by its
				// vk (it uses sc instead).  But it's safest to do all of them
				// in case future changes ever ruin that assumption:
				kvk[VK_LMENU].used_as_suffix = kvk[VK_RMENU].used_as_suffix
					= ksc[SC_LALT].used_as_suffix = ksc[SC_RALT].used_as_suffix = true;
				break;
			case VK_SHIFT:
				// The neutral key itself is also set to be a suffix further below.
				kvk[VK_LSHIFT].used_as_suffix = kvk[VK_RSHIFT].used_as_suffix
					= ksc[SC_LSHIFT].used_as_suffix = ksc[SC_RSHIFT].used_as_suffix = true;
				break;
			case VK_CONTROL:
				kvk[VK_LCONTROL].used_as_suffix = kvk[VK_RCONTROL].used_as_suffix
					= ksc[SC_LCONTROL].used_as_suffix = ksc[SC_RCONTROL].used_as_suffix = true;
				break;
			// Later might want to add cases for VK_LCONTROL and such, but for right now,
			// these keys should never come up since they're done by scan code?
			}
		}

		pThisKey->used_as_suffix = true;

		HotkeyIDType hotkey_id_with_flags = aHK[i]->mID;
		if (!aHK[i]->mDoSuppress)
		{
			hotkey_id_with_flags |= HOTKEY_NO_SUPPRESS;
			// Due to the fact that the hook does handle things as hotkeys, but rather at
			// a lower level (prefixes and suffixes), there's no easy way to support toggling
			// suppression for individual hotkeys (perhaps the simplest way to do so would be
			// to create a new array of structs to contain the list of which hotkey_id's are
			// special/non-suppressed).  So for now, once a suffix key has had its suppression
			// turned off, it stays off.  But currently, this is inconsequential I think, since
			// only the naked (unmodified) key, when used as a hotkey, supports non-suppression.
			// UPDATE: The above limitation now applies only to mouse buttons, since the
			// non-suppression of keyboard keys is handled via the HOTKEY_NO_SUPPRESS bit in
			// the hotkey_id (mouse events can't be easily handled this way since the hook
			// would have to generate substitute mouse events, which may be a lot of work to
			// code):
			if (aHK[i]->mType == HK_MOUSE_HOOK)
				pThisKey->no_mouse_suppress = true;
		}
		// else leave the bit set to zero so that the key will be suppressed (most hotkeys are like this).

		// If this is a naked (unmodified) modifier key, make it a prefix if it ever modifies any
		// other hotkey.  This processing might be later combined with the hotkeys activation function
		// to eliminate redundancy / improve efficiency, but then that function would probably need to
		// init everything else here as well:
		if (pThisKey->as_modifiersLR && !aHK[i]->mModifiers && !aHK[i]->mModifiersLR
			&& !aHK[i]->mModifierVK && !aHK[i]->mModifierSC)
			SetModifierAsPrefix(aHK[i]->mVK, aHK[i]->mSC);

		if (aHK[i]->mModifierVK)
		{
			if (kvk[aHK[i]->mModifierVK].as_modifiersLR)
				// The hotkey's ModifierVK is itself a modifier.
				SetModifierAsPrefix(aHK[i]->mModifierVK, 0, true);
			else
				kvk[aHK[i]->mModifierVK].used_as_prefix = true;

			if (pThisKey->nModifierVK < MAX_MODIFIER_VKS_PER_SUFFIX)  // else currently no error-reporting.
			{
				pThisKey->ModifierVK[pThisKey->nModifierVK].vk = aHK[i]->mModifierVK;
				if (aHK[i]->mHookAction)
					pThisKey->ModifierVK[pThisKey->nModifierVK].id_with_flags = aHK[i]->mHookAction;
				else
					pThisKey->ModifierVK[pThisKey->nModifierVK].id_with_flags = hotkey_id_with_flags;
				++pThisKey->nModifierVK;
				continue;
			}
		}
		else
		{
			if (aHK[i]->mModifierSC)
			{
				if (kvk[aHK[i]->mModifierSC].as_modifiersLR)
					// The hotkey's ModifierSC is itself a modifier.
					SetModifierAsPrefix(0, aHK[i]->mModifierSC, true);
				else
					ksc[aHK[i]->mModifierSC].used_as_prefix = true;
				if (pThisKey->nModifierSC < MAX_MODIFIER_SCS_PER_SUFFIX)  // else currently no error-reporting.
				{
					pThisKey->ModifierSC[pThisKey->nModifierSC].sc = aHK[i]->mModifierSC;
					if (aHK[i]->mHookAction)
						pThisKey->ModifierSC[pThisKey->nModifierSC].id_with_flags = aHK[i]->mHookAction;
					else
						pThisKey->ModifierSC[pThisKey->nModifierSC].id_with_flags = hotkey_id_with_flags;
					++pThisKey->nModifierSC;
					continue;
				}
			}
		}

		// At this point, since the above didn't "continue", this hotkey is one without a ModifierVK/SC.
		// Put it into a temporary array, which will be later sorted.

		hk_sorted[hk_sorted_count].id_with_flags = aHK[i]->mHookAction ? aHK[i]->mHookAction : hotkey_id_with_flags;
		hk_sorted[hk_sorted_count].vk = aHK[i]->mVK;
		hk_sorted[hk_sorted_count].sc = aHK[i]->mSC;
		hk_sorted[hk_sorted_count].modifiers = aHK[i]->mModifiers;
		hk_sorted[hk_sorted_count].modifiersLR = aHK[i]->mModifiersLR;
		hk_sorted[hk_sorted_count].AllowExtraModifiers = aHK[i]->mAllowExtraModifiers;
		++hk_sorted_count;
	}

	if (hk_sorted_count)
	{
		// It's necessary to get them into this order to avoid problems that would be caused by
		// AllowExtraModifiers:
		qsort((void *)hk_sorted, hk_sorted_count, sizeof(hk_sorted_type), sort_most_general_before_least);

		// For each hotkey without a ModifierVK/SC (which override normal modifiers), expand its modifiers and
		// modifiersLR into its column in the kvkm or kscm arrays.

		mod_type modifiers, i_modifiers_merged;
		int modifiersLR;  // Don't make this modLR_type to avoid integer overflow, since it's a loop-counter.
		for (i = 0; i < hk_sorted_count; ++i)
		{
			i_modifiers_merged = hk_sorted[i].modifiers;
			if (hk_sorted[i].modifiersLR)
				i_modifiers_merged |= ConvertModifiersLR(hk_sorted[i].modifiersLR);

//if (hk_sorted[i].vk == VK_MENU) PostMessage(g_hWnd, AHK_HOOK_TEST_MSG, hk_sorted[i].modifiers, hk_sorted[i].modifiersLR);
			for (modifiersLR = 0; modifiersLR <= MODLR_MAX; ++modifiersLR)  // For each possible LR value.
			{
				modifiers = ConvertModifiersLR(modifiersLR);
				if (hk_sorted[i].AllowExtraModifiers)
				{
					// True if modifiersLR is a superset of i's modifier value.  In other words,
					// modifiersLR has the minimum required keys but also has some
					// extraneous keys, which are allowed in this case:
					if (i_modifiers_merged != (modifiers & i_modifiers_merged))
						continue;
				}
				else
					if (i_modifiers_merged != modifiers)
						continue;

				// In addition to the above, modifiersLR must also have the *specific* left or right keys
				// found in i's modifiersLR.  In other words, i's modifiersLR must be a perfect subset
				// of modifiersLR:
				if (hk_sorted[i].modifiersLR) // make sure that any more specific left/rights are also present.
					if (hk_sorted[i].modifiersLR != (modifiersLR & hk_sorted[i].modifiersLR))
						continue;

				// If above didn't "continue", modifiersLR is a valid hotkey combination so set it as such:
				if (!hk_sorted[i].vk)
					// scan codes don't need the switch() stmt below because, for example,
					// the hook knows to look up left-control by only SC_LCONTROL,
					// not VK_LCONTROL.
					Kscm(modifiersLR, hk_sorted[i].sc) = hk_sorted[i].id_with_flags;
				else
				{
					Kvkm(modifiersLR, hk_sorted[i].vk) = hk_sorted[i].id_with_flags;
					switch (hk_sorted[i].vk)
					{
					case VK_MENU:
						Kvkm(modifiersLR, VK_LMENU) = Kvkm(modifiersLR, VK_RMENU)
							= Kscm(modifiersLR, SC_LALT) = Kscm(modifiersLR, SC_RALT)
							= hk_sorted[i].id_with_flags;
						break;
					case VK_LMENU: // In case the program is ever changed to support these VKs directly.
						Kvkm(modifiersLR, VK_LMENU) = Kscm(modifiersLR, SC_LALT) = hk_sorted[i].id_with_flags;
						break;
					case VK_RMENU:
						Kvkm(modifiersLR, VK_RMENU) = Kscm(modifiersLR, SC_RALT) = hk_sorted[i].id_with_flags;
						break;
					case VK_SHIFT:
						Kvkm(modifiersLR, VK_LSHIFT) = Kvkm(modifiersLR, VK_RSHIFT)
							= Kscm(modifiersLR, SC_LSHIFT) = Kscm(modifiersLR, SC_RSHIFT)
							= hk_sorted[i].id_with_flags;
						break;
					case VK_LSHIFT:
						Kvkm(modifiersLR, VK_LSHIFT) = Kscm(modifiersLR, SC_LSHIFT) = hk_sorted[i].id_with_flags;
						break;
					case VK_RSHIFT:
						Kvkm(modifiersLR, VK_RSHIFT) = Kscm(modifiersLR, SC_RSHIFT) = hk_sorted[i].id_with_flags;
						break;
					case VK_CONTROL:
						Kvkm(modifiersLR, VK_LCONTROL) = Kvkm(modifiersLR, VK_RCONTROL)
							= Kscm(modifiersLR, SC_LCONTROL) = Kscm(modifiersLR, SC_RCONTROL)
							= hk_sorted[i].id_with_flags;
						break;
					case VK_LCONTROL:
						Kvkm(modifiersLR, VK_LCONTROL) = Kscm(modifiersLR, SC_LCONTROL) = hk_sorted[i].id_with_flags;
						break;
					case VK_RCONTROL:
						Kvkm(modifiersLR, VK_RCONTROL) = Kscm(modifiersLR, SC_RCONTROL) = hk_sorted[i].id_with_flags;
						break;
					}
				}
			}
		}
	}

	// It seems best to reset these whenever either hook is activated, since there's
	// currently no way to know if the function has been called before.  Currently,
	// it's only called once (either by first use of "SetNumLock(etc.), AlwaysOn" or
	// by the activation of the entire hotkey set (if the set requires the hook)
	// so this is the first time:
	ZeroMemory(KeyLog, sizeof(KeyLog));
	ZeroMemory(g_PhysicalKeyState, sizeof(g_PhysicalKeyState));
	pPrefixKey = NULL;
	g_modifiersLR_logical = g_modifiersLR_physical = g_modifiersLR_get = 0;
	disguise_next_lwin_up = disguise_next_rwin_up = alt_tab_menu_is_visible = false;
	// Even if OS is Win9x, try LL hooks anyway.  This will probably fail on WinNT if it doesn't have SP3+
	if (!g_hhkLowLevelKeybd && (aHooksToActivate & HOOK_KEYBD))
		if (g_hhkLowLevelKeybd = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeybdProc, g_hInstance, 0))
			hooks_currently_active |= HOOK_KEYBD;
	if (!g_hhkLowLevelMouse && (aHooksToActivate & HOOK_MOUSE))
		if (g_hhkLowLevelMouse = SetWindowsHookEx(WH_MOUSE_LL, LowLevelMouseProc, g_hInstance, 0))
			hooks_currently_active |= HOOK_MOUSE;

	return hooks_currently_active;
}



void HookTerm()
{
	if (g_hhkLowLevelKeybd)
		if (UnhookWindowsHookEx(g_hhkLowLevelKeybd))
			g_hhkLowLevelKeybd = NULL;
	if (g_hhkLowLevelMouse)
		if (UnhookWindowsHookEx(g_hhkLowLevelMouse))
			g_hhkLowLevelMouse = NULL;
	if (kvk) delete [] kvk;
	if (ksc) delete [] ksc;
	if (kvkm) delete [] kvkm;
	if (kscm) delete [] kscm;
	kvk = ksc = NULL;
	kvkm = kscm = NULL;
}



char *GetHookStatus(char *aBuf, size_t aBufSize)
{
	if (!aBuf || !aBufSize) return aBuf;

	char KeyLogText[2048], KeyName[128], KeyNameOS[128]
		, LRgText[128], LRhText[128], LRpText[128], IgnoreText[1024];
	int item, i, j;
	*KeyLogText = *IgnoreText = '\0';

	if (!g_hhkLowLevelKeybd && !g_hhkLowLevelMouse)
		strlcpy(KeyLogText, "\r\n"
			"Note: The KeyLog itself is not shown because neither\r\n"
			"the keyboard nor the mouse hook is installed.\r\n"
			"You can force them to be installed by adding either or\r\n"
			"both of the following lines to this script:\r\n"
			"#InstallKeybdHook\r\n"
			"#InstallMouseHook\r\n", sizeof(KeyLogText));
	else
	{
		// Start at the oldest key, which is KeyLogNext:
		for (item = KeyLogNext, i = 0; i < MAX_LOGGED_KEYS; ++i, ++item)
		{
			if (item >= MAX_LOGGED_KEYS)
				item = 0;
			switch(KeyLog[item].vk)
			{
			case VK_CONTROL:
			case VK_MENU:
			case VK_SHIFT:
			case VK_LWIN:
			case VK_RWIN:
				*KeyName = '\0';  // Let GetKeyNameText() resolve it instead.
				break;
			default:
				for (j = 0; j < g_key_to_vk_count; ++j)
					if (g_key_to_vk[j].vk == KeyLog[item].vk)
						break;
				if (j < g_key_to_vk_count)
					strcpy(KeyName, g_key_to_vk[j].key_name);
				else
					if (isprint(KeyLog[item].vk))
					{
						KeyName[0] = KeyLog[item].vk;
						KeyName[1] = '\0';
					}
					else
						*KeyName = '\0';
			}
			*KeyNameOS = '\0';
			if (KeyLog[item].sc)
				// Use 0x02000000 to tell it that we want it to give left/right specific info, lctrl/rctrl etc.
				if (!GetKeyNameText((long)(KeyLog[item].sc) << 16
					, KeyNameOS, sizeof(KeyNameOS)/sizeof(TCHAR)))
					snprintf(KeyNameOS, sizeof(KeyNameOS), "lParam=%08X", (LPARAM)(KeyLog[item].sc) << 16);
			// Minimize the length of the text in this stmt because if it's too long,
			// the MessageBox dialog will trucate it and not display the most important
			// (most recent) keystrokes:
			snprintfcat(KeyLogText, sizeof(KeyLogText), "\r\nvk%02X sc%03X %c %c %s %s"
				, KeyLog[item].vk, KeyLog[item].sc
				, KeyLog[item].key_up ? 'u' : 'd'
				// It can't be both ignored and suppressed, so display only one:
				, KeyLog[item].event_type
				, KeyName, stricmp(KeyNameOS, KeyName) ? KeyNameOS : ""
				);
		}
	}
	snprintf(aBuf, aBufSize,
		"Modifiers (Hook's Logical) = %s\r\n"
		"Modifiers (Hook's Physical) = %s\r\n" // Font isn't fixed-width, so don't bother trying to line them up.
		"Modifiers (GetKeyState()) = %s (old, from last call to GetModifierLRState())\r\n"
		"Prefix Key Down=%s\r\n"
		"%s"
		, ModifiersLRToText(g_modifiersLR_logical, LRhText)
		, ModifiersLRToText(g_modifiersLR_physical, LRpText)
		, ModifiersLRToText(g_modifiersLR_get, LRgText)
		, pPrefixKey == NULL ? "No" : "Yes"
		// This can't work because the hotkey that triggers the info-dialog will always make it true:
		//, pPrefixKey == NULL ? "" : (pPrefixKey->was_just_used ? " (yes)" : " (no)")
		, KeyLogText);
	return aBuf;
}
