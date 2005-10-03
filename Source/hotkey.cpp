/*
AutoHotkey

Copyright 2003-2005 Chris Mallett (support@autohotkey.com)

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
#include "hotkey.h"
#include "globaldata.h"  // For g_os and other global vars.
#include "window.h" // For MsgBox()
//#include "application.h" // For ExitApp()

// Initialize static members:
bool Hotkey::sHotkeysAreLocked = false;
HookType Hotkey::sWhichHookNeeded = 0;
HookType Hotkey::sWhichHookAlways = 0;
DWORD Hotkey::sTimePrev = {0};
DWORD Hotkey::sTimeNow = {0};
Hotkey *Hotkey::shk[MAX_HOTKEYS] = {NULL};
HotkeyIDType Hotkey::sNextID = 0;
const HotkeyIDType &Hotkey::sHotkeyCount = Hotkey::sNextID;
bool Hotkey::sJoystickHasHotkeys[MAX_JOYSTICKS] = {false};
DWORD Hotkey::sJoyHotkeyCount = 0;



void Hotkey::AllActivate()
// This function can also be called to install the keyboard hook if the state
// of g_ForceNumLock and such have changed, even if the hotkeys are already
// active.
{
	// Do this part only if it hasn't been done before (as indicated by sHotkeysAreLocked)
	// because it's not reviewed/designed to be run more than once.  UPDATE (for dynamic
	// hotkeys): I think the gist of the likely problems with running this part more than
	// once is that some hotkeys were already promoted (the first time this ran) to be hook
	// hotkeys based on their interdependencies with other hotkeys.  If those
	// interdependencies change due to dynamic changes to the set of hotkeys, some hotkeys
	// might stay promoted when they should in fact be demoted back to being registered,
	// which will result in the hook being installed to implement those keys when it
	// technically doesn't need to be.  However, this is likely extremely rare, and fairly
	// inconsequential even when it does happen.  This could perhaps be resolved by
	// storing the original mType with each hotkey so that those types can be restored
	// by the caller, if desired, before calling us.

	// v1.0.37.05: A prefix key such as "a" in "a & b" should cause any use of "a" as a suffix
	// (such as ^!a) also to be a hook hotkey.  Otherwise, the ^!a hotkey won't fire because the
	// hook prevents the OS's hotkey monitor from seeing that the hotkey was pressed.  NOTE:
	// This is done only for vitual keys because prefix keys that are done by scan code (mModifierSC)
	// should already be hook hotkeys when used as suffix keys (there may be a few unusual exceptions,
	// but they seem too rare to justify the extra code size).
	bool vk_is_prefix[VK_ARRAY_COUNT] = {false};
	int j;
	for (j = 0; j < sHotkeyCount; ++j)
	{
		Hotkey &hot = *shk[j]; // For performance and convenience.
		if (hot.mModifierVK && hot.mEnabled && (!g_IsSuspended || hot.IsExemptFromSuspend())) // Ordered for short-circuit performance.
			vk_is_prefix[hot.mModifierVK] = true;
	}

	bool is_neutral;
	modLR_type modifiersLR;
	for (int i = 0; i < sHotkeyCount; ++i)
	{
		Hotkey &hot = *shk[i]; // For performance and convenience.

		if (hot.mKeyUp && hot.mVK) // No need to do the below for mSC hotkeys since their down hotkeys would already be handled by the hook.
		{
			// For each key-up hotkey, search for any its counterpart that's a down-hotkey (if any).
			// Such a hotkey should also be handled by the hook because if registered, such as
			// "#5" (reg) and "#5 up" (hook), the hook would suppress the down event because it
			// is unaware that down-hotkey exists (it's suppressed to prevent the key from being
			// stuck in a logically down state).
			for (j = 0; j < sHotkeyCount; ++j)
			{
				// No need to check the following because they are already hook hotkeys:
				// mModifierVK/SC
				// mAllowExtraModifiers
				// mNoSuppress
				// In addition, there is no need to check shk[j]->mKeyUp because that can't be
				// true if it's mType is HK_NORMAL:
				if (shk[j]->mType == HK_NORMAL && shk[j]->mVK == hot.mVK && shk[j]->mModifiersConsolidatedLR == hot.mModifiersConsolidatedLR)
				{
					// v1.0.39: Unregister in case previous iteration registered it.
					// This could be avoided by putting this loop outside the main loop, but that would
					// require two passes through the hotkey list, hurting performance and increasing code size.
					shk[j]->Unregister();
					shk[j]->mType = HK_KEYBD_HOOK;
				}
			}
		}

		// For simplicity, don't try to undo keys that are already considered to be
		// handled by the hook, since it's not easy to know if they were set that
		// way using "#UseHook, on" or really qualified some other way.
		// Instead, just remove any modifiers that are obviously redundant from all
		// keys (do them all due to cases where RegisterHotkey() fails and the key
		// is then auto-enabled via the hook).  No attempt is currently made to
		// correct a silly hotkey such as "lwin & lwin".  In addition, weird hotkeys
		// such as <^Control and ^LControl are not currently validated and might
		// yield unpredictable results:
		if (modifiersLR = KeyToModifiersLR(hot.mVK, hot.mSC, &is_neutral))
			// This hotkey's action-key is itself a modifier, so ensure that it's not defined
			// to modify itself.  Other sections might rely on us doing this:
			if (is_neutral)
				// Since the action-key is a neutral modifier (not left or right specific),
				// turn off any neutral modifiers that may be on:
				hot.mModifiers &= ~ConvertModifiersLR(modifiersLR);
			else
				hot.mModifiersLR &= ~modifiersLR;

		// HK_MOUSE_HOOK type, and most HK_KEYBD types, are handled by the hotkey constructor.
		// What we do here is change the type of any normal or undetermined key if there are other
		// keys that overlap with it (i.e. because only now are all these keys available for checking).
		if (hot.mType == HK_UNDETERMINED || hot.mType == HK_NORMAL)
		{
			// The idea here is to avoid the use of the keyboard hook if at all possible (since
			// it may reduce system performance slightly).  With that in mind, rather than just
			// forcing NumpadEnter and Enter to be entirely separate keys, both handled by the
			// hook, we allow mod+Enter to take over both keys if their is no mod+NumPadEnter key
			// configured with identical modifiers.  UPDATE: I don't think I like this method because
			// it will produce undocumented behavior in some cases (e.g. if user has ^NumpadEnter as
			// a hotkey, ^Enter will also trigger the hotkey, which is not what would be expected).
			// Therefore, I'm changing it now to have all dual-state keys handled by the hook so that
			// the counterpart key will never trigger an unexpected firing:
			if (vk_to_sc(hot.mVK, true) // This virtual key corresponds to two scan codes...
				|| vk_is_prefix[hot.mVK]) // ... or it's a suffix that is also used as a prefix (allows ^!a to work without $ when "a & b" is a hotkey).
			{
				// v1.0.38.02: The line below now makes an explicit VK hotkey such as "VK24::" (which is VK_HOME)
				// to be handled via RegisterHotkey() vs. the hook.  Someone asked for this ability, but even if it
				// weren't for that it seems more correct to recognize an explicitly-specified VK as a "neutral VK"
				// (i.e. one that fires for both scan codes if the VK has two scan codes). The user can always
				// specify "SCnnn::" as a hotkey to avoid this fire-on-both-scan-codes behavior.
				hot.mType = (g_os.IsWin9x() || hot.mVK_WasSpecifiedByNumber) ? HK_NORMAL : HK_KEYBD_HOOK;
				// Older comments (some might be obsolete in light of the v1.0.38.02 change above):
				// If the hook must be used, we'll mark it as needing the hook so that
				// other reporting (e.g. ListHotkeys) can easily tell which keys won't work
				// on Win9x.  The first condition: e.g. a naked NumpadEnd or NumpadEnter shouldn't
				// be allowed to be Registered() because it would cause the hotkey to also fire on
				// END or ENTER (the non-Numpad versions of these keys).  UPDATE: But it seems
				// best to allow this on Win9x because it's more flexible to do so (i.e. some
				// people might have a use for it):
				//if (!hot.mModifiers)
				//	hot.mType = HK_KEYBD_HOOK;
				//
				// Second condition (now disabled): Since both keys (e.g. NumpadEnd and End) are
				// configured as hotkeys with the same modifiers, only one of them can be registered.
				// It's probably best to allow one of them to be registered, arbitrarily, so that some
				// functionality is offered.  That's why this is disabled:
				// This hotkey's vk has a counterpart key with the same vk.  Check if either of those
				// two keys exists as a "scan code" hotkey.  If so, This hotkey must be handled
				// by the hook to prevent it from firing for both scan codes.  modifiersLR should be
				// zero here because otherwise type would have already been set to HK_KEYBD_HOOK:
				//if (FindHotkeyBySC(vk_to_sc(hot.mVK), hot.mModifiers, hot.mModifiersLR) != HOTKEY_ID_INVALID))
				//	hot.mType = HK_KEYBD_HOOK;
			}

			// Fall back to default checks if more specific ones above didn't set it to use the hook:
			if (hot.mType != HK_KEYBD_HOOK)
			{
				// Keys modified by CTRL/SHIFT/ALT/WIN can always be registered normally because these
				// modifiers are never used (are overridden) when that key is used as a ModifierVK
				// for another key.  Example: if key1 is a ModifierVK for another key, ^key1
				// (or any other modified versions of key1) can be registered as a hotkey because
				// that doesn't affect the hook's ability to use key1 as a prefix:
				if (hot.mModifiers)
					hot.mType = HK_NORMAL;
				else
				{
					if ((hot.mVK == VK_LWIN || hot.mVK == VK_RWIN))  // "!hot.mModifiers" already true
						// To prevent the start menu from appearing for a naked LWIN or RWIN, must
						// handle this key with the hook (the presence of a normal modifier makes
						// this unnecessary, at least under WinXP, because the start menu is
						// never invoked when a modifier key is held down with lwin/rwin).
						// But make it NORMAL on Win9x since the hook isn't yet supported.
						// At least that way there's a chance some people might find it useful,
						// perhaps by doing their own workaround for the start menu appearing:
						hot.mType = g_os.IsWin9x() ? HK_NORMAL : HK_KEYBD_HOOK;
					else
					{
						// If this hotkey is an unmodified modifier (e.g. control = calc.exe) and
						// there are any other hotkeys that rely specifically on this modifier,
						// have the hook handle this hotkey so that it will only fire on key-up
						// rather than key-down.  Note: cases where this key's modifiersLR or
						// ModifierVK/SC are non-zero -- as well as hotkeys that use sc vs. vk
						// -- have already been set to use the keybd hook, so don't need to be
						// handled here.  UPDATE: All the following cases have been already set
						// to be HK_KEYBD_HOOK:
						// - left/right ctrl/alt/shift (since RegisterHotkey() doesn't support them).
						// - Any key with a ModifierVK/SC
						// - The naked lwin or rwin key (due to the check above)
						// Therefore, the only case left to be detected by this next line is the
						// one in which the user configures the naked neutral key VK_SHIFT,
						// VK_MENU, or VK_CONTROL.  As a safety precaution, always handle those
						// neutral keys with the hook so that their action will only fire
						// when the key is released (thus allowing each key to be used for its
						// normal modifying function):
						if (hot.mVK == VK_SHIFT || hot.mVK == VK_MENU || hot.mVK == VK_CONTROL)
							// In addition, the following are already known to be true or we wouldn't
							// have gotten this far:
							// !hot.mModifiers && !hot.mModifiersLR
							// !hot.mModifierVK && !hot.mModifierSC
							// If it's Win9x, take a stab at making it a normal registered hotkey
							// in case it's of some use to someone to allow that:
							hot.mType = g_os.IsWin9x() ? HK_NORMAL : HK_KEYBD_HOOK;
						else
							// Check if this key is used as the modifier (prefix) for any other key.  If it is,
							// the keyboard hook must handle this key also because otherwise the key-down event
							// would trigger the registered hotkey immediately, rather than waiting to see if
							// this key is be held down merely to modify some other key.
							hot.mType = !g_os.IsWin9x() && FindHotkeyWithThisModifier(hot.mVK, hot.mSC)
								!= HOTKEY_ID_INVALID ? HK_KEYBD_HOOK : HK_NORMAL;
					}
				}
				if (hot.mVK == VK_APPS)
					// Override anything set above:
					// For now, always use the hook to handle hotkeys that use Appskey as a suffix.
					// This is because registering such keys with RegisterHotkey() will fail to suppress
					// (hide) the key-up events from the system, and the key-up for Apps key, at least in
					// apps like Explorer, is a special event that results in the context menu appearing
					// (though most other apps seem to use the key-down event rather than the key-up,
					// so they would probably work okay).  Note: Of possible future use is the fact that
					// if the Alt key is held down before pressing Appskey, it's native function does
					// not occur.  This may be similar to the fact that LWIN and RWIN don't cause the
					// start menu to appear if a shift key is held down.  For Win9x, take a stab at
					// registering it in case its limited capability is useful to someone:
					hot.mType = g_os.IsWin9x() ? HK_NORMAL : HK_KEYBD_HOOK;
			}
		}

		// Check if this mouse hotkey also requires the keyboard hook (e.g. #LButton).
		// Some mouse hotkeys, such as those with normal modifiers, don't require it
		// since the mouse hook has logic to handle that situation.  But those that
		// are composite hotkeys such as "RButton & Space" or "Space & RButton" need
		// the keyboard hook:
		if (hot.mType == HK_MOUSE_HOOK && (
			hot.mModifierSC || hot.mSC // i.e. since it's an SC, the modifying key isn't a mouse button.
			|| hot.mHookAction // v1.0.25.05: At least some alt-tab actions require the keyboard hook. For example, a script consisting only of "MButton::AltTabAndMenu" would not work properly otherwise.
			// v1.0.25.05: The line below was added to prevent the Start Menu from appearing, which
			// requires the keyboard hook. ALT hotkeys don't need it because the mouse hook sends
			// a CTRL keystroke to disguise them, a trick that is unfortunately not reliable for
			// when it happens while the while key is down (though it does disguise a Win-up).
			|| ((hot.mModifiersConsolidatedLR & (MOD_LWIN|MOD_RWIN)) && !(hot.mModifiersConsolidatedLR & (MOD_LALT|MOD_RALT)))
			// For v1.0.30, above has been expanded to include Win+Shift and Win+Control modifiers.
			|| (hot.mVK && !IsMouseVK(hot.mVK)) // e.g. "RButton & Space"
			|| (hot.mModifierVK && !IsMouseVK(hot.mModifierVK)))   ) // e.g. "Space & RButton"
			hot.mType = HK_BOTH_HOOKS;  // Needed by ChangeHookState().
			// For the above, the following types of mouse hotkeys do not need the keyboard hook:
			// 1) mAllowExtraModifiers: Already handled since the mouse hook fetches the modifier state
			//    manually when the keyboard hook isn't installed.
			// 2) mModifiersConsolidatedLR (i.e. the mouse button is modified by a normal modifier
			//    such as CTRL): Same reason as #1.
			// 3) As a subset of #2, mouse hotkeys that use WIN as a modifier will not have the
			//    Start Menu suppressed unless the keyboard hook is installed.  It's debatable,
			//    but that seems a small price to pay (esp. given how rare it is just to have
			//    the mouse hook with no keyboard hook) to avoid the overhead of the keyboard hook.
		
		// If the hotkey is normal, try to register it.  If the register fails, use the hook to try
		// to override any other script or program that might have it registered (as documented):
		if (hot.mType == HK_NORMAL && (!g_IsSuspended || hot.IsExemptFromSuspend()) && !hot.Register())
			hot.mType = HK_KEYBD_HOOK;
			// The old Win9x warning dialog was removed to reduce code size (since usage of
			// Win9x is becoming very rare).

		// If this is a hook hotkey and the OS is Win9x, the old warning has been removed to
		// reduce code size (since usage of Win9x is becoming very rare).  Older comment:
		// Since it's flagged as a hook in spite of the fact that the OS is Win9x, it means
		// that some previous logic determined that it's not even worth trying to register
		// it because it's just plain not supported.
		switch (hot.mType) // It doesn't matter if the OS is Win9x because in that case, other sections just ignore hook hotkeys.
		{
		case HK_KEYBD_HOOK: sWhichHookNeeded |= HOOK_KEYBD; break;
		case HK_MOUSE_HOOK: sWhichHookNeeded |= HOOK_MOUSE; break;
		case HK_BOTH_HOOKS: sWhichHookNeeded |= HOOK_KEYBD|HOOK_MOUSE; break;
		}
	} // for()

	// Check if anything else requires the hook.
	// But do this part outside of the above block because these values may have changed since
	// this function was first called.  The Win9x warning message was removed so that scripts can be
	// run on multiple OSes without a continual warning message just because it happens to be running
	// on Win9x:
	if (   !(sWhichHookNeeded & HOOK_KEYBD)
		&& (!(g_ForceNumLock == NEUTRAL && g_ForceCapsLock == NEUTRAL && g_ForceScrollLock == NEUTRAL)
			|| Hotstring::AtLeastOneEnabled())   ) // Called last for performance due to short-circuit boolean.
		sWhichHookNeeded |= HOOK_KEYBD;

	// Install or deinstall either or both hooks, if necessary, based on these param values.
	// Also, tell it to always display warning if this is a reinstall of the hook(s).
	// Older info about whether to display the "already installed" warning dialog:
	// If script isn't being restarted, shouldn't need much more than one check.
	// If it is being restarted, sometimes the prior instance takes a long time to
	// release the mutex, perhaps due to being partially swapped out (esp. if the
	// prior instance has been running for many hours or days).  UPDATE: This is
	// still displaying a warning sometimes when restarting, even with a 1 second
	// grace/wait period.  So it seems best not to display the warning at all
	// when in restart mode, since obviously (by definition) this script is
	// already running so of course the user wants to restart it unconditionally
	// 99% of the time.  The only exception would be when the user's recent
	// changes to the script (i.e. those for which the restart is being done)
	// now require one of the hooks that hadn't been required before (rare).
	// So for now, when in restart mode, just acquire the mutex but don't display
	// any warning if another instance also has the mutex:
	if (g_IsSuspended)
		ChangeHookState(shk, sHotkeyCount, sWhichHookNeeded, sWhichHookAlways, false);
	else
		ChangeHookState(shk, sHotkeyCount, sWhichHookNeeded, sWhichHookAlways
			, (!g_ForceLaunch && !g_script.mIsRestart) || sHotkeysAreLocked);

	// Fix for v1.0.34: If the auto-execute section uses the Hotkey command but returns before doing
	// something that calls MsgSleep, the main timer won't have been turned on.  For example:
	// Hotkey, Joy1, MySubroutine
	// ;Sleep 1  ; This was a workaround before this fix.
	// return
	// By putting the following check here rather than in AutoHotkey.cpp, that problem is resolved.
	// In addition...
	if (sJoyHotkeyCount)  // Joystick hotkeys require the timer to be always on.
		SET_MAIN_TIMER

	// Signal that no new hotkeys should be defined after this point (i.e. that the definition
	// stage is complete).  Do this only after the the above so that the above can use the old value.
	// UPDATE (for dynamic hotkeys): This now indicates that all static hotkeys have been defined:
	sHotkeysAreLocked = true;
}



ResultType Hotkey::AllDeactivate(bool aObeySuspend, bool aChangeHookStatus, bool aKeepHookIfNeeded)
// Returns OK or FAIL (but currently not failure modes).
{
	if (!sHotkeysAreLocked) // The hotkey definition stage hasn't yet been run, so there's no need.
		return OK;
	if (aChangeHookStatus)
	{
		if (aObeySuspend && g_IsSuspended) // Have the hook keep active only those that are exempt from suspension.
			ChangeHookState(shk, sHotkeyCount, sWhichHookNeeded, sWhichHookAlways, false);
		else
		{
			HookType hooks_to_be_active;
			if (aKeepHookIfNeeded)
			{
				hooks_to_be_active = GetActiveHooks();
				// Caller has already changed the set of hotkeys, perhaps to disable one of them.  If there
				// are any remaining enabled hotkeys which require a particular hook, do not remove that hook.
				// This is done for performance reasons (removing the hook is probably a high overhead operation,
				// and sometimes it seems to take a while to get it installed again) and also to avoid having
				// a time gap in the hook's tracking of key states:
				int i;
				if (g_KeybdHook && !Hotstring::AtLeastOneEnabled())
				{
					for (i = 0; i < sHotkeyCount; ++i)
						if (   shk[i]->mEnabled && (shk[i]->mType == HK_KEYBD_HOOK || shk[i]->mType == HK_BOTH_HOOKS)   )
							break;
					if (i == sHotkeyCount) // No enabled hotkey that requires this hook was found, so remove the hook.
						hooks_to_be_active &= ~HOOK_KEYBD;
				}
				if (g_MouseHook)
				{
					for (i = 0; i < sHotkeyCount; ++i)
						if (   shk[i]->mEnabled && (shk[i]->mType == HK_MOUSE_HOOK || shk[i]->mType == HK_BOTH_HOOKS)   )
							break;
					if (i == sHotkeyCount) // No enabled hotkey that requires this hook was found, so remove the hook.
						hooks_to_be_active &= ~HOOK_MOUSE;
				}
			}
			else // remove all hooks
				hooks_to_be_active = 0;
			AddRemoveHooks(hooks_to_be_active); // No change is made if the hooks are already in the correct state.
		}
	}
	// Unregister all hotkeys except when aObeySuspend is true.  In that case, don't
	// unregister those whose subroutines have ACT_SUSPEND as their first line.  This allows
	// such hotkeys to stay in effect so that the user can press them to turn off the suspension.
	// This also resets the mRunAgainAfterFinished flag for each hotkey that is being deactivated
	// here, including hook hotkeys:
	for (int i = 0; i < sHotkeyCount; ++i)
	{
		// IIRC, !g_IsSuspended is used below in case this function is called from something
		// like AllDestruct(), in which case the Unregister() should happen even if hotkeys
		// aren't suspended:
		if (!aObeySuspend || !g_IsSuspended || !shk[i]->IsExemptFromSuspend())
		{
			shk[i]->Unregister();
			shk[i]->mRunAgainAfterFinished = false;  // ACT_SUSPEND, at least, relies on us to do this.
		}
	}
	// Hot-strings are handled by ToggleSuspendState.

	return OK;
}



ResultType Hotkey::AllDestruct()
// Returns OK or FAIL (but currently not failure modes).
{
	// Tell it to deactivate all unconditionally, even if the script is suspended and
	// some hotkeys are exempt from suspension:
	AllDeactivate(false);
	for (int i = 0; i < sHotkeyCount; ++i)
		delete shk[i];  // unregisters before destroying
	sNextID = 0;
	return OK;
}



void Hotkey::AllDestructAndExit(int aExitCode)
{
	// Might be needed to prevent hang-on-exit.  Once this is done, no message boxes or other dialogs
	// can be displayed.  MSDN: "The exit value returned to the system must be the wParam parameter
	// of the WM_QUIT message."  In our case, PostQuitMessage() should announce the same exit code
	// that we will eventually call exit() with:
	PostQuitMessage(aExitCode);
	AllDestruct();
	// Do this only at the last possible moment prior to exit() because otherwise
	// it may free memory that is still in use by objects that depend on it.
	// This is actually kinda wrong because when exit() is called, the destructors
	// of static, global, and main-scope objects will be called.  If any of these
	// destructors try to reference memory freed() by DeleteAll(), there could
	// be trouble.
	// It's here mostly for traditional reasons.  I'm 99.99999 percent sure that there would be no
	// penalty whatsoever to omitting this, since any modern OS will reclaim all
	// memory dynamically allocated upon program termination.  Indeed, omitting
	// deletes and free()'s for simple objects will often improve the reliability
	// and performance since the OS is far more efficient at reclaiming the memory
	// than us doing it manually (which involves a potentially large number of deletes
	// due to all the objects and sub-objects to be destructed in a typical C++ program).
	// UPDATE: In light of the first paragraph above, it seems best not to do this at all,
	// instead letting all implicitly-called destructors run prior to program termination,
	// at which time the OS will reclaim all remaining memory:
	//SimpleHeap::DeleteAll();

	// In light of the comments below, and due to the fact that anyone using this app
	// is likely to want the anti-focus-stealing measure to always be disabled, I
	// think it's best not to bother with this ever, since its results are
	// unpredictable:
/*	if (g_os.IsWin98orLater() || g_os.IsWin2000orLater())
		// Restore the original timeout value that was set by WinMain().
		// Also disables the compiler warning for the PVOID cast.
		// Note: In many cases, this call will fail to set the value (perhaps even if
		// SystemParametersInfo() reports success), probably because apps aren't
		// supposed to change this value unless they currently have the input
		// focus or something similar (and this app probably doesn't meet the criteria
		// at this stage).  So I think what will happen is: the value set
		// in WinMain() will stay in effect until the user reboots, at which time
		// the default value store in the registry will once again be in effect.
		// This limitation seems harmless.  Indeed, it's probably a good thing not to
		// set it back afterward so that windows behave more consistently for the user
		// regardless of whether this program is currently running.
#ifdef _MSC_VER
	#pragma warning( disable : 4312 )
#endif
		SystemParametersInfo(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, (PVOID)g_OriginalTimeout, SPIF_SENDCHANGE);
#ifdef _MSC_VER
	#pragma warning( default : 4312 ) 
#endif
*/
	// I know this isn't the preferred way to exit the program.  However, due to unusual
	// conditions such as the script having MsgBoxes or other dialogs displayed on the screen
	// at the time the user exits (in which case our main event loop would be "buried" underneath
	// the event loops of the dialogs themselves), this is the only reliable way I've found to exit
	// so far.  The caller has already called PostQuitMessage(), which might not help but it doesn't hurt:
	exit(aExitCode); // exit() is insignificant in code size.  It does more than ExitProcess(), but perhaps nothing more that this application actually requires.
	// By contrast to _exit(), exit() flushes all file buffers before terminating the process. It also
	// calls any functions registered via atexit or _onexit.
}



ResultType Hotkey::PerformID(HotkeyIDType aHotkeyID)
// Returns OK or FAIL.
{
	// Currently, hotkey_id can't be < 0 due to its type, so we only check if it's too large:
	if (aHotkeyID >= sHotkeyCount)
	{
		MsgBox("Hotkey ID too large");
		return FAIL;  // Not a critical error in case some other app is sending us bogus messages?
	}

	static bool dialog_is_displayed = false;  // Prevents double-display caused by key buffering.
	if (dialog_is_displayed) // Another recursion layer is already displaying the warning dialog below.
		return OK; // Don't allow new hotkeys to fire during that time.

	// Help prevent runaway hotkeys (infinite loops due to recursion in bad script files):
	static UINT throttled_key_count = 0;  // This var doesn't belong in struct since it's used only here.
	UINT time_until_now;
	int display_warning;
	if (!sTimePrev)
		sTimePrev = GetTickCount();

	++throttled_key_count;
	sTimeNow = GetTickCount();
	// Calculate the amount of time since the last reset of the sliding interval.
	// Note: A tickcount in the past can be subtracted from one in the future to find
	// the true difference between them, even if the system's uptime is greater than
	// 49 days and the future one has wrapped but the past one hasn't.  This is
	// due to the nature of DWORD math.  The only time this calculation will be
	// unreliable is when the true difference between the past and future
	// tickcounts itself is greater than about 49 days:
	time_until_now = (sTimeNow - sTimePrev);
	if (display_warning = (throttled_key_count > (DWORD)g_MaxHotkeysPerInterval
		&& time_until_now < (DWORD)g_HotkeyThrottleInterval))
	{
		// The moment any dialog is displayed, hotkey processing is halted since this
		// app currently has only one thread.
		char error_text[2048];
		// Using %f with wsprintf() yields a floating point runtime error dialog.
		// UPDATE: That happens if you don't cast to float, or don't have a float var
		// involved somewhere.  Avoiding floats altogether may reduce EXE size
		// and maybe other benefits (due to it not being "loaded")?
		snprintf(error_text, sizeof(error_text), "%u hotkeys have been received in the last %ums.\n\n"
			"Do you want to continue?\n(see #MaxHotkeysPerInterval in the help file)"  // In case its stuck in a loop.
			, throttled_key_count, time_until_now);

		// Turn off any RunAgain flags that may be on, which in essense is the same as de-buffering
		// any pending hotkey keystrokes that haven't yet been fired:
		ResetRunAgainAfterFinished();

		// This is now needed since hotkeys can still fire while a messagebox is displayed.
		// Seems safest to do this even if it isn't always necessary:
		dialog_is_displayed = true;
		g_AllowInterruption = false;
		if (MsgBox(error_text, MB_YESNO) == IDNO)
			g_script.ExitApp(EXIT_CRITICAL); // Might not actually Exit if there's an OnExit subroutine.
		g_AllowInterruption = true;
		dialog_is_displayed = false;
	}
	// The display_warning var is needed due to the fact that there's an OR in this condition:
	if (display_warning || time_until_now > (DWORD)g_HotkeyThrottleInterval)
	{
		// Reset the sliding interval whenever it expires.  Doing it this way makes the
		// sliding interval more sensitive than alternate methods might be.
		// Also reset it if a warning was displayed, since in that case it didn't expire.
		throttled_key_count = 0;
		sTimePrev = sTimeNow;
	}
	if (display_warning)
		// At this point, even though the user chose to continue, it seems safest
		// to ignore this particular hotkey event since it might be WinClose or some
		// other command that would have unpredictable results due to the displaying
		// of the dialog itself.
		return OK;

	return shk[aHotkeyID]->Perform();
}



void Hotkey::TriggerJoyHotkeys(int aJoystickID, DWORD aButtonsNewlyDown)
{
	for (int i = 0; i < sHotkeyCount; ++i)
	{
		Hotkey &hk = *shk[i]; // For performance and convenience.
		// Fix for v1.0.34: If hotkey isn't enabled, or hotkeys are suspended and this one isn't
		// exempt, don't fire it.  These checks are necessary only for joystick hotkeys because 
		// normal hotkeys are completely deactivated when turned off or suspended, but the joystick
		// is still polled even when some joystick hotkeys are disabled.
		if (hk.mType != HK_JOYSTICK || hk.mVK != aJoystickID || !hk.mEnabled
			|| (g_IsSuspended && !hk.IsExemptFromSuspend()))
			continue;
		// Determine if this hotkey's button is among those newly pressed.
		if (   aButtonsNewlyDown & ((DWORD)0x01 << (hk.mSC - JOYCTRL_1))   )
			// Post it to the thread because the message pump itself (not the WindowProc) will handle it:
			PostMessage(NULL, WM_HOTKEY, (WPARAM)i, 0);
	}
}



ResultType Hotkey::Dynamic(char *aHotkeyName, Label *aJumpToLabel, HookActionType aHookAction, char *aOptions)
// Creates, updates, enables, or disables a hotkey dynamically (while the script is running).
// aJumpToLabel can be NULL while aHookAction is 0 only when the caller is updating an existing
// hotkey to have new options (i.e. its current label is not to be changed).
// Returns OK or FAIL.
{
	int hotkey_id = FindHotkeyByName(aHotkeyName);
	Hotkey *hk = (hotkey_id == HOTKEY_ID_INVALID) ? NULL : shk[hotkey_id];
	ResultType result = OK; // Set default.

	switch(aHookAction)
	{
	case 0:
		if (hk)
		{
			if (aJumpToLabel)
				result = hk->UpdateHotkey(aJumpToLabel, 0);  // Update its label to be aJumpToLabel.
			// else do nothing, the caller is probably just trying to change this hotkey's options.
		}
		else // Hotkey needs to be created.
		{
			if (!aJumpToLabel) // Caller is trying to set new aOptions for a non-existent hotkey.
				return g_script.ScriptError(ERR_NONEXISTENT_HOTKEY ERR_ABORT, aHotkeyName);
			// Otherwise, aJumpToLabel is the new target label.
			result = AddHotkey(aJumpToLabel, 0, aHotkeyName);
			if (result == OK)
			{
				hk = shk[sNextID - 1];
				// This is reported only once, when the hotkey is first created, since the user knows
				// about it then and there's no need to report it every time the hotkey is disabled then
				// reenabled, etc.  This need not be checked for XP/2k/NT because the hook will be used
				// to override a registered hotkey:
				if (g_os.IsWin9x())
				{
					if (hk->mType == HK_NORMAL && hk->mEnabled && !hk->mIsRegistered
						|| TYPE_IS_HOOK(hk->mType)) // Type was determined by either AddHotkey() or AllActivate().
						return g_script.ScriptError("Hotkey in use or unsupported." ERR_ABORT, aHotkeyName);
				}
			}
		}
		break;
	case HOTKEY_ID_ON:
	case HOTKEY_ID_OFF:
	case HOTKEY_ID_TOGGLE:
		if (!hk)
			return g_script.ScriptError(ERR_NONEXISTENT_HOTKEY ERR_ABORT, aHotkeyName);
		switch (aHookAction)
		{
		case HOTKEY_ID_ON: result = hk->Enable(); break;
		case HOTKEY_ID_OFF: result = hk->Disable(); break;
		case HOTKEY_ID_TOGGLE: result = hk->mEnabled ? hk->Disable() : hk->Enable(); break;
		}
		break;
	default: // It's one of the alt-tab actions handled by the hook (no label required).
		result = hk ? hk->UpdateHotkey(NULL, aHookAction) : AddHotkey(NULL, aHookAction, aHotkeyName);
	}

	if (result != OK) // The error was already displayed by the above.
		return result;
	if (!*aOptions)
		return OK;  // New hotkeys will have been created using the current values of g_MaxThreadsBuffer, etc.

	result = OK; // Set new default from this point forward.
	int max_threads_per_hotkey;
	bool max_threads_buffer;
	int priority;

	if (!hk)
	{
		hotkey_id = sNextID - 1;  // The ID of the newly added hotkey.
		hk = shk[hotkey_id];
		// Set defaults for all options.  Don't use g_MaxThreadsBuffer, etc., as the default
		// because it seems more useful to have the option letters override a known default
		// and since otherwise, the option letters would have to support a minus prefix or
		// such to allow the option to be put back to its default state (i.e. if
		// g_MaxThreadsBuffer is ON, you would need to use -B to have the hotkey use it as
		// off, which adds unnecessary complexity I think).  UPDATE: Because a hotkey's
		// options can be changed after they were initially set, it seems best to allow
		// the minus sign approach after all (in this case via B vs. B0 option strings),
		// since otherwise there would be no way to turn off the options after they are
		// turned on.
		max_threads_buffer = g_MaxThreadsBuffer;
		max_threads_per_hotkey = g_MaxThreadsPerHotkey;
		priority = 0; // standard default
	}
	else // Set the defaults to be whatever they already were set to when the hotkey was originally created.
	{
		max_threads_buffer = hk->mMaxThreadsBuffer;
		max_threads_per_hotkey = hk->mMaxThreads;
		priority = hk->mPriority;
	}
	for (char *cp = aOptions; *cp; ++cp)
	{
		switch(toupper(*cp))
		{
		case 'O': // v1.0.38.02.  Note that there is no "Off" counterpart because it seems too rarely needed.
			if (toupper(cp[1]) == 'N') // Full validation for maintainability.
			{
				++cp; // Omit the 'N' from further consideration in case it ever becomes a valid option letter.
				result = hk->Enable();
			}
			break;
		case 'B':
			max_threads_buffer = (cp[1] != '0');  // i.e. if the char is NULL or something other than '0'.
			break;
		// For options such as P & T: Use atoi() vs. ATOI() to avoid interpreting something like 0x01B
		// as hex when in fact the B was meant to be an option letter:
		case 'P':
			priority = atoi(cp + 1);
			break;
		case 'T':
			max_threads_per_hotkey = atoi(cp + 1);
			if (max_threads_per_hotkey > MAX_THREADS_LIMIT)
				// For now, keep this limited to prevent stack overflow due to too many pseudo-threads.
				max_threads_per_hotkey = MAX_THREADS_LIMIT;
			else if (max_threads_per_hotkey < 1)
				max_threads_per_hotkey = 1;
			break;
		// Otherwise: Ignore other characters, such as the digits that comprise the number after the T option.
		}
	}
	hk->mMaxThreadsBuffer = max_threads_buffer;
	hk->mMaxThreads = max_threads_per_hotkey;
	hk->mPriority = priority;
	return result;
}



ResultType Hotkey::UpdateHotkey(Label *aJumpToLabel, HookActionType aHookAction)
{
	bool reset = false;
	if (aJumpToLabel != mJumpToLabel)
	{
		if (mJumpToLabel && mJumpToLabel->mName == mName)
			// Since mName points to the original label it must be currently be a static hotkey
			// (one not created by the Hotkey command).  We need to preserve it's original name,
			// so whenever a static hotkey is made dynamic (a one-time event since it can never
			// go in the other direction), we allocate some memory for the original label name
			// because it is the Hotkey's name (Mod+Keyname, e.g. ^c).  Once the mName is
			// allocated in this way, it will never be made to point to another string because
			// that would cause a memory leak:
			mName = SimpleHeap::Malloc(mJumpToLabel->mName);
		// If either the new or old label is NULL, we're changing from a normal hotkey to a special
		// Alt-Tab hotkey, or vice versa.  Due to the complexity of registered vs. hook hotkeys, for now
		// just start from scratch so that there is high confidence that the hook and all registered hotkeys,
		// including their interdependencies, will be re-initialized correctly:
		if (!aJumpToLabel || !mJumpToLabel)
		{
			mJumpToLabel = aJumpToLabel;  // Done after the above.  Should be NULL only when there's an aHookAction.
			reset = true;
		}
		else
		{
			bool exempt_prev = IsExemptFromSuspend();
			mJumpToLabel = aJumpToLabel;
			if (IsExemptFromSuspend() != exempt_prev) // The new label's exempt status is different than the old's.
				reset = true;
		}
	}
	if (aHookAction != mHookAction)
	{
		mHookAction = aHookAction;
		// For now just keep it simple to maximize confidence that everything will be re-initialized okay:
		reset = true;
	}
	if (reset)
	{
		AllDeactivate(false, true, true);
		AllActivate();
	}
	return OK;
}



ResultType Hotkey::AddHotkey(Label *aJumpToLabel, HookActionType aHookAction, char *aName)
// Caller must ensure that either aJumpToLabel or aName is not NULL.
// aName is NULL whenever the caller is creating a static hotkey, at loadtime (i.e. one that
// points to a hotkey label rather than a normal label).  The only time aJumpToLabel should
// be NULL is when the caller is creating a dynamic hotkey that has an aHookAction.
// Return OK or FAIL.
{
	if (   !(shk[sNextID] = new Hotkey(sNextID, aJumpToLabel, aHookAction, aName))   )
		return FAIL;
	if (!shk[sNextID]->mConstructedOK)
	{
		delete shk[sNextID];  // SimpleHeap allows deletion of most recently added item.
		return FAIL;  // The constructor already displayed the error.
	}
	++sNextID;
	if (sHotkeysAreLocked)
	{
		// Reasons to reinitialize the entire hotkey set, perhaps even if the newly added hotkey
		// is just a registered one: Hotkeys affect each other.  See examples in AllActivate()
		// such as "Check if this key is used as the modifier (prefix) for any other key.  If it
		// is, the keyboard hook must handle this key also because otherwise the key-down event
		// would trigger the registered hotkey immediately, rather than waiting to see if this
		// key is be held down merely to modify some other key.".  For now, call AllDeactivate()
		// for cases such as the following: the newly added hotkey has an interaction/dependency
		// with another hotkey, causing it to be promoted from a registered hotkey to a hook hotkey.
		// In such a case, the key should be unregistered for maximum reliability, even though'
		// the hook could probably override the registration in most cases:
		AllDeactivate(false, false);  // Avoid removing the hooks when adding a key.
		AllActivate();
	}
	return OK;
}



Hotkey::Hotkey(HotkeyIDType aID, Label *aJumpToLabel, HookActionType aHookAction, char *aName) // Constructor
	: mID(HOTKEY_ID_INVALID)  // Default until overridden.
	// Caller must ensure that either aName or aJumpToLabel isn't NULL.
	, mVK(0)
	, mSC(0)
	, mModifiers(0)
	, mModifiersLR(0)
	, mAllowExtraModifiers(false)
	, mKeyUp(false)
	, mNoSuppress(0)  // Default is to suppress both prefixes and suffixes.
	, mModifierVK(0)
	, mModifierSC(0)
	, mModifiersConsolidatedLR(0)
	, mType(HK_UNDETERMINED)
	, mVK_WasSpecifiedByNumber(false)
	, mUnregisterDuringThread(false)
	, mIsRegistered(false)
	, mEnabled(true)
	, mHookAction(aHookAction)
	, mJumpToLabel(aJumpToLabel)  // Can be NULL for dynamic hotkeys that are hook actions such as Alt-Tab.
	, mExistingThreads(0)
	, mMaxThreads(g_MaxThreadsPerHotkey)  // The value of g_MaxThreadsPerHotkey can vary during load-time.
	, mMaxThreadsBuffer(g_MaxThreadsBuffer) // same comment as above
	, mPriority(0) // default priority is always 0
	, mRunAgainAfterFinished(false), mRunAgainTime(0), mConstructedOK(false)

// It's better to receive the hotkey_id as a param, since only the caller has better knowledge and
// verification of the fact that this hotkey's id is always set equal to it's index in the array
// (for performance reasons).
{
	if (sHotkeyCount >= MAX_HOTKEYS || sNextID > HOTKEY_ID_MAX)  // The latter currently probably can't happen.
	{
		// This will actually cause the script to terminate if this hotkey is a static (load-time)
		// hotkey.  In the future, some other behavior is probably better:
		MsgBox("The max number of hotkeys has been reached.");  // Brief msg since so rare.
		return;
	}

	char *hotkey_name = aName ? aName : aJumpToLabel->mName;
	if (!TextInterpret(hotkey_name)) // The called function already displayed the error.
		return;

	if (mType == HK_JOYSTICK)
		mModifiersConsolidatedLR = 0;
	else
	{
		char error_text[512];
		if (   (mHookAction == HOTKEY_ID_ALT_TAB || mHookAction == HOTKEY_ID_ALT_TAB_SHIFT)
			&& !mModifierVK && !mModifierSC   )
		{
			if (mModifiers)
			{
				// Neutral modifier has been specified.  Future enhancement: improve this
				// to try to guess which key, left or right, should be used based on the
				// location of the suffix key on the keyboard.
				snprintf(error_text, sizeof(error_text), "The AltTab hotkey \"%s\" must specify which key (L or R).", hotkey_name);
				if (g_script.mIsReadyToExecute) // Dynamically registered via the Hotkey command.
				{
					snprintfcat(error_text, sizeof(error_text), "\n\n%s", ERR_ABORT_NO_SPACES);
					g_script.ScriptError(error_text);
				}
				else
					MsgBox(error_text);
				return;  // Key is invalid so don't give it an ID.
			}
			if (mModifiersLR)
			{
				switch (mModifiersLR)
				{
				case MOD_LCONTROL: mModifierVK = g_os.IsWin9x() ? VK_CONTROL : VK_LCONTROL; break;
				case MOD_RCONTROL: mModifierVK = g_os.IsWin9x() ? VK_CONTROL : VK_RCONTROL; break;
				case MOD_LSHIFT: mModifierVK = g_os.IsWin9x() ? VK_SHIFT : VK_LSHIFT; break;
				case MOD_RSHIFT: mModifierVK = g_os.IsWin9x() ? VK_SHIFT : VK_RSHIFT; break;
				case MOD_LALT: mModifierVK = g_os.IsWin9x() ? VK_MENU : VK_LMENU; break;
				case MOD_RALT: mModifierVK = g_os.IsWin9x() ? VK_MENU : VK_RMENU; break;
				case MOD_LWIN: mModifierVK = VK_LWIN; break; // Win9x should support LWIN/RWIN.
				case MOD_RWIN: mModifierVK = VK_RWIN; break;
				default:
					snprintf(error_text, sizeof(error_text), "The AltTab hotkey \"%s\" must have exactly "
						"one modifier/prefix.", hotkey_name);
					if (g_script.mIsReadyToExecute) // Dynamically registered via the Hotkey command.
					{
						snprintfcat(error_text, sizeof(error_text), ERR_ABORT);
						g_script.ScriptError(error_text);
					}
					else
						MsgBox(error_text);
					return;  // Key is invalid so don't give it an ID.
				}
				// Since above didn't return:
				mModifiersLR = 0;  // Since ModifierVK/SC is now its substitute.
			}
			// Update: This is no longer needed because the hook attempts to compensate.
			// However, leaving it enabled may improve performance and reliability.
			// Update#2: No, it needs to be disabled, otherwise alt-tab won't work right
			// in the rare case where an ALT key itself is defined as "AltTabMenu":
			//else
				// It has no ModifierVK/SC and no modifiers, so it's a hotkey that is defined
				// to fire only when the Alt-Tab menu is visible.  Since the ALT key must be
				// down for that menu to be visible (on all OSes?), add the ALT key to this
				// keys modifiers so that it will be detected as a hotkey whenever the
				// Alt-Tab menu is visible:
			//	modifiers |= MOD_ALT;
		}

		if (mType != HK_MOUSE_HOOK && mType != HK_JOYSTICK)
			if ((g_ForceKeybdHook || mModifiersLR || mAllowExtraModifiers || mNoSuppress || aHookAction || mKeyUp)
				&& !g_os.IsWin9x())
				// Do this for both NO_SUPPRESS_SUFFIX and NO_SUPPRESS_PREFIX.  In the case of
				// NO_SUPPRESS_PREFIX, the hook is needed anyway since the only way to get
				// NO_SUPPRESS_PREFIX in effect is with a hotkey that has a ModifierVK/SC.
				mType = HK_KEYBD_HOOK;

		// Currently, these take precedence over each other in the following order, so don't
		// just bitwise-or them together in case there's any ineffectual stuff stored in
		// the fields that have no effect (e.g. modifiers have no effect if there's a mModifierVK):
		if (mModifierVK)
			mModifiersConsolidatedLR = KeyToModifiersLR(mModifierVK);
		else if (mModifierSC)
			mModifiersConsolidatedLR = KeyToModifiersLR(0, mModifierSC);
		else
		{
			mModifiersConsolidatedLR = mModifiersLR;
			if (mModifiers)
				mModifiersConsolidatedLR |= ConvertModifiers(mModifiers);
		}
	}

	// To avoid memory leak, this is done only when it is certain the hotkey will be created:
	if (   !(mName = aName ? SimpleHeap::Malloc(aName) : hotkey_name)   )
	{
		g_script.ScriptError(ERR_OUTOFMEM);  // Very rare.
		return;
	}

	// Always assign the ID last, right before a successful return, so that the caller is notified
	// that the constructor succeeded:
	mConstructedOK = true;
	mID = aID;
	// Don't do this because the caller still needs the old/unincremented value:
	//++sHotkeyCount;  // Hmm, seems best to do this here, but revisit this sometime.
}



ResultType Hotkey::TextInterpret(char *aName)
// Returns OK or FAIL.
{
	// Make a copy that can be modified:
	char hotkey_name[256];
	strlcpy(hotkey_name, aName, sizeof(hotkey_name));
	char *term1 = hotkey_name;
	char *term2 = strstr(term1, COMPOSITE_DELIMITER);
	if (!term2)
		return TextToKey(TextToModifiers(term1), aName, false);
	if (*term1 == '~')
	{
		mNoSuppress |= NO_SUPPRESS_PREFIX;
		term1 = omit_leading_whitespace(term1 + 1);
	}
    char *end_of_term1 = omit_trailing_whitespace(term1, term2) + 1;
	// Temporarily terminate the string so that the 2nd term is hidden:
	char ctemp = *end_of_term1;
	*end_of_term1 = '\0';
	ResultType result = TextToKey(term1, aName, true);
	*end_of_term1 = ctemp;  // Undo the termination.
	if (result == FAIL)
		return FAIL;
	term2 += COMPOSITE_DELIMITER_LENGTH;
	term2 = omit_leading_whitespace(term2);
	// Even though modifiers on keys already modified by a mModifierVK are not supported, call
	// TextToModifiers() anyway to use its output (for consistency).  The modifiers it sets
	// are currently ignored because the mModifierVK takes precedence.
	return TextToKey(TextToModifiers(term2), aName, false);
}



char *Hotkey::TextToModifiers(char *aText)
// Takes input param <text> to support receiving only a subset of object.text.
// Returns the location in <text> of the first non-modifier key.
// Checks only the first char(s) for modifiers in case these characters appear elsewhere (e.g. +{+}).
// But come to think of it, +{+} isn't valid because + itself is already shift-equals.  So += would be
// used instead, e.g. +==action.  Similarly, all the others, except =, would be invalid as hotkeys also.
{
	if (!aText || !*aText) return aText; // Below relies on this having ensured that aText isn't blank.

	// Explicitly avoids initializing modifiers to 0 because the caller may have already included
	// some set some modifiers in there.
	char *marker;
	bool key_left, key_right;

	// Improved for v1.0.37.03: The loop's condition is now marker[1] vs. marker[0] so that
	// the last character is never considered a modifier.  This allows a modifier symbol
	// to double as the name of a suffix key.  It also fixes issues on layouts where the
	// symbols +^#! do not require the shift key to be held down, such as the German layout.
	for (marker = aText, key_left = false, key_right = false; marker[1]; ++marker)
	{
		switch (*marker)
		{
		case '>':
			key_right = true;
			break;
		case '<':
			key_left = true;
			break;
		case '*':
			mAllowExtraModifiers = true;
			break;
		case '~':
			mNoSuppress |= NO_SUPPRESS_SUFFIX;
			break;
		case '$':
			if (g_os.IsWin9x())
				mUnregisterDuringThread = true;
			else
				mType = HK_KEYBD_HOOK;
			// else ignore the flag and try to register normally, which in most cases seems better
			// than disabling the hotkey.
			break;
		case '!':
			if ((!key_right && !key_left))
			{
				mModifiers |= MOD_ALT;
				break;
			}
			// Both left and right may be specified, e.g. ><+a means both shift keys must be held down:
			if (key_left)
			{
				mModifiersLR |= MOD_LALT;
				key_left = false;
			}
			if (key_right)
			{
				mModifiersLR |= MOD_RALT;
				key_right = false;
			}
			break;
		case '^':
			if ((!key_right && !key_left))
			{
				mModifiers |= MOD_CONTROL;
				break;
			}
			if (key_left)
			{
				mModifiersLR |= MOD_LCONTROL;
				key_left = false;
			}
			if (key_right)
			{
				mModifiersLR |= MOD_RCONTROL;
				key_right = false;
			}
			break;
		case '+':
			if ((!key_right && !key_left))
			{
				mModifiers |= MOD_SHIFT;
				break;
			}
			if (key_left)
			{
				mModifiersLR |= MOD_LSHIFT;
				key_left = false;
			}
			if (key_right)
			{
				mModifiersLR |= MOD_RSHIFT;
				key_right = false;
			}
			break;
		case '#':
			if ((!key_right && !key_left))
			{
				mModifiers |= MOD_WIN;
				break;
			}
			if (key_left)
			{
				mModifiersLR |= MOD_LWIN;
				key_left = false;
			}
			if (key_right)
			{
				mModifiersLR |= MOD_RWIN;
				key_right = false;
			}
			break;
		default:
			return marker;  // Return immediately whenever a non-modifying char is found.
		}
	}
	return marker;
}



ResultType Hotkey::TextToKey(char *aText, char *aHotkeyName, bool aIsModifier)
// Caller must ensure that aText is a modifiable string.
// Takes input param aText to support receiving only a subset of mName.
// In private members, sets the values of vk/sc or ModifierVK/ModifierSC depending on aIsModifier.
// It may also merge new modifiers into the existing value of modifiers, so the caller
// should never reset modifiers after calling this.
// Returns OK or FAIL.
{
	char error_text[512];
	if (!aText || !*aText)
	{
		snprintf(error_text, sizeof(error_text), "\"%s\" is not a valid hotkey."
			" A shifted hotkey such as ? should be defined as +/.", aHotkeyName);
		if (g_script.mIsReadyToExecute) // Dynamically registered via the Hotkey command.
		{
			snprintfcat(error_text, sizeof(error_text), ERR_ABORT);
			g_script.ScriptError(error_text);
		}
		else
			MsgBox(error_text);
		return FAIL;
	}

	vk_type temp_vk; // No need to initialize this one.
	sc_type temp_sc = 0;
	modLR_type modifiersLR = 0;
	bool is_mouse = false;
	int joystick_id;

	if (!aIsModifier)
	{
		// Previous steps should make it unnecessary to call omit_leading_whitespace(aText).
		char *cp = StrChrAny(aText, " \t"); // Find first space or tab.
		if (cp && !stricmp(omit_leading_whitespace(cp), "Up"))
		{
			// This is a key-up hotkey, such as "Ctrl Up::".
			mKeyUp = true;
			*cp = '\0'; // Terminate at the first space so that the word "up" is removed from further consideration by us and callers.
		}
	}

	if (temp_vk = TextToVK(aText, &modifiersLR, true)) // Assign.
	{
		if (aIsModifier)
		{
			if (temp_vk == VK_WHEEL_DOWN || temp_vk == VK_WHEEL_UP)
			{
				snprintf(error_text, sizeof(error_text), "\"%s\" is not allowed as a prefix key.", aText);
				if (g_script.mIsReadyToExecute) // Dynamically registered via the Hotkey command.
				{
					snprintfcat(error_text, sizeof(error_text), ERR_ABORT);
					g_script.ScriptError(error_text);
				}
				else
					MsgBox(error_text);
				return FAIL;
			}
		}
		else
			// This is done here rather than at some later stage because we have access to the raw
			// name of the suffix key (with any leading modifiers such as ^ omitted from the beginning):
			mVK_WasSpecifiedByNumber = !strnicmp(aText, "VK", 2);
		is_mouse = IsMouseVK(temp_vk);
		if (modifiersLR & (MOD_LSHIFT | MOD_RSHIFT))
			if (temp_vk >= 'A' && temp_vk <= 'Z')  // VK of an alpha char is the same as the ASCII code of its uppercase version.
				modifiersLR &= ~(MOD_LSHIFT | MOD_RSHIFT);
				// Above: Making alpha chars case insensitive seems much more friendly.  In other words,
				// if the user defines ^Z as a hotkey, it will really be ^z, not ^+z.  By removing SHIFT
				// from the modifiers here, we're only removing it from our modifiers, not the global
				// modifiers that have already been set elsewhere for this key (e.g. +Z will still be +z).
	}
	else // No virtual key was found.  Is there a scan code?
		if (   !(temp_sc = TextToSC(aText))   )
			if (   !(temp_sc = (sc_type)ConvertJoy(aText, &joystick_id, true))   )  // Is there a joystick control/button?
			{
				snprintf(error_text, sizeof(error_text), "\"%s\" is not a valid key name within a hotkey label.", aText);
				if (g_script.mIsReadyToExecute) // Dynamically registered via the Hotkey command.
				{
					snprintfcat(error_text, sizeof(error_text), ERR_ABORT);
					g_script.ScriptError(error_text);
				}
				else
					MsgBox(error_text);
				return FAIL;
			}
			else
			{
				++sJoyHotkeyCount;
				mType = HK_JOYSTICK;
				temp_vk = (vk_type)joystick_id;  // 0 is the 1st joystick, 1 the 2nd, etc.
				sJoystickHasHotkeys[joystick_id] = true;
			}


/*
If ever do this, be sure to avoid doing it for keys that must be tracked by scan code (e.g. those in the
scan code array).
	if (!temp_vk && !is_mouse)  // sc must be non-zero or else it would have already returned above.
		if (temp_vk = sc_to_vk(temp_sc))
		{
			snprintf(error_text, sizeof(error_text), "DEBUG: \"%s\" (scan code %X) was successfully mapped to virtual key %X", text, temp_sc, temp_vk);
			MsgBox(error_text);
			temp_sc = 0; // Maybe set this just for safety, even though a non-zero vk should always take precedence over it.
		}
*/
	if (is_mouse)
		mType = HK_MOUSE_HOOK;

	if (aIsModifier)
	{
		mModifierVK = temp_vk;
		mModifierSC = temp_sc;
		if (!is_mouse && mType != HK_JOYSTICK)
			mType = HK_KEYBD_HOOK;  // Always use the hook for keys that have a mModifierVK or SC
	}
	else
	{
		mVK = temp_vk;
		mSC = temp_sc;
		// Turn on any additional modifiers.  e.g. SHIFT to realize '#'.
		// Fix for v1.0.37.03: To avoid using the keyboard hook for something like "+::", which in
		// turn would allow the hotkey fire only for LShift+Equals rather than RShift+Equals, convert
		// modifiers from left-right to neutral.  But exclude right-side modifiers (except RWin) so that
		// things like AltGr are more precisely handled (the implications of this policy could use
		// further review).  Currently, right-Alt (via AltGr) is the only possible right-side key.
		mModifiers |= ConvertModifiersLR(modifiersLR & (MOD_RWIN|MOD_LWIN|MOD_LCONTROL|MOD_LALT|MOD_LSHIFT));
		mModifiersLR |= (modifiersLR & (MOD_RSHIFT|MOD_RALT|MOD_RCONTROL)); // Not MOD_RWIN since it belongs above.
		if (!is_mouse && mType != HK_JOYSTICK)
		{
			// For these, if it's Win9x, attempt to register them normally to give the user at least
			// some partial functiality.  The key will probably be toggled to its opposite state when
			// it's used as a hotkey, but the user may be able to concoct a script workaround for that:
			if (g_os.IsWin9x())
			{
				// Fix for v1.0.25: If no VK could be found, try to find one so that the attempt
				// to register the hotkey will have a chance of success.  This change is known to
				// permit all the keys normally handled by scan code to work as hotkeys on Win9x.
				// Namely: Del, Ins, Home, End, PgUp/PgDn, and the arrow keys.
				// One of the main reasons for handling these keys by scan code is that
				// they each have a counterpart on the Numpad, and in many cases a user would
				// not want both to become hotkeys.  It could be argued that this should be changed
				// For Win NT/2k/XP too, but then it would have to be documented that the $ prefix
				// would have to be used to prevent the "both keys of the pair" behavior.  This
				// confusion plus the chance of breaking existing scripts doesn't seem worth the
				// small benefit of being able to avoid the keyboard hook in cases where these
				// would be the only hotkeys using it.  If there is ever a need, a new hotkey
				// prefix or option can be added someday to handle this, perhaps #LinkPairedKeys
				// to avoid having yet another reserved hotkey prefix/symbol.
				if (!mVK)
					mVK = sc_to_vk(temp_sc);
			}
			else // OS isn't Win9x.
			{
				if (mType != HK_MOUSE_HOOK) // Added in v1.0.39 to make a hotkey such as "LButton & LCtrl" install the mouse hook.
				{
					
					switch (mVK)
					{
					case 0: // Scan codes having no available virtual key must always be handled by the hook.
					// In addition, to support preventing the toggleable keys from toggling, handle those
					// with the hook also:
					case VK_NUMLOCK:
					case VK_CAPITAL:
					case VK_SCROLL:
					// Finally, the non-neutral (left-right) modifier keys (except LWin and RWin) must also
					// be done with the hook because even if RegisterHotkey() claims to succeed on them,
					// I'm 99% sure I tried it and the hotkeys don't actually work with that method:
					case VK_LCONTROL:
					case VK_RCONTROL:
					case VK_LSHIFT:
					case VK_RSHIFT:
					case VK_LMENU:
					case VK_RMENU:
						mType = HK_KEYBD_HOOK;
						break;
					}
				}
			}
		}
	}
	return OK;
}



ResultType Hotkey::Register()
// Returns OK or FAIL.
{
	if (mIsRegistered || !mEnabled) return OK;  // It's normal for a disabled hotkey to return OK.
	// Can't use the API method to register such hotkeys.  They are handled by the hook:
	if (mType != HK_NORMAL) return FAIL;

	// Indicate that the key modifies itself because RegisterHotkey() requires that +SHIFT,
	// for example, be used to register the naked SHIFT key.  So what we do here saves the
	// user from having to specify +SHIFT in the script:
	mod_type modifiers_prev = mModifiers;
	switch (mVK)
	{
	case VK_LWIN:
	case VK_RWIN: mModifiers |= MOD_WIN; break;
	case VK_CONTROL: mModifiers |= MOD_CONTROL; break;
	case VK_SHIFT: mModifiers |= MOD_SHIFT; break;
	case VK_MENU: mModifiers |= MOD_ALT; break;
	}

/*
	if (   !(mIsRegistered = RegisterHotKey(NULL, id, modifiers, vk))   )
		// If another app really is already using this hotkey, there's no point in trying to have the keyboard
		// hook try to handle it instead because registered hotkeys take precedence over keyboard hook.
		// UPDATE: For WinNT/2k/XP, this warning can be disabled because registered hotkeys can be
		// overridden by the hook.  But something like this is probably needed for Win9x:
		char error_text[MAX_EXEC_STRING];
		snprintf(error_text, sizeof(error_text), "RegisterHotKey() of hotkey \"%s\" (id=%d, virtual key=%d, modifiers=%d) failed,"
			" perhaps because another application (or Windows itself) is already using it."
			"  You could try adding the line \"%s, On\" prior to its line in the script."
			, text, id, vk, modifiers, g_cmd[CMD_FORCE_KEYBD_HOOK]);
		MsgBox(error_text);
		return FAIL;
	return 1;
*/
	// Must register them to our main window (i.e. don't use NULL to indicate our thread),
	// otherwise any modal dialogs, such as MessageBox(), that call DispatchMessage()
	// internally wouldn't be able to find anyone to send hotkey messages to, so they
	// would probably be lost:
	if (mIsRegistered = RegisterHotKey(g_hWnd, mID, mModifiers, mVK))
		return OK;

	// On failure, reset the modifiers in case this function changed them.  This is done
	// in case this hotkey will now be handled by the hook, which doesn't want any
	// extra modifiers that were added above:
	mModifiers = modifiers_prev;
	return FAIL;
}



ResultType Hotkey::Unregister()
// Returns OK or FAIL.
{
	if (!mIsRegistered) return OK;
	// Don't report any errors in here, at least not when we were called in conjunction
	// with cleanup and exit.  Such reporting might cause an infinite loop, leading to
	// a stack overflow if the reporting itself encounters an error and tries to exit,
	// which in turn would call us again:
	if (mIsRegistered = !UnregisterHotKey(g_hWnd, mID))  // I've see it fail in one rare case.
		return FAIL;
	return OK;
}



void Hotkey::InstallKeybdHook()
{
	sWhichHookAlways |= HOOK_KEYBD;
	ChangeHookState(shk, sHotkeyCount, sWhichHookNeeded, sWhichHookAlways, false);
}



int Hotkey::FindHotkeyByName(char *aName)
// Returns the the HotkeyID if found, HOTKEY_ID_INVALID otherwise.
{
	// Originally I thought it might be best to exclude ~ and $ from the comparison by using
	// stricmp_exclude().  However, if this is done, there would be no way to change a hotkey
	// from hook to normal (or vice versa), and also no way to change the pass-through nature
	// of a hotkey once it is created (since hotkeys can be disabled but never destroyed).
	// Although special handling could be added for this, it seems best to allow these duplicates
	// to be defined, for both max simplicity and flexibility.  The hook will be tested to ensure
	// that it can properly handle duplicate definitions such as $^h and ~^h (testing reveals
	// that the first hotkey to be defined will take precedence).
	for (int i = 0; i < sHotkeyCount; ++i)
		if (!stricmp(shk[i]->mName, aName)) // Case insensitive so that something like ^A is a match for ^a
			return i;
	return HOTKEY_ID_INVALID;  // No match found.
}



int Hotkey::FindHotkeyWithThisModifier(vk_type aVK, sc_type aSC)
// Returns the the HotkeyID if found, HOTKEY_ID_INVALID otherwise.
// Answers the question: What is the first hotkey with mModifierVK or mModifierSC equal to those given?
// A non-zero vk param will take precendence over any non-zero value for sc.
{
	if (!aVK & !aSC) return HOTKEY_ID_INVALID;
	for (int i = 0; i < sHotkeyCount; ++i)
		if (   (aVK && aVK == shk[i]->mModifierVK) || (aSC && aSC == shk[i]->mModifierSC)   )
			return i;
	return HOTKEY_ID_INVALID;  // No match found.
}



int Hotkey::FindHotkeyContainingModLR(modLR_type aModifiersLR) // , int hotkey_id_to_omit)
// Returns the the HotkeyID if found, HOTKEY_ID_INVALID otherwise.
// Find the first hotkey whose modifiersLR contains *any* of the modifiers shows in the parameter value.
// The caller tells us the ID of the hotkey to omit from the search because that one
// would always be found (since something like "lcontrol=calc.exe" in the script
// would really be defines as  "<^control=calc.exe".
// Note: By intent, this function does not find hotkeys whose normal/neutral modifiers
// contain <modifiersLR>.
{
	if (!aModifiersLR) return HOTKEY_ID_INVALID;
	for (int i = 0; i < sHotkeyCount; ++i)
		// Bitwise set-intersection: indicates if anything in common:
		if (shk[i]->mModifiersLR & aModifiersLR)
		//if (i != hotkey_id_to_omit && shk[i]->mModifiersLR & modifiersLR)
			return i;
	return HOTKEY_ID_INVALID;  // No match found.
}


/*
int Hotkey::FindHotkeyBySC(sc2_type aSC2, mod_type aModifiers, modLR_type aModifiersLR)
// Returns the the HotkeyID if found, HOTKEY_ID_INVALID otherwise.
// Answers the question: What is the first hotkey with the given sc & modifiers *regardless* of
// any non-zero mModifierVK or mModifierSC it may have?  The mModifierSC/vk is ignored because
// the caller wants to know whether this key would be blocked if its counterpart were registered.
// For example, the hook wouldn't see "MEDIA_STOP & NumpadENTER" at all if NumPadENTER was
// already registered via RegisterHotkey(), since RegisterHotkey() doesn't honor any modifiers
// other than the standard ones.
{
	for (int i = 0; i < sHotkeyCount; ++i)
		if (!shk[i]->mVK && (shk[i]->mSC == aSC2.a || shk[i]->mSC == aSC2.b))
			if (shk[i]->mModifiers == aModifiers && shk[i]->mModifiersLR == aModifiersLR)  // Ensures an exact match.
				return i;
	return HOTKEY_ID_INVALID;  // No match found.
}
*/


char *Hotkey::ListHotkeys(char *aBuf, int aBufSize)
// Translates this script's list of variables into text equivalent, putting the result
// into aBuf and returning the position in aBuf of its new string terminator.
{
	char *aBuf_orig = aBuf;
	// Save vertical space by limiting newlines here:
	aBuf += snprintf(aBuf, BUF_SPACE_REMAINING, "Type\tOff?\tRunning\tName\r\n"
							 "-------------------------------------------------------------------\r\n");
	// Start at the oldest and continue up through the newest:
	for (int i = 0; i < sHotkeyCount; ++i)
		aBuf = shk[i]->ToText(aBuf, BUF_SPACE_REMAINING, true);
	return aBuf;
}



char *Hotkey::ToText(char *aBuf, int aBufSize, bool aAppendNewline)
// aBufSize is an int so that any negative values passed in from caller are not lost.
// Caller has ensured that aBuf isn't NULL.
// Translates this var into its text equivalent, putting the result into aBuf and
// returning the position in aBuf of its new string terminator.
{
	char *aBuf_orig = aBuf;
	char existing_threads_str[128];
	if (mExistingThreads)
		_itoa(mExistingThreads, existing_threads_str, 10);
	else
		*existing_threads_str = '\0'; // Make it blank to avoid clutter in the hotkey display.
	char htype[32];
	switch (mType)
	{
	case HK_NORMAL: strlcpy(htype, "reg", sizeof(htype)); break;
	case HK_KEYBD_HOOK: strlcpy(htype, "k-hook", sizeof(htype)); break;
	case HK_MOUSE_HOOK: strlcpy(htype, "m-hook", sizeof(htype)); break;
	case HK_BOTH_HOOKS: strlcpy(htype, "2-hooks", sizeof(htype)); break;
	case HK_JOYSTICK: strlcpy(htype, "joypoll", sizeof(htype)); break;
	default: *htype = '\0';
	}
	aBuf += snprintf(aBuf, BUF_SPACE_REMAINING, "%s%s\t%s\t%s\t%s"
		, htype, (mType == HK_NORMAL && !mIsRegistered) ? "(no)" : ""
		, mEnabled ? "" : "OFF"
		, existing_threads_str
		, mName);
	if (aAppendNewline && BUF_SPACE_REMAINING >= 2)
	{
		*aBuf++ = '\r';
		*aBuf++ = '\n';
		*aBuf = '\0';
	}
	return aBuf;
}


///////////////
// Hot Strings
///////////////

// Init static variables:
Hotstring **Hotstring::shs = NULL;
HotstringIDType Hotstring::sHotstringCount = 0;
HotstringIDType Hotstring::sHotstringCountMax = 0;


ResultType Hotstring::Perform()
// Returns OK or FAIL.  Caller has already ensured that the backspacing (if specified by mDoBackspace)
// has been done.
{
	if (mExistingThreads >= mMaxThreads && !ACT_IS_ALWAYS_ALLOWED(mJumpToLabel->mJumpToLine->mActionType))
		return FAIL;
	// See Hotkey::Perform() for details about this.  For hot strings -- which also use the
	// g_script.mThisHotkeyStartTime value to determine whether g_script.mThisHotkeyModifiersLR
	// is still timely/accurate -- it seems best to set to "no modifiers":
	g_script.mThisHotkeyModifiersLR = 0;
	++mExistingThreads;  // This is the thread count for this particular hotstring only.
	ResultType result = mJumpToLabel->mJumpToLine->ExecUntil(UNTIL_RETURN);
	--mExistingThreads;
	return result ? OK : FAIL;	// Return OK on all non-failure results.
}



void Hotstring::DoReplace(LPARAM alParam)
// LOWORD(alParam) is the char from the set of EndChars that the user had to press to trigger the hotkey.
// This is not applicable if mEndCharRequired is false, in which case caller should have passed zero.
{
	char SendBuf[LINE_SIZE + 20] = "";  // Allow extra room for the optional backspaces.
	char *start_of_replacement = SendBuf;  // Set default.
	if (mDoBackspace)
	{
		// Subtract 1 from backspaces because the final key pressed by the user to make a
		// match was already suppressed by the hook (it wasn't sent through to the active
		// window).  So what we do is backspace over all the other keys prior to that one,
		// put in the replacement text (if applicable), then send the EndChar through
		// (if applicable) to complete the sequence.
		int backspace_count = mStringLength - 1;
		if (mEndCharRequired)
			++backspace_count;
		for (int i = 0; i < backspace_count; ++i)
			*start_of_replacement++ = '\b';  // Use raw backspaces, not {BS n}, in case the send will be raw.
		*start_of_replacement = '\0';
	}
	if (*mReplacement)
	{
		// Below casts to int to preserve any negaive results:
		snprintf(start_of_replacement, (int)(sizeof(SendBuf) - (start_of_replacement - SendBuf)), "%s", mReplacement);
		CaseConformModes case_conform_mode = (CaseConformModes)HIWORD(alParam);
		if (case_conform_mode == CASE_CONFORM_ALL_CAPS)
			CharUpper(start_of_replacement);
		else if (case_conform_mode == CASE_CONFORM_FIRST_CAP)
			*start_of_replacement = (char)CharUpper((LPTSTR)(UCHAR)*start_of_replacement);
	}
	int old_delay = g.KeyDelay;
	int old_press_duration = g.PressDuration;
	g.KeyDelay = mKeyDelay; // This is relatively safe since SendKeys() normally can't be interrupted by a new thread.
	g.PressDuration = -1;   // Always -1, since Send command can be used in body of hotstring to have a custom press duration.
	if (*SendBuf)
		SendKeys(SendBuf, mSendRaw);
	if (   *mReplacement && !mOmitEndChar && (*SendBuf = (char)LOWORD(alParam))   ) // Assign
	{
		// Send the final character in raw mode so that chars such as !{} are sent properly.
		SendBuf[1] = '\0'; // Terminate.
		// Fix for v1.0.25.08: CollectInput() translates the user's Enter keystrokes into `n
		// instead of `r for use with the Input command. Since the an Input command may be in progress
		// while monitoring hotstrings, it is probably best not to try to change things in the hook
		// (unless the policy of translating Enter to `n is someday reversed).  Instead, any `n
		// received here, which corresponds to the user's physical press of Enter or Shift-Enter
		// (Ctrl-Enter doesn't seem to have an ascii counterpart), is translated back to `r.  This
		// prevents the user's Enter keystroke (which would be the end-char that triggers this
		// hotstring) from being translated into Ctrl-Enter.  Ctrl-Enter has a different effect in
		// most word processors than Enter, producing a page break or something else undesirable.
		// Update for v1.0.25.12: The below is no longer necessary because SendKeys() treats
		// \n the same as \r now:
		//if (*SendBuf == '\n')
		//	*SendBuf = '\r';
		SendKeys(SendBuf, true);
	}
	g.KeyDelay = old_delay;  // Restore
	g.PressDuration = old_press_duration;  // Restore
}



ResultType Hotstring::AddHotstring(Label *aJumpToLabel, char *aOptions, char *aHotstring, char *aReplacement
	, bool aHasContinuationSection)
// Returns OK or FAIL.
// Caller has ensured that aHotstringOptions is blank if there are no options.  Otherwise, aHotstringOptions
// should end in a colon, which marks the end of the options list.  aHotstring is the hotstring itself
// (e.g. "ahk"), which does not have to be unique, unlike the label name, which was made unique by also
// including any options in with the label name (e.g. ::ahk:: is a different label than :c:ahk::).
// Caller has also ensured that aHotstring is not blank.
{
	// The length is limited for performance reasons, notably so that the hook does not have to move
	// memory around in the buffer it uses to watch for hotstrings:
	if (strlen(aHotstring) > MAX_HOTSTRING_LENGTH)
		return g_script.ScriptError("Hotstring max abbreviation length is " MAX_HOTSTRING_LENGTH_STR ".", aHotstring);

	if (!shs)
	{
		if (   !(shs = (Hotstring **)malloc(HOTSTRING_BLOCK_SIZE * sizeof(Hotstring *)))   )
			return g_script.ScriptError(ERR_OUTOFMEM); // Short msg. since so rare.
		sHotstringCountMax = HOTSTRING_BLOCK_SIZE;
	}
	else if (sHotstringCount >= sHotstringCountMax) // Realloc to preserve contents and keep contiguous array.
	{
		// Expand the array by one block.  Use a temp var. because realloc() returns NULL on failure
		// but leaves original block allocated.
		void *realloc_temp = realloc(shs, (sHotstringCountMax + HOTSTRING_BLOCK_SIZE) * sizeof(Hotstring *));
		if (!realloc_temp)
			return g_script.ScriptError(ERR_OUTOFMEM);  // Short msg. since so rare.
		shs = (Hotstring **)realloc_temp;
		sHotstringCountMax += HOTSTRING_BLOCK_SIZE;
	}

	if (   !(shs[sHotstringCount] = new Hotstring(aJumpToLabel, aOptions, aHotstring, aReplacement, aHasContinuationSection))   )
		return g_script.ScriptError(ERR_OUTOFMEM); // Short msg. since so rare.
	if (!shs[sHotstringCount]->mConstructedOK)
	{
		delete shs[sHotstringCount];  // SimpleHeap allows deletion of most recently added item.
		return FAIL;  // The constructor already displayed the error.
	}

	++sHotstringCount;
	return OK;
}



Hotstring::Hotstring(Label *aJumpToLabel, char *aOptions, char *aHotstring, char *aReplacement, bool aHasContinuationSection)
	: mJumpToLabel(aJumpToLabel)  // Can be NULL for dynamic hotkeys that are hook actions such as Alt-Tab.
	, mString(NULL), mReplacement(""), mStringLength(0)
	, mSuspended(false)
	, mExistingThreads(0)
	, mMaxThreads(g_MaxThreadsPerHotkey)  // The value of g_MaxThreadsPerHotkey can vary during load-time.
	, mPriority(g_HSPriority), mKeyDelay(g_HSKeyDelay)  // And all these can vary too.
	, mCaseSensitive(g_HSCaseSensitive), mConformToCase(g_HSConformToCase), mDoBackspace(g_HSDoBackspace)
	, mOmitEndChar(g_HSOmitEndChar), mSendRaw(aHasContinuationSection ? true : g_HSSendRaw)
	, mEndCharRequired(g_HSEndCharRequired), mDetectWhenInsideWord(g_HSDetectWhenInsideWord), mDoReset(g_HSDoReset)
	, mConstructedOK(false)
{
	// Insist on certain qualities so that they never need to be checked other than here:
	if (!mJumpToLabel) // Caller has already ensured that aHotstring is not blank.
		return;

	ParseOptions(aOptions, mPriority, mKeyDelay, mCaseSensitive, mConformToCase, mDoBackspace, mOmitEndChar
		, mSendRaw, mEndCharRequired, mDetectWhenInsideWord, mDoReset);

	// To avoid memory leak, this is done only when it is certain the hotkey will be created:
	if (   !(mString = SimpleHeap::Malloc(aHotstring))   )
	{
		g_script.ScriptError(ERR_OUTOFMEM); // Short msg since very rare.
		return;
	}
	mStringLength = (UCHAR)strlen(mString);
	if (*aReplacement)
	{
		// To avoid wasting memory due to SimpleHeap's block-granularity, only allocate there if the replacement
		// string is short (note that replacement strings can be over 16,000 characters long).  Since
		// hotstrings can be disabled but never entirely deleted, it's not a memory leak in either case
		// since memory allocated by either method will be freed when the program exits.
		size_t size = strlen(aReplacement) + 1;
		if (   !(mReplacement = (size > MAX_ALLOC_SIMPLE) ? (char *)malloc(size) : SimpleHeap::Malloc(size))   )
		{
			g_script.ScriptError(ERR_OUTOFMEM); // Short msg since very rare.
			return;
		}
		strcpy(mReplacement, aReplacement);
	}
	else // Leave mReplacement blank, but make this false so that the hook doesn't do extra work.
		mConformToCase = false;


	mConstructedOK = true; // Done at the very end.
}




void Hotstring::ParseOptions(char *aOptions, int &aPriority, int &aKeyDelay, bool &aCaseSensitive
	, bool &aConformToCase, bool &aDoBackspace, bool &aOmitEndChar, bool &aSendRaw, bool &aEndCharRequired
	, bool &aDetectWhenInsideWord, bool &aDoReset)
{
	// In this case, colon rather than zero marks the end of the string.  However, the string
	// might be empty so check for that too.  In addition, this is now called from
	// IsDirective(), so that's another reason to check for normal string termination.
	char *cp1;
	for (char *cp = aOptions; *cp && *cp != ':'; ++cp)
	{
		cp1 = cp + 1;
		switch(toupper(*cp))
		{
		case '*':
			aEndCharRequired = (*cp1 == '0');
			break;
		case '?':
			aDetectWhenInsideWord = (*cp1 != '0');
			break;
		case 'B': // Do backspacing.
			aDoBackspace = (*cp1 != '0');
			break;
		case 'C':
			if (*cp1 == '0') // restore both settings to default.
			{
				aConformToCase = true;
				aCaseSensitive = false;
			}
			else if (*cp1 == '1')
			{
				aConformToCase = false;
				aCaseSensitive = false;
			}
			else // treat as plain "C"
			{
				aConformToCase = false;  // No point in conforming if its case sensitive.
				aCaseSensitive = true;
			}
			break;
		case 'O':
			aOmitEndChar = (*cp1 != '0');
			break;
		// For options such as K & P: Use atoi() vs. ATOI() to avoid interpreting something like 0x01C
		// as hex when in fact the C was meant to be an option letter:
		case 'K':
			aKeyDelay = atoi(cp + 1);
			break;
		case 'P':
			aPriority = atoi(cp + 1);
			break;
		case 'R':
			aSendRaw = (*cp1 != '0');
			break;
		case 'Z':
			aDoReset = (*cp1 != '0');
			break;
		// Otherwise: Ignore other characters, such as the digits that comprise the number after the P option.
		}
	}
}