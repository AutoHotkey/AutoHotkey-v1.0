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

// In the VC++ project, this file's properties should be set to "exclude this file from build"
// so that it doesn't try to compile it separately.

/*
One of the main objectives of a KeyboardProc or MouseProc (hook) is to minimize the amount of CPU overhead
caused by every input event being handled by the procedure.  One way this is done is to return immediately
on simple conditions that are relatively frequent (such as receiving a key that's not involved in any
hotkey combination).

Another way is to avoid API or system calls that might have a high overhead.  That's why the state of
every prefix key is tracked independently, rather than calling the WinAPI to determine if the
key is actually down at the moment of consideration.
*/


#ifdef INCLUDE_KEYBD_HOOK
	// pEvent is a macro for convenience and readability:
	#undef pEvent
	#define pEvent ((PKBDLLHOOKSTRUCT)lParam)
#else // Mouse Hook:
	#undef pEvent
	#define pEvent ((PMSLLHOOKSTRUCT)lParam)
#endif


// KEY_PHYS_IGNORE events must be mostly ignored because currently there is no way for a given
// hook instance to detect if it sent the event or some other instance.  Therefore, to treat
// such events as true physical events might cause infinite loops or other side-effects in
// the instance that generated the event.  More review of this is needed if KEY_PHYS_IGNORE
// events ever need to be treated as true physical events by the instances of the hook that
// didn't originate them:
#define IS_IGNORED (event.dwExtraInfo == KEY_IGNORE || event.dwExtraInfo == KEY_PHYS_IGNORE \
	|| event.dwExtraInfo == KEY_IGNORE_ALL_EXCEPT_MODIFIER)


#ifdef INCLUDE_KEYBD_HOOK
// Used to help make a workaround for the way the keyboard driver generates physical
// shift-key events to release the shift key whenever it is physically down during
// the press or release of a dual-state Numpad key. These keyboard driver generated
// shift-key events only seem to happen when Numlock is ON, the shift key is logically
// or physically down, and a dual-state numpad key is pressed or released (i.e. the shift
// key might not have been down for the press, but if it's down for the release, the driver
// will suddenly start generating shift events).  I think the purpose of these events is to
// allow the shift keyto temporarily alter the state of the Numlock key for the purpose of
// sending that one key, without the shift key actually being "seen" as down while the key
// itself is sent (since some apps may have special behavior when they detect the shift key
// is down).

// Note: numlock, numpaddiv/mult/sub/add/enter are not affected by this because they have only
// a single state (i.e. they are unaffected by the state of the Numlock key).  Also, these
// driver-generated events occur at a level lower than the hook, so it doesn't matter whether
// the hook suppresses the keys involved (i.e. the shift events still happen anyway).

// So which keys are not physical even though they're non-injected?:
// 1) The shift-up that precedes a down of a dual-state numpad key (only happens when shift key is logically down).
// 2) The shift-down that precedes a pressing down (or releasing in certain very rare cases caused by the
//    exact sequence of keys received) of a key WHILE the numpad key in question is still down.
//    Although this case may seem rare, it's happened to both Robert Yaklin and myself when doing various
//    sorts of hotkeys.
// 3) The shift-up that precedes an up of a dual-state numpad key.  This only happens if the shift key is
//    logically down for any reason at this exact moment, which can be achieved via the send command.
// 4) The shift-down that follows the up of a dual-state numpad key (i.e. the driver is restoring the shift state
//    to what it was before).  This can be either immediate or "lazy".  It's lazy whenever the user had pressed
//    another key while a numpad key was being held down (i.e. case #2 above), in which case the driver waits
//    indefinitely for the user to press any other key and then immediately sneaks in the shift key-down event
//    right before it in the input stream (insertion).
// 5) Similar to #4, but if the driver needs to generate a shift-up for an unexpected Numpad-up event,
//    the restoration of the shift key will be "lazy".  This case was added in response to the below
//    example, wherein the shift key got stuck physically down (incorrectly) by the hook:
// 68  048	 	d	0.00	Num 8          	
// 6B  04E	 	d	0.09	Num +          	
// 68  048	i	d	0.00	Num 8          	
// 68  048	i	u	0.00	Num 8          	
// A0  02A	i	d	0.02	Shift          	part of the macro
// 01  000	i	d	0.03	LButton        	
// A0  02A	 	u	0.00	Shift          	driver, for the next key
// 26  048	 	u	0.00	Num 8          	
// A0  02A	 	d	0.49	Shift          	driver lazy down (but not detected as non-physical)
// 6B  04E	 	d	0.00	Num +          	


// The below timeout is for the subset of driver-generated shift-events that occur immediately
// before or after some other keyboard event.  The elapsed time is usually zero, but using 22ms
// here just in case slower systems or systems under load have longer delays between keystrokes:
#define SHIFT_KEY_WORKAROUND_TIMEOUT 22
static bool pad_state[PAD_TOTAL_COUNT];  // Initialized by ChangeHookState()
static bool next_phys_shift_down_is_not_phys = false;
static vk_type prior_vk = 0;
static sc_type prior_sc = 0;
static bool prior_event_was_key_up = false;
static bool prior_event_was_physical = false;
static DWORD prior_event_tickcount = 0;
static modLR_type prior_modifiersLR_physical = 0;
static BYTE prior_shift_state = 0;  // i.e. default to "key is up".
static BYTE prior_lshift_state = 0;
#endif

	
#ifdef INCLUDE_KEYBD_HOOK
inline bool DualStateNumpadKeyIsDown()
{
	// Note: GetKeyState() might not agree with us that the key is physically down because
	// the hook may have suppressed it (e.g. if it's a hotkey).  Therefore, pad_state
	// is the only way to know for user if the user is physically holding down a *qualified*
	// Numpad key.  "Qualified" means that it must be a dual-state key and Numlock must have
	// been ON at the time the key was first pressed down.  This last criteria is needed because
	// physically holding down the shift-key will change VK generated by the driver to appear
	// to be that of the numpad without the numlock key on.  In other words, we can't just
	// consult the g_PhysicalKeyState array because it won't tell whether a key such as
	// NumpadEnd is truly phyiscally down:
	for (int i = 0; i < PAD_TOTAL_COUNT; ++i)
		if (pad_state[i])
			return true;
	return false;
}



inline IsDualStateNumpadKey(vk_type aVK, sc_type aSC)
{
	if (aSC & 0x100)  // If it's extended, it can't be a numpad key.
		return false;

	switch (aVK)
	{
	// It seems best to exclude the VK_DECIMAL and VK_NUMPAD0 through VK_NUMPAD9 from the below
	// list because the callers want to know whether this is a numpad key being *modified* by
	// the shift key (i.e. the shift key is being held down to temporarily transform the numpad
	// key into its opposite state, overriding the fact that Numlock is ON):
	case VK_DELETE: // NumpadDot (VK_DECIMAL)
	case VK_INSERT: // Numpad0
	case VK_END:    // Numpad1
	case VK_DOWN:   // Numpad2
	case VK_NEXT:   // Numpad3
	case VK_LEFT:   // Numpad4
	case VK_CLEAR:  // Numpad5 (this has been verified to be the VK that is sent, at least on my keyboard).
	case VK_RIGHT:  // Numpad6
	case VK_HOME:   // Numpad7
	case VK_UP:     // Numpad8
	case VK_PRIOR:  // Numpad9
		return true;
	}

	return false;
}
#endif



#ifdef INCLUDE_KEYBD_HOOK
bool EventIsPhysical(KBDLLHOOKSTRUCT &event, vk_type vk, sc_type sc, bool key_up)
// Always use the parameter vk rather than event.vkCode because the caller or caller's caller
// might have adjusted vk, namely to make it a left/right specific modifier key rather than a
// neutral one.
{
	// MSDN: "The keyboard input can come from the local keyboard driver or from calls to the keybd_event
	// function. If the input comes from a call to keybd_event, the input was "injected"".
	// My: This also applies to mouse events, so use it for them too:
	if (event.flags & LLKHF_INJECTED)
		return false;
	// So now we know it's a physical event.  But certain LSHIFT key-down events are driver-generated.
	// We want to be able to tell the difference because the Send command and other aspects
	// of keyboard functionality need us to be accurate about which keys the user is physically
	// holding down at any given time:
	if (   (vk == VK_LSHIFT || vk == VK_SHIFT) && !key_up   ) // But not RSHIFT.
	{
		if (next_phys_shift_down_is_not_phys && !DualStateNumpadKeyIsDown())
		{
			next_phys_shift_down_is_not_phys = false;
			return false;
		}
		// Otherwise (see notes about SHIFT_KEY_WORKAROUND_TIMEOUT above for details):
		if (prior_event_was_key_up && IsDualStateNumpadKey(prior_vk, prior_sc)
			&& (DWORD)(GetTickCount() - prior_event_tickcount) < (DWORD)SHIFT_KEY_WORKAROUND_TIMEOUT   )
			return false;
	}
	// Otherwise, it's physical:
	g_TimeLastInputPhysical = event.time;
	return true;
}

#else // Mouse hook:
inline bool EventIsPhysical(MSLLHOOKSTRUCT &event, bool key_up)
{
	// g_TimeLastInputPhysical is handled elsewhere so that mouse movements are handled too
	// (this function is only ever called for action of the mouse buttons).
	return !(event.flags & LLMHF_INJECTED);
}
#endif



#ifdef INCLUDE_KEYBD_HOOK
void UpdateModifierState(KBDLLHOOKSTRUCT &event, vk_type vk, sc_type sc, bool key_up, bool aIsSuppressed)
// Always use the parameter vk rather than event.vkCode because the caller or caller's caller
// might have adjusted vk, namely to make it a left/right specific modifier key rather than a
// neutral one.
{
	// This part is done even if the key is being ignored because we always want their status
	// to be correct *regardless* of whether the key is ignored.  This is especially important
	// in cases such as Shift-Alt-Tab and Alt-Tab both have substitutes.  NOTE: We don't want
	// the CapsLock/Numlock/Scrolllock section up here with it because for those cases, we
	// genuinely want to ignore them entirely when the hook itself sends a keybd_event for one
	// of them.
	// Since low-level (but not high level) keyboard hook supports left/right VKs, use
	// them in pref. to scan code because it's much more likely to be compatible with
	// non-English or non-std keyboards.

	// Below excludes KEY_IGNORE_ALL_EXCEPT_MODIFIER since that type of event should not
	// be ignored by this function.  UPDATE: KEY_PHYS_IGNORE is now considered to be something
	// that should not be ignored because if more than one instance has the hook installed,
	// it is possible for g_modifiersLR_logical_non_ignored to say that a key is down in one
	// instance when that instance's g_modifiersLR_logical doesn't say it's down, which is
	// definitely wrong.  So is now omitted from the below:
	bool is_not_ignored = event.dwExtraInfo != KEY_IGNORE;

	switch (vk)
	{
	// Normally (for physical key presses) the vk will be left/right specific.  However,
	// if another app calls keybd_event() or a similar function to inject input,
	// the generic key will be received if that's what was sent.
	// Try to keep the most often-pressed keys at the top for potentially better
	// performance (depends on how the compiler implements the switch stmt: jump table
	// vs. if-then-else tree):
	case VK_LSHIFT:
		if (key_up)
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical &= ~MOD_LSHIFT;
				// Even if is_not_ignored == true, this is updated unconditionally on key-up events
				// to ensure that g_modifiersLR_logical_non_ignored never says a key is down when
				// g_modifiersLR_logical says its up, which might otherwise happen in cases such
				// as alt-tab.  See this comment further below, where the operative word is "relied":
				// "key pushed ALT down, or relied upon it already being down, so go up".  UPDATE:
				// The above is no longer a concern because KeyEvent() now defaults to the mode
				// which causes our var "is_not_ignored" to be true here.  Only the Send command
				// overrides this default, and it takes responsibility for ensuring that the older
				// comment above never happens by forcing any down-modifiers to be up if they're
				// not logically down as reflected in g_modifiersLR_logical.  There's more
				// explanation for g_modifiersLR_logical_non_ignored in keyboard.h:
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored &= ~MOD_LSHIFT;
			}
			if (EventIsPhysical(event, vk, sc, key_up)) // Note that ignored events can be physical via KEYEVENT_PHYS()
			{
				g_modifiersLR_physical &= ~MOD_LSHIFT;
				g_PhysicalKeyState[VK_LSHIFT] = 0;
				g_PhysicalKeyState[VK_SHIFT] = g_PhysicalKeyState[VK_RSHIFT];  // Neutral is down if right is down.
			}
		}
		else
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical |= MOD_LSHIFT;
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored |= MOD_LSHIFT;
			}
			if (EventIsPhysical(event, vk, sc, key_up))
			{
				g_modifiersLR_physical |= MOD_LSHIFT;
				// Neutral key is defined as being down if either L/R is down:
				g_PhysicalKeyState[VK_LSHIFT] = g_PhysicalKeyState[VK_SHIFT] = STATE_DOWN;
			}
		}
		break;
	case VK_RSHIFT:
		if (key_up)
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical &= ~MOD_RSHIFT;
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored &= ~MOD_RSHIFT;
			}
			if (EventIsPhysical(event, vk, sc, key_up))
			{
				g_modifiersLR_physical &= ~MOD_RSHIFT;
				g_PhysicalKeyState[VK_RSHIFT] = 0;
				g_PhysicalKeyState[VK_SHIFT] = g_PhysicalKeyState[VK_LSHIFT];
			}
		}
		else
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical |= MOD_RSHIFT;
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored |= MOD_RSHIFT;
			}
			if (EventIsPhysical(event, vk, sc, key_up))
			{
				g_modifiersLR_physical |= MOD_RSHIFT;
				g_PhysicalKeyState[VK_RSHIFT] = g_PhysicalKeyState[VK_SHIFT] = STATE_DOWN;
			}
		}
		break;
	case VK_LCONTROL:
		if (key_up)
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical &= ~MOD_LCONTROL;
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored &= ~MOD_LCONTROL;
			}
			if (EventIsPhysical(event, vk, sc, key_up))
			{
				g_modifiersLR_physical &= ~MOD_LCONTROL;
				g_PhysicalKeyState[VK_LCONTROL] = 0;
				g_PhysicalKeyState[VK_CONTROL] = g_PhysicalKeyState[VK_RCONTROL];
			}
		}
		else
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical |= MOD_LCONTROL;
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored |= MOD_LCONTROL;
			}
			if (EventIsPhysical(event, vk, sc, key_up))
			{
				g_modifiersLR_physical |= MOD_LCONTROL;
				g_PhysicalKeyState[VK_LCONTROL] = g_PhysicalKeyState[VK_CONTROL] = STATE_DOWN;
			}
		}
		break;
	case VK_RCONTROL:
		if (key_up)
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical &= ~MOD_RCONTROL;
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored &= ~MOD_RCONTROL;
			}
			if (EventIsPhysical(event, vk, sc, key_up))
			{
				g_modifiersLR_physical &= ~MOD_RCONTROL;
				g_PhysicalKeyState[VK_RCONTROL] = 0;
				g_PhysicalKeyState[VK_CONTROL] = g_PhysicalKeyState[VK_LCONTROL];
			}
		}
		else
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical |= MOD_RCONTROL;
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored |= MOD_RCONTROL;
			}
			if (EventIsPhysical(event, vk, sc, key_up))
			{
				g_modifiersLR_physical |= MOD_RCONTROL;
				g_PhysicalKeyState[VK_RCONTROL] = g_PhysicalKeyState[VK_CONTROL] = STATE_DOWN;
			}
		}
		break;
	case VK_LMENU:
		if (key_up)
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical &= ~MOD_LALT;
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored &= ~MOD_LALT;
			}
			if (EventIsPhysical(event, vk, sc, key_up))
			{
				g_modifiersLR_physical &= ~MOD_LALT;
				g_PhysicalKeyState[VK_LMENU] = 0;
				g_PhysicalKeyState[VK_MENU] = g_PhysicalKeyState[VK_RMENU];
			}
		}
		else
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical |= MOD_LALT;
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored |= MOD_LALT;
			}
			if (EventIsPhysical(event, vk, sc, key_up))
			{
				g_modifiersLR_physical |= MOD_LALT;
				g_PhysicalKeyState[VK_LMENU] = g_PhysicalKeyState[VK_MENU] = STATE_DOWN;
			}
		}
		break;
	case VK_RMENU:
		if (key_up)
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical &= ~MOD_RALT;
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored &= ~MOD_RALT;
			}
			if (EventIsPhysical(event, vk, sc, key_up))
			{
				g_modifiersLR_physical &= ~MOD_RALT;
				g_PhysicalKeyState[VK_RMENU] = 0;
				g_PhysicalKeyState[VK_MENU] = g_PhysicalKeyState[VK_LMENU];
			}
		}
		else
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical |= MOD_RALT;
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored |= MOD_RALT;
			}
			if (EventIsPhysical(event, vk, sc, key_up))
			{
				g_modifiersLR_physical |= MOD_RALT;
				g_PhysicalKeyState[VK_RMENU] = g_PhysicalKeyState[VK_MENU] = STATE_DOWN;
			}
		}
		break;
	case VK_LWIN:
		if (key_up)
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical &= ~MOD_LWIN;
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored &= ~MOD_LWIN;
			}
			if (EventIsPhysical(event, vk, sc, key_up))
			{
				g_modifiersLR_physical &= ~MOD_LWIN;
				g_PhysicalKeyState[VK_LWIN] = 0;
			}
		}
		else
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical |= MOD_LWIN;
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored |= MOD_LWIN;
			}
			if (EventIsPhysical(event, vk, sc, key_up))
			{
				g_modifiersLR_physical |= MOD_LWIN;
				g_PhysicalKeyState[VK_LWIN] = STATE_DOWN;
			}
		}
		break;
	case VK_RWIN:
		if (key_up)
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical &= ~MOD_RWIN;
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored &= ~MOD_RWIN;
			}
			if (EventIsPhysical(event, vk, sc, key_up))
			{
				g_modifiersLR_physical &= ~MOD_RWIN;
				g_PhysicalKeyState[VK_RWIN] = 0;
			}
		}
		else
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical |= MOD_RWIN;
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored |= MOD_RWIN;
			}
			if (EventIsPhysical(event, vk, sc, key_up))
			{
				g_modifiersLR_physical |= MOD_RWIN;
				g_PhysicalKeyState[VK_RWIN] = STATE_DOWN;
			}
		}
		break;

	// This should rarely if ever occur under WinNT/2k/XP -- perhaps only if an app calls keybd_event()
	// and explicitly specifies one of these VKs to be sent.  UPDATE: THIS DOES HAPPEN ON WINDOWS NT.
	// So apparently, NT works with the neutral VKs for modifiers rather than the left/right specific
	// ones like 2k/XP.  UPDATE: This section is no longer needed because the keyboard hook translates
	// neutral modifier events into left/right-specific events early on (before we're ever called).
	// ALSO NOTE: The below has not yet been changed to keep g_modifiersLR_logical_non_ignored updated.
	//case VK_SHIFT:
	//	if (sc == SC_RSHIFT)
	//		if (key_up)
	//		{
	//			if (!aIsSuppressed)
	//				g_modifiersLR_logical &= ~MOD_RSHIFT;
	//			if (EventIsPhysical(event, vk, sc, key_up))
	//			{
	//				g_modifiersLR_physical &= ~MOD_RSHIFT;
	//				g_PhysicalKeyState[VK_RSHIFT] = 0;
	//				g_PhysicalKeyState[VK_SHIFT] = g_PhysicalKeyState[VK_LSHIFT];
	//			}
	//		}
	//		else
	//		{
	//			if (!aIsSuppressed)
	//				g_modifiersLR_logical |= MOD_RSHIFT;
	//			if (EventIsPhysical(event, vk, sc, key_up))
	//			{
	//				g_modifiersLR_physical |= MOD_RSHIFT;
	//				g_PhysicalKeyState[VK_RSHIFT] = g_PhysicalKeyState[VK_SHIFT] = STATE_DOWN;
	//			}
	//		}
	//	else // Assume the left even if scan code doesn't match what would be expected.
	//	// Else even if it's not SC_LSHIFT, assume that it's the left-shift key anyway
	//	// (since one of them has to be the event, have to choose one):
	//		if (key_up)
	//		{
	//			if (!aIsSuppressed)
	//				g_modifiersLR_logical &= ~MOD_LSHIFT;
	//			if (EventIsPhysical(event, vk, sc, key_up))
	//			{
	//				g_modifiersLR_physical &= ~MOD_LSHIFT;
	//				g_PhysicalKeyState[VK_LSHIFT] = 0;
	//				g_PhysicalKeyState[VK_SHIFT] = g_PhysicalKeyState[VK_RSHIFT];
	//			}
	//		}
	//		else
	//		{
	//			if (!aIsSuppressed)
	//				g_modifiersLR_logical |= MOD_LSHIFT;
	//			if (EventIsPhysical(event, vk, sc, key_up))
	//			{
	//				g_modifiersLR_physical |= MOD_LSHIFT;
	//				g_PhysicalKeyState[VK_LSHIFT] = g_PhysicalKeyState[VK_SHIFT] = STATE_DOWN;
	//			}
	//		}
	//	break;
	//case VK_CONTROL:
	//	if (sc == SC_RCONTROL)
	//		if (key_up)
	//		{
	//			if (!aIsSuppressed)
	//				g_modifiersLR_logical &= ~MOD_RCONTROL;
	//			if (EventIsPhysical(event, vk, sc, key_up))
	//			{
	//				g_modifiersLR_physical &= ~MOD_RCONTROL;
	//				g_PhysicalKeyState[VK_RCONTROL] = 0;
	//				g_PhysicalKeyState[VK_CONTROL] = g_PhysicalKeyState[VK_LCONTROL];
	//			}
	//		}
	//		else
	//		{
	//			if (!aIsSuppressed)
	//				g_modifiersLR_logical |= MOD_RCONTROL;
	//			if (EventIsPhysical(event, vk, sc, key_up))
	//			{
	//				g_modifiersLR_physical |= MOD_RCONTROL;
	//				g_PhysicalKeyState[VK_RCONTROL] = g_PhysicalKeyState[VK_CONTROL] = STATE_DOWN;
	//			}
	//		}
	//	else // Assume the left even if scan code doesn't match what would be expected.
	//		if (key_up)
	//		{
	//			if (!aIsSuppressed)
	//				g_modifiersLR_logical &= ~MOD_LCONTROL;
	//			if (EventIsPhysical(event, vk, sc, key_up))
	//			{
	//				g_modifiersLR_physical &= ~MOD_LCONTROL;
	//				g_PhysicalKeyState[VK_LCONTROL] = 0;
	//				g_PhysicalKeyState[VK_CONTROL] = g_PhysicalKeyState[VK_RCONTROL];
	//			}
	//		}
	//		else
	//		{
	//			if (!aIsSuppressed)
	//				g_modifiersLR_logical |= MOD_LCONTROL;
	//			if (EventIsPhysical(event, vk, sc, key_up))
	//			{
	//				g_modifiersLR_physical |= MOD_LCONTROL;
	//				g_PhysicalKeyState[VK_LCONTROL] = g_PhysicalKeyState[VK_CONTROL] = STATE_DOWN;
	//			}
	//		}
	//	break;
	//case VK_MENU:
	//	if (sc == SC_RALT)
	//		if (key_up)
	//		{
	//			if (!aIsSuppressed)
	//				g_modifiersLR_logical &= ~MOD_RALT;
	//			if (EventIsPhysical(event, vk, sc, key_up))
	//			{
	//				g_modifiersLR_physical &= ~MOD_RALT;
	//				g_PhysicalKeyState[VK_RMENU] = 0;
	//				g_PhysicalKeyState[VK_MENU] = g_PhysicalKeyState[VK_LMENU];
	//			}
	//		}
	//		else
	//		{
	//			if (!aIsSuppressed)
	//				g_modifiersLR_logical |= MOD_RALT;
	//			if (EventIsPhysical(event, vk, sc, key_up))
	//			{
	//				g_modifiersLR_physical |= MOD_RALT;
	//				g_PhysicalKeyState[VK_RMENU] = g_PhysicalKeyState[VK_MENU] = STATE_DOWN;
	//			}
	//		}
	//	else // Assume the left even if scan code doesn't match what would be expected.
	//		if (key_up)
	//		{
	//			if (!aIsSuppressed)
	//				g_modifiersLR_logical &= ~MOD_LALT;
	//			if (EventIsPhysical(event, vk, sc, key_up))
	//			{
	//				g_modifiersLR_physical &= ~MOD_LALT;
	//				g_PhysicalKeyState[VK_LMENU] = 0;
	//				g_PhysicalKeyState[VK_MENU] = g_PhysicalKeyState[VK_RMENU];
	//			}
	//		}
	//		else
	//		{
	//			if (!aIsSuppressed)
	//				g_modifiersLR_logical |= MOD_LALT;
	//			if (EventIsPhysical(event, vk, sc, key_up))
	//			{
	//				g_modifiersLR_physical |= MOD_LALT;
	//				g_PhysicalKeyState[VK_LMENU] = g_PhysicalKeyState[VK_MENU] = STATE_DOWN;
	//			}
	//		}
	//	break;
	}
}



void UpdateKeyState(KBDLLHOOKSTRUCT &event, vk_type vk, sc_type sc, bool key_up, bool aIsSuppressed)
// Always use the parameter vk rather than event.vkCode because the caller or caller's caller
// might have adjusted vk, namely to make it a left/right specific modifier key rather than a
// neutral one.
{
	// See above notes near the first mention of SHIFT_KEY_WORKAROUND_TIMEOUT for details.
	// This part of the workaround can be tested via "NumpadEnd::KeyHistory".  Turn on numlock,
	// hold down shift, and press numpad1. The hotkey will fire and the status should display
	// that the shift key is physically, but not logically down at that exact moment:
	if (prior_event_was_physical && (prior_vk == VK_LSHIFT || prior_vk == VK_SHIFT)  // But not RSHIFT.
		&& (DWORD)(GetTickCount() - prior_event_tickcount) < (DWORD)SHIFT_KEY_WORKAROUND_TIMEOUT)
	{
		bool current_is_dual_state = IsDualStateNumpadKey(vk, sc);
		// Verified: Both down and up events for the *current* (not prior) key qualify for this:
		bool fix_it = (!prior_event_was_key_up && DualStateNumpadKeyIsDown())  // Case #4 of the workaround.
			|| (prior_event_was_key_up && key_up && current_is_dual_state); // Case #5
		if (fix_it)
			next_phys_shift_down_is_not_phys = true;
		// In the first case, both the numpad key-up and down events are eligible:
		if (   fix_it || (prior_event_was_key_up && current_is_dual_state)   )
		{
			// Since the prior event (the shift key) already happened (took effect) and since only
			// now is it known that it shouldn't have been physical, undo the effects of it having
			// been physical:
			g_modifiersLR_physical = prior_modifiersLR_physical;
			g_PhysicalKeyState[VK_SHIFT] = prior_shift_state;
			g_PhysicalKeyState[VK_LSHIFT] = prior_lshift_state;
		}
	}


	// Must do this part prior to UpdateModifierState() because we want to store the values
	// as they are prior to the potentially-erroneously-physical shift key event takes effect.
	// The state of these is also saved because we can't assume that a shift-down, for
	// example, CHANGED the state to down, because it may have been already down before that:
	prior_modifiersLR_physical = g_modifiersLR_physical;
	prior_shift_state = g_PhysicalKeyState[VK_SHIFT];
	prior_lshift_state = g_PhysicalKeyState[VK_LSHIFT];

	// If this function was called from SuppressThisKey(), these comments apply:
	// Currently SuppressThisKey is only called with a modifier in the rare case
	// when disguise_next_lwin/rwin_up is in effect.  But there may be other cases in the
	// future, so we need to make sure the physical state of the modifiers is updated
	// in our tracking system even though the key is being suppressed:
	if (kvk[vk].as_modifiersLR)
		UpdateModifierState(event, vk, sc, key_up, aIsSuppressed);  // Update our tracking of LWIN/RWIN/RSHIFT etc.

	// Now that we're done using the old values (the above used them and also UpdateModifierState()'s
	// calls to EventIsPhysical()), update these to their new values:
	prior_vk = vk;
	prior_sc = sc;
	prior_event_was_key_up = key_up;
	prior_event_was_physical = EventIsPhysical(event, vk, sc, key_up);
	prior_event_tickcount = GetTickCount();
}
#endif // Keyboard hook



#ifdef INCLUDE_KEYBD_HOOK
	#undef SuppressThisKey
	#define SuppressThisKey SuppressThisKeyFunc(event, vk, sc, key_up, pKeyHistoryCurr)
	LRESULT SuppressThisKeyFunc(KBDLLHOOKSTRUCT &event, vk_type vk, sc_type sc, bool key_up, KeyHistoryItem *pKeyHistoryCurr)
	// Always use the parameter vk rather than event.vkCode because the caller or caller's caller
	// might have adjusted vk, namely to make it a left/right specific modifier key rather than a
	// neutral one.
#else // Mouse Hook:
	#undef SuppressThisKey
	#define SuppressThisKey SuppressThisKeyFunc(pKeyHistoryCurr)
	LRESULT SuppressThisKeyFunc(KeyHistoryItem *pKeyHistoryCurr)
#endif
{
	if (pKeyHistoryCurr->event_type == ' ') // then it hasn't been already set somewhere else
		pKeyHistoryCurr->event_type = 's';
	// This handles the troublesome Numlock key, which on some (most/all?) keyboards
	// will change state independent of the keyboard's indicator light even if its
	// keydown and up events are suppressed.  This is certainly true on the
	// MS Natural Elite keyboard using default drivers on WinXP.  SetKeyboardState()
	// doesn't resolve this, so the only alternative to the below is to use the
	// Win9x method of setting the Numlock state explicitly whenever the key is released.
	// That might be complicated by the fact that the unexpected state change described
	// here can't be detected by GetKeyboardState() and such (it sees the state indicated
	// by the numlock light on the keyboard, which is wrong).  In addition, doing it this
	// way allows Numlock to be a prefix key for something like Numpad7, which would
	// otherwise be impossible because Numpad7 would become NumpadHome the moment
	// Numlock was pressed down.  Note: this problem doesn't appear to affect Capslock
	// or Scrolllock for some reason, possibly hardware or driver related.
	// Note: the check for KEY_IGNORE isn't strictly necessary, but here just for safety
	// in case this is ever called for a key that should be ignored.  If that were
	// to happen and we didn't check for it, and endless loop of keyboard events
	// might be caused due to the keybd events sent below.
#ifdef INCLUDE_KEYBD_HOOK
	if (vk == VK_NUMLOCK && !key_up && !IS_IGNORED)
	{
		// This seems to undo the faulty indicator light problem and toggle
		// the key back to the state it was in prior to when the user pressed it.
		// Originally, I had two keydowns and before that some keyups too, but
		// testing reveals that only a single key-down is needed.  UPDATE:
		// It appears that all 4 of these key events are needed to make it work
		// in every situation, especially the case when ForceNumlock is on but
		// numlock isn't used for any hotkeys.
		// Note: The only side-effect I've discovered of this method is that the
		// indicator light can't be toggled after the program is exitted unless the
		// key is pressed twice:
		KeyEvent(KEYUP, VK_NUMLOCK);
		KeyEvent(KEYDOWNANDUP, VK_NUMLOCK);
		KeyEvent(KEYDOWN, VK_NUMLOCK);
	}
	UpdateKeyState(event, vk, sc, key_up, true);
#endif

	// Use PostMessage() rather than directly calling the function to write the key to
	// the log file.  This is done so that we can return sooner, which reduces the keyboard
	// and mouse lag that would otherwise be caused by us not being in a pumping-messages state.
	// IN ADDITION, this method may also help to keep the keystrokes in order (due to recursive
	// calls to the hook via keybd_event(), etc), I'm not sure.  UPDATE: No, that seems impossible
	// when you think about it.  Also, to simplify this the function is now called directly rather
	// than posting a message to avoid complications that might be caused by the script being
	// uninterruptible for a long period (rare), which would thus cause the posted msg to stay
	// buffered for a long time:
#ifdef ENABLE_KEY_HISTORY_FILE
	if (g_KeyHistoryToFile && pKeyHistoryCurr)
		KeyHistoryToFile(NULL, pKeyHistoryCurr->event_type, pKeyHistoryCurr->key_up
			, pKeyHistoryCurr->vk, pKeyHistoryCurr->sc);  // A fairly low overhead operation.
#endif

	return 1;
}



#ifdef INCLUDE_KEYBD_HOOK
#define shs Hotstring::shs  // For convenience.
inline bool CollectInput(KBDLLHOOKSTRUCT &event, vk_type vk, sc_type sc, bool key_up, bool is_ignored)
// Returns true if the caller should treat the key as visible (non-suppressed).
// Always use the parameter vk rather than event.vkCode because the caller or caller's caller
// might have adjusted vk, namely to make it a left/right specific modifier key rather than a
// neutral one.
{
	// Generally, we return this value to our caller so that it will treat the event as visible
	// if either there's no input in progress or if there is but it's visible.  Below relies on
	// boolean evaluation order:
	bool treat_as_visible = g_input.status != INPUT_IN_PROGRESS || g_input.Visible
		|| kvk[vk].pForceToggle;  // Never suppress toggleable keys such as CapsLock.

	if (key_up)
		// Always pass modifier-up events through unaltered.  At the very least, this is needed for
		// cases where a user presses a #z hotkey, for example, to initiate an Input.  When the user
		// releases the LWIN/RWIN key during the input, that up-event should not be suppressed
		// otherwise the modifier key would get "stuck down".  
		return kvk[vk].as_modifiersLR ? true : treat_as_visible;

	// Hotstrings monitor neither ignored input nor input that is invisible due to suppression by
	// the Input command.  One reason for not monitoring ignored input is to avoid any chance of
	// an infinite loop of keystrokes caused by one hotstring triggering itself directly or
	// indirectly via a different hotstring:
	bool do_monitor_hotstring = shs && !is_ignored && treat_as_visible;
	bool do_input = g_input.status == INPUT_IN_PROGRESS && !(g_input.IgnoreAHKInput && is_ignored);

	UCHAR end_key_attributes;
	if (do_input)
	{
		end_key_attributes = g_input.EndVK[vk];
		if (!end_key_attributes)
			end_key_attributes = g_input.EndSC[sc];
		if (end_key_attributes) // A terminating keystroke has now occurred unless the shift state isn't right.
		{
			// Caller has ensured that only one of the flags below is set (if any):
			bool shift_must_be_down = end_key_attributes & END_KEY_WITH_SHIFT;
			bool shift_must_not_be_down = end_key_attributes & END_KEY_WITHOUT_SHIFT;
			bool shift_state_matters = shift_must_be_down && !shift_must_not_be_down
				|| !shift_must_be_down && shift_must_not_be_down; // i.e. exactly one of them.
			if (    !shift_state_matters
				|| (shift_must_be_down && (g_modifiersLR_logical & (MOD_LSHIFT | MOD_RSHIFT)))
				|| (shift_must_not_be_down && !(g_modifiersLR_logical & (MOD_LSHIFT | MOD_RSHIFT)))   )
			{
				// The shift state is correct to produce the desired end-key.
				g_input.status = INPUT_TERMINATED_BY_ENDKEY;
				g_input.EndedBySC = g_input.EndSC[sc];
				g_input.EndingVK = vk;
				g_input.EndingSC = sc;
				// Don't change this line:
				g_input.EndingRequiredShift = shift_must_be_down && (g_modifiersLR_logical & (MOD_LSHIFT | MOD_RSHIFT));
				if (!do_monitor_hotstring)
					return treat_as_visible;
				// else need to return only after the input is collected for the hotstring.
			}
		}
	}

	// Reset hotstring detection if the user seems to be navigating within an editor.  This is done
	// so that hotstrings do not fire in unexpected places.
	if (do_monitor_hotstring && g_HSBufLength)
	{
		switch (vk)
		{
		case VK_LEFT:
		case VK_RIGHT:
		case VK_DOWN:
		case VK_UP:
		case VK_NEXT:
		case VK_PRIOR:
		case VK_HOME:
		case VK_END:
			*g_HSBuf = '\0';
			g_HSBufLength = 0;
			break;
		}
	}

	// Don't unconditionally transcribe modified keys such as Ctrl-C because calling ToAsciiEx() on
	// some such keys (e.g. Ctrl-LeftArrow or RightArrow if I recall correctly), disrupts the native
	// function of those keys.  That is the reason for the existence of the
	// g_input.TranscribeModifiedKeys option.
	if (g_modifiersLR_physical
		&& !(g_input.status == INPUT_IN_PROGRESS && g_input.TranscribeModifiedKeys)
		&& g_modifiersLR_physical != MOD_LSHIFT && g_modifiersLR_physical != MOD_RSHIFT
		&& g_modifiersLR_physical != (MOD_LSHIFT & MOD_RSHIFT)
		&& !((g_modifiersLR_physical & (MOD_LALT | MOD_RALT)) && (g_modifiersLR_physical & (MOD_LCONTROL | MOD_RCONTROL))))
		// Since in some keybd layouts, AltGr (Ctrl+Alt) will produce valid characters (such as the @ symbol,
		// which is Ctrl+Alt+Q in the German/IBM layout and Ctrl+Alt+2 in the Spanish layout), an attempt
		// will now be made to transcribe all of the following modifier combinations:
		// - Anything with no modifiers at all.
		// - Anything that uses ONLY the shift key.
		// - Anything with Ctrl+Alt together in it, including Ctrl+Alt+Shift, etc. -- but don't do
		//   "anything containing the Alt key" because that causes weird side-effects with
		//   Alt+LeftArrow/RightArrow and maybe other keys too).
		// Older/Obsolete: If any modifiers except SHIFT are physically down, don't transcribe the key since
		// most users wouldn't want that.  An additional benefit of this policy is that registered hotkeys will
		// normally be excluded from the input (except those rare ones that have only SHIFT as a modifier).
		// Note that ToAscii() will translate ^i to a tab character, !i to plain i, and many other modified
		// letters as just the plain letter key, which we don't want.
		return treat_as_visible;

	static vk_type pending_dead_key_vk = 0;
	static sc_type pending_dead_key_sc = 0; // Need to track this separately because sometimes default mapping isn't correct.
	static bool pending_dead_key_used_shift = false;

	// v1.0.21: Only true (unmodified) backspaces are recognized by the below.  Another reason to do
	// this is that ^backspace has a native function (delete word) different than backspace in many editors.
	if (vk == VK_BACK && !g_modifiersLR_physical) // Backspace
	{
		// Note that it might have been in progress upon entry to this function but now isn't due to
		// INPUT_TERMINATED_BY_ENDKEY above:
		if (do_input && g_input.status == INPUT_IN_PROGRESS && g_input.BackspaceIsUndo) // Backspace is being used as an Undo key.
			if (g_input.BufferLength)
				g_input.buffer[--g_input.BufferLength] = '\0';
		if (do_monitor_hotstring && g_HSBufLength)
			g_HSBuf[--g_HSBufLength] = '\0';
		if (pending_dead_key_vk) // Doing this produces the expected behavior when a backspace occurs immediately after a dead key.
			pending_dead_key_vk = 0;
		return treat_as_visible;
	}

	BYTE ch[3], key_state[256];
	memcpy(key_state, g_PhysicalKeyState, 256);
	// As of v1.0.25.10, the below fixes the Input command so that when it is capturing artificial input,
	// such as from the Send command or a hotstring's replacement text, the captured input will reflect
	// any modifiers that are logically but not physically down:
	AdjustKeyState(key_state, g_modifiersLR_logical);
	// Make the state of capslock accurate so that ToAscii() will return upper vs. lower if appropriate:
	if (IsKeyToggledOn(VK_CAPITAL))
		key_state[VK_CAPITAL] |= STATE_ON;
	else
		key_state[VK_CAPITAL] &= ~STATE_ON;

	// Use ToAsciiEx() vs. ToAscii() because there is evidence from Putty author that
	// ToAsciiEx() works better with more keyboard layouts under 2k/XP than ToAscii()
	// does (though if true, there is no MS explanation).
	int byte_count = ToAsciiEx(vk, event.scanCode  // Uses the original scan code, not the adjusted "sc" one.
		, key_state, (LPWORD)ch, g_MenuIsVisible ? 1 : 0
		, GetKeyboardLayout(0)); // Fetch layout every time in case it changes while the program is running.
	if (!byte_count) // No translation for this key.
		return treat_as_visible;

	// More notes about dead keys: The dead key behavior of Enter/Space/Backspace is already properly
	// maintained when an Input or hotstring monitoring is in effect.  In addition, keys such as the
	// following already work okay (i.e. the user can press them in between the pressing of a dead
	// key and it's finishing/base/trigger key without disrupting the production of diacritic letters)
	// because ToAsciiEx() finds no translation-to-char for them:
	// pgup/dn/home/end/ins/del/arrowkeys/f1-f24/etc.
	// Note that if a pending dead key is followed by the press of another dead key (including itself),
	// the sequence should be triggered and both keystrokes should appear in the active window.
	// That case has been tested too, and works okay with the layouts tested so far.
	// I've only discovered two keys which need special handling, and they are handled below:
	// VK_TAB & VK_ESCAPE
	// These keys have an ascii translation but should not trigger/complete a pending dead key,
	// at least not on the Spanish and Danish layouts, which are the two I've tested so far.

	// Dead keys in Danish layout as they appear on a US English keyboard: Equals and Plus /
	// Right bracket & Brace / probably others.

	// It's not a dead key, but if there's a dead key pending and this incoming key is capable of
	// completing/triggering it, do a workaround for the side-effects of ToAsciiEx().  This workaround
	// allows dead keys to continue to operate properly in the user's foreground window, while still
	// being capturable by the Input command and recognizable by any defined hotstrings whose
	// abbreviations use diacritic letters:
	if (pending_dead_key_vk && vk != VK_TAB && vk != VK_ESCAPE)
	{
		vk_type vk_to_send = pending_dead_key_vk;
		pending_dead_key_vk = 0; // First reset this because below results in a recursive call to keyboard hook.
		// If there's an Input in progress and it's invisible, the foreground app won't see the keystrokes,
		// thus no need to re-insert the dead key into the keyboard buffer.  Note that the Input might have
		// been in progress upon entry to this function but now isn't due to INPUT_TERMINATED_BY_ENDKEY above.
		if (treat_as_visible)
		{
			// Tell the recursively called next instance of the keyboard hook not do the following for
			// the below KEYEVENT_PHYS: Do not call ToAsciiEx() on it and do not capture it as part of
			// the Input itself.  Although this is only needed for the case where the statement
			// "(do_input && g_input.status == INPUT_IN_PROGRESS && !g_input.IgnoreAHKInput)" is true
			// (since hotstrings don't capture/monitor AHK-generated input), it's simpler and about the
			// same in performance to do it unconditonally:
			vk_to_ignore_next_time_down = vk_to_send;
			// Ensure the correct shift-state is set for the below event.  The correct shift key (left or
			// right) must be used to prevent sticking keys and other side-effects:
			vk_type which_shift_down = 0;
			if (g_modifiersLR_logical & MOD_LSHIFT)
				which_shift_down = VK_LSHIFT;
			else if (g_modifiersLR_logical & MOD_RSHIFT)
				which_shift_down = VK_RSHIFT;
			vk_type which_shift_to_send = which_shift_down ? which_shift_down : VK_LSHIFT;
			if (pending_dead_key_used_shift != (bool)which_shift_down)
				KeyEvent(pending_dead_key_used_shift ? KEYDOWN : KEYUP, which_shift_to_send);
			// Since it's a substitute for the previously suppressed physical dead key event, mark it as physical:
			KEYEVENT_PHYS(KEYDOWNANDUP, vk_to_send, pending_dead_key_sc);
			if (pending_dead_key_used_shift != (bool)which_shift_down) // Restore the original shift state.
				KeyEvent(pending_dead_key_used_shift ? KEYUP : KEYDOWN, which_shift_to_send);
		}
	}
	else if (byte_count < 0) // It's a dead key not already handled by the above (i.e. that does not immediately follow a pending dead key).
	{
		if (treat_as_visible)
		{
			pending_dead_key_vk = vk;
			pending_dead_key_sc = sc;
			pending_dead_key_used_shift = g_modifiersLR_logical & (MOD_LSHIFT | MOD_RSHIFT);
		}
		// Dead keys must always be hidden, otherwise they would be shown twice literally due to
		// having been "damaged" by ToAsciiEx():
		return false;
	}

	if (ch[0] == '\r')  // Translate \r to \n since \n is more typical and useful in Windows.
		ch[0] = '\n';
	if (ch[1] == '\r')  // But it's never referred to if byte_count < 2
		ch[1] = '\n';

	bool suppress_hotstring_final_char = false; // Set default.

	if (do_monitor_hotstring)
	{
		HWND fore = GetForegroundWindow();
		if (fore != g_HShwnd)
		{
			// Since the buffer tends to correspond to the text to the left of the caret in the
			// active window, if the active window changes, it seems best to reset the buffer
			// to avoid misfires.
			g_HShwnd = fore;
			*g_HSBuf = '\0';
			g_HSBufLength = 0;
		}
		else if (HS_BUF_SIZE - g_HSBufLength < 3) // Not enough room for up-to-2 chars.
		{
			// Make room in buffer by removing chars from the front that are no longer needed for HS detection:
			// Bug-fixed the below for v1.0.21:
			g_HSBufLength = (int)strlen(g_HSBuf + HS_BUF_DELETE_COUNT);  // The new length.
			memmove(g_HSBuf, g_HSBuf + HS_BUF_DELETE_COUNT, g_HSBufLength + 1); // +1 to include the zero terminator.
		}

		g_HSBuf[g_HSBufLength++] = ch[0];
		if (byte_count > 1)
			// MSDN: "This usually happens when a dead-key character (accent or diacritic) stored in the
			// keyboard layout cannot be composed with the specified virtual key to form a single character."
			g_HSBuf[g_HSBufLength++] = ch[1];
		g_HSBuf[g_HSBufLength] = '\0';

		if (g_HSBufLength)
		{
			char *cphs, *cpbuf, *cpcase_start, *cpcase_end;
			int characters_with_case;
			bool first_char_with_case_is_upper, first_char_with_case_has_gone_by;
			CaseConformModes case_conform_mode;

			// Searching through the hot strings in the original, physical order is the documented
			// way in which precedence is determined, i.e. the first match is the only one that will
			// be triggered.
			for (HotstringIDType u = 0; u < Hotstring::sHotstringCount; ++u)
			{
				Hotstring &hs = *shs[u];  // For performance and convenience.
				if (hs.mSuspended)
					continue;
				if (hs.mEndCharRequired)
				{
					if (g_HSBufLength <= hs.mStringLength) // Ensure the string is long enough for loop below.
						continue;
					if (!strchr(g_EndChars, g_HSBuf[g_HSBufLength - 1])) // It's not an end-char, so no match.
						continue;
					cpbuf = g_HSBuf + g_HSBufLength - 2; // Init once for both loops. -2 to omit end-char.
				}
				else // No ending char required.
				{
					if (g_HSBufLength < hs.mStringLength) // Ensure the string is long enough for loop below.
						continue;
					cpbuf = g_HSBuf + g_HSBufLength - 1; // Init once for both loops.
				}
				cphs = hs.mString + hs.mStringLength - 1; // Init once for both loops.
				// Check if this item is a match:
				if (hs.mCaseSensitive)
				{
					for (; cphs >= hs.mString; --cpbuf, --cphs)
						if (*cpbuf != *cphs)
							break;
				}
				else // case insensitive
					// use toupper() vs. CharUpper() for consistency with Input, IfInString, etc.
					// Update: To support hotstrings such as the following without being case
					// sensitive, it seems best to use CharUpper() instead, e.g.
					// (char)CharUpper((LPTSTR)(UCHAR)*ch).  Case insensitive hotstring example:
					//::ακ::Replacement Text
					// Update #2: On balance, it's not a clear win to use CharUpper since it
					// is expected to perform significantly worse than toupper.  Others have
					// said these Windows API functions that support diacritic letters can be
					// dramatically slower than the C-lib functions, though I haven't specifically
					// seen anything about CharUpper() being bad.  But since performance is of
					// particular concern here in the hook -- especially if there are hundreds
					// of hotstrings that need to be checked after each keystroke -- it seems
					// best to stick to toupper() (note that the Input command's searching loops
					// [further below] also use toupper() via stricmp().  One justification for
					// this is that it is rare to have diacritic letters in hotstrings, and even
					// rarer that someone would require them to be case insensitive.  There
					// are ways to script hotstring variants to work around this limitation.
					for (; cphs >= hs.mString; --cpbuf, --cphs)
						if (toupper(*cpbuf) != toupper(*cphs))
							break;
				// Relies on short-circuit boolean order:
				if (cphs < hs.mString && (hs.mDetectWhenInsideWord || cpbuf < g_HSBuf || !IsCharAlphaNumeric(*cpbuf)))
				{
					// MATCH FOUND
					// Since default KeyDelay is 0, and since that is expected to be typical, it seems
					// best to unconditionally post a message rather than trying to handle the backspacing
					// and replacing here.  This is because a KeyDelay of 0 might be fairly slow at
					// sending keystrokes if the system is under heavy load, in which case we would
					// not be returning to our caller in a timely fashion, which would case the OS to
					// think the hook is unreponsive, which in turn would cause it to timeout and
					// route the key through anyway (testing confirms this).
					if (!hs.mConformToCase)
						case_conform_mode = CASE_CONFORM_NONE;
					else
					{
						// Find out what case the user typed the string in so that we can have the
						// replacement produced in similar case:
						cpcase_end = g_HSBuf + g_HSBufLength;
						if (hs.mEndCharRequired)
							--cpcase_end;
						// Bug-fix for v1.0.19: First find out how many of the characters in the abbreviation
						// have upper and lowercase versions (i.e. exclude digits, punctuation, etc):
						for (characters_with_case = 0, first_char_with_case_is_upper = first_char_with_case_has_gone_by = false
							, cpcase_start = cpcase_end - hs.mStringLength
							; cpcase_start < cpcase_end; ++cpcase_start)
							if (IsCharLower(*cpcase_start) || IsCharUpper(*cpcase_start)) // A case-potential char.
							{
								if (!first_char_with_case_has_gone_by)
								{
									first_char_with_case_has_gone_by = true;
									if (IsCharUpper(*cpcase_start))
										first_char_with_case_is_upper = true; // Override default.
								}
								++characters_with_case;
							}
						if (!characters_with_case) // All characters in the abbreviation are caseless.
							case_conform_mode = CASE_CONFORM_NONE;
						else if (characters_with_case == 1)
							// Since there is only a single character with case potential, it seems best as
							// a default behavior to capitalize the first letter of the replacment whenever
							// that character was typed in uppercase.  The behavior can be overridden by
							// turning off the case-conform mode.
							case_conform_mode = first_char_with_case_is_upper ? CASE_CONFORM_FIRST_CAP : CASE_CONFORM_NONE;
						else // At least two characters have case potential. If all of them are upper, use ALL_CAPS.
						{
							if (!first_char_with_case_is_upper) // It can't be either FIRST_CAP or ALL_CAPS.
								case_conform_mode = CASE_CONFORM_NONE;
							else // First char is uppercase, and if all the others are too, this will be ALL_CAPS.
							{
								case_conform_mode = CASE_CONFORM_FIRST_CAP; // Set default.
								// Bug-fix for v1.0.19: Changed !IsCharUpper() below to IsCharLower() so that
								// caseless characters such as the @ symbol do not disqualify an abbreviation
								// from being considered "all uppercase":
								for (cpcase_start = cpcase_end - hs.mStringLength; cpcase_start < cpcase_end; ++cpcase_start)
									if (IsCharLower(*cpcase_start)) // Use IsCharLower to better support chars from non-English languages.
										break; // Any lowercase char disqualifies CASE_CONFORM_ALL_CAPS.
								if (cpcase_start == cpcase_end) // All case-possible characters are uppercase.
									case_conform_mode = CASE_CONFORM_ALL_CAPS;
								//else leave it at the default set above.
							}
						}
					}

					// Put the end char in the LOWORD and the case_conform_mode in the HIWORD.
					// Casting to UCHAR might be necessary to avoid problems when MAKELONG
					// casts a signed char to an unsigned WORD:
					PostMessage(g_hWnd, AHK_HOTSTRING, u, MAKELONG(hs.mEndCharRequired
						? (UCHAR)g_HSBuf[g_HSBufLength - 1] : 0, case_conform_mode));
					// Clean up:
					if (*hs.mReplacement)
					{
						// Since the buffer no longer reflects what is actually on screen to the left
						// of the caret position (since a replacement is about to be done), reset the
						// buffer, except for any end-char (since that might legitimately form part
						// of another hot string adjacent to the one just typed).  The end-char
						// sent by DoReplace() won't be captured (since it's "ignored input", which
						// is why it's put into the buffer manually here):
						if (hs.mEndCharRequired)
						{
							*g_HSBuf = g_HSBuf[g_HSBufLength - 1];
							g_HSBufLength = 1;
						}
						else
							g_HSBufLength = 0;
						g_HSBuf[g_HSBufLength] = '\0';
					}
					else if (hs.mDoBackspace)
					{
						// It's not a replacement, but we're doing backspaces, so adjust buf for backspaces
						// and the fact that the final char of the HS (if no end char) or the end char
						// (if end char requird) will have been suppressed and never made it to the
						// active window.  A simpler way to understand is to realize that the buffer now
						// contains (for recognition purposes, in its right side) the hotstring and its
						// end char (if applicable), so remove both:
						g_HSBufLength -= hs.mStringLength;
						if (hs.mEndCharRequired)
							--g_HSBufLength;
						g_HSBuf[g_HSBufLength] = '\0';
					}
					if (hs.mDoBackspace)
					{
						// Have caller suppress this final key pressed by the user, since it would have
						// to be backspaced over anyway.  Even if there is a visible Input command in
						// progress, this should still be okay since the input will still see the key,
						// it's just that the active window won't see it, which is okay since once again
						// it would have to be backspaced over anyway.  UPDATE: If an Input is in progress,
						// it should not receive this final key because otherwise the hotstring's backspacing
						// would backspace one too few times from the Input's point of view, thus the input
						// would have one extra, unwanted character left over (namely the first character
						// of the hotstring's abbreviation).  However, this method is not a complete
						// solution because it fails to work under a situation such as the following:
						// A hotstring script is started, followed by a separate script that uses the
						// Input command.  The Input script's hook will take precedence (since it was
						// started most recently), thus when the Hotstring's script's hook does sends
						// its replacement text, the Input script's hook will get a hold of it first
						// before the Hotstring's script has a chance to suppress it.  In other words,
						// The Input command command capture the ending character and then there will
						// be insufficient backspaces sent to clear the abbrevation out of it.  This
						// situation is quite rare so for now it's just mentioned here as a known limitation.
						treat_as_visible = false; // It might already have been false due to an invisible-input in progress, etc.
						suppress_hotstring_final_char = true; // This var probably must be separate from treat_as_visible to support invisible inputs.
					}
					break;
				}
			} // for()
		} // if buf not empty
	} // Yes, do collect hotstring input.

	// Note that it might have been in progress upon entry to this function but now isn't due to
	// INPUT_TERMINATED_BY_ENDKEY above:
	if (!do_input || g_input.status != INPUT_IN_PROGRESS || suppress_hotstring_final_char)
		return treat_as_visible;

	// Since above didn't return, the only thing remaining to do below is handle the input that's
	// in progress (which we know is the case otherwise other opportunities to return above would
	// have done so).  Hotstrings (if any) have already been fully handled by the above.

	#define ADD_INPUT_CHAR(ch) \
		if (g_input.BufferLength < g_input.BufferLengthMax)\
		{\
			g_input.buffer[g_input.BufferLength++] = ch;\
			g_input.buffer[g_input.BufferLength] = '\0';\
		}
	ADD_INPUT_CHAR(ch[0])
	if (byte_count > 1)
		// MSDN: "This usually happens when a dead-key character (accent or diacritic) stored in the
		// keyboard layout cannot be composed with the specified virtual key to form a single character."
		ADD_INPUT_CHAR(ch[1])
	if (!g_input.MatchCount) // The match list is empty.
	{
		if (g_input.BufferLength >= g_input.BufferLengthMax)
			g_input.status = INPUT_LIMIT_REACHED;
		return treat_as_visible;
	}
	// else even if BufferLengthMax has been reached, check if there's a match because a match should take
	// precedence over the length limit.

	// Otherwise, check if the buffer now matches any of the key phrases:
	if (g_input.FindAnywhere)
	{
		if (g_input.CaseSensitive)
		{
			for (UINT i = 0; i < g_input.MatchCount; ++i)
			{
				if (strstr(g_input.buffer, g_input.match[i]))
				{
					g_input.status = INPUT_TERMINATED_BY_MATCH;
					return treat_as_visible;
				}
			}
		}
		else // Not case sensitive.
		{
			for (UINT i = 0; i < g_input.MatchCount; ++i)
			{
				if (strcasestr(g_input.buffer, g_input.match[i]))
				{
					g_input.status = INPUT_TERMINATED_BY_MATCH;
					return treat_as_visible;
				}
			}
		}
	}
	else // Exact match is required
	{
		if (g_input.CaseSensitive)
		{
			for (UINT i = 0; i < g_input.MatchCount; ++i)
			{
				if (!strcmp(g_input.buffer, g_input.match[i]))
				{
					g_input.status = INPUT_TERMINATED_BY_MATCH;
					return treat_as_visible;
				}
			}
		}
		else // Not case sensitive.
		{
			for (UINT i = 0; i < g_input.MatchCount; ++i)
			{
				if (!stricmp(g_input.buffer, g_input.match[i]))
				{
					g_input.status = INPUT_TERMINATED_BY_MATCH;
					return treat_as_visible;
				}
			}
		}
	}

	// Otherwise, no match found.
	if (g_input.BufferLength >= g_input.BufferLengthMax)
		g_input.status = INPUT_LIMIT_REACHED;
	return treat_as_visible;
}
#undef shs  // To avoid naming conflicts
#endif  // i.e. above is only for the keyboard hook, not the mouse hook.



#ifdef INCLUDE_KEYBD_HOOK
#undef AllowKeyToGoToSystem
#define AllowKeyToGoToSystem AllowIt(g_KeybdHook, code, wParam, lParam, vk, sc, key_up, pKeyHistoryCurr)
#define AllowKeyToGoToSystemButDisguiseWinAlt AllowIt(g_KeybdHook, code, wParam, lParam \
	, vk, sc, key_up, pKeyHistoryCurr, true)
LRESULT AllowIt(HHOOK hhk, int code, WPARAM wParam, LPARAM lParam, vk_type vk, sc_type sc, bool key_up
	, KeyHistoryItem *pKeyHistoryCurr, bool DisguiseWinAlt = false)
// Always use the parameter vk rather than event.vkCode because the caller or caller's caller
// might have adjusted vk, namely to make it a left/right specific modifier key rather than a
// neutral one.
{
	// Prevent toggleable keys from being toggled (if the user wanted that) by suppressing it.
	// Seems best to suppress key-up events as well as key-down, since a key-up by itself,
	// if seen by the system, doesn't make much sense and might have unwanted side-effects
	// in rare cases (e.g. if the foreground app takes note of these types of key events).
	// Don't do this for ignored keys because that could cause an endless loop of
	// numlock events due to the keybd events that SuppressThisKey sends.
	// It's a little more readable and comfortable not to rely on short-circuit
	// booleans and instead do these conditions as separate IF statements.
	KBDLLHOOKSTRUCT &event = *(PKBDLLHOOKSTRUCT)lParam; // Needed by IS_IGNORED below.
	bool is_ignored = IS_IGNORED;
	if (!is_ignored)
	{
		if (kvk[vk].pForceToggle) // Key is a toggleable key.
		{
			// Dereference to get the global var's value:
			if (*(kvk[vk].pForceToggle) != NEUTRAL) // Prevent toggle.
				return SuppressThisKey;
		}
	}

	// This is done unconditionally so that even if a qualified Input is not in progress, the
	// variable will be correctly reset anyway:
	if (vk_to_ignore_next_time_down && vk_to_ignore_next_time_down == vk && !key_up)
		vk_to_ignore_next_time_down = 0;  // i.e. this ignore-for-the-sake-of-CollectInput() ticket has now been used.
	else if ((Hotstring::shs && !is_ignored) || (g_input.status == INPUT_IN_PROGRESS && !(g_input.IgnoreAHKInput && is_ignored)))
		if (!CollectInput(*pEvent, vk, sc, key_up, is_ignored)) // Key should be invisible (suppressed).
			return SuppressThisKey;

	// Do these here since the above "return SuppressThisKey" will have already done it in that case.
#ifdef ENABLE_KEY_HISTORY_FILE
	if (g_KeyHistoryToFile && pKeyHistoryCurr)
		KeyHistoryToFile(NULL, pKeyHistoryCurr->event_type, pKeyHistoryCurr->key_up
			, pKeyHistoryCurr->vk, pKeyHistoryCurr->sc);  // A fairly low overhead operation.
#endif

	UpdateKeyState(*pEvent, vk, sc, key_up, false);

	// UPDATE: The Win-L and Ctrl-Alt-Del workarounds below are still kept in effect in spite of the
	// anti-stick workaround done via GetModifierLRState().  This is because ResetHook() resets more
	// than just the modifiers and physical key state, which seems appropriate since the user might
	// be away for a long period of time while the computer is locked or the security screen is displayed.
	// Win-L uses logical keys, unlike Ctrl-Alt-Del which uses physical keys (i.e. Win-L can be simulated,
	// but Ctrl-Alt-Del must be physically pressed by the user):
	if (   vk == 'L' && !key_up && (g_modifiersLR_logical == MOD_LWIN  // i.e. *no* other keys but WIN.
		|| g_modifiersLR_logical == MOD_RWIN || g_modifiersLR_logical == (MOD_LWIN | MOD_RWIN))
		&& g_os.IsWinXPorLater())
	{
		// Since the user has pressed Win-L with *no* other modifier keys held down, and since
		// this key isn't being suppressed (since we're here in this function), the computer
		// is about to be locked.  When that happens, the hook is apparently disabled or
		// deinstalled until the user logs back in.  Because it is disabled, it will not be
		// notified when the user releases the LWIN or RWIN key, so we should assume that
		// it's now not in the down position.  This avoids it being thought to be down when the
		// user logs back in, which might cause hook hotkeys to accidentally fire.
		// Update: I've received an indication from a single Win2k user (unconfirmed from anyone
		// else) that the Win-L hotkey doesn't work on Win2k.  AutoIt3 docs confirm this.
		// Thus, it probably doesn't work on NT either.  So it's been changed to happen only on XP:
		ResetHook(true); // We already know that *only* the WIN key is down.
		// Above will reset g_PhysicalKeyState, especially for the windows keys and the 'L' key
		// (in our case), in preparation for re-logon:
	}

	// Although the delete key itself can be simulated (logical or physical), the user must be physically
	// (not logically) holding down CTRL and ALT for the ctrl-alt-del sequence to take effect,
	// which is why g_modifiersLR_physical is used vs. g_modifiersLR_logical (which is used above since
	// it's different).  Also, this is now done for XP -- in addition to NT4 & Win2k -- in case XP is
	// configured to display the NT/2k style security window instead of the task manager.  This is
	// probably very common because whenever the welcome screen is diabled, that's the default behavior?:
	// Control Panel > User Accounts > Use the welcome screen for fast and easy logon
	if (   (vk == VK_DELETE || vk == VK_DECIMAL) && !key_up         // Both of these qualify, see notes.
		&& (g_modifiersLR_physical & (MOD_LCONTROL | MOD_RCONTROL)) // At least one CTRL key is physically down.
		&& (g_modifiersLR_physical & (MOD_LALT | MOD_RALT))         // At least one ALT key is physically down.
		&& !(g_modifiersLR_physical & (MOD_LSHIFT | MOD_RSHIFT))    // Neither shift key is phys. down (WIN is ok).
		&& g_os.IsWinNT4orLater()   )
	{
		// Similar to the above case except for Windows 2000.  I suspect it also applies to NT,
		// but I'm not sure.  It seems safer to apply it to NT until confirmed otherwise.
		// Note that Ctrl-Alt-Delete works with *either* delete key, and it works regardless
		// of the state of Numlock (at least on XP, so it's probably that way on Win2k/NT also,
		// though it would be nice if this too is someday confirmed).  Here's the key history
		// someone for when the pressed ctrl-alt-del and then pressed esc to dismiss the dialog
		// on Win2k (Win2k invokes a 6-button dialog, with choices such as task manager and lock
		// workstation, if I recall correctly -- unlike XP which invokes task mgr by default):
		// A4  038	 	d	21.24	Alt            	
		// A2  01D	 	d	0.00	Ctrl           	
		// A2  01D	 	d	0.52	Ctrl           	
		// 2E  053	 	d	0.02	Num Del        	<-- notice how there's no following up event
		// 1B  001	 	u	2.80	Esc             <-- notice how there's no preceding down event
		// Other notes: On XP at least, shift key must not be down, otherwise Ctrl-Alt-Delete does
		// not take effect.  Windows key can be down, however.
		// Since the user will be gone for an unknown amount of time, it seems best just to reset
		// all hook tracking of the modifiers to the "up" position.  The user can always press them
		// down again upon return.  It also seems best to reset both logical and physical, just for
		// peace of mind and simplicity:
		ResetHook(true);
		// The above will also reset g_PhysicalKeyState so that especially the following will not
		// be thought to be physically down:CTRL, ALT, and DEL keys.  This is done in preparation
		// for returning from the security screen.  The neutral keys (VK_MENU and VK_CONTROL)
		// must also be reset -- not just because it's correct but because CollectInput() relies on it.
	}

	// Bug-fix for v1.0.20: The below section was moved out of LowLevelKeybdProc() to here because
	// alt_tab_menu_is_visible should not be set to true prior to knowing whether the current tab-down
	// event will be suppressed.  This is because if it is suppressed, the menu will not become visible
	// after all since the system will never see the tab-down event.
	// Having this extra check here, in addition to the other(s) that set alt_tab_menu_is_visible to be
	// true, allows AltTab and ShiftAltTab hotkeys to function even when the AltTab menu was invoked by
	// means other than an AltTabMenu or AltTabAndMenu hotkey.  The alt-tab menu becomes visible only
	// under these exact conditions, at least under WinXP:
	if (vk == VK_TAB && !key_up && !alt_tab_menu_is_visible
		&& ((g_modifiersLR_logical & MOD_LALT) || (g_modifiersLR_logical & MOD_RALT))
		&& !(g_modifiersLR_logical & MOD_LCONTROL) && !(g_modifiersLR_logical & MOD_RCONTROL))
		alt_tab_menu_is_visible = true;

	if (!kvk[vk].as_modifiersLR)
		return CallNextHookEx(hhk, code, wParam, lParam);

	// Due to above, we now know it's a modifier.

	// Don't do it this way because then the alt key itself can't be reliable used as "AltTabMenu"
	// (due to ShiftAltTab causing alt_tab_menu_is_visible to become false):
	//if (   alt_tab_menu_is_visible && !((g_modifiersLR_logical & MOD_LALT) || (g_modifiersLR_logical & MOD_RALT))
	//	&& !(key_up && pKeyHistoryCurr->event_type == 'h')   )  // In case the alt key itself is "AltTabMenu"
	if (   alt_tab_menu_is_visible && (vk == VK_MENU || vk == VK_LMENU || vk == VK_RMENU) && key_up
		// In case the alt key itself is "AltTabMenu":
		&& pKeyHistoryCurr->event_type != 'h' && pKeyHistoryCurr->event_type != 's'   )
		// It's important to reset in this case because if alt_tab_menu_is_visible were to
		// stay true and the user presses ALT in the future for a purpose other than to
		// display the Alt-tab menu, we would incorrectly believe the menu to be displayed:
		alt_tab_menu_is_visible = false;

	bool vk_is_win = vk == VK_LWIN || vk == VK_RWIN;
	if (DisguiseWinAlt && key_up && (vk_is_win || vk == VK_LMENU || vk == VK_RMENU || vk == VK_MENU))
	{
		// I think the best way to do this is to suppress the given key-event and substitute
		// some new events to replace it.  This is because otherwise we would probably have to
		// Sleep() or wait for the shift key-down event to take effect before calling
		// CallNextHookEx(), so that the shift key will be in effect in time for the win
		// key-up event to be disguised properly.  UPDATE: Currently, this doesn't check
		// to see if a shift key is already down for some other reason; that would be
		// pretty rare anyway, and I have more confidence in the reliability of putting
		// the shift key down every time.  UPDATE #2: Ctrl vs. Shift is now used to avoid
		// issues with the system's language-switch hotkey.  See detailed comments in
		// SetModifierLRState() about this.
		// Also, check the current logical state of CTRL to see if it's already down, for these reasons:
		// 1) There is no need to push it down again, since the release of ALT or WIN will be
		//    successfully disguised as long as it's down currently.
		// 2) If it's already down, the up-event part of the disguise keystroke would put it back
		//    up, which might mess up other things that rely upon it being down.
		bool disguise_it = true;  // Starting default.
		if ((g_modifiersLR_logical & MOD_LCONTROL) || (g_modifiersLR_logical & MOD_RCONTROL))
			disguise_it = false; // LCTRL or RCTRL is already down, so disguise is already in effect.
		else if (   vk_is_win && ((g_modifiersLR_logical & MOD_LSHIFT) || (g_modifiersLR_logical & MOD_RSHIFT)
			|| (g_modifiersLR_logical & MOD_LALT) || (g_modifiersLR_logical & MOD_RALT))   )
			disguise_it = false; // WIN key disguise is easier to satisfy, so don't need it in these cases either.
		// Since the below call to KeyEvent() calls the keybd hook reentrantly, a quick down-and-up
		// on Control is all that is necessary to disguise the key.  This is because the OS will see
		// that the Control keystroke occurred while ALT or WIN is still down because we haven't
		// done CallNextHookEx() yet.
		if (disguise_it)
			KeyEvent(KEYDOWNANDUP, VK_CONTROL); // Fix for v1.0.25: Use Ctrl vs. Shift to avoid triggering the LAlt+Shift language-change hotkey.
	}
	return CallNextHookEx(hhk, code, wParam, lParam);
}

#else // Mouse hook:
#define AllowKeyToGoToSystem AllowIt(g_MouseHook, code, wParam, lParam, pKeyHistoryCurr)
LRESULT AllowIt(HHOOK hhk, int code, WPARAM wParam, LPARAM lParam, KeyHistoryItem *pKeyHistoryCurr)
{
	// Since a mouse button that is physically down is not necessarily logically down -- such as
	// when the mouse button is a suppressed hotkey -- only update the logical state (which is the
	// state the OS believes the key to be in) when this even is non-supressed (i.e. allowed to
	// go to the system):
#ifdef FUTURE_USE_MOUSE_BUTTONS_LOGICAL
	KBDLLHOOKSTRUCT &event = *(PMSDLLHOOKSTRUCT)lParam;  // For convenience, maintainability, and possibly performance.
	switch (wParam)
	{
		case WM_LBUTTONUP: g_mouse_buttons_logical &= ~MK_LBUTTON; break;
		case WM_RBUTTONUP: g_mouse_buttons_logical &= ~MK_RBUTTON; break;
		case WM_MBUTTONUP: g_mouse_buttons_logical &= ~MK_MBUTTON; break;
		// Seems most correct to map NCX and X to the same VK since any given mouse is unlikely to
		// have both sets of these extra buttons?:
		case WM_NCXBUTTONUP:
		case WM_XBUTTONUP:
			g_mouse_buttons_logical &= ~(   (HIWORD(event.mouseData)) == XBUTTON1 ? MK_XBUTTON1 : MK_XBUTTON2   );
			break;
		case WM_LBUTTONDOWN: g_mouse_buttons_logical |= MK_LBUTTON; break;
		case WM_RBUTTONDOWN: g_mouse_buttons_logical |= MK_RBUTTON; break;
		case WM_MBUTTONDOWN: g_mouse_buttons_logical |= MK_MBUTTON; break;
		case WM_NCXBUTTONDOWN:
		case WM_XBUTTONDOWN:
			g_mouse_buttons_logical |= (HIWORD(event.mouseData) == XBUTTON1) ? MK_XBUTTON1 : MK_XBUTTON2;
			break;
	}
#endif
#ifdef ENABLE_KEY_HISTORY_FILE
	if (g_KeyHistoryToFile && pKeyHistoryCurr)
		KeyHistoryToFile(NULL, pKeyHistoryCurr->event_type, pKeyHistoryCurr->key_up
			, pKeyHistoryCurr->vk, pKeyHistoryCurr->sc);  // A fairly low overhead operation.
#endif
	return CallNextHookEx(hhk, code, wParam, lParam);
}
#endif



#ifdef INCLUDE_KEYBD_HOOK
LRESULT CALLBACK LowLevelKeybdProc(int code, WPARAM wParam, LPARAM lParam)
{
	if (code != HC_ACTION)  // MSDN docs specify that Both LL keybd & mouse hook should return in this case.
		return CallNextHookEx(g_KeybdHook, code, wParam, lParam);

	KeyHistoryItem *pKeyHistoryCurr = NULL; // Needs to be done early.
	KBDLLHOOKSTRUCT &event = *(PKBDLLHOOKSTRUCT)lParam;  // For convenience, maintainability, and possibly performance.

	// Change the event to be physical if that is indicated in its dwExtraInfo attribute.
	// This is done for cases when the hook is installed multiple times and one instance of
	// it wants to inform the others that this event should be considered physical for the
	// purpose of updating modifier and key states:
	if (event.dwExtraInfo == KEY_PHYS_IGNORE)
		event.flags &= ~LLKHF_INJECTED;

	// Make all keybd events physical to try to fool the system into accepting CTRL-ALT-DELETE.
	// This didn't work, which implies that Ctrl-Alt-Delete is trapped at a lower level than
	// this hook (folks have said that it's trapped in the keyboard driver itself):
	//event.flags &= ~LLKHF_INJECTED;

	// Note: Some scan codes are shared by more than one key (e.g. Numpad7 and NumpadHome).  This is why
	// the keyboard hook must be able to handle hotkeys by either their virtual key or their scan code.
	// i.e. if sc were always used in preference to vk, we wouldn't be able to distinguish between such keys.

	bool key_up = (wParam == WM_KEYUP || wParam == WM_SYSKEYUP);
	vk_type vk = (vk_type)event.vkCode;
	sc_type sc = (sc_type)event.scanCode;
	if (vk && !sc) // It might be possible for another app to call keybd_event with a zero scan code.
		sc = g_vk_to_sc[vk].a;
	// MapVirtualKey() does *not* include 0xE0 in HIBYTE if key is extended.  In case it ever
	// does in the future (or if event.scanCode ever does), force sc to be an 8-bit value
	// so that it's guaranteed consistent and to ensure it won't exceed SC_MAX (which might cause
	// array indexes to be out-of-bounds).  The 9th bit is later set to 1 if the key is extended:
	sc &= 0xFF;
	// Change sc to be extended if indicated.  But avoid doing so for VK_RSHIFT, which is
	// apparently considered extended by the API when it shouldn't be.  Update: Well, it looks like
	// VK_RSHIFT really is an extended key, at least on WinXP (and probably be extension on the other
	// NT based OSes as well).  What little info I could find on the 'net about this is contradictory,
	// but it's clear that some things just don't work right if the non-extended scan code is sent.  For
	// example, the shift key will appear to get stuck down in the foreground app if the non-extended
	// scan code is sent with VK_RSHIFT key-up event:
	if ((event.flags & LLKHF_EXTENDED)) // && vk != VK_RSHIFT)
		sc |= 0x100;

	// The below must be done prior to any returns that indirectly call UpdateModifierState():
	// Update: It seems best to do the below unconditionally, even if the OS is Win2k or WinXP,
	// since it seems like this translation will add value even in those cases:
	// To help ensure consistency with Windows XP and 2k, for which this hook has been primarily
	// designed and tested, translate neutral modifier keys into their left/right specific VKs,
	// since beardboy's testing shows that NT receives the neutral keys like Win9x does:
	switch (vk)
	{
	case VK_CONTROL: vk = (sc == SC_RCONTROL) ? VK_RCONTROL : VK_LCONTROL; break;
	case VK_MENU: vk = (sc == SC_RALT) ? VK_RMENU : VK_LMENU; break;
	case VK_SHIFT: vk = (sc == SC_RSHIFT) ? VK_RSHIFT : VK_LSHIFT; break;
	}

#else // Mouse Hook:
LRESULT CALLBACK LowLevelMouseProc(int code, WPARAM wParam, LPARAM lParam)
{
	KeyHistoryItem *pKeyHistoryCurr = NULL;
	// code != HC_ACTION should be evaluated PRIOR to considering the values
	// of wParam and lParam, because those values may be invalid or untrustworthy
	// whenever code < 0.  So the order in this short-circuit boolean expression
	// may be important:
	if (code != HC_ACTION)
		return AllowKeyToGoToSystem;

	MSLLHOOKSTRUCT &event = *(PMSLLHOOKSTRUCT)lParam;  // For convenience, maintainability, and possibly performance.

	// Make all mouse events physical to try to simulate mouse clicks in games that normally ignore
	// artificial input.
	//event.flags &= ~LLMHF_INJECTED;

	if (!(event.flags & LLMHF_INJECTED)) // Physical mouse movement or button action (uses LLMHF vs. LLKHF).
		g_TimeLastInputPhysical = event.time;

	if (wParam == WM_MOUSEMOVE) // Only after updating for physical input, above, is this checked.
		return AllowKeyToGoToSystem;

	// MSDN: WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MOUSEMOVE, WM_MOUSEWHEEL, WM_RBUTTONDOWN, or WM_RBUTTONUP.
	// But what about the middle button?  It's undocumented, but it is received.
	// What about doubleclicks (e.g. WM_LBUTTONDBLCLK): I checked: They are NOT received.
	// This is expected because each click in a doubleclick could be separately suppressed by
	// the hook, which would make become a non-doubleclick.
	vk_type vk = 0;
	short wheel_delta = 0;
	bool key_up = true;  // Init to safest value.
	switch (wParam)
	{
		case WM_MOUSEWHEEL:
			// MSDN: "A positive value indicates that the wheel was rotated forward,
			// away from the user; a negative value indicates that the wheel was rotated
			// backward, toward the user. One wheel click is defined as WHEEL_DELTA,
			// which is 120."  Must typecast to short (not int) otherwise the conversion
			// to negative/positive number won't be correct.  Also, I think the delta
			// can be greater than 120 only if the system can't keep up with how fast
			// the wheel is being turned (thus not generating an event for every
			// turn-click)?:
			wheel_delta = GET_WHEEL_DELTA_WPARAM(event.mouseData);
			vk = wheel_delta < 0 ? VK_WHEEL_DOWN : VK_WHEEL_UP;
			key_up = false; // Always consider wheel movements to be "key down" events.
			break;
		case WM_LBUTTONUP: vk = VK_LBUTTON;	break;
		case WM_RBUTTONUP: vk = VK_RBUTTON; break;
		case WM_MBUTTONUP: vk = VK_MBUTTON; break;
		// Seems most correct to map NCX and X to the same VK since any given mouse is unlikely to
		// have both sets of these extra buttons?:
		case WM_NCXBUTTONUP:
		case WM_XBUTTONUP: vk = (HIWORD(event.mouseData) == XBUTTON1) ? VK_XBUTTON1 : VK_XBUTTON2; break;
		case WM_LBUTTONDOWN: vk = VK_LBUTTON; key_up = false; break;
		case WM_RBUTTONDOWN: vk = VK_RBUTTON; key_up = false; break;
		case WM_MBUTTONDOWN: vk = VK_MBUTTON; key_up = false; break;
		case WM_NCXBUTTONDOWN:
		case WM_XBUTTONDOWN: vk = (HIWORD(event.mouseData) == XBUTTON1) ? VK_XBUTTON1 : VK_XBUTTON2; key_up = false; break;
	}

#endif

	bool is_ignored = IS_IGNORED;
	// This is done for more than just convenience.  It solves problems that would otherwise arise
	// due to the value of a global var such as KeyHistoryNext changing due to the reentrancy of
	// this procedure.  For example, a call to KeyEvent() in here would alter the value of
	// KeyHistoryNext, in most cases before we had a chance to finish using the old value.  In other
	// words, we use an automatic variable so that every instance of this function will get its
	// own copy of the variable whose value will stays constant until that instance returns:
	pKeyHistoryCurr = g_KeyHistory + g_KeyHistoryNext;
	if (++g_KeyHistoryNext >= MAX_HISTORY_KEYS)
		g_KeyHistoryNext = 0;
	pKeyHistoryCurr->vk = vk;
#ifdef INCLUDE_KEYBD_HOOK
	// Intentionally log a zero if it comes in that way, prior to using MapVirtualKey() to try to resolve it:
	pKeyHistoryCurr->sc = sc;
#else
	pKeyHistoryCurr->sc = 0;
#endif
	pKeyHistoryCurr->key_up = key_up;
	pKeyHistoryCurr->event_type = is_ignored ? 'i' : ' ';
	g_HistoryTickNow = GetTickCount();
	pKeyHistoryCurr->elapsed_time = (g_HistoryTickNow - g_HistoryTickPrev) / (float)1000;
	g_HistoryTickPrev = g_HistoryTickNow;
	HWND fore_win = GetForegroundWindow();
	if (fore_win)
		GetWindowText(fore_win, pKeyHistoryCurr->target_window, sizeof(pKeyHistoryCurr->target_window));
	else
		strcpy(pKeyHistoryCurr->target_window, "N/A");

#ifdef INCLUDE_KEYBD_HOOK
	// If the scan code is extended, the key that was pressed is not a dual-state numpad key,
	// i.e. it could be the counterpart key, such as End vs. NumpadEnd, located elsewhere on
	// the keyboard, but we're not interested in those.  Also, Numlock must be ON because
	// otherwise the driver will not generate those false-physical shift key events:
	if (!(sc & 0x100) && (IsKeyToggledOn(VK_NUMLOCK)))
	{
		switch (vk)
		{
		case VK_DELETE: case VK_DECIMAL: pad_state[PAD_DECIMAL] = !key_up; break;
		case VK_INSERT: case VK_NUMPAD0: pad_state[PAD_NUMPAD0] = !key_up; break;
		case VK_END:    case VK_NUMPAD1: pad_state[PAD_NUMPAD1] = !key_up; break;
		case VK_DOWN:   case VK_NUMPAD2: pad_state[PAD_NUMPAD2] = !key_up; break;
		case VK_NEXT:   case VK_NUMPAD3: pad_state[PAD_NUMPAD3] = !key_up; break;
		case VK_LEFT:   case VK_NUMPAD4: pad_state[PAD_NUMPAD4] = !key_up; break;
		case VK_CLEAR:  case VK_NUMPAD5: pad_state[PAD_NUMPAD5] = !key_up; break;
		case VK_RIGHT:  case VK_NUMPAD6: pad_state[PAD_NUMPAD6] = !key_up; break;
		case VK_HOME:   case VK_NUMPAD7: pad_state[PAD_NUMPAD7] = !key_up; break;
		case VK_UP:     case VK_NUMPAD8: pad_state[PAD_NUMPAD8] = !key_up; break;
		case VK_PRIOR:  case VK_NUMPAD9: pad_state[PAD_NUMPAD9] = !key_up; break;
		}
	}

	// Track physical state of keyboard & mouse buttons since GetAsyncKeyState() doesn't seem
	// to do so, at least under WinXP.  Also, if it's a modifier, let another section handle it
	// because it's not as simple as just setting the value to true or false (e.g. if LShift
	// goes up, the state of VK_SHIFT should stay down if VK_RSHIFT is down, or up otherwise).
	// Also, even if this input event will wind up being suppressed (usually because of being
	// a hotkey), still update the physical state anyway, because we want the physical state to
	// be entirely independent of the logical state (i.e. we want the key to be reported as
	// physically down even if it isn't logically down):
	if (!kvk[vk].as_modifiersLR && EventIsPhysical(event, vk, sc, key_up))
		g_PhysicalKeyState[vk] = key_up ? 0 : STATE_DOWN;

	// Pointer to the key record for the current key event.  Establishes this_key as an alias
	// for the array element in kvk or ksc that corresponds to the vk or sc, respectively.
	// I think the compiler can optimize the performance of reference variables better than
	// pointers because the pointer indirection step is avoided.  In any case, this must be
	// a true alias to the object, not a copy of it, because it's address (&this_key) is compared
	// to other addresses for equality further below.
	key_type &this_key = *(ksc[sc].sc_takes_precedence ? (ksc + sc) : (kvk + vk));

#else // Mouse hook.
	if (EventIsPhysical(event, key_up))
		g_PhysicalKeyState[vk] = key_up ? 0 : STATE_DOWN;
	key_type &this_key = *(kvk + vk);
#endif

	// Do this after above since AllowKeyToGoToSystem requires that sc be properly determined.
	// Another reason to do it after the above is due to the fact that KEY_PHYS_IGNORE permits
	// an ignored key to be considered physical input, which is handled above:
	if (is_ignored)
	{
		// This is a key sent by our own app that we want to ignore.
		// It's important never to change this to call the SuppressKey function because
		// that function would cause an infinite loop when the Numlock key is pressed,
		// which would likely hang the entire system:
		// UPDATE: This next part is for cases where more than one script is using the hook
		// simultaneously.  In such cases, it desirable for the KEYEVENT_PHYS() of one
		// instance to affect the down-state of the current prefix-key in the other
		// instances.  This check is done here -- even though there may be a better way to
		// implement it -- to minimize the chance of side-effects that a more fundamental
		// change might cause (i.e. a more fundamental change would require a lot more
		// testing, though it might also fix more things):
		if (event.dwExtraInfo == KEY_PHYS_IGNORE && key_up && pPrefixKey == &this_key)
		{
			this_key.is_down = false;
			this_key.down_performed_action = false;  // Seems best, but only for PHYS_IGNORE.
			pPrefixKey = NULL;
		}
		return AllowKeyToGoToSystem;
	}

#ifdef INCLUDE_KEYBD_HOOK
	// The below DISGUISE events are done only after ignored events are returned from, above.
	// In other words, only non-ignored events (usually physical) are disguised.
	// Do this only after the above because the SuppressThisKey macro relies
	// on the vk variable being available.  It also relies upon the fact that sc has
	// already been properly determined:
	// In rare cases it may be necessary to disguise both left and right, which is why
	// it's not done as a generic windows key:
	if (   key_up && ((disguise_next_lwin_up && vk == VK_LWIN) || (disguise_next_rwin_up && vk == VK_RWIN)
		 || (disguise_next_lalt_up && vk == VK_LMENU) || (disguise_next_ralt_up && vk == VK_RMENU))   )
	{
		// Do this first to avoid problems with reentrancy triggered by the KeyEvent() calls further below.
		switch (vk)
		{
		case VK_LWIN: disguise_next_lwin_up = false; break;
		case VK_RWIN: disguise_next_rwin_up = false; break;
		// For now, assume VK_MENU the left alt key.  This neutral key is probably never received anyway
		// due to the nature of this type of hook on NT/2k/XP and beyond.  Later, this can be furher
		// optimized to check the scan code and such (what's being done here isn't that essential to
		// start with, so it's not a high priority -- but when it is done, be sure to review the
		// above IF statement also).  UPDATE: Teh above is no longer a concern since neutral keys
		// are translated above into their left/right-specific counterpart:
		case VK_LMENU: disguise_next_lalt_up = false; break;
		case VK_RMENU: disguise_next_ralt_up = false; break;
		}
		// Send our own up-event to replace this one.  But since ours has the Shift key
		// held down for it, the Start Menu or foreground window's menu bar won't be invoked.
		// It's necessary to send an up-event so that it's state, as seen by the system,
		// is put back into the up position, which would be needed if its previous
		// down-event wasn't suppressed (probably due to the fact that this win or alt
		// key is a prefix but not a suffix).
		// Fix for v1.0.25: Use CTRL vs. Shift to avoid triggering the LAlt+Shift language-change hotkey.
		// This is definitely needed for ALT, but is done here for WIN also in case ALT is down,
		// which might cause the use of SHIFT as the disguise key to trigger the language switch.
		if (!((g_modifiersLR_logical & MOD_LCONTROL) || (g_modifiersLR_logical & MOD_RCONTROL))) // CTRL is not down.
			KeyEvent(KEYDOWNANDUP, VK_CONTROL);
		// Since the above call to KeyEvent() calls the keybd hook reentrantly, a quick down-and-up
		// on Control is all that is necessary to disguise the key.  This is because the OS will see
		// that the Control keystroke occurred while ALT or WIN is still down because we haven't
		// done CallNextHookEx() yet.
		return AllowKeyToGoToSystem;
	}

#else // Mouse hook
	// If no vk, there's no mapping for this key, so currently there's no way to process it.
	// Also, if the script is displaying a menu (tray, main, or custom popup menu), always
	// pass left-button events through -- even if LButton is defined as a hotkey -- so
	// that menu items can be properly selected.  This is necessary because if LButton is
	// a hotkey, it can't launch now anyway due to the script being uninterruptible while
	// a menu is visible.  And since it can't launch, it can't do its typical "MouseClick
	// left" to send a true mouse-click through as a replacement for the suppressed
	// button-down and button-up events caused by the hotkey:
	if (!vk || (g_MenuIsVisible && vk == VK_LBUTTON))
	{
		// Bug-fix for v1.0.22: If "LControl & LButton::" (and perhaps similar combinations)
		// is a hotkey, the foreground window would think that the mouse is stuck down, at least
		// if the user clicked outside the menu to dismiss it.  Specifically, this comes about
		// as follows:
		// The wrong up-event is suppressed:
		// ...because down_performed_action was true when it should have been false
		// ...because the while-menu-was-displayed up-event never set it to false
		// ...because it returned too early here before it could get to that part further below.
		if (vk)
		{
			this_key.down_performed_action = false; // Seems ok in this case to do this for both key_up and !key_up.
			this_key.is_down = !key_up;
		}
		return AllowKeyToGoToSystem;
	}
#endif

// Uncomment this section to have it report the vk and sc of every key pressed (can be useful):
//PostMessage(g_hWnd, AHK_HOOK_TEST_MSG, vk, sc);
//return AllowKeyToGoToSystem;

#ifdef INCLUDE_KEYBD_HOOK
	if (pPrefixKey && pPrefixKey != &this_key && !key_up && !this_key.as_modifiersLR)
		// Any key-down event (other than those already ignored and returned from,
		// above) should probably be considered an attempt by the user to use the
		// prefix key that's currently being held down as a "modifier".  That way,
		// if pPrefixKey happens to also be a suffix, its suffix action won't fire
		// when the key is released, which is probably the correct thing to do 90%
		// or more of the time.  But don't consider the modifiers themselves to have
		// been modified by  prefix key, since that is almost never desirable:
		pPrefixKey->was_just_used = AS_PREFIX;
#else
	// Update: The above is now done only for keyboard hook, not the mouse.  This is because
	// most people probably would not want a prefix key's suffix-action to be stopped
	// from firing just because a non-hotkey mouse button was pressed while the key
	// was held down (i.e. for games).  Update #2: A small exception to this has been made:
	// Prefix keys that are also modifiers (ALT/SHIFT/CTRL/WIN) will now not fire their
	// suffix action on key-up if they modified a mouse button event (since Ctrl-LeftClick,
	// for example, is a valid native action and we don't want to give up that flexibility).
	if (pPrefixKey && pPrefixKey != &this_key && !key_up && pPrefixKey->as_modifiersLR)
		pPrefixKey->was_just_used = AS_PREFIX;
#endif
	// WinAPI docs state that for both virtual keys and scan codes:
	// "If there is no translation, the return value is zero."
	// Therefore, zero is never a key that can be validly configured (and likely it's never received here anyway).
	// UPDATE: For performance reasons, this check isn't even done.  Even if sc and vk are both zero, both kvk[0]
	// and ksc[0] should have all their attributes initialized to FALSE so nothing should happen for that key
	// anyway.
	//if (!vk && !sc)
	//	return AllowKeyToGoToSystem;

	if (!this_key.used_as_prefix && !this_key.used_as_suffix)
		return AllowKeyToGoToSystem;

	int down_performed_action, was_down_before_up;
	if (key_up)
	{
		// Save prior to reset.  These var's should only be used further below in conjunction with key_up
		// being TRUE.  Otherwise, their values will be unreliable (refer to some other key, probably).
		was_down_before_up = this_key.is_down;
		down_performed_action = this_key.down_performed_action;  // Save prior to possible reset below.
		// Reset these values in preparation for the next call to this procedure that involves this key:
		this_key.down_performed_action = false;
	}
	this_key.is_down = !key_up;
	bool modifiers_were_corrected = false;

#ifdef INCLUDE_KEYBD_HOOK
	// The below was added to fix hotkeys that have a neutral suffix such as "Control & LShift".
	// It may also fix other things and help future enhancements:
	if (this_key.as_modifiersLR)
	{
		// The neutral modifier "Win" is not currently supported.
		kvk[VK_CONTROL].is_down = kvk[VK_LCONTROL].is_down || kvk[VK_RCONTROL].is_down;
		kvk[VK_MENU].is_down = kvk[VK_LMENU].is_down || kvk[VK_RMENU].is_down;
		kvk[VK_SHIFT].is_down = kvk[VK_LSHIFT].is_down || kvk[VK_RSHIFT].is_down;
		// No longer possible because vk is translated early on from neutral to left-right specific:
		// I don't think these ever happen with physical keyboard input, but it might with artificial input:
		//case VK_CONTROL: kvk[sc == SC_RCONTROL ? VK_RCONTROL : VK_LCONTROL].is_down = !key_up; break;
		//case VK_MENU: kvk[sc == SC_RALT ? VK_RMENU : VK_LMENU].is_down = !key_up; break;
		//case VK_SHIFT: kvk[sc == SC_RSHIFT ? VK_RSHIFT : VK_LSHIFT].is_down = !key_up; break;
	}
#else  // Mouse Hook
// If the mouse hook is installed without the keyboard hook, update g_modifiersLR_logical
// manually so that it can be referred to by the mouse hook after this point:
	if (!g_KeybdHook)
	{
		g_modifiersLR_logical = g_modifiersLR_logical_non_ignored = GetModifierLRState(true);
		modifiers_were_corrected = true;
	}
#endif

	modLR_type modifiersLRnew;
	HotkeyIDType hotkey_id = HOTKEY_ID_INVALID;  // Set default.
	bool no_suppress = false;  // Hotkeys are normally suppressed, so set this behavior as default.
	#define GET_HOTKEY_ID_AND_FLAGS(id_with_flags) \
		hotkey_id = id_with_flags;\
		no_suppress = hotkey_id & HOTKEY_NO_SUPPRESS;\
		hotkey_id &= HOTKEY_ID_MASK

	///////////////////////////////////////////////////////////////////////////////////////
	// CASE #1 of 4: PREFIX key has been pressed down.  But use it in this capacity only if
	// no other prefix is already in effect or if this key isn't a suffix.  Update: Or if
	// this key-down is the same as the prefix already down, since we want to be able to
	// a prefix when it's being used in its role as a modified suffix (see below comments).
	///////////////////////////////////////////////////////////////////////////////////////
	if (this_key.used_as_prefix && !key_up && (!pPrefixKey || !this_key.used_as_suffix || &this_key == pPrefixKey))
	{
		// This check is necessary in cases such as the following, in which the "A" key continues
		// to repeat becauses pressing a mouse button (unlike pressing a keyboard key) does not
		// stop the prefix key from repeating:
		// $a::send, a
		// a & lbutton::
		if (&this_key != pPrefixKey)
		{
			// Override any other prefix key that might be in effect with this one, in case the
			// prior one, due to be old for example, was invalid somehow.  UPDATE: It seems better
			// to leave the old one in effect to support the case where one prefix key is modifying
			// a second one in its role as a suffix.  In other words, if key1 is a prefix and
			// key2 is both a prefix and a suffix, we want to leave key1 in effect as a prefix,
			// rather than key2.  Hence, a null-check was added in the above if-stmt:
			pPrefixKey = &this_key;
			// It should be safe to init this because even if the current key is repeating,
			// it should be impossible to receive here the key-downs that occurred after
			// the first, because there's a return-on-repeat check farther above (update: that check
			// is gone now).  Even if that check weren't done, it's safe to reinitialize this to zero
			// because on most (all?) keyboards & OSs, the moment the user presses another key while
			// this one is held down, the key-repeating ceases and does not resume for
			// this key (though the second key will begin to repeat if it too is held down).
			// In other words, the fear that this would be wrongly initialized and thus cause
			// this prefix's suffix-action to fire upon key-release seems unfounded.
			// It seems easier (and may perform better than alternative ways) to init this
			// here rather than say, upon the release of the prefix key:
			pPrefixKey->was_just_used = 0;
		}

		// This new section was added May 30, 2004, to fix scenarios such as the following example:
		// a & b::Msgbox a & b
		// $^a::MsgBox a
		// Previously, the ^a hotkey would only fire on key-up (unless it was registered, in which
		// case it worked as intended on the down-event).  When the user presses a, it's okay (and
		// probably desirable) to have recorded that event as a prefix-key-down event (above).
		// But in addition to that, we now check if this is a normal, modified hotkey that should
		// fire now rather than waiting for the key-up event.  This is done because it makes sense,
		// it's more correct, and also it makes the behavior of a hooked ^a hotkey consistent with
		// that of a registered ^a.

		// Prior to considering whether to fire a hotkey, correct the hook's modifier state.
		// Although this is rarely needed, there are times when the OS disables the hook, thus
		// it is possible for it to miss keystrokes.  See comments in GetModifierLRState()
		// for more info:
		if (!modifiers_were_corrected)
		{
			modifiers_were_corrected = true;
			GetModifierLRState(true);
		}

		// non_ignored is always used when considering whether a key combination is in place to
		// trigger a hotkey:
#ifdef INCLUDE_KEYBD_HOOK
		modifiersLRnew = g_modifiersLR_logical_non_ignored;
		if (this_key.as_modifiersLR)
			// Hotkeys are not defined to modify themselves, so look for a match accordingly.
			modifiersLRnew &= ~this_key.as_modifiersLR;
		// For this case to be checked, there must be at least one modifier key currently down (other
		// than this key itself if it's a modifier), because if there isn't and this prefix is also
		// a suffix, its suffix action should only fire on key-up (i.e. not here, but later on):
		if (modifiersLRnew)
			GET_HOTKEY_ID_AND_FLAGS(ksc[sc].sc_takes_precedence ? Kscm(modifiersLRnew, sc) : Kvkm(modifiersLRnew, vk));
		// Alt-tab need not be checked here (like it is in the similar section below that calls 
		// GET_HOTKEY_ID_AND_FLAGS) because all such hotkeys use (or were converted at load-time to use)
		// a modifier_vk, not a set of modifiers or modifierlr's:
		//if (hotkey_id == HOTKEY_ID_INVALID && alt_tab_menu_is_visible)
		//...
#else // Mouse hook:
		if (g_modifiersLR_logical_non_ignored) // See above for explanation.
			GET_HOTKEY_ID_AND_FLAGS(Kvkm(g_modifiersLR_logical_non_ignored, vk));
#endif

		if (hotkey_id == HOTKEY_ID_INVALID)
		{
			// In this case, a key-down event can't trigger a suffix, so return immediately:
#ifdef INCLUDE_KEYBD_HOOK
			return (this_key.as_modifiersLR || (this_key.no_suppress & NO_SUPPRESS_PREFIX)
				// The order on this line important; it relies on short-circuit boolean:
				|| (this_key.pForceToggle && *this_key.pForceToggle == NEUTRAL))
				? AllowKeyToGoToSystem : SuppressThisKey;
#else
			return (this_key.no_suppress & NO_SUPPRESS_PREFIX) ? AllowKeyToGoToSystem : SuppressThisKey;
#endif
		}
	}

	//////////////////////////////////////////////////////////////////////////////////
	// CASE #2 of 4: SUFFIX key (that's not a prefix, or is one but has just been used
	// in its capacity as a suffix instead) has been released.
	// This is done before Case #3 for performance reasons.
	//////////////////////////////////////////////////////////////////////////////////
	if (this_key.used_as_suffix && pPrefixKey != &this_key && key_up) // Since key_up, hotkey_id == INVALID.
	{
		// If it did perform an action, suppress this key-up event.  Do this even
		// if this key is a modifier because it's previous key-down would have
		// already been suppressed (since this case is for suffixes that aren't
		// also prefixes), thus the key-up can be safely suppressed as well.
		// It's especially important to do this for keys whose up-events are
		// special actions within the OS, such as AppsKey, Lwin, and Rwin.
		// Toggleable keys are also suppressed here on key-up because their
		// previous key-down event would have been suppressed in order for
		// down_performed_action to be true.  UPDATE: Added handling for
		// NO_SUPPRESS_NEXT_UP_EVENT and also applied this next part to both
		// mouse and keyboard:
		bool suppress_up_event;
		if (this_key.no_suppress & NO_SUPPRESS_NEXT_UP_EVENT)
		{
			suppress_up_event = false;
			this_key.no_suppress &= ~NO_SUPPRESS_NEXT_UP_EVENT;  // This ticket has been used up.
		}
		else // the default is to suppress the up-event.
			suppress_up_event = true;
		if (down_performed_action)
			return suppress_up_event ? SuppressThisKey : AllowKeyToGoToSystem;
		// Otherwise let it be processed normally:
		return AllowKeyToGoToSystem;
	}

	//////////////////////////////////////////////
	// CASE #3 of 4: PREFIX key has been released.
	//////////////////////////////////////////////
	if (   (this_key.used_as_prefix) && key_up   )  // Since key_up, hotkey_id == INVALID.
	{
		if (pPrefixKey == &this_key)
			pPrefixKey = NULL;
		// Else it seems best to keep the old one in effect.  This could happen, for example,
		// if the user holds down prefix1, holds down prefix2, then releases prefix1.
		// In that case, we would want to keep the most recent prefix (prefix2) in effect.
		// This logic would fail to work properly in a case like this if the user releases
		// prefix2 but still has prefix1 held down.  The user would then have to release
		// prefix1 and press it down again to get the hook to realize that it's in effect.
		// This seems very unlikely to be something commonly done by anyone, so for now
		// it's just documented here as a limitation.

		if (this_key.it_put_alt_down) // key pushed ALT down, or relied upon it already being down, so go up:
		{
			this_key.it_put_alt_down = false;
			KeyEvent(KEYUP, VK_MENU);
		}
		if (this_key.it_put_shift_down) // similar to above
		{
			this_key.it_put_shift_down = false;
			KeyEvent(KEYUP, VK_SHIFT);
		}
		// The order of expressions in this IF is important; it relies on short-circuit boolean:
#ifdef INCLUDE_KEYBD_HOOK
		if (this_key.pForceToggle && *this_key.pForceToggle == NEUTRAL)
		{
			// It's done this way because CapsLock, for example, is a key users often
			// press quickly while typing.  I suspect many users are like me in that
			// they're in the habit of not having releasing the CapsLock key quite yet
			// before they resume typing, expecting it's new mode to be in effect.
			// This resolves that problem by always toggling the state of a toggleable
			// key upon key-down.  If this key has just acted in its role of a prefix
			// to trigger a suffix action, toggle its state back to what it was before
			// because the firing of a hotkey should not have the side-effect of also
			// toggling the key:
			// Toggle the key by replacing this key-up event with a new sequence
			// of our own.  This entire-replacement is done so that the system
			// will see all three events in the right order:
			if (this_key.was_just_used == AS_PREFIX_FOR_HOTKEY)
			{
				KEYEVENT_PHYS(KEYUP, vk, sc); // Mark it as physical for any other hook instances.
				KeyEvent(KEYDOWNANDUP, vk, sc);
				return SuppressThisKey;
			}

			// Otherwise, if it was used to modify a non-suffix key, or it was just
			// pressed and released without any keys in between, don't suppress its up-event
			// at all.  UPDATE: Don't return here if it didn't modify anything because
			// this prefix might also be a suffix. Let later sections handle it then.
			if (this_key.was_just_used == AS_PREFIX)
				return AllowKeyToGoToSystem;
		}
		else
#endif
			// Seems safest to suppress this key if the user pressed any non-modifier key while it
			// was held down.  As a side-effect of this, if the user holds down numlock, for
			// example, and then presses another key that isn't actionable (i.e. not a suffix),
			// the numlock state won't be toggled even it's normally configured to do so.
			// This is probably the right thing to do in most cases.
			// Older note:
			// In addition, this suppression is relied upon to prevent toggleable keys from toggling
			// when they are used to modify other keys.  For example, if "Capslock & A" is a hotkey,
			// the state of the Capslock key should not be changed when the hotkey is pressed.
			// Do this check prior to the below check (give it precedence).
			if (this_key.was_just_used)  // AS_PREFIX or AS_PREFIX_FOR_HOTKEY.
#ifndef INCLUDE_KEYBD_HOOK // Mouse hook
				return (this_key.no_suppress & NO_SUPPRESS_PREFIX) ? AllowKeyToGoToSystem : SuppressThisKey;
#else // Keyboard hook
			{
				if (this_key.as_modifiersLR)
					return (this_key.was_just_used == AS_PREFIX_FOR_HOTKEY) ? AllowKeyToGoToSystemButDisguiseWinAlt
						: AllowKeyToGoToSystem;  // i.e. don't disguise Win or Alt key if it didn't fire a hotkey.
				else if (this_key.no_suppress & NO_SUPPRESS_PREFIX)
					return AllowKeyToGoToSystem;
				else
					return SuppressThisKey;
			}
#endif

		// Since above didn't return, this key-up for this prefix key wasn't used in it's role
		// as a prefix.  If it's not a suffix, we're done, so just return.  Don't do
		// "DisguiseWinAlt" because we want the key's native key-up function to take effect.
		// Also, allow key-ups for toggleable keys that the user wants to be toggleable to
		// go through to the system, because the prior key-down for this prefix key
		// wouldn't have been suppressed and thus this up-event goes with it (and this
		// up-even is also needed by the OS, at least WinXP, to properly set the indicator
		// light and toggle state):
		if (!this_key.used_as_suffix)
#ifdef INCLUDE_KEYBD_HOOK
			return (this_key.as_modifiersLR || (this_key.no_suppress & NO_SUPPRESS_PREFIX)
				// The order on this line important; it relies on short-circuit boolean:
				|| (this_key.pForceToggle && *this_key.pForceToggle == NEUTRAL))
				? AllowKeyToGoToSystem : SuppressThisKey;
#else
			return (this_key.no_suppress & NO_SUPPRESS_PREFIX) ? AllowKeyToGoToSystem : SuppressThisKey;
#endif

		// Since the above didn't return, this key is both a prefix and a suffix, but
		// is currently operating in its capacity as a suffix.
		if (!was_down_before_up)
			// If this key wasn't thought to be down prior to this up-event, it's probably because
			// it is registered with another prefix by RegisterHotkey().  In this case, the keyup
			// should be passed back to the system rather than performing it's key-up suffix
			// action.  UPDATE: This can't happen with a low-level hook.  But if there's another
			// low-level hook installed that receives events before us, and it's not
			// well-implemented (i.e. it sometimes sends ups without downs), this check
			// may help prevent unexpected behavior.
			return AllowKeyToGoToSystem;

		// Since no suffix action was triggered while it was held down, fall through rather than
		// returning, so that the key's own suffix action will be considered.
	}

	////////////////////////////////////////////////////////////////////////////////////////////////////
	// CASE #4 of 4: SUFFIX key has been pressed down (or released if it's a key-up event, in which case
	// it fell through from CASE #3 above).  Update: This case can also happen if it fell through
	// from case #1 (i.e. it already determined the value of hotkey_id).
	////////////////////////////////////////////////////////////////////////////////////////////////////
	// First correct modifiers, because at this late a state, the likelihood of firing a hotkey is high.
	// For details, see comments for "modifiers_were_corrected" above:
	if (!modifiers_were_corrected)
	{
		modifiers_were_corrected = true;
		GetModifierLRState(true);
	}
	if (pPrefixKey && !key_up && hotkey_id == HOTKEY_ID_INVALID) // Helps performance by avoiding all the below checking.
	{
		// Action here is considered first, and takes precedence since a suffix's ModifierVK/SC should
		// take effect regardless of whether any win/ctrl/alt/shift modifiers are currently down, even if
		// those modifiers themselves form another valid hotkey with this suffix.  In other words,
		// ModifierVK/SC combos take precedence over normally-modified combos:
		int i;
		for (i = 0; i < this_key.nModifierVK; ++i)
			if (kvk[this_key.ModifierVK[i].vk].is_down)
			{
				// Since the hook is now designed to receive only left/right specific modifier keys
				// -- never the neutral keys -- don't say that a neutral prefix key is down because
				// then it would never be released properly by the other main prefix/suffix handling
				// cases of the hook.  Instead, always identify which prefix key (left or right) is
				// in effect:
				switch (this_key.ModifierVK[i].vk)
				{
				case VK_SHIFT: pPrefixKey = kvk + (kvk[VK_RSHIFT].is_down ? VK_RSHIFT : VK_LSHIFT); break;
				case VK_CONTROL: pPrefixKey = kvk + (kvk[VK_RCONTROL].is_down ? VK_RCONTROL : VK_LCONTROL); break;
				case VK_MENU: pPrefixKey = kvk + (kvk[VK_RMENU].is_down ? VK_RMENU : VK_LMENU); break;
				default: pPrefixKey = kvk + this_key.ModifierVK[i].vk;
				}
				// Do this, even though it was probably already done close to the top of the function,
				// just in case this for-loop changed the value pPrefixKey (perhaps because there
				// is currently more than one prefix being held down):
				pPrefixKey->was_just_used = AS_PREFIX_FOR_HOTKEY;
				GET_HOTKEY_ID_AND_FLAGS(this_key.ModifierVK[i].id_with_flags);
				break;
			}
 		if (hotkey_id == HOTKEY_ID_INVALID)  // Now check scan codes since above didn't find one.
		{
			for (i = 0; i < this_key.nModifierSC; ++i)
				if (ksc[this_key.ModifierSC[i].sc].is_down)
				{
					pPrefixKey = ksc + this_key.ModifierSC[i].sc;
					pPrefixKey->was_just_used = AS_PREFIX_FOR_HOTKEY;
					GET_HOTKEY_ID_AND_FLAGS(this_key.ModifierSC[i].id_with_flags);
					break;
				}
		}
		if (hotkey_id == HOTKEY_ID_INVALID)
		{
			// Search again, but this time do it with this_key translated into its neutral counterpart.
			// This avoids the need to display a warning dialog for an example such as the following,
			// which was previously unsupported:
			// AppsKey & Control::MsgBox %A_ThisHotkey%
			// Note: If vk was a neutral modifier when it first came in (e.g. due to NT4), it was already
			// translated early on (above) to be non-neutral.
			vk_type vk_neutral = 0;  // Set default.  Note that VK_LWIN/VK_RWIN have no neutral VK.
			switch (vk)
			{
			case VK_LCONTROL:
			case VK_RCONTROL: vk_neutral = VK_CONTROL; break;
			case VK_LMENU:
			case VK_RMENU:    vk_neutral = VK_MENU; break;
			case VK_LSHIFT:
			case VK_RSHIFT:   vk_neutral = VK_SHIFT; break;
			}
			if (vk_neutral)
			{
				// These next two for() loops are nearly the same as the ones above, so see comments there
				// and maintain them together:
				for (i = 0; i < kvk[vk_neutral].nModifierVK; ++i)
					if (kvk[kvk[vk_neutral].ModifierVK[i].vk].is_down)
					{
						// See the nearly identical section above for comments on the below:
						switch (kvk[vk_neutral].ModifierVK[i].vk)
						{
						case VK_SHIFT: pPrefixKey = kvk + (kvk[VK_RSHIFT].is_down ? VK_RSHIFT : VK_LSHIFT); break;
						case VK_CONTROL: pPrefixKey = kvk + (kvk[VK_RCONTROL].is_down ? VK_RCONTROL : VK_LCONTROL); break;
						case VK_MENU: pPrefixKey = kvk + (kvk[VK_RMENU].is_down ? VK_RMENU : VK_LMENU); break;
						default: pPrefixKey = kvk + kvk[vk_neutral].ModifierVK[i].vk;
						}
						pPrefixKey->was_just_used = AS_PREFIX_FOR_HOTKEY;
						GET_HOTKEY_ID_AND_FLAGS(kvk[vk_neutral].ModifierVK[i].id_with_flags);
						break;
					}
 				if (hotkey_id == HOTKEY_ID_INVALID)  // Now check scan codes since above didn't find one.
					for (i = 0; i < kvk[vk_neutral].nModifierSC; ++i)
						if (ksc[kvk[vk_neutral].ModifierSC[i].sc].is_down)
						{
							pPrefixKey = ksc + kvk[vk_neutral].ModifierSC[i].sc;
							pPrefixKey->was_just_used = AS_PREFIX_FOR_HOTKEY;
							GET_HOTKEY_ID_AND_FLAGS(kvk[vk_neutral].ModifierSC[i].id_with_flags);
							break;
						}
			}
		}

		// Alt-tab: handled directly here rather than via posting a message back to the main window.
		// In part, this is because it would be difficult to design a way to tell the main window when
		// to release the alt-key:
		if (hotkey_id == HOTKEY_ID_ALT_TAB || hotkey_id == HOTKEY_ID_ALT_TAB_SHIFT)
		{
			// Not sure if it's necessary to set this in this case.  Review.
			if (!key_up) // i.e. don't do this for key-up events that triggered an action.
				this_key.down_performed_action = true;
		
			if (   !(g_modifiersLR_logical & MOD_LALT) && !(g_modifiersLR_logical & MOD_RALT)   )  // Neither alt-key is down.
				// Note: Don't set the ignore-flag in this case because we want the hook to notice it.
				// UPDATE: It might be best, after all, to have the hook ignore these keys.  That's because
				// we want to avoid any possibility that other hotkeys will fire off while the user is
				// alt-tabbing (though we can't stop that from happening if they were registered with
				// RegisterHotkey).  In other words, since the
				// alt-tab window is in the foreground until the user released the substitute-alt key,
				// don't allow other hotkeys to be activated.  One good example that this helps is the case
				// where <key1> & rshift is defined as alt-tab but <key1> & <key2> is defined as shift-alt-tab.
				// In that case, if we didn't ignore these events, one hotkey might unintentionally trigger
				// the other.
				KeyEvent(KEYDOWN, VK_MENU);
				// And leave it down until a key-up event on the prefix key occurs.
#ifdef INCLUDE_KEYBD_HOOK
			if (vk == VK_LCONTROL || vk == VK_RCONTROL)
				// Even though this suffix key would have been suppressed, it seems that the
				// OS's alt-tab functionality sees that it's down somehow and thus this is necessary
				// to allow the alt-tab menu to appear.  This doesn't need to be done for any other
				// modifier than Control, nor any normal key since I don't think normal keys
				// being in a down-state causes any problems with alt-tab:
				KeyEvent(KEYUP, vk, sc);
#endif
			// Update the prefix key's
			// flag to indicate that it was this key that originally caused the alt-key to go down,
			// so that we know to set it back up again when the key is released.  UPDATE: Actually,
			// it's probably better if this flag is set regardless of whether ALT is already down.
			// That way, in case it's state go stuck down somehow, it will be reset by an Alt-TAB
			// (i.e. alt-tab will always behave as expected even if ALT was down before starting).
			// Note: pPrefixKey must already be non-NULL or this couldn't be an alt-tab event:
			pPrefixKey->it_put_alt_down = true;
			if (hotkey_id == HOTKEY_ID_ALT_TAB_SHIFT)
			{
				if (!(g_modifiersLR_logical & MOD_LSHIFT) && !(g_modifiersLR_logical & MOD_RSHIFT))  // Neither shift-key is down.
					KeyEvent(KEYDOWN, VK_SHIFT);  // Same notes apply to this key.
				pPrefixKey->it_put_shift_down = true;
			}
			// And this may do weird things if VK_TAB itself is already assigned a as a naked hotkey, since
			// it will recursively call the hook, resulting in the launch of some other action.  But it's hard
			// to imagine someone ever reassigning the naked VK_TAB key (i.e. with no modifiers).
			// UPDATE: The new "ignore" method should prevent that.  Or in the case of low-level hook:
			// keystrokes sent by our own app by default will not fire hotkeys.  UPDATE: Even though
			// the LL hook will have suppressed this key, it seems that the OS's alt-tab menu uses
			// some weird method (apparently not GetAsyncState(), because then our attempt to put
			// it up would fail) to determine whether the shift-key is down, so we need to still do this:
			else if (hotkey_id == HOTKEY_ID_ALT_TAB) // i.e. it's not shift-alt-tab
			{
				// Force it to be alt-tab as the user intended.
#ifdef INCLUDE_KEYBD_HOOK
				if (vk == VK_LSHIFT || vk == VK_RSHIFT)  // Needed.  See above comments. vk == VK_SHIFT not needed.
					// If a shift key is the suffix key, this must be done every time,
					// not just the first:
					KeyEvent(KEYUP, vk, sc);
				// UPDATE: Don't do "else" because sometimes the opposite key may be down, so the
				// below needs to be unconditional:
				//else
#endif
				if ((g_modifiersLR_logical & MOD_LSHIFT) || (g_modifiersLR_logical & MOD_RSHIFT))
					// In this case, it's not necessary to put the shift key back down because the
					// alt-tab menu only disappears after the prefix key has been released (and it's
					// not realistic that a user would try to trigger another hotkey while the
					// alt-tab menu is visible).  In other words, the user will be releasing the
					// shift key anyway as part of the alt-tab process, so it's not necessary to put
					// it back down for the user here (the shift stays in effect as a prefix for us
					// here because it's sent as an ignore event -- but the prefix will be correctly
					// canceled when the user releases the shift key).
					KeyEvent(KEYUP, (g_modifiersLR_logical & MOD_RSHIFT) ? VK_RSHIFT : VK_LSHIFT);
			}
			if ((g_modifiersLR_logical & MOD_LCONTROL) || (g_modifiersLR_logical & MOD_RCONTROL))
				// Any down control key prevents alt-tab from working.  This is similar to
				// what's done for the shift-key above, so see those comments for details.
				// Note: Since this is the low-level hook, the current OS must be something
				// beyond other than Win9x, so there's no need to conditionally send'
				// VK_CONTROL instead of the left/right specific key of the pair:
				KeyEvent(KEYUP, (g_modifiersLR_logical & MOD_RCONTROL) ? VK_RCONTROL : VK_LCONTROL);
			KeyEvent(KEYDOWNANDUP, VK_TAB);
			if (hotkey_id == HOTKEY_ID_ALT_TAB_SHIFT && pPrefixKey->it_put_shift_down
				&& ((vk >= VK_NUMPAD0 && vk <= VK_NUMPAD9) || vk == VK_DECIMAL)) // dual-state numpad key.
			{
				// In this case, if there is a numpad key involved, it's best to put the shift key
				// back up in between every alt-tab to avoid problems caused due to the fact that
				// the shift key being down would CHANGE the VK being received when the key is
				// released (due to the fact that SHIFT temporarily disables numlock).
				KeyEvent(KEYUP, VK_SHIFT);
				pPrefixKey->it_put_shift_down = false;  // Reset for next time since we put it back up already.
			}
			pKeyHistoryCurr->event_type = 'h'; // h = hook hotkey (not one registered with RegisterHotkey)
			return SuppressThisKey;
		} // end of alt-tab section.
	} // end of section that searches for a suffix modified by the prefix that's currently held down.

	if (hotkey_id == HOTKEY_ID_INVALID)  // Keep checking since above didn't find one.
	{
		modifiersLRnew = g_modifiersLR_logical_non_ignored;
#ifdef INCLUDE_KEYBD_HOOK
		if (this_key.as_modifiersLR)
			// Hotkeys are not defined to modify themselves, so look for a match accordingly.
			modifiersLRnew &= ~this_key.as_modifiersLR;
		GET_HOTKEY_ID_AND_FLAGS(ksc[sc].sc_takes_precedence ? Kscm(modifiersLRnew, sc) : Kvkm(modifiersLRnew, vk));
		// Bug fix for v1.0.20: The below second attempt is no longer made if the current keystroke
		// is a tab-down/up  This is because doing so causes any naked TAB that has been defined as
		// a hook hotkey to incorrectly fire when the user holds down ALT and presses tab two or more
		// times to advance through the alt-tab menu.  Here is the sequence:
		// $TAB is defined as a hotkey in the script.
		// User holds down ALT and presses TAB two or more times.
		// The Alt-tab menu becomes visible on the first TAB keystroke.
		// The $TAB hotkey fires on the second keystroke because of the below (now-fixed) logic.
		// By the way, the overall idea behind the below might be considered faulty because
		// you could argue that non-modified hotkeys should never be allowed to fire while ALT is
		// down just because the alt-tab menu is visible.  However, it seems justified because
		// the benefit (which I believe was originally and particularly that an unmodified mouse button
		// or wheel hotkey could be used to advance through the menu even though ALT is artificially
		// down due to support displaying the menu) outweighs the cost, which seems low since
		// it would be rare that anyone would press another hotkey while they are navigating through
		// the Alt-Tab menu.
		if (hotkey_id == HOTKEY_ID_INVALID && alt_tab_menu_is_visible && vk != VK_TAB)
		{
			// Try again, this time without the ALT key in case the user is trying to
			// activate an alt-tab related key (i.e. a special hotkey action such as AltTab
			// that relies on the Alt key being logically but not physically down).
			modifiersLRnew &= ~(MOD_LALT | MOD_RALT);
			GET_HOTKEY_ID_AND_FLAGS(ksc[sc].sc_takes_precedence ? Kscm(modifiersLRnew, sc) : Kvkm(modifiersLRnew, vk));
		}
#else // Mouse hook:
		GET_HOTKEY_ID_AND_FLAGS(Kvkm(g_modifiersLR_logical_non_ignored, vk));
		if (hotkey_id == HOTKEY_ID_INVALID && alt_tab_menu_is_visible)
		{
			modifiersLRnew &= ~(MOD_LALT | MOD_RALT);
			GET_HOTKEY_ID_AND_FLAGS(Kvkm(modifiersLRnew, vk));
		}
#endif
		if (hotkey_id == HOTKEY_ID_INVALID)
		{
			// Even though at this point this_key is a valid suffix, no actionable ModifierVK/SC
			// or modifiers were pressed down, so just let the system process this normally
			// (except if it's a toggleable key).  This case occurs whenever a suffix key (which
			// is also a prefix) is released but the key isn't configured to perform any action
			// upon key-release.  Currently, I think the only way a key-up event will result
			// in a hotkey action is for the release of a naked/modifierless prefix key.
			// Example of a configuration that would result in this case whenever Rshift alone
			// is pressed then released:
			// RControl & RShift = Alt-Tab
			// RShift & RControl = Shift-Alt-Tab
			if (key_up)
				// These sequence is basically the same as the one used in Case #3
				// when a prefix key that isn't a suffix failed to modify anything
				// and was then released, so make modifications made here or there
				// are considered for inclusion in the other one.  UPDATE: Since
				// the previous sentence is a bit obsolete, describe this better:
				// If it's a toggleable key that the user wants to allow to be
				// toggled, just allow this up-event to go through because the
				// previous down-event for it (in its role as a prefix) would not
				// have been suppressed:
#ifdef INCLUDE_KEYBD_HOOK
				// NO_SUPPRESS_PREFIX can occur if it fell through from Case #3 but the right
				// modifier keys aren't down to have triggered a key-up hotkey:
				return (this_key.as_modifiersLR || (this_key.no_suppress & NO_SUPPRESS_PREFIX)
					// The order on this line important; it relies on short-circuit boolean:
					|| (this_key.pForceToggle && *this_key.pForceToggle == NEUTRAL))
					? AllowKeyToGoToSystem : SuppressThisKey;
#else
				return (this_key.no_suppress & NO_SUPPRESS_PREFIX) ? AllowKeyToGoToSystem : SuppressThisKey;
#endif
			// For execution to have reached this point, the current key must be both a prefix and
			// suffix, but be acting in its capacity as a suffix.  Since no hotkey action will fire,
			// and since the key wasn't used as a prefix, I think that must mean that not all of
			// the required modifiers aren't present.  Example "a & b = calc ... lshift & a = notepad".
			// In that case, if the 'a' key is pressed and released by itself, perhaps its native
			// function should be performed by suppressing this key-up event, replacing it with a
			// down and up of our own.  However, it seems better not to do this, for now, since this
			// is really just a subset of allowing all prefixes to perform their native functions
			// upon key-release their value of was_just_used is false, which is probably
			// a bad idea in many cases (e.g. if user configures VK_VOLUME_MUTE button to be a
			// prefix, it might be undesirable for the volume to be muted if the button is pressed
			// but the user changes his mind and doesn't use it to modify anything, so just releases
			// it (at least it seems that I do this).  In any case, this default behavior can be
			// changed by explicitly configuring 'a', in the example above, to be "Send, a".
			// Here's a more complete example:
			// a & b = notepad
			// lcontrol & a = calc
			// a = Send, a
			// So the bottom line is that by default, a prefix key's native function is always
			// suppressed except if it's a toggleable key such as num/caps/scroll-lock.
			return AllowKeyToGoToSystem;
		}
	}
	// Now everything is in place for an action to be performed:

	// If only a windows key was held down (and no other modifiers) to activate this hotkey,
	// suppress the next win-up event so that the start menu won't appear (if other modifiers are present,
	// there's no need to do this because the Start Menu doesn't appear, at least on WinXP).
	// The appearance of the Start Menu would be caused by the fact that the hotkeys suffix key
	// was suppressed, therefore the OS does not see that the WIN key "modified" anything while
	// it was held down.  Note that if the WIN key is auto-repeating due to the user having held
	// it down long enough, when the user presses the hotkey's suffix key, the auto-repeating
	// stops, probably because auto-repeat is a very low-level feature.  It's also interesting
	// that unlike non-modifier keys such as letters, the auto-repeat does not resume after
	// the user releases the suffix, even if the WIN key is kept held down for a long time.
	// When the user finally releases the WIN key, that release will be disguised if called
	// for by the logic below.
	if (!(g_modifiersLR_logical & ~(MOD_LWIN | MOD_RWIN))) // Only lwin, rwin, both, or neither are currently down.
	{
		// If it's used as a prefix, there's no need (and it would probably break something)
		// to disguise the key this way since the prefix-handling logic already does that
		// whenever necessary:
		if ((g_modifiersLR_logical & MOD_LWIN) && !kvk[VK_LWIN].used_as_prefix)
			disguise_next_lwin_up = true;
		if ((g_modifiersLR_logical & MOD_RWIN) && !kvk[VK_RWIN].used_as_prefix)
			disguise_next_rwin_up = true;
	}
	// For maximum reliability on the maximum range of systems, it seems best to do the above
	// for ALT keys also, to prevent them from invoking the icon menu or menu bar of the
	// foreground window (rarer than the Start Menu problem, above, I think).
	// Update for v1.0.25: This is usually only necessary for hotkeys whose only modifier is ALT.
	// For example, Shift-Alt hotkeys do not need it if Shift is pressed after Alt because Alt
	// "modified" the shift so the OS knows it's not a naked ALT press to activate the menu bar.
	// Conversely, if Shift is pressed prior to Alt, but released before Alt, I think the shift-up
	// counts as a "modification" and the same rule applies.  However, if shift is released after Alt,
	// that would activate the menu bar unless the ALT key is disguised below.  This issue does
	// not apply to the WIN key above because apparently it is disguised automatically
	// whenever some other modifier was involved with it in any way and at any time during the
	// keystrokes that comprise the hotkey.
	else if ((g_modifiersLR_logical & MOD_LALT) && !kvk[VK_LMENU].used_as_prefix)
		if (g_KeybdHook)
			disguise_next_lalt_up = true;
		else
			// Since no keyboard hook, no point in setting the variable because it would never be acted up.
			// Instead, disguise the key now with a CTRL keystroke. Note that this is not done for
			// mouse buttons that use the WIN key as a prefix because it does not work reliably for them
			// (i.e. sometimes the Start Menu appears, even if two CTRL keystrokes are sent rather than one).
			// Therefore, as of v1.0.25.05, mouse button hotkeys that use only the WIN key as a modifier cause
			// the keyboard hook to be installed.  This determine is made during the hotkey loading stage.
			KeyEvent(KEYDOWNANDUP, VK_CONTROL);
	else if ((g_modifiersLR_logical & MOD_RALT) && !kvk[VK_RMENU].used_as_prefix)
		// The two else if's above: If it's used as a prefix, there's no need (and it would probably break something)
		// to disguise the key this way since the prefix-handling logic already does that whenever necessary.
		if (g_KeybdHook)
			disguise_next_ralt_up = true;
		else
			KeyEvent(KEYDOWNANDUP, VK_CONTROL);

	switch (hotkey_id)
	{
		case HOTKEY_ID_ALT_TAB_MENU_DISMISS: // This case must occur before HOTKEY_ID_ALT_TAB_MENU due to non-break.
			if (!alt_tab_menu_is_visible)
				// Even if the menu really is displayed by other means, we can't easily detect it
				// because it's not a real window?
				return AllowKeyToGoToSystem;  // Let the key do its native function.
			// else fall through to the next case.
		case HOTKEY_ID_ALT_TAB_MENU:  // These cases must occur before the Alt-tab ones due to conditional break.
		case HOTKEY_ID_ALT_TAB_AND_MENU:
		{
			vk_type which_alt_down = 0;
			if (g_modifiersLR_logical & MOD_LALT)
				which_alt_down = VK_LMENU;
			else if (g_modifiersLR_logical & MOD_RALT)
				which_alt_down = VK_RMENU;

			if (alt_tab_menu_is_visible)  // Can be true even if which_alt_down is zero.
			{
				if (hotkey_id != HOTKEY_ID_ALT_TAB_AND_MENU) // then it is MENU or DISMISS.
				{
					// Since it is possible for the menu to be visible when neither ALT
					// key is down, always send an alt-up event if one isn't down
					// so that the menu is dismissed as intended:
					KeyEvent(KEYUP, which_alt_down ? which_alt_down : VK_MENU);
					if (this_key.as_modifiersLR && vk != VK_LWIN && vk != VK_RWIN)
						// Something strange seems to happen with the foreground app
						// thinking the modifier is still down (even though it was suppressed
						// entirely [confirmed!]).  For example, if the script contains
						// the line "lshift::AltTabMenu", pressing lshift twice would
						// otherwise cause the newly-activated app to think the shift
						// key is down.  Sending an extra UP here seems to fix that,
						// hopefully without breaking anything else.  Note: It's not
						// done for Lwin/Rwin because most (all?) apps don't care whether
						// LWin/RWin is down, and sending an up event might risk triggering
						// the start menu in certain hotkey configurations.  This policy
						// might not be the right one for everyone, however:
						KeyEvent(KEYUP, vk); // Can't send sc here since it's not defined for the mouse hook.
					alt_tab_menu_is_visible = false;
					break;
				}
				// else HOTKEY_ID_ALT_TAB_AND_MENU, do nothing (don't break) because we want
				// the switch to fall through to the Alt-Tab case.
			}
			else // alt-tab menu is not visible
			{
				// Unlike CONTROL, SHIFT, AND ALT, the LWIN/RWIN keys don't seem to need any
				// special handling to make them work with the alt-tab features.

				bool vk_is_alt = vk == VK_LMENU || vk == VK_RMENU;  // Tranlated & no longer needed: || vk == VK_MENU;
				bool vk_is_shift = vk == VK_LSHIFT || vk == VK_RSHIFT;  // || vk == VK_SHIFT;
				bool vk_is_control = vk == VK_LCONTROL || vk == VK_RCONTROL;  // || vk == VK_CONTROL;

				vk_type which_shift_down = 0;
				if (g_modifiersLR_logical & MOD_LSHIFT)
					which_shift_down = VK_LSHIFT;
				else if (g_modifiersLR_logical & MOD_RSHIFT)
					which_shift_down = VK_RSHIFT;
				else if (!key_up && vk_is_shift)
					which_shift_down = vk;

				vk_type which_control_down = 0;
				if (g_modifiersLR_logical & MOD_LCONTROL)
					which_control_down = VK_LCONTROL;
				else if (g_modifiersLR_logical & MOD_RCONTROL)
					which_control_down = VK_RCONTROL;
				else if (!key_up && vk_is_control)
					which_control_down = vk;

				bool shift_put_up = false;
				if (which_shift_down)
				{
					KeyEvent(KEYUP, which_shift_down);
					shift_put_up = true;
				}

				bool control_put_up = false;
				if (which_control_down)
				{
					// In this case, the control key must be put up because the OS, at least
					// WinXP, knows the control key is down even though the down event was
					// suppressed by the hook.  So put it up and leave it up, because putting
					// it back down would cause it to be down even after the user releases
					// it (since the up-event of a hotkey is also suppressed):
					KeyEvent(KEYUP, which_control_down);
					control_put_up = true;
				}

				// Alt-tab menu is not visible, or was not made visible by us.  In either case,
				// try to make sure it's displayed:
				// Don't put alt down if it's already down, it might mess up cases where the
				// ALT key itself is assigned to be one of the alt-tab actions:
				if (vk_is_alt)
					if (key_up)
						// The system won't see it as down for the purpose of alt-tab, so remove this
						// modifier from consideration.  This is necessary to allow something like this
						// to work:
						// LAlt & WheelDown::AltTab
						// LAlt::AltTabMenu   ; Since LAlt is a prefix key above, it will be a key-up hotkey here.
						which_alt_down = 0;
					else // Because there hasn't been a chance to update g_modifiersLR_logical yet:
						which_alt_down = vk;
				if (!which_alt_down)
					KeyEvent(KEYDOWN, VK_MENU); // Use the generic/neutral ALT key so it works with Win9x.

				KeyEvent(KEYDOWN, VK_TAB);
				// Only put it put it back down if it wasn't the hotkey itself, because
				// the system would never have known it was down because the down-event
				// on the hotkey would have been suppressed.  And since the up-event
				// will also be suppressed, putting it down like this would result in
				// it being permanently down even after the user releases the key!:
				if (shift_put_up && !vk_is_shift) // Must do this regardless of the value of key_up.
					KeyEvent(KEYDOWN, which_shift_down);
				
				// Update: Can't do this one because going down on control will instantly
				// dismiss the alt-tab menu, which we don't want if we're here.
				//if (control_put_up && !vk_is_control) // Must do this regardless of the value of key_up.
				//	KeyEvent(KEYDOWN, which_control_down);

				// At this point, the alt-tab menu has displayed and advanced by one icon
				// (to the next window in the z-order).  Rather than sending a shift-tab to
				// go back to the first icon in the menu, it seems best to leave it where
				// it is because usually the user will want to go forward at least one item.
				// Going backward through the menu is a lot more rare for most people.
				alt_tab_menu_is_visible = true;
				break;
			}
		}
		case HOTKEY_ID_ALT_TAB:
		case HOTKEY_ID_ALT_TAB_SHIFT:
		{
			// Since we're here, this ALT-TAB hotkey didn't have a prefix or it would have
			// already been handled and we would have returned above.  Therefore, this
			// hotkey is defined as taking effect only if the alt-tab menu is currently
			// displayed, otherwise it will just be passed through to perform it's native
			// function.  Example:
			// MButton::AltTabMenu
			// WheelDown::AltTab     ; But if the menu is displayed, the wheel will function normally.
			// WheelUp::ShiftAltTab  ; But if the menu is displayed, the wheel will function normally.
			if (!alt_tab_menu_is_visible)
				// Even if the menu really is displayed by other means, we can't easily detect it
				// because it's not a real window?
				return AllowKeyToGoToSystem;

			// Unlike CONTROL, SHIFT, AND ALT, the LWIN/RWIN keys don't seem to need any
			// special handling to make them work with the alt-tab features.

			// Must do this to prevent interference with Alt-tab when these keys
			// are used to do the navigation.  Don't put any of these back down
			// after putting them up since that would probably cause them to become
			// stuck down due to the fact that the user's physical release of the
			// key will be suppressed (since it's a hotkey):
			if (!key_up && (vk == VK_LCONTROL || vk == VK_RCONTROL || vk == VK_LSHIFT || vk == VK_RSHIFT))
				// Don't do the ALT key because it causes more problems than it solves
				// (in fact, it might not solve any at all).
				KeyEvent(KEYUP, vk); // Can't send sc here since it's not defined for the mouse hook.

			// Even when the menu is visible, it's possible that neither of the ALT keys
			// is down, at least under XP (probably NT and 2k also).  Not sure about Win9x:
			if (   !((g_modifiersLR_logical & MOD_LALT) || (g_modifiersLR_logical & MOD_RALT))
				|| (key_up && (vk == VK_LMENU || vk == VK_RMENU))   )
				KeyEvent(KEYDOWN, VK_MENU);
				// And never put it back up because that would dismiss the menu.
			// Otherwise, use keystrokes to navigate through the menu:
			bool shift_put_down = false;
			if (   hotkey_id == HOTKEY_ID_ALT_TAB_SHIFT
				&& !((g_modifiersLR_logical & MOD_LSHIFT) || (g_modifiersLR_logical & MOD_RSHIFT))   )
			{
				KeyEvent(KEYDOWN, VK_SHIFT);
				shift_put_down = true;
			}
			KeyEvent(KEYDOWNANDUP, VK_TAB);
			if (shift_put_down)
				KeyEvent(KEYUP, VK_SHIFT);
			break;
		}
		default:
		// UPDATE to below: Since this function is only called from a single thread (namely ours),
		// albeit recursively, it's apparently not reentrant (unless our own main app itself becomes
		// multithreaded someday, and even then it might not matter?) there's no advantage to using
		// PostMessage() because the message can't be acted upon until after we return from this
		// function.  Therefore, avoid the overhead (and possible delays while system is under
		// heavy load?) of using PostMessage and simply execute the hotkey right here before
		// returning.  UPDATE AGAIN: No, I don't think this will work reliably because this
		// function is called invisibly by GetMessage(), without it even telling us that
		// it's calling it.  Therefore, if we call a subroutine in the script from here,
		// we can't return until after the subroutine is over, thus GetMessage() will
		// probably hang, or be forced to start a new thread or something?  An alternative to
		// using PostMessage() (if it really is susceptible to delays due to system being under
		// load, which is far from certain), is to change the value of a global var to signal
		// to MsgSleep() that a hotkey has been fired.  However, this doesn't seem likely
		// to work because a call to GetMessage() will likely call this function without
		// actually returning any messages to its caller, thus the hotkeys would never be
		// seen during periods when there are no messages.  PostMessage() works reliably, so
		// it seems best not to change it without good reason and without a full understanding
		// of what's really going on.
#ifdef INCLUDE_KEYBD_HOOK
			PostMessage(g_hWnd, AHK_HOOK_HOTKEY, hotkey_id, 0);  // Returns non-zero on success.
#else
			// In the case of a mouse hotkey whose native function the user didn't want suppressed,
			// tell our hotkey handler to also dismiss any menus that the mouseclick itself may
			// have invoked:
			PostMessage(g_hWnd, AHK_HOOK_HOTKEY, hotkey_id, no_suppress);
#endif
			// Don't execute it directly because if whatever it does takes a long time, this keystroke
			// and instance of the function will be left hanging until it returns:
			//Hotkey::PerformID(hotkey_id);
	}

	pKeyHistoryCurr->event_type = 'h'; // h = hook hotkey (not one registered with RegisterHotkey)

#ifdef INCLUDE_KEYBD_HOOK
	if (key_up && this_key.used_as_prefix && this_key.pForceToggle) // Key is a toggleable key.
		if (*this_key.pForceToggle == NEUTRAL) // Dereference to get the global var's value.
		{
			// In this case, since all the above conditions are true, the key-down
			// event for this key-up (which fired a hotkey) would not have been
			// suppressed.  Thus, we should toggle the state of the key back
			// the what it was before the user pressed it (due to the policy that
			// the natural function of a key should never take effect when that
			// key is used as a hotkey suffix).  You could argue that instead
			// of doing this, we should *pForceToggle's value to make the
			// key untoggleable whenever it's both a prefix and a naked
			// (key-up triggered) suffix.  However, this isn't too much harder
			// and has the added benefit of allowing the key to be toggled if
			// a modifier is held down before it (e.g. alt-CapsLock would then
			// be able to toggle the CapsLock key):
			KEYEVENT_PHYS(KEYUP, vk, sc); // Mark it as physical for any other hook instances.
			KeyEvent(KEYDOWNANDUP, vk, sc);
			return SuppressThisKey;
		}

	if (this_key.as_modifiersLR && key_up)
		// Since this hotkey is fired on a key-up event, and since it's a modifier, must
		// not suppress the key because otherwise the system's state for this modifier
		// key would be stuck down due to the fact that the previous down-event for this
		// key (which is presumably a prefix *and* a suffix) was not suppressed:
		return AllowKeyToGoToSystemButDisguiseWinAlt;
#endif

	if (key_up)
	{
		if (no_suppress) // Plus we know it's not a modifier since otherwise it would've returned above.
		{
			// Since this hotkey is firing on key-up but the user specified not to suppress its native
			// function, send a down event to make up for the fact that the original down event was
			// suppressed (since key-up hotkeys' down events are always suppressed because they
			// are also prefix keys by definition).  UPDATE: Now that it is possible for a prefix key
			// to be non-suppressed, this is done only if the prior down event wasn't suppressed:
#ifdef INCLUDE_KEYBD_HOOK
			if (!(this_key.no_suppress & NO_SUPPRESS_PREFIX))
				KeyEvent(KEYDOWN, vk, sc);
				// Now allow the up-event to go through.  The DOWN should always wind up taking effect
				// before the UP because the above should already have "finished" by now, since
				// it resulted in a recursive call to this function (using our current thread
				// rather than some other re-entrant thread):
//#else // Mouse hook.
// Currently not supporting the mouse buttons for the above method, becuase KeyEvent()
// doesn't support the translation of a mouse-VK into a mouse_event() call.
// Such a thing might not work anyway because the hook probably received extra
// info such as the location where the mouse click should occur and other things.
// That info plus anything else relevant in MSLLHOOKSTRUCT would have to be
// translated into the correct info for a call to mouse_event().
#endif
			return AllowKeyToGoToSystem;
		}
	}
	else // Key Down
	{
		// Do this only for DOWN (not UP) events that triggered an action:
		this_key.down_performed_action = true;
		// Also update this in case the currently-down Prefix key is both a modifier
		// and a normal prefix key (in which case it isn't stored in this_key's array
		// of VK and SC prefixes, so this value wouldn't have yet been set).
		// Update: The below is done even if &prefix_key != &this_key, which happens
		// when we reached this point after having fallen through from Case #1 above.
		// The reason for this is that we just fired a hotkey action for this key,
		// so we don't want it's action to fire again upon key-up:
		if (pPrefixKey)
			pPrefixKey->was_just_used = AS_PREFIX_FOR_HOTKEY;
		if (no_suppress) // Plus we know it's not a modifier since otherwise it would've returned above.
		{
			// Since this hotkey is firing on key-down but the user specified not to suppress its native
			// function, substitute an DOWN+UP pair of events for this event, since we want the
			// DOWN to precede the UP.  It's necessary to send the UP because the user's physical UP
			// will be suppressed automatically when this function is called for that event.
			// UPDATE: The below method causes side-effects due to the fact that it is simulated
			// input vs. physical input, e.g. when used with the Input command, which distinguishes
			// between "ignored" and physical input.  Therefore, let this down event pass through
			// and set things up so that the corresponding up-event is also not suppressed:
			//KeyEvent(KEYDOWNANDUP, vk, sc);
			// No longer relevant due to the above change:
			// Now let it just fall through to suppress this down event, because we can't use it
			// since doing so would result in the UP event having preceded the DOWN, which would
			// be the wrong order.
			this_key.no_suppress |= NO_SUPPRESS_NEXT_UP_EVENT;
			return AllowKeyToGoToSystem;
		}
#ifdef INCLUDE_KEYBD_HOOK
		else if (vk == VK_LMENU || vk == VK_RMENU)
			// Since this is a hotkey that fires on ALT-DOWN and it's a normal (suppressed) hotkey,
			// send an up-event to "turn off" the OS's low-level handling for the alt key with
			// respect to having it modify keypresses.  For example, the following hotkey would
			// fail to work properly without this workaround because the OS apparently sees that
			// the ALT key is physically down even though it is not logically down:
			// RAlt::Send f  ; Actually triggers !f, which activates the FILE menu if the active window has one.
			// RAlt::Send {PgDn}  ; Fails to work because ALT-PgDn usually does nothing.
			// NOTE: The above is something of a separate issue than the "Alt triggers the menu
			// bar" problem noted in the FAQ, since that has to do with the fact that modifiers
			// are never suppressed if they are prefixes, and thus cause the menu bar to be
			// activated.  Some other type of workaround would be needed for that, but such
			// a workaround might not be possible without breaking some existing scripts that
			// rely on the current ALT key prefix behavior).
			KeyEvent(KEYUP, vk, sc);
#endif
	}
	
	// Otherwise:
	return SuppressThisKey;
}
