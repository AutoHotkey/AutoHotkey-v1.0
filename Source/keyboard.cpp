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
#include "keyboard.h"
#include "globaldata.h" // for g.KeyDelay
#include "application.h" // for MsgSleep()
#include "util.h"  // for strlicmp()
#include "window.h" // for IsWindowHung()



// Added for v1.0.25.  Search on sPrevEventType for more comments:
static KeyEventTypes sPrevEventType;
static vk_type sPrevVK = 0;



inline void DoKeyDelay(int aDelay = g.KeyDelay)
// A small inline to help with tracking things in our effort to track the physical
// state of the modifier keys, since GetAsyncKeyState() does not appear to be
// reliable (not properly implemented), at least on Windows XP.  UPDATE: Tracking the
// modifiers this way would require that the hotkey's modifiers be put back down between
// every key event, which can sometimes interfere with the send itself (i.e. since the ALT
// key can activate the menu bar in the foreground window).  So using the new physical
// modifier tracking method instead.
{
	if (aDelay < 0) // To support user-specified KeyDelay of -1 (fastest send rate).
		return;
	if (g_os.IsWin9x())
	{
		// Do a true sleep on Win9x because the MsgSleep() method is very inaccurate on Win9x
		// for some reason (a MsgSleep(1) causes a sleep between 10 and 55ms, for example):
		Sleep(aDelay);
		return;
	}
	SLEEP_WITHOUT_INTERRUPTION(aDelay);
}



void SendKeys(char *aKeys, bool aSendRaw, HWND aTargetWindow)
// The aKeys string must be modifiable (not constant), since for performance reasons,
// it's allowed to be temporarily altered by this function.  mThisHotkeyModifiersLR, if non-zero,
// is the set of modifiers used to trigger the hotkey that called the subroutine
// containing the Send that got us here.  If any of those modifiers are still down,
// they will be released prior to sending the batch of keys specified in <aKeys>.
{
	if (!aKeys || !*aKeys) return;

	// Maybe best to call immediately so that the amount of time during which we haven't been pumping
	// messsages is more accurate:
	LONG_OPERATION_INIT

	// Below is now called with "true" so that the hook's modifier state will be corrected (if necessary)
	// prior to every send.
	modLR_type modifiersLR_current = GetModifierLRState(true); // Current "logical" modifier state.

	// Make a best guess of what the physical state of the keys is prior to starting,
	// since GetAsyncKeyState() is unreliable (it seems to always report the logical vs.
	// physical state, at least under Windows XP).  Note: We're only want those physical
	// keys that are also logically down (it's possible for a key to be down physically
	// but not logically such as well R-control, for example, is a suffix hotkey and the
	// user is physically holding it down):
	modLR_type modifiersLR_down_physically_and_logically, modifiersLR_down_physically_but_not_logically;
	if (g_KeybdHook)
	{
		// Since hook is installed, use its more reliable tracking to determine which
		// modifiers are down.
		// Update: modifiersLR_down_physically_but_not_logically is now used to distinguish
		// between the following two cases, allowing modifiers to be properly restored to
		// the down position when the hook is installed:
		// 1) naked modifier key used only as suffix: when the user phys. presses it, it isn't
		//    logically down because the hook suppressed it.
		// 2) A modifier that is a prefix, that triggers a hotkey via a suffix, and that hotkey sends
		//    that modifier.  The modifier will go back up after the SEND, so the key will be physically
		//    down but not logically.
		modifiersLR_down_physically_but_not_logically = g_modifiersLR_physical & ~g_modifiersLR_logical;
		modifiersLR_down_physically_and_logically = g_modifiersLR_physical & g_modifiersLR_logical; // intersect
	}
	else // Use best-guess instead.
	{
		modifiersLR_down_physically_but_not_logically = 0; // There's no way of knowing, so assume none.
		// Even if TickCount has wrapped due to system being up more than about 49 days,
		// DWORD math still gives the right answer as long as g_script.mThisHotkeyStartTime
		// itself isn't more than about 49 days ago:
		if ((GetTickCount() - g_script.mThisHotkeyStartTime) < (DWORD)g_HotkeyModifierTimeout) // Elapsed time < timeout-value
			modifiersLR_down_physically_and_logically = modifiersLR_current & g_script.mThisHotkeyModifiersLR; // Bitwise AND is set intersection.
		else
			// Since too much time as passed since the user pressed the hotkey, it seems best,
			// based on the action that will occur below, to assume that no hotkey modifiers
			// are physically down:
			modifiersLR_down_physically_and_logically = 0;
	}
	// Any of the external modifiers that are down but NOT due to the hotkey are probably
	// logically down rather than physically (perhaps from a prior command such as
	// "Send, {CtrlDown}".  Since there's no way to be sure due to the unreliability of
	// GetAsyncKeyState() under XP and perhaps other OSes, it seems best to assume that
	// they are logically vs. physically down.  This value contains the modifiers that
	// we will not attempt to change (e.g. "Send, A" will not release the LWin
	// before sending "A" if this value indicates that LWin is down).  The below sets
	// the value to be all the down-keys in modifiersLR_current except any that are physically
	// down due to the hotkey itself.  UPDATE: To improve the above, we now exclude from
	// the set of persistent modifiers any that weren't made persistent by this script.
	// Such a policy seems likely to do more good than harm as there have been cases where
	// a modifier was detected as persistent just because #HotkeyModifier had timed out
	// while the user was still holding down the key, but then when the user released it,
	// this logic here would think it's still persistent and push it back down again
	// to enforce it as "always-down" during the send operation.  Thus, the key would
	// basically get stuck down even after the send was over:
	g_modifiersLR_persistent &= modifiersLR_current & ~modifiersLR_down_physically_and_logically;
	mod_type modifiers_persistent = ConvertModifiersLR(g_modifiersLR_persistent);
	// The above two variables should be kept in sync with each other from now on.

//MsgBox(GetTickCount() - g_script.mThisHotkeyStartTime);
//char mod_str[256];
//MsgBox(ModifiersLRToText(aModifiersLR, mod_str));
//MsgBox(ModifiersLRToText(modifiersLR_current, mod_str));
//MsgBox(ModifiersLRToText(modifiersLR_down_physically_and_logically, mod_str));
//MsgBox(ModifiersLRToText(g_modifiersLR_persistent, mod_str));

	// Might be better to do this prior to changing capslock state:
	bool threads_are_attached = false; // Set default.
	DWORD my_thread, target_thread;
	if (aTargetWindow)
	{
		my_thread  = GetCurrentThreadId();
		target_thread = GetWindowThreadProcessId(aTargetWindow, NULL);
		if (target_thread && target_thread != my_thread && !IsWindowHung(aTargetWindow))
			threads_are_attached = AttachThreadInput(my_thread, target_thread, TRUE) != 0;
	}

	// The default behavior is to turn the capslock key off prior to sending any keys
	// because otherwise lowercase letters would come through as uppercase:
	ToggleValueType prior_capslock_state;
	if (threads_are_attached || !g_os.IsWin9x())
		// Only under either of the above conditions can the state of Capslock be reliably
		// retrieved and changed:
		prior_capslock_state = g.StoreCapslockMode ? ToggleKeyState(VK_CAPITAL, TOGGLED_OFF) : TOGGLE_INVALID;
	else // OS is Win9x and threads are not attached.
	{
		// Attempt to turn off capslock, but never attempt to turn it back on because we can't
		// reliably detect whether it was on beforehand.  Update: This didn't do any good, so
		// it's disabled for now:
		//CapslockOffWin9x();
		prior_capslock_state = TOGGLE_INVALID;
	}

	bool blockinput_prev = g_BlockInput;
	bool do_selective_blockinput = (g_BlockInputMode == TOGGLE_SEND || g_BlockInputMode == TOGGLE_SENDANDMOUSE)
		&& !aTargetWindow && g_os.IsWinNT4orLater();
	if (do_selective_blockinput)
		Line::ScriptBlockInput(true); // Turn it on unconditionally even if it was on, since Ctrl-Alt-Del might have disabled it.

	char single_char_string[2];
	vk_type vk = 0;
	sc_type sc = 0;
	modLR_type key_as_modifiersLR = 0;
	modLR_type modifiersLR_for_next_key = 0;
	// Above: For v1.0.35, it was changed to modLR vs. mod so that AltGr keys such as backslash and {
	// are supported on layouts such as German when sending to apps such as Putty that are fussy about
	// which ALT key is held down to produce the character.

	// For v1.0.25, the below is static to track it in between sends, so that the below will continue
	// to work:
	// Send {LWinDown}
	// Send {LWinUp}  ; Should still open the Start Menu even though it's a separate Send.
	static vk_type sPreviousEventModifierDown = 0;
	vk_type this_event_modifier_down;

	for (; *aKeys; ++aKeys, sPreviousEventModifierDown = this_event_modifier_down)
	{
		LONG_OPERATION_UPDATE_FOR_SENDKEYS
		this_event_modifier_down = 0; // Set default for this iteration, overridden selectively below.

		if (!aSendRaw && strchr("^+!#{}", *aKeys))
		{
			switch (*aKeys)
			{
			case '^':
				if (!(modifiers_persistent & MOD_CONTROL))
					modifiersLR_for_next_key |= MOD_LCONTROL;
				// else don't add it, because the value of modifiersLR_for_next_key may also used to determine
				// which keys to release after the key to which this modifier applies is sent.
				// We don't want persistent modifiers to ever be released because that's how
				// AutoIt2 behaves and it seems like a reasonable standard.
				continue;
			case '+':
				if (!(modifiers_persistent & MOD_SHIFT))
					modifiersLR_for_next_key |= MOD_LSHIFT;
				continue;
			case '!':
				if (!(modifiers_persistent & MOD_ALT))
					modifiersLR_for_next_key |= MOD_LALT;
				continue;
			case '#':
				if (g_script.mIsAutoIt2) // Since AutoIt2 ignores these, ignore them if script is in AutoIt2 mode.
					continue;
				if (!(modifiers_persistent & MOD_WIN))
					modifiersLR_for_next_key |= MOD_LWIN;
				continue;
			case '}': continue;  // Important that these be ignored.  Be very careful about changing this, see below.
			case '{':
			{
				char *end_pos = strchr(aKeys + 1, '}');
				if (!end_pos) // ignore it
					continue;
				size_t key_text_length = end_pos - aKeys - 1;
				if (!key_text_length)
				{
					if (end_pos[1] == '}')
					{
						// The literal string "{}}" has been encountered, which is interpreted as a single "}".
						++end_pos;
						key_text_length = 1;
					}
					else // Empty braces {} were encountered.
						continue;  // do nothing: let it proceed to the }, which will then be ignored.
				}
				size_t key_name_length = key_text_length; // Set default.

				*end_pos = '\0';  // temporarily terminate the string here.
				UINT repeat_count = 1; // Set default.
				KeyEventTypes event_type = KEYDOWNANDUP; // Set default.
				char old_char;
				char *space_pos = StrChrAny(aKeys + 1, " \t");  // Relies on the fact that {} key names contain no spaces.
				if (space_pos)
				{
					old_char = *space_pos;
					*space_pos = '\0';  // Temporarily terminate here so that TextToVK() can properly resolve a single char.
					key_name_length = space_pos - aKeys - 1; // Override the default value set above.
					char *next_word = omit_leading_whitespace(space_pos + 1);
					UINT next_word_length = (UINT)(end_pos - next_word);
					if (next_word_length > 0)
					{
						if (!stricmp(next_word, "down"))
							event_type = KEYDOWN;
						else if (!stricmp(next_word, "up"))
							event_type = KEYUP;
						else
						{
							repeat_count = ATOI(next_word);
							if (repeat_count < 0) // But seems best to allow zero itself, for possibly use with environment vars
								repeat_count = 0;
						}
					}
				}

				vk = TextToVK(aKeys + 1, &modifiersLR_for_next_key, true, false); // false must be passed due to below.
				sc = vk ? 0 : TextToSC(aKeys + 1);  // If sc is 0, it will be resolved by KeyEvent() later.
				if (!vk && !sc && toupper(*(aKeys + 1)) == 'V' && toupper(*(aKeys + 2)) == 'K')
				{
					char *sc_string = StrChrAny(aKeys + 3, "Ss"); // Look for the "SC" that demarks the scan code.
					if (sc_string && toupper(*(sc_string + 1)) == 'C')
						sc = strtol(sc_string + 2, NULL, 16);  // Convert from hex.
					// else leave sc set to zero and just get the specified VK.  This supports Send {VKnn}.
					vk = (vk_type)strtol(aKeys + 3, NULL, 16);  // Convert from hex.
				}

				if (space_pos)  // undo the temporary termination
					*space_pos = old_char;
				*end_pos = '}';  // undo the temporary termination

				// For v1.0.25, the following situation is fixed by the code below: If LWin or LAlt
				// becomes a persistent modifier (e.g. via Send {LWin down}) and the user physically
				// releases LWin immediately before: 1) the {LWin up} is scheduled; and 2) SendKey()
				// returns.  Then SendKey() will push the modifier back down so that it is in effect
				// for other things done by its caller (SendKeys) and also so that if the Send
				// operation ends, the key will still be down as the user intended (to modify future
				// keystrokes, physical or simulated).  However, since that down-event is followed
				// immediately by an up-event, the Start Menu appears for WIN-key or the active
				// window's menu bar is activated for ALT-key.  SOLUTION: Disguise Win-up and Alt-up
				// events in these cases.  This workaround has been successfully tested.  It's also
				// limited is scope so that a script can still explicitly invoke the start menu with
				// "Send {LWin}" or the menu bar with "Send {Alt}".
				// The check of sPreviousEventModifierDown allows "Send {LWinDown}{LWinUp}" etc., to
				// continue to work:
				#define DISGUISE_IF_NEEDED \
					if (sPrevEventType == KEYDOWN && sPreviousEventModifierDown != vk \
						&& ((vk == VK_LWIN || vk == VK_RWIN) && (sPrevVK == VK_LWIN || sPrevVK == VK_RWIN)\
							|| (vk == VK_LMENU || (vk == VK_RMENU && !g_LayoutHasAltGr)) && (sPrevVK == VK_LMENU || sPrevVK == VK_RMENU)))\
						KeyEvent(KEYDOWNANDUP, VK_CONTROL); // Disguise it to suppress Start Menu or prevent activation of active window's menu bar.

				if (repeat_count)
				{
					if (vk || sc)
					{
						if (key_as_modifiersLR = KeyToModifiersLR(vk, sc)) // Assign
						{
							if (!aTargetWindow)
							{
								if (event_type == KEYDOWN) // i.e. make {Shift down} have the same effect {ShiftDown}
								{
									this_event_modifier_down = vk;
									modifiers_persistent = ConvertModifiersLR(g_modifiersLR_persistent |= key_as_modifiersLR);
								}
								else if (event_type == KEYUP) // *not* KEYDOWNANDUP, since that would be an intentional activation of the Start Menu or menu bar.
								{
									DISGUISE_IF_NEEDED
									modifiers_persistent = ConvertModifiersLR(g_modifiersLR_persistent &= ~key_as_modifiersLR);
								}
								// else must never change modifiers_persistent in response to KEYDOWNANDUP
								// because that would break existing scripts.  This is because that same
								// modifier key may have been pushed down via {ShiftDown} rather than "{Shift Down}".
								// In other words, {Shift} should never undo the effects of a prior {ShiftDown}
								// or {Shift down}.
							}
							//else don't add this event to modifiers_persistent because it will not be
							// manifest via keybd_event.  Instead, it will done via less intrusively
							// (less interference with foreground window) via SetKeyboardState() and
							// PostMessage().  This change is for ControlSend in v1.0.21 and has been
							// documented.
						}
						// Below: modifiers_persistent stays in effect (pressed down) even if the key
						// being sent includes that same modifier.  Surprisingly, this is how AutoIt2
						// behaves also, which is good.  Example: Send, {AltDown}!f  ; this will cause
						// Alt to still be down after the command is over, even though F is modified
						// by Alt.
						SendKey(vk, sc, modifiersLR_for_next_key, g_modifiersLR_persistent
							, repeat_count, event_type, key_as_modifiersLR, aTargetWindow);
					}

					else if (key_name_length == 1) // No vk/sc means a char of length one is sent via special method.
						SendKeySpecial(aKeys[1], modifiersLR_for_next_key, g_modifiersLR_persistent, repeat_count
							, event_type, aTargetWindow);

					// See comment "else must never change modifiers_persistent" above about why
					// !aTargetWindow is used below:
					else if (vk = TextToSpecial(aKeys + 1, (UINT)key_text_length, event_type
						, g_modifiersLR_persistent, modifiers_persistent, !aTargetWindow)) // Assign.
					{
						if (!aTargetWindow)
						{
							if (event_type == KEYDOWN)
								this_event_modifier_down = vk;
							else // KEYUP (TextToSpecial() never returns KEYDOWNANDUP)
								DISGUISE_IF_NEEDED
						}
						for (UINT i = 0; i < repeat_count; ++i)
						{
							// Don't tell it to save & restore modifiers because special keys like this one
							// should have maximum flexibility (i.e. nothing extra should be done so that the
							// user can have more control):
							KeyEvent(event_type, vk, 0, aTargetWindow, true);
							LONG_OPERATION_UPDATE_FOR_SENDKEYS
						}
					}

					else if (key_text_length > 4 && !strnicmp(aKeys + 1, "ASC ", 4) && !aTargetWindow) // {ASC nnnnn}
					{
						// Include the trailing space in "ASC " to increase uniqueness (selectivity).
						// Also, sending the ASC sequence to window doesn't work, so don't even try:
						SendASC(omit_leading_whitespace(aKeys + 4), aTargetWindow); // aTargetWindow is always NULL, it's just for maintainability.
						// Do this only once at the end of the sequence:
						DoKeyDelay();
					}
					//else do nothing since it isn't recognized as any of the above "else if" cases (see below).
				} // repeat_count is non-zero

				// If what's between {} is unrecognized, such as {Bogus}, it's safest not to send
				// the contents literally since that's almost certainly not what the user intended.
				// In addition, reset the modifiers, since they were intended to apply only to
				// the key inside {}.  Also, the below is done even if repeat-count is zero:
				modifiersLR_for_next_key = 0;
				aKeys = end_pos;  // In prep for aKeys++ done by the loop.
				continue;
			} // case '{'
			} // switch()
		} // if (!aSendRaw && strchr("^+!#{}", *aKeys))

		else // Encountered a character other than ^+!#{} ... or we're in raw mode.
		{
			// Best to call this separately, rather than as first arg in SendKey, since it changes the
			// value of modifiers and the updated value is *not* guaranteed to be passed.
			// In other words, SendKey(TextToVK(...), modifiers, ...) would often send the old
			// value for modifiers.
			single_char_string[0] = *aKeys;
			single_char_string[1] = '\0';
			vk = TextToVK(single_char_string, &modifiersLR_for_next_key, true, true);
			sc = 0;
			if (vk)
				SendKey(vk, sc, modifiersLR_for_next_key, g_modifiersLR_persistent, 1, KEYDOWNANDUP, 0, aTargetWindow);
			else // Try to send it by alternate means.
				SendKeySpecial(*aKeys, modifiersLR_for_next_key, g_modifiersLR_persistent, 1, KEYDOWNANDUP, aTargetWindow);
			modifiersLR_for_next_key = 0;  // Safest to reset this regardless of whether a key was sent.
		}
	} // for()

	// Don't press back down the modifiers that were used to trigger this hotkey if there's
	// any doubt that they're still down, since doing so when they're not physically down
	// would cause them to be stuck down, which might cause unwanted behavior when the unsuspecting
	// user resumes typing:
	if (g_KeybdHook
		|| g_HotkeyModifierTimeout < 0 // User specified that the below should always be done.
		|| (GetTickCount() - g_script.mThisHotkeyStartTime) < (DWORD)g_HotkeyModifierTimeout) // Elapsed time < timeout-value
	{
		// If possible, update the set of modifier keys that are being physically held down.
		// This is done because the user may have released some keys during the send operation
		// (especially if KeyDelay > 0 and the Send is a large one).
		// Update: Include all keys that are phsyically down except those that were down
		// physically but not logically at the *start* of the send operation (since the send
		// operation may have changed the logical state).  In other words, we want to restore
		// the keys to their former logical-down position to match the fact that the user is still
		// holding them down physically.  The previously-down keys we don't do this for are those 
		// that were physically but not logically down, such as a naked Control key that's used
		// as a suffix without being a prefix.  See above comments for more details:
		if (g_KeybdHook)
			modifiersLR_down_physically_and_logically = g_modifiersLR_physical
				& ~modifiersLR_down_physically_but_not_logically; // intersect
		// Restore the state of the modifiers to be those believed to be physically held down
		// by the user.  Do not restore any that were logically "persistent", as detected upon
		// entrance to this function (e.g. due to something such as a prior "Send, {LWinDown}"),
		// since their state should already been correct if things above are designed right:
		modifiersLR_current = GetModifierLRState();
		modLR_type keys_to_press_down = modifiersLR_down_physically_and_logically & ~modifiersLR_current;
		// Use KEY_IGNORE_ALL_EXCEPT_MODIFIER to tell the hook to adjust g_modifiersLR_logical_non_ignored
		// because these keys being put back down match the physical pressing of those same keys by the
		// user, and we want such modifiers to be taken into account for the purpose of deciding whether
		// other hotkeys should fire (or the same one again if auto-repeating):
		if (keys_to_press_down)
		{
			SetModifierLRStateSpecific(keys_to_press_down, modifiersLR_current, KEYDOWN, aTargetWindow);
			// Since modifiers were changed by the above, do a key-delay if the special intra-keystroke
			// delay is in effect.
			// Since there normally isn't a delay between a change in modifiers and the first keystroke,
			// if a PressDuration is in effect, also do it here to improve reliability (I have observed
			// cases where modifiers need to be left alone for a short time in order for the keystrokes
			// that follow to be be modified by the intended set of modifiers).
			DoKeyDelay(g.PressDuration);
		}

		if (g_KeybdHook)
		{
			// Ensure that g_modifiersLR_logical_non_ignored does not contain any down-modifiers
			// that aren't down in g_modifiersLR_logical.  This is done mostly for peace-of-mind,
			// since there might be ways, via combinations of physical user input and the Send
			// commands own input (overlap and interference) for one to get out of sync with the
			// other.  The below uses ^ to find the differences between the two, then uses & to
			// find which are down in non_ignored that aren't in logical, then inverts those bits
			// in g_modifiersLR_logical_non_ignored, which sets those keys to be in the up position:
			g_modifiersLR_logical_non_ignored &= ~((g_modifiersLR_logical ^ g_modifiersLR_logical_non_ignored)
				& g_modifiersLR_logical_non_ignored);
		}
	}

	if (prior_capslock_state == TOGGLED_ON) // The current user setting requires us to turn it back on.
		ToggleKeyState(VK_CAPITAL, TOGGLED_ON);

	// Might be better to do this after changing capslock state, since having the threads attached
	// tends to help with updating the global state of keys (perhaps only under Win9x in this case):
	if (threads_are_attached)
		AttachThreadInput(my_thread, target_thread, FALSE);

	if (do_selective_blockinput && !blockinput_prev) // Turn it back off only if it wasn't ON before we started.
		Line::ScriptBlockInput(false);
}



int SendKey(vk_type aVK, sc_type aSC, modLR_type aModifiersLR, modLR_type aModifiersLRPersistent
	, int aRepeatCount, KeyEventTypes aEventType, modLR_type aKeyAsModifiersLR, HWND aTargetWindow)
// vk or sc may be zero, but not both.
// Returns the number of keys actually sent for caller convenience.
// The function is reponsible for first setting the correct modifier state,
// as specified by the caller, before sending the key.  After sending,
// it should put the system's modifier state back to the way it was
// originally, except, for safely, it seems best not to put back down
// any modifiers that were originally down unless those keys are physically
// down.
{
	if (!aVK && !aSC) return 0;

	// Avoid changing modifier states and other things if there is nothing to be sent.
	// Otherwise, menu bar might activated due to ALT keystrokes that don't modify any key,
	// the Start Menu might appear due to WIN keystrokes that don't modify anything, etc:
	if (aRepeatCount <= 0) return aRepeatCount;

	// Maybe best to call immediately so that the amount of time during which we haven't been pumping
	// messsages is more accurate:
	LONG_OPERATION_INIT

	// I thought maybe it might be best not to release unwanted modifier keys that are already down
	// (perhaps via something like "Send, {altdown}{esc}{altup}"), but that harms the case where
	// modifier keys are down somehow, unintentionally: The send command wouldn't behave as expected.
	// e.g. "Send, abc" while the control key is held down by other means, would send ^a^b^c,
	// possibly dangerous.  So it seems best to default to making sure all modifiers are in the
	// proper down/up position prior to sending any Keybd events.  UPDATE: This has been changed
	// so that only modifiers that were actually used to trigger that hotkey are released during
	// the send.  Other modifiers that are down may be down intentially, e.g. due to a previous
	// call to send, something like Send, {ShiftDown}.
	// UPDATE: It seems best to save the initial state only once, prior to sending the key-group,
	// because only at the beginning can the original state be determined without having to
	// save and restore it in each loop iteration.
	// UPDATE: Not saving and restoring at all anymore, due to interference (side-effects)
	// caused by the extra keybd events.

	modLR_type modifiersLR_specified = aModifiersLR | aModifiersLRPersistent;

	// Sending mouse clicks via ControlSend is not supported, so in that case fall back to the
	// old method of sending the VK directly (which probably has no effect 99% of the time):
	if (IsMouseVK(aVK) && !aTargetWindow)
	{
		// Pass "true" so that WIN and ALT are disguised if they have to be released due to
		// a hotkey such as !a::Send {LButton}
		if (SetModifierLRState(modifiersLR_specified, GetModifierLRState(), aTargetWindow, true, KEY_IGNORE))
			// Modifiers were changed by the above.
			DoKeyDelay(g.PressDuration); // See comments in SendKeys() about why this is done.
		// It will do its own MouseDelay:
		Line::MouseClick(aVK, COORD_UNSPECIFIED, COORD_UNSPECIFIED, aRepeatCount, g.DefaultMouseSpeed, aEventType);
	}
	else
	{
		for (int i = 0; i < aRepeatCount; ++i)
		{
			LONG_OPERATION_UPDATE_FOR_SENDKEYS
			// These modifiers above stay in effect for each of these keypresses.
			// Always on the first iteration, and thereafter only if the send won't be essentially
			// instantaneous.  The modifiers are checked before every key is sent because
			// if a high repeat-count was specified, the user may have time to release one or more
			// of the modifier keys that were used to trigger a hotkey.  That physical release
			// will cause a key-up event which will cause the state of the modifiers, as seen
			// by the system, to change.  Example: If user releases control-key during the operation,
			// some of the D's won't be control-D's:
			// ^c::Send,^{d 15}
			// Also: Seems best to do SetModifierState() even if Keydelay < 0:
			// Update: If this key is itself a modifier, don't change the state of the other
			// modifier keys just for it, since most of the time that is unnecessary and in
			// some cases, the extra generated keystrokes would cause complications/side-effects.
			if (!aKeyAsModifiersLR)
			{
				// Pass "true" so that WIN and ALT are disguised if they have to be released due to
				// a hotkey such as !a::Send test
				// See keyboard.h for explantion of KEY_IGNORE:
				if (SetModifierLRState(modifiersLR_specified, GetModifierLRState(), aTargetWindow, true, KEY_IGNORE))
					// Modifiers were changed by the above.
					DoKeyDelay(g.PressDuration); // See comments in SendKeys() about why this is done.
			}
			KeyEvent(aEventType, aVK, aSC, aTargetWindow, true);
		}
	}

	// The final iteration by the above loop does its key delay prior to us changing the
	// modifiers below.  This is a good thing because otherwise the modifiers would sometimes
	// be released so soon after the keys they modify that the modifiers are not in effect.
	// This can be seen sometimes when/ ctrl-shift-tabbing back through a multi-tabbed dialog:
	// The last ^+{tab} might otherwise not take effect because the CTRL key would be released
	// too quickly.

	// Release any modifiers that were pressed down just for the sake of the above
	// event (i.e. leave any persistent modifiers pressed down).  The caller should
	// already have verified that aModifiersLR does not contain any of the modifiers
	// in aModifiersLRPersistent.  Also, call GetModifierLRState() again explicitly
	// rather than trying to use a saved value from above, in case the above itself
	// changed the value of the modifiers (i.e. aVk/aSC is a modifier).  Admittedly,
	// that would be pretty strange but it seems the most correct thing to do.
	if (!aKeyAsModifiersLR) // See prior use of this var for explanation.
	{
		// It seems best not to use KEY_IGNORE_ALL_EXCEPT_MODIFIER in this case, though there's
		// a slight chance that a script or two might be broken by not doing so.  The chance
		// is very slight because the only thing KEY_IGNORE_ALL_EXCEPT_MODIFIER would allow is
		// something like the following example.  Note that the hotkey below must be a hook
		// hotkey (even more rare) because registered hotkeys will still see the logical modifier
		// state and thus fire regardless of whether g_modifiersLR_logical_non_ignored says that
		// they shouldn't:
		// #b::Send, {CtrlDown}{AltDown}
		// $^!a::MsgBox You pressed the A key after pressing the B key.
		// In the above, making ^!a a hook hotkey prevents it from working in conjunction with #b.
		// UPDATE: It seems slightly better to have it be KEY_IGNORE_ALL_EXCEPT_MODIFIER for these reasons:
		// 1) Persistent modifiers are fairly rare.  When they're in effect, it's usually for a reason
		//    and probably a pretty good one and from a user who knows what they're doing.
		// 2) The condition that g_modifiersLR_logical_non_ignored was added to fix occurs only when
		//    the user physically presses a suffix key (or auto-repeats one by holding it down)
		//    during the course of a SendKeys() operation.  Since the persistent modifiers were
		//    (by definition) already in effect prior to the Send, putting them back down for the
		//    purpose of firing hook hotkeys does not seem unreasonable, and may in fact add value.
		// Also, the below is called with "false" to avoid generating unnecessary disguise-keystrokes.
		// They are not needed because if our keystrokes were modified by either WIN or ALT, the
		// release of the WIN or ALT key will already be disguised due to its having modified
		// something while it was down.
		if (SetModifierLRState(aModifiersLRPersistent, GetModifierLRState(), aTargetWindow, false, KEY_IGNORE_ALL_EXCEPT_MODIFIER))
			// Modifiers were changed by the above.
			DoKeyDelay(g.PressDuration); // See comments in SendKeys() about why this is done.
	}

	return aRepeatCount;
}



///////////////////////////////////////////////////////////////////////////////
// SendKeySpecial's non-ANSI methods are based on AutoIt v3 source code, which is:
// Copyright 1999-2003 Jonathan Bennett and others listed at
// http://www.autoitscript.com/autoit3/docs/credits.htm
// License: GNU GPL version 2 or (at your option) any later version.
///////////////////////////////////////////////////////////////////////////////
int SendKeySpecial(char aChar, modLR_type aModifiersLR, modLR_type aModifiersLRPersistent
	, int aRepeatCount, KeyEventTypes aEventType, HWND aTargetWindow)
// Returns the number of keys actually sent for caller convenience.
// This function uses some of the same code as SendKey() above, so maintain them together.
{
	// Avoid changing modifier states and other things if there is nothing to be sent.
	// Otherwise, menu bar might activated due to ALT keystrokes that don't modify any key,
	// the Start Menu might appear due to WIN keystrokes that don't modify anything, etc:
	if (aRepeatCount <= 0) return aRepeatCount;

	static char cAnsiToAscii [128] =
	{ 
// 80   €            ‚      ƒ      „      …      †      ‡      ˆ      ‰      Š      ‹      Œ            Ž      
        0,     0,     0,  '\x9f',   0,     0,     0,     0,     0,     0,     0,     0,     0,     0,     0,     0,
// 90         ‘      ’      “      ”     •:f9    –      —      ˜      ™      š      ›      œ            ž      Ÿ
        0,     0,     0,     0,     0  ,   0  ,   0,     0,     0,     0,     0,     0,     0,     0,     0,     0,
// A0          ¡      ¢      £      ¤      ¥      ¦      §      ¨      ©      ª      «      ¬      ­      ®      ¯
        0,  '\xad','\x9b','\x9c',   0  ,'\x9d','\xb3','\x15',   0  ,   0  ,'\xa6','\xae','\xaa',   0  ,   0 ,    0,
// B0   °      ±      ²      ³      ´      µ      ¶      ·      ¸      ¹      º      »      ¼      ½      ¾      ¿
     '\xf8','\xf1','\xfd',   0  ,   0  ,'\xe6' ,'\x14','\xfa',   0     ,0  ,'\xa7','\xaf','\xac','\xab',   0  ,'\xa8',
// C0   À      Á      Â      Ã      Ä      Å      Æ      Ç      È      É      Ê      Ë      Ì      Í      Î      Ï
     '\x62','\x22','\x32','\x42','\x8e','\x8f','\x92','\x80','\x64','\x90','\x34','\x54','\x66','\x26','\x36','\x56',
// D0   Ð      Ñ      Ò      Ó      Ô      Õ      Ö      ×      Ø      Ù      Ú      Û      Ü      Ý      Þ      ß
        0,  '\xa5','\x68','\x28','\x38','\x48','\x99',   0  ,   0  ,'\x6a','\x2a','\x3a','\x9a','\x2c',   0  ,'\xe1',
// E0   à      á      â      ã      ä      å      æ      ç      è      é      ê      ë      ì      í      î      ï
     '\x85','\xa0','\x83','\x41','\x84','\x86','\x91','\x87','\x8a','\x82','\x88','\x89','\x8d','\xa1','\x8c','\x8b',
// F0   ð      ñ      ò      ó      ô      õ      ö      ÷      ø      ù      ú      û      ü      ý      þ      ÿ
        0,  '\xa4','\x95','\xa2','\x93','\x47','\x94','\xf6',   0  ,'\x97','\xa3','\x96','\x81','\x2b',   0  ,'\x98'
	};

//                               0   1   2   3   4   5   6   7   8   9   A   B   C   D   E   F
	static char g_cDiadicLetter[16] = {' ','a','A','e','E','i','I','o','O','u','U','y','Y','n','N',' '};

	modLR_type modifiersLR_specified = aModifiersLR | aModifiersLRPersistent;

	char asc_string[16] = "";

	// At the very least, this section should be kept to provide support for Daish ø & Ø chars.
	// However, it also extends support for many other chars that the AutoIt3's method cannot produce.
	// Most of these might be symbols, but some are probably useful to some people.
	// Here is the complete list of all ANSI chars above 127 (also known as "extended ASCII"?):
	// €‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ ¡¢£¤¥¦§¨©ª«¬­®¯°±²³´µ¶·¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ
	// Without the below method, only the following of the above can be produced:
	// ƒ¡¢£¥¦§ª«¬°±²µ¶·º»¼½¿AAAAÄÅÆÇEÉEEIIIIÑOOOOÖUUUÜYßàáâaäåæçèéêëìíîïñòóôoö÷ùúûüyÿ
	// The above has been tested on both Win98SE (but not the Win98 command prompt) and WinXP and seems
	// to work okay.  And since it uses ANSI keypad method, it should work on nearly all languages/layouts.
	if (aChar < 0) // Try using ANSI, which should be standard in Windows on nearly all language/layouts?
		snprintf(asc_string, sizeof(asc_string), "0%d", (int)(UCHAR)aChar);  // Must have leading zero.

	int asc_int = 0;
	if (!*asc_string)
	{
		asc_int = cAnsiToAscii[(int)((aChar - 128) & 0xff)] & 0xff;
		if (asc_int && (asc_int < 32 || asc_int >= 128))  // CHANGED FROM AU3: No sense in sending {Asc 0}.
			// simulation using {ASC nnn}
			// Only the char code between whose corresponding value
			// in cAnsiToAscii[] >= 128 can be sent directly
			_itoa(asc_int, asc_string, 10);
	}
	
	if (*asc_string) // It may have been made non-blank by the action inside the block above.
	{
		LONG_OPERATION_INIT
		for (int i = 0; i < aRepeatCount; ++i)
		{
			LONG_OPERATION_UPDATE_FOR_SENDKEYS
			SendASC(asc_string, aTargetWindow); // aTargetWindow is always NULL, it's just for maintainability.
			DoKeyDelay();
		}
		// See notes in SendKey():
		if (SetModifierLRState(aModifiersLRPersistent, GetModifierLRState(), aTargetWindow, false))
			DoKeyDelay(g.PressDuration); // See comments in SendKeys() about why this is done.
		return aRepeatCount;
	}

	// Otherwise:
	// simulation using diadic keystroke
	// pick up the diadic char according to cTradAnsiLetter
	// 0-3 part is the index in szDiadic defining the first character to be sent
	//        this implementation will allow to change the szDiadic value
	//			and if set to blank not to send the diadic char
	//			This will allow to treat without extra char when language keyboard
	//			does not support this diadic char Ex; ~ in german

	// List of Diadic characters will be updated according to keyboard layout
	//                     0   1   2   3   4   5   6   7
	static char g_cDiadic[8] = {' ',' ','´','^','~','¨','`',' '};
	static g_cDiadic_initialized = false;
	if (!g_cDiadic_initialized)
	{
		g_cDiadic_initialized = true;
		char szKLID[KL_NAMELENGTH];
		GetKeyboardLayoutName(szKLID);		// Get input locale identifier name
		//	update Diadics char according to keyboard possibility
		for (int i=1; i<=7; ++i)
		{	// check VK code to check if diadic char can be sent
			// English keyboard cannot sent diadics char
			if ( VkKeyScan(g_cDiadic[i]) == -1 || (strcmp(&szKLID[6], "09") == 0) )
				g_cDiadic[i] = ' ';				// reset Diadic setting
		}
		// need to check if a German keyboard  in use
		// because ~ does not work as a diadic char
		if (strcmp(&szKLID[6], "07") == 0)	// german keyboard
			g_cDiadic[4] = ' ';
	}

	char asc_string1[16] = "";
	bool send1 = false;
	char ch1 = g_cDiadic[asc_int >> 4];
	if (ch1 != ' ')		// something can be try to send diadic followed by non accent char
	{
		if (VkKeyScan(ch1) != -1)
			send1 = true;
		else
		{
			int asc_int1 = cAnsiToAscii[(int)((ch1 - 128) & 0xff)] & 0xff;
			if (asc_int1 < 32 || asc_int1 >= 128)
			{
				_itoa(asc_int1, asc_string1, 10);
				send1 = true;
			}
		}
	}

	// pick up the basic letter according to cTradAnsiLetter
	// 4-7 part is the index in szDiadic defining the second character to be sent
	char asc_string2[16] = "";
	bool send2 = false;
	char ch2 = g_cDiadicLetter[asc_int & 0x0f];
	if (ch2 != ' ')		// something can be try to send diadic followed by non accent char
	{
		if (VkKeyScan(ch2) != -1)
			send2 = true;
		else
		{
			int asc_int2 = cAnsiToAscii[(int)((ch2 - 128) & 0xff)] & 0xff;
			if (asc_int2 < 32 || asc_int2 >= 128)
			{
				_itoa(asc_int2, asc_string2, 10);
				send2 = true;
			}
		}
	}

	if (!send1 && !send2) // Can't simulate aChar.
		return 0;

	LONG_OPERATION_INIT
	for (int i = 0; i < aRepeatCount; ++i)
	{
		LONG_OPERATION_UPDATE_FOR_SENDKEYS
		if (send1)
		{
			if (*asc_string1)
				SendASC(asc_string1, aTargetWindow); // aTargetWindow is always NULL, it's just for maintainability.
			else
				SendChar(ch1, modifiersLR_specified, KEYDOWNANDUP, aTargetWindow);
		}
		if (send2)
		{
			if (*asc_string2)
				SendASC(asc_string2, aTargetWindow); // aTargetWindow is always NULL, it's just for maintainability.
			else
				SendChar(ch2, modifiersLR_specified, KEYDOWNANDUP, aTargetWindow);
		}
		DoKeyDelay();
	}
	// See notes in SendKey():
	if (SetModifierLRState(aModifiersLRPersistent, GetModifierLRState(), aTargetWindow, false))
		DoKeyDelay(g.PressDuration); // See comments in SendKeys() about why this is done.
	return aRepeatCount;
}



int SendASC(char *aAscii, HWND aTargetWindow)
// aAscii is a string to support explicit leading zeros because sending 216, for example, is not
// the same as sending 0216.
// Returns the number of keys sent (doesn't need to be exact).
{
	// This is just here to catch bugs in callers who do it wrong.  See notes in SendKeys() for explanation:
	if (aTargetWindow) return 0;

	int value = ATOI(aAscii);

	// This is not correct because it is possible to generate unicode characters by typing
	// Alt+256 and beyond:
	// if (value < 0 || value > 255) return 0; // Sanity check.

	// Known issue: If the hotkey that triggers this Send command is CONTROL-ALT
	// (and maybe either CTRL or ALT separately, as well), the {ASC nnnn} method
	// might not work reliably due to strangeness with that OS feature, at least on
	// WinXP.  I already tried adding delays between the keystrokes and it didn't help.

	// Make sure modifier state is correct: ALT pressed down and other modifiers UP
	// because CTRL and SHIFT seem to interfere with this technique if they are down,
	// at least under WinXP (though the Windows key doesn't seem to be a problem):
	modLR_type modifiersLR_to_release = GetModifierLRState()
		& (MOD_LCONTROL | MOD_RCONTROL | MOD_LSHIFT | MOD_RSHIFT);
	if (modifiersLR_to_release)
	{
		// Note: It seems best never to put them back down, because the act of doing so
		// may do more harm than good (i.e. the keystrokes may caused unexpected
		// side-effects.  Specify KEY_IGNORE so that this action does not affect the
		// modifiers that the hook uses to determine which hotkey should be triggered
		// for a suffix key that has more than one set of triggering modifiers
		// (for when the user is holding down that suffix to auto-repeat it --
		// see keyboard.h for details):
		SetModifierLRStateSpecific(modifiersLR_to_release, GetModifierLRState(), KEYUP, aTargetWindow, KEY_IGNORE);
		DoKeyDelay(g.PressDuration); // See comments in SendKeys() about why this is done.
	}

	int keys_sent = 0;  // Track this value and return it to the caller.

	modLR_type which_alt_was_already_down;
	if (!(which_alt_was_already_down = (GetModifierLRState() & (MOD_LALT|MOD_RALT)))) // Neither ALT key is down.
	{
		KeyEvent(KEYDOWN, VK_MENU); // Send the neutral ALT key to support Win9x/NT4.
		++keys_sent;
	}

	// Caller relies upon us to stop upon reaching the first non-digit character:
	for (char *cp = aAscii; *cp >= '0' && *cp <= '9'; ++cp)
	{
		// A comment from AutoIt3: ASCII 0 is 48, NUMPAD0 is 96, add on 48 to the ASCII.
		// Also, don't do WinDelay after each keypress in this case because it would make
		// such keys take up to 3 or 4 times as long to send (AutoIt3 avoids doing the
		// delay also).  Note that strings longer than 4 digits are allowed because
		// some or all OSes support Unicode characters 0 through 65535.
		KeyEvent(KEYDOWNANDUP, *cp + 48);
		++keys_sent;
	}

	// Must release the key regardless of whether it was already down, so that the
	// sequence will take effect immediately rather than waiting for the user to
	// release the ALT key (if it's physically down).  It's the caller's responsibility
	// to put it back down if it needs to be.
	// Fix for v1.0.36.06: Release the right-Alt key if that was the one already down prior
	// to the operation. Logic in both GetModifierLRState() above and KeyEvent() below ensures
	// that VK_RMENU either never comes about for Win9x/NT4 or that even if it does, KeyEvent()
	// translates it to a more appropriate event for those OSes.  This change fixed a hotkey
	// such as the following if it was fired by holding down AltGr: <^>!i:: Send {ASC 00256}
	KeyEvent(KEYUP, which_alt_was_already_down == MOD_RALT ? VK_RMENU : VK_MENU);
	return ++keys_sent;
}



int SendChar(char aChar, modLR_type aModifiersLR, KeyEventTypes aEventType, HWND aTargetWindow)
// Returns the number of keys sent (doesn't need to be exact).
{
	SHORT mod_plus_vk = VkKeyScan(aChar);
	char keyscan_modifiers = HIBYTE(mod_plus_vk);
	if (keyscan_modifiers == -1) // No translation could be made.
		return 0;

	// Combine the modifiers needed to enact this key with those that the caller wanted to be in effect:
	mod_type modifiers = ConvertModifiersLR(aModifiersLR);
	if (keyscan_modifiers & 0x01)
		modifiers |= MOD_SHIFT;
	if (keyscan_modifiers & 0x02)
		modifiers |= MOD_CONTROL;
	if (keyscan_modifiers & 0x04)
		modifiers |= MOD_ALT;

	// It's the caller's responsibility to restore the modifiers if it needs to:
	if (SetModifierState(modifiers, GetModifierLRState(), aTargetWindow, true, KEY_IGNORE))
		// Modifiers were changed by the above.
		DoKeyDelay(g.PressDuration); // See comments in SendKeys() about why this is done.
	KeyEvent(aEventType, LOBYTE(mod_plus_vk), 0, aTargetWindow, true);
	return 1;
}



ResultType KeyEvent(KeyEventTypes aEventType, vk_type aVK, sc_type aSC, HWND aTargetWindow
	, bool aDoKeyDelay, DWORD aExtraInfo)
// sc or vk, but not both, can be zero to indicate unspecified.
// For keys like NumpadEnter -- that have have a unique scancode but a non-unique virtual key --
// caller can just specify the sc.  In addition, the scan code should be specified for keys
// like NumpadPgUp and PgUp.  In that example, the caller would send the same scan code for
// both except that PgUp would be extended.   sc_to_vk() would map both of them to the same
// virtual key, which is fine since it's the scan code that matters to apps that can
// differentiate between keys with the same vk.

// Later, switch to using SendInput() on OS's that support it.  But this is non-trivial due to
// the following observation when it was attempted:
// The problem with SendInput is that there are many assumptions in SendKeys() and all the functions
// it calls about things already being in effect as the Send sends its keystrokes.  But when the
// sendinput method is used, these events haven't yet occurred so the logic is not correct,
// and it would probably be very complex to fix it.
{
	if (!aVK && !aSC) return FAIL;

	// Even if the sc_to_vk() mapping results in a zero-value vk, don't return.
	// I think it may be valid to send keybd_events	that have a zero vk.
	// In any case, it's unlikely to hurt anything:
	if (!aVK)
		aVK = sc_to_vk(aSC);
	else
		if (!aSC)
			// In spite of what the MSDN docs say, the scan code parameter *is* used, as evidenced by
			// the fact that the hook receives the proper scan code as sent by the below, rather than
			// zero like it normally would.  Even though the hook would try to use MapVirtualKey() to
			// convert zero-value scan codes, it's much better to send it here also for full compatibility
			// with any apps that may rely on scan code (and such would be the case if the hook isn't
			// active because the user doesn't need it; also for some games maybe).  In addition, if the
			// current OS is Win9x, we must map it here manually (above) because otherwise the hook
			// wouldn't be able to differentiate left/right on keys such as RCONTROL, which is detected
			// via its scan code.
			aSC = vk_to_sc(aVK);

	// Do this only after the above, so that the SC is left/right specific if the VK was such,
	// even on Win9x (though it's probably never called that way for Win9x; it's probably aways
	// called with either just the proper left/right SC or that plus the neutral VK).
	// Under WinNT/2k/XP, sending VK_LCONTROL and such result in the high-level (but not low-level
	// I think) hook receiving VK_CONTROL.  So somewhere interally it's being translated (probably
	// by keybd_event itself).  In light of this, translate the keys here manually to ensure full
	// support under Win9x (which might not support this internal translation).  The scan code
	// looked up above should still be correct for left-right centric keys even under Win9x.
	if (g_os.IsWin9x())
	{
		// Convert any non-neutral VK's to neutral for these OSes, since apps and the OS itself
		// can't be expected to support left/right specific VKs while running under Win9x:
		switch(aVK)
		{
		case VK_LCONTROL:
		case VK_RCONTROL: aVK = VK_CONTROL; break;
		case VK_LSHIFT:
		case VK_RSHIFT: aVK = VK_SHIFT; break;
		case VK_LMENU:
		case VK_RMENU: aVK = VK_MENU; break;
		}
	}

	// aTargetWindow is almost always passed in as NULL by our caller, even if the overall command
	// being executed is ControlSend.  This is because of the following reasons:
	// 1) Modifiers need to be logically changed via keybd_event() when using ControlSend against
	//    a cmd-prompt, console, or possibly other types of windows.
	// 2) If a hotkey triggered the ControlSend that got us here and that hotkey is a naked modifier
	//    such as RAlt:: or modified modifier such as ^#LShift, that modifier would otherwise auto-repeat
	//    an possibly interfere with the send operation.  This auto-repeat occurs because unlike a normal
	//    send, there are no calls to keybd_event() (keybd_event() stop the auto-repeat as a side-effect).
	// One exception to this is something like "ControlSend, Edit1, {Control down}", which explicitly
	// calls us with a target window.  This exception is by design and has been bug-fixed and documented
	// in ControlSend for v1.0.21:
	if (aTargetWindow && KeyToModifiersLR(aVK, aSC))
	{
		// When sending modifier keystrokes directly to a window, use the AutoIt3 SetKeyboardState()
		// technique to improve the reliability of changes to modifier states.  If this is not done,
		// sometimes the state of the SHIFT key (and perhaps other modifiers) will get out-of-sync
		// with what's intended, resulting in uppercase vs. lowercase problems (and that's probably
		// just the tip of the iceberg).  For this to be helpful, our caller must have ensured that
		// our thread is attached to aTargetWindow's (but it seems harmless to do the below even if
		// that wasn't done for any reason).  Doing this here in this function rather than at a
		// higher level probably isn't best in terms of performance (e.g. in the case where more
		// than one modifier is being changed, the multiple calls to Get/SetKeyboardState() could
		// be consolidated into one call), but it is much easier to code and maintain this way
		// since many different functions might call us to change the modifier state:
		BYTE state[256];
		GetKeyboardState((PBYTE)&state);
		if (aEventType == KEYDOWN)
			state[aVK] |= 0x80;
		else if (aEventType == KEYUP)
			state[aVK] &= ~0x80;
		// else KEYDOWNANDUP, in which case it seems best (for now) not to change the state at all.
		// It's rarely if ever called that way anyway.

		// If aVK is a left/right specific key, be sure to also update the state of the neutral key:
		switch(aVK)
		{
		case VK_LCONTROL: 
		case VK_RCONTROL:
			if ((state[VK_LCONTROL] & 0x80) || (state[VK_RCONTROL] & 0x80))
				state[VK_CONTROL] |= 0x80;
			else
				state[VK_CONTROL] &= ~0x80;
			break;
		case VK_LSHIFT:
		case VK_RSHIFT:
			if ((state[VK_LSHIFT] & 0x80) || (state[VK_RSHIFT] & 0x80))
				state[VK_SHIFT] |= 0x80;
			else
				state[VK_SHIFT] &= ~0x80;
			break;
		case VK_LMENU:
		case VK_RMENU:
			if ((state[VK_LMENU] & 0x80) || (state[VK_RMENU] & 0x80))
				state[VK_MENU] |= 0x80;
			else
				state[VK_MENU] &= ~0x80;
			break;
		}

		SetKeyboardState((PBYTE)&state);
		// Even after doing the above, we still continue on to send the keystrokes
		// themselves to the window, for greater reliability (same as AutoIt3).
	}

	if (aTargetWindow)
	{
		// lowest 16 bits: repeat count: always 1 for up events, probably 1 for down in our case.
		// highest order bits: 11000000 (0xC0) for keyup, usually 00000000 (0x00) for keydown.
		LPARAM lParam = (LPARAM)(aSC << 16);
		if (aEventType != KEYUP)  // i.e. always do it for KEYDOWNANDUP
			PostMessage(aTargetWindow, WM_KEYDOWN, aVK, lParam | 0x00000001);
		// The press-duration delay is done only when this is a down-and-up because otherwise,
		// the normal g.KeyDelay will be in effect.  In other words, it seems undesirable in
		// most cases to do both delays for only "one half" of a keystroke:
		if (aDoKeyDelay && aEventType == KEYDOWNANDUP)
			DoKeyDelay(g.PressDuration);
		if (aEventType != KEYDOWN)  // i.e. always do it for KEYDOWNANDUP
			PostMessage(aTargetWindow, WM_KEYUP, aVK, lParam | 0xC0000001);
	}
	else // Keystrokes are to be sent with keybd_event() rather than PostMessage().
	{
		// Users of the below want them updated only for keybd_event() keystrokes (not PostMessage ones):
		sPrevEventType = aEventType;
		sPrevVK = aVK;
		// Turn off BlockInput momentarily to support sending of the ALT key.  This is not done for
		// Win9x because input cannot be simulated during BlockInput on that platform anyway; thus
		// it seems best (due to backward compatibility) not to turn off BlockInput then.
		// Jon Bennett noted: "As many of you are aware BlockInput was "broken" by a SP1 hotfix under
		// Windows XP so that the ALT key could not be sent. I just tried it under XP SP2 and it seems
		// to work again."  In light of this, it seems best to unconditionally and momentarily disable
		// input blocking regardless of which OS is being used (except Win9x, since no simulated input
		// is even possible for those OSes).
		bool we_turned_blockinput_off = g_BlockInput && (aVK == VK_MENU || aVK == VK_LMENU || aVK == VK_RMENU)
			&& g_os.IsWinNT4orLater();
		if (we_turned_blockinput_off)
			Line::ScriptBlockInput(false);

		vk_type control_vk;
		bool lcontrol_was_down, do_detect_altgr;
		if (do_detect_altgr = (aVK == VK_RMENU && !g_LayoutHasAltGr)) // Keyboard layout isn't yet marked as having an AltGr key, so auto-detect it here as well as other places.
		{
			control_vk = g_os.IsWin2000orLater() ? VK_LCONTROL : VK_CONTROL;
			lcontrol_was_down = GetAsyncKeyState(control_vk) & 0x80000000; // In this case, must use GetAsyncKeyState() vs. GetKeyState() to detect the change, at least on XP.
			// Add extra detection of AltGr if hook is installed, which has been show to be useful for some
			// scripts where the other AltGr detection methods don't occur in a timely enough fashion.
			// The following method relies upon the fact that it's impossible for the hook to receive
			// events from the user while it's processing our keybd_event() here.  This is because
			// any physical keystrokes that happen to occur at the exact moment of our keybd_event()
			// will stay queued until the main event loop routes them to the hook via GetMessage().
			g_HookReceiptOfLControlMeansAltGr = true;
		}

		if (aEventType != KEYUP)  // i.e. always do it for KEYDOWNANDUP
		{
			// v1.0.35.07 was when g_IgnoreNextLControlDown/Up were originally added.
			// The following global is used to flag as our own the keyboard driver's LControl-down keystroke
			// that is triggered by RAlt-down (AltGr).  This prevents it from triggering hotkeys such as
			// "*Control::".  It probably fixes other obscure side-effects and bugs also, since the
			// event should be considered script-generated even though indirect.  Note: The problem with
			// having the hook detect AltGr's automatic LControl-down is that the keyboard driver seems
			// to generate the LControl-down *before* notifying the system of the RAlt-down.  That makes
			// it impossible for the hook to flag the LControl keystroke in advance, so it would have to
			// retroactively undo the effects.  But that is impossible because the previous keystroke might
			// already have wrongly fired a hotkey.
			if (aVK == VK_RMENU && g_LayoutHasAltGr) // VK_RMENU vs. VK_MENU should be safe since this logic is only needed for the hook, which is never in effect on Win9x.
				g_IgnoreNextLControlDown = true; // Must be set prior to keybd_event() to be effective.
			keybd_event(aVK
				, LOBYTE(aSC)  // naked scan code (the 0xE0 prefix, if any, is omitted)
				, (HIBYTE(aSC) ? KEYEVENTF_EXTENDEDKEY : 0)
				, aExtraInfo);
			// The following is done by us rather than by the hook to avoid problems where:
			// 1) The hook is removed at a critical point during the operation, preventing the variable from
			//    being reset to false.
			// 2) For some reason this AltGr keystroke done above did not cause LControl to go down (perhaps
			//    because the current keyboard layout doesn't have AltGr as we thought), which would be a bug
			//    because some other Ctrl keystroke would then be wrongly ignored.
			g_IgnoreNextLControlDown = false; // Unconditional reset.

			if (do_detect_altgr) // i.e. g_LayoutHasAltGr is currently false, so make it true if called for.
			{
				do_detect_altgr = false; // Indicate that the second half has been done (for later below).
				g_HookReceiptOfLControlMeansAltGr = false; // Must reset promptly in case key-delay below routes physical keystrokes to hook.
				// Do it the following way rather than setting g_LayoutHasAltGr directly to the result of
				// the boolean expression because keybd_event() above may have changed the value of
				// g_LayoutHasAltGr to true, in which case that should be given precedence over the below.
				if (!lcontrol_was_down && (GetAsyncKeyState(control_vk) & 0x80000000)) // It wasn't down before but now it is.  Thus, RAlt is really AltGr.
					g_LayoutHasAltGr = true;
			}

			if (aVK == VK_NUMLOCK && g_os.IsWin9x()) // Under Win9x, Numlock needs special treatment.
				ToggleNumlockWin9x();

			if (!g_KeybdHook) // Hook isn't logging, so we'll log just the keys we send, here.
				UpdateKeyEventHistory(false, aVK, aSC);
		}
		// The press-duration delay is done only when this is a down-and-up because otherwise,
		// the normal g.KeyDelay will be in effect.  In other words, it seems undesirable in
		// most cases to do both delays for only "one half" of a keystroke:
		if (aDoKeyDelay && aEventType == KEYDOWNANDUP)
			DoKeyDelay(g.PressDuration);
		if (aEventType != KEYDOWN)  // i.e. always do it for KEYDOWNANDUP
		{
			// See comments above for details about g_LayoutHasAltGr and g_IgnoreNextLControlUp:
			if (aVK == VK_RMENU && g_LayoutHasAltGr) // VK_RMENU vs. VK_MENU should be safe since this logic is only needed for the hook, which is never in effect on Win9x.
				g_IgnoreNextLControlUp = true; // Must be set prior to keybd_event() to be effective.
			keybd_event(aVK, LOBYTE(aSC), (HIBYTE(aSC) ? KEYEVENTF_EXTENDEDKEY : 0)
				| KEYEVENTF_KEYUP, aExtraInfo);
			g_IgnoreNextLControlUp = false; // Unconditional reset.
			if (do_detect_altgr) // This should be true only when aEventType==KEYUP. See similar section above for comments.
			{
				g_HookReceiptOfLControlMeansAltGr = false;
				// Do it the following way rather than setting g_LayoutHasAltGr directly to the result of
				// the boolean expression because keybd_event() above may have changed the value of
				// g_LayoutHasAltGr to true, in which case that should be given precedence over the below.
				if (lcontrol_was_down && !(GetAsyncKeyState(control_vk) & 0x80000000)) // RAlt is really AltGr.
					g_LayoutHasAltGr = true;
			}
			if (!g_KeybdHook) // Hook isn't logging, so we'll log just the keys we send, here.
				UpdateKeyEventHistory(true, aVK, aSC);
		}

		if (we_turned_blockinput_off)  // Turn it back on.
			Line::ScriptBlockInput(true);
	}

	if (aDoKeyDelay)
		DoKeyDelay();
	return OK;
}



void UpdateKeyEventHistory(bool aKeyUp, vk_type aVK, sc_type aSC)
{
	if (!g_KeyHistory) // Don't access the array if it doesn't exist (i.e. key history is disabled).
		return;
	g_KeyHistory[g_KeyHistoryNext].key_up = aKeyUp;
	g_KeyHistory[g_KeyHistoryNext].vk = aVK;
	g_KeyHistory[g_KeyHistoryNext].sc = aSC;
	g_KeyHistory[g_KeyHistoryNext].event_type = 'i'; // Callers all want this.
	g_HistoryTickNow = GetTickCount();
	g_KeyHistory[g_KeyHistoryNext].elapsed_time = (g_HistoryTickNow - g_HistoryTickPrev) / (float)1000;
	g_HistoryTickPrev = g_HistoryTickNow;
	HWND fore_win = GetForegroundWindow();
	if (fore_win)
	{
		if (fore_win != g_HistoryHwndPrev)
			GetWindowText(fore_win, g_KeyHistory[g_KeyHistoryNext].target_window, sizeof(g_KeyHistory[g_KeyHistoryNext].target_window));
		else // i.e. avoid the call to GetWindowText() if possible.
			*g_KeyHistory[g_KeyHistoryNext].target_window = '\0';
	}
	else
		strcpy(g_KeyHistory[g_KeyHistoryNext].target_window, "N/A");
	g_HistoryHwndPrev = fore_win; // Update unconditionally in case it's NULL.
	if (++g_KeyHistoryNext >= g_MaxHistoryKeys)
		g_KeyHistoryNext = 0;
}



ToggleValueType ToggleKeyState(vk_type aVK, ToggleValueType aToggleValue)
// Toggle the given aVK into another state.  For performance, it is the caller's responsibility to
// ensure that aVK is a toggleable key such as capslock, numlock, or scrolllock.
// Returns the state the key was in before it was changed (but it's only a best-guess under Win9x).
{
	// Can't use GetAsyncKeyState() because it doesn't have this info:
	ToggleValueType starting_state = IsKeyToggledOn(aVK) ? TOGGLED_ON : TOGGLED_OFF;
	if (aToggleValue != TOGGLED_ON && aToggleValue != TOGGLED_OFF) // Shouldn't be called this way.
		return starting_state;
	if (starting_state == aToggleValue) // It's already in the desired state, so just return the state.
		return starting_state;
	if (aVK == VK_NUMLOCK)
	{
		if (g_os.IsWin9x())
		{
			// For Win9x, we want to set the state unconditionally to be sure it's right.  This is because
			// the retrieval of the Capslock state, for example, is unreliable, at least under Win98se
			// (probably due to lack of an AttachThreadInput() having been done).  Although the
			// SetKeyboardState() method used by ToggleNumlockWin9x is not required for caps & scroll lock keys,
			// it is required for Numlock:
			ToggleNumlockWin9x();
			return starting_state;  // Best guess, but might be wrong.
		}
		// Otherwise, NT/2k/XP:
		// Sending an extra up-event first seems to prevent the problem where the Numlock
		// key's indicator light doesn't change to reflect its true state (and maybe its
		// true state doesn't change either).  This problem tends to happen when the key
		// is pressed while the hook is forcing it to be either ON or OFF (or it suppresses
		// it because it's a hotkey).  Needs more testing on diff. keyboards & OSes:
		KeyEvent(KEYUP, aVK);
	}
	// Since it's not already in the desired state, toggle it:
	KeyEvent(KEYDOWNANDUP, aVK);
	// Fix for v1.0.36.06: // If it's Capslock and it didn't turn off as specified, it's probably because
	// the OS is configured to turn Capslock off only in response to pressing the SHIFT key (via Ctrl Panel's
	// Regional settings).  So send shift to do it instead:
	if (aVK == VK_CAPITAL && aToggleValue == TOGGLED_OFF && IsKeyToggledOn(aVK))
 		KeyEvent(KEYDOWNANDUP, VK_SHIFT);

	return starting_state;
}



void ToggleNumlockWin9x()
// Numlock requires a special method to toggle the state and its indicator light under Win9x.
// Capslock and Scrolllock do not need this method, since keybd_event() works for them.
{
	BYTE state[256];
	GetKeyboardState((PBYTE)&state);
	state[VK_NUMLOCK] ^= 0x01;  // Toggle the low-order bit to the opposite state.
	SetKeyboardState((PBYTE)&state);
}



//void CapslockOffWin9x()
//{
//	BYTE state[256];
//	GetKeyboardState((PBYTE)&state);
//	state[VK_CAPITAL] &= ~0x01;
//	SetKeyboardState((PBYTE)&state);
//}



/*
void SetKeyState (vk_type vk, int aKeyUp)
// Later need to adapt this to support Win9x by using SetKeyboardState for those OSs.
{
	if (!vk) return;
	int key_already_up = !(GetKeyState(vk) & 0x8000);
	if ((key_already_up && aKeyUp) || (!key_already_up && !aKeyUp))
		return;
	KeyEvent(aKeyUp, vk);
}
*/



modLR_type SetModifierState(mod_type aModifiersNew, modLR_type aModifiersLRnow, HWND aTargetWindow
	, bool aDisguiseWinAlt, DWORD aExtraInfo)
// Returns the set of modifiers that *changed* (i.e. went from down to up or vice versa).
{
	// Can't do this because the two values aren't compatible (one is LR and the other neutral):
	//if (aModifiersNew == aModifiersLRnow) return aModifiersLRnow
/*
char error_text[512];
snprintf(error_text, sizeof(error_text), "new=%02X, LRnow=%02X", aModifiersNew, aModifiersLRnow);
MsgBox(error_text);
*/
	// It's done this way in case RSHIFT, for example, is down, thus giving us the shift key
	// already without having to put the (normal/default) LSHIFT key down.
	mod_type modifiers_now = ConvertModifiersLR(aModifiersLRnow);
	modLR_type modifiersLRnew = aModifiersLRnow; // Start with what they are now.

	// If neither should be on, turn them both off.  If one should be on, turn on only one.
	// But if both are on when only one should be (rare), leave them both on:
	if ((modifiers_now & MOD_CONTROL) && !(aModifiersNew & MOD_CONTROL))
		modifiersLRnew &= ~(MOD_LCONTROL | MOD_RCONTROL);
	else if (!(modifiers_now & MOD_CONTROL) && (aModifiersNew & MOD_CONTROL))
		modifiersLRnew |= MOD_LCONTROL;
	if ((modifiers_now & MOD_ALT) && !(aModifiersNew & MOD_ALT))
		modifiersLRnew &= ~(MOD_LALT | MOD_RALT);
	else if (!(modifiers_now & MOD_ALT) && (aModifiersNew & MOD_ALT))
		modifiersLRnew |= MOD_LALT;
	if ((modifiers_now & MOD_WIN) && !(aModifiersNew & MOD_WIN))
		modifiersLRnew &= ~(MOD_LWIN | MOD_RWIN);
	else if (!(modifiers_now & MOD_WIN) && (aModifiersNew & MOD_WIN))
		modifiersLRnew |= MOD_LWIN;
	if ((modifiers_now & MOD_SHIFT) && !(aModifiersNew & MOD_SHIFT))
		modifiersLRnew &= ~(MOD_LSHIFT | MOD_RSHIFT);
	else if (!(modifiers_now & MOD_SHIFT) && (aModifiersNew & MOD_SHIFT))
		modifiersLRnew |= MOD_LSHIFT;

	if (modifiersLRnew == aModifiersLRnow)  // They're already in the right state.
		return 0;
	// Otherwise, change the state:
	return SetModifierLRState(modifiersLRnew, aModifiersLRnow, aTargetWindow, aDisguiseWinAlt, aExtraInfo);
}



modLR_type SetModifierLRState(modLR_type modifiersLRnew, modLR_type aModifiersLRnow, HWND aTargetWindow
	, bool aDisguiseWinAlt, DWORD aExtraInfo)
// Puts modifiers into the specified state, releasing or pressing down keys as needed.
// Returns the set of modifiers that *changed* (i.e. went from down to up or vice versa).
// Note that by design and as documented for ControlSend, aTargetWindow is not used as the target for the
// various calls to KeyEvent() here.  It is only used as a workaround for the GUI window issue described
// at the bottom.
{
	if (aModifiersLRnow == modifiersLRnew) // They're already in the right state, so avoid doing all the checks.
		return 0;

	// Notes about modifier key behavior on Windows XP (these probably apply to NT/2k also, and has also
	// been tested to be true on Win98): The WIN and ALT keys are the problem keys, because if either is
	// released without having modified something (even another modifier), the WIN key will cause the
	// Start Menu to appear, and the ALT key will activate the menu bar of the active window (if it has one).
	// For example, a hook hotkey such as "$#c::Send text" (text must start with a lowercase letter
	// to reproduce the issue, because otherwise WIN would be auto-disguised as a side effect of the SHIFT
	// keystroke) would cause the Start Menu to appear if the disguise method below weren't used.
	// See comments in SetModifierLRStateSpecific() for more details about this.

	// Set up some conditions so that the keystrokes that disguise the release of Win or Alt
	// are only sent when necessary (which helps avoid complications caused by keystroke interaction,
	// while improving performance):
	bool ctrl_not_down = !(aModifiersLRnow & (MOD_LCONTROL | MOD_RCONTROL)); // Neither CTRL key is down now.
	bool ctrl_will_not_be_down = !(modifiersLRnew & (MOD_LCONTROL | MOD_RCONTROL)) // Nor will it be.
		&& !(g_LayoutHasAltGr && (modifiersLRnew & MOD_RALT)); // Nor will it be pushed down indirectly due to AltGr.

	bool ctrl_nor_shift_nor_alt_down = ctrl_not_down                             // Neither CTRL key is down now.
		&& !(aModifiersLRnow & (MOD_LSHIFT | MOD_RSHIFT | MOD_LALT | MOD_RALT)); // Nor is any SHIFT/ALT key.

	bool ctrl_or_shift_or_alt_will_be_down = !ctrl_will_not_be_down            // CTRL will be down.
		|| (modifiersLRnew & (MOD_LSHIFT | MOD_RSHIFT | MOD_LALT | MOD_RALT)); // or SHIFT or ALT will be.

	// If the required disguise keys aren't down now but will be, defer the release of Win and/or Alt
	// until after the disguise keys are in place (since in that case, the caller wanted them down
	// as part of the normal operation here):
	bool defer_win_release = ctrl_nor_shift_nor_alt_down && ctrl_or_shift_or_alt_will_be_down;
	bool defer_alt_release = ctrl_not_down && !ctrl_will_not_be_down;

	bool release_lwin = (aModifiersLRnow & MOD_LWIN) && !(modifiersLRnew & MOD_LWIN);
	bool release_rwin = (aModifiersLRnow & MOD_RWIN) && !(modifiersLRnew & MOD_RWIN);
	bool release_lalt = (aModifiersLRnow & MOD_LALT) && !(modifiersLRnew & MOD_LALT);
	bool release_ralt = (aModifiersLRnow & MOD_RALT) && !(modifiersLRnew & MOD_RALT);

	// ** WIN (must be done before ALT in case it is relying on ALT being down to disguise the release WIN).
	// If ALT is going to be pushed down further below, defer_win_release should be true, which will make sure
	// the WIN key isn't released until after the ALT key is pushed down here at the top.
	if (release_lwin)
	{
		if (!defer_win_release)
		{
			// Fixed for v1.0.25: To avoid triggering the system's LAlt+Shift language hotkey, the
			// Control key is now used to suppress LWIN/RWIN (preventing the Start Menu from appearing)
			// rather than the Shift key.  This is definitely needed for ALT, but is done here for
			// WIN also in case ALT is down, which might cause the use of SHIFT as the disguise key
			// to trigger the language switch.
			if (ctrl_nor_shift_nor_alt_down && aDisguiseWinAlt) // Nor will they be pushed down later below, otherwise defer_win_release would have been true and we couldn't get to this point.
				KeyEvent(KEYDOWNANDUP, VK_CONTROL, 0, NULL, false, aExtraInfo); // Disguise key release to suppress Start Menu.
				// The above event is safe because if we're here, it means VK_CONTROL will not be
				// pressed down further below.  In other words, we're not defeating the job
				// of this function by sending these disguise keystrokes.
			KeyEvent(KEYUP, VK_LWIN, 0, NULL, false, aExtraInfo);
		}
		// else release it only after the normal operation of the function pushes down the disguise keys.
	}
	else if (!(aModifiersLRnow & MOD_LWIN) && (modifiersLRnew & MOD_LWIN)) // Press down is needed.
		// Note that no disguising of WIN or ALT is done for ALT-down since those events should not
		// need disguising since a Send command's activity uses them to modify other keys, which
		// is an automatic disguise.  Even if it didn't modify other keys with them, activation of
		// the Start Menu or the menu bar might be the intended outcome of the script.
		KeyEvent(KEYDOWN, VK_LWIN, 0, NULL, false, aExtraInfo);

	if (release_rwin)
	{
		if (!defer_win_release)
		{
			if (ctrl_nor_shift_nor_alt_down && aDisguiseWinAlt)
				KeyEvent(KEYDOWNANDUP, VK_CONTROL, 0, NULL, false, aExtraInfo); // Disguise key release to suppress Start Menu.
			KeyEvent(KEYUP, VK_RWIN, 0, NULL, false, aExtraInfo);
		}
		// else release it only after the normal operation of the function pushes down the disguise keys.
	}
	else if (!(aModifiersLRnow & MOD_RWIN) && (modifiersLRnew & MOD_RWIN))
		KeyEvent(KEYDOWN, VK_RWIN, 0, NULL, false, aExtraInfo);

	// ** ALT
	if (release_lalt)
	{
		if (!defer_alt_release)
		{
			if (ctrl_not_down && aDisguiseWinAlt)
				KeyEvent(KEYDOWNANDUP, VK_CONTROL, 0, NULL, false, aExtraInfo); // Disguise key release to suppress menu activation.
			KeyEvent(KEYUP, VK_LMENU, 0, NULL, false, aExtraInfo);
		}
	}
	else if (!(aModifiersLRnow & MOD_LALT) && (modifiersLRnew & MOD_LALT))
		KeyEvent(KEYDOWN, VK_LMENU, 0, NULL, false, aExtraInfo);

	if (release_ralt)
	{
		if (!defer_alt_release || g_LayoutHasAltGr) // No need to defer if RAlt==AltGr. But don't change the value of defer_alt_release because LAlt uses it too.
		{
			if (g_LayoutHasAltGr)
			{
				// Indicate that control is both up and required so that the section after this one won't
				// push it down.  This change might not be strictly necessary to fix anything but it does
				// eliminate redundant keystrokes.  See similar section below for more details.
				aModifiersLRnow &= ~MOD_LCONTROL; // To reflect what KeyEvent(KEYUP, VK_RMENU) below will do.
				modifiersLRnew &= ~MOD_LCONTROL;  //
			}
			else // No AltGr, so check if disguise is necessary (AltGr itself never needs disguise).
				if (ctrl_not_down && aDisguiseWinAlt)
					KeyEvent(KEYDOWNANDUP, VK_CONTROL, 0, NULL, false, aExtraInfo); // Disguise key release to suppress menu activation.
			KeyEvent(KEYUP, VK_RMENU, 0, NULL, false, aExtraInfo);
		}
	}
	else if (!(aModifiersLRnow & MOD_RALT) && (modifiersLRnew & MOD_RALT))
	{
		KeyEvent(KEYDOWN, VK_RMENU, 0, NULL, false, aExtraInfo);
		if (g_LayoutHasAltGr) // Note that KeyEvent() might have just altered the value of g_LayoutHasAltGr.
		{
			// Indicate that control is both down and required so that the section after this one won't
			// release it.  Without this fix, a hotkey that sends an AltGr char such as "^ä:: SendRaw, {"
			// would fail to work under German layout because left-alt would be released after right-alt
			// goes down.
			aModifiersLRnow |= MOD_LCONTROL; // To reflect what KeyEvent() did above.
			modifiersLRnew |= MOD_LCONTROL;  // All callers want LControl to be down if they wanted AltGr to be down.
		}
	}

	// CONTROL and SHIFT are done only after the above because the above might rely on them
	// being down before for certain early operations.

	// ** CONTROL
	if ((aModifiersLRnow & MOD_LCONTROL) && !(modifiersLRnew & MOD_LCONTROL))
		KeyEvent(KEYUP, VK_LCONTROL, 0, NULL, false, aExtraInfo);
	else if (!(aModifiersLRnow & MOD_LCONTROL) && (modifiersLRnew & MOD_LCONTROL))
		KeyEvent(KEYDOWN, VK_LCONTROL, 0, NULL, false, aExtraInfo);
	if ((aModifiersLRnow & MOD_RCONTROL) && !(modifiersLRnew & MOD_RCONTROL))
		KeyEvent(KEYUP, VK_RCONTROL, 0, NULL, false, aExtraInfo);
	else if (!(aModifiersLRnow & MOD_RCONTROL) && (modifiersLRnew & MOD_RCONTROL))
		KeyEvent(KEYDOWN, VK_RCONTROL, 0, NULL, false, aExtraInfo);
	
	// ** SHIFT
	if ((aModifiersLRnow & MOD_LSHIFT) && !(modifiersLRnew & MOD_LSHIFT))
		KeyEvent(KEYUP, VK_LSHIFT, 0, NULL, false, aExtraInfo);
	else if (!(aModifiersLRnow & MOD_LSHIFT) && (modifiersLRnew & MOD_LSHIFT))
		KeyEvent(KEYDOWN, VK_LSHIFT, 0, NULL, false, aExtraInfo);
	if ((aModifiersLRnow & MOD_RSHIFT) && !(modifiersLRnew & MOD_RSHIFT))
		KeyEvent(KEYUP, VK_RSHIFT, 0, NULL, false, aExtraInfo);
	else if (!(aModifiersLRnow & MOD_RSHIFT) && (modifiersLRnew & MOD_RSHIFT))
		KeyEvent(KEYDOWN, VK_RSHIFT, 0, NULL, false, aExtraInfo);

	if (defer_win_release) // Must be done before ALT because it might rely on ALT being down to disguise release of WIN key.
	{
		if (release_lwin)
			KeyEvent(KEYUP, VK_LWIN, 0, NULL, false, aExtraInfo);
		if (release_rwin)
			KeyEvent(KEYUP, VK_RWIN, 0, NULL, false, aExtraInfo);
	}
	if (defer_alt_release)
	{
		if (release_lalt)
			KeyEvent(KEYUP, VK_LMENU, 0, NULL, false, aExtraInfo);
		if (release_ralt && !g_LayoutHasAltGr) // If g_LayoutHasAltGr==true, RAlt would already have been released earlier since defer_alt_release would have been ignored for it.
			KeyEvent(KEYUP, VK_RMENU, 0, NULL, false, aExtraInfo);
	}

	// When calling KeyEvent(), probably best not to specify a scan code unless
	// absolutely necessary, since some keyboards may have non-standard scan codes
	// which KeyEvent() will resolve into the proper vk tranlations for us.
	// Decided not to Sleep() between keystrokes, even zero, out of concern that this
	// would result in a significant delay (perhaps more than 10ms) while the system
	// is under load.

	// Since the above didn't return early, keybd_event() has been used to change the state
	// of at least one modifier.  As a result, if the caller gave a non-NULL aTargetWindow,
	// it wants us to check if that window belongs to our thread.  If it does, we should do
	// a short msg queue check to prevent an apparent synchronization problem when using
	// ControlSend against the script's own GUI or other windows.  Here is an example of a
	// line whose modifier would not be in effect in time for its keystroke to be modified
	// by it:
	// ControlSend, Edit1, ^{end}, Test Window
	// Update: Another bug-fix for v1.0.21, as was the above: If the keyboard hook is installed,
	// the modifier keystrokes must have a way to get routed through the hook BEFORE the
	// keystrokes get sent via PostMessage().  If not, the correct modifier state will usually
	// not be in effect (or at least not be in sync) for the keys sent via PostMessage() afterward:
	if (aTargetWindow) // ControlSend mode is in effect.
	{
		if (g_KeybdHook)  // This check must come first (it should take precedence if both conditions are true).
			// -1 has been verified to be insufficient, at least for the very first letter sent if it is
			// supposed to be capitalized:
			SLEEP_WITHOUT_INTERRUPTION(0)
		else if (GetWindowThreadProcessId(aTargetWindow, NULL) == GetCurrentThreadId())
			SLEEP_WITHOUT_INTERRUPTION(-1)
	}

	return aModifiersLRnow ^ modifiersLRnew; // Calculate the set of modifiers that changed (currently excludes AltGr's change of LControl's state).
}



void SetModifierLRStateSpecific(modLR_type aModifiersLR, modLR_type aModifiersLRnow, KeyEventTypes aEventType
	, HWND aTargetWindow, DWORD aExtraInfo)
// Press or release only the specific keys whose bits are set to 1 in aModifiersLR.
// aEventType should be either KEYDOWN or KEYUP (not KEYDOWNANDUP).
// If aEventType == KEYDOWN, it is assumed that the caller wants any ALT or WIN keystroke
// disguised (if needed) to prevent the Start Menu from appearing or the menu bar from being
// activated.  In other words, it is assumed that the caller will not be sending any keystrokes
// afterward that would avert the need for disguising.
{
	if (!aModifiersLR) // Nothing to do (especially avoids the aTargetWindow check at the bottom).
		return;

	bool ctrl_not_down = !(aModifiersLRnow & (MOD_LCONTROL | MOD_RCONTROL)); // Neither CTRL key is down now.
	bool ctrl_wont_be_down = aEventType == KEYUP || !((aModifiersLR & (MOD_LCONTROL | MOD_RCONTROL))
		|| (g_LayoutHasAltGr && (aModifiersLR & MOD_RALT))); // Nor will it be pushed down indirectly due to AltGr.
	// Control won't be going down if aEventType==KEYUP because nothing is ever pressed down in that mode.

	// To prevent it from activating the menu bar, the release of the ALT key should be disguised
	// unless a CTRL key is currently down.  This is because CTRL always seems to avoid the
	// activation of the menu bar (unlike SHIFT, which sometimes allows the menu to be activated,
	// though this is hard to reproduce on XP).  Another reason not to use SHIFT is that the OS
	// uses LAlt+Shift as a hotkey to switch languages.  Such a hotkey would be triggered if SHIFT
	// were pressed down to disguise the release of LALT.

	// Alt-down events are also disguised whenever they won't be accompanied by a Ctrl-down.
	// This is necessary whenever our caller does not plan to disguise the key itself.  For example,
	// if "!a::Send Test" is a registered hotkey, two things must be done to avoid complications:
	// 1) Prior to sending the word test, ALT must be released in a way that does not activate the
	//    menu bar.  This is done by sandwiching it between a CTRL-down and a CTRL-up.
	// 2) After the send is complete, SendKeys() will restore the ALT key to the down position if
	//    the user is still physically holding ALT down (this is done to make the logical state of
	//    the key match its physical state, which allows the same hotkey to be fired twice in a row
	//    without the user having to release and press down the ALT key physically).
	// The #2 case above is the one handled below by ctrl_wont_be_down.  It is especially necessary
	// when the user releases the ALT key prior to releasing the hotkey suffix, which would otherwise
	// cause the menu bar (if any) of the active window to be activated.
	bool disguise_alt_key = ctrl_not_down && ctrl_wont_be_down; // Since this applies to both Left and Right Alt, don't take g_LayoutHasAltGr into account here. That is done later below.

	// Some of the same comments above for ALT key apply to the WIN key.  More about this issue:
	// Although the disguise of the down-event is usually not needed, it is needed in the rare case
	// where the user releases the WIN or ALT key prior to releasing the hotkey's suffix.
	// Although the hook could be told to disguise the physical release of ALT or WIN in these
	// cases, it's best not to rely on the hook since it is not always installed.

	// Registered WIN and ALT hotkeys that don't use the Send command work okay except ALT hotkeys,
	// which if the user releases ALT prior the hotkey's suffix key, cause the menu bar to be activated.
	// Since it is unusual for users to do this and because it is standard behavior for  ALT hotkeys
	// registered in the OS, fixing it via the hook seems like a low priority, and perhaps isn't worth
	// the added code complexity/size.  But if there is ever a need to do so, the following note applies:
	// If the hook is installed, could tell it to disguise any need-to-be-disguised Alt-up that occurs
	// after receipt of the registered ALT hotkey.  But what if that hotkey uses the send command:
	// there might be intereference?  Doesn't seem so, because the hook only disguises non-ignored events.

	// The WIN key is successfully disguised under a greater number of conditions than ALT:
	bool disguise_win_key = ctrl_not_down && ctrl_wont_be_down
		&& !(aModifiersLRnow & (MOD_LSHIFT | MOD_RSHIFT)) // And neither SHIFT key is down.
		&& (aEventType == KEYUP || !(aModifiersLR & (MOD_LSHIFT | MOD_RSHIFT))) // Nor will it be.
		&& !(aModifiersLRnow & (MOD_LALT | MOD_RALT))     // And neither ALT key is down.
		&& (aEventType == KEYUP || !(aModifiersLR & (MOD_LALT | MOD_RALT))); // Nor will it be.

	// Handle ALT and WIN prior to the other modifiers because the "disguise" methods below are
	// only needed upon release of ALT or WIN.  This is because such releases tend to have a better
	// chance of being "disguised" if SHIFT or CTRL is down at the time of the release.  Thus, the
	// release of SHIFT or CTRL (if called for) is deferred until afterward.
	// Also, WIN is a little more troublesome than ALT, so it is done first in case the ALT key
	// is down but will be going up, since the ALT key being down might help the WIN key.
	// For example, if you hold down CTRL, then hold down LWIN long enough for it to auto-repeat,
	// then release CTRL before releasing LWIN, the Start Menu would appear, at least on XP.
	// But it does not appear if CTRL is released after LWIN.

	// WIN must be handled before ALT because if aEventType indicates that all these keys are being released,
	// we would want to leave ALT down (if it is down and will be going up later) until after LWIN/RWIN
	// are released so that no extra steps are needed to prevent the Start Menu from appearing.  This is
	// because the ALT key can disguise the WIN key, but not vice versa.
	if (aModifiersLR & MOD_LWIN)
	{
		// Fixed for v1.0.25: To avoid triggering the system's LAlt+Shift language hotkey, the
		// Control key is now used to suppress LWIN/RWIN (preventing the Start Menu from appearing)
		// rather than the Shift key.  This is definitely needed for ALT, but is done here for
		// WIN also in case ALT is down, which might cause the use of SHIFT as the disguise key
		// to trigger the language switch.
		if (disguise_win_key)
			KeyEvent(KEYDOWN, VK_CONTROL, 0, NULL, false, aExtraInfo); // Ensures that the Start Menu does not appear.
		KeyEvent(aEventType, VK_LWIN, 0, NULL, false, aExtraInfo);
		if (disguise_win_key)
			KeyEvent(KEYUP, VK_CONTROL, 0, NULL, false, aExtraInfo); // Ensures that the Start Menu does not appear.
		// The above and the next three sections like it must not be done as a KEYDOWNANDUP like
		// SetModifierLRState() uses, since our aEventType might be either down or up.  In the case of DOWN,
		// a the KEYDOWNANDUP for CTRL could come *after* the key is pressed down, probably with the same effect
		// as the above.  However, the above seems slightly simpler and has been tested.
	}
	if (aModifiersLR & MOD_RWIN)
	{
		if (disguise_win_key)
			KeyEvent(KEYDOWN, VK_CONTROL, 0, NULL, false, aExtraInfo); // Ensures that the Start Menu does not appear.
		KeyEvent(aEventType, VK_RWIN, 0, NULL, false, aExtraInfo);
		if (disguise_win_key)
			KeyEvent(KEYUP, VK_CONTROL, 0, NULL, false, aExtraInfo);
	}

	if (aModifiersLR & MOD_LALT)
	{
		if (disguise_alt_key)
			KeyEvent(KEYDOWN, VK_CONTROL, 0, NULL, false, aExtraInfo); // Ensures that menu bar is not activated.
		KeyEvent(aEventType, VK_LMENU, 0, NULL, false, aExtraInfo);
		if (disguise_alt_key)
			KeyEvent(KEYUP, VK_CONTROL, 0, NULL, false, aExtraInfo);
	}
	if (aModifiersLR & MOD_RALT)
	{
		// For the below: There should never be a need to disguise AltGr.  Doing so would likely cause unwanted
		// side-effects. Also, disguise_alt_key does not take g_LayoutHasAltGr into account because
		// disguise_alt_key also applies to the left alt key.
		if (disguise_alt_key && !g_LayoutHasAltGr)
		{
			KeyEvent(KEYDOWN, VK_CONTROL, 0, NULL, false, aExtraInfo); // Ensures that menu bar is not activated.
			KeyEvent(aEventType, VK_RMENU, 0, NULL, false, aExtraInfo);
			KeyEvent(KEYUP, VK_CONTROL, 0, NULL, false, aExtraInfo);
		}
		else // No disguise needed.
		{
			KeyEvent(aEventType, VK_RMENU, 0, NULL, false, aExtraInfo);
			if (g_LayoutHasAltGr) // Note that KeyEvent() might have just changed the value of g_LayoutHasAltGr.
				aModifiersLR &= ~MOD_LCONTROL; // Ensure that the step later below won't send an LControl keystroke (neither up nor down).
		}
	} // RAlt

	if (aModifiersLR & MOD_LSHIFT) KeyEvent(aEventType, VK_LSHIFT, 0, NULL, false, aExtraInfo);
	if (aModifiersLR & MOD_RSHIFT) KeyEvent(aEventType, VK_RSHIFT, 0, NULL, false, aExtraInfo);
	if (aModifiersLR & MOD_LCONTROL) KeyEvent(aEventType, VK_LCONTROL, 0, NULL, false, aExtraInfo);
	if (aModifiersLR & MOD_RCONTROL) KeyEvent(aEventType, VK_RCONTROL, 0, NULL, false, aExtraInfo);

	// See comments at the bottom of SetModifierLRState() about this:
	if (aTargetWindow && GetWindowThreadProcessId(aTargetWindow, NULL) == GetCurrentThreadId())
		SLEEP_WITHOUT_INTERRUPTION(-1)

	// Other notes about disguising ALT and WIN:
	// Registered Alt hotkeys don't quite work if the Alt key is released prior to the suffix.
	// Key history for Alt-B hotkey released this way, which undesirably activates the menu bar:
	// A4  038	 	d	0.03	Alt            	
	// 42  030	 	d	0.03	B              	
	// A4  038	 	u	0.24	Alt            	
	// 42  030	 	u	0.19	B              	
	// Testing shows that the above does not happen for a normal (non-hotkey) alt keystroke such as Alt-8,
	// so the above behavior is probably caused by the fact that B-down is suppressed by the OS's hotkey
	// routine, but not B-up.
	// The above also happens with registered WIN hotkeys, but only if the Send cmd resulted in the WIN
	// modifier being pushed back down afterward to match the fact that the user is still holding it down.
	// This behavior applies to ALT hotkeys also.
	// One solution: if the hook is installed, have it keep track of when the start menu or menu bar
	// *would* be activated.  These tracking vars can be consulted by the Send command, and the hook
	// can also be told when to use them after a registered hotkey has been pressed, so that the Alt-up
	// or Win-up keystroke that belongs to it can be disguised.

	// The following are important ways in which other methods of disguise might not be sufficient:
	// Sequence: shift-down win-down shift-up win-up: invokes Start Menu when WIN is held down long enough
	// to auto-repeat.  Same when Ctrl or Alt is used in lieu of Shift.
	// Sequence: shift-down alt-down alt-up shift-up: invokes menu bar.  However, as long as another key,
	// even Shift, is pressed down *after* alt is pressed down, menu bar is not activated, e.g. alt-down
	// shift-down shift-up alt-up.  In addition, CTRL always prevents ALT from activating the menu bar,
	// even with the following sequences:
	// ctrl-down alt-down alt-up ctrl-up
	// alt-down ctrl-down ctrl-up alt-up
	// (also seems true for all other permutations of Ctrl/Alt)
}



inline mod_type GetModifierState()
{
	return ConvertModifiersLR(GetModifierLRState());
}



modLR_type GetModifierLRState(bool aExplicitlyGet)
// Try to report a more reliable state of the modifier keys than GetKeyboardState
// alone could.
{
	// Rather than old/below method, in light of the fact that new low-level hook is being tried,
	// try relying on only the hook's tracked value rather than calling Get() (if the hook
	// is active:
	if (g_KeybdHook && !aExplicitlyGet)
		return g_modifiersLR_logical;

	// I decided to call GetKeyboardState() rather than tracking the state of these keys with the
	// hook itself because that method wasn't reliable.  Hopefully, this method will always
	// report the correct physical state of the keys (unless the OS itself thinks they're stuck
	// down even when they're physically up, which seems to happen sometimes on some keyboards).
	// It's probably better to make any vars (large ones at least) static for performance reasons.
	// Normal/automatic vars have to be reallocated on the stack every time the function is called.
	// This should be safe because it now seems that a KeyboardProc() is not called re-entrantly
	// (though it is possible it's called that way sometimes?  Couldn't find an answer).
	// This is the alternate, lower-performance method.

	// Now, at the last possible moment (for performance), set the correct status for all
	// the bits in g_modifiersLR_get.

	// Use GetKeyState() rather than GetKeyboardState() because it's the only way to get
	// accurate key state when a console window is active, it seems.  I've also seen other
	// cases where GetKeyboardState() is incorrect (at least under WinXP) when GetKeyState(),
	// in its place, yields the correct info.  Very strange.

	modLR_type modifiersLR = 0;  // Allows all to default to up/off to simplify the below.
	if (g_os.IsWin9x() || g_os.IsWinNT4())
	{
		// Assume it's the left key since there's no way to tell which of the pair it
		// is? (unless the hook is installed, in which case it's value would have already
		// been returned, above).
		if (IsKeyDown9xNT(VK_SHIFT)) modifiersLR |= MOD_LSHIFT;
		if (IsKeyDown9xNT(VK_CONTROL)) modifiersLR |= MOD_LCONTROL;
		if (IsKeyDown9xNT(VK_MENU)) modifiersLR |= MOD_LALT;
		if (IsKeyDown9xNT(VK_LWIN)) modifiersLR |= MOD_LWIN;
		if (IsKeyDown9xNT(VK_RWIN)) modifiersLR |= MOD_RWIN;
	}
	else
	{
		if (IsKeyDown2kXP(VK_LSHIFT)) modifiersLR |= MOD_LSHIFT;
		if (IsKeyDown2kXP(VK_RSHIFT)) modifiersLR |= MOD_RSHIFT;
		if (IsKeyDown2kXP(VK_LCONTROL)) modifiersLR |= MOD_LCONTROL;
		if (IsKeyDown2kXP(VK_RCONTROL)) modifiersLR |= MOD_RCONTROL;
		if (IsKeyDown2kXP(VK_LMENU)) modifiersLR |= MOD_LALT;
		if (IsKeyDown2kXP(VK_RMENU)) modifiersLR |= MOD_RALT;
		if (IsKeyDown2kXP(VK_LWIN)) modifiersLR |= MOD_LWIN;
		if (IsKeyDown2kXP(VK_RWIN)) modifiersLR |= MOD_RWIN;
	}

	if (g_KeybdHook)
	{
		// Since hook is installed, fix any modifiers that it incorrectly thinks are down.
		// Though rare, this situation does arise during periods when the hook cannot track
		// the user's keystrokes, such as when the OS is doing something with the hardware,
		// e.g. switching to TV-out or changing video resolutions.  There are probably other
		// situations where this happens -- never fully explored and identified -- so it
		// seems best to do this, at least when the caller specified aExplicitlyGet = true.
		// To limit the scope of this workaround, only change the state of the hook modifiers
		// to be "up" for those keys the hook thinks are logically down but which the OS thinks
		// are logically up.  Note that it IS possible for a key to be physically down without
		// being logically down (i.e. during a Send command where the user is phyiscally holding
		// down a modifier, but the Send command needs to put it up temporarily), so do not
		// change the hook's physical state for such keys in that case.
		modLR_type hook_wrongly_down = g_modifiersLR_logical & ~modifiersLR;
		if (hook_wrongly_down)
		{
			// Adjust the physical and logical hook state to release the keys that are wrongly down.
			// If a key is wrongly logically down, it seems best to release it both physically and
			// logically, since the hook's failure to see the up-event probably makes its physical
			// state wrong in most such cases.
			g_modifiersLR_physical &= ~hook_wrongly_down;
			g_modifiersLR_logical &= ~hook_wrongly_down;
			g_modifiersLR_logical_non_ignored &= ~hook_wrongly_down;
			// Also adjust physical state so that the GetKeyState command will retrieve the correct values:
			AdjustKeyState(g_PhysicalKeyState, g_modifiersLR_physical);
		}
	}

	return modifiersLR;

	// Only consider a modifier key to be really down if both the hook's tracking of it
	// and GetKeyboardState() agree that it should be down.  The should minimize the impact
	// of the inherent unreliability present in each method (and each method is unreliable in
	// ways different from the other).  I have verified through testing that this eliminates
	// many misfires of hotkeys.  UPDATE: Both methods are fairly reliable now due to starting
	// to send scan codes with keybd_event(), using MapVirtualKey to resolve zero-value scan
	// codes in the keyboardproc(), and using GetKeyState() rather than GetKeyboardState().
	// There are still a few cases when they don't agree, so return the bitwise-and of both
	// if the keyboard hook is active.  Bitwise and is used because generally it's safer
	// to assume a modifier key is up, when in doubt (e.g. to avoid firing unwanted hotkeys):
//	return g_KeybdHook ? (g_modifiersLR_logical & g_modifiersLR_get) : g_modifiersLR_get;
}



void AdjustKeyState(BYTE aKeyState[], modLR_type aModifiersLR)
// Caller has ensured that aKeyState is a 256-BYTE array of key states, in the same format used
// by GetKeyboardState() and ToAscii().
{
	aKeyState[VK_LSHIFT] = (aModifiersLR & MOD_LSHIFT) ? STATE_DOWN : 0;
	aKeyState[VK_RSHIFT] = (aModifiersLR & MOD_RSHIFT) ? STATE_DOWN : 0;
	aKeyState[VK_LCONTROL] = (aModifiersLR & MOD_LCONTROL) ? STATE_DOWN : 0;
	aKeyState[VK_RCONTROL] = (aModifiersLR & MOD_RCONTROL) ? STATE_DOWN : 0;
	aKeyState[VK_LMENU] = (aModifiersLR & MOD_LALT) ? STATE_DOWN : 0;
	aKeyState[VK_RMENU] = (aModifiersLR & MOD_RALT) ? STATE_DOWN : 0;
	aKeyState[VK_LWIN] = (aModifiersLR & MOD_LWIN) ? STATE_DOWN : 0;
	aKeyState[VK_RWIN] = (aModifiersLR & MOD_RWIN) ? STATE_DOWN : 0;
	// Update the state of neutral keys only after the above, in case both keys of the pair were wrongly down:
	aKeyState[VK_SHIFT] = (aKeyState[VK_LSHIFT] || aKeyState[VK_RSHIFT]) ? STATE_DOWN : 0;
	aKeyState[VK_CONTROL] = (aKeyState[VK_LCONTROL] || aKeyState[VK_RCONTROL]) ? STATE_DOWN : 0;
	aKeyState[VK_MENU] = (aKeyState[VK_LMENU] || aKeyState[VK_RMENU]) ? STATE_DOWN : 0;
}



modLR_type KeyToModifiersLR(vk_type aVK, sc_type aSC, bool *pIsNeutral)
// Convert the given virtual key / scan code to its equivalent bitwise modLR value.
// Callers rely upon the fact that we convert a neutral key such as VK_SHIFT into MOD_LSHIFT,
// not the bitwise combo of MOD_LSHIFT|MOD_RSHIFT.
{
	if (pIsNeutral) *pIsNeutral = false;  // Set default for this output param, unless overridden later.
	if (!aVK && !aSC) return 0;

	if (aVK) // Have vk take precedence over any non-zero sc.
		switch(aVK)
		{
		case VK_SHIFT: if (pIsNeutral) *pIsNeutral = true; return MOD_LSHIFT;
		case VK_LSHIFT: return MOD_LSHIFT;
		case VK_RSHIFT:	return MOD_RSHIFT;
		case VK_CONTROL: if (pIsNeutral) *pIsNeutral = true; return MOD_LCONTROL;
		case VK_LCONTROL: return MOD_LCONTROL;
		case VK_RCONTROL: return MOD_RCONTROL;
		case VK_MENU: if (pIsNeutral) *pIsNeutral = true; return MOD_LALT;
		case VK_LMENU: return MOD_LALT;
		case VK_RMENU: return MOD_RALT;
		case VK_LWIN: return MOD_LWIN;
		case VK_RWIN: return MOD_RWIN;
		default: return 0;
		}
	// If above didn't return, rely on the non-zero sc instead:
	switch(aSC)
	{
	case SC_LSHIFT: return MOD_LSHIFT;
	case SC_RSHIFT:	return MOD_RSHIFT;
	case SC_LCONTROL: return MOD_LCONTROL;
	case SC_RCONTROL: return MOD_RCONTROL;
	case SC_LALT: return MOD_LALT;
	case SC_RALT: return MOD_RALT;
	case SC_LWIN: return MOD_LWIN;
	case SC_RWIN: return MOD_RWIN;
	}
	return 0;
}



modLR_type ConvertModifiers(mod_type aModifiers)
// Convert the input param to a modifiersLR value and return it.
{
	modLR_type modifiersLR = 0;
	if (aModifiers & MOD_WIN) modifiersLR |= (MOD_LWIN | MOD_RWIN);
	if (aModifiers & MOD_ALT) modifiersLR |= (MOD_LALT | MOD_RALT);
	if (aModifiers & MOD_CONTROL) modifiersLR |= (MOD_LCONTROL | MOD_RCONTROL);
	if (aModifiers & MOD_SHIFT) modifiersLR |= (MOD_LSHIFT | MOD_RSHIFT);
	return modifiersLR;
}



mod_type ConvertModifiersLR(modLR_type aModifiersLR)
// Convert the input param to a normal modifiers value and return it.
{
	mod_type modifiers = 0;
	if (aModifiersLR & (MOD_LWIN | MOD_RWIN)) modifiers |= MOD_WIN;
	if (aModifiersLR & (MOD_LALT | MOD_RALT)) modifiers |= MOD_ALT;
	if (aModifiersLR & (MOD_LSHIFT | MOD_RSHIFT)) modifiers |= MOD_SHIFT;
	if (aModifiersLR & (MOD_LCONTROL | MOD_RCONTROL)) modifiers |= MOD_CONTROL;
	return modifiers;
}



char *ModifiersLRToText(modLR_type aModifiersLR, char *aBuf)
// Caller has ensured that aBuf is not NULL.
{
	*aBuf = '\0';
	if (aModifiersLR & MOD_LWIN) strcat(aBuf, "LWin ");
	if (aModifiersLR & MOD_RWIN) strcat(aBuf, "RWin ");
	if (aModifiersLR & MOD_LSHIFT) strcat(aBuf, "LShift ");
	if (aModifiersLR & MOD_RSHIFT) strcat(aBuf, "RShift ");
	if (aModifiersLR & MOD_LCONTROL) strcat(aBuf, "LCtrl ");
	if (aModifiersLR & MOD_RCONTROL) strcat(aBuf, "RCtrl ");
	if (aModifiersLR & MOD_LALT) strcat(aBuf, "LAlt ");
	if (aModifiersLR & MOD_RALT) strcat(aBuf, "RAlt ");
	return aBuf;
}



char *SCtoKeyName(sc_type aSC, char *aBuf, int aBufSize)
// aBufSize is an int so that any negative values passed in from caller are not lost.
// Always produces a non-empty string.
{
	for (int i = 0; i < g_key_to_sc_count; ++i)
	{
		if (g_key_to_sc[i].sc == aSC)
		{
			strlcpy(aBuf, g_key_to_sc[i].key_name, aBufSize);
			return aBuf;
		}
	}
	// Since above didn't return, no match was found.  Use the default format for an unknown scan code:
	snprintf(aBuf, aBufSize, "SC%03x", aSC);
	return aBuf;
}



char *VKtoKeyName(vk_type aVK, sc_type aSC, char *aBuf, int aBufSize)
// aBufSize is an int so that any negative values passed in from caller are not lost.
// Caller may omit aSC and it will be derived if needed.
{
	for (int i = 0; i < g_key_to_vk_count; ++i)
	{
		if (g_key_to_vk[i].vk == aVK)
		{
			strlcpy(aBuf, g_key_to_vk[i].key_name, aBufSize);
			return aBuf;
		}
	}
	// Since above didn't return, no match was found.  Ask the OS for the name instead (it's probably
	// a letter key such as A through Z, but could be anything for which we don't have a listing):
	return GetKeyName(aVK, aSC, aBuf, aBufSize);
}



sc_type TextToSC(char *aText)
{
	if (!aText || !*aText) return 0;
	for (int i = 0; i < g_key_to_sc_count; ++i)
		if (!stricmp(g_key_to_sc[i].key_name, aText))
			return g_key_to_sc[i].sc;
	// Do this only after the above, in case any valid key names ever start with SC:
	if (toupper(*aText) == 'S' && toupper(*(aText + 1)) == 'C')
		return strtol(aText + 2, NULL, 16);  // Convert from hex.
	return 0; // Indicate "not found".
}



vk_type TextToVK(char *aText, modLR_type *pModifiersLR, bool aExcludeThoseHandledByScanCode, bool aAllowExplicitVK)
// If modifiers_p is non-NULL, place the modifiers that are needed to realize the key in there.
// e.g. M is really +m (shift-m), # is really shift-3.
{
	if (!aText || !*aText) return 0;

	// Don't trim() aText or modify it because that will mess up the caller who expects it to be unchanged.
	// Instead, for now, just check it as-is.  The only extra whitespace that should exist, due to trimming
	// of text during load, is that on either side of the COMPOSITE_DELIMITER (e.g. " then ").

	if (strlen(aText) == 1)
	{
		vk_type vk;
		char keyscan_modifiers;
		if (*aText == '\n')
		{
			// For v1.0.25.12, it seems best to avoid the many recent problems with linefeed (`n) being sent
			// as Ctrl+Enter by changing it to always send a plain Enter, just like carriage return (`r).
			vk = VK_RETURN;
			keyscan_modifiers = 0;
		}
		else
		{
			SHORT mod_plus_vk = VkKeyScan(*aText);
			vk = LOBYTE(mod_plus_vk);
			keyscan_modifiers = HIBYTE(mod_plus_vk);
			if (keyscan_modifiers == -1 && vk == (UCHAR)-1) // No translation could be made.
				return 0;
		}

		// For v1.0.35, pModifiersLR was changed to modLR vs. mod so that AltGr keys such as backslash and {
		// are supported on layouts such as German when sending to apps such as Putty that are fussy about
		// which ALT key is held down to produce the character.  The following section detects AltGr by the
		// assuming that any character that requires both CTRL and ALT (with optional SHIFT) to be held
		// down is in fact an AltGr key (I don't think there are any that aren't AltGr in this case, but
		// confirmation would be nice).  Also, this is not done for Win9x because the distinction between
		// right and left-alt is not well-supported and it might do more harm than good (testing is
		// needed on fussy apps like Putty on Win9x).  UPDATE: Windows NT4 is now excluded from this
		// change because apparently it wants the left Alt key's virtual key and not the right's (though
		// perhaps it would prefer the right scan code vs. the left in apps such as Putty, but until that
		// is proven, the complexity is not added here).  Otherwise, on French and other layouts on NT4,
		// AltGr-produced characters such as backslash do not get sent properly.  In hindsight, this is
		// not suprising because the keyboard hook also receives neutral modifier keys on NT4 rather than
		// a more specific left/right key.

		// For v1.0.36.06: The following should be a 99% reliable indicator that current layout has an AltGr key.
		// But is there a better way?  Maybe could use IOCTL to query the keyboard driver's AltGr flag whenever
		// the main event loop receives WM_INPUTLANGCHANGEREQUEST.  Making such a thing work on both Win9x and
		// 2000/XP might be an issue.
		bool requires_altgr = (keyscan_modifiers & 0x06) == 0x06;
		if (requires_altgr) // This character requires both CTRL and ALT (and possibly SHIFT, since I think Shift+AltGr combinations exist).
			g_LayoutHasAltGr = true;

		// The win docs for VkKeyScan() are a bit confusing, referring to flag "bits" when it should really
		// say flag "values".  In addition, it seems that these flag values are incompatible with
		// MOD_ALT, MOD_SHIFT, and MOD_CONTROL, so they must be translated:
		if (pModifiersLR) // The caller wants this info added to the output param.
		{
			// Best not to reset this value because some callers want to retain what was in it before,
			// merely merging these new values into it:
			//*pModifiers = 0;
			if (requires_altgr && g_os.IsWin2000orLater())
			{
				// v1.0.35: The critical difference below is right vs. left ALT.  Must not include MOD_LCONTROL
				// because simulating the RAlt keystroke on these keyboard layouts will automatically
				// press LControl down.
				*pModifiersLR |= MOD_RALT;
				if (keyscan_modifiers & 0x01) // Added for v1.0.36.06 because presence of AltGr should not preclude the presence of Shift.
					*pModifiersLR |= MOD_LSHIFT;
			}
			else // Do normal/default translation.
			{
				if (keyscan_modifiers & 0x01)
					*pModifiersLR |= MOD_LSHIFT;
				if (keyscan_modifiers & 0x02)
					*pModifiersLR |= MOD_LCONTROL;
				if (keyscan_modifiers & 0x04)
					*pModifiersLR |= MOD_LALT;
			}
		}
		return vk;
	}

// Use above in favor of this:
//	if (strlen(text) == 1 && toupper(*text) >= 'A' && toupper(*text) <= 'Z')
//		return toupper(*text);  // VK is the same as the ASCII code in this case, maybe for other chars too?

	if (aAllowExplicitVK && toupper(aText[0]) == 'V' && toupper(aText[1]) == 'K')
		return (vk_type)strtol(aText + 2, NULL, 16);  // Convert from hex.

	for (int i = 0; i < g_key_to_vk_count; ++i)
		if (!stricmp(g_key_to_vk[i].key_name, aText))
			return g_key_to_vk[i].vk;

	if (aExcludeThoseHandledByScanCode)
		return 0; // Zero is not a valid virtual key, so it should be a safe failure indicator.

	// Otherwise check if aText is the name of a key handled by scan code and if so, map that
	// scan code to its corresponding virtual key:
	sc_type sc = TextToSC(aText);
	return sc ? sc_to_vk(sc) : 0;
}



vk_type TextToSpecial(char *aText, UINT aTextLength, KeyEventTypes &aEventType, modLR_type &aModifiersLR
	, mod_type &aModifiers, bool aUpdatePersistent)
// Returns vk for key-down, negative vk for key-up, or zero if no translation.
// We also update whatever's in *pModifiers and *pModifiersLR to reflect the type of key-action
// specified in <aText>.  This makes it so that {altdown}{esc}{altup} behaves the same as !{esc}.
// Note that things like LShiftDown are not supported because: 1) they are rarely needed; and 2)
// they can be down via "lshift down".
{
	if (!strlicmp(aText, "ALTDOWN", aTextLength))
	{
		if (aUpdatePersistent)
		{
			if (!(aModifiersLR & (MOD_LALT | MOD_RALT))) // i.e. do nothing if either left or right is already present.
				aModifiersLR |= MOD_LALT; // If neither is down, use the left one because it's more compatible.
			aModifiers |= MOD_ALT;
		}
		aEventType = KEYDOWN;
		return VK_MENU;
	}
	if (!strlicmp(aText, "ALTUP", aTextLength))
	{
		// Unlike for Lwin/Rwin, it seems best to have these neutral keys (e.g. ALT vs. LALT or RALT)
		// restore either or both of the ALT keys into the up position.  The user can use {LAlt Up}
		// to be more specific and avoid this behavior:
		if (aUpdatePersistent)
		{
			aModifiersLR &= ~(MOD_LALT | MOD_RALT);
			aModifiers &= ~MOD_ALT;
		}
		aEventType = KEYUP;
		return VK_MENU;
	}
	if (!strlicmp(aText, "SHIFTDOWN", aTextLength))
	{
		if (aUpdatePersistent)
		{
			if (!(aModifiersLR & (MOD_LSHIFT | MOD_RSHIFT))) // i.e. do nothing if either left or right is already present.
				aModifiersLR |= MOD_LSHIFT; // If neither is down, use the left one because it's more compatible.
			aModifiers |= MOD_SHIFT;
		}
		aEventType = KEYDOWN;
		return VK_SHIFT;
	}
	if (!strlicmp(aText, "SHIFTUP", aTextLength))
	{
		if (aUpdatePersistent)
		{
			aModifiersLR &= ~(MOD_LSHIFT | MOD_RSHIFT); // See "ALTUP" for explanation.
			aModifiers &= ~MOD_SHIFT;
		}
		aEventType = KEYUP;
		return VK_SHIFT;
	}
	if (!strlicmp(aText, "CTRLDOWN", aTextLength) || !strlicmp(aText, "CONTROLDOWN", aTextLength))
	{
		if (aUpdatePersistent)
		{
			if (!(aModifiersLR & (MOD_LCONTROL | MOD_RCONTROL))) // i.e. do nothing if either left or right is already present.
				aModifiersLR |= MOD_LCONTROL; // If neither is down, use the left one because it's more compatible.
			aModifiers |= MOD_CONTROL;
		}
		aEventType = KEYDOWN;
		return VK_CONTROL;
	}
	if (!strlicmp(aText, "CTRLUP", aTextLength) || !strlicmp(aText, "CONTROLUP", aTextLength))
	{
		if (aUpdatePersistent)
		{
			aModifiersLR &= ~(MOD_LCONTROL | MOD_RCONTROL); // See "ALTUP" for explanation.
			aModifiers &= ~MOD_CONTROL;
		}
		aEventType = KEYUP;
		return VK_CONTROL;
	}
	if (!strlicmp(aText, "LWINDOWN", aTextLength))
	{
		if (aUpdatePersistent)
		{
			aModifiersLR |= MOD_LWIN;
			aModifiers |= MOD_WIN;
		}
		aEventType = KEYDOWN;
		return VK_LWIN;
	}
	if (!strlicmp(aText, "LWINUP", aTextLength))
	{
		if (aUpdatePersistent)
		{
			aModifiersLR &= ~MOD_LWIN;
			if (!(aModifiersLR & MOD_RWIN))  // If both WIN keys are now up, the neutral modifier also is set to up.
				aModifiers &= ~MOD_WIN;
		}
		aEventType = KEYUP;
		return VK_LWIN;
	}
	if (!strlicmp(aText, "RWINDOWN", aTextLength))
	{
		if (aUpdatePersistent)
		{
			aModifiersLR |= MOD_RWIN;
			aModifiers |= MOD_WIN;
		}
		aEventType = KEYDOWN;
		return VK_RWIN;
	}
	if (!strlicmp(aText, "RWINUP", aTextLength))
	{
		if (aUpdatePersistent)
		{
			aModifiersLR &= ~MOD_RWIN;
			if (!(aModifiersLR & MOD_LWIN))  // If both WIN keys are now up, the neutral modifier also is set to up.
				aModifiers &= ~MOD_WIN;
		}
		aEventType = KEYUP;
		return VK_RWIN;
	}
	// Otherwise, leave aEventType unchanged and return zero to indicate failure:
	return 0;
}



#ifdef ENABLE_KEY_HISTORY_FILE
ResultType KeyHistoryToFile(char *aFilespec, char aType, bool aKeyUp, vk_type aVK, sc_type aSC)
{
	static char target_filespec[MAX_PATH] = "";
	static FILE *fp = NULL;
	static HWND last_foreground_window = NULL;
	static DWORD last_tickcount = GetTickCount();

	if (!g_KeyHistory) // Since key history is disabled, keys are not being tracked by the hook, so there's nothing to log.
		return OK;     // Files should not need to be closed since they would never have been opened in the first place.

	if (!aFilespec && !aVK && !aSC) // Caller is signaling to close the file if it's open.
	{
		if (fp)
		{
			fclose(fp);
			fp = NULL;
		}
		return OK;
	}

	if (aFilespec && *aFilespec && stricmp(aFilespec, target_filespec)) // Target filename has changed.
	{
		if (fp)
		{
			fclose(fp);
			fp = NULL;  // To indicate to future calls to this function that it's closed.
		}
		strlcpy(target_filespec, aFilespec, sizeof(target_filespec));
	}

	if (!aVK && !aSC) // Caller didn't want us to log anything this time.
		return OK;
	if (!*target_filespec)
		return OK; // No target filename has ever been specified, so don't even attempt to open the file.

	if (!aVK)
		aVK = sc_to_vk(aSC);
	else
		if (!aSC)
			aSC = vk_to_sc(aVK);

	char buf[2048] = "", win_title[1024] = "<Init>", key_name[128] = "";
	HWND curr_foreground_window = GetForegroundWindow();
	DWORD curr_tickcount = GetTickCount();
	bool log_changed_window = (curr_foreground_window != last_foreground_window);
	if (log_changed_window)
	{
		if (curr_foreground_window)
			GetWindowText(curr_foreground_window, win_title, sizeof(win_title));
		else
			strlcpy(win_title, "<None>", sizeof(win_title));
		last_foreground_window = curr_foreground_window;
	}

	snprintf(buf, sizeof(buf), "%02X" "\t%03X" "\t%0.2f" "\t%c" "\t%c" "\t%s" "%s%s\n"
		, aVK, aSC
		, (float)(curr_tickcount - last_tickcount) / (float)1000
		, aType
		, aKeyUp ? 'u' : 'd'
		, GetKeyName(aVK, aSC, key_name, sizeof(key_name))
		, log_changed_window ? "\t" : ""
		, log_changed_window ? win_title : ""
		);
	last_tickcount = curr_tickcount;
	if (!fp)
		if (   !(fp = fopen(target_filespec, "a"))   )
			return OK;
	fputs(buf, fp);
	return OK;
}
#endif



char *GetKeyName(vk_type aVK, sc_type aSC, char *aBuf, int aBufSize)
// aBufSize is an int so that any negative values passed in from caller are not lost.
// Caller has ensured that aBuf isn't NULL.
{
	if (aBufSize < 3)
		return aBuf;

	*aBuf = '\0'; // Set default.
	if (!aVK && !aSC)
		return aBuf;

	if (!aVK)
		aVK = sc_to_vk(aSC);
	else
		if (!aSC)
			aSC = vk_to_sc(aVK);

	// Use 0x02000000 to tell it that we want it to give left/right specific info, lctrl/rctrl etc.
	if (!aSC || !GetKeyNameText((long)(aSC) << 16, aBuf, (int)(aBufSize/sizeof(TCHAR))))
	{
		int j;
		for (j = 0; j < g_key_to_vk_count; ++j)
			if (g_key_to_vk[j].vk == aVK)
				break;
		if (j < g_key_to_vk_count) // Match found.
			strlcpy(aBuf, g_key_to_vk[j].key_name, aBufSize);
		else
		{
			if (isprint(aVK))
			{
				aBuf[0] = aVK;
				aBuf[1] = '\0';
			}
			else
				strlcpy(aBuf, "not found", aBufSize);
		}
	}
	return aBuf;
}



sc_type vk_to_sc(vk_type aVK, bool aReturnSecondary)
// For v1.0.37.03, vk_to_sc() was converted into a function rather than being an array because if the
// script's keyboard layout changes while it's running, the array would get out-of-date.
// If caller passes true for aReturnSecondary, the non-primary scan code will be returned for
// virtual keys that two scan codes (if there's only one scan code, callers rely on zero being returned).
{
	// Try to minimize the number mappings done manually because MapVirtualKey is a more reliable
	// way to get the mapping if user has non-standard or custom keyboard layout.

	sc_type sc = 0;

	switch (aVK)
	{
	// Yield a manually translation for virtual keys that MapVirtualKey() doesn't support or for which it
	// doesn't yield consistent result (such as Win9x supporting only SHIFT rather than VK_LSHIFT/VK_RSHIFT).
	case VK_LSHIFT:   sc = SC_LSHIFT; break; // Modifiers are listed first for performance.
	case VK_RSHIFT:   sc = SC_RSHIFT; break;
	case VK_LCONTROL: sc = SC_LCONTROL; break;
	case VK_RCONTROL: sc = SC_RCONTROL; break;
	case VK_LMENU:    sc = SC_LALT; break;
	case VK_RMENU:    sc = SC_RALT; break;
	case VK_LWIN:     sc = SC_LWIN; break; // Earliest versions of Win95/NT might not support these, so map them manually.
	case VK_RWIN:     sc = SC_RWIN; break; //

	// According to http://support.microsoft.com/default.aspx?scid=kb;en-us;72583
	// most or all numeric keypad keys cannot be mapped reliably under any OS. The article is
	// a little unclear about which direction, if any, that MapVirtualKey() does work in for
	// the numpad keys, so for peace-of-mind map them all manually for now:
	case VK_NUMPAD0:  sc = SC_NUMPAD0; break;
	case VK_NUMPAD1:  sc = SC_NUMPAD1; break;
	case VK_NUMPAD2:  sc = SC_NUMPAD2; break;
	case VK_NUMPAD3:  sc = SC_NUMPAD3; break;
	case VK_NUMPAD4:  sc = SC_NUMPAD4; break;
	case VK_NUMPAD5:  sc = SC_NUMPAD5; break;
	case VK_NUMPAD6:  sc = SC_NUMPAD6; break;
	case VK_NUMPAD7:  sc = SC_NUMPAD7; break;
	case VK_NUMPAD8:  sc = SC_NUMPAD8; break;
	case VK_NUMPAD9:  sc = SC_NUMPAD9; break;
	case VK_DECIMAL:  sc = SC_NUMPADDOT; break;
	case VK_NUMLOCK:  sc = SC_NUMLOCK; break;
	case VK_DIVIDE:   sc = SC_NUMPADDIV; break;
	case VK_MULTIPLY: sc = SC_NUMPADMULT; break;
	case VK_SUBTRACT: sc = SC_NUMPADSUB; break;
	case VK_ADD:      sc = SC_NUMPADADD; break;
	}

	if (sc) // Above found a match.
		return aReturnSecondary ? 0 : sc; // Callers rely on zero being returned for VKs that don't have secondary SCs.

	// Use the OS API's MapVirtualKey() to resolve any not manually done above:
	if (   !(sc = MapVirtualKey(aVK, 0))   )
		return 0; // Indicate "no mapping".

	// Turn on the extended flag for those that need it.
	// Because MapVirtualKey can only accept (and return) naked scan codes (the low-order byte),
	// handle extended scan codes that have a non-extended counterpart manually.
	// Older comment: MapVirtualKey() should include 0xE0 in HIBYTE if key is extended, BUT IT DOESN'T.
	// There doesn't appear to be any built-in function to determine whether a vk's scan code
	// is extended or not.  See MSDN topic "keyboard input" for the below list.
	// Note: NumpadEnter is probably the only extended key that doesn't have a unique VK of its own.
	// So in that case, probably safest not to set the extended flag.  To send a true NumpadEnter,
	// as well as a true NumPadDown and any other key that shares the same VK with another, the
	// caller should specify the sc param to circumvent the need for KeyEvent() to use the below:
	switch (aVK)
	{
	case VK_APPS:     // Application key on keyboards with LWIN/RWIN/Apps.  Not listed in MSDN as "extended"?
	case VK_CANCEL:   // Ctrl-break
	case VK_SNAPSHOT: // PrintScreen
	case VK_DIVIDE:   // NumpadDivide (slash)
	case VK_NUMLOCK:
	// Below are extended but were already handled and returned from higher above:
	//case VK_LWIN:
	//case VK_RWIN:
	//case VK_RMENU:
	//case VK_RCONTROL:
	//case VK_RSHIFT: // WinXP needs this to be extended for keybd_event() to work properly.
		sc |= 0x0100;
		break;

	// The following virtual keys have more than one physical key, and thus more than one scan code.
	// If the caller passed true for aReturnSecondary, the extended version of the scan code will be
	// returned (all of the following VKs have two SCs):
	case VK_RETURN:
	case VK_INSERT:
	case VK_DELETE:
	case VK_PRIOR: // PgUp
	case VK_NEXT:  // PgDn
	case VK_HOME:
	case VK_END:
	case VK_UP:
	case VK_DOWN:
	case VK_LEFT:
	case VK_RIGHT:
		return aReturnSecondary ? (sc | 0x0100) : sc; // Below relies on the fact that these cases return early.
	}

	// Since above didn't return, if aReturnSecondary==true, return 0 to indicate "no secondary SC for this VK".
	return aReturnSecondary ? 0 : sc; // Callers rely on zero being returned for VKs that don't have secondary SCs.
}



vk_type sc_to_vk(sc_type aSC)
{
	// These are mapped manually because MapVirtualKey() doesn't support them correctly, at least
	// on some -- if not all -- OSs.  The main app also relies upon the values assigned below to
	// determine which keys should be handled by scan code rather than vk:
	switch (aSC)
	{
	// Even though neither of the SHIFT keys are extended -- and thus could be mapped with MapVirtualKey()
	// -- it seems better to define them explicitly because under Win9x (maybe just Win95).
	// I'm pretty sure MapVirtualKey() would return VK_SHIFT instead of the left/right VK.
	case SC_LSHIFT:      return VK_LSHIFT; // Modifiers are listed first for performance.
	case SC_RSHIFT:      return VK_RSHIFT;
	case SC_LCONTROL:    return VK_LCONTROL;
	case SC_RCONTROL:    return VK_RCONTROL;
	case SC_LALT:        return VK_LMENU;
	case SC_RALT:        return VK_RMENU;

	// Numpad keys require explicit mapping because MapVirtualKey() doesn't support them on all OSes.
	// See comments in vk_to_sc() for details.
	case SC_NUMLOCK:     return VK_NUMLOCK;
	case SC_NUMPADDIV:   return VK_DIVIDE;
	case SC_NUMPADMULT:  return VK_MULTIPLY;
	case SC_NUMPADSUB:   return VK_SUBTRACT;
	case SC_NUMPADADD:   return VK_ADD;
	case SC_NUMPADENTER: return VK_RETURN;

	// The following are ambiguous because each maps to more than one VK.  But be careful
	// changing the value to the other choice because some callers rely upon the values
	// assigned below to determine which keys should be handled by scan code rather than vk:
	case SC_NUMPADDEL:   return VK_DELETE;
	case SC_NUMPADCLEAR: return VK_CLEAR;
	case SC_NUMPADINS:   return VK_INSERT;
	case SC_NUMPADUP:    return VK_UP;
	case SC_NUMPADDOWN:  return VK_DOWN;
	case SC_NUMPADLEFT:  return VK_LEFT;
	case SC_NUMPADRIGHT: return VK_RIGHT;
	case SC_NUMPADHOME:  return VK_HOME;
	case SC_NUMPADEND:   return VK_END;
	case SC_NUMPADPGUP:  return VK_PRIOR;
	case SC_NUMPADPGDN:  return VK_NEXT;

	// No callers currently need the following alternate virtual key mappings.  If it is ever needed,
	// could have a new aReturnSecondary parameter that if true, causes these to be returned rather
	// than the above:
	//case SC_NUMPADDEL:   return VK_DECIMAL;
	//case SC_NUMPADCLEAR: return VK_NUMPAD5; // Same key as Numpad5 on most keyboards?
	//case SC_NUMPADINS:   return VK_NUMPAD0;
	//case SC_NUMPADUP:    return VK_NUMPAD8;
	//case SC_NUMPADDOWN:  return VK_NUMPAD2;
	//case SC_NUMPADLEFT:  return VK_NUMPAD4;
	//case SC_NUMPADRIGHT: return VK_NUMPAD6;
	//case SC_NUMPADHOME:  return VK_NUMPAD7;
	//case SC_NUMPADEND:   return VK_NUMPAD1;
	//case SC_NUMPADPGUP:  return VK_NUMPAD9;
	//case SC_NUMPADPGDN:  return VK_NUMPAD3;	
	}

	// Use the OS API call to resolve any not manually set above.  This should correctly
	// resolve even elements such as SC_INSERT, which is an extended scan code, because
	// it passes in only the low-order byte which is SC_NUMPADINS.  In the case of SC_INSERT
	// and similar ones, MapVirtualKey() will return the same vk for both, which is correct.
	// Only pass the LOBYTE because I think it fails to work properly otherwise.
	// Also, DO NOT pass 3 for the 2nd param of MapVirtualKey() because apparently
	// that is not compatible with Win9x so it winds up returning zero for keys
	// such as UP, LEFT, HOME, and PGUP (maybe other sorts of keys too).  This
	// should be okay even on XP because the left/right specific keys have already
	// been resolved above so don't need to be looked up here (LWIN and RWIN
	// each have their own VK's so shouldn't be problem for the below call to resolve):
	return MapVirtualKey((BYTE)aSC, 1);
}



bool IsMouseVK(vk_type aVK)
{
	switch (aVK)
	{
	case VK_LBUTTON:
	case VK_RBUTTON:
	case VK_MBUTTON:
	case VK_XBUTTON1:
	case VK_XBUTTON2:
	case VK_WHEEL_DOWN:
	case VK_WHEEL_UP:
		return true;
	}
	return false;
}
