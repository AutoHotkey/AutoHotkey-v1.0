/*
AutoHotkey

Copyright 2003-2006 Chris Mallett (support@autohotkey.com)

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
#include "application.h"
#include "globaldata.h" // for access to g_clip, the "g" global struct, etc.
#include "window.h" // for serveral MsgBox and window functions
#include "util.h" // for strlcpy()
#include "resources\resource.h"  // For ID_TRAY_OPEN.


bool MsgSleep(int aSleepDuration, MessageMode aMode)
// Returns true if it launched at least one thread, and false otherwise.
// aSleepDuration can be be zero to do a true Sleep(0), or less than 0 to avoid sleeping or
// waiting at all (i.e. messages are checked and if there are none, the function will return
// immediately).  aMode is either RETURN_AFTER_MESSAGES (default) or WAIT_FOR_MESSAGES.
// If the caller doesn't specify aSleepDuration, this function will return after a
// time less than or equal to SLEEP_INTERVAL (i.e. the exact amount of the sleep
// isn't important to the caller).  This mode is provided for performance reasons
// (it avoids calls to GetTickCount and the TickCount math).  However, if the
// caller's script subroutine is suspended due to action by us, an unknowable
// amount of time may pass prior to finally returning to the caller.
{
	// This is done here for performance reasons.  UPDATE: This probably never needs
	// to close the clipboard now that Line::ExecUntil() also calls CLOSE_CLIPBOARD_IF_OPEN:
	CLOSE_CLIPBOARD_IF_OPEN;

	// While in mode RETURN_AFTER_MESSAGES, there are different things that can happen:
	// 1) We launch a new hotkey subroutine, interrupting/suspending the old one.  But
	//    subroutine calls this function again, so now it's recursed.  And thus the
	//    new subroutine can be interrupted yet again.
	// 2) We launch a new hotkey subroutine, but it returns before any recursed call
	//    to this function discovers yet another hotkey waiting in the queue.  In this
	//    case, this instance/recursion layer of the function should process the
	//    hotkey messages linearly rather than recursively?  No, this doesn't seem
	//    necessary, because we can just return from our instance/layer and let the
	//    caller handle any messages waiting in the queue.  Eventually, the queue
	//    should be emptied, especially since most hotkey subroutines will run
	//    much faster than the user could press another hotkey, with the possible
	//    exception of the key-repeat feature triggered by holding a key down.
	//    Even in that case, the worst that would happen is that messages would
	//    get dropped off the queue because they're too old (I think that's what
	//    happens).
	// Based on the above, when mode is RETURN_AFTER_MESSAGES, we process
	// all messages until a hotkey message is encountered, at which time we
	// launch that subroutine only and then return when it returns to us, letting
	// the caller handle any additional messages waiting on the queue.  This avoids
	// the need to have a "run the hotkeys linearly" mode in a single iteration/layer
	// of this function.  Note: The WM_QUIT message does not receive any higher
	// precedence in the queue than other messages.  Thus, if there's ever concern
	// that that message would be lost, as a future change perhaps can use PeekMessage()
	// with a filter to explicitly check to see if our queue has a WM_QUIT in it
	// somewhere, prior to processing any messages that might take result in
	// a long delay before the remainder of the queue items are processed (there probably
	// aren't any such conditions now, so nothing to worry about?)

	// Above is somewhat out-of-date.  The objective now is to spend as much time
	// inside GetMessage() as possible, since it's the keystroke/mouse engine
	// whenever the hooks are installed.  Any time we're not in GetMessage() for
	// any length of time (say, more than 20ms), keystrokes and mouse events
	// will be lagged.  PeekMessage() is probably almost as good, but it probably
	// only clears out any waiting keys prior to returning.  CONFIRMED: PeekMessage()
	// definitely routes to the hook, perhaps only if called regularly (i.e. a single
	// isolated call might not help much).

	// This var allows us to suspend the currently-running subroutine and run any
	// hotkey events waiting in the message queue (if there are more than one, they
	// will be executed in sequence prior to resuming the suspended subroutine).
	// Never static because we could be recursed (e.g. when one hotkey iterruptes
	// a hotkey that has already been interrupted) and each recursion layer should
	// have it's own value for this:
	global_struct global_saved;

	// Decided to support a true Sleep(0) for aSleepDuration == 0, as well
	// as no delay at all if aSleepDuration < 0.  This is needed to implement
	// "SetKeyDelay, 0" and possibly other things.  I believe a Sleep(0)
	// is always <= Sleep(1) because both of these will wind up waiting
	// a full timeslice if the CPU is busy.

	// Reminder for anyone maintaining or revising this code:
	// Giving each subroutine its own thread rather than suspending old ones is
	// probably not a good idea due to the exclusive nature of the GUI
	// (i.e. it's probably better to suspend existing subroutines rather than
	// letting them continue to run because they might activate windows and do
	// other stuff that would interfere with the window automation activities of
	// other threads)

	// If caller didn't specify, the exact amount of the Sleep() isn't
	// critical to it, only that we handles messages and do Sleep()
	// a little.
	// Most of this initialization section isn't needed if aMode == WAIT_FOR_MESSAGES,
	// but it's done anyway for consistency:
	bool allow_early_return;
	if (aSleepDuration == INTERVAL_UNSPECIFIED)
	{
		aSleepDuration = SLEEP_INTERVAL;  // Set interval to be the default length.
		allow_early_return = true;
	}
	else
		// The timer resolution makes waiting for half or less of an
		// interval too chancy.  The correct thing to do on average
		// is some kind of rounding, which this helps with:
		allow_early_return = (aSleepDuration <= SLEEP_INTERVAL_HALF);

	// Record the start time when the caller first called us so we can keep
	// track of how much time remains to sleep (in case the caller's subroutine
	// is suspended until a new subroutine is finished).  But for small sleep
	// intervals, don't worry about it.
	// Note: QueryPerformanceCounter() has very high overhead compared to GetTickCount():
	DWORD start_time = allow_early_return ? 0 : GetTickCount();

	// This check is also done even if the main timer will be set (below) so that
	// an initial check is done rather than waiting 10ms more for the first timer
	// message to come in.  Some of our many callers would want this, and although some
	// would not need it, there are so many callers that it seems best to just do it
	// unconditionally, especially since it's not a high overhead call (e.g. it returns
	// immediately if the tickcount is still the same as when it was last run).
	// Another reason for doing this check immediately is that our msg queue might
	// contains a time-consuming msg prior to our WM_TIMER msg, e.g. a hotkey msg.
	// In that case, the hotkey would be processed and launched without us first having
	// emptied the queue to discover the WM_TIMER msg.  In other words, WM_TIMER msgs
	// might get buried in the queue behind others, so doing this check here should help
	// ensure that timed subroutines are checked often enough to keep them running at
	// their specified frequencies.
	// Note that ExecUntil() no longer needs to call us solely for prevention of lag
	// caused by the keyboard & mouse hooks, so checking the timers early, rather than
	// immediately going into the GetMessage() state, should not be a problem:
	POLL_JOYSTICK_IF_NEEDED  // Do this first since it's much faster.
	bool return_value = false; //  Set default.  Also, this is used by the macro below.
	CHECK_SCRIPT_TIMERS_IF_NEEDED

	// Because this function is called recursively: for now, no attempt is
	// made to improve performance by setting the timer interval to be
	// aSleepDuration rather than a standard short interval.  That would cause
	// a problem if this instance of the function invoked a new subroutine,
	// suspending the one that called this instance.  The new subroutine
	// might need a timer of a longer interval, which would mess up
	// this layer.  One solution worth investigating is to give every
	// layer/instance its own timer (the ID of the timer can be determined
	// from info in the WM_TIMER message).  But that can be a real mess
	// because what if a deeper recursion level receives our first
	// WM_TIMER message because we were suspended too long?  Perhaps in
	// that case we wouldn't our WM_TIMER pulse because upon returning
	// from those deeper layers, we would check to see if the current
	// time is beyond our finish time.  In addition, having more timers
	// might be worse for overall system performance than having a single
	// timer that pulses very frequently (because the system must keep
	// them all up-to-date).  UPDATE: Timer is now also needed whenever an
	// aSleepDuration greater than 0 is about to be done and there are some
	// script timers that need to be watched (this happens when aMode == WAIT_FOR_MESSAGES).
	// UPDATE: Make this a macro so that it is dynamically resolved every time, in case
	// the value of g_script.mTimerEnabledCount changes on-the-fly.
	// UPDATE #2: The below has been changed in light of the fact that the main timer is
	// now kept always-on whenever there is at least one enabled timed subroutine.
	// This policy simplifies ExecUntil() and long-running commands such as FileSetAttrib.
	// UPDATE #3: Use aMode == RETURN_AFTER_MESSAGES, not g_nThreads > 0, because the
	// "Edit This Script" menu item (and possibly other places) might result in an indirect
	// call to us and we will need the timer to avoid getting stuck in the GetMessageState()
	// with hotkeys being disallowed due to filtering:
	bool this_layer_needs_timer = (aSleepDuration > 0 && aMode == RETURN_AFTER_MESSAGES);
	if (this_layer_needs_timer)
	{
		++g_nLayersNeedingTimer;  // IsCycleComplete() is responsible for decrementing this for us.
		SET_MAIN_TIMER
		// Reasons why the timer might already have been on:
		// 1) g_script.mTimerEnabledCount is greater than zero.
		// 2) another instance of MsgSleep() (beneath us in the stack) needs it (see the comments
		//    in IsCycleComplete() near KILL_MAIN_TIMER for details).
	}

	// Only used when aMode == RETURN_AFTER_MESSAGES:
	// True if the current subroutine was interrupted by another:
	//bool was_interrupted = false;
	bool sleep0_was_done = false;
	bool empty_the_queue_via_peek = false;

	int i, object_count;
	bool msg_was_handled;
	HWND fore_window, focused_control, focused_parent, criterion_found_hwnd;
	DWORD fore_pid;
	char fore_class_name[32];
	UserMenuItem *menu_item;
	Hotkey *hk;
	HotkeyVariant *variant;
	ActionTypeType type_of_first_line;
	int priority;
	Hotstring *hs;
	GuiType *pgui; // This is just a temp variable and should not be referred to once the below has been determined.
	GuiControlType *pcontrol, *ptab_control;
	GuiIndexType gui_index;  // Don't use pgui because pointer can become invalid if ExecUntil() executes "Gui Destroy".
	bool *pgui_label_is_running;
	Label *gui_label;
	RECT rect;
	DWORD event_info;
	HDROP hdrop_to_free;
	DWORD drop_count;
	DWORD tick_before, tick_after;
	LRESULT msg_reply;
	MSG msg;

	for (;;)
	{
		tick_before = GetTickCount();
		if (aSleepDuration > 0 && !empty_the_queue_via_peek)
		{
			// The following comment is mostly obsolete as of v1.0.39 (which introduces a thread
			// dedicated to the hooks).  However, using GetMessage() is still superior to
			// PeekMessage() for performance reason.  Add to that the risk of breaking things
			// and it seems clear that it's best to retain GetMessage().
			// Older comment:
			// Use GetMessage() whenever possible -- rather than PeekMessage() or a technique such
			// MsgWaitForMultipleObjects() -- because it's the "engine" that passes all keyboard
			// and mouse events immediately to the low-level keyboard and mouse hooks
			// (if they're installed).  Otherwise, there's greater risk of keyboard/mouse lag.
			// PeekMessage(), depending on how, and how often it's called, will also do this, but
			// I'm not as confident in it.
			if (GetMessage(&msg, NULL, 0, MSG_FILTER_MAX) == -1) // -1 is an error, 0 means WM_QUIT
				continue; // Error probably happens only when bad parameters were passed to GetMessage().
			//else let any WM_QUIT be handled below.
			// The below was added for v1.0.20 to solve the following issue: If BatchLines is 10ms
			// (its default) and there are one or more 10ms script-timers active, those timers would
			// actually only run about every 20ms.  In addition to solving that problem, the below
			// might also improve reponsiveness of hotkeys, menus, buttons, etc. when the CPU is
			// under heavy load:
			tick_after = GetTickCount();
			if (tick_after - tick_before > 3)  // 3 is somewhat arbitrary, just want to make sure it rested for a meaningful amount of time.
				g_script.mLastScriptRest = tick_after;
		}
		else // aSleepDuration < 1 || empty_the_queue_via_peek
		{
			// In the above cases, we don't want to be stuck in GetMessage() for even 10ms:
			if (!PeekMessage(&msg, NULL, 0, MSG_FILTER_MAX, PM_REMOVE)) // No more messages
			{
				// Since the Peek() didn't find any messages, our timeslice was probably just
				// yielded if the CPU is under heavy load.  If so, it seems best to count that as
				// a "rest" so that 10ms script-timers will run closer to the desired frequency
				// (see above comment for more details).
				// These next few lines exact match the ones above, so keep them in sync:
				tick_after = GetTickCount();
				if (tick_after - tick_before > 3)
					g_script.mLastScriptRest = tick_after;
				// It is not necessary to actually do the Sleep(0) when aSleepDuration == 0
				// because the most recent PeekMessage() has just yielded our prior timeslice.
				// This is because when Peek() doesn't find any messages, it automatically
				// behaves as though it did a Sleep(0). UPDATE: This is apparently not quite
				// true.  Although Peek() does yield, it is somehow not as long or as good as
				// Sleep(0).  This is evidenced by the fact that some of my script's
				// WinWaitClose's now finish too quickly when DoKeyDelay(0) is done for them,
				// but replacing DoKeyDelay(0) with Sleep(0) makes it work as it did before.
				if (aSleepDuration == 0 && !sleep0_was_done)
				{
					Sleep(0);
					sleep0_was_done = true;
					// Now start a new iteration of the loop that will see if we
					// received any messages during the up-to-20ms delay (perhaps even more)
					// that just occurred.  It's done this way to minimize keyboard/mouse
					// lag (if the hooks are installed) that will occur if any key or
					// mouse events are generated during that 20ms.  Note: It seems that
					// the OS knows not to yield our timeslice twice in a row: once for
					// the Sleep(0) above and once for the upcoming PeekMessage() (if that
					// PeekMessage() finds no messages), so it does not seem necessary
					// to check HIWORD(GetQueueStatus(QS_ALLEVENTS)).  This has been confirmed
					// via the following test, which shows that while BurnK6 (CPU maxing program)
					// is foreground, a Sleep(0) really does a Sleep(60).  But when it's not
					// foreground, it only does a Sleep(20).  This behavior is UNAFFECTED by
					// the added presence of of a HIWORD(GetQueueStatus(QS_ALLEVENTS)) check here:
					//SplashTextOn,,, xxx
					//WinWait, xxx  ; set last found window
					//Loop
					//{
					//	start = %a_tickcount%
					//	Sleep, 0
					//	elapsed = %a_tickcount%
					//	elapsed -= %start%
					//	WinSetTitle, %elapsed%
					//}
					continue;
				}
				// Otherwise: aSleepDuration is non-zero or we already did the Sleep(0)
				// Macro notes:
				// Must decrement prior to every RETURN to balance it.
				// Do this prior to checking whether timer should be killed, below.
				// Kill the timer only if we're about to return OK to the caller since the caller
				// would still need the timer if FAIL was returned above.  But don't kill it if
				// there are any enabled timed subroutines, because the current policy it to keep
				// the main timer always-on in those cases.  UPDATE: Also avoid killing the timer
				// if there are any script threads running.  To do so might cause a problem such
				// as in this example scenario: MsgSleep() is called for any reason with a delay
				// large enough to require the timer.  The timer is set.  But a msg arrives that
				// MsgSleep() dispatches to MainWindowProc().  If it's a hotkey or custom menu,
				// MsgSleep() is called recursively with a delay of -1.  But when it finishes via
				// IsCycleComplete(), the timer would be wrongly killed because the underlying
				// instance of MsgSleep still needs it.  Above is even more wide-spread because if
				// MsgSleep() is called recursively for any reason, even with a duration >10, it will
				// wrongly kill the timer upon returning, in some cases.  For example, if the first call to
				// MsgSleep(-1) finds a hotkey or menu item msg, and executes the corresponding subroutine,
				// that subroutine could easily call MsgSleep(10+) for any number of reasons, which
				// would then kill the timer.
				// Also require that aSleepDuration > 0 so that MainWindowProc()'s receipt of a
				// WM_HOTKEY msg, to which it responds by turning on the main timer if the script
				// is uninterruptible, is not defeated here.  In other words, leave the timer on so
				// that when the script becomes interruptible once again, the hotkey will take effect
				// almost immediately rather than having to wait for the displayed dialog to be
				// dismissed (if there is one).
				// If timer doesn't exist, the new-thread-launch routine of MsgSleep() relies on this
				// to turn it back on whenever a layer beneath us needs it.  Since the timer
				// is never killed while g_script.mTimerEnabledCount is >0, it shouldn't be necessary
				// to check g_script.mTimerEnabledCount here.
				#define RETURN_FROM_MSGSLEEP \
				{\
					if (this_layer_needs_timer)\
						--g_nLayersNeedingTimer;\
					if (g_MainTimerExists)\
					{\
						if (aSleepDuration > 0 && !g_nLayersNeedingTimer && !g_script.mTimerEnabledCount && !Hotkey::sJoyHotkeyCount)\
							KILL_MAIN_TIMER \
					}\
					else if (g_nLayersNeedingTimer)\
						SET_MAIN_TIMER \
					return return_value;\
				}
				// Function should always return OK in this case.  Also, was_interrupted
				// will always be false because if this "aSleepDuration < 1" call
				// really was interrupted, it would already have returned in the
				// hotkey cases of the switch().  UPDATE: was_interrupted can now
				// be true since the hotkey case in the switch() doesn't return,
				// relying on us to do it after making sure the queue is empty.
				// The below is checked here rather than in IsCycleComplete() because
				// that function is sometimes called more than once prior to returning
				// (e.g. empty_the_queue_via_peek) and we only want this to be decremented once:
				IsCycleComplete(aSleepDuration, start_time, allow_early_return);
				RETURN_FROM_MSGSLEEP
			}
			// else Peek() found a message, so process it below.
		} // PeekMessage() vs. GetMessage()

		// Since above didn't return or "continue", a message has been received that is eligible
		// for further processing.

		// For max. flexibility, it seems best to allow the message filter to have the first
		// crack at looking at the message, before even TRANSLATE_AHK_MSG:
		if (g_MsgMonitorCount && MsgMonitor(msg.hwnd, msg.message, msg.wParam, msg.lParam, &msg, msg_reply))  // Count is checked here to avoid function-call overhead.
			continue; // MsgMonitor has returned "true", indicating that this message should be omitted from further processing.
			// Above does "continue" and ignores msg_reply.  This is because testing shows that
			// messages received via Get/PeekMessage() were always sent via PostMessage.  If an
			// another thread sends ours a message, MSDN implies that Get/PeekMessage() internally
			// calls the message's WindowProc directly and sends the reply back to the other thread.
			// That makes sense because it seems unlikely that DispatchMessage contains any means
			// of replying to a message because it has no way of knowing whether the MSG struct
			// arrived via Post vs. SendMessage.

		// If this message might be for one of our GUI windows, check that before doing anything
		// else with the message.  This must be done first because some of the standard controls
		// also use WM_USER messages, so we must not assume they're generic thread messages just
		// because they're >= WM_USER.  The exception is AHK_GUI_ACTION should always be handled
		// here rather than by IsDialogMessage().  Note: sObjectCount is checked first to help
		// performance, since all messages must come through this bottleneck.
		if (GuiType::sObjectCount && msg.hwnd && msg.hwnd != g_hWnd && !(msg.message == AHK_GUI_ACTION
			|| msg.message == AHK_USER_MENU || msg.message == AHK_CLIPBOARD_CHANGE))
		{
			if (msg.message == WM_KEYDOWN)
			{
				// Relies heavily on short-circuit boolean order:
				if (  (msg.wParam == VK_NEXT || msg.wParam == VK_PRIOR || msg.wParam == VK_TAB
					|| msg.wParam == VK_LEFT || msg.wParam == VK_RIGHT)
					&& (focused_control = GetFocus()) && (focused_parent = GetNonChildParent(focused_control))
					&& (pgui = GuiType::FindGui(focused_parent)) && pgui->mTabControlCount
					&& (pcontrol = pgui->FindControl(focused_control)) && pcontrol->type != GUI_CONTROL_HOTKEY   )
				{
					ptab_control = NULL; // Set default.
					if (pcontrol->type == GUI_CONTROL_TAB) // The focused control is a tab control itself.
					{
						ptab_control = pcontrol;
						// For the below, note that Alt-left and Alt-right are automatically excluded,
						// as desired, since any key modified only by alt would be WM_SYSKEYDOWN vs. WM_KEYDOWN.
						if (msg.wParam == VK_LEFT || msg.wParam == VK_RIGHT)
						{
							pgui->SelectAdjacentTab(*ptab_control, msg.wParam == VK_RIGHT, false, false);
							// Pass false for both the above since that's the whole point of having arrow
							// keys handled separately from the below: Focus should stay on the tabs
							// rather than jumping to the first control of the tab, it focus should not
							// wrap around to the beginning or end (to conform to standard behavior for
							// arrow keys).
							continue; // Suppress this key even if the above failed (probably impossible in this case).
						}
						//else fall through to the next part.
					}
					// If focus is in a multiline edit control, don't act upon Control-Tab (and
					// shift-control-tab -> for simplicity & consistency) since Control-Tab is a special
					// keystroke that inserts a literal tab in the edit control:
					if (   msg.wParam != VK_LEFT && msg.wParam != VK_RIGHT
						&& (GetKeyState(VK_CONTROL) & 0x8000) // Even if other modifiers are down, it still qualifies. Use GetKeyState() vs. GetAsyncKeyState() because the former's definition is more suitable.
						&& (msg.wParam != VK_TAB || pcontrol->type != GUI_CONTROL_EDIT
							|| !(GetWindowLong(pcontrol->hwnd, GWL_STYLE) & ES_MULTILINE))   )
					{
						// If ptab_control wasn't determined above, check if focused control is owned by a tab control:
						if (!ptab_control && !(ptab_control = pgui->FindTabControl(pcontrol->tab_control_index))   )
							// Fall back to the first tab control (for consistency & simplicty, seems best
							// to always use the first rather than something fancier such as "nearest in z-order".
							ptab_control = pgui->FindTabControl(0);
						if (ptab_control)
						{
							pgui->SelectAdjacentTab(*ptab_control
								, msg.wParam == VK_NEXT || (msg.wParam == VK_TAB && !(GetKeyState(VK_SHIFT) & 0x8000)) // Use GetKeyState() vs. GetAsyncKeyState() because the former's definition is more suitable.
								, true, true);
							// Update to the below: Must suppress the tab key at least, to prevent it
							// from navigating *and* changing the tab.  And since this one is suppressed,
							// might as well suppress the others for consistency.
							// Older: Since WM_KEYUP is not handled/suppressed here, it seems best not to
							// suppress this WM_KEYDOWN either (it should do nothing in this case
							// anyway, but for balance this seems best): Fall through to the next section.
							continue;
						}
						//else fall through to the below.
					}
					//else fall through to the below.
				} // Interception of keystrokes for navigation in tab control.

				// v1.0.34: Fix for the fact that a multiline edit control will send WM_CLOSE to its parent
				// when user presses ESC while it has focus.  The following check is similar to the block's above.
				// The alternative to this approach would have been to override the edit control's WindowProc,
				// but the following seemed to be less code. Although this fix is only necessary for multiline
				// edits, its done for all edits since it doesn't do any harm.  In addition, there is no need to
				// check what modifiers are down because we never receive the keystroke for Ctrl-Esc and Alt-Esc
				// (the OS handles those beforehand) and both Win-Esc and Shift-Esc are identical to a naked Esc
				// inside an edit.  The following check relies heavily on short-circuit eval. order.
				if (   (msg.wParam == VK_ESCAPE || msg.wParam == VK_TAB) // v1.0.38.03: Added VK_TAB handling for "WantTab".
					&& (focused_control = GetFocus()) && (focused_parent = GetNonChildParent(focused_control))
					&& (pgui = GuiType::FindGui(focused_parent)) && (pcontrol = pgui->FindControl(focused_control))
					&& pcontrol->type == GUI_CONTROL_EDIT)
				{
					if (msg.wParam == VK_ESCAPE)
					{
						pgui->Escape();
						continue; // Omit this keystroke from any further processing.
					}
					else // msg.wParam == VK_TAB
						if (pcontrol->attrib & GUI_CONTROL_ATTRIB_ALTBEHAVIOR) // It has the "WantTab" property.
						{
							// For flexibility, do this even for single-line edit controls, though in that
							// case the tab keystroke will produce an "empty box" character.
							// Strangely, if a message pump other than this one (MsgSleep) is running,
							// such as that of a MsgBox, "WantTab" is already in effect unconditionally,
							// perhaps because MsgBox and others respond to WM_GETDLGCODE with DLGC_WANTTAB.
							SendMessage(pcontrol->hwnd, EM_REPLACESEL, TRUE, (LPARAM)"\t");
							continue; // Omit this keystroke from any further processing.
						}
				}
			} // if (msg.message == WM_KEYDOWN)

			for (i = 0, object_count = 0, msg_was_handled = false; i < MAX_GUI_WINDOWS; ++i)
			{
				// Note: indications are that IsDialogMessage() should not be called with NULL as
				// its first parameter (perhaps as an attempt to get allow dialogs owned by our
				// thread to be handled at once). Although it might work on some versions of Windows,
				// it's undocumented and shouldn't be relied on.
				// Also, can't call IsDialogMessage against msg.hwnd because that is not a complete
				// solution: at the very least, tab key navigation will not work in GUI windows.
				// There are probably other side-effects as well.
				if (g_gui[i])
				{
					if (g_gui[i]->mHwnd)
					{
						g.CalledByIsDialogMessageOrDispatch = true;
						if (IsDialogMessage(g_gui[i]->mHwnd, &msg))
						{
							msg_was_handled = true;
							g.CalledByIsDialogMessageOrDispatch = false;
							break;
						}
						g.CalledByIsDialogMessageOrDispatch = false;
					}
					if (GuiType::sObjectCount == ++object_count) // No need to keep searching.
						break;
				}
			}
			if (msg_was_handled) // This message was handled by IsDialogMessage() above.
				continue; // Continue with the main message loop.
		}

		TRANSLATE_AHK_MSG(msg.message, msg.wParam)

		switch(msg.message)
		{
		case WM_TIMER:
			if (msg.lParam) // Since this WM_TIMER is intended for a TimerProc, dispatch the msg instead.
				break;
			// It seems best to poll the joystick for every WM_TIMER message (i.e. every 10ms or so on
			// NT/2k/XP).  This is because if the system is under load, it might be 20ms, 40ms, or even
			// longer before we get a timeslice again and that is a long time to be away from the poll
			// (a fast button press-and-release might occur in less than 50ms, which could be missed if
			// the polling frequency is too low):
			POLL_JOYSTICK_IF_NEEDED // Do this first since it's much faster.
			// v1.0.38.04: The following line is done prior to the timer-check to reduce situations
			// in which a timer thread is interrupted before it can execute even a single line.
			// Search for mLastPeekTime in MsgSleep() for detailed explanation.
			g_script.mLastPeekTime = GetTickCount(); // It's valid to reset this because by definition, "msg" just came in via Get() or Peek(), both of which qualify as a Peek() for this purpose.
			CHECK_SCRIPT_TIMERS_IF_NEEDED
			if (aMode == WAIT_FOR_MESSAGES)
				// Timer should have already been killed if we're in this state.
				// But there might have been some WM_TIMER msgs already in the queue
				// (they aren't removed when the timer is killed).  Or perhaps
				// a caller is calling us with this aMode even though there
				// are suspended subroutines (currently never happens).
				// In any case, these are always ignored in this mode because
				// the caller doesn't want us to ever return.  UPDATE: This can now
				// happen if there are any enabled timed subroutines we need to keep an
				// eye on, which is why the mTimerEnabledCount value is checked above
				// prior to starting a new iteration.
				continue;
			if (aSleepDuration < 1) // In this case, WM_TIMER messages have already fulfilled their function, above.
				continue;
			// Otherwise aMode == RETURN_AFTER_MESSAGES:
			// Realistically, there shouldn't be any more messages in our queue
			// right now because the queue was stripped of WM_TIMER messages
			// prior to the start of the loop, which means this WM_TIMER
			// came in after the loop started.  So in the vast majority of
			// cases, the loop would have had enough time to empty the queue
			// prior to this message being received.  Therefore, just return rather
			// than trying to do one final iteration in peek-mode (which might
			// complicate things, i.e. the below function makes certain changes
			// in preparation for ending this instance/layer, only to be possibly,
			// but extremely rarely, interrupted/recursed yet again if that final
			// peek were to detect a recursable message):
			if (IsCycleComplete(aSleepDuration, start_time, allow_early_return))
				RETURN_FROM_MSGSLEEP
			// Otherwise, stay in the blessed GetMessage() state until the time has expired:
			continue;

		case AHK_GUI_ACTION:   // The user pressed a button on a GUI window, or some other actionable event. Listed first for performance.
		case WM_HOTKEY:        // As a result of this app having previously called RegisterHotkey(), or from TriggerJoyHotkeys().
		case AHK_HOOK_HOTKEY:  // Sent from this app's keyboard or mouse hook.
		case AHK_HOTSTRING:    // Sent from keybd hook to activate a non-auto-replace hotstring.
		case AHK_USER_MENU:    // The user selected a custom menu item.
		case AHK_CLIPBOARD_CHANGE: // Listed so that hdrop_to_free is initialized.
			// MSG_FILTER_MAX should prevent us from receiving these messages (except AHK_USER_MENU)
			// whenever g_AllowInterruption or g.AllowThreadToBeInterrupted is false.
			hdrop_to_free = NULL;  // Set default for this message's processing (simplifies code).
			event_info = UINT_MAX; // Set default for all of the above event types.  A flag that indicates "no event info".
			switch(msg.message)
			{
			case AHK_GUI_ACTION: // Listed first for performance.
				// Assume that it is possible that this message's GUI window has been destroyed
				// (and maybe even recreated) since the time the msg was posted.  If this can happen,
				// that's another reason for finding which GUI this control is associate with (it also
				// needs to be found so that we can call the correct GUI window object to perform
				// the action):
				if (   !(pgui = GuiType::FindGui(msg.hwnd))   ) // No associated GUI object, so ignore this event.
					continue;
				gui_index = pgui->mWindowIndex;  // Needed in case ExecUntil() performs "Gui Destroy" further below.
				pgui_label_is_running = NULL; // Set default (in cases other than AHK_GUI_ACTION it is not used, so not initialized).
				switch(msg.wParam)
				{
				case AHK_GUI_SIZE: // This is the signal to run the window's OnEscape label. Listed first for performance.
					if (   !(gui_label = pgui->mLabelForSize)   ) // In case it became NULL since the msg was posted.
						continue;
					pgui_label_is_running = &pgui->mLabelForSizeIsRunning;
					break;
				case AHK_GUI_CLOSE:  // This is the signal to run the window's OnClose label.
					if (   !(gui_label = pgui->mLabelForClose)   ) // In case it became NULL since the msg was posted.
						continue;
					pgui_label_is_running = &pgui->mLabelForCloseIsRunning;
					break;
				case AHK_GUI_ESCAPE: // This is the signal to run the window's OnEscape label.
					if (   !(gui_label = pgui->mLabelForEscape)   ) // In case it became NULL since the msg was posted.
						continue;
					pgui_label_is_running = &pgui->mLabelForEscapeIsRunning;
					break;
				case AHK_GUI_CONTEXTMENU:
					if (   !(gui_label = pgui->mLabelForContextMenu)   ) // In case it became NULL since the msg was posted.
						continue;
					// UPDATE: Must allow multiple threads because otherwise the user cannot right-click twice
					// consecutively (the second click is blocked because the menu is still displayed at the
					// instant of the click.  The following comment is probably not entirely correct because
					// the display of a popup menu via "Menu, MyMenu, Show" will spin off a new thread if the
					// user selects an item in the menu:
					// Unlike most other Gui labels, it seems best by default to allow GuiContextMenu to be
					// launched multiple times so that multiple items in the menu can be running simultaneously
					// as separate threads.  Therefore, leave pgui_label_is_running at its default of NULL.
					break;
				case AHK_GUI_DROPFILES: // This is the signal to run the window's DropFiles label.
					hdrop_to_free = pgui->mHdrop; // This variable simplifies the code further below.
					if (   !(gui_label = pgui->mLabelForDropFiles) // In case it became NULL since the msg was posted.
						|| !hdrop_to_free // Checked just in case, so that the below can query it.
						|| !(drop_count = DragQueryFile(hdrop_to_free, 0xFFFFFFFF, NULL, 0))   ) // Probably impossible, but if it ever can happen, seems best to ignore it.
					{
						if (hdrop_to_free) // Checked again in case short-circuit boolean above never checked it.
						{
							DragFinish(hdrop_to_free); // Since the drop-thread will not be launched, free the memory.
							pgui->mHdrop = NULL; // Indicate that this GUI window is ready for another drop.
						}
						continue;
					}
					// It is not necessary to check if the label is running in this case because
					// the caller who posted this message to us has ensured that it isn't running.
					// Therefore, leave pgui_label_is_running at its default of NULL.
					break;
				default: // This is an action from a particular control in the GUI window.
					event_info = HIWORD(msg.wParam); // Assigned as an unsigned value to preserve the USHORT's full unsigned capacity.
					msg.wParam = LOWORD(msg.wParam); // Remove the extra-info word from further consideration below.
					if (msg.wParam >= pgui->mControlCount) // Index beyond the quantity of controls, so ignore this event.
						continue; // Discarding an invalid message here is relied upon both other sections below.
					pcontrol = &pgui->mControl[msg.wParam]; // Set for use here and in other places below.
					if (   !(gui_label = pcontrol->jump_to_label)   )
					{
						// On if there's no label is the implicit action considered.
						if (pcontrol->attrib & GUI_CONTROL_ATTRIB_IMPLICIT_CANCEL)
							pgui->Cancel();
						continue; // Fully handled by the above; or there was no label.
						// This event might lack both a label and an action if its control was changed to be
						// non-actionable since the time the msg was posted.
					}
					// And leave pgui_label_is_running at its default of NULL, which signals below to use
					// alternate variable.
				} // switch(msg.wParam)
				type_of_first_line = gui_label->mJumpToLine->mActionType; // Above would already have discarded this message if it there was no label.
				break; // case AHK_GUI_ACTION

			case AHK_USER_MENU: // user-defined menu item
				if (   !(menu_item = g_script.FindMenuItemByID((UINT)msg.lParam))   ) // Item not found.
					continue; // ignore the msg
				// And just in case a menu item that lacks a label (such as a separator) is ever
				// somehow selected (perhaps via someone sending direct messages to us, bypassing
				// the menu):
				if (!menu_item->mLabel)
					continue;
				type_of_first_line = menu_item->mLabel->mJumpToLine->mActionType;
				break;

			case AHK_HOTSTRING:
				if (msg.wParam >= Hotstring::sHotstringCount) // Invalid hotstring ID (perhaps spoofed by external app)
					continue; // Do nothing.
				hs = Hotstring::shs[msg.wParam];  // For performance and convenience.
				if (hs->mHotCriterion)
				{
					// For details, see comments in the hotkey section of this switch().
					if (   !(criterion_found_hwnd = HotCriterionAllowsFiring(hs->mHotCriterion, hs->mHotWinTitle, hs->mHotWinText))   )
						// Hotstring is no longer eligible to fire even though it was when the hook sent us
						// the message.  Abort the firing even though the hook may have already started
						// executing the hotstring by suppressing the final end-character or other actions.
						// It seems preferable to abort midway through the execution than to continue sending
						// keystrokes to the wrong window, or when the hotstring has become suspended.
						continue;
					// For details, see comments in the hotkey section of this switch().
					if (!(hs->mHotCriterion == HOT_IF_ACTIVE || hs->mHotCriterion == HOT_IF_EXIST))
						criterion_found_hwnd = NULL; // For "NONE" and "NOT", there is no last found window.
				}
				else // No criterion, so it's a global hotstring.  It can always fire, but it has no "last found window".
					criterion_found_hwnd = NULL;
				// Do a simple replacement for the hotstring if that's all that's called for.
				// Don't create a new quasi-thread or any of that other complexity done further
				// below.  But also do the backspacing (if specified) for a non-autoreplace hotstring,
				// even if it can't launch due to MaxThreads, MaxThreadsPerHotkey, or some other reason:
				hs->DoReplace(msg.lParam);  // Does only the backspacing if it's not an auto-replace hotstring.
				if (*hs->mReplacement) // Fully handled by the above; i.e. it's an auto-replace hotstring.
					continue;
				// Otherwise, continue on and let a new thread be created to handle this hotstring.
				// Since this isn't an auto-replace hotstring, set this value to support
				// the built-in variable A_EndChar:
				g_script.mEndChar = (char)LOWORD(msg.lParam);
				type_of_first_line = hs->mJumpToLabel->mJumpToLine->mActionType;
				break;

			case AHK_CLIPBOARD_CHANGE: // Due to the presence of an OnClipboardChange label in the script.
				// Caller has ensured that mOnClipboardChangeLabel is a non-NULL, valid pointer.
				type_of_first_line = g_script.mOnClipboardChangeLabel->mJumpToLine->mActionType;
				break;

			default: // hotkey
				if (msg.wParam >= Hotkey::sHotkeyCount) // Invalid hotkey ID.
					continue;
				hk = Hotkey::shk[msg.wParam];
				// Check if criterion allows firing.
				// For maintainability, this is done here rather than a little further down
				// past the MAX_THREADS_LIMIT and thread-priority checks.  Those checks hardly
				// ever abort a hotkey launch anyway.
				//
				// If message is WM_HOTKEY, it's either:
				// 1) A joystick hotkey from TriggerJoyHotkeys(), in which case the lParam is ignored.
				// 2) A hotkey message sent by the OS, in which case lParam contains currently-unused info set by the OS.
				//
				// An incoming WM_HOTKEY can be subject to #IfWin at this stage under the following conditions:
				// 1) Joystick hotkey, because it relies on us to do the check so that the check is done only
				//    once rather than twice.
				// 2) Win9x's support for #IfWin, which never uses the hook but instead simply does nothing if
				//    none of the hotkey's criteria is satisfied.
				// 3) #IfWin keybd hotkeys that were made non-hook because they have a non-suspended, global variant.
				//
				// If message is AHK_HOOK_HOTKEY:
				// Rather than having the hook pass the qualified variant to us, it seems preferable
				// to search through all the criteria again and rediscover it.  This is because conditions
				// may have changed since the message was posted, and although the hotkey might still be
				// eligible for firing, a different variant might now be called for (e.g. due to a change
				// in the active window).  Since most criteria hotkeys have at most only a few criteria,
				// and since most such criteria are #IfWinActive rather than Exist, the performance will
				// typically not be reduced much at all.  Futhermore, trading performance for greater
				// reliability seems worth it in this case.
				// 
				// The inefficiency of calling HotCriterionAllowsFiring() twice for each hotkey --
				// once in the hook and again here -- seems justifed for the following reasons:
				// - It only happens twice if the hotkey a hook hotkey (multi-variant keyboard hotkeys
				//   that have a global variant are usually non-hook, even on NT/2k/XP).
				// - The hook avoids doing its first check of WinActive/Exist if it sees that the hotkey
				//   has a non-suspended, global variant.  That way, hotkeys that are hook-hotkeys for
				//   reasons other than #IfWin (such as mouse, overriding OS hotkeys, or hotkeys
				//   that are too fancy for RegisterHotkey) will not have to do the check twice.
				// - It provides the ability to set the last-found-window for #IfWinActive/Exist
				//   (though it's not needed for the "Not" counterparts).  This HWND could be passed
				//   via the message, but that would require malloc-there and free-here, and might
				//   result in memory leaks if its ever possible for messages to get discarded by the OS.
				// - It allows hotkeys that were eligible for firing at the time the message was
				//   posted but that have since become ineligible to be aborted.  This seems like a
				//   good precaution for most users/situations because such hotkey subroutines will
				//   often assume (for scripting simplicity) that the specified window is active or
				//   exists when the subroutine executes its first line.
				// - Most criterion hotkeys use #IfWinActive, which is a very fast call.  Also, although
				//   WinText and/or "SetTitleMatchMode Slow" slow down window searches, those are rarely
				//   used too.
				if (   !(variant = hk->CriterionAllowsFiring(&criterion_found_hwnd))   )
					continue; // No criterion is eligible, so ignore this hotkey event (see other comments).
					// If this is AHK_HOOK_HOTKEY, criterion was eligible at time message was posted,
					// but not now.  Seems best to abort (see other comments).
				// Now that above has ensured variant is non-NULL:
				if (!(variant->mHotCriterion == HOT_IF_ACTIVE || variant->mHotCriterion == HOT_IF_EXIST))
					criterion_found_hwnd = NULL; // For "NONE" and "NOT", there is no last found window.
				type_of_first_line = variant->mJumpToLabel->mJumpToLine->mActionType;
			} // switch(msg.message)

			if (g_nThreads >= g_MaxThreadsTotal)
			{
				// The below allows 1 thread beyond the limit in case the script's configured
				// #MaxThreads is exactly equal to the absolute limit.  This is because we want
				// subroutines whose first line is something like ExitApp to take effect even
				// when we're at the absolute limit:
				if (g_nThreads > MAX_THREADS_LIMIT || !ACT_IS_ALWAYS_ALLOWED(type_of_first_line))
				{
					// Allow only a limited number of recursion levels to avoid any chance of
					// stack overflow.  So ignore this message.  Later, can devise some way
					// to support "queuing up" these launch-thread events for use later when
					// there is "room" to run them, but that might cause complications because
					// in some cases, the user didn't intend to hit the key twice (e.g. due to
					// "fat fingers") and would have preferred to have it ignored.  Doing such
					// might also make "infinite key loops" harder to catch because the rate
					// of incoming hotkeys would be slowed down to prevent the subroutines from
					// running concurrently.
					if (hdrop_to_free) // This is only non-NULL when pgui is non-NULL.
					{
						DragFinish(hdrop_to_free); // Since the drop-thread will not be launched, free the memory.
						pgui->mHdrop = NULL; // Indicate that this GUI window is ready for another drop.
					}
					continue;
				}
				// If the above "continued", it seems best not to re-queue/buffer the key since
				// it might be a while before the number of threads drops back below the limit.
			}

			switch(msg.message)
			{
			case AHK_GUI_ACTION: // Listed first for performance.
				// It seems best by default not to allow multiple threads for the same control.
				// Such events are discarded because it seems likely that most script designers
				// would want to see the effects of faulty design (e.g. long running timers or
				// hotkeys that interrupt gui threads) rather than having events for later,
				// when they might suddenly take effect unexpectedly:
				if (pgui_label_is_running)
				{
					if (*pgui_label_is_running)
						continue;
				}
				else if (msg.wParam != AHK_GUI_DROPFILES && msg.wParam != AHK_GUI_CONTEXTMENU) // It's a control's label (the above IF catches AHK_GUI_CLOSE and related).
					if (pcontrol->attrib & GUI_CONTROL_ATTRIB_LABEL_IS_RUNNING) // pcontrol was set higher above for control-specific events.
						continue;
					//else this control's label is eligible to run now.
				//else the check wasn't needed because it was done elsewhere (AHK_GUI_DROPFILES) or the
				// action is not thread-restricted (AHK_GUI_CONTEXTMENU).
				priority = 0;  // Always use default for now.
				break;
			case AHK_USER_MENU: // user-defined menu item
				// Ignore/discard a hotkey or custom menu item event if the current thread's priority
				// is higher than it's:
				priority = menu_item->mPriority;
				break;
			case AHK_HOTSTRING:
				priority = hs->mPriority;
				break;
			case AHK_CLIPBOARD_CHANGE: // Due to the presence of an OnClipboardChange label in the script.
				if (g_script.mOnClipboardChangeIsRunning)
					continue;
				priority = 0;  // Always use default for now.
				break;
			default: // hotkey
				// Due to the key-repeat feature and the fact that most scripts use a value of 1
				// for their #MaxThreadsPerHotkey, this check will often help average performance
				// by avoiding a lot of unncessary overhead that would otherwise occur:
				if (!hk->PerformIsAllowed(*variant))
				{
					// The key is buffered in this case to boost the responsiveness of hotkeys
					// that are being held down by the user to activate the keyboard's key-repeat
					// feature.  This way, there will always be one extra event waiting in the queue,
					// which will be fired almost the instant the previous iteration of the subroutine
					// finishes (this above descript applies only when MaxThreadsPerHotkey is 1,
					// which it usually is).
					hk->RunAgainAfterFinished(*variant);
					continue;
				}
				priority = variant->mPriority;
			}

			if (priority < g.Priority) // Ignore this event because its priority is too low.
			{
				if (hdrop_to_free) // This is only non-NULL when pgui is non-NULL.
				{
					DragFinish(hdrop_to_free); // Since the drop-thread will not be launched, free the memory.
					pgui->mHdrop = NULL; // Indicate that this GUI window is ready for another drop.
				}
				continue;
			}

			// Always kill the main timer, for performance reasons and for simplicity of design,
			// prior to embarking on new subroutine whose duration may be long (e.g. if BatchLines
			// is very high or infinite, the called subroutine may not return to us for seconds,
			// minutes, or more; during which time we don't want the timer running because it will
			// only fill up the queue with WM_TIMER messages and thus hurt performance).
			// UPDATE: But don't kill it if it should be always-on to support the existence of
			// at least one enabled timed subroutine or joystick hotkey:
			if (!g_script.mTimerEnabledCount && !Hotkey::sJoyHotkeyCount)
				KILL_MAIN_TIMER;

			if (aMode == RETURN_AFTER_MESSAGES)
			{
				// Assert: g_nThreads should be greater than 0 in this mode, which means
				// that there must be another thread besides the one we're about to create.
				// That thread will be interrupted and suspended to allow this new one to run.
				//was_interrupted = true;
				// Save the current foreground window in case the subroutine that's about
				// to be suspended is working with it.  Then, when the subroutine is
				// resumed, we can ensure this window is the foreground one.  UPDATE:
				// this has been disabled because it often is the incorrect thing to do
				// (e.g. if the suspended hotkey wasn't working with the window, but is
				// a long-running subroutine, hotkeys that activate windows will have
				// those windows deactivated instantly when their subroutine is over,
				// since the suspended subroutine resumes and would reassert its foreground
				// window:
				//g.hWndToRestore = aRestoreActiveWindow ? GetForegroundWindow() : NULL;

				// Also save the ErrorLevel of the subroutine that's about to be suspended.
				// Current limitation: If the user put something big in ErrorLevel (very unlikely
				// given its nature, but allowed) it will be truncated by this, if too large.
				// Also: Don't use var->Get() because need better control over the size:
				strlcpy(g.ErrorLevel, g_ErrorLevel->Contents(), sizeof(g.ErrorLevel));
				// Could also use copy constructor but that would probably incur more overhead?:
				// Next, change the value of globals to reflect the fact that we're about
				// to launch a new subroutine.
			}
			// v1.0.37.06: The following must ben done regardless of aMode because the idle thread is now
			// resumed via ResumeUnderlyingThread():
			CopyMemory(&global_saved, &g, sizeof(global_struct));

			switch(msg.message)
			{
			case AHK_GUI_ACTION: // Listed first for performance.
			case AHK_CLIPBOARD_CHANGE:
				break; // Do nothing at this stage.
			case AHK_USER_MENU: // user-defined menu item
				// Safer to make a full copies than point to something potentially volatile.
				strlcpy(g_script.mThisMenuItemName, menu_item->mName, sizeof(g_script.mThisMenuItemName));
				strlcpy(g_script.mThisMenuName, menu_item->mMenu->mName, sizeof(g_script.mThisMenuName));
				break;
			default: // hotkey or hotstring
				// Just prior to launching the hotkey, update these values to support built-in
				// variables such as A_TimeSincePriorHotkey:
				g_script.mPriorHotkeyName = g_script.mThisHotkeyName;
				g_script.mPriorHotkeyStartTime = g_script.mThisHotkeyStartTime;
				// Unlike hotkeys -- which can have a name independent of their label by being created or updated
				// with the HOTKEY command -- a hot string's unique name is always its label since that includes
				// the options that distinguish between (for example) :c:ahk:: and ::ahk::
				g_script.mThisHotkeyName = (msg.message == AHK_HOTSTRING) ? hs->mJumpToLabel->mName : hk->mName;
				g_script.mThisHotkeyStartTime = GetTickCount(); // Fixed for v1.0.35.10 to not happen for GUI threads.
			}

			if (g_nFileDialogs)
				// Since there is a quasi-thread with an open file dialog underneath the one
				// we're about to launch, set the current directory to be the one the user
				// would expect to be in effect.  This is not a 100% fix/workaround for the
				// fact that the dialog changes the working directory as the user navigates
				// from folder to folder because the dialog can still function even when its
				// quasi-thread is suspended (i.e. while our new thread being launched here
				// is in the middle of running).  In other words, the user can still use
				// the dialog of a suspended quasi-thread, and thus change the working
				// directory indirectly.  But that should be very rare and I don't see an
				// easy way to fix it completely without using a "HOOK function to monitor
				// the WM_NOTIFY message", calling SetCurrentDirectory() after every script
				// line executes (which seems too high in overhead to be justified), or
				// something similar.  Note changing to a new directory here does not seem
				// to hurt the ongoing FileSelectFile() dialog.  In other words, the dialog
				// does not seem to care that its changing of the directory as the user
				// navigates is "undone" here:
				SetCurrentDirectory(g_WorkingDir);

			// Make every newly launched subroutine start off with the global default values that
			// the user set up in the auto-execute part of the script (e.g. KeyDelay, WinDelay, etc.).
			// However, we do not set ErrorLevel to anything special here (except for GUI threads, later
			// below) because it's more flexible that way (i.e. the user may want one hotkey subroutine
			// to use the value of ErrorLevel set by another):
			InitNewThread(priority, false, true, type_of_first_line);

			// Fix for v1.0.37.06:  Must do the following only after InitNewThread() has updated
			// g.IsPaused for the new thread.
			// If the current quasi-thread is paused, the thread we're about to launch
			// will not be, so the icon needs to be checked:
			g_script.UpdateTrayIcon();

			// Do this last, right before launching the thread:
			// It seems best to reset mLinesExecutedThisCycle unconditionally (now done by InitNewThread),
			// because the user has pressed a hotkey or selected a custom menu item, so would expect
			// maximum responsiveness (e.g. in a game where split second timing can matter) rather than
			// the risk that a "rest" will be done immediately by ExecUntil() just because
			// mLinesExecutedThisCycle happens to be large some prior subroutine.  The same applies to
			// mLastScriptRest, which is why that is reset also:
			g_script.mLastScriptRest = g_script.mLastPeekTime = GetTickCount();
			// v1.0.38.04: The above now resets mLastPeekTime too to reduce situations in which a thread
			// doesn't even run one line before being interrupted by another thread.  Here's how that would
			// happen: ExecUntil() would see that a Peek() is due and call PeekMessage().  The Peek() will
			// yield if we have no messages and the CPU is under heavy load, and thus the script might not
			// get another timeslice for 20ms (or even longer if there is more than one other needy process).
			// Even if the Peek() doesn't yield (i.e. we have messages), those messages might take a long time
			// to process (such as WM_PAINT) even though the script is uninterruptible.  Either way, when the
			// Peek-check completes, a long time might have passed, and the thread might now be interruptible
			// due to the interruptible-timer having expired (which is probably possible only in the no-yield
			// scenario above, since in the case of yield, ExecUntil wouldn't check messages again after the
			// yield).  Thus, the Peek-check's MsgSleep() might launch an interrupting thread before the prior
			// thread had a chance to execute even one line.  Resetting mLastPeekTime above should alleviate that,
			// perhaps even completely resolve it due to the way tickcounts tend not to change early on in
			// a timeslice (perhaps because timeslices fall exactly upon tick-count boundaries).  If it doesn't
			// completely resolve it, mLastPeekTime could instead be set to zero as a special value that
			// ExecUntil recognizes to do the following processing, but this processing reduces performance
			// by 2.5% in a simple addition-loop benchmark:
			//if (g_script.mLastPeekTime)
			//	LONG_OPERATION_UPDATE
			//else
			//	g_script.mLastPeekTime = GetTickCount();

			// Perform the new thread's subroutine:
			return_value = true; // We will return this value to indicate that we launched at least one new thread.

			switch (msg.message)
			{
			case AHK_GUI_ACTION: // Listed first for performance.
				// This indicates whether a double-click or other non-standard event launched it:
				switch (msg.wParam)
				{
				case AHK_GUI_CONTEXTMENU:
					// Caller turned on the high-order bit if this context-menu was generated via the
					// keyboard (such as AppsKey or Shift-F10).
					if (msg.lParam & 0x80000000)
					{
						msg.lParam &= ~0x80000000; // Turn off the high-order bit to reveal the control index for later use below.
						g.GuiEvent = GUI_EVENT_NORMAL;
					}
					else
						g.GuiEvent = GUI_EVENT_RCLK;
					break;
				case AHK_GUI_DROPFILES:
					g.GuiEvent = GUI_EVENT_DROPFILES;
					break;
				case AHK_GUI_CLOSE:
				case AHK_GUI_ESCAPE:
				case AHK_GUI_SIZE:
					g.GuiEvent = (GuiEventType)msg.lParam;
					break;
				default: // Control-specific event.
					g.GuiEvent = (GuiEventType)msg.lParam;
					if (pcontrol->type == GUI_CONTROL_LISTBOX && g.GuiEvent == GUI_EVENT_DBLCLK)
						event_info = 1 + (int)SendMessage(pcontrol->hwnd, LB_GETCARETINDEX, 0, 0); // +1 to convert to one-based index.
					if (pcontrol->type != GUI_CONTROL_LISTVIEW)
						break; // No further action (leave event_info as the default set higher above).
					switch (g.GuiEvent) // Don't use toupper() on it to avoid possible corruption of non-alpha values.
					{
					case 'A': // LVN_ITEMACTIVATE
					case 'E': // LVN_BEGINLABELEDIT
					case 'e': // LVN_ENDLABELEDIT
					case GUI_EVENT_DBLCLK: // Double-click.
					// v1.0.36.03: For the following, it's somewhat debatable to set event_info when the ListView
					// isn't single-select, but the usefulness seems to outweigh any confusion it might cause:
					case GUI_EVENT_NORMAL: // Left-click
					case GUI_EVENT_RCLK:   // Right-click
					case 'D': // LVN_BEGINDRAG
					case 'd': // LVN_BEGINRDRAG (right-drag)
					case 'R': // NM_RDBLCLK (right-double-click)
					//case 'I': // LVN_ITEMCHANGED. This one is not done because in the future, might want to have it report the row-number of the changed row rather than always the currently focused row.
						event_info = 1 + ListView_GetNextItem(pcontrol->hwnd, -1, LVNI_FOCUSED);
						// Testing shows that only one item at a time can have focus, even when mulitple
						// items are selected.
						break;
					}
				} // switch (msg.message)

				g.GuiWindowIndex = pgui->mWindowIndex; // g.GuiControlIndex is conditionally set later below.
				g.GuiDefaultWindowIndex = pgui->mWindowIndex; // GUI threads default to operating upon their own window.

				// Set A_EventInfo and ErrorLevel for anything that calls for it:
				switch (msg.wParam)
				{
				case AHK_GUI_SIZE:
					// Note that SIZE_MAXSHOW/SIZE_MAXHIDE don't seem to ever be received under the conditions
					// described at MSDN, even if the window has WS_POPUP style.  Therefore, event_info will
					// probably never contain those values, and as a result they are not documented in the help file.
					event_info = (DWORD)pgui->mSizeType;
					break;
				case AHK_GUI_CONTEXTMENU:
					if ((GuiIndexType)msg.lParam < pgui->mControlCount) // Index is in bounds, so this context menu is for a control.
					{
						GuiControlType &control = pgui->mControl[msg.lParam];
						if (control.type == GUI_CONTROL_LISTBOX) // +1 to convert to one-based index.
							event_info = 1 + (int)SendMessage(control.hwnd, LB_GETCARETINDEX, 0, 0); // Cast to int to preserve any -1 value.
						else if (control.type == GUI_CONTROL_LISTVIEW)
							event_info = 1 + ListView_GetNextItem(control.hwnd, -1, LVNI_FOCUSED);
							// Testing shows that only one item at a time can have focus, even when mulitple
							// items are selected.
						//else leave event_info as the default set higher above.
					}
					//else leave ErrorLevel as the default set higher above.
					break;
				case AHK_GUI_DROPFILES:
					event_info = drop_count;
					break;
				//default: No action for the other cases.
				}

				// event_info is a separate variable because it is sometimes set before g.EventInfo is available
				// for the new thread.  It also serves to indicate whether g_ErrorLevel should be made blank
				// vs. assigned a number.
				if (event_info == UINT_MAX) // No event info.
					// Reset to blank to reserve it for potential future uses (may help avoid breaking existing
					// scripts if ErrorLevel is ever set), but only for AHK_GUI_ACTION to retain backward
					// compatibility and because A_EventInfo should probably be used in lieu of ErrorLevel
					// for any future event info to be passed to the script.
					g_ErrorLevel->Assign();
				else
				{
					// Override the thread-default of zero with this event info number:
					g.EventInfo = event_info;
					g_ErrorLevel->Assign(event_info); // For backward compatibility.
				}

				// Set last found window (as documented).  It's not necessary to check IsWindow/IsWindowVisible/
				// DetectHiddenWindows since GetValidLastUsedWindow() takes care of that whenever the script
				// actually tries to use the last found window.  UPDATE: Definitely don't want to check
				// IsWindowVisible/DetectHiddenWindows now that the last-found window is exempt from
				// DetectHiddenWindows if the last-found window is one of the script's GUI windows [v1.0.25.13]:
				g.hWndLastUsed = pgui->mHwnd;

				if (pgui_label_is_running) // i.e. GuiClose, GuiEscape, and related window-level events.
					*pgui_label_is_running = true;
					// and leave g.GuiControlIndex at its default
				else if (msg.wParam == AHK_GUI_CONTEXTMENU || msg.wParam == AHK_GUI_DROPFILES) // No label-is-running attribute to set.
				{
					g.GuiControlIndex = (GuiIndexType)msg.lParam; // Index is in lParam vs. wParam in this case.
					// By definition, g.GuiControlIndex should be the focused control.  However, testing shows
					// that it can also be the only control in a window that lacks any focus-capable controls.
					// If the window has no controls at all, testing shows that g.GuiControlIndex will be out
					// of bounds, so the below will fall back to cursor position for AppsKey.
					if (msg.wParam == AHK_GUI_CONTEXTMENU && g.GuiEvent == GUI_EVENT_NORMAL // Context menu was invoked via keyboard.
						&& g.GuiControlIndex < pgui->mControlCount) // Context menu event occurred while a control had focus.
					{
						pgui->ControlGetPosOfFocusedItem(pgui->mControl[g.GuiControlIndex], g.GuiPoint);
					}
					else// This is a drop-files event or a context-menu event that was invoked via normal mouse click.
						// This also happens when there is no focused control, in which case it seems best to report
						// the cursor's position rather than some arbitrary center-point, or top-left point in the
						// parent window.  This is because it might be more convenient for the user to move the
						// mouse to select a menu item (since menu will be close to mouse cursor).
						g.GuiPoint = msg.pt; // v1.0.38: More accurate/customary to use msg.pt than GetCursorPos().
					// Convert screen coordinates to window coordinates.  Window coords seem best by default because
					// the "Menu Show" command uses them by default.  A CoordMode option can be added to change this
					// if it is ever needed.
					GetWindowRect(pgui->mHwnd, &rect);
					g.GuiPoint.x -= rect.left;
					g.GuiPoint.y -= rect.top;
					if (msg.wParam == AHK_GUI_DROPFILES)
						// Visually indicate that drops aren't allowed while and existing drop is still being
						// processed. Fix for v1.0.31.02: The window's current ExStyle is fetched every time
						// in case a non-GUI command altered it (such as making it transparent):
						SetWindowLong(pgui->mHwnd, GWL_EXSTYLE, GetWindowLong(pgui->mHwnd, GWL_EXSTYLE) & ~WS_EX_ACCEPTFILES);
				}
				else // It's a control's action, so set its attribute.
				{
					pcontrol->attrib |= GUI_CONTROL_ATTRIB_LABEL_IS_RUNNING;
					g.GuiControlIndex = (GuiIndexType)msg.wParam;
				}

				// LAUNCH THREAD:
				gui_label->mJumpToLine->ExecUntil(UNTIL_RETURN);

				// Bug-fix for v1.0.22: If the above ExecUntil() performed a "Gui Destroy", the
				// pointers below are now invalid so should not be dereferenced.  In such a case,
				// hdrop_to_free will already have been freed as part of the window destruction
				// process, so don't do it here.  g_gui[gui_index] is checked to ensure the window
				// still exists:
				if (pgui = g_gui[gui_index]) // Assign.  This refresh is done as explained below.
				{
					// Bug-fix for v1.0.30.04: If the thread that was just launched above destroyed
					// its own GUI window, but then recreated it, that window's members obviously aren't
					// guaranteed to have the same memory addresses that they did prior to destruction.
					// Even g_gui[gui_index] would probably be a different address, so pgui would be
					// invalid too.  Therefore, refresh the original pointers (pgui is refreshed above).
					// See similar switch() higher above for comments about the below:
					switch(msg.wParam)
					{
					case AHK_GUI_SIZE:   pgui->mLabelForSizeIsRunning = false; break;   // Safe to reset even if there is
					case AHK_GUI_CLOSE:  pgui->mLabelForCloseIsRunning = false; break;  // no label due to the window having
					case AHK_GUI_ESCAPE: pgui->mLabelForEscapeIsRunning = false; break; // been destroyed and recreated.
					case AHK_GUI_CONTEXTMENU: break; // Do nothing, but avoid the default case below.
					case AHK_GUI_DROPFILES:
						if (pgui->mHdrop) // It's no longer safer to refer to hdrop_to_free (see comments above).
						{
							DragFinish(pgui->mHdrop); // Since the DropFiles quasi-thread is finished, free the HDROP resources.
							pgui->mHdrop = NULL; // Indicate that this GUI window is ready for another drop.
						}
						// Fix for v1.0.31.02: The window's current ExStyle is fetched every time in case a non-GUI
						// command altered it (such as making it transparent):
						SetWindowLong(pgui->mHwnd, GWL_EXSTYLE, GetWindowLong(pgui->mHwnd, GWL_EXSTYLE) | WS_EX_ACCEPTFILES);
						break;
					default: // It's a control's action, so set its attribute.
						if (msg.wParam < pgui->mControlCount) // Recheck to ensure that control still exists (in case window was recreated as explained above).
							pgui->mControl[msg.wParam].attrib &= ~GUI_CONTROL_ATTRIB_LABEL_IS_RUNNING; // Don't use pcontrol here for the reasons above.
					}
				} // if (this gui window wasn't destroyed-without-recreation by the thread we just launched).
				break;

			case AHK_USER_MENU: // user-defined menu item
				// Below: the menu type is passed with the message so that its value will be in sync
				// with the timestamp of the message (in case this message has been stuck in the
				// queue for a long time):
				if (msg.wParam < MAX_GUI_WINDOWS) // Poster specified that this menu item was from a gui's menu bar (since wParam is unsigned, any incoming -1 is seen as greater than max).
				{
					// msg.wParam is the index rather than a pointer to avoid any chance of problems with
					// a gui object or its window having been destroyed while the msg was waiting in the queue.
					// As documented, set the last found window if possible/applicable.  It's not necessary to
					// check IsWindow/IsWindowVisible/DetectHiddenWindows since GetValidLastUsedWindow()
					// takes care of that whenever the script actually tries to use the last found window.
					if (g_gui[msg.wParam])
					{
						g.hWndLastUsed = g_gui[msg.wParam]->mHwnd; // OK if NULL.
						// This flags GUI menu items as being GUI so that the script has a way of detecting
						// whether a given submenu's item was selected from inside a menu bar vs. a popup:
						g.GuiEvent = GUI_EVENT_NORMAL;
						g.GuiWindowIndex = g.GuiDefaultWindowIndex = (GuiIndexType)msg.wParam; // But leave GuiControl at its default, which flags this event as from a menu item.
					}
				}
				menu_item->mLabel->mJumpToLine->ExecUntil(UNTIL_RETURN);
				break;

			case AHK_HOTSTRING:
				g.hWndLastUsed = criterion_found_hwnd; // v1.0.42. Even if the window is invalid for some reason, IsWindow() and such are called whenever the script accesses it (GetValidLastUsedWindow()).
				hs->Perform();
				break;

			case AHK_CLIPBOARD_CHANGE:
				g.EventInfo = CountClipboardFormats() ? (IsClipboardFormatAvailable(CF_TEXT) || IsClipboardFormatAvailable(CF_HDROP) ? 1 : 2) : 0;
				g_ErrorLevel->Assign(g.EventInfo); // For backward compatibility.
				// ACT_IS_ALWAYS_ALLOWED() was already checked above.
				// The message poster has ensured that g_script.mOnClipboardChangeLabel is non-NULL and valid.
				g_script.mOnClipboardChangeIsRunning = true;
				g_script.mOnClipboardChangeLabel->mJumpToLine->ExecUntil(UNTIL_RETURN); // EXECUTE.
				g_script.mOnClipboardChangeIsRunning = false;
				break;

			default: // hotkey
				g.hWndLastUsed = criterion_found_hwnd; // v1.0.42. Even if the window is invalid for some reason, IsWindow() and such are called whenever the script accesses it (GetValidLastUsedWindow()).
				hk->Perform(*variant);
			}

			// v1.0.37.06: Call ResumeUnderlyingThread() even if aMode==WAIT_FOR_MESSAGES; this is for
			// maintainability and also in case the pause command has been used to unpause the idle thread.
			ResumeUnderlyingThread(&global_saved, true);

			if (aMode == WAIT_FOR_MESSAGES) // This is the "idle thread", meaning that the end of the thread above has returned the script to an idle state.
			{
				// v1.0.38.04: The following line is for maintainability and reliability.  It avoids the need
				// for other sections to figure out whether they should reset g.ThreadIsCritical to false
				// when a thread is resumed (a resumed thread might still be critical if it was interrupted
				// by an emergency thread such as OnExit or OnMessage).
				g.AllowThreadToBeInterrupted = true; // This one is probably necessary due to the way it conforms to ThreadIsCritical in other sections.
				g.ThreadIsCritical = false; // Not strictly necessary but improves maintainability.
				g.AllowTimers = true; // Same as above.
				g.Priority = PRIORITY_MINIMUM; // Ensure minimum priority so that idle state can always be "interrupted".
			}
			else // Some thread other than the idle thread.
			{
				if (IsCycleComplete(aSleepDuration, start_time, allow_early_return))
				{
					// Check for messages once more in case the subroutine that just completed
					// above didn't check them that recently.  This is done to minimize the time
					// our thread spends *not* pumping messages, which in turn minimizes keyboard
					// and mouse lag if the hooks are installed (even though this is no longer
					// true due to v1.0.39's dedicated hook thread, it seems best to continue this
					// practice to maximize responsiveness of hotkeys, the app itself [e.g. tray
					// menu], and also to retain backward compatibility).  Set the state of this
					// function/layer/instance so that it will use peek-mode.  UPDATE: Don't change
					// the value of aSleepDuration to -1 because IsCycleComplete() needs to know the
					// original sleep time specified by the caller to determine whether
					// to decrement g_nLayersNeedingTimer:
					empty_the_queue_via_peek = true;
					allow_early_return = true;
					// And now let it fall through to the "continue" statement below.
				}
				else if (this_layer_needs_timer) // Ensure the timer is back on since we still need it here.
					SET_MAIN_TIMER // This won't do anything if it's already on.
				// and now if the cycle isn't complete, stay in the blessed GetMessage() state until the time
				// has expired.
			}
			continue;

#ifdef _DEBUG
		case AHK_HOOK_TEST_MSG:
		{
			char dlg_text[512];
			snprintf(dlg_text, sizeof(dlg_text), "TEST MSG: %d (0x%X)  %d (0x%X)"
				"\nCurrent Thread: 0x%X"
				, msg.wParam, msg.wParam, msg.lParam, msg.lParam
				, GetCurrentThreadId());
			MsgBox(dlg_text);
			continue;
		}
#endif

		case WM_KEYDOWN:
			if (msg.hwnd == g_hWndEdit && msg.wParam == VK_ESCAPE)
			{
				// This won't work if a MessageBox() window is displayed because its own internal
				// message pump will dispatch the message to our edit control, which will just
				// ignore it.  And avoiding setting the focus to the edit control won't work either
				// because the user can simply click in the window to set the focus.  But for now,
				// this is better than nothing:
				ShowWindow(g_hWnd, SW_HIDE);  // And it's okay if this msg gets dispatched also.
				continue;
			}
			// Otherwise, break so that the messages will get dispatched.  We need the other
			// WM_KEYDOWN msgs to be dispatched so that the cursor is keyboard-controllable in
			// the edit window:
			break;

		case WM_INPUTLANGCHANGEREQUEST:
			// WM_INPUTLANGCHANGE does not appear to be received.  I think that msg is sent directly to
			// the Window Proc of the Edit control in the main window. By contrast, MSDN says:
			// "[WM_INPUTLANGCHANGE] is posted, not sent, to the application, so the return value is ignored."
			// In addition, although WM_INPUTLANGCHANGEREQUEST is dispatched via DispatchMessage(),
			// MainWindowProc() apparently never receives it (once again, this is because msg.hwnd is that
			// of the main window's Edit control rather than the main window itself). But all we need to do
			// is reset the AltGr indicator, so it's okay to handle it here.
			g_LayoutHasAltGr = false; // Seems safest to assume it doesn't have one until proven otherwise.
			break; // MSDN says, "to accept the change, the application should pass the message to DefWindowProc."

		case WM_QUIT:
			// The app normally terminates before WM_QUIT is ever seen here because of the way
			// WM_CLOSE is handled by MainWindowProc().  However, this is kept here in case anything
			// external ever explicitly posts a WM_QUIT to our thread's queue:
			g_script.ExitApp(EXIT_WM_QUIT);
			continue; // Since ExitApp() won't necessarily exit.
		} // switch()

		// If a "continue" statement wasn't encountered somewhere in the switch(), we want to
		// process this message in a more generic way.
		// This little part is from the Miranda source code.  But it doesn't seem
		// to provide any additional functionality: You still can't use keyboard
		// keys to navigate in the dialog unless it's the topmost dialog.
		// UPDATE: The reason it doesn't work for non-topmost MessageBoxes is that
		// this message pump isn't even the one running.  It's the pump of the
		// top-most MessageBox itself, which apparently doesn't properly dispatch
		// all types of messages to other MessagesBoxes.  However, keeping this
		// here is probably a good idea because testing reveals that it does
		// sometimes receive messages intended for MessageBox windows (which makes
		// sense because our message pump here retrieves all thread messages).
		// It might cause problems to dispatch such messages directly, since
		// IsDialogMessage() is supposed to be used in lieu of DispatchMessage()
		// for these types of messages.
		// NOTE: THE BELOW IS CONFIRMED to be needed, at least for a FileSelectFile()
		// dialog whose quasi-thread has been suspended, and probably for some of the other
		// types of dialogs as well:
		if ((fore_window = GetForegroundWindow()) != NULL)  // There is a foreground window.
		{
			GetWindowThreadProcessId(fore_window, &fore_pid);
			if (fore_pid == GetCurrentProcessId())  // It belongs to our process.
			{
				GetClassName(fore_window, fore_class_name, sizeof(fore_class_name));
				if (!strcmp(fore_class_name, "#32770"))  // MessageBox(), InputBox(), or FileSelectFile() window.
				{
					g.CalledByIsDialogMessageOrDispatch = true; // In case there is any way IsDialogMessage() can call one of our own window proc's rather than that of a MsgBox, etc.
					if (IsDialogMessage(fore_window, &msg))  // This message is for it, so let it process it.
					{
						// If it is likely that a FileSelectFile dialog is active, this
						// section attempt to retain the current directory as the user
						// navigates from folder to folder.  This is done because it is
						// possible that our caller is a quasi-thread other than the one
						// that originally launched the FileSelectFile (i.e. that dialog
						// is in a suspended thread), in which case the user's navigation
						// would cause the active threads working dir to change unexpectedly
						// unless the below is done.  This is not a complete fix since if
						// a message pump other than this one is running (e.g. that of a
						// MessageBox()), these messages will not be detected:
						if (g_nFileDialogs) // See MsgSleep() for comments on this.
							// The below two messages that are likely connected with a user
							// navigating to a different folder within the FileSelectFile dialog.
							// That avoids changing the directory for every message, since there
							// can easily be thousands of such messages every second if the
							// user is moving the mouse.  UPDATE: This doesn't work, so for now,
							// just call SetCurrentDirectory() for every message, which does work.
							// A brief test of CPU utilization indicates that SetCurrentDirectory()
							// is not a very high overhead call when it is called many times here:
							//if (msg.message == WM_ERASEBKGND || msg.message == WM_DELETEITEM)
							SetCurrentDirectory(g_WorkingDir);
						g.CalledByIsDialogMessageOrDispatch = false;
						continue;  // This message is done, so start a new iteration to get another msg.
					}
					g.CalledByIsDialogMessageOrDispatch = false;
				}
			}
		}
		// Translate keyboard input for any of our thread's windows that need it:
		if (!g_hAccelTable || !TranslateAccelerator(g_hWnd, g_hAccelTable, &msg))
		{
			TranslateMessage(&msg);
			g.CalledByIsDialogMessageOrDispatch = true; // Relies on the fact that the types of messages we dispatch can't result in a recursive call back to this function.
			DispatchMessage(&msg); // This is needed to send keyboard input and other messages to various windows and for some WM_TIMERs.
			g.CalledByIsDialogMessageOrDispatch = false;
		}
	} // infinite-loop
}



ResultType IsCycleComplete(int aSleepDuration, DWORD aStartTime, bool aAllowEarlyReturn)
// This function is used just to make MsgSleep() more readable/understandable.
{
	// Note: Even if TickCount has wrapped due to system being up more than about 49 days,
	// DWORD math still gives the right answer as long as aStartTime itself isn't more
	// than about 49 days ago. Note: must cast to int or any negative result will be lost
	// due to DWORD type:
	DWORD tick_now = GetTickCount();
	if (!aAllowEarlyReturn && (int)(aSleepDuration - (tick_now - aStartTime)) > SLEEP_INTERVAL_HALF)
		// Early return isn't allowed and the time remaining is large enough that we need to
		// wait some more (small amounts of remaining time can't be effectively waited for
		// due to the 10ms granularity limit of SetTimer):
		return FAIL; // Tell the caller to wait some more.

	// Update for v1.0.20: In spite of the new resets of mLinesExecutedThisCycle that now appear
	// in MsgSleep(), it seems best to retain this reset here for peace of mind, maintainability,
	// and because it might be necessary in some cases (a full study was not done):
	// Reset counter for the caller of our caller, any time the thread
	// has had a chance to be idle (even if that idle time was done at a deeper
	// recursion level and not by this one), since the CPU will have been
	// given a rest, which is the main (perhaps only?) reason for using BatchLines
	// (e.g. to be more friendly toward time-critical apps such as games,
	// video capture, video playback).  UPDATE: mLastScriptRest is also reset
	// here because it has a very similar purpose.
	if (aSleepDuration > -1)
	{
		g_script.mLinesExecutedThisCycle = 0;
		g_script.mLastScriptRest = tick_now;
	}
	// v1.0.38.04: Reset mLastPeekTime because caller has just done a GetMessage() or PeekMessage(),
	// both of which should have routed events to the keyboard/mouse hooks like LONG_OPERATION_UPDATE's
	// PeekMessage() and thus satisified the reason that mLastPeekTime is tracked in the first place.
	// This might also improve performance slightly by avoiding extra Peek() calls, while also reducing
	// premature thread interruptions.
	g_script.mLastPeekTime = tick_now;
	return OK;
}



bool CheckScriptTimers()
// Returns true if it launched at least one thread, and false otherwise.
// It's best to call this function only directly from MsgSleep() or when there is an instance of
// MsgSleep() closer on the call stack than the nearest dialog's message pump (e.g. MsgBox).
// This is because threads some events might get queued up for our thread during the execution
// of the timer subroutines here.  When those subroutines finish, if we return directly to a dialog's
// message pump, and such pending messages might be discarded or mishandled.
// Caller should already have checked the value of g_script.mTimerEnabledCount to ensure it's
// greater than zero, since we don't check that here (for performance).
// This function will go through the list of timed subroutines only once and then return to its caller.
// It does it only once so that it won't keep a thread beneath it permanently suspended if the sum
// total of all timer durations is too large to be run at their specified frequencies.
// This function is allowed to be called recursively, which handles certain situations better:
// 1) A hotkey subroutine interrupted and "buried" one of the timer subroutines in the stack.
//    In this case, we don't want all the timers blocked just because that one is, so recursive
//    calls from ExecUntil() are allowed, and they might discover other timers to run.
// 2) If the script is idle but one of the timers winds up taking a long time to execute (perhaps
//    it gets stuck in a long WinWait), we want a recursive call (from MsgSleep() in this example)
//    to launch any other enabled timers concurrently with the first, so that they're not neglected
//    just because one of the timers happens to be long-running.
// Of course, it's up to the user to design timers so that they don't cause problems when they
// interrupted hotkey subroutines, or when they themselves are interrupted by hotkey subroutines
// or other timer subroutines.
{
	// When this is true, such as during a SendKeys() operation, it seems best not to launch any new
	// timed subroutines.  The reasons for this are similar to the reasons for not allowing hotkeys
	// to fire during such times.  Those reasons are discussed in other comments.  In addition,
	// it seems best as a policy not to allow timed subroutines to run while the script's current
	// quasi-thread is paused.  Doing so would make the tray icon flicker (were it even updated below,
	// which it currently isn't) and in any case is probably not what the user would want.  Most of the
	// time, the user would want all timed subroutines stopped while the current thread is paused.
	// And even if this weren't true, the confusion caused by the subroutines still running even when
	// the current thread is paused isn't worth defaulting to the opposite approach.  In the future,
	// and if there's demand, perhaps a config option can added that allows a different default behavior.
	// UPDATE: It seems slightly better (more consistent) to disallow all timed subroutines whenever
	// there is even one paused thread anywhere in the "stack".  UPDATE for v1.0.37.06: g_IdleIsPaused is
	// now checked in case the "idle thread" is paused (since that thread is not counted in
	// g_nPausedThreads).  However, g_nPausedThreads must still be checked in case the uppermost thread
	// isn't paused but some other thread isn't (as documented, timers don't run when any thread is paused).
	if (!INTERRUPTIBLE || g_nPausedThreads > 0 || g_IdleIsPaused || !g.AllowTimers || g_nThreads >= g_MaxThreadsTotal)
		return false; // Above: To be safe (prevent stack faults) don't allow max threads to be exceeded.

	ScriptTimer *timer;
	UINT launched_threads;
	DWORD tick_start;
	global_struct global_saved;

	// Note: It seems inconsequential if a subroutine that the below loop executes causes a
	// new timer to be added to the linked list while the loop is still enumerating the timers.

	for (launched_threads = 0, timer = g_script.mFirstTimer; timer != NULL; timer = timer->mNextTimer)
	{
		// Call GetTickCount() every time in case a previous iteration of the loop took a long
		// time to execute.  Also, as of v1.0.36.03, the following subtracts two DWORDs to support
		// intervals of 49.7 vs. 24.8 days.  This should work as long as the distance between the
		// values being compared isn't greater than 49.7 days. This is because 1 minus 2 in unsigned
		// math yields 0xFFFFFFFF milliseconds (49.7 days).
		if (timer->mEnabled && timer->mExistingThreads < 1 && timer->mPriority >= g.Priority // thread priorities
			&& (tick_start = GetTickCount()) - timer->mTimeLastRun >= (DWORD)timer->mPeriod)
		{
			if (!launched_threads)
			{
				// Since this is the first subroutine that will be launched during this call to
				// this function, we know it will wind up running at least one subroutine, so
				// certain changes are made:
				// Increment the count of quasi-threads only once because this instance of this
				// function will never create more than 1 thread (i.e. if there is more than one
				// enabled timer subroutine, the will always be run sequentially by this instance).
				// If g_nThreads is zero, incrementing it will also effectively mark the script as
				// non-idle, the main consequence being that an otherwise-idle script can be paused
				// if the user happens to do it at the moment a timed subroutine is running, which
				// seems best since some timed subroutines might take a long time to run:
				++g_nThreads;
				// Next, save the current state of the globals so that they can be restored just prior
				// to returning to our caller:
				strlcpy(g.ErrorLevel, g_ErrorLevel->Contents(), sizeof(g.ErrorLevel)); // Save caller's errorlevel.
				CopyMemory(&global_saved, &g, sizeof(global_struct));
				// But never kill the main timer, since the mere fact that we're here means that
				// there's at least one enabled timed subroutine.  Though later, performance can
				// be optimized by killing it if there's exactly one enabled subroutine, or if
				// all the subroutines are already in a running state (due to being buried beneath
				// the current quasi-thread).  However, that might introduce unwanted complexity
				// in other places that would need to start up the timer again because we stopped it, etc.
			}

			// Fix for v1.0.31: mTimeLastRun is now given its new value *before* the thread is launched
			// rather than after.  This allows a timer to be reset by its own thread -- by means of
			// "SetTimer, TimerName", which is otherwise impossible because the reset was being
			// overridden by us here when the thread finished.
			// Seems better to store the start time rather than the finish time, though it's clearly
			// debatable.  The reason is that it's sometimes more important to ensure that a given
			// timed subroutine is *begun* at the specified interval, rather than assuming that
			// the specified interval is the time between when the prior run finished and the new
			// one began.  This should make timers behave more consistently (i.e. how long a timed
			// subroutine takes to run SHOULD NOT affect its *apparent* frequency, which is number
			// of times per second or per minute that we actually attempt to run it):
			timer->mTimeLastRun = tick_start;
			++launched_threads;

			if (g_nFileDialogs) // See MsgSleep() for comments on this.
				SetCurrentDirectory(g_WorkingDir);

			// Note that it is not necessary to call UpdateTrayIcon() here since timers won't run
			// if there is any paused thread, thus the icon can't currently be showing "paused".

			// This next line is necessary in case a prior iteration of our loop invoked a different
			// timed subroutine that changed any of the global struct's values.  In other words, make
			// every newly launched subroutine start off with the global default values that
			// the user set up in the auto-execute part of the script (e.g. KeyDelay, WinDelay, etc.).
			// However, we do not set ErrorLevel to NONE here because it's more flexible that way
			// (i.e. the user may want one hotkey subroutine to use the value of ErrorLevel set by another):
			InitNewThread(timer->mPriority, false, false, timer->mLabel->mJumpToLine->mActionType); // False as last param because ++g_nThreads should be done only once rather than each Init().
			// The above also resets g_script.mLinesExecutedThisCycle to zero, which should slightly
			// increase the expectation that any short timed subroutine will run all the way through
			// to completion rather than being interrupted by the press of a hotkey, and thus potentially
			// buried in the stack.  However, mLastScriptRest is not set to GetTickCount() here because
			// unlike other events -- which are typically in response to an explicit action by the user
			// such as pressing a button or hotkey -- times are lower priority and more relaxed.
			// Also, mLastScriptRest really should only be set when a call to Get/PeekMsg has just
			// occurred, so it should be left as the responsibilty of the section in MsgSleep that
			// launches new threads.

			++timer->mExistingThreads;
			timer->mLabel->mJumpToLine->ExecUntil(UNTIL_RETURN);
			--timer->mExistingThreads;

			KILL_UNINTERRUPTIBLE_TIMER
		} // if timer is due to launch.
	} // for() each timer.

	if (launched_threads) // Since at least one subroutine was run above, restore various values for our caller.
	{
		ResumeUnderlyingThread(&global_saved, false); // Last param "false" because KILL_UNINTERRUPTIBLE_TIMER was already done above.
		return true;
	}
	return false;
}



void PollJoysticks()
// It's best to call this function only directly from MsgSleep() or when there is an instance of
// MsgSleep() closer on the call stack than the nearest dialog's message pump (e.g. MsgBox).
// This is because events posted to the thread indirectly by us here would be discarded or mishandled
// by a non-standard (dialog) message pump.
// Polling the joysticks this way rather than using joySetCapture() is preferable for several reasons:
// 1) I believe joySetCapture() internally polls the joystick anyway, via a system timer, so it probably
//    doesn't perform much better (if at all) than polling "manually".
// 2) joySetCapture() only supports 4 buttons;
// 3) joySetCapture() will fail if another app is already capturing the joystick;
// 4) Even if the joySetCapture() succeeds, other programs (e.g. older games), would be prevented from
//    capturing the joystick while the script in question is running.
{
	// Even if joystick hotkeys aren't currently allowed to fire, poll it anyway so that hotkey
	// messages can be buffered for later.
	static DWORD sButtonsPrev[MAX_JOYSTICKS] = {0}; // Set initial state to "all buttons up for all joysticks".
	JOYINFOEX jie;
	DWORD buttons_newly_down;

	for (int i = 0; i < MAX_JOYSTICKS; ++i)
	{
		if (!Hotkey::sJoystickHasHotkeys[i])
			continue;
		// Reset these every time in case joyGetPosEx() ever changes them:
		jie.dwSize = sizeof(JOYINFOEX);
		jie.dwFlags = JOY_RETURNBUTTONS; // vs. JOY_RETURNALL
		if (joyGetPosEx(i, &jie) != JOYERR_NOERROR) // Skip this joystick and try the others.
			continue;
		// The exclusive-or operator determines which buttons have changed state.  After that,
		// the bitwise-and operator determines which of those have gone from up to down (the
		// down-to-up events are currently not significant).
		buttons_newly_down = (jie.dwButtons ^ sButtonsPrev[i]) & jie.dwButtons;
		sButtonsPrev[i] = jie.dwButtons;
		if (!buttons_newly_down)
			continue;
		// See if any of the joystick hotkeys match this joystick ID and one of the buttons that
		// has just been pressed on it.  If so, queue up (buffer) the hotkey events so that they will
		// be processed when messages are next checked:
		Hotkey::TriggerJoyHotkeys(i, buttons_newly_down);
	}
}



bool MsgMonitor(HWND aWnd, UINT aMsg, WPARAM awParam, LPARAM alParam, MSG *apMsg, LRESULT &aMsgReply)
// Returns false if the message is not being monitored, or it is but the called function indicated
// that the message should be given its normal processing.  Returns true when the caller should
// not process this message but should instead immediately reply with aMsgReply (if a reply is possible).
// When false is returned, caller should ignore the value of aMsgReply.
{
	// This function directly launches new threads rather than posting them as something like
	// AHK_GUI_ACTION (which would allow them to be queued by means of MSG_FILTER_MAX) because a message
	// monitor function in the script can return "true" to exempt the message from further processing.
	// Consequently, the MSG_FILTER_MAX queuing effect will only occur for monitored messages that are
	// numerically greater than WM_HOTKEY. Other messages will not be subject to the filter and thus
	// will arrive here even when the script is currently uninterruptible, in which case it seems best
	// to discard the message because the current design doesn't allow for interruptions. The design
	// could be reviewed to find out what the consequences of interruption would be.  Also, the message
	// could be suppressed (via return 1) and reposted, but if there are other messages already in
	// the queue that qualify to fire a msg-filter (or even messages such as WM_LBUTTONDOWN that have
	// a normal effect that relies on ordering), the messages would then be processed out of their
	// original order, which would be very undesirable in many cases.
	//
	// In light of the above, INTERRUPTIBLE_IF_NECESSARY is used instead of INTERRUPTIBLE_IN_EMERGENCY
	// to reduce on the unreliability of message filters that are numerically less than WM_HOTKEY.
	// For example, if the user presses a hotkey and an instant later a qualified WM_LBUTTONDOWN arrives,
	// the filter will still be able to run by interrupting the uinterruptible thread.  In this case,
	// ResumeUnderlyingThread() sets g.AllowThreadToBeInterrupted to false for us in case the
	// timer "TIMER_ID_UNINTERRUPTIBLE" fired for the new thread rather than for the old one (this
	// prevents the interrupted thread from becoming permanently uninterruptible).
	if (!INTERRUPTIBLE_IN_EMERGENCY)
		return false;

	// Linear search vs. binary search should perform better on average because the vast majority
	// of message monitoring scripts are expected to monitor only a few message numbers.
	int msg_index;
	for (msg_index = 0; msg_index < g_MsgMonitorCount; ++msg_index)
		if (g_MsgMonitor[msg_index].msg == aMsg)
			break;
	if (msg_index == g_MsgMonitorCount) // No match found, so the script isn't monitoring this message.
		return false; // Tell the caller to give this message any additional/default processing.
	// Otherwise, the script is monitoring this message, so continue on.

	MsgMonitorStruct &monitor = g_MsgMonitor[msg_index]; // For performance and convenience.
	Func &func = *monitor.func;                          // Above, but also in case monitor item gets deleted while the function is running (e.g. by the function itself).

	// Many of the things done below are similar to the thread-launch procedure used in MsgSleep(),
	// so maintain them together and see MsgSleep() for more detailed commments.
	if (g_nThreads >= g_MaxThreadsTotal)
		// Below: Only a subset of ACT_IS_ALWAYS_ALLOWED is done here because:
		// 1) The omitted action types seem too obscure to grant always-run permission for msg-monitor events.
		// 2) Reduction in code size.
		if (g_nThreads > MAX_THREADS_LIMIT
			|| func.mJumpToLine->mActionType != ACT_EXITAPP && func.mJumpToLine->mActionType != ACT_RELOAD)
			return false;
	if (monitor.label_is_running || g.Priority > 0) // Monitor is already running its function or existing thread's priority is too high to be interrupted.
		return false;

	// Need to check if backup is needed in case script explicitly called the function rather than using
	// it solely as a callback.
	// See ExpandExpression() for detailed comments about the following section.
	VarBkp *var_backup = NULL;   // If needed, it will hold an array of VarBkp objects.
	int var_backup_count; // The number of items in the above array.
	if (func.mInstances > 0) // Backup is needed.
		if (!BackupFunctionVars(func, var_backup, var_backup_count)) // Out of memory.
			return false;
			// Since we're in the middle of processing messages, and since out-of-memory is so rare,
			// it seems justifiable not to have any error reporting and instead just avoid launching
			// the new thread.

	// Since above didn't return, the launch of the new thread is now considered unavoidable.

	// See MsgSleep() for comments about the following section.
	strlcpy(g.ErrorLevel, g_ErrorLevel->Contents(), sizeof(g.ErrorLevel));
	global_struct global_saved;
	CopyMemory(&global_saved, &g, sizeof(global_struct));
	if (g_nFileDialogs)
		SetCurrentDirectory(g_WorkingDir);

	monitor.label_is_running = true;
	InitNewThread(0, false, true, func.mJumpToLine->mActionType);
	g_script.UpdateTrayIcon();

	// Set last found window (as documented).  Can be NULL.
	// Nested controls like ComboBoxes require more than a simple call to GetParent().
	if (g.hWndLastUsed = GetNonChildParent(aWnd)) // Assign parent window as the last found window (it's ok if it's hidden).
	{
		GuiType *pgui = GuiType::FindGui(g.hWndLastUsed);
		if (pgui) // This parent window is a GUI window.
		{
			g.GuiWindowIndex = pgui->mWindowIndex;  // Update the built-in variable A_GUI.
			g.GuiDefaultWindowIndex = pgui->mWindowIndex; // Consider this a GUI thread; so it defaults to operating upon its own window.
			GuiIndexType control_index = GUI_HWND_TO_INDEX(aWnd);
			if (control_index != -1) // Match found (compares directly to -1 due to unsigned).
				g.GuiControlIndex = control_index;
			//else leave it at its default, which was set when the new thread was initialized.
		}
		//else leave the above members at their default values set when the new thread was initialized.
	}
	if (apMsg)
	{
		g.GuiPoint = apMsg->pt;
		g.EventInfo = apMsg->time;
	}
	//else leave them at their init-thread defaults.

	// See ExpandExpression() for detailed comments about the following section.
	if (func.mParamCount > 0)
	{
		// Copy the appropriate values into each of the function's formal parameters.
		func.mParam[0].var->Assign((DWORD)awParam); // Assign parameter #1: wParam
		if (func.mParamCount > 1) // Assign parameter #2: lParam
		{
			// v1.0.38.01: LPARAM is now written out as a DWORD because the majority of system messages
			// use LPARAM as a pointer or other unsigned value.  This shouldn't affect most scripts because
			// of the way ATOI64() and ATOU() wrap a negative number back into the unsigned domain for
			// commands such as PostMessage/SendMessage.
			func.mParam[1].var->Assign((DWORD)alParam);
			if (func.mParamCount > 2) // Assign parameter #3: Message number (in case this function monitors more than one).
			{
				func.mParam[2].var->AssignHWND((HWND)aMsg); // Write msg number as hex because it's a lot more common. Casting issues make it easier to retain the name "AssignHWND".
				if (func.mParamCount > 3) // Assign parameter #4: HWND (listed last since most scripts won't use it for anything).
					func.mParam[3].var->AssignHWND(aWnd); // Can be a parent or child window.
			}
		}
	}

	// v1.0.38.04: Below was added to maximize responsiveness to incoming messages.  The reasoning
	// is similar to why the same thing is done in MsgSleep() prior to its launch of a thread, so see
	// MsgSleep for more comments:
	g_script.mLastScriptRest = g_script.mLastPeekTime = GetTickCount();
	char *return_value = ""; // Init to default in case function doesn't return a value or it EXITs or fails.

	// See ExpandExpression() for detailed comments about the following section.
	Func *prev_func = g.CurrentFunc; // This will be non-NULL when a function is called from inside another function.
	g.CurrentFunc = &func;
	++func.mInstances;
	ResultType result = func.mJumpToLine->ExecUntil(UNTIL_BLOCK_END, &return_value);
	--func.mInstances;
	g.CurrentFunc = prev_func;

	int i;
	for (i = 0; i < func.mVarCount; ++i)
		func.mVar[i]->Free(VAR_ALWAYS_FREE_EXCLUDE_STATIC, true); // Pass "true" to exclude aliases, since their targets should not be freed (they don't belong to this function).
	for (i = 0; i < func.mLazyVarCount; ++i)
		func.mLazyVar[i]->Free(VAR_ALWAYS_FREE_EXCLUDE_STATIC, true);
	if (var_backup) // This is the indicator that a backup was made, a restore is also needed.
		RestoreFunctionVars(func, var_backup, var_backup_count); // It avoids restoring statics.

	ResumeUnderlyingThread(&global_saved, true);
	// Check that the msg_index item still exists (it may have been deleted during the thread that just finished,
	// either by the thread itself or some other thread that interrupted it).  BIF_OnMessage has been sure to
	// reset deleted array elements to have a NULL func.  Even so, the following scenario could happen:
	// 1) The message element is deleted.
	// 2) It is recreated to be the same as before, but now it has a different array index.
	// 3) It's label_is_running member would have been set to false upon creation, and the thread for the same
	//    message might have launched the same function we did above, or some other.
	// 4) Everything seems okay in this case, especially given its rarity.
	//
	// But what if step 2 above created the same msg+func in the same position as before?  It's label_is_running
	// member would have been wrongly set to false, which would have allowed this msg-monitor thread to launch
	// while it was techically still running above.  This scenario seems too rare and the consequences too small
	// to justify the extra code size, so it is documented here as a known limitation.
	//
	// Thus, if "monitor" is defunct due to deletion, setting its label_is_running to false is harmless.
	// However, "monitor" might have been reused by BIF_OnMessage() to create a new msg-monitor, so the
	// thing that must be checked is the message number to avoid wrongly setting some other msg-monitor's
	// label_is_running to false:
	if (monitor.msg == aMsg)
		monitor.label_is_running = false;
	//else "monitor" is now some other msg-monitor, so do don't change it (see above comments).

	if (!*return_value || result == EARLY_EXIT || result == FAIL) // Tell the caller to process this message normally.
		return false; // The caller should ignore the value of aMsgReply in this case.
	// Otherwise true will be returned, in which case the aMsgReply contains the reply to be sent for
	// this message.
	aMsgReply = (LPARAM)ATOI64(return_value); // Use 64-bit in case it's an unsigned number greater than 0x7FFFFFFF, in which case this allows it to wrap around to a negative.
	return true;
}



void InitNewThread(int aPriority, bool aSkipUninterruptible, bool aIncrementThreadCount
	, ActionTypeType aTypeOfFirstLine)
// The value of aTypeOfFirstLine is ignored when aSkipUninterruptible==true.
// To reduce the expectation that a newly launched hotkey or timed subroutine will
// be immediately interrupted by a timed subroutine or hotkey, interruptions are
// forbidden for a short time (user-configurable).  If the subroutine is a quick one --
// finishing prior to when ExecUntil() or the Timer would have set g_AllowInterruption to be
// true -- we will set it to be true afterward so that it gets done as quickly as possible.
// The following rules of precedence apply:
// If either UninterruptibleTime or UninterruptedLineCountMax is zero, newly launched subroutines
// are always interruptible.  Otherwise: If both are negative, newly launched subroutines are
// never interruptible.  If only one is negative, newly launched subroutines cannot be interrupted
// due to that component, only the other one (which is now known to be positive otherwise the
// first rule of precedence would have applied).
{
	if (aIncrementThreadCount)
		++g_nThreads; // It is the caller's responsibility to avoid calling us if the thread count is too high.

	// v1.0.38.04: mLinesExecutedThisCycle is now reset in this function for maintainability. For simplicity,
	// the reset is unconditional because it is desirable 99% of the time.
	// See comments in CheckScriptTimers() for why g_script.mLastScriptRest isn't altered here.
	g_script.mLinesExecutedThisCycle = 0; // Make it start fresh to avoid unnecessary delays due to SetBatchLines.

	bool underlying_thread_is_paused = g.IsPaused;  // Indicate in the upcoming thread whether the thread to be interrupted is paused.
	CopyMemory(&g, &g_default, sizeof(global_struct));
	g.UnderlyingThreadIsPaused = underlying_thread_is_paused;
	g.Priority = aPriority;

	if (aSkipUninterruptible)
		return;

	// v1.0.38.04: Mark the thread critical here, for maintainability and also to avoid doing
	// SET_UNINTERRUPTIBLE_TIMER when the script has taken explicit control of this thread's
	// interruptibility.
	// OLDER (somewhat obsolete but good background): There is no need to set g.ThreadIsCritical here
	// because when actually encountered, "Critical [On]" will do it (and we're sure that first line
	// will execute because: 1) g.AllowThreadToBeInterrupted was set to false above; and 2) the
	// the uinterruptible timer will not be activated. Thus, the thread cannot be interrupted by
	// anything short of OnExit/OnMessage, in which case g.ThreadIsCritical's value doesn't matter,
	// and it will be set later when "Critical [On]" is actually encountered by the resumed thread.
	// Update: But there might be a tiny window of opportunity for a critical thread to get interrupted
	// by a non-emergency thread after it was previously interrupted by an emergency thread, therefore
	// the following line is here for maintainability and reliability:
	if (!g.ThreadIsCritical) // If the thread default isn't "critical", make this thread critical only if it's explicitly marked that way.
		g.ThreadIsCritical = (aTypeOfFirstLine == ACT_CRITICAL);
	//else it's already critical, so leave it that way until "Critical Off" (which may be the very first line) is encountered at runtime.

	if (g_script.mUninterruptibleTime && g_script.mUninterruptedLineCountMax // Both components must be non-zero to start off uninterruptible.
		|| g.ThreadIsCritical) // v1.0.38.04.
	{
		// Use g.AllowThreadToBeInterrupted vs. g_AllowInterruption in case g_AllowInterruption
		// just happens to have been set to true for some other reason (e.g. SendKeys()):
		g.AllowThreadToBeInterrupted = false;
		// v1.0.38.04: If the first line is "Critical" (even "Critical Off"), don't activate the timer
		// because "Critical [On]" means uninterruptible until it's turned off.  And "Critical Off"
		// means immediately interruptible.
		if (g_script.mUninterruptibleTime > 0  // Known to be either negative or positive (but not zero) at this point.
			&& !g.ThreadIsCritical)
			// It's much better to set a timer than have ExecUntil() watch for the time
			// to expire.  This is because it performs better, but more importantly
			// consider the case when ExecUntil() calls a WinWait, FileSetAttrib, or any
			// single command that might take a long time to execute.  While that command
			// is executing, ExecUntil() would be unable to keep an eye on things, thus
			// the script would stay uninterruptible for far longer than intended unless
			// many checks were added in various places, which would be cumbersome to
			// maintain.  By using a timer, the message loop (which is called frequently
			// by just about every long-running script command) will be able to make the
			// script interruptible again much closer to the desired moment in time.
			SET_UNINTERRUPTIBLE_TIMER
		// else if it's negative, it's considered to be infinite, so no timer need be set.
	}
}



void ResumeUnderlyingThread(global_struct *pSavedStruct, bool aKillInterruptibleTimer)
{
	--g_nThreads; // Below relies on this having been done early.

	if (aKillInterruptibleTimer)
		KILL_UNINTERRUPTIBLE_TIMER // g.AllowThreadToBeInterrupted is set later below, after the g struct has been restored.

	bool underlying_thread_is_paused = g.UnderlyingThreadIsPaused; // Done this way for performance (to avoid multiple indirections).
	CopyMemory(&g, pSavedStruct, sizeof(global_struct));
	g_ErrorLevel->Assign(g.ErrorLevel);

	// The below relies on the above having restored "g":
	if (g.IsPaused != underlying_thread_is_paused) // The thread to be resumed has been marked with a new state by the thread above.
	{
		if (g.IsPaused = underlying_thread_is_paused) // The thread to be resumed is being put into a paused state.
		{
			if (g_nThreads) // The script is not about to become idle (i.e. a non-idle thread is being paused now).
				++g_nPausedThreads;
			else
				g_IdleIsPaused = true;
		}
		else // The thread to be resumed is being unpaused.
		{
			if (g_nThreads) // The script is not about to become idle.
				--g_nPausedThreads;
			else // The script is about to become idle, so by definition the "idle thread" is being unpaused.
				g_IdleIsPaused = false;
		}
		CheckMenuItem(GetMenu(g_hWnd), ID_FILE_PAUSE, g.IsPaused ? MF_CHECKED : MF_UNCHECKED);
	}

	// If the thread to be resumed was paused and has not been unpaused above, it will automatically be
	// resumed in a paused state because when we return from this function, we should be returning to
	// an instance of ExecUntil() (our caller), which should be in a pause loop still.  Conversely,
	// if the thread to be resumed wasn't paused but was just paused above, the icon will be changed now
	// but the thread won't actually pause until ExecUntil() goes into its pause loop (which should be
	// immediately after the current command finishes, if execution is right in the middle of a command
	// due to the command having done a MsgSleep to allow a thread to interrupt).
	// Older comment: Always update the tray icon in case the paused state of the subroutine
	// we're about to resume is different from our previous paused state.  Do this even
	// when the macro is used by CheckScriptTimers(), which although it might not techically
	// need it, lends maintainability and peace of mind.
	// UPDATE: Doing "g.AllowThreadToBeInterrupted = true" seems like a good idea to be safe,
	// at least in the case where CheckScriptTimers() calls this macro at a time when there
	// is no thread other than the "idle thread" to resume.  A resumed thread should always
	// be interruptible anyway, since otherwise it couldn't have been interrupted in the
	// first place to get us here. UPDATE #2: Making it "true" is now also relied upon by MsgMonitor(),
	// which sometimes launches a new thread even when the current thread is interruptible.
	// UPDATE #3: In v1.0.38.04, make it reflect the state of ThreadIsCritical for cases where
	// a critical thread was interrupted by an OnExit or OnMessage thread.  Upon being resumed after
	// such an emergency interruption, a critical thread should be uninterruptible again.
	g_script.UpdateTrayIcon();
	g.AllowThreadToBeInterrupted = !g.ThreadIsCritical;
}



VOID CALLBACK MsgBoxTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime)
{
	// Unfortunately, it appears that MessageBox() will return zero rather
	// than AHK_TIMEOUT, specified below -- at least under WinXP.  This
	// makes it impossible to distinguish between a MessageBox() that's been
	// timed out (destroyed) by this function and one that couldn't be
	// created in the first place due to some other error.  But since
	// MessageBox() errors are rare, we assume that they timed out if
	// the MessageBox() returns 0.  UPDATE: Due to the fact that TimerProc()'s
	// are called via WM_TIMER messages in our msg queue, make sure that the
	// window really exists before calling EndDialog(), because if it doesn't,
	// chances are that EndDialog() was already called with another value.
	// UPDATE #2: Actually that isn't strictly needed because MessageBox()
	// ignores the AHK_TIMEOUT value we send here.  But it feels safer:
	if (IsWindow(hWnd))
		EndDialog(hWnd, AHK_TIMEOUT);
	KillTimer(hWnd, idEvent);
	// v1.0.33: The following was added to fix the fact that a MsgBox with only an OK button
	// does not acutally send back the code sent by EndDialog() above.  The HWND is checked
	// in case "g" is no longer the original thread due to another thread having interrupted it.
	// Consequently, MsgBox's with an OK button won't be 100% reliable with the timeout feature
	// if an interrupting thread is running at the time the box times out.  This is in the help
	// file as a known limitation:
	if (g.DialogHWND == hWnd) // Regardless of whether IsWindow() is true.
		g.MsgBoxTimedOut = true;
}



VOID CALLBACK AutoExecSectionTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime)
// See the comments in AutoHotkey.cpp for an explanation of this function.
{
	// Since this was called, it means the AutoExec section hasn't yet finished (otherwise
	// this timer would have been killed before we got here).  UPDATE: I don't think this is
	// necessarily true.  I think it's possible for the WM_TIMER msg (since even TimerProc()
	// timers use WM_TIMER msgs) to be still buffered in the queue even though its timer
	// has been killed (killing the timer does not auto-purge any pending messages for
	// that timer, and it is risky/problematic to try to do so manually).  Therefore, although
	// we kill the timer here, we also do a double check further below to make sure
	// the desired action hasn't already occurred.  Finally, the macro is used here because
	// it's possible that the timer has already been killed, so we don't want to risk any
	// problems that might arise from killing a non-existent timer (which this prevents):
	// The below also sets "g.AllowThreadToBeInterrupted = true".  Notes about this:
	// And since the AutoExecute is taking a long time (or might never complete), we now allow
	// interruptions such as hotkeys and timed subroutines. Use g.AllowThreadToBeInterrupted
	// vs. g_AllowInterruption in case commands in the AutoExecute section need exclusive use of
	// g_AllowInterruption (i.e. they might change its value to false and then back to true,
	// which would interfere with our use of that var):
	KILL_AUTOEXEC_TIMER

	// This is a double-check because it's possible for the WM_TIMER message to have
	// been received (thus calling this TimerProc() function) even though the timer
	// was already killed by AutoExecSection().  In that case, we don't want to update
	// the global defaults again because the g struct might have incorrect/unintended
	// values by now:
	if (!g_script.AutoExecSectionIsRunning)
		return;

	// Otherwise, it's still running (or paused). So update global DEFAULTS, which are for all threads launched in the future:
	CopyMemory(&g_default, &g, sizeof(global_struct));
	global_clear_state(g_default);  // Only clear g_default, not g.  This also ensures that IsPaused gets set to false in case it's true in "g".
}



VOID CALLBACK UninterruptibleTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime)
{
	// v1.0.38.04: Make AllowThreadToBeInterrupted conform to ThreadIsCritical. This is necessary
	// in the case where the timer was killed by the "Critical" command but not before a WM_TIMER message
	// got posted to our queue.
	g.AllowThreadToBeInterrupted = !g.ThreadIsCritical;
	// But unconditionally kill the timer since it should be killed even when g.ThreadIsCritical==true.
	// This is because any thread that uses the command "Critical" has taken charge of its own interruptibility.
	KILL_UNINTERRUPTIBLE_TIMER // Best to use the macro so that g_UninterruptibleTimerExists is reset to false.
}



VOID CALLBACK InputTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime)
{
	KILL_INPUT_TIMER
	g_input.status = INPUT_TIMED_OUT;
}
