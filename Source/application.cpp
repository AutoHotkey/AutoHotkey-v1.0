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

#include "stdafx.h" // pre-compiled headers
#include "application.h"
#include "globaldata.h" // for access to g_clip, the "g" global struct, etc.
#include "window.h" // for serveral MsgBox and window functions
#include "util.h" // for strlcpy()
#include "resources\resource.h"  // For ID_TRAY_OPEN.


static UINT_PTR g_MainTimerID = 0;
// Must be non-zero, so start at 1.  The first timers in the series are used by the
// MessageBoxes:
//enum OurTimers {TIMER_ID_MAIN = MAX_MSGBOXES + 2};  // +2 to give an extra margin of safety.

#define KILL_MAIN_TIMER \
if (g_MainTimerID && KillTimer(NULL, g_MainTimerID))\
	g_MainTimerID = 0;

// Realistically, SetTimer() called this way should never fail?  But the
// event loop can't function properly without it, at least when there
// are suspended subroutines.
// MSDN docs for SetTimer(): "Windows 2000/XP: If uElapse is less than 10,
// the timeout is set to 10."
#define SET_MAIN_TIMER \
if (!g_MainTimerID && !(g_MainTimerID = (BOOL)SetTimer(NULL, 0, SLEEP_INTERVAL, (TIMERPROC)NULL)))\
	g_script.ExitApp("SetTimer() unexpectedly failed.");



ResultType MsgSleep(int aSleepDuration, MessageMode aMode, bool aRestoreActiveWindow)
// Returns a non-meaningful value (so that it can return the result of something, thus
// effectively ignoring the result).  Callers should ignore it.  aSleepDuration can be
// zero to do a true Sleep(0), or less than 0 to avoid sleeping or waiting at all
// (i.e. messages are checked and if there are none, the function will return immediately).
// aMode is RETURN_AFTER_MESSAGES (default) or WAIT_FOR_MESSAGES.
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
	// I know of no way to simulate a Sleep(0), so for now we do this.
	// UPDATE: It's more urgent that messages be checked than for the Sleep
	// duration to be zero, so for now, just let zero be handled like any other.
	//if (aMode == RETURN_AFTER_MESSAGES && aSleepDuration == 0)
	//{
	//	Sleep(0);
	//	return OK;
	//}
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
	// will be lagged.  PeekMessage is probably almost as good, but it probably
	// only clears out any waiting keys prior to returning.

	// This var allows us to suspend the currently-running subroutine and run any
	// hotkey events waiting in the message queue (if there are more than one, they
	// will be executed in sequence prior to resuming the suspended subroutine).
	// Never static because we could be recursed (e.g. when one hotkey iterruptes
	// a hotkey that has already been interrupted) and each recursion layer should
	// have it's own value for this?:
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
	// other stuff that would interfere with the GUI activities of other threads)

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
	// intervals, don't worry about it:
	// Note: QueryPerformanceCounter() has very high overhead compared to GetTickCount():
	DWORD start_time = allow_early_return ? 0 : GetTickCount();

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
	// them all up-to-date):
	bool timer_is_needed = aSleepDuration > 0 && aMode == RETURN_AFTER_MESSAGES;
	if (timer_is_needed)
		SET_MAIN_TIMER

	// Remove any WM_TIMER messages that are already in our queue because we don't
	// want encounter them so early.  This is because most callers would want us to yield
	// at least a little bit of time to other processes that need it, rather
	// than returning early without our thread having been idle at all inside
	// GetMessage().  This is because GetMessage() will put our thread to sleep in
	// a way vastly more effective than Sleep()... a way that will wake up our thread
	// almost the instant a message arrives for us (except when other processe(s)
	// have the CPU maxed), or the instant any keyboard or mouse events occur if our keybd
	// or mouse hook is installed (but in that case, our thread is apparently used by
	// GetMessage() itself to call the hook functions without us having to do anything).
	// If there are other CPU-intensive processes using all of their timeslices, I think
	// it might be 20ms, 40ms, or more during which our thread is suspended by the OS's
	// round-robin scheduler.  That's one way (and there are probably others) that there
	// could be many WM_TIMER messages already the queue.
	PurgeTimerMessages();

	// Only used when aMode == RETURN_AFTER_MESSAGES:
	// True if the current subroutine was interrupted by another:
	bool was_interrupted = false;
	bool sleep0_was_done = false;
	bool empty_the_queue_via_peek = false;

	HWND fore_window;
	DWORD fore_pid;
	char fore_class_name[32];

	MSG msg;
	for (;;)
	{
		// The script is idle (doing nothing and having no dialogs or suspended subroutines)
		// whenever the aMode is the one set when we were first called by WinMain().  All
		// other callers call this function with the other mode, and the script should not
		// be considered idle in those cases:
		g_IsIdle = (aMode == WAIT_FOR_MESSAGES);

		if (aSleepDuration > 0 && !empty_the_queue_via_peek)
		{
			// Use GetMessage() whenever possible, rather than PeekMessage() or a technique such
			// MsgWaitForMultipleObjects() because it's the "engine" that passes all keyboard
			// and mouse events immediately to the low-level keyboard and mouse hooks
			// (if they're installed).  Otherwise, there's greater risk of keyboard/mouse lag.
			if (GetMessage(&msg, NULL, 0, 0) == -1) // -1 is an error, 0 means WM_QUIT
			{
				// This probably can't happen since the above GetMessage() is getting any
				// message belonging to a thread we already know exists (i.e. the one we're
				// in now):
				MsgBox("GetMessage() unexpectedly returned an error.  Press OK to continue running.");
				continue;
			}
			// else let any WM_QUIT be handled below.
		}
		else
		{
			// aSleepDuration <= 0 or "empty_the_queue_via_peek" is true, so we don't want
			// to be stuck in GetMessage() for even 10ms:
			if (!PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) // No more messages
			{
				// Since we're here, it means this recursion layer/instance of this
				// function didn't encounter any hotkey messages because if it had,
				// it would have already returned due to the WM_HOTKEY cases below.
				// So there should be no need to restore the value of global variables?
				if (aSleepDuration == 0 && !sleep0_was_done)
				{
					// Support a true Sleep(0) because it's the only way to yield
					// CPU time in this exact way.  It's used for things like
					// "SetKeyDelay, 0", which is defined as a Sleep(0) between
					// every keystroke
					// Out msg queue is empty, so do the sleep now, which might
					// yield the rest of our entire timeslice (probably 20ms
					// since we likely haven't used much of it) if the CPU is
					// under load:
					Sleep(0);
					sleep0_was_done = true;
					// Now start a new iteration of the loop that will see if we
					// received any messages during the up-to-20ms delay that
					// just occurred.  It's done this way to minimize keyboard/mouse
					// lag (if the hooks are installed) that will occur if any key or
					// mouse events are generated during that 20ms:
					continue;
				}
				else // aSleepDuration is non-zero or we already did the Sleep(0)
					// Function should always return OK in this case.  Also, was_interrupted
					// will always be false because if this "aSleepDuration <= 0" call
					// really was interrupted, it would already have returned in the
					// hotkey cases of the switch().  UPDATE: was_interrupted can now
					// be true since the hotkey case in the switch() doesn't return,
					// relying on us to do it after making sure the queue is empty:
					return IsCycleComplete(aSleepDuration, start_time, allow_early_return, was_interrupted);
			}
			// else Peek() found a message, so process it below.
		}

		switch(msg.message)
		{
		case WM_QUIT:
			// Any other cleanup needed before this?  If the app owns any windows,
			// they're cleanly destroyed upon termination?
			// Note: If PostQuitMessage() was called to generate this message,
			// no new dialogs (e.g. MessageBox) can be created (perhaps no new
			// windows of any kind):
			g_script.ExitApp();
			break;
		case WM_TIMER:
			if (aMode == WAIT_FOR_MESSAGES)
				// Timer should have already been killed if we're in this state.
				// But there might have been some WM_TIMER msgs already in the queue
				// (they aren't removed when the timer is killed).  Or perhaps
				// a caller is calling us with this aMode even though there
				// are suspended subroutines (currently never happens).
				// In any case, these are always ignored in this mode because
				// the caller doesn't want us to ever return:
				break;
			if (aSleepDuration <= 0)
				// This case is also handled specially above, so ignore WM_TIMER
				// messages, which might occur due to the hotkey case of the switch()
				// doing on final iteration in peek-mode (and perhaps in other ways):
				break;
			// Otherwise amode == RETURN_AFTER_MESSAGES:
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
			if (IsCycleComplete(aSleepDuration, start_time, allow_early_return, was_interrupted))
				return OK;
			// Otherwise, stay in the blessed GetMessage() state until
			// the time has expired:
			break;
		case AHK_HOOK_TEST_MSG:
		{
			char dlg_text[512];
			snprintf(dlg_text, sizeof(dlg_text), "TEST MSG: %d (0x%X)  %d (0x%X)"
				"\nCurrent Thread: 0x%X"
				, msg.wParam, msg.wParam, msg.lParam, msg.lParam
				, GetCurrentThreadId());
			MsgBox(dlg_text);
			break;
		}
		case WM_HOTKEY: // As a result of this app having previously called RegisterHotkey().
		case AHK_HOOK_HOTKEY:  // Sent from this app's keyboard or mouse hook.
			if (g_IgnoreHotkeys)
				break;
			if (g_nInterruptedSubroutines >= 10)
				// Allow only a limited number of recursion levels to avoid any chance of
				// stack overflow.  So ignore this message:
				break;
			if (aMode == RETURN_AFTER_MESSAGES)
			{
				was_interrupted = true;
				++g_nInterruptedSubroutines;
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
				g.hWndToRestore = NULL;
				// Also save the ErrorLevel of the subroutine that's about to be suspended.
				// Current limitation: If the user put something big in ErrorLevel (very unlikely
				// given its nature, but allowed) it will be truncated by this, if too large.
				// Also: Don't use var->Get() because need better control over the size:
				strlcpy(g.ErrorLevel, g_ErrorLevel->Contents(), sizeof(g.ErrorLevel));
				// Could also use copy constructor but that would probably incur more overhead?:
				CopyMemory(&global_saved, &g, sizeof(global_struct));
				// Next, change the value of globals to reflect the fact that we're about
				// to launch a new subroutine.
			}

			// Always kill the main timer, for performance reasons and for simplicity of design,
			// prior to embarking on new subroutine whose duration may be long (e.g. if BatchLines
			// is very high or infinite, the called subroutine may not return to us for seconds,
			// minutes, or more; during which time we don't want the timer running because it will
			// only fill up the queue with WM_TIMER messages and thus hurt performance.
			KILL_MAIN_TIMER;

			// Make every newly launched subroutine start off with the global default values that
			// the user set up in the auto-execute part of the script (e.g. KeyDelay, WinDelay, etc.).
			// However, we do not set ErrorLevel to NONE here because it's more flexible that way
			// (i.e. the user may want one hotkey subroutine to use the value of ErrorLevel set by another):
			CopyMemory(&g, &g_default, sizeof(global_struct));

			// Always reset these two, after the saving to global_saved and restoring to defaults above,
			// regardless of aMode:
			g.hWndLastUsed = NULL;
			g.StartTime = GetTickCount();
			g_script.mLinesExecutedThisCycle = 0;  // Doing this is somewhat debatable.
			g_IsIdle = false;  // Make sure the state is correct since we're about to launch a subroutine.

			// Just prior to launching the hotkey, update these values to support built-in
			// variables such as A_TimeSincePriorHotkey:
			g_script.mPriorHotkeyLabel = g_script.mThisHotkeyLabel;
			g_script.mPriorHotkeyStartTime = g_script.mThisHotkeyStartTime;
			g_script.mThisHotkeyLabel = Hotkey::GetLabel((HotkeyIDType)msg.wParam);
			g_script.mThisHotkeyStartTime = GetTickCount();

			// Perform the new hotkey's subroutine:
			if (Hotkey::PerformID((HotkeyIDType)msg.wParam) == CRITICAL_ERROR)
				// The above should have already displayed the error.  Rather than
				// calling PostQuitMessage(CRITICAL_ERROR) (in case the above didn't),
				// it seems best just to exit here directly, since there may be messages
				// in our queue waiting and we don't want to process them:
				g_script.ExitApp(NULL, CRITICAL_ERROR);
			g_LastPerformedHotkeyType = Hotkey::GetType((HotkeyIDType)msg.wParam); // For use with the Keylog cmd.

			if (aMode == RETURN_AFTER_MESSAGES)
			{
				if (g_nInterruptedSubroutines > 0) // Check just in case because never want it to go negative.
					--g_nInterruptedSubroutines;
				// Restore the global values for the subroutine that was interrupted:
				// Resume the original subroutine by returning, even if there
				// are still messages waiting in the queue.  Always return OK
				// to the caller because CRITICAL_ERROR was handled above:
				CopyMemory(&g, &global_saved, sizeof(global_struct));
				g_ErrorLevel->Assign(g.ErrorLevel);  // Restore the variable from the stored value.
				// Last param of below call tells it never to restore, because the final
				// loop iteration will do that if it's appropriate (i.e. avoid doing it
				// twice in case the final peek-iteration discovers a recursable action):
				if (g_UnpauseWhenResumed && g.IsPaused)
				{
					g_UnpauseWhenResumed = false;  // We've "used up" this unpause ticket.
					g.IsPaused = false;
					--g_nPausedSubroutines;
				}
				// else if it it's paused, it will automatically be resumed in a paused state
				// because when we return from this function, we should be returning to
				// an instance of ExecUntil() (our caller), which should be in a pause loop still.
				// But always update the tray icon in case the paused state of the subroutine
				// we're about to resume is different from our previous paused state:
				g_script.UpdateTrayIcon();

				if (IsCycleComplete(aSleepDuration, start_time, allow_early_return, was_interrupted, true))
				{
					// Check for messages once more in case the subroutine that just completed
					// above didn't check them that recently.  This is done to minimize the time
					// our thread spends *not* pumping messages, which in turn minimizes keyboard
					// and mouse lag if the hooks are installed.  Set the state of this function
					// layer/instance so that it will use peek-mode.  UPDATE: Don't change the
					// value of aSleepDuration to -1 because IsCycleComplete() needs to know the
					// original sleep time specified by the caller to determine whether
					// to restore the caller's active window after the caller's subroutine
					// is resumed-after-suspension:
					empty_the_queue_via_peek = true;
					allow_early_return = true;
					// Now it just falls out of the switch() and the loop does another iteration.
				}
				else
					if (timer_is_needed)  // Turn the timer back on if it is needed.
						SET_MAIN_TIMER
						// and stay in the blessed GetMessage() state until the time has expired.
			}
			break;
		default:
			// This is done here rather than as a case of the switch() because making
			// it a case would prevent all keyboard input from being dispatched, and we
			// need it to be dispatched for the cursor to be keyboard-controllable in
			// the edit window:
			if (msg.hwnd == g_hWndEdit && msg.message == WM_KEYDOWN && msg.wParam == VK_ESCAPE)
				// This won't work if a MessageBox() window is displayed because its own internal
				// message pump will dispatch the message to our edit control, which will just
				// ignore it.  And avoiding setting the focus to the edit control won't work either
				// because the user can simply click in the window to set the focus.  But for now,
				// this is better than nothing:
				ShowWindow(g_hWnd, SW_HIDE);  // And it's okay if this msg gets dispatched also.
			if (msg.hwnd == g_hWndEdit && msg.message == WM_KEYDOWN && msg.wParam == VK_F5)
				SendMessage(g_hWnd, WM_COMMAND, ID_TRAY_OPEN, 0);
			// This little part is from the Miranda source code.  But it doesn't seem
			// to provide any additional functionality: You still can't use keyboard
			// keys to navigate in the dialog unless it's the topmost dialog.
			// UPDATE: The reason it doens't work for non-topmost MessageBoxes is that
			// this message pump isn't even the one running.  It's the pump of the
			// top-most MessageBox itself, which apparently doesn't properly dispatch
			// all types of messages to other MessagesBoxes.  However, keeping this
			// here is probably a good idea because testing reveals that it does
			// sometimes receive messages intended for MessageBox windows (which makes
			// sense because our message pump here retrieves all thread messages).
			// It might cause problems to dispatch such messages directly, since
			// IsDialogMessage() is supposed to be used in lieu of DispatchMessage()
			// for these types of messages:
			if ((fore_window = GetForegroundWindow()) != NULL)  // There is a foreground window.
			{
				GetWindowThreadProcessId(fore_window, &fore_pid);
				if (fore_pid == GetCurrentProcessId())  // It belongs to our process.
				{
					GetClassName(fore_window, fore_class_name, sizeof(fore_class_name));
					if (!strcmp(fore_class_name, "#32770"))  // MessageBox(), InputBox(), or FileSelectFile() window.
						if (IsDialogMessage(fore_window, &msg))  // This message is for it, so let it process it.
							continue;  // This message is done.
				}
			}
			// Translate keyboard input for any of our thread's windows that need it:
			TranslateMessage(&msg);
			// This just routes the message to the WindowProc() of the
			// window the message was intended for (e.g. if our main window is visible
			// or if we have a dialog window... though I think IsDialog() is supposed
			// to be used for those, not Dispatch():
			DispatchMessage(&msg);
		} // switch()
	} // infinite-loop
}



ResultType IsCycleComplete(int aSleepDuration, DWORD aStartTime, bool aAllowEarlyReturn
	, bool aWasInterrupted, bool aNeverRestoreWnd)
// This function is used just to make MsgSleep() more readable/understandable.
{
	// Note: Even if TickCount has wrapped due to system being up more than about 49 days,
	// DWORD math still gives the right answer as long as aStartTime itself isn't more
	// than about 49 days ago. Note: must cast to int or any negative result will be lost
	// due to DWORD type:
	DWORD time_now = GetTickCount();
	if (!aAllowEarlyReturn && (int)(aSleepDuration - (time_now - aStartTime)) > SLEEP_INTERVAL_HALF)
		// Early return isn't allowed and the time remaining is large enough that we need to
		// wait some more (small amounts of remaining time can't be effectively waited for
		// due to the 10ms granularity limit of SetTimer):
		return FAIL; // Tell the caller to wait some more.

	// Reset counter for the caller of our caller, any time the thread
	// has had a chance to be idle (even if that idle time was done at a deeper
	// recursion level and not by this one), since the CPU will have been
	// given a rest, which is the main (perhaps only?) reason for using BatchLines
	// (e.g. to be more friendly toward time-critical apps such as games,
	// video capture, video playback):
	if (aSleepDuration >= 0)
		g_script.mLinesExecutedThisCycle = 0;

	// Normally only update g_script.mLastSleepTime when a non-zero MsgSleep() has
	// occurred.  This is because lesser sleeps never put us into the GetMessage()
	// state, and thus the keyboard & mouse hooks might cause key/mouse lag if
	// we fail call GetMessage() consistently every 10 to 50ms.
	if (aSleepDuration > 0)
		g_script.mLastSleepTime = time_now;

	// Kill the timer only if we're about to return OK to the caller since the caller
	// would still need the timer if FAIL was returned above.
	// Note: Since we're here, g_MainTimerID should be true.
	KILL_MAIN_TIMER

	// Check g.WaitingForDialog because don't want to restore the foreground
	// window for a subroutine that was waiting for user-input in a MessageBox
	// or InputBox window.  This is because that subroutine was obviously not
	// "working with" any window, and it's far better to avoid restoring when
	// there's any doubt:
	if (g.hWndToRestore != NULL && !aNeverRestoreWnd && aWasInterrupted
		&& aSleepDuration <= SLEEP_INTERVAL && !g.WaitingForDialog && !g_TrayMenuIsVisible)
	{
		// If the suspended subroutine was suspended during a very short rest period,
		// it probably didn't expect the active window to change.  So restore it
		// to the foreground.  This is a little bit iffy because the suspended
		// subroutine might not even have been using the active window, and restoring
		// it this way might then be incorrect.  However, one case it does make sense
		// (and there are probably others) is during a Send command that gets interrupted.
		// The old window should be reactivated to avoid the keystrokes going to the wrong
		// window.  I've tested it with this case and the Send command works well even when
		// interrupted by another hotkey that changes the active window.
		// But, this all needs more review to determine what would be correct most often to do:
		SetForegroundWindowEx(g.hWndToRestore);

		// e.g. In case window animation is on and it takes a while to actually be active.
		// Note: this is indirectly recursive, but doesn't seem capable of causing
		// an infinite loop & stack fault:
		DoWinDelay;
	}
	return OK;
}



VOID CALLBACK DialogTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime)
{
	// Unfortunately, it appears that MessageBox() will return zero rather
	// than AHK_TIMEOUT, specified below -- at least under WinXP.  This
	// makes it impossible to distinguish between a MessageBox() that's been
	// timed out (destroyed) by this function and one that couldn't be
	// created in the first place due to some other error.  But since
	// MessageBox() errors are rare, we assume that they timed out if
	// the MessageBox() returns 0:
	EndDialog(hWnd, AHK_TIMEOUT);
	KillTimer(hWnd, idEvent);
}
