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

#ifndef application_h
#define application_h

#include "defines.h"

// 9 might be better than 10 because if the granularity/timer is a little
// off on certain systems, a Sleep(10) might really result in a Sleep(20),
// whereas a Sleep(9) is almost certainly a Sleep(10) on OS's such as
// NT/2k/XP.  UPDATE: Roundoff issues with scripts having
// even multiples of 10 in them, such as "Sleep,300", shouldn't be hurt
// by this because they use GetTickCount() to verify how long the
// sleep duration actually was.  UPDATE again: Decided to go back to 10
// because I'm pretty confident that that always sleeps 10 on NT/2k/XP
// unless the system is under load, in which case any Sleep between 0
// and 20 inclusive seems to sleep for exactly(?) one timeslice.
// A timeslice appears to be 20ms in duration.  Anyway, using 10
// allows "SetKeyDelay, 10" to be really 10 rather than getting
// rounded up to 20 due to doing first a Sleep(10) and then a Sleep(1).
// For now, I'm avoiding using timeBeginPeriod to improve the resolution
// of Sleep() because of possible incompatibilities on some systems,
// and also because it may degrade overall system performance.
// UPDATE: Will get rounded up to 10 anyway by SetTimer().  However,
// future OSs might support timer intervals of less than 10.
#define SLEEP_INTERVAL 10
#define SLEEP_INTERVAL_HALF (int)(SLEEP_INTERVAL / 2)

// Use some negative value unlikely to ever be passed explicitly:
#define INTERVAL_UNSPECIFIED (INT_MIN + 303)
#define NO_SLEEP -1
enum MessageMode {RETURN_AFTER_MESSAGES, WAIT_FOR_MESSAGES};

// The first timers in the series are used by the MessageBoxes.  Start at +2 to give
// an extra margin of safety:
enum OurTimers {TIMER_ID_MAIN = MAX_MSGBOXES + 2, TIMER_ID_UNINTERRUPTIBLE, TIMER_ID_AUTOEXEC
	, TIMER_ID_INPUT, TIMER_ID_DEREF};

// MUST MAKE main timer and uninterruptible timers associated with our main window so that
// MainWindowProc() will be able to process them when it is called by the DispatchMessage()
// of a non-standard message pump such as MessageBox().  In other words, don't let the fact
// that the script is displaying a dialog interfere with the timely receipt and processing
// of the WM_TIMER messages, including those "hidden messages" which cause DefWindowProc()
// (I think) to call the TimerProc() of timers that use that method.
// Realistically, SetTimer() called this way should never fail?  But the event loop can't
// function properly without it, at least when there are suspended subroutines.
// MSDN docs for SetTimer(): "Windows 2000/XP: If uElapse is less than 10,
// the timeout is set to 10."
#define SET_MAIN_TIMER \
if (!g_MainTimerExists && !(g_MainTimerExists = SetTimer(g_hWnd, TIMER_ID_MAIN, SLEEP_INTERVAL, (TIMERPROC)NULL)))\
	g_script.ExitApp(EXIT_CRITICAL, "SetTimer"); // Just a brief msg to cut down on mem overhead, since it should basically never happen.

// When someone calls SET_UNINTERRUPTIBLE_TIMER, by definition the current script subroutine is
// becoming non-interruptible.  Therefore, their should never be a need to have more than one
// of these timer going simultanously since they're only created for the launch of a new
// quasi-thread, which is forbidden when the current thread is uninterruptible.
// Remember than the 2nd param of SetTimer() is ignored when the 1st param is NULL.
// For this one, the timer is not recreated if it already exists because I think SetTimer(), when
// called with NULL as a first parameter, may wind up creating more than one Timer and we only
// want one of these to exist at a time.
// The caller should ensure that aTimeoutValue is <= INT_MAX because otherwise SetTimer()'s behavior
// will vary depending on OS type & version.
// Also have this one abort on unexpected error, since failure to set the timer might result in the
// script becoming permanently uninterruptible (which prevents new hotkeys from being activated
// even though the program is still responsive).
#define SET_UNINTERRUPTIBLE_TIMER \
if (!g_UninterruptibleTimerExists && !(g_UninterruptibleTimerExists = SetTimer(g_hWnd, TIMER_ID_UNINTERRUPTIBLE \
	, g_script.mUninterruptibleTime < 10 ? 10 : g_script.mUninterruptibleTime, UninteruptibleTimeout)))\
	g_script.ExitApp(EXIT_CRITICAL, "SetTimer() unexpectedly failed.");

// See AutoExecSectionTimeout() for why g.AllowThisThreadToBeInterrupted is used rather than the other var.
// Also, from MSDN: "When you specify a TimerProc callback function, the default window procedure calls the
// callback function when it processes WM_TIMER. Therefore, you need to dispatch messages in the calling thread,
// even when you use TimerProc instead of processing WM_TIMER."  My: This is why all TimerProc type timers
// should probably have a window rather than passing NULL as first param of SetTimer().
#define SET_AUTOEXEC_TIMER(aTimeoutValue) \
{\
	g.AllowThisThreadToBeInterrupted = false;\
	if (!g_AutoExecTimerExists && !(g_AutoExecTimerExists = SetTimer(g_hWnd, TIMER_ID_AUTOEXEC, aTimeoutValue, AutoExecSectionTimeout)))\
		g_script.ExitApp(EXIT_CRITICAL, "SetTimer() unexpectedly failed.");\
}

#define SET_INPUT_TIMER(aTimeoutValue) \
	if (!g_InputTimerExists)\
		g_InputTimerExists = SetTimer(g_hWnd, TIMER_ID_INPUT, aTimeoutValue, InputTimeout);

// For this one, SetTimer() is called unconditionally because our caller wants the timer reset
// (as though it were killed and recreated) uncondtionally.  MSDN's comments are a little vague
// about this, but testing shows that calling SetTimer() against an existing timer does completely
// reset it as though it were killed and recreated.  Note also that g_hWnd is used vs. NULL so that
// the timer will fire even when a msg pump other than our own is running, such as that of a MsgBox.
#define SET_DEREF_TIMER(aTimeoutValue) g_DerefTimerExists = SetTimer(g_hWnd, TIMER_ID_DEREF, aTimeoutValue, DerefTimeout);

#define KILL_MAIN_TIMER \
if (g_MainTimerExists && KillTimer(g_hWnd, TIMER_ID_MAIN))\
	g_MainTimerExists = false;

// Although the caller doesn't always need g.AllowThisThreadToBeInterrupted reset to true,
// it's much more maintainable and nicer to do it unconditionally due to the complexity of
// managing quasi-threads.  At the very least, it's needed for when the "idle thread"
// is "resumed" (see MsgSleep for explanation).
#define MAKE_THREAD_INTERRUPTIBLE \
{\
	g.AllowThisThreadToBeInterrupted = true;\
	if (g_UninterruptibleTimerExists && KillTimer(g_hWnd, TIMER_ID_UNINTERRUPTIBLE))\
		g_UninterruptibleTimerExists = false;\
}

// Notes about the below macro:
// If our thread's message queue has any message pending whose HWND member is NULL -- or even
// normal messages which would be routed back to the thread by the WindowProc() -- clean them
// out of the message queue before launching the dialog's message pump below.  That message pump
// doesn't know how to properly handle such messages (it would either lose them or dispatch them
// at times we don't want them dispatched).  But first ensure the current quasi-thread is
// interruptible, since it's about to display a dialog so there little benefit (and a high cost)
// to leaving it uninterruptible.  The "high cost" is that MsgSleep (our main message pump) would
// filter out (leave queued) certain messages if the script were uninterruptible.  Then when it
// returned, the dialog message pump below would start, and it would discard or misroute the
// messages.
// If this is not done, the following scenario would also be a problem:
// A newly launched thread (in its period of uninterruptibility) displays a dialog.  As a consequence,
// the dialog's message pump starts dispatching all messages.  If during this brief time (before the
// thread becomes interruptible) a hotkey/hotstring/custom menu item/gui thread is dispatched to one
// of our WindowProc's, and then posted to our thread via PostMessage(NULL,...), the item would be lost
// because the dialog message pump discards messages that lack an HWND (since it doesn't know how to
// dispatch them to a Window Proc).
// GetQueueStatus() is used because unlike PeekMessage() or GetMessage(), it might not yield
// our timeslice if the CPU is under heavy load, which would be good to improve performance here.
#define DIALOG_PREP \
{\
	MAKE_THREAD_INTERRUPTIBLE \
	if (HIWORD(GetQueueStatus(QS_ALLEVENTS)))\
		MsgSleep(-1);\
}


// See above comment about g.AllowThisThreadToBeInterrupted.
// Also, must restore to true in this case since auto-exec section isn't run as a new thread
// (i.e. there's nothing to resume).
#define KILL_AUTOEXEC_TIMER \
{\
	g.AllowThisThreadToBeInterrupted = true;\
	if (g_AutoExecTimerExists && KillTimer(g_hWnd, TIMER_ID_AUTOEXEC))\
		g_AutoExecTimerExists = false;\
}

#define KILL_INPUT_TIMER \
if (g_InputTimerExists && KillTimer(g_hWnd, TIMER_ID_INPUT))\
	g_InputTimerExists = false;

#define KILL_DEREF_TIMER \
if (g_DerefTimerExists && KillTimer(g_hWnd, TIMER_ID_DEREF))\
	g_DerefTimerExists = false;

// Callers should note that using INTERVAL_UNSPECIFIED might not rest the CPU at all if there is
// already at least one msg waiting in our thread's msg queue:
ResultType MsgSleep(int aSleepDuration = INTERVAL_UNSPECIFIED, MessageMode aMode = RETURN_AFTER_MESSAGES);

// This macro is used to Sleep without the possibility of a new hotkey subroutine being launched.
// Timed subroutines will also be prevented from running while it is enabled.
// It should be used when an operation needs to sleep, but doesn't want to be interrupted (suspended)
// by any hotkeys the user might press during that time.  Reasons why the caller wouldn't want to
// be suspended:
// 1) If it's doing something with a window -- such as sending keys or clicking the mouse or trying
//    to activate it -- that might get messed up if a new hotkey fires in the middle of the operation.
// 2) If its a command that's still using some of its parameters that might reside in the deref buffer.
//    In this case, the launching of a new hotkey would likely overwrite those values, causing
//    unpredictable behavior.
#define SLEEP_WITHOUT_INTERRUPTION(aSleepTime) \
{\
	g_AllowInterruption = false;\
	MsgSleep(aSleepTime);\
	g_AllowInterruption = true;\
}

// Whether we should allow the script's current quasi-thread to be interrupted by
// either a newly pressed hotkey or a timed subroutine that is due to be run.
// Note that the 2 variables used here are independent of each other to support
// the case where an uninterruptible operation such as SendKeys() happens to occur
// while g.AllowThisThreadToBeInterrupted is true, in which case we would want the
// completion of that operation to affect only the status of g_AllowInterruption,
// not g.AllowThisThreadToBeInterrupted.
#define INTERRUPTIBLE (g.AllowThisThreadToBeInterrupted && g_AllowInterruption && !g_MenuIsVisible)

// To reduce the expectation that a newly launched hotkey or timed subroutine will
// be immediately interrupted by a timed subroutine or hotkey, interruptions are
// forbidden for a short time (user-configurable).  If the subroutine is a quick one --
// finishing prior to when PerformID()'s call of ExecUntil() or the Timer would have set
// g_AllowInterruption to be true -- we will set it to be true afterward so that it
// gets done as quickly as possible.
// The following rules of precedence apply:
// If either UninterruptibleTime or UninterruptedLineCountMax is zero, newly launched subroutines
// are always interruptible.  Otherwise: If both are negative, newly launched subroutines are
// never interruptible.  If only one is negative, newly launched subroutines cannot be interrupted
// due to that component, only the other one (which is now known to be positive otherwise the
// first rule of precedence would have applied).
// Notes that apply to the macro:
// Both must be non-zero.
// ...
// Use g.AllowThisThreadToBeInterrupted vs. g_AllowInterruption in case g_AllowInterruption
// just happens to have been set to true for some other reason (e.g. SendKeys()).
// ...
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
// ...
// Known to be either negative or positive (but not zero) at this point.
// ...
// else if it's negative, it's considered to be infinite, so no timer need be set.

#define INIT_NEW_THREAD \
CopyMemory(&g, &g_default, sizeof(global_struct));\
if (g_script.mUninterruptibleTime && g_script.mUninterruptedLineCountMax)\
{\
	g.AllowThisThreadToBeInterrupted = false;\
	if (g_script.mUninterruptibleTime > 0)\
		SET_UNINTERRUPTIBLE_TIMER \
}

// The DISABLE_UNINTERRUPTIBLE_SUB macro below must always kill the timer if it exists -- even if
// the timer hasn't expired yet.  This is because if the timer were to fire when interruptibility had
// already been previously restored, it's possible that it would set g.AllowThisThreadToBeInterrupted
// to be true even when some other code had had the opporutunity to set it to false by intent.
// In other words, once g.AllowThisThreadToBeInterrupted is set to true the first time, it should not be
// set a second time "just to be sure" because by then it may already by in use by someone else
// for some other purpose.
// It's possible for the SetBatchLines command to have changed the values of g_script.mUninterruptibleTime
// and g_script.mUninterruptedLineCountMax since the time INIT_NEW_THREAD was called.  If they were
// changed so that subroutines are always interruptible, that seems to be handled correctly.
// If they were changed so that subroutines are never interruptible, that seems to be okay too.
// It doesn't seem like any combination of starting vs. ending interruptibility is a particular
// problem, so no special handling is done here (just keep it simple).
// UPDATE: g.AllowThisThreadToBeInterrupted is always made true even if both settings are negative,
// since our callers would all want us to do it unconditionally.  This is because there's no need to
// keep it false even when all subroutines are permanently uninterruptible, since it will be made
// false every time a new subroutine launches.
// Macro notes:
// Reset in case the timer hasn't yet expired and PerformID()'s call of ExecUntil()
// didn't get a chance to do it:
// ...
// Since this timer is of the type that calls a function directly, rather than placing
// msgs in our msg queue, it should not be necessary to worry about removing its messages
// from the msg queue.
// UPDATE: The below is now the same as KILL_UNINTERRUPTIBLE_TIMER because all of its callers
// have finished running the current thread (i.e. the current thread is about to be destroyed):
//#define DISABLE_UNINTERRUPTIBLE_SUB \
//{\
//	g.AllowThisThreadToBeInterrupted = true;\
//	KILL_UNINTERRUPTIBLE_TIMER \
//}
//#define DISABLE_UNINTERRUPTIBLE_SUB	KILL_UNINTERRUPTIBLE_TIMER


// The unpause logic is done immediately after the most recently suspended thread's
// global settings are restored so that that thread is set up properly to be resumed.
// Comments about macro:
//    g_UnpauseWhenResumed = false --> because we've "used up" this unpause ticket.
//    g_ErrorLevel->Assign(g.ErrorLevel) --> restores the variable from the stored value.
// If the thread to be resumed has not been unpaused, it will automatically be resumed in
// a paused state because when we return from this function, we should be returning to
// an instance of ExecUntil() (our caller), which should be in a pause loop still.
// But always update the tray icon in case the paused state of the subroutine
// we're about to resume is different from our previous paused state.  Do this even
// when the macro is used by CheckScriptTimers(), which although it might not techically
// need it, lends maintainability and peace of mind.
// UPDATE: Doing "g.AllowThisThreadToBeInterrupted = true" seems like a good idea to be safe,
// at least in the case where CheckScriptTimers() calls this macro at a time when there
// is no thread other than the "idle thread" to resume.  A resumed thread should always
// be interruptible anyway, since otherwise it couldn't have been interrupted in the
// first place to get us here:
#define RESUME_UNDERLYING_THREAD \
{\
	CopyMemory(&g, &global_saved, sizeof(global_struct));\
	g_ErrorLevel->Assign(g.ErrorLevel);\
	if (g_UnpauseWhenResumed && g.IsPaused)\
	{\
		g_UnpauseWhenResumed = false;\
		g.IsPaused = false;\
		--g_nPausedThreads;\
		CheckMenuItem(GetMenu(g_hWnd), ID_FILE_PAUSE, MF_UNCHECKED);\
	}\
	g_script.UpdateTrayIcon();\
	g.AllowThisThreadToBeInterrupted = true;\
}


// Have this be dynamically resolved each time.  For example, when MsgSleep() uses this
// while in mode WAIT_FOR_MESSSAGES, its msg loop should use this macro in case the
// value of g_AllowInterruption changes from one iteration to the next.  Thankfully,
// MS made WM_HOTKEY have a very high value, so filtering in this way should not exclude
// any other important types of messages:
#define MSG_FILTER_MAX (INTERRUPTIBLE ? 0 : WM_HOTKEY - 1)

// Do a true Sleep() for short sleeps on Win9x because it is much more accurate than the MsgSleep()
// method on that OS, at least for when short sleeps are done on Win98SE:
#define DoWinDelay \
	if (g.WinDelay >= 0)\
	{\
		if (g.WinDelay < 25 && g_os.IsWin9x())\
			Sleep(g.WinDelay);\
		else\
			MsgSleep(g.WinDelay);\
	}

#define DoControlDelay \
	if (g.ControlDelay >= 0)\
	{\
		if (g.ControlDelay < 25 && g_os.IsWin9x())\
			Sleep(g.ControlDelay);\
		else\
			MsgSleep(g.ControlDelay);\
	}

ResultType IsCycleComplete(int aSleepDuration, DWORD aStartTime, bool aAllowEarlyReturn);

// These should only be called from MsgSleep() (or something called by MsgSleep()) because
// we don't want to be in the situation where a thread launched by CheckScriptTimers() returns
// first to a dialog's message pump rather than MsgSleep's pump.  That's because our thread
// might then have queued messages that would be stuck in the queue (due to the possible absence
// of the main timer) until the dialog's msg pump ended.
void CheckScriptTimers();
#define CHECK_SCRIPT_TIMERS_IF_NEEDED if (g_script.mTimerEnabledCount) CheckScriptTimers();

void PollJoysticks();
#define POLL_JOYSTICK_IF_NEEDED if (Hotkey::sJoyHotkeyCount) PollJoysticks();

VOID CALLBACK MsgBoxTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);
VOID CALLBACK AutoExecSectionTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);
VOID CALLBACK UninteruptibleTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);
VOID CALLBACK InputTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);
VOID CALLBACK DerefTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);

#endif
