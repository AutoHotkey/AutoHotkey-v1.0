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
enum OurTimers {TIMER_ID_MAIN = MAX_MSGBOXES + 2, TIMER_ID_UNINTERRUPTIBLE, TIMER_ID_AUTOEXEC, TIMER_ID_INPUT};

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
	g_script.ExitApp("SetTimer"); // Just a brief msg to cut down on mem overhead, since it should basically never happen.

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
#define SET_UNINTERRUPTIBLE_TIMER(aTimeoutValue) \
if (!g_UninterruptibleTimerExists && !(g_UninterruptibleTimerExists = SetTimer(g_hWnd, TIMER_ID_UNINTERRUPTIBLE, aTimeoutValue, UninteruptibleTimeout)))\
	g_script.ExitApp("SetTimer() unexpectedly failed.");

#define SET_AUTOEXEC_TIMER(aTimeoutValue) \
if (!g_AutoExecTimerExists && !(g_AutoExecTimerExists = SetTimer(g_hWnd, TIMER_ID_AUTOEXEC, aTimeoutValue, AutoExecSectionTimeout)))\
	g_script.ExitApp("SetTimer() unexpectedly failed.");

#define SET_INPUT_TIMER(aTimeoutValue) \
	if (!g_InputTimerExists)\
		g_InputTimerExists = SetTimer(g_hWnd, TIMER_ID_INPUT, aTimeoutValue, InputTimeout);

#define KILL_MAIN_TIMER \
if (g_MainTimerExists && KillTimer(g_hWnd, TIMER_ID_MAIN))\
	g_MainTimerExists = false;

#define KILL_UNINTERRUPTIBLE_TIMER \
if (g_UninterruptibleTimerExists && KillTimer(g_hWnd, TIMER_ID_UNINTERRUPTIBLE))\
	g_UninterruptibleTimerExists = false;

#define KILL_AUTOEXEC_TIMER \
if (g_AutoExecTimerExists && KillTimer(g_hWnd, TIMER_ID_AUTOEXEC))\
	g_AutoExecTimerExists = false;

#define KILL_INPUT_TIMER \
if (g_InputTimerExists && KillTimer(g_hWnd, TIMER_ID_INPUT))\
	g_InputTimerExists = false;


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
// while g_AllowInterruptionForSub is true, in which case we would want the completion
// of that operation to affect only the status of g_AllowInterruption, not
// g_AllowInterruptionForSub.
#define INTERRUPTIBLE (g_AllowInterruptionForSub && g_AllowInterruption && !g_MenuIsVisible)

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
// Use g_AllowInterruptionForSub vs. g_AllowInterruption in case g_AllowInterruption
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
#define ENABLE_UNINTERRUPTIBLE_SUB \
if (g.UninterruptibleTime && g.UninterruptedLineCountMax)\
{\
	g_AllowInterruptionForSub = false;\
	g_script.mUninterruptedLineCount = 0;\
	if (g.UninterruptibleTime > 0)\
		SET_UNINTERRUPTIBLE_TIMER(g.UninterruptibleTime < 10 ? 10 : g.UninterruptibleTime) \
}

// The DISABLE_UNINTERRUPTIBLE_SUB macro below must always kill the timer if it exists -- even if
// the timer hasn't expired yet.  This is because if the timer were to fire when interruptibility
// had already been previously restored, it's possible that it would set g_AllowInterruptionForSub
// to be true even when some other code had had the opporutunity to set it to false by intent.
// In other words, once g_AllowInterruptionForSub is set to true the first time, it should not be
// set a second time "just to be sure" because by then it may already by in use by someone else
// for some other purpose.
// It's possible for the SetBatchLines command to have changed the values of g.UninterruptibleTime
// and g.UninterruptedLineCountMax since the time ENABLE_UNINTERRUPTIBLE_SUB was called.  If they were
// changed so that subroutines are always interruptible, that seems to be handled correctly.
// If they were changed so that subroutines are never interruptible, that seems to be okay too.
// It doesn't seem like any combination of starting vs. ending interruptibility is a particular
// problem, so no special handling is done here (just keep it simple).
// UPDATE: g_AllowInterruptionForSub is always made true even if both settings are negative, since
// our callers would all want us to do it unconditionally.  This is because there's no need to
// keep it false even when all subroutines are permanently uninterruptible, since it will be made
// false every time a new subroutine launches.
// Macro notes:
// Reset in case the timer hasn't yet expired and PerformID()'s call of ExecUntil()
// didn't get a chance to do it:
// ...
// Since this timer is of the type that calls a function directly, rather than placing
// msgs in our msg queue, it should not be necessary to worry about removing its messages
// from the msg queue.
#define DISABLE_UNINTERRUPTIBLE_SUB \
{\
	g_AllowInterruptionForSub = true;\
	KILL_UNINTERRUPTIBLE_TIMER \
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

void CheckScriptTimers();
#define CHECK_SCRIPT_TIMERS_IF_NEEDED if (g_script.mTimerEnabledCount) CheckScriptTimers();

void PollJoysticks();
#define POLL_JOYSTICK_IF_NEEDED if (Hotkey::sJoyHotkeyCount) PollJoysticks();

VOID CALLBACK MsgBoxTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);
VOID CALLBACK AutoExecSectionTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);
VOID CALLBACK UninteruptibleTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);
VOID CALLBACK InputTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);

#endif
