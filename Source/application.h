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

#include <limits.h>
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

ResultType MsgSleep(int aSleepDuration = INTERVAL_UNSPECIFIED, MessageMode aMode = RETURN_AFTER_MESSAGES
	, bool aRestoreActiveWindow = true);

// This macro is used to Sleep without the possibility of a new hotkey subroutine being launched.
// It should be used when an operation needs to sleep, but doesn't want to be interrupted (suspended)
// by any hotkeys the user might press during that time.  Reasons why the caller wouldn't want to
// be suspended:
// 1) If it's doing something with a window -- such as sending keys or clicking the mouse or trying
//    to activate it -- that might get messed up if a new hotkey fires in the middle of the operation.
// 2) If its a command that's still using some of its parameters that might reside in the deref buffer.
//    In this case, the launching of a new hotkey would likely overwrite those values, causing
//    unpredictable behavior.
#define SLEEP_AND_IGNORE_HOTKEYS(aSleepTime) \
{\
	g_IgnoreHotkeys = true;\
	MsgSleep(aSleepTime);\
	g_IgnoreHotkeys = false;\
}
#define DoWinDelay if (g.WinDelay >= 0) MsgSleep(g.WinDelay)
#define DoControlDelay if (g.ControlDelay >= 0) MsgSleep(g.ControlDelay)

ResultType IsCycleComplete(int aSleepDuration, DWORD aStartTime, bool aAllowEarlyReturn
	, bool aWasInterrupted, bool aNeverRestoreWnd = false);

VOID CALLBACK DialogTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);

inline void PurgeTimerMessages()
// Maybe best to have it an inline in case the caller also has declared var named "msg".
{
	MSG msg;
	while (PeekMessage(&msg, NULL, WM_TIMER, WM_TIMER, PM_REMOVE)); // self-contained loop.
	// There should be no danger of the above accidentally removing the MsgBox() timeout
	// timers, since those timers call a function rather than post a message to our queue.
}

#endif
