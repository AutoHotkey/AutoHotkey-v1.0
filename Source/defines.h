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

#ifndef defines_h
#define defines_h

#include "stdafx.h" // pre-compiled headers

// Disable silly performance warning about converting int to bool:
// Unlike other typecasts from a larger type to a smaller, I'm 99% sure
// that all compilers are supposed to do something special for bool,
// not just truncate.  Example:
// bool x = 0xF0000000
// The above should give the value "true" to x, not false which is
// what would happen if:
// char x = 0xF0000000
//
#ifdef _MSC_VER
#pragma warning(disable:4800)
#endif

#define NAME_P "AutoHotkey"
#define NAME_VERSION "1.0.16"
#define NAME_PV NAME_P " v" NAME_VERSION

// Window class names: Changing these may result in new versions not being able to detect any old instances
// that may be running (such as the use of FindWindow() in WinMain()).  It may also have other unwanted
// effects, such as anything in the OS that relies on the class name that the user may have changed the
// settings for, such as whether to hide the tray icon (though it probably doesn't use the class name
// in that example).
// MSDN: "Because window classes are process specific, window class names need to be unique only within
// the same process. Also, because class names occupy space in the system's private atom table, you
// should keep class name strings as short a possible:
#define WINDOW_CLASS_MAIN "AutoHotkey"
#define WINDOW_CLASS_SPLASH "AutoHotkey2"

#define EXT_AUTOIT2 ".aut"
#define EXT_AUTOHOTKEY ".ahk"
#define CONVERSION_FLAG (EXT_AUTOIT2 EXT_AUTOHOTKEY)
#define CONVERSION_FLAG_LENGTH 8

// Items that may be needed for VC++ 6.X:
#ifndef SPI_GETFOREGROUNDLOCKTIMEOUT
	#define SPI_GETFOREGROUNDLOCKTIMEOUT        0x2000
	#define SPI_SETFOREGROUNDLOCKTIMEOUT        0x2001
#endif
#ifndef VK_XBUTTON1
	#define VK_XBUTTON1       0x05    /* NOT contiguous with L & RBUTTON */
	#define VK_XBUTTON2       0x06    /* NOT contiguous with L & RBUTTON */
	#define WM_NCXBUTTONDOWN                0x00AB
	#define WM_NCXBUTTONUP                  0x00AC
	#define WM_NCXBUTTONDBLCLK              0x00AD
	#define GET_WHEEL_DELTA_WPARAM(wParam)  ((short)HIWORD(wParam))
	#define WM_XBUTTONDOWN                  0x020B
	#define WM_XBUTTONUP                    0x020C
	#define WM_XBUTTONDBLCLK                0x020D
	#define GET_KEYSTATE_WPARAM(wParam)     (LOWORD(wParam))
	#define GET_NCHITTEST_WPARAM(wParam)    ((short)LOWORD(wParam))
	#define GET_XBUTTON_WPARAM(wParam)      (HIWORD(wParam))
	/* XButton values are WORD flags */
	#define XBUTTON1      0x0001
	#define XBUTTON2      0x0002
#endif
#ifndef HIMETRIC_INCH
	#define HIMETRIC_INCH 2540
#endif


#define IS_32BIT(signed_value_64) (signed_value_64 >= INT_MIN && signed_value_64 <= INT_MAX)
#define GET_BIT(buf,n) (((buf) & (1 << (n))) >> (n))
#define SET_BIT(buf,n,val) ((val) ? ((buf) |= (1<<(n))) : (buf &= ~(1<<(n))))

// FAIL = 0 to remind that FAIL should have the value zero instead of something arbitrary
// because some callers may simply evaluate the return result as true or false
// (and false is a failure):
enum ResultType {FAIL = 0, OK, WARN = OK, CRITICAL_ERROR
	, CONDITION_TRUE, CONDITION_FALSE
	, LOOP_BREAK, LOOP_CONTINUE
	, EARLY_RETURN, EARLY_EXIT};

enum ExitReasons {EXIT_NONE, EXIT_CRITICAL, EXIT_ERROR, EXIT_DESTROY, EXIT_LOGOFF, EXIT_SHUTDOWN
	, EXIT_WM_QUIT, EXIT_WM_CLOSE, EXIT_MENU, EXIT_EXIT, EXIT_RELOAD, EXIT_SINGLEINSTANCE};

enum SingleInstanceType {ALLOW_MULTI_INSTANCE, SINGLE_INSTANCE, SINGLE_INSTANCE_REPLACE
	, SINGLE_INSTANCE_IGNORE, SINGLE_INSTANCE_OFF}; // ALLOW_MULTI_INSTANCE must be zero.

enum MenuVisibleType {MENU_VISIBLE_NONE, MENU_VISIBLE_POPUP, MENU_VISIBLE_MAIN}; // NONE must be zero.

// These are used for things that can be turned on, off, or left at a
// neutral default value that is neither on nor off.  INVALID must
// be zero:
enum ToggleValueType {TOGGLE_INVALID = 0, TOGGLED_ON, TOGGLED_OFF, ALWAYS_ON, ALWAYS_OFF
	, TOGGLE, TOGGLE_PERMIT, NEUTRAL};

// For convenience in many places.  Must cast to int to avoid loss of negative values.
#define BUF_SPACE_REMAINING ((int)(aBufSize - (aBuf - aBuf_orig)))

// MsgBox timeout value.  This can't be zero because that is used as a failure indicator:
// Also, this define is in this file to prevent problems with mutual
// dependency between script.h and window.h.  Update: It can't be -1 either because
// that value is used to indicate failure by DialogBox():
#define AHK_TIMEOUT -2
// And these to prevent mutual dependency problem between window.h and globaldata.h:
#define MAX_MSGBOXES 7
#define MAX_INPUTBOXES 4
#define MAX_PROGRESS_WINDOWS 10  // Allow a lot for downloads and such.
#define MAX_PROGRESS_WINDOWS_STR "10" // Keep this in sync with above.
#define MAX_SPLASHIMAGE_WINDOWS 10
#define MAX_SPLASHIMAGE_WINDOWS_STR "10" // Keep this in sync with above.
#define MAX_TOOLTIPS 20
#define MAX_TOOLTIPS_STR "20"
#define MAX_FILEDIALOGS 4
#define MAX_FOLDERDIALOGS 4
#define MAX_NUMBER_LENGTH 20
// Above is the maximum length of a 64-bit number when expressed as decimal or hex string.
// e.g. -9223372036854775808 or (unsigned) 18446744073709551616

// Hot strings:
// memmove() and proper detection of long hotstrings rely on buf being at least this large:
#define HS_BUF_SIZE (MAX_HOTSTRING_LENGTH * 2 + 10)
#define HS_BUF_DELETE_COUNT (HS_BUF_SIZE / 2)
#define HS_MAX_END_CHARS 100

// Bitwise storage of boolean flags.  This section is kept in this file because
// of mutual dependency problems between hook.h and other header files:
typedef UCHAR HookType;
#define HOOK_KEYBD 0x01
#define HOOK_MOUSE 0x02
#define HOOK_FAIL  0xFF

#define EXTERN_G extern global_struct g
#define EXTERN_OSVER extern OS_Version g_os
#define EXTERN_CLIPBOARD extern Clipboard g_clip
#define EXTERN_SCRIPT extern Script g_script
#define CLOSE_CLIPBOARD_IF_OPEN	if (g_clip.mIsOpen) g_clip.Close()
#define CLIPBOARD_CONTAINS_ONLY_FILES (!IsClipboardFormatAvailable(CF_TEXT) && IsClipboardFormatAvailable(CF_HDROP))


// These macros used to keep app responsive during a long operation.  They may prove to
// be unnecessary if a 2nd thread can be dedicated to checking the message loop, which might
// then prevent keyboard and mouse lag whenever either of the hooks is installed.
// The sleep duration must be greater than zero when the hooks are installed, so that
// MsgSleep() will enter the GetMessage() state, which is the only state that seems to
// pass off keyboard & mouse input to the hooks.
// A value of 8 for how_often_to_sleep seems best if avoiding keyboard/mouse lag is
// top priority.  UPDATE: 18 is now being used because 8 sometimes causes a delay after every keystroke
// (and perhaps ever file in FileSetAttrib() and such), possibly because the system's tickcount
// gets synchronized exactly with the calls to GetTickCount, which means that the first tick
// is fetched less than 1ms before the system is about to update to a new tickcount.  In any case,
// 18 seems like a good trade-off of performance vs. lag (lag is barely noticeable).
// For example, a value of 15 (or maybe it was 25) makes the mouse cursor
// move with an almost imperceptible jumpiness.  Of course, the granularity of GetTickCount()
// is usually 10ms, so it's best to choose a value such as 8, 18, 28, etc. to be sure
// that the proper interval is really being used.
// Making time_of_last_sleep static so that recursive functions, such as FileSetAttrib(),
// will sleep as often as intended even if the target files require frequent recursion.
// Making this static is not friendly to reentrant calls to the function (i.e. calls maded
// as a consequence of the current script subroutine being interrupted by another during
// this instance's MsgSleep()).  However, it doesn't seem to be that much of a consequence
// since the exact interval period of the MsgSleep()'s isn't that important.  It's also
// pretty unlikely that the interrupting subroutine will also just happen to call the same
// function rather than some other.  UPDATE: These macros were greatly simplified when
// it was discovered that PeekMessage(), when called directly as below, is enough to prevent
// keyboard and mouse lag when the hooks are installed:
#define LONG_OPERATION_INIT MSG msg; DWORD tick_now;

// This is the same as the above except it also handled bytes_to_read for URLDownloadToFile():
#define LONG_OPERATION_INIT_FOR_URL \
	MSG msg; DWORD tick_now;\
	DWORD bytes_to_read = Hotkey::HookIsActive() ? 1024 : sizeof(bufData);

// MsgSleep() is used rather than SLEEP_WITHOUT_INTERRUPTION to allow other hotkeys to
// launch and interrupt (suspend) the operation.  It seems best to allow that, since
// the user may want to press some fast window activation hotkeys, for example,
// during the operation.  The operation will be resumed after the interrupting subroutine
// finishes.
// Notes applying to the macro:
// Store tick_now for use later, in case the Peek() isn't done, though not all callers need it later.
// ...
// Since the Peek() might (must?) yield, and thus take a long time to return even when no msg is found,
// must update tick_now again to avoid having to Peek() immediately after the next iteration:
// ...
// Perversely, the code may run faster when "g_script.mLastPeekTime = tick_now" is a sep. operation rather
// than combined in a chained assignment statement.
#define LONG_OPERATION_UPDATE \
{\
	tick_now = GetTickCount();\
	if (tick_now - g_script.mLastPeekTime > 5)\
	{\
		if (PeekMessage(&msg, NULL, 0, 0, PM_NOREMOVE))\
			MsgSleep(-1);\
		tick_now = GetTickCount();\
		g_script.mLastPeekTime = tick_now;\
	}\
}

// Same as the above except for SendKeys() and related functions (uses SLEEP_WITHOUT_INTERRUPTION vs. MsgSleep):
#define LONG_OPERATION_UPDATE_FOR_SENDKEYS \
{\
	tick_now = GetTickCount();\
	if (tick_now - g_script.mLastPeekTime > 5)\
	{\
		if (PeekMessage(&msg, NULL, 0, 0, PM_NOREMOVE))\
			SLEEP_WITHOUT_INTERRUPTION(-1) \
		tick_now = GetTickCount();\
		g_script.mLastPeekTime = tick_now;\
	}\
}

// Defining these here avoids awkwardness due to the fact that globaldata.cpp
// does not (for design reasons) include globaldata.h:
typedef UCHAR ActionTypeType; // If ever have more than 256 actions, will have to change this.
struct Action
{
	char *Name;
	// Just make them int's, rather than something smaller, because the collection
	// of actions will take up very little memory.  Int's are probably faster
	// for the processor to access since they are the native word size, or something:
	int MinParams, MaxParams;
	// Array indicating which args must be purely numeric.  The first arg is
	// number 1, the second 2, etc (i.e. it doesn't start at zero).  The list
	// is ended with a zero, much like a string.  The compiler will notify us
	// (verified) if the number of elements ever needs to be increased:
	#define MAX_NUMERIC_PARAMS 7
	ActionTypeType NumericParams[MAX_NUMERIC_PARAMS];
};

// Same reason as above struct.  It's best to keep this struct as small as possible
// because it's used as a local (stack) var by at least one recursive function:
enum TitleMatchModes {MATCHMODE_INVALID = FAIL, FIND_IN_LEADING_PART, FIND_ANYWHERE, FIND_EXACT, FIND_FAST, FIND_SLOW};

// Bitwise flags for the UCHAR CoordMode:
#define COORD_MODE_PIXEL 0x1
#define COORD_MODE_MOUSE 0x2
#define COORD_MODE_TOOLTIP 0x4

// Each instance of this struct generally corresponds to a quasi-thread.  The function that creates
// a new thread typically saves the old thread's struct values on its stack so that they can later
// be copied back into the g struct when the thread is resumed:
struct global_struct
{
	TitleMatchModes TitleMatchMode;
	bool TitleFindFast; // Whether to use the fast mode of searching window text, or the more thorough slow mode.
	bool DetectHiddenWindows; // Whether to detect the titles of hidden parent windows.
	bool DetectHiddenText;    // Whether to detect the text of hidden child windows.
	__int64 LinesPerCycle; // Use 64-bits for this so that user can specify really large values.
	int IntervalBeforeRest;
	bool AllowThisThreadToBeInterrupted;  // Whether this thread can be interrupted by custom menu items, hotkeys, or timers.
	int UninterruptedLineCount; // Stored as a g-struct attribute in case OnExit sub interrupts it while uninterruptible.
	int Priority;  // This thread's priority relative to others.
	int WinDelay;  // negative values may be used as special flags.
	int ControlDelay;  // negative values may be used as special flags.
	int KeyDelay;  // negative values may be used as special flags.
	int MouseDelay;  // negative values may be used as special flags.
	UCHAR DefaultMouseSpeed;
	UCHAR CoordMode; // Bitwise collection of flags.
	bool StoreCapslockMode;
	bool AutoTrim;
	bool StringCaseSense;
	char FormatFloat[32];
	bool FormatIntAsHex;
	char ErrorLevel[128]; // Big in case user put something bigger than a number in g_ErrorLevel.
	HWND hWndLastUsed;  // In many cases, it's better to use g_ValidLastUsedWindow when referring to this.
	//HWND hWndToRestore;
	int MsgBoxResult;  // Which button was pressed in the most recent MsgBox.
	bool IsPaused;
};

inline void global_clear_state(global_struct *gp)
// Reset those values which represent the condition or state created by previously executed commands.
{
	*gp->ErrorLevel = '\0'; // This isn't the actual ErrorLevel: it's used to save and restore it.
	// But don't reset g_ErrorLevel itself because we want to handle that conditional behavior elsewhere.
	gp->hWndLastUsed = NULL;
	//gp->hWndToRestore = NULL;
	gp->MsgBoxResult = 0;
	gp->IsPaused = false;
	gp->UninterruptedLineCount = 0;
}

inline void global_init(global_struct *gp)
// This isn't made a real constructor to avoid the overhead, since there are times when we
// want to declare a local var of type global_struct without having it initialized.
{
	// Init struct with application defaults.  They're in a struct so that it's easier
	// to save and restore their values when one hotkey interrupts another, going into
	// deeper recursion.  When the interrupting subroutine returns, the former
	// subroutine's values for these are restored prior to resuming execution:
	global_clear_state(gp);
	gp->TitleMatchMode = FIND_IN_LEADING_PART; // Standard default for AutoIt2 and 3.
	gp->TitleFindFast = true; // Since it's so much faster in many cases.
	gp->DetectHiddenWindows = false;  // Same as AutoIt2 but unlike AutoIt3; seems like a more intuitive default.
	gp->DetectHiddenText = true;  // Unlike AutoIt, which defaults to false.  This setting performs better.
	// Not sure what the optimal default is.  1 seems too low (scripts would be very slow by default):
	gp->LinesPerCycle = -1;
	gp->IntervalBeforeRest = 10;  // sleep for 10ms every 10ms
	gp->AllowThisThreadToBeInterrupted = true; // Separate from g_AllowInterruption so that they can have independent values.
	#define PRIORITY_MINIMUM INT_MIN
	gp->Priority = 0;
	gp->WinDelay = 100;  // AutoIt3's default is 250, which seems a little too high nowadays.
	gp->ControlDelay = 20;
	gp->KeyDelay = 10;   // AutoIt3's default.
	gp->MouseDelay = 10;
	// AutoIt3's default:
	#define DEFAULT_MOUSE_SPEED 2
	#define MAX_MOUSE_SPEED 100
	#define MAX_MOUSE_SPEED_STR "100"
	#define COORD_UNSPECIFIED INT_MIN
	gp->DefaultMouseSpeed = DEFAULT_MOUSE_SPEED;
	gp->CoordMode = 0;  // All the flags it contains are off by default.
	gp->StoreCapslockMode = true;  // AutoIt2 (and probably 3's) default, and it makes a lot of sense.
	gp->AutoTrim = true;  // AutoIt2's default, and overall the best default in most cases.
	gp->StringCaseSense = false;  // AutoIt2 default, and it does seem best.
	strcpy(gp->FormatFloat, "%0.6f");
	gp->FormatIntAsHex = false;
	// For FormatFloat:
	// I considered storing more than 6 digits to the right of the decimal point (which is the default
	// for most Unices and MSVC++ it seems).  But going beyond that makes things a little weird for many
	// numbers, due to the inherent imprecision of floating point storage.  For example, 83648.4 divided
	// by 2 shows up as 41824.200000 with 6 digits, but might show up 41824.19999999999700000000 with
	// 20 digits.  The extra zeros could be chopped off the end easily enough, but even so, going beyond
	// 6 digits seems to do more harm than good for the avg. user, overall.  A default of 6 is used here
	// in case other/future compilers have a different default (for backward compatibility, we want
	// 6 to always be in effect as the default for future releases).
}

#endif
