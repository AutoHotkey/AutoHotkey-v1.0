/*
AutoHotkey

Copyright 2003 Christopher L. Mallett

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
#define WINDOW_CLASS_NAME NAME_P
#define NAME_VERSION "0.2.22"
#define NAME_PV NAME_P " v" NAME_VERSION

#define EXT_AUTOIT2 ".aut"
#define EXT_AUTOHOTKEY ".ahk"
#define CONVERSION_FLAG (EXT_AUTOIT2 EXT_AUTOHOTKEY)
#define CONVERSION_FLAG_LENGTH 8

// 10 is the number base system:
#define ITOA(value, buf) _itoa(value, buf, 10)
#define ITOA64(value, buf) _i64toa(value, buf, 10)

// This is not the best way to do this, but it does make MS compilers more
// like MinGW and other Gnu-ish compilers.  Also, I'm not sure what (if
// any) differences exist between _snprintf() and snprintf(), such as
// their return values, so this might not be completely safe:
#ifndef snprintf
#define snprintf _snprintf
#endif
#ifndef vsnprintf
#define vsnprintf _vsnprintf
#endif

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



#define GET_BIT(buf,n) (((buf) & (1 << (n))) >> (n))
#define SET_BIT(buf,n,val) ((val) ? ((buf) |= (1<<(n))) : (buf &= ~(1<<(n))))

// FAIL = 0 to remind that FAIL should have the value zero instead of something arbitrary
// because some callers may simply evaluate the return result as true or false
// (and false is a failure):
enum ResultType {FAIL = 0, OK, WARN = OK, CRITICAL_ERROR
	, CONDITION_TRUE, CONDITION_FALSE
	, LOOP_BREAK, LOOP_CONTINUE
	, EARLY_RETURN, EARLY_EXIT};

// These are used for things that can be turned on, off, or left at a
// neutral default value that is neither on nor off.  INVALID must
// be zero:
enum ToggleValueType {TOGGLE_INVALID = 0, TOGGLED_ON, TOGGLED_OFF, ALWAYS_ON, ALWAYS_OFF
	, TOGGLE, TOGGLE_PERMIT, NEUTRAL};

// For convenience in many places.  Must cast to int to avoid loss of negative values.
#define BUF_SPACE_REMAINING ((int)(aBufSize - (aBuf - aBuf_orig)))

// MsgBox timeout value.  This can't be zero because that is used as a failure indicator:
// Also, this define is in this file to prevent problems with mutual
// dependency between script.h and window.h:
#define AHK_TIMEOUT -1
// And these to prevent mutual dependency problem between window.h and globaldata.h:
#define MAX_MSGBOXES 7
#define MAX_INPUTBOXES 4
#define MAX_FILEDIALOGS 4
#define MAX_FOLDERDIALOGS 4

// Bitwise storage of boolean flags.  This section is kept in this file because
// of mutual dependency problems between hook.h and other header files:
typedef UCHAR HookType;
#define HOOK_KEYBD 0x01
#define HOOK_MOUSE 0x02
#define HOOK_FAIL  0xFF

#define EXTERN_CLIPBOARD extern Clipboard g_clip
#define CLOSE_CLIPBOARD_IF_OPEN	if (g_clip.mIsOpen) g_clip.Close()
#define CLIPBOARD_CONTAINS_ONLY_FILES (!IsClipboardFormatAvailable(CF_TEXT) && IsClipboardFormatAvailable(CF_HDROP))

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
enum TitleMatchModes {MATCHMODE_INVALID = FAIL, FIND_IN_LEADING_PART, FIND_ANYWHERE, FIND_FAST, FIND_SLOW};
struct global_struct
{
	bool TitleFindAnywhere;  // Whether match can be found anywhere in a window title, rather than leading part.
	bool TitleFindFast; // Whether to use the fast mode of searching window text, or the more thorough slow mode.
	bool DetectHiddenWindows; // Whether to detect the titles of hidden parent windows.
	bool DetectHiddenText;    // Whether to detect the text of hidden child windows.
	__int64 LinesPerCycle; // Use 64-bits for this so that user can specify really large values.
	int WinDelay;  // negative values may be used as special flags.
	int ControlDelay;  // negative values may be used as special flags.
	int KeyDelay;  // negative values may be used as special flags.
	UCHAR DefaultMouseSpeed;
	bool StoreCapslockMode;
	bool AutoTrim;
	bool StringCaseSense;
	char FormatFloat[32];
	char ErrorLevel[128]; // Big in case user put something bigger than a number in g_ErrorLevel.
	HWND hWndLastUsed;  // In many cases, it's better to use g_ValidLastUsedWindow when referring to this.
	HWND hWndToRestore;
	int MsgBoxResult;  // Which button was pressed in the most recent MsgBox.
	bool WaitingForDialog;
	bool IsPaused;
};

inline void global_clear_state(global_struct *gp)
// Reset those values which represent the condition or state created by previously executed commands.
{
	*gp->ErrorLevel = '\0'; // This isn't the actual ErrorLevel: it's used to save and restore it.
	// But don't reset g_ErrorLevel itself because we want to handle that conditional behavior elsewhere.
	gp->hWndLastUsed = gp->hWndToRestore = NULL;
	gp->MsgBoxResult = 0;
	gp->WaitingForDialog = false;
	gp->IsPaused = false;

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
	gp->TitleFindAnywhere = false; // Standard default for AutoIt2 and 3: Leading part of window title must match.
	gp->TitleFindFast = true; // Since it's so much faster in many cases.
	gp->DetectHiddenWindows = false;  // Same as AutoIt2 but unlike AutoIt3; seems like a more intuitive default.
	gp->DetectHiddenText = true;  // Unlike AutoIt, which defaults to false.  This setting performs better.
	// Not sure what the optimal default is.  1 seems too low (scripts would be very slow by default):
	#define DEFAULT_BATCH_LINES 10
	gp->LinesPerCycle = DEFAULT_BATCH_LINES;
	gp->WinDelay = 100;  // AutoIt3's default is 250, which seems a little too high nowadays.
	gp->ControlDelay = 20;
	gp->KeyDelay = 10;   // AutoIt3's default.
	// AutoIt3's default:
	#define DEFAULT_MOUSE_SPEED 2
	#define MAX_MOUSE_SPEED 100
	#define MAX_MOUSE_SPEED_STR "100"
	#define COORD_UNSPECIFIED (INT_MIN)
	gp->DefaultMouseSpeed = DEFAULT_MOUSE_SPEED;
	gp->StoreCapslockMode = true;  // AutoIt2 (and probably 3's) default, and it makes a lot of sense.
	gp->AutoTrim = true;  // AutoIt2's default, and overall the best default in most cases.
	gp->StringCaseSense = false;  // AutoIt2 default, and it does seem best.
	strcpy(gp->FormatFloat, "%0.6f");
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
