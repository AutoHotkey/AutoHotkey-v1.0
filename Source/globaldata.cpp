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
// These includes should probably a superset of those in globaldata.h:
#include "hook.h" // For KeyLogItem and probably other things.
#include "clipboard.h"  // For the global clipboard object
#include "script.h" // For the global script object and g_ErrorLevel
#include "os_version.h" // For the global OS_Version object

// Since at least some of some of these (e.g. g_modifiersLR_logical) should not
// be kept in the struct since it's not correct to save and restore their
// state, don't keep anything in the global_struct except those things
// which are necessary to save and restore (even though it would clean
// up the code and might make maintaining it easier):
HWND g_hWnd = NULL;
HWND g_hWndEdit = NULL;
HWND g_hWndSplash = NULL;
HINSTANCE g_hInstance; // Set by WinMain().
modLR_type g_modifiersLR_logical = 0;
modLR_type g_modifiersLR_physical = 0;
modLR_type g_modifiersLR_get = 0;

// Used by the hook to track physical state of all virtual keys, since GetAsyncKeyState() does
// not work as advertised, at least under WinXP:
bool g_PhysicalKeyState[VK_MAX + 1] = {false};

int g_HotkeyModifierTimeout = 100;
HHOOK g_hhkLowLevelKeybd = NULL;
HHOOK g_hhkLowLevelMouse = NULL;
bool g_ForceLaunch = false;
bool g_AllowOnlyOneInstance = false;
bool g_NoTrayIcon = false;
bool g_AllowSameLineComments = true;
char g_LastPerformedHotkeyType = HK_NORMAL;
bool g_IsIdle = false;  // Set false as the initial state for use during the auto-execute part of the script.
bool g_IsSuspended = false;  // Make this separate from g_IgnoreHotkeys since that is frequently turned off & on.
bool g_IgnoreHotkeys = false;
int g_nInterruptedSubroutines = 0;
int g_nPausedSubroutines = 0;
bool g_UnpauseWhenResumed = false;  // Start off "false" because the Unpause mode must be explicitly triggered.

// On my system, the repeat-rate (which is probably set to XP's default) is such that between 20
// and 25 keys are generated per second.  Therefore, 50 in 2000ms seems like it should allow the
// key auto-repeat feature to work on most systems without triggering the warning dialog.
// In any case, using auto-repeat with a hotkey is pretty rare for most people, so it's best
// to keep these values conservative:
int g_MaxHotkeysPerInterval = 50;
int g_HotkeyThrottleInterval = 2000; // Milliseconds.

bool g_TrayMenuIsVisible = false;
int g_nMessageBoxes = 0;
int g_nInputBoxes = 0;
int g_nFileDialogs = 0;
int g_nFolderDialogs = 0;
InputBoxType g_InputBox[MAX_INPUTBOXES];

char g_delimiter = ',';
char g_DerefChar = '%';
char g_EscapeChar = '`';

// Global objects:
Script g_script;
Var *g_ErrorLevel = NULL; // Allows us (in addition to the user) to set this var to indicate success/failure.
// This made global for performance reasons (determining size of clipboard data then
// copying contents in or out without having to close & reopen the clipboard in between):
Clipboard g_clip;
OS_Version g_os;  // OS version object, courtesy of AutoIt3.

DWORD g_OriginalTimeout;

global_struct g, g_default;

bool g_ForceKeybdHook = false;
ToggleValueType g_ForceNumLock = NEUTRAL;
ToggleValueType g_ForceCapsLock = NEUTRAL;
ToggleValueType g_ForceScrollLock = NEUTRAL;

vk2_type g_sc_to_vk[SC_MAX + 1] = {{0}};
sc2_type g_vk_to_sc[VK_MAX + 1] = {{0}};


// The order of initialization here must match the order in the enum contained in script.h
// It's in there rather than in globaldata.h so that the action-type constants can be referred
// to without having access to the global array itself (i.e. it avoids having to include
// globaldata.h in modules that only need access to the enum's constants, which in turn prevents
// many mutual dependency problems between modules).  Note: Action names must not contain any
// spaces or tabs because within a script, those characters can be used in lieu of a delimiter
// to separate the action-type-name from the first parameter.
// Note about the sub-array: Since the parent array array is global, it would be automatically
// zero-filled if we didn't provide specific initialization.  But since we do, I'm not sure
// what value the unused elements in the NumericParams subarray will have.  Therefore, it seems
// safest to always terminate these subarrays with an explicit zero, below.

// STEPS TO ADD A NEW COMMAND:
// 1) Add an entry to the command enum in script.h.
// 2) Add an entry to the below array (it's position here MUST exactly match that in the enum).
//    The subarray should indicate the param numbers that must be numeric (first param is 1,
//    not zero).  That subarray should be terminated with an explicit zero to be safe and
//    so that the compiler will complain if the sub-array size needs to be increased to
//    accommodate all the elements in the new sub-array, including room for it's 0 terminator.
//    If any of the numeric params allow negative or float values, add an entries to
//    ArgAllowsNegative() and ArgAllowsFloat().
//    If any of the params are mandatory (can't be blank), add an entry to CheckForMandatoryArgs().
//    Note: If you use a value for MinParams than is greater than zero, remember than any params
//    beneath that threshold will also be required to be non-blank (i.e. user can't omit them even
//    if later, non-blank params are provided).
// 3) If the new command has any params that are output or input vars, change Line::ArgIsVar().
// 4) Add any desired load-time validation in Script::AddLine() in an appropriate section.
// 5) Implement the command in Line::Perform() or Line::EvaluateCondition (if it's an IF).
//    If the command waits for anything (e.g. calls MsgSleep()), be sure to make a local
//    copy of any ARG values that are needed during the wait period, because if another hotkey
//    subroutine suspends the current one while its waiting, it could also overwrite the ARG
//    deref buffer with its own values.

Action g_act[] =
{
	{"<invalid command>", 0, 0, NULL}  // ACT_INVALID.  Give it a name in case it's ever displayed.

	// ACT_ASSIGN, ACT_ADD/SUB/MULT/DIV: Give them names for display purposes.
	// Note: Line::ToText() relies on the below names being the correct symbols for the operation:
	// 1st param is the target, 2nd (optional) is the value:
	, {"=", 1, 2, NULL} // For this one, omitting the second param sets the var to be empty.

	// Subtraction(but not addition) allow 2nd to be blank due to 3rd param.
	// Also, it seems ok to allow date-time operations with += and -=, even though these
	// operators may someday be enhanced to handle complex expressions, since it seems
	// possible to parse out the TimeUnits parameter even from a complex expression (since
	// such expressions wouldn't be expected to use commas for anything else?):
	, {"+=", 2, 3, {2, 0}}
	, {"-=", 1, 3, {2, 0}}
	, {"*=", 2, 2, {2, 0}}
	, {"/=", 2, 2, {2, 0}}

	// This command is never directly parsed, but we need to have it here as a translation
	// target for the old "repeat" command.  This is because that command treats a zero
	// first-param as an infinite loop.  Since that param can be a dereferenced variable,
	// there's no way to reliably translate each REPEAT command into a LOOP command at
	// load-time.  Thus, we support both types of loops as actual commands that are
	// handled separately at runtime.
	, {"Repeat", 0, 1, {1, 0}}  // Iteration Count: was mandatory in AutoIt2 but doesn't seem necessary here.
	, {"Else", 0, 0, NULL}

	// Comparison operators take 1 param (if they're being compared to blank) or 2.
	// For example, it's okay (though probably useless) to compare a string to the empty
	// string this way: "If var1 >=".  Note: Line::ToText() relies on the below names:
	, {"=", 1, 2, NULL}, {"<>", 1, 2, NULL}, {">", 1, 2, NULL}
	, {">=", 1, 2, NULL}, {"<", 1, 2, NULL}, {"<=", 1, 2, NULL}
	, {"is", 2, 2, NULL}, {"is not", 2, 2, NULL}

	// For these, allow a minimum of zero, otherwise, the first param (WinTitle) would
	// be considered mandatory-non-blank by default.  It's easier to make all the params
	// optional and validate elsewhere that at least one of the four isn't blank.
	// Also, All the IFs must be physically adjacent to each other in this array
	// so that ACT_IF_FIRST and ACT_IF_LAST can be used to detect if a command is an IF:
	, {"IfWinExist", 0, 4, NULL}, {"IfWinNotExist", 0, 4, NULL}  // Title, text, exclude-title, exclude-text
	// Passing zero params results in activating the LastUsed window:
	, {"IfWinActive", 0, 4, NULL}, {"IfWinNotActive", 0, 4, NULL} // same
	, {"IfInString", 2, 2, NULL} // String var, search string
	, {"IfNotInString", 2, 2, NULL} // String var, search string
	, {"IfExist", 1, 1, NULL} // File or directory.
	, {"IfNotExist", 1, 1, NULL} // File or directory.
	// IfMsgBox must be physically adjacent to the other IFs in this array:
	, {"IfMsgBox", 1, 1, NULL} // MsgBox result (e.g. OK, YES, NO)
	, {"MsgBox", 0, 4, {4, 0}} // Text (if only 1 param) or: Mode-flag, Title, Text, Timeout.
	, {"InputBox", 1, 4, NULL} // Output var, title, prompt, hide-text (e.g. passwords)
	, {"SplashTextOn", 0, 4, {1, 2, 0}} // Width, height, title, text
	, {"SplashTextOff", 0, 0, NULL}

	, {"StringLeft", 3, 3, {3, 0}}  // output var, input var, number of chars to extract
	, {"StringRight", 3, 3, {3, 0}} // same
	, {"StringMid", 4, 4, {3, 4, 0}} // Output Variable, Input Variable, Start char, Number of chars to extract
	, {"StringTrimLeft", 3, 3, {3, 0}}  // output var, input var, number of chars to trim
	, {"StringTrimRight", 3, 3, {3, 0}} // same
	, {"StringLower", 2, 2, NULL} // output var, input var
	, {"StringUpper", 2, 2, NULL} // output var, input var
	, {"StringLen", 2, 2, NULL} // output var, input var
	, {"StringGetPos", 3, 4, NULL}  // Output Variable, Input Variable, Search Text, R or Right (from right)
	, {"StringReplace", 3, 5, NULL} // Output Variable, Input Variable, Search String, Replace String, do-all.

	, {"EnvSet", 1, 2, NULL} // EnvVar, Value
	, {"EnvUpdate", 0, 0, NULL}

	, {"Run", 1, 3, NULL}, {"RunWait", 1, 3, NULL}  // TargetFile, Working Dir, WinShow-Mode
	, {"GetKeyState", 2, 3, NULL} // OutputVar, key name, mode (optional) P = Physical, T = Toggle
	, {"Send", 1, 1, NULL} // But that first param can be a deref that resolves to a blank param
	// For these, the "control" param can be blank.  The window's first visible control will
	// be used.  For this first one, allow a minimum of zero, otherwise, the first param (control)
	// would be considered mandatory-non-blank by default.  It's easier to make all the params
	// optional and validate elsewhere that the 2nd one specifically isn't blank:
	, {"ControlSend", 0, 6, NULL} // Control, Chars-to-Send, std. 4 window params.
	, {"ControlLeftClick", 0, 5, NULL} // Control, std. 4 window params
	, {"ControlGetFocus", 1, 5, NULL}  // OutputVar, std. 4 window params
	, {"ControlFocus", 0, 5, NULL}     // Control, std. 4 window params
	, {"ControlSetText", 1, 6, NULL}   // Control, new text, std. 4 window params
	, {"ControlGetText", 1, 6, NULL}   // Output-var, Control, std. 4 window params

	, {"SetDefaultMouseSpeed", 1, 1, {1, 0}} // speed (numeric)
	, {"MouseMove", 2, 3, {1, 2, 3, 0}} // x, y, speed
	, {"MouseClick", 1, 6, {2, 3, 4, 5, 0}} // which-button, x, y, ClickCount, speed, d=hold-down/u=release
	, {"MouseClickDrag", 1, 6, {2, 3, 4, 5, 6, 0}} // which-button, x1, y1, x2, y2, speed
	, {"MouseGetPos", 0, 2, NULL} // 2 optional output variables: one for xpos, and one for ypos. MinParams must be 0.

	, {"StatusBarGetText", 1, 6, {2, 0}} // Output-var, part# (numeric), std. 4 window params
	, {"StatusBarWait", 0, 8, {2, 3, 6, 0}} // Wait-text(blank ok),seconds,part#,title,text,interval,exclude-title,exclude-text
	, {"ClipWait", 0, 1, {1, 0}} // Seconds-to-wait (0 = 500ms)

	, {"Sleep", 1, 1, {1, 0}} // Sleep time in ms (numeric)
	, {"Random", 1, 3, {2, 3, 0}} // Output var, Min, Max (Note: MinParams is 1 so that param2 can be blank).
	, {"Goto", 1, 1, NULL}, {"Gosub", 1, 1, NULL} // Label (or dereference that resolves to a label).
	, {"Return", 0, 0, NULL}, {"Exit", 0, 1, {1, 0}} // ExitCode (currently ignored)
	, {"Loop", 0, 3, NULL} // Iteration Count or file-search (e.g. c:\*.*), FileLoopMode, Recurse? (custom validation for these last two)
	, {"Break", 0, 0, NULL}, {"Continue", 0, 0, NULL}
	, {"{", 0, 0, NULL}, {"}", 0, 0, NULL}

	, {"WinActivate", 0, 4, NULL} // Passing zero params results in activating the LastUsed window.
	, {"WinActivateBottom", 0, 4, NULL} // Min. 0 so that 1st params can be blank and later ones not blank.

	// These all use Title, Text, Timeout (in seconds not ms), Exclude-title, Exclude-text.
	// See above for why zero is the minimum number of params for each:
	, {"WinWait", 0, 5, {3, 0}}, {"WinWaitClose", 0, 5, {3, 0}}
	, {"WinWaitActive", 0, 5, {3, 0}}, {"WinWaitNotActive", 0, 5, {3, 0}}

	, {"WinMinimize", 0, 4, NULL}, {"WinMaximize", 0, 4, NULL}, {"WinRestore", 0, 4, NULL} // std. 4 params
	, {"WinHide", 0, 4, NULL}, {"WinShow", 0, 4, NULL} // std. 4 params
	, {"WinMinimizeAll", 0, 0, NULL}, {"WinMinimizeAllUndo", 0, 0, NULL}
	, {"WinClose", 0, 5, {3, 0}} // title, text, time-to-wait-for-close (0 = 500ms), exclude title/text
	, {"WinKill", 0, 5, {3, 0}} // same as WinClose.
	, {"WinMove", 0, 8, {3, 4, 5, 6, 0}} // title, text, xpos, ypos, width, height, exclude-title, exclude_text
	// Note for WinMove: xpos/ypos/width/height can be the string "default", but that is explicitly
	// checked for in spite of requiring it to be numeric in the definition here.
	, {"WinMenuSelectItem", 0, 11, NULL} // WinTitle, WinText, Menu name, 6 optional sub-menu names, ExcludeTitle/Text

	// WinSetTitle: Allow a minimum of zero params so that title isn't forced to be non-blank.
	// Also, if the user passes only one param, the title of the "last used" window will be
	// set to the string in the first param:
	, {"WinSetTitle", 0, 5, NULL} // title, text, newtitle, exclude-title, exclude-text
	, {"WinGetTitle", 1, 5, NULL} // Output-var, std. 4 window params
	, {"WinGetPos", 0, 8, NULL} // Four optional output vars: xpos, ypos, width, height.  Std. 4 window params.
	, {"WinGetText", 1, 5, NULL} // Output var, std 4 window params.

	, {"PixelGetColor", 3, 3, {2, 3, 0}} // OutputVar, X-coord, Y-coord
	, {"PixelSearch", 0, 8, {3, 4, 5, 6, 7, 8, 0}} // OutputX, OutputY, left, top, right, bottom, Color, Variation
	// Note in the above: 0 min args so that the output vars can be optional.

	// See above for why minimum is 1 vs. 2:
	, {"GroupAdd", 1, 6, NULL} // Group name, WinTitle, WinText, Label, exclude-title/text
	, {"GroupActivate", 1, 2, NULL}
	, {"GroupDeactivate", 1, 2, NULL}
	, {"GroupClose", 1, 2, NULL}

	, {"DriveSpaceFree", 2, 2, NULL} // Output-var, path (e.g. c:\)
	, {"SoundSetWaveVolume", 1, 1, {1, 0}} // Volume percent-level (0-100)
	, {"SoundPlay", 1, 2, NULL} // Filename [, wait]

	, {"FileAppend", 2, 2, NULL} // text, filename
	, {"FileReadLine", 3, 3, NULL} // Output variable, filename, line-number (custom validation, not numeric validation)
	, {"FileCopy", 2, 3, {3, 0}} // source, dest, flag
	, {"FileMove", 2, 3, {3, 0}} // source, dest, flag
	, {"FileDelete", 1, 1, NULL} // filename
	, {"FileCreateDir", 1, 1, NULL} // dir name
	, {"FileRemoveDir", 1, 1, NULL} // dir name

	, {"FileGetAttrib", 1, 2, NULL} // OutputVar, Filespec (if blank, uses loop's current file)
	, {"FileSetAttrib", 1, 4, NULL} // Attribute(s), FilePattern, OperateOnFolders?, Recurse? (custom validation for these last two)
	, {"FileGetTime", 1, 3, NULL} // OutputVar, Filespec, WhichTime (modified/created/accessed)
	, {"FileSetTime", 0, 5, {1, 0}} // datetime (YYYYMMDDHH24MISS), FilePattern, WhichTime, OperateOnFolders?, Recurse?
	, {"FileGetSize", 1, 3, NULL} // OutputVar, Filespec, B|K|M (bytes, kb, or mb)
	, {"FileGetVersion", 1, 2, NULL} // OutputVar, Filespec

	, {"FileSelectFile", 1, 3, {2, 0}} // output var, flag, working dir
	, {"FileSelectFolder", 1, 4, NULL} // output var, root directory, allow create folder (0=no, 1=yes), greeting

	, {"IniRead", 4, 5, NULL}   // OutputVar, Filespec, Section, Key, Default (value to return if key not found)
	, {"IniWrite", 4, 4, NULL}  // Value, Filespec, Section, Key
	, {"IniDelete", 3, 3, NULL} // Filespec, Section, Key

	, {"RegRead", 1, 5, NULL} // output var, (ValueType [optional]), RegKey, RegSubkey, ValueName
	, {"RegWrite", 4, 5, NULL} // ValueType, RegKey, RegSubKey, ValueName, Value (set to blank if omitted?)
	, {"RegDelete", 2, 3, NULL} // RegKey, RegSubKey, ValueName

	, {"SetKeyDelay", 1, 1, {1, 0}} // Delay in ms (numeric, negative allowed)
	, {"SetWinDelay", 1, 1, {1, 0}} // Delay in ms (numeric, negative allowed)
	, {"SetControlDelay", 1, 1, {1, 0}} // Delay in ms (numeric, negative allowed)
	, {"SetBatchLines", 1, 1, {1, 0}} // Number of script lines to execute before sleeping.
	, {"SetTitleMatchMode", 1, 1, NULL} // Allowed values: 1, 2, slow, fast
	, {"SetFormat", 1, 2, {2, 0}} // OptionName, FormatString

	, {"Suspend", 0, 1, NULL} // On/Off/Toggle/Permit/Blank (blank is the same as toggle)
	, {"Pause", 0, 1, NULL} // On/Off/Toggle/Blank (blank is the same as toggle)
	, {"AutoTrim", 1, 1, NULL} // On/Off
	, {"StringCaseSense", 1, 1, NULL} // On/Off
	, {"DetectHiddenWindows", 1, 1, NULL} // On/Off
	, {"DetectHiddenText", 1, 1, NULL} // On/Off

	, {"SetNumlockState", 0, 1, NULL} // On/Off/AlwaysOn/AlwaysOff or blank (unspecified) to return to normal.
	, {"SetScrollLockState", 0, 1, NULL} // same
	, {"SetCapslockState", 0, 1, NULL} // same
	, {"SetStoreCapslockMode", 1, 1, NULL} // On/Off

	, {"KeyLog", 0, 2, NULL}, {"ListLines", 0, 0, NULL}
	, {"ListVars", 0, 0, NULL}, {"ListHotkeys", 0, 0, NULL}

	, {"Edit", 0, 0, NULL}
	, {"Reload", 0, 0, NULL}
	, {"ExitApp", 0, 1, NULL}  // Optional exit-code
	, {"Shutdown", 1, 1, {1, 0}} // Seems best to make the first param (the flag/code) mandatory.
};
// Below is the most maintainable way to determine the actual count?
// Due to C++ lang. restrictions, can't easily make this a const because constants
// automatically get static (internal) linkage, thus such a var could never be
// used outside this module:
int g_ActionCount = sizeof(g_act) / sizeof(Action);



Action g_old_act[] =
{
	{"<invalid command>", 0, 0, NULL}  // OLD_INVALID.  Give it a name in case it's ever displayed.
	, {"SetEnv", 1, 2, NULL}
	, {"EnvAdd", 2, 3, {2, 0}}, {"EnvSub", 1, 3, {2, 0}} // EnvSub (but not Add) allow 2nd to be blank due to 3rd param.
	, {"EnvMult", 2, 2, {2, 0}}, {"EnvDiv", 2, 2, {2, 0}}
	, {"IfEqual", 1, 2, NULL}, {"IfNotEqual", 1, 2, NULL}
	, {"IfGreater", 1, 2, NULL}, {"IfGreaterOrEqual", 1, 2, NULL}
	, {"IfLess", 1, 2, NULL}, {"IfLessOrEqual", 1, 2, NULL}
	, {"LeftClick", 2, 2, {1, 2, 0}}, {"RightClick", 2, 2, {1, 2, 0}}
	, {"LeftClickDrag", 4, 4, {1, 2, 3, 4, 0}}, {"RightClickDrag", 4, 4, {1, 2, 3, 4, 0}}
	  // Allow zero params, unlike AutoIt.  These params should match those for REPEAT in the above array:
	, {"Repeat", 0, 1, {1, 0}}, {"EndRepeat", 0, 0, NULL}
	, {"WinGetActiveTitle", 1, 1, NULL} // <Title Var>
	, {"WinGetActiveStats", 5, 5, NULL} // <Title Var>, <Width Var>, <Height Var>, <Xpos Var>, <Ypos Var>
};
int g_OldActionCount = sizeof(g_old_act) / sizeof(Action);


key_to_vk_type g_key_to_vk[] =
{ {"Numpad0", VK_NUMPAD0}
, {"Numpad1", VK_NUMPAD1}
, {"Numpad2", VK_NUMPAD2}
, {"Numpad3", VK_NUMPAD3}
, {"Numpad4", VK_NUMPAD4}
, {"Numpad5", VK_NUMPAD5}
, {"Numpad6", VK_NUMPAD6}
, {"Numpad7", VK_NUMPAD7}
, {"Numpad8", VK_NUMPAD8}
, {"Numpad9", VK_NUMPAD9}
, {"NumpadMult", VK_MULTIPLY}
, {"NumpadDiv", VK_DIVIDE}
, {"NumpadAdd", VK_ADD}
, {"NumpadSub", VK_SUBTRACT}
// , {"NumpadEnter", VK_RETURN}  // Must do this one via scan code, see below for explanation.
, {"NumpadDot", VK_DECIMAL}
, {"Numlock", VK_NUMLOCK}
, {"ScrollLock", VK_SCROLL}
, {"CapsLock", VK_CAPITAL}

, {"Escape", VK_ESCAPE}
, {"Esc", VK_ESCAPE}
, {"Tab", VK_TAB}
, {"Space", VK_SPACE}
, {"Backspace", VK_BACK}
, {"BS", VK_BACK}

// These keys each have a counterpart on the number pad with the same VK.  Use the VK for these,
// since they are probably more likely to be assigned to hotkeys (thus minimizing the use of the
// keyboard hook, and use the scan code (SC) for their counterparts.  UPDATE: To support handling
// these keys with the hook (i.e. the sc_takes_precedence flag in the hook), do them by scan code
// instead.  This allows Numpad keys such as Numpad7 to be differentiated from NumpadHome, which
// would otherwise be impossible since both of them share the same scan code (i.e. if the
// sc_takes_precedence flag is set for the scan code of NumpadHome, that will effectively prevent
// the hook from telling the difference between it and Numpad7 since the hook is currently set
// to handle an incoming key by either vk or sc, but not both.

// Even though ENTER is probably less likely to be assigned than NumpadEnter, must have ENTER as
// the primary vk because otherwise, if the user configures only naked-NumPadEnter to do something,
// RegisterHotkey() would register that vk and ENTER would also be configured to do the same thing.
, {"Enter", VK_RETURN}
, {"Return", VK_RETURN}

, {"NumpadDel", VK_DELETE}
, {"NumpadIns", VK_INSERT}
, {"NumpadClear", VK_CLEAR}  // same physical key as Numpad5 on most keyboards?
, {"NumpadUp", VK_UP}
, {"NumpadDown", VK_DOWN}
, {"NumpadLeft", VK_LEFT}
, {"NumpadRight", VK_RIGHT}
, {"NumpadHome", VK_HOME}
, {"NumpadEnd", VK_END}
, {"NumpadPgUp", VK_PRIOR}
, {"NumpadPgDn", VK_NEXT}

, {"PrintScreen", VK_SNAPSHOT}
, {"CtrlBreak", VK_CANCEL}  // Might want to verify this, and whether it has any peculiarities.
, {"Pause", VK_PAUSE}
, {"Break", VK_PAUSE}
, {"Help", VK_HELP}  // VK_HELP is probably not the extended HELP key.  Not sure what this one is.

, {"AppsKey", VK_APPS}

// UPDATE: For the NT/2k/XP version, now doing these by VK since it's likely to be
// more compatible with non-standard or non-English keyboards:
, {"LControl", VK_LCONTROL}
, {"RControl", VK_RCONTROL}
, {"LCtrl", VK_LCONTROL} // Support this alternate to be like AutoIt3.
, {"RCtrl", VK_RCONTROL} // Support this alternate to be like AutoIt3.
, {"LShift", VK_LSHIFT}
, {"RShift", VK_RSHIFT}
, {"LAlt", VK_LMENU}
, {"RAlt", VK_RMENU}
// These two are always left/right centric and I think their vk's are always supported by the various
// Windows API calls, unlike VK_RSHIFT, etc. (which are seldom supported):
, {"LWin", VK_LWIN}
, {"RWin", VK_RWIN}

// The left/right versions of these are handled elsewhere since their virtual keys aren't fully API-supported:
, {"Control", VK_CONTROL}
, {"Alt", VK_MENU}
, {"Shift", VK_SHIFT}
/*
These were used to confirm the fact that you can't use RegisterHotkey() on VK_LSHIFT, even if the shift
modifier is specified along with it:
, {"LShift", VK_LSHIFT}
, {"RShift", VK_RSHIFT}
*/
, {"F1", VK_F1}
, {"F2", VK_F2}
, {"F3", VK_F3}
, {"F4", VK_F4}
, {"F5", VK_F5}
, {"F6", VK_F6}
, {"F7", VK_F7}
, {"F8", VK_F8}
, {"F9", VK_F9}
, {"F10", VK_F10}
, {"F11", VK_F11}
, {"F12", VK_F12}
, {"F13", VK_F13}
, {"F14", VK_F14}
, {"F15", VK_F15}
, {"F16", VK_F16}
, {"F17", VK_F17}
, {"F18", VK_F18}
, {"F19", VK_F19}
, {"F20", VK_F20}
, {"F21", VK_F21}
, {"F22", VK_F22}
, {"F23", VK_F23}
, {"F24", VK_F24}

// Mouse buttons:
, {"LButton", VK_LBUTTON}
, {"RButton", VK_RBUTTON}
, {"MButton", VK_MBUTTON}
// Supported in only in Win2k and beyond:
, {"XButton1", VK_XBUTTON1}
, {"XButton2", VK_XBUTTON2}
// Custom/fake VKs for use by the mouse hook (supported only in WinNT SP3 and beyond?):
, {"WheelDown", VK_WHEEL_DOWN}
, {"WheelUp", VK_WHEEL_UP}

, {"Browser_Back", VK_BROWSER_BACK}
, {"Browser_Forward", VK_BROWSER_FORWARD}
, {"Browser_Refresh", VK_BROWSER_REFRESH}
, {"Browser_Stop", VK_BROWSER_STOP}
, {"Browser_Search", VK_BROWSER_SEARCH}
, {"Browser_Favorites", VK_BROWSER_FAVORITES}
, {"Browser_Home", VK_BROWSER_HOME}
, {"Volume_Mute", VK_VOLUME_MUTE}
, {"Volume_Down", VK_VOLUME_DOWN}
, {"Volume_Up", VK_VOLUME_UP}
, {"Media_Next", VK_MEDIA_NEXT_TRACK}  // Use the AutoIt3 convention of omitting "_Track" from the name.
, {"Media_Prev", VK_MEDIA_PREV_TRACK}  // Use the AutoIt3 convention of omitting "_Track" from the name.
, {"Media_Stop", VK_MEDIA_STOP}
, {"Media_Play_Pause", VK_MEDIA_PLAY_PAUSE}
, {"Launch_Mail", VK_LAUNCH_MAIL}
, {"Launch_Media", VK_LAUNCH_MEDIA_SELECT} // Use the AutoIt3 name.
, {"Launch_App1", VK_LAUNCH_APP1}
, {"Launch_App2", VK_LAUNCH_APP2}

// Probably safest to terminate it this way, with a flag value.  (plus this makes it a little easier
// to code some loops, maybe).  Can also calculate how many elements are in the array using sizeof(array)
// divided by sizeof(element).  UPDATE: Decided not to do this in case ever decide to sort this array; don't
// want to rely on the fact that this will wind up in the right position after the sort (even though it
// should):
//, {"", 0}
};



key_to_sc_type g_key_to_sc[] =
// Even though ENTER is probably less likely to be assigned than NumpadEnter, must have ENTER as
// the primary vk because otherwise, if the user configures only naked-NumPadEnter to do something,
// RegisterHotkey() would register that vk and ENTER would also be configured to do the same thing.
{ {"NumpadEnter", SC_NUMPADENTER}

, {"Delete", SC_DELETE}
, {"Del", SC_DELETE}
, {"Insert", SC_INSERT}
, {"Ins", SC_INSERT}
// , {"Clear", SC_CLEAR}  // Seems unnecessary because there is no counterpart to the Numpad5 clear key?
, {"Up", SC_UP}
, {"Down", SC_DOWN}
, {"Left", SC_LEFT}
, {"Right", SC_RIGHT}
, {"Home", SC_HOME}
, {"End", SC_END}
, {"PgUp", SC_PGUP}
, {"PgDn", SC_PGDN}

// If user specified left or right, must use scan code to distinguish *both* halves of the pair since
// each half has same vk *and* since their generic counterparts (e.g. CONTROL vs. L/RCONTROL) are
// already handled by vk.  Note: RWIN and LWIN don't need to be handled here because they each have
// their own virtual keys.
// UPDATE: For the NT/2k/XP version, now doing these by VK since it's likely to be
// more compatible with non-standard or non-English keyboards:
/*
, {"LControl", SC_LCONTROL}
, {"RControl", SC_RCONTROL}
, {"LShift", SC_LSHIFT}
, {"RShift", SC_RSHIFT}
, {"LAlt", SC_LALT}
, {"RAlt", SC_RALT}
*/
};



// Can calc the counts only after the arrays are initialized above:
int g_key_to_vk_count = sizeof(g_key_to_vk) / sizeof(key_to_vk_type);
int g_key_to_sc_count = sizeof(g_key_to_sc) / sizeof(key_to_sc_type);

KeyLogItem g_KeyLog[MAX_LOGGED_KEYS] = {{0}};
int g_KeyLogNext = 0;
bool g_KeyLogToFile = false;
