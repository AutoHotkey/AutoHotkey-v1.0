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

#ifndef script_h
#define script_h

#include "stdafx.h" // pre-compiled headers
#include "defines.h"
#include "SimpleHeap.h" // for overloaded new/delete operators.
#include "keyboard.h" // for modLR_type
#include "var.h" // for a script's variables.
#include "WinGroup.h" // for a script's Window Groups.
#include "Util.h" // for FileTimeToYYYYMMDD(), strlcpy()
#include "resources\resource.h"  // For tray icon.
#ifdef AUTOHOTKEYSC
	#include "lib/exearc_read.h"
#endif

#include "os_version.h" // For the global OS_Version object
EXTERN_OSVER; // For the access to the g_os version object without having to include globaldata.h
EXTERN_G;

// AutoIt2 supports lines up to 16384 characters long, and we want to be able to do so too
// so that really long lines from aut2 scripts, such as a chain of IF commands, can be
// brought in and parsed:
#define LINE_SIZE (16384 + 1)

enum ExecUntilMode {NORMAL_MODE, UNTIL_RETURN, UNTIL_BLOCK_END, ONLY_ONE_LINE};

// It's done this way so that mAttribute can store a pointer or one of these constants.
// If it is storing a pointer for a given Action Type, be sure never to compare it
// for equality against these constants because by coincidence, the pointer value
// might just match one of them:
#define ATTR_NONE ((void *)0)
#define ATTR_LOOP_UNKNOWN ((void *)1)
#define ATTR_LOOP_NORMAL ((void *)2)
#define ATTR_LOOP_FILE ((void *)3)
#define ATTR_LOOP_REG ((void *)4)
#define ATTR_LOOP_READ_FILE ((void *)5)
#define ATTR_LOOP_PARSE ((void *)6)
typedef void *AttributeType;

enum FileLoopModeType {FILE_LOOP_INVALID, FILE_LOOP_FILES_ONLY, FILE_LOOP_FILES_AND_FOLDERS, FILE_LOOP_FOLDERS_ONLY};
enum VariableTypeType {VAR_TYPE_INVALID, VAR_TYPE_NUMBER, VAR_TYPE_INTEGER, VAR_TYPE_FLOAT, VAR_TYPE_TIME
	, VAR_TYPE_DIGIT, VAR_TYPE_ALPHA, VAR_TYPE_ALNUM, VAR_TYPE_SPACE};

// But the array that goes with these actions is in globaldata.cpp because
// otherwise it would be a little cumbersome to declare the extern version
// of the array in here (since it's only extern to modules other than
// script.cpp):
enum enum_act {
// Seems best to make this one zero so that it will be the ZeroMemory() default within
// any POD structures that contain an action_type field:
  ACT_INVALID = FAIL  // These should both be zero for initialization and function-return-value purposes.
, ACT_ASSIGN, ACT_ADD, ACT_SUB, ACT_MULT, ACT_DIV
, ACT_ASSIGN_FIRST = ACT_ASSIGN, ACT_ASSIGN_LAST = ACT_DIV
, ACT_REPEAT // Never parsed directly, only provided as a translation target for the old command (see other notes).
, ACT_ELSE   // Parsed at a lower level than most commands to support same-line ELSE-actions (e.g. "else if").
, ACT_IFEQUAL, ACT_IFNOTEQUAL, ACT_IFGREATER, ACT_IFGREATEROREQUAL, ACT_IFLESS, ACT_IFLESSOREQUAL
, ACT_IFIS, ACT_IFISNOT
, ACT_FIRST_COMMAND // i.e the above aren't considered commands for parsing/searching purposes.
, ACT_IFWINEXIST = ACT_FIRST_COMMAND
, ACT_IFWINNOTEXIST, ACT_IFWINACTIVE, ACT_IFWINNOTACTIVE
, ACT_IFINSTRING, ACT_IFNOTINSTRING
, ACT_IFEXIST, ACT_IFNOTEXIST, ACT_IFMSGBOX
, ACT_IF_FIRST = ACT_IFEQUAL, ACT_IF_LAST = ACT_IFMSGBOX  // Keep this updated with any new IFs that are added.
, ACT_MSGBOX, ACT_INPUTBOX, ACT_SPLASHTEXTON, ACT_SPLASHTEXTOFF, ACT_TOOLTIP, ACT_TRAYTIP, ACT_INPUT
, ACT_TRANSFORM, ACT_STRINGLEFT, ACT_STRINGRIGHT, ACT_STRINGMID
, ACT_STRINGTRIMLEFT, ACT_STRINGTRIMRIGHT, ACT_STRINGLOWER, ACT_STRINGUPPER
, ACT_STRINGLEN, ACT_STRINGGETPOS, ACT_STRINGREPLACE, ACT_STRINGSPLIT
, ACT_ENVSET, ACT_ENVUPDATE
, ACT_RUNAS, ACT_RUN, ACT_RUNWAIT, ACT_URLDOWNLOADTOFILE
, ACT_GETKEYSTATE
, ACT_SEND, ACT_CONTROLSEND, ACT_CONTROLCLICK, ACT_CONTROLMOVE, ACT_CONTROLFOCUS, ACT_CONTROLGETFOCUS
, ACT_CONTROLSETTEXT, ACT_CONTROLGETTEXT, ACT_CONTROL, ACT_CONTROLGET
, ACT_COORDMODE, ACT_SETDEFAULTMOUSESPEED, ACT_MOUSEMOVE, ACT_MOUSECLICK, ACT_MOUSECLICKDRAG, ACT_MOUSEGETPOS
, ACT_STATUSBARGETTEXT
, ACT_STATUSBARWAIT
, ACT_CLIPWAIT
, ACT_SLEEP, ACT_RANDOM
, ACT_GOTO, ACT_GOSUB, ACT_HOTKEY, ACT_SETTIMER, ACT_RETURN, ACT_EXIT
, ACT_LOOP, ACT_BREAK, ACT_CONTINUE
, ACT_BLOCK_BEGIN, ACT_BLOCK_END
, ACT_WINACTIVATE, ACT_WINACTIVATEBOTTOM
, ACT_WINWAIT, ACT_WINWAITCLOSE, ACT_WINWAITACTIVE, ACT_WINWAITNOTACTIVE
, ACT_WINMINIMIZE, ACT_WINMAXIMIZE, ACT_WINRESTORE
, ACT_WINHIDE, ACT_WINSHOW
, ACT_WINMINIMIZEALL, ACT_WINMINIMIZEALLUNDO
, ACT_WINCLOSE, ACT_WINKILL, ACT_WINMOVE, ACT_WINMENUSELECTITEM
, ACT_WINSET, ACT_WINSETTITLE, ACT_WINGETTITLE, ACT_WINGETCLASS, ACT_WINGETPOS, ACT_WINGETTEXT
, ACT_POSTMESSAGE, ACT_SENDMESSAGE
// Keep rarely used actions near the bottom for parsing/performance reasons:
, ACT_PIXELGETCOLOR, ACT_PIXELSEARCH
, ACT_GROUPADD, ACT_GROUPACTIVATE, ACT_GROUPDEACTIVATE, ACT_GROUPCLOSE
, ACT_DRIVESPACEFREE, ACT_DRIVEGET
, ACT_SOUNDGET, ACT_SOUNDSET, ACT_SOUNDGETWAVEVOLUME, ACT_SOUNDSETWAVEVOLUME, ACT_SOUNDPLAY
, ACT_FILEAPPEND, ACT_FILEREADLINE, ACT_FILEDELETE, ACT_FILERECYCLE, ACT_FILERECYCLEEMPTY
, ACT_FILEINSTALL, ACT_FILECOPY, ACT_FILEMOVE, ACT_FILECOPYDIR, ACT_FILEMOVEDIR
, ACT_FILECREATEDIR, ACT_FILEREMOVEDIR
, ACT_FILEGETATTRIB, ACT_FILESETATTRIB, ACT_FILEGETTIME, ACT_FILESETTIME
, ACT_FILEGETSIZE, ACT_FILEGETVERSION
, ACT_SETWORKINGDIR, ACT_FILESELECTFILE, ACT_FILESELECTFOLDER, ACT_FILECREATESHORTCUT
, ACT_INIREAD, ACT_INIWRITE, ACT_INIDELETE
, ACT_REGREAD, ACT_REGWRITE, ACT_REGDELETE
, ACT_SETKEYDELAY, ACT_SETMOUSEDELAY, ACT_SETWINDELAY, ACT_SETCONTROLDELAY, ACT_SETBATCHLINES, ACT_SETINTERRUPT
, ACT_SETTITLEMATCHMODE, ACT_SETFORMAT
, ACT_SUSPEND, ACT_PAUSE
, ACT_AUTOTRIM, ACT_STRINGCASESENSE, ACT_DETECTHIDDENWINDOWS, ACT_DETECTHIDDENTEXT, ACT_BLOCKINPUT
, ACT_SETNUMLOCKSTATE, ACT_SETSCROLLLOCKSTATE, ACT_SETCAPSLOCKSTATE, ACT_SETSTORECAPSLOCKMODE
, ACT_KEYHISTORY, ACT_LISTLINES, ACT_LISTVARS, ACT_LISTHOTKEYS
, ACT_EDIT, ACT_RELOAD, ACT_MENU
, ACT_EXITAPP
, ACT_SHUTDOWN
// Make these the last ones before the count so they will be less often processed.  This helps
// performance because this one doesn't actually have a keyword so will never result
// in a match anyway.  UPDATE: No longer used because Run/RunWait is now required, which greatly
// improves syntax checking during load:
//, ACT_EXEC
// It's safer not to do this here.  It's better set by a
// calculation immediately after the array is declared and initialized,
// at which time we know its true size:
// , ACT_COUNT
};

enum enum_act_old {
  OLD_INVALID = FAIL  // These should both be zero for initialization and function-return-value purposes.
  , OLD_SETENV, OLD_ENVADD, OLD_ENVSUB, OLD_ENVMULT, OLD_ENVDIV
  , OLD_IFEQUAL, OLD_IFNOTEQUAL, OLD_IFGREATER, OLD_IFGREATEROREQUAL, OLD_IFLESS, OLD_IFLESSOREQUAL
  , OLD_LEFTCLICK, OLD_RIGHTCLICK, OLD_LEFTCLICKDRAG, OLD_RIGHTCLICKDRAG
  , OLD_REPEAT, OLD_ENDREPEAT
  , OLD_WINGETACTIVETITLE, OLD_WINGETACTIVESTATS
};

// It seems best not to include ACT_SUSPEND in the below, since the user may have marked
// a large number of subroutines as "Suspend, Permit".  Even PAUSE is iffy, since the user
// may be using it as "Pause, off/toggle", but it seems best to support PAUSE:
#define ACT_IS_ALWAYS_ALLOWED(ActionType) (ActionType == ACT_EXITAPP || ActionType == ACT_PAUSE \
	|| ActionType == ACT_EDIT || ActionType == ACT_RELOAD || ActionType == ACT_KEYHISTORY \
	|| ActionType == ACT_LISTLINES || ActionType == ACT_LISTVARS || ActionType == ACT_LISTHOTKEYS)
#define ACT_IS_IF(ActionType) (ActionType >= ACT_IF_FIRST && ActionType <= ACT_IF_LAST)
#define ACT_IS_ASSIGN(ActionType) (ActionType >= ACT_ASSIGN_FIRST && ActionType <= ACT_ASSIGN_LAST)

#define ATTACH_THREAD_INPUT \
	bool threads_are_attached = false;\
	DWORD my_thread  = GetCurrentThreadId();\
	DWORD target_thread = GetWindowThreadProcessId(target_window, NULL);\
	if (target_thread && target_thread != my_thread && !IsWindowHung(target_window))\
		threads_are_attached = AttachThreadInput(my_thread, target_thread, TRUE) != 0;

#define DETACH_THREAD_INPUT \
	if (threads_are_attached)\
		AttachThreadInput(my_thread, target_thread, FALSE);

// The starting ID is arbitrary, but it seems best to avoid low IDs to avoid any chance of
// a conflict with system menu id's, for example (though this is unlikely since I believe
// they all  generate WM_SYSCOMMAND vs. WM_COMMAND).  Also need to avoid a conflict with the
// main menu's IDs in resource.h.  ID_TRAY_USER is the id of the first user defined menu item
// (i.e. if the script uses the MENU command to customize the tray menu).
enum TrayMenuItems {ID_TRAY_OPEN = 16000, ID_TRAY_HELP, ID_TRAY_WINDOWSPY, ID_TRAY_RELOADSCRIPT
	, ID_TRAY_EDITSCRIPT, ID_TRAY_SUSPEND, ID_TRAY_PAUSE, ID_TRAY_EXIT, ID_TRAY_USER};


#define ERR_ABORT_NO_SPACES "The current thread will exit."
#define ERR_ABORT "  " ERR_ABORT_NO_SPACES
#define WILL_EXIT "The program will exit."
#define OLD_STILL_IN_EFFECT "The script was not reloaded; the old version will remain in effect."
#define PLEASE_REPORT "  Please report this as a bug."
#define ERR_UNRECOGNIZED_ACTION "This line does not contain a recognized action."
#define ERR_MISSING_OUTPUT_VAR "This command requires that at least one of its output variables be provided."
#define ERR_ELSE_WITH_NO_IF "This ELSE doesn't appear to belong to any IF-statement."
#define ERR_SETTIMER "This timer's target label does not exist."
#define ERR_HOTKEY_LABEL "Parameter #2 is not a valid label or action."
#define ERR_MENULABEL "This menu item's target label does not exist."
#define ERR_GROUPADD_LABEL "The target label in parameter #4 does not exist."
#define ERR_WINDOW_PARAM "This command requires that at least one of its window parameters be non-blank."
#define ERR_LOOP_FILE_MODE "If not blank or a variable reference, parameter #2 must be either 0, 1, 2."
#define ERR_LOOP_REG_MODE  "If not blank or a variable reference, parameter #3 must be either 0, 1, 2."
#define ERR_ON_OFF "If not blank or a variable reference, this parameter must be either ON or OFF."
#define ERR_ON_OFF_ALWAYS "If not blank or a variable reference, this parameter must be either ON, OFF, ALWAYSON, or ALWAYSOFF."
#define ERR_ON_OFF_TOGGLE "If not blank or a variable reference, this parameter must be either ON, OFF, or TOGGLE."
#define ERR_ON_OFF_TOGGLE_PERMIT "If not blank or a variable reference, this parameter must be either ON, OFF, TOGGLE, or PERMIT."
#define ERR_TITLEMATCHMODE "TitleMatchMode must be either 1, 2, 3, slow, fast, or a variable reference."
#define ERR_TITLEMATCHMODE2 "The variable does not contain a valid TitleMatchMode." ERR_ABORT
#define ERR_TRANSFORMCOMMAND "Parameter #2 is not a valid transform command."
#define ERR_MENUCOMMAND "Parameter #2 is not a valid menu command."
#define ERR_MENUCOMMAND2 "Parameter #2's variable does not contain a valid menu command." ERR_ABORT
#define ERR_CONTROLCOMMAND "Parameter #1 is not a valid Control command."
#define ERR_CONTROLGETCOMMAND "Parameter #2 is not a valid ControlGet command."
#define ERR_DRIVECOMMAND "Parameter #2 is not a valid DriveGet command."
#define ERR_WINSET "Parameter #1 is not a valid WinSet attribute."
#define ERR_IFMSGBOX "This line specifies an invalid MsgBox result."
#define ERR_REG_KEY "The key name must be either HKEY_LOCAL_MACHINE, HKEY_USERS, HKEY_CURRENT_USER, HKEY_CLASSES_ROOT, HKEY_CURRENT_CONFIG, or the abbreviations for these."
#define ERR_REG_VALUE_TYPE "The value type must be either REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ, REG_DWORD, or REG_BINARY."
#define ERR_RUN_SHOW_MODE "Parameter #3 must be either blank, a variable reference, or one of these words: MIN, MAX, HIDE."
#define ERR_COMPARE_TIMES "Parameter #3 must be either blank, a variable reference, or one of these words: Seconds, Minutes, Hours, Days."
#define ERR_INVALID_DATETIME "This date-time string contains at least one invalid component."
#define ERR_FILE_TIME "Parameter #3 must be either blank, M, C, A, or a variable reference."
#define ERR_MOUSE_BUTTON "This line specifies an invalid mouse button."
#define ERR_MOUSE_COORD "The X & Y coordinates must be either both absent or both present."
#define ERR_MOUSE_UPDOWN "Parameter #6 must be either blank, D, U, or a variable reference."
#define ERR_DIVIDEBYZERO "This line would attempt to divide by zero."
#define ERR_PERCENT "Parameter #1 must be a number between -100 and 100 (inclusive), or a variable reference."
#define ERR_MOUSE_SPEED "The Mouse Speed must be a number between 0 and " MAX_MOUSE_SPEED_STR ", blank, or a variable reference."
#define ERR_MEM_ASSIGN "Out of memory while assigning to this variable." ERR_ABORT
#define ERR_VAR_IS_RESERVED "This variable is reserved and cannot be assigned to."
#define ERR_DEFINE_CHAR "The character being defined must not be identical to another special or reserved character."
#define ERR_INCLUDE_FILE "A filename must be specified for #Include."
#define ERR_DEFINE_COMMENT "The comment flag must not be one of the hotkey definition symbols (e.g. ! ^ + $ ~ * < >)."

//----------------------------------------------------------------------------------

bool Util_Shutdown(int nFlag);
BOOL Util_ShutdownHandler(HWND hwnd, DWORD lParam);
void Util_WinKill(HWND hWnd);

enum MainWindowModes {MAIN_MODE_NO_CHANGE, MAIN_MODE_LINES, MAIN_MODE_VARS
	, MAIN_MODE_HOTKEYS, MAIN_MODE_KEYHISTORY, MAIN_MODE_REFRESH};
ResultType ShowMainWindow(MainWindowModes aMode = MAIN_MODE_NO_CHANGE, bool aRestricted = true);
ResultType GetAHKInstallDir(char *aBuf);


struct InputBoxType
{
	char *title;
	char *text;
	int width;
	int height;
	int xpos;
	int ypos;
	Var *output_var;
	char password_char;
	char *default_string;
	DWORD timeout;
	HWND hwnd;
};

// From AutoIt3's inputbox:
template <class T>
inline void swap(T &v1, T &v2) {
	T tmp=v1;
	v1=v2;
	v2=tmp;
}

#define INPUTBOX_DEFAULT INT_MIN
ResultType InputBox(Var *aOutputVar, char *aTitle, char *aText, bool aHideInput
	, int aWidth, int aHeight, int aX, int aY, double aTimeout, char *aDefault);
BOOL CALLBACK InputBoxProc(HWND hWndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
VOID CALLBACK InputBoxTimeout(HWND hWnd, UINT uMsg, UINT idEvent, DWORD dwTime);
BOOL CALLBACK EnumChildFocusFind(HWND aWnd, LPARAM lParam);
BOOL CALLBACK EnumChildGetText(HWND aWnd, LPARAM lParam);
LRESULT CALLBACK MainWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam);


typedef UINT LineNumberType;
#define LOADING_FAILED UINT_MAX

// -2 for the beginning and ending g_DerefChars:
#define MAX_VAR_NAME_LENGTH (UCHAR_MAX - 2)
#define MAX_DEREFS_PER_ARG 512
typedef UCHAR DerefLengthType;
struct DerefType
{
	char *marker;
	Var *var;
	DerefLengthType length;
};

typedef UCHAR ArgTypeType;  // UCHAR vs. an enum, to save memory.
#define ARG_TYPE_NORMAL     (UCHAR)0
#define ARG_TYPE_INPUT_VAR  (UCHAR)1
#define ARG_TYPE_OUTPUT_VAR (UCHAR)2
struct ArgStruct
{
	ArgTypeType type;
	char *text;
	DerefType *deref;  // Will hold a NULL-terminated array of var-deref locations within <text>.
};


// Some of these lengths and such are based on the MSDN example at
// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/sysinfo/base/enumerating_registry_subkeys.asp:
#define MAX_KEY_LENGTH 255
#define MAX_VALUE_NAME 16383
#define REG_SUBKEY -2 // Custom type, not standard in Windows.
struct RegItemStruct
{
	HKEY root_key_type, root_key;  // root_key_type is always a local HKEY, whereas root_key can be a remote handle.
	char subkey[MAX_PATH];  // The branch of the registry where this subkey or value is located.
	char name[MAX_VALUE_NAME]; // The subkey or value name.
	DWORD type; // Value Type (e.g REG_DWORD). This is the length used by MSDN in their example code.
	FILETIME ftLastWriteTime; // Non-initialized.
	void InitForValues() {ftLastWriteTime.dwHighDateTime = ftLastWriteTime.dwLowDateTime = 0;}
	void InitForSubkeys() {type = REG_SUBKEY;}  // To distinguish REG_DWORD and such from the subkeys themselves.
	RegItemStruct(HKEY aRootKeyType, HKEY aRootKey, char *aSubKey)
		: root_key_type(aRootKeyType), root_key(aRootKey), type(REG_NONE)
	{
		*name = '\0';
		// Make a local copy on the caller's stack so that if the current script subroutine is
		// interrupted to allow another to run, the contents of the deref buffer is saved here:
		strlcpy(subkey, aSubKey, sizeof(subkey));
		// Even though the call may work with a trailing backslash, it's best to remove it
		// so that consistent results are delivered to the user.  For example, if the script
		// is enumerating recursively into a subkey, subkeys deeper down will not include the
		// trailing backslash when they are reported.  So the user's own subkey should not
		// have one either so that when A_ScriptSubKey is referenced in the script, it will
		// always show up as the value without a trailing backslash:
		size_t length = strlen(subkey);
		if (length && subkey[length - 1] == '\\')
			subkey[length - 1] = '\0';
	}
};

struct LoopReadFileStruct
{
	FILE *mReadFile, *mWriteFile;
	#define READ_FILE_LINE_SIZE (64 * 1024)  // This is also used by FileReadLine().
	char mCurrentLine[READ_FILE_LINE_SIZE];
	LoopReadFileStruct(FILE *aReadFile, FILE *aWriteFile)
		: mReadFile(aReadFile), mWriteFile(aWriteFile)
	{
		*mCurrentLine = '\0';
	}
};


typedef UCHAR ArgCountType;
#define MAX_ARGS 20


// Note that currently this value must fit into a sc_type variable because that is how TextToKey()
// stores it in the hotkey class.  sc_type is currently a UINT, and will always be at least a
// WORD in size, so it shouldn't be much of an issue:
#define MAX_JOYSTICKS 16  // The maximum allowed by any Windows operating system.
#define MAX_JOY_BUTTONS 32 // Also the max that Windows supports.
enum JoyControls {JOYCTRL_INVALID, JOYCTRL_XPOS, JOYCTRL_YPOS, JOYCTRL_ZPOS
, JOYCTRL_RPOS, JOYCTRL_UPOS, JOYCTRL_VPOS, JOYCTRL_POV
, JOYCTRL_NAME, JOYCTRL_BUTTONS, JOYCTRL_AXES, JOYCTRL_INFO
, JOYCTRL_1, JOYCTRL_2, JOYCTRL_3, JOYCTRL_4, JOYCTRL_5, JOYCTRL_6, JOYCTRL_7, JOYCTRL_8  // Buttons.
, JOYCTRL_9, JOYCTRL_10, JOYCTRL_11, JOYCTRL_12, JOYCTRL_13, JOYCTRL_14, JOYCTRL_15, JOYCTRL_16
, JOYCTRL_17, JOYCTRL_18, JOYCTRL_19, JOYCTRL_20, JOYCTRL_21, JOYCTRL_22, JOYCTRL_23, JOYCTRL_24
, JOYCTRL_25, JOYCTRL_26, JOYCTRL_27, JOYCTRL_28, JOYCTRL_29, JOYCTRL_30, JOYCTRL_31, JOYCTRL_32
, JOYCTRL_BUTTON_MAX = JOYCTRL_32
};

enum TransformCmds {TRANS_CMD_INVALID, TRANS_CMD_ASC, TRANS_CMD_CHR, TRANS_CMD_MOD
	, TRANS_CMD_POW, TRANS_CMD_EXP, TRANS_CMD_SQRT, TRANS_CMD_LOG, TRANS_CMD_LN
	, TRANS_CMD_ROUND, TRANS_CMD_CEIL, TRANS_CMD_FLOOR, TRANS_CMD_ABS
	, TRANS_CMD_SIN, TRANS_CMD_COS, TRANS_CMD_TAN, TRANS_CMD_ASIN, TRANS_CMD_ACOS, TRANS_CMD_ATAN
	, TRANS_CMD_BITAND, TRANS_CMD_BITOR, TRANS_CMD_BITXOR, TRANS_CMD_BITNOT
	, TRANS_CMD_BITSHIFTLEFT, TRANS_CMD_BITSHIFTRIGHT
};

enum MenuCommands {MENU_COMMAND_INVALID, MENU_COMMAND_ADD, MENU_COMMAND_RENAME
	, MENU_COMMAND_CHECK, MENU_COMMAND_UNCHECK, MENU_COMMAND_TOGGLECHECK
	, MENU_COMMAND_ENABLE, MENU_COMMAND_DISABLE, MENU_COMMAND_TOGGLEENABLE
	, MENU_COMMAND_STANDARD, MENU_COMMAND_NOSTANDARD
	, MENU_COMMAND_DEFAULT, MENU_COMMAND_NODEFAULT
	, MENU_COMMAND_DELETE, MENU_COMMAND_DELETEALL
	, MENU_COMMAND_ICON, MENU_COMMAND_NOICON, MENU_COMMAND_MAINWINDOW, MENU_COMMAND_NOMAINWINDOW
};

enum ControlCmds {CONTROL_CMD_INVALID, CONTROL_CMD_CHECK, CONTROL_CMD_UNCHECK
	, CONTROL_CMD_ENABLE, CONTROL_CMD_DISABLE, CONTROL_CMD_SHOW, CONTROL_CMD_HIDE
	, CONTROL_CMD_SHOWDROPDOWN, CONTROL_CMD_HIDEDROPDOWN
	, CONTROL_CMD_TABLEFT, CONTROL_CMD_TABRIGHT
	, CONTROL_CMD_ADD, CONTROL_CMD_DELETE, CONTROL_CMD_CHOOSE
	, CONTROL_CMD_CHOOSESTRING, CONTROL_CMD_EDITPASTE};

enum ControlGetCmds {CONTROLGET_CMD_INVALID, CONTROLGET_CMD_CHECKED, CONTROLGET_CMD_ENABLED
	, CONTROLGET_CMD_VISIBLE, CONTROLGET_CMD_TAB, CONTROLGET_CMD_FINDSTRING
	, CONTROLGET_CMD_CHOICE, CONTROLGET_CMD_LINECOUNT, CONTROLGET_CMD_CURRENTLINE
	, CONTROLGET_CMD_CURRENTCOL, CONTROLGET_CMD_LINE, CONTROLGET_CMD_SELECTED};

enum DriveCmds {DRIVE_CMD_INVALID, DRIVE_CMD_LIST, DRIVE_CMD_FILESYSTEM, DRIVE_CMD_LABEL
	, DRIVE_CMD_SETLABEL, DRIVE_CMD_SERIAL, DRIVE_CMD_TYPE, DRIVE_CMD_STATUS, DRIVE_CMD_CAPACITY};

enum WinSetAttributes {WINSET_INVALID, WINSET_TRANSPARENT, WINSET_ALWAYSONTOP};



class Line
{
private:
	static char *sDerefBuf;  // Buffer to hold the values of any args that need to be dereferenced.
	static char *sDerefBufMarker;  // Location of next free byte.
	static size_t sDerefBufSize;

	// Static because only one line can be Expanded at a time (not to mention the fact that we
	// wouldn't want the size of each line to be expanded by this size):
	static char *sArgDeref[MAX_ARGS];

	ResultType EvaluateCondition();
	ResultType PerformLoop(WIN32_FIND_DATA *apCurrentFile, RegItemStruct *apCurrentRegItem
		, LoopReadFileStruct *apCurrentReadFile, char *aCurrentField, bool &aContinueMainLoop, Line *&aJumpToLine
		, AttributeType aAttr, FileLoopModeType aFileLoopMode, bool aRecurseSubfolders, char *aFilePattern
		, __int64 aIterationLimit, bool aIsInfinite);
	ResultType PerformLoopReg(WIN32_FIND_DATA *apCurrentFile, LoopReadFileStruct *apCurrentReadFile
		, char *aCurrentField, bool &aContinueMainLoop, Line *&aJumpToLine, FileLoopModeType aFileLoopMode
		, bool aRecurseSubfolders, HKEY aRootKeyType, HKEY aRootKey, char *aRegSubkey);
	ResultType PerformLoopParse(WIN32_FIND_DATA *apCurrentFile, RegItemStruct *apCurrentRegItem
		, LoopReadFileStruct *apCurrentReadFile, bool &aContinueMainLoop, Line *&aJumpToLine);
	ResultType Line::PerformLoopParseCSV(WIN32_FIND_DATA *apCurrentFile, RegItemStruct *apCurrentRegItem
		, LoopReadFileStruct *apCurrentReadFile, bool &aContinueMainLoop, Line *&aJumpToLine);
	ResultType PerformLoopReadFile(WIN32_FIND_DATA *apCurrentFile, RegItemStruct *apCurrentRegItem
		, char *aCurrentField, bool &aContinueMainLoop, Line *&aJumpToLine, FILE *aReadFile, FILE *aWriteFile);
	ResultType Perform(WIN32_FIND_DATA *aCurrentFile = NULL, RegItemStruct *aCurrentRegItem = NULL
		, LoopReadFileStruct *aCurrentReadFile = NULL);
	ResultType PerformAssign();
	ResultType StringSplit(char *aArrayName, char *aInputString, char *aDelimiterList, char *aOmitList);
	ResultType ScriptGetKeyState(char *aKeyName, char *aOption);
	ResultType DriveSpace(char *aPath, bool aGetFreeSpace);
	ResultType DriveGet(char *aCmd, char *aValue);
	ResultType SoundSetGet(char *aSetting, DWORD aComponentType, int aComponentInstance
		, DWORD aControlType, UINT aMixerID);
	ResultType SoundGetWaveVolume(HWAVEOUT aDeviceID);
	ResultType SoundSetWaveVolume(char *aVolume, HWAVEOUT aDeviceID);
	ResultType SoundPlay(char *aFilespec, bool aSleepUntilDone);
	ResultType URLDownloadToFile(char *aURL, char *aFilespec);
	ResultType FileSelectFile(char *aOptions, char *aWorkingDir, char *aGreeting, char *aFilter);

	// Bitwise flags:
	#define FSF_ALLOW_CREATE 0x01
	#define FSF_EDITBOX      0x02
	ResultType FileSelectFolder(char *aRootDir, DWORD aOptions, char *aGreeting);

	ResultType FileCreateShortcut(char *aTargetFile, char *aShortcutFile, char *aWorkingDir, char *aArgs
		, char *aDescription, char *aIconFile, char *aHotkey);
	ResultType FileCreateDir(char *aDirSpec);
	ResultType FileReadLine(char *aFilespec, char *aLineNumber);
	ResultType FileAppend(char *aFilespec, char *aBuf, FILE *aTargetFileAlreadyOpen = NULL);
	ResultType FileDelete(char *aFilePattern);
	ResultType FileRecycle(char *aFilePattern);
	ResultType FileRecycleEmpty(char *aDriveLetter);
	ResultType FileInstall(char *aSource, char *aDest, char *aFlag);

	// AutoIt3 functions:
	bool Util_CopyDir (const char *szInputSource, const char *szInputDest, bool bOverwrite);
	bool Util_MoveDir (const char *szInputSource, const char *szInputDest, bool bOverwrite);
	bool Util_RemoveDir (const char *szInputSource, bool bRecurse);
	int Util_CopyFile(const char *szInputSource, const char *szInputDest, bool bOverwrite, bool bMove);
	void Util_ExpandFilenameWildcard(const char *szSource, const char *szDest, char *szExpandedDest);
	void Util_ExpandFilenameWildcardPart(const char *szSource, const char *szDest, char *szExpandedDest);
	bool Util_CreateDir(const char *szDirName);
	bool Util_DoesFileExist(const char *szFilename);
	bool Util_IsDir(const char *szPath);
	void Util_GetFullPathName(const char *szIn, char *szOut);
	void Util_StripTrailingDir(char *szPath);
	bool Util_IsDifferentVolumes(const char *szPath1, const char *szPath2);

	ResultType FileGetAttrib(char *aFilespec);
	int FileSetAttrib(char *aAttributes, char *aFilePattern, FileLoopModeType aOperateOnFolders
		, bool aDoRecurse, bool aCalledRecursively = false);
	ResultType FileGetTime(char *aFilespec, char aWhichTime);
	int FileSetTime(char *aYYYYMMDD, char *aFilePattern, char aWhichTime
		, FileLoopModeType aOperateOnFolders, bool aDoRecurse, bool aCalledRecursively = false);
	ResultType FileGetSize(char *aFilespec, char *aGranularity);
	ResultType FileGetVersion(char *aFilespec);

	ResultType IniRead(char *aFilespec, char *aSection, char *aKey, char *aDefault);
	ResultType IniWrite(char *aValue, char *aFilespec, char *aSection, char *aKey);
	ResultType IniDelete(char *aFilespec, char *aSection, char *aKey);
	ResultType RegRead(HKEY aRootKey, char *aRegSubkey, char *aValueName);
	ResultType RegWrite(DWORD aValueType, HKEY aRootKey, char *aRegSubkey, char *aValueName, char *aValue);
	ResultType RegDelete(HKEY aRootKey, char *aRegSubkey, char *aValueName);
	static bool RegRemoveSubkeys(HKEY hRegKey);

	ResultType PerformShowWindow(ActionTypeType aActionType, char *aTitle = "", char *aText = ""
		, char *aExcludeTitle = "", char *aExcludeText = "");
	ResultType ToolTip(char *aText, char *aX, char *aY);
	ResultType TrayTip(char *aTitle, char *aText, char *aTimeout, char *aOptions);
	ResultType Transform(char *aCmd, char *aValue1, char *aValue2);
	ResultType Input(char *aOptions, char *aEndKeys, char *aMatchList);
	ResultType WinMove(char *aTitle, char *aText, char *aX, char *aY
		, char *aWidth = "", char *aHeight = "", char *aExcludeTitle = "", char *aExcludeText = "");
	ResultType WinMenuSelectItem(char *aTitle, char *aText, char *aMenu1, char *aMenu2
		, char *aMenu3, char *aMenu4, char *aMenu5, char *aMenu6, char *aMenu7
		, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlSend(char *aControl, char *aKeysToSend, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlClick(vk_type aVK, int aClickCount, char aEventType, char *aControl
		, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlMove(char *aControl, char *aX, char *aY, char *aWidth, char *aHeight
		, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlGetFocus(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlFocus(char *aControl, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlSetText(char *aControl, char *aNewText, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlGetText(char *aControl, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType Control(char *aCmd, char *aValue, char *aControl, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlGet(char *aCommand, char *aValue, char *aControl, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType StatusBarGetText(char *aPart, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType StatusBarWait(char *aTextToWaitFor, char *aSeconds, char *aPart, char *aTitle, char *aText
		, char *aInterval, char *aExcludeTitle, char *aExcludeText);
	ResultType ScriptPostMessage(char *aMsg, char *awParam, char *alParam, char *aControl
		, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType ScriptSendMessage(char *aMsg, char *awParam, char *alParam, char *aControl
		, char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType WinSet(char *aAttrib, char *aValue, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType WinSetTitle(char *aTitle, char *aText, char *aNewTitle
		, char *aExcludeTitle = "", char *aExcludeText = "");
	ResultType WinGetTitle(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType WinGetClass(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType WinGetText(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType WinGetPos(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType PixelSearch(int aLeft, int aTop, int aRight, int aBottom, int aColor, int aVariation);
	ResultType PixelGetColor(int aX, int aY);

	static ResultType SetToggleState(vk_type aVK, ToggleValueType &ForceLock, char *aToggleText);
	static ResultType MouseClickDrag(vk_type aVK // Which button.
		, int aX1, int aY1, int aX2, int aY2, int aSpeed, bool aMoveRelative);
	static ResultType MouseClick(vk_type aVK // Which button.
		, int aX = COORD_UNSPECIFIED, int aY = COORD_UNSPECIFIED // These values signal us not to move the mouse.
		, int aClickCount = 1, int aSpeed = DEFAULT_MOUSE_SPEED, char aEventType = '\0', bool aMoveRelative = false);
	static void MouseMove(int aX, int aY, int aSpeed = DEFAULT_MOUSE_SPEED, bool aMoveRelative = false);
	ResultType MouseGetPos();
	static void MouseEvent(DWORD aEventFlags, DWORD aX = 0, DWORD aY = 0, DWORD aData = 0)
	// A small inline to help us remember to use KEY_IGNORE so that our own mouse
	// events won't be falsely detected as hotkeys by the hooks (if they are installed).
	{
		mouse_event(aEventFlags, aX, aY, aData, KEY_IGNORE);
	}

public:
	ActionTypeType mActionType; // What type of line this is.
	UCHAR mFileNumber;  // Which file the line came from.  0 is the first, and it's the main script file.
	LineNumberType mLineNumber;  // The line number in the file from which the script was loaded, for debugging.
	AttributeType mAttribute;
	Line *mPrevLine, *mNextLine; // The prev & next lines adjacent to this one in the linked list; NULL if none.
	Line *mRelatedLine;  // e.g. the "else" that belongs to this "if"
	Line *mParentLine; // Indicates the parent (owner) of this line.
	// Probably best to always use ARG1 even if other things have supposedly verfied
	// that it exists, since it's count-check should make the dereference of a NULL
	// pointer (or accessing non-existent array elements) virtually impossible.
	// Empty-string is probably more universally useful than NULL, since some
	// API calls and other functions might not appreciate receiving NULLs.  In addition,
	// always remembering to have to check for NULL makes things harder to maintain
	// and more bug-prone.  The below macros rely upon the fact that the individual
	// elements of mArg cannot be NULL (because they're explicitly set to be blank
	// when the user has omitted an arg in between two non-blank args).  Later, might
	// want to review if any of the API calls used expect a string whose contents are
	// modifiable.
	#define RAW_ARG1 (mArgc > 0 ? mArg[0].text : "")
	#define RAW_ARG2 (mArgc > 1 ? mArg[1].text : "")
	#define RAW_ARG3 (mArgc > 2 ? mArg[2].text : "")
	#define RAW_ARG4 (mArgc > 3 ? mArg[3].text : "")
	#define RAW_ARG5 (mArgc > 4 ? mArg[4].text : "")
	#define RAW_ARG6 (mArgc > 5 ? mArg[5].text : "")
	#define RAW_ARG7 (mArgc > 6 ? mArg[6].text : "")
	#define RAW_ARG8 (mArgc > 7 ? mArg[7].text : "")
	#define LINE_RAW_ARG1 (line->mArgc > 0 ? line->mArg[0].text : "")
	#define LINE_RAW_ARG2 (line->mArgc > 1 ? line->mArg[1].text : "")
	#define LINE_RAW_ARG3 (line->mArgc > 2 ? line->mArg[2].text : "")
	#define LINE_RAW_ARG4 (line->mArgc > 3 ? line->mArg[3].text : "")
	#define LINE_RAW_ARG5 (line->mArgc > 4 ? line->mArg[4].text : "")
	#define LINE_RAW_ARG6 (line->mArgc > 5 ? line->mArg[5].text : "")
	#define LINE_RAW_ARG7 (line->mArgc > 6 ? line->mArg[6].text : "")
	#define LINE_RAW_ARG8 (line->mArgc > 7 ? line->mArg[7].text : "")
	#define LINE_RAW_ARG9 (line->mArgc > 8 ? line->mArg[8].text : "")
	#define LINE_ARG1 (line->mArgc > 0 ? line->sArgDeref[0] : "")
	#define LINE_ARG2 (line->mArgc > 1 ? line->sArgDeref[1] : "")
	#define LINE_ARG3 (line->mArgc > 2 ? line->sArgDeref[2] : "")
	#define LINE_ARG4 (line->mArgc > 3 ? line->sArgDeref[3] : "")
	#define SAVED_ARG1 (mArgc > 0 ? arg[0] : "")
	#define SAVED_ARG2 (mArgc > 1 ? arg[1] : "")
	#define SAVED_ARG3 (mArgc > 2 ? arg[2] : "")
	#define SAVED_ARG4 (mArgc > 3 ? arg[3] : "")
	#define SAVED_ARG5 (mArgc > 4 ? arg[4] : "")
	#define ARG1 (mArgc > 0 ? sArgDeref[0] : "")
	#define ARG2 (mArgc > 1 ? sArgDeref[1] : "")
	#define ARG3 (mArgc > 2 ? sArgDeref[2] : "")
	#define ARG4 (mArgc > 3 ? sArgDeref[3] : "")
	#define ARG5 (mArgc > 4 ? sArgDeref[4] : "")
	#define ARG6 (mArgc > 5 ? sArgDeref[5] : "")
	#define ARG7 (mArgc > 6 ? sArgDeref[6] : "")
	#define ARG8 (mArgc > 7 ? sArgDeref[7] : "")
	#define ARG9 (mArgc > 8 ? sArgDeref[8] : "")
	#define ARG10 (mArgc > 9 ? sArgDeref[9] : "")
	#define ARG11 (mArgc > 10 ? sArgDeref[10] : "")
	#define TWO_ARGS ARG1, ARG2
	#define THREE_ARGS  ARG1, ARG2, ARG3
	#define FOUR_ARGS   ARG1, ARG2, ARG3, ARG4
	#define FIVE_ARGS   ARG1, ARG2, ARG3, ARG4, ARG5
	#define SIX_ARGS    ARG1, ARG2, ARG3, ARG4, ARG5, ARG6
	#define SEVEN_ARGS  ARG1, ARG2, ARG3, ARG4, ARG5, ARG6, ARG7
	#define EIGHT_ARGS  ARG1, ARG2, ARG3, ARG4, ARG5, ARG6, ARG7, ARG8
	#define NINE_ARGS   ARG1, ARG2, ARG3, ARG4, ARG5, ARG6, ARG7, ARG8, ARG9
	#define TEN_ARGS    ARG1, ARG2, ARG3, ARG4, ARG5, ARG6, ARG7, ARG8, ARG9, ARG10
	#define ELEVEN_ARGS ARG1, ARG2, ARG3, ARG4, ARG5, ARG6, ARG7, ARG8, ARG9, ARG10, ARG11
	// If the arg's text is non-blank, it means the variable is a dynamic name such as array%i%
	// that must be resolved at runtime rather than load-time.  Therefore, this macro must not
	// be used without first having checked that arg.text is blank:
	#define VAR(arg) ((Var *)arg.deref)
	// Uses arg number (i.e. the first arg is 1, not 0).  Caller must ensure that ArgNum >= 1 and that
	// the arg in question is an input or output variable (since that isn't checked there):
	#define ARG_HAS_VAR(ArgNum) (mArgc >= ArgNum && (*mArg[ArgNum-1].text || mArg[ArgNum-1].deref))
	ArgCountType mArgc; // How many arguments exist in mArg[].
	ArgStruct *mArg; // Will be used to hold a dynamic array of dynamic Args.

	// Shouldn't go much higher than 200 since the main window's Edit control is currently limited
	// to 32K to be compatible with the Win9x limit.  Avg. line length is probably under 100 for
	// the vast majority of scripts, so 200 seems unlikely to exceed the buffer size.  Even in the
	// worst case where the buffer size is exceeded, the text is simply truncated, so it's not too
	// bad:
	#define LINE_LOG_SIZE 200
	static Line *sLog[LINE_LOG_SIZE];
	static DWORD sLogTick[LINE_LOG_SIZE];
	static int sLogNext;

	#define MAX_SCRIPT_FILES (UCHAR_MAX + 1)
	static char *sSourceFile[MAX_SCRIPT_FILES];
	static int nSourceFiles; // An int vs. UCHAR so that it can be exactly 256 without overflowing.

	ResultType ExecUntil(ExecUntilMode aMode, Line **apJumpToLine = NULL
		, WIN32_FIND_DATA *aCurrentFile = NULL, RegItemStruct *aCurrentRegItem = NULL
		, LoopReadFileStruct *aCurrentReadFile = NULL, char *aCurrentField = NULL
		, __int64 aCurrentLoopIteration = 0); // Use signed, since script/ITOA64 aren't designed to work with unsigned.

	Var *ResolveVarOfArg(int aArgIndex, bool aCreateIfNecessary = true);
	ResultType ExpandArgs();
	VarSizeType GetExpandedArgSize(bool aCalcDerefBufSize);
	char *ExpandArg(char *aBuf, int aArgIndex);

	bool FileIsFilteredOut(WIN32_FIND_DATA &aCurrentFile, FileLoopModeType aFileLoopMode, char *aFilePath);

	Line *GetJumpTarget(bool aIsDereferenced);
	ResultType IsJumpValid(Line *aDestination);

	static ArgTypeType ArgIsVar(ActionTypeType aActionType, int aArgIndex);
	static int ConvertEscapeChar(char *aFilespec, char aOldChar, char aNewChar, bool aFromAutoIt2 = false);
	static size_t ConvertEscapeCharGetLine(char *aBuf, int aMaxCharsToRead, FILE *fp);
	ResultType CheckForMandatoryArgs();

	bool ArgHasDeref(int aArgNum)
	// This function should always be called in lieu of doing something like "strchr(arg.text, g_DerefChar)"
	// because that method is unreliable due to the possible presence of literal (escaped) g_DerefChars
	// in the text.
	// aArgNum starts at 1 (for the first arg), so zero is invalid).
	{
		if (!aArgNum)
		{
			LineError("ArgHasDeref() called with a zero (bad habit to get into).", WARN);
			++aArgNum;  // But let it continue.
		}
		if (aArgNum > mArgc) // arg doesn't exist
			return false;
		if (mArg[aArgNum - 1].type != ARG_TYPE_NORMAL) // Always do this check prior to the next.
			return false;
		// Relies on short-circuit boolean evaluation order to prevent NULL-deref:
		return mArg[aArgNum - 1].deref && mArg[aArgNum - 1].deref[0].marker;
	}

	ResultType ArgMustBeDereferenced(Var *aVar, int aArgIndexToExclude)
	// Returns CONDITION_TRUE, CONDITION_FALSE, or FAIL.
	{
		if (aVar->mType == VAR_CLIPBOARD)
			// Even if the clipboard is both an input and an output var, it still
			// doesn't need to be dereferenced into the temp buffer because the
			// clipboard has two buffers of its own.  The only exception is when
			// the clipboard has only files on it, in which case those files need
			// to be converted into plain text:
			return CLIPBOARD_CONTAINS_ONLY_FILES ? CONDITION_TRUE : CONDITION_FALSE;
		if (aVar->mType != VAR_NORMAL || !aVar->Length())
			// Reserved vars must always be dereferenced due to their volatile nature.
			// Normal vars of length zero are dereferenced because they might exist
			// as system environment variables, whose contents are also potentially
			// volatile (i.e. they are sometimes changed by outside forces):
			return CONDITION_TRUE;
		// Since the above didn't return, we know that this is a NORMAL input var of
		// non-zero length.  Such input vars only need to be dereferenced if they are
		// also used as an output var by the current script line:
		Var *output_var;
		for (int iArg = 0; iArg < mArgc; ++iArg)
			if (iArg != aArgIndexToExclude && mArg[iArg].type == ARG_TYPE_OUTPUT_VAR)
			{
				if (   !(output_var = ResolveVarOfArg(iArg, false))   )
					return FAIL;  // It will have already displayed the error.
				if (output_var == aVar)
					return CONDITION_TRUE;
			}
		// Otherwise:
		return CONDITION_FALSE;
	}

	bool ArgAllowsNegative(int aArgNum)
	// aArgNum starts at 1 (for the first arg), so zero is invalid.
	{
		if (!aArgNum)
		{
			LineError("ArgAllowsNegative() called with a zero (bad habit to get into).");
			++aArgNum;  // But let it continue.
		}
		switch(mActionType)
		{
		case ACT_ADD:
		case ACT_SUB:
		case ACT_MULT:
		case ACT_DIV:
		case ACT_SETKEYDELAY:
		case ACT_SETMOUSEDELAY:
		case ACT_SETWINDELAY:
		case ACT_SETCONTROLDELAY:
		case ACT_RANDOM:
		case ACT_WINMOVE:
		case ACT_CONTROLMOVE:
		case ACT_SETINTERRUPT:
		case ACT_PIXELGETCOLOR:
		case ACT_TOOLTIP:  // Seems best to allow negative for this one even though the tip will be put in a visible region.
			return true;

		// Since mouse coords are relative to the foreground window, they can be negative:
		case ACT_MOUSECLICK:
			return (aArgNum == 2 || aArgNum == 3);
		case ACT_MOUSECLICKDRAG:
			return (aArgNum >= 2 && aArgNum <= 5);  // Allow dragging to/from negative coordinates.
		case ACT_MOUSEMOVE:
			return (aArgNum == 1 || aArgNum == 2);
		case ACT_PIXELSEARCH:
			return (aArgNum >= 3 || aArgNum <= 7); // i.e. Color values can be negative, but the last param cannot.
		case ACT_INPUTBOX:
			return (aArgNum == 7 || aArgNum == 8); // X & Y coords, even if they're absolute vs. relative.

		case ACT_SOUNDSET:
		case ACT_SOUNDSETWAVEVOLUME:
			return aArgNum == 1;
		}
		return false;  // Since above didn't return, negative is not allowed.
	}

	bool ArgAllowsFloat(int aArgNum)
	// aArgNum starts at 1 (for the first arg), so zero is invalid.
	{
		if (!aArgNum)
		{
			LineError("ArgAllowsFloat() called with a zero (bad habit to get into).");
			++aArgNum;  // But let it continue.
		}
		switch(mActionType)
		{
		case ACT_RANDOM:
		case ACT_MSGBOX:
		case ACT_WINCLOSE:
		case ACT_WINKILL:
		case ACT_WINWAIT:
		case ACT_WINWAITCLOSE:
		case ACT_WINWAITACTIVE:
		case ACT_WINWAITNOTACTIVE:
		case ACT_STATUSBARWAIT:
		case ACT_CLIPWAIT:
		// For these, allow even the variable (the first arg) to to be a float so that
		// the runtime checks won't catch it as an error:
		case ACT_ADD:
		case ACT_SUB:
		case ACT_MULT:
		case ACT_DIV:
			return true;

		case ACT_SOUNDSET:
		case ACT_SOUNDSETWAVEVOLUME:
			return aArgNum == 1;

		case ACT_INPUTBOX:
			return aArgNum == 10;
		}
		return false;  // Since above didn't return, negative is not allowed.
	}

	static HKEY RegConvertRootKey(char *aBuf, bool *aIsRemoteRegistry = NULL)
	{
		// Even if the computer name is a single letter, it seems like using a colon as delimiter is ok
		// (e.g. a:HKEY_LOCAL_MACHINE), since we wouldn't expect the root key to be used as a filename
		// in that exact way, i.e. a drive letter should be followed by a backslash 99% of the time in
		// this context.
		// Research indicates that colon is an illegal char in a computer name (at least for NT,
		// and by extension probably all other OSes).  So it should be safe to use it as a delimiter
		// for the remote registry feature.  But just in case, get the right-most one,
		// e.g. Computer:01:HKEY_LOCAL_MACHINE  ; the first colon is probably illegal on all OSes.
		// Additional notes from the Internet:
		// "A Windows NT computer name can be up to 15 alphanumeric characters with no blank spaces
		// and must be unique on the network. It can contain the following special characters:
		// ! @ # $ % ^ & ( ) -   ' { } .
		// It may not contain:
		// \ * + = | : ; " ? ,
		// The following is a list of illegal characters in a computer name:
		// regEx.Pattern = "`|~|!|@|#|\$|\^|\&|\*|\(|\)|\=|\+|{|}|\\|;|:|'|<|>|/|\?|\||%"

		char *colon_pos = strrchr(aBuf, ':');
		char *key_name = colon_pos ? omit_leading_whitespace(colon_pos + 1) : aBuf;
		if (aIsRemoteRegistry) // Caller wanted the below put into the output parameter.
			*aIsRemoteRegistry = (colon_pos != NULL);
		HKEY root_key = NULL; // Set default.
		if (!stricmp(key_name, "HKEY_LOCAL_MACHINE") || !stricmp(key_name, "HKLM"))  root_key = HKEY_LOCAL_MACHINE;
		if (!stricmp(key_name, "HKEY_CLASSES_ROOT") || !stricmp(key_name, "HKCR"))   root_key = HKEY_CLASSES_ROOT;
		if (!stricmp(key_name, "HKEY_CURRENT_CONFIG") || !stricmp(key_name, "HKCC")) root_key = HKEY_CURRENT_CONFIG;
		if (!stricmp(key_name, "HKEY_CURRENT_USER") || !stricmp(key_name, "HKCU"))   root_key = HKEY_CURRENT_USER;
		if (!stricmp(key_name, "HKEY_USERS") || !stricmp(key_name, "HKU"))           root_key = HKEY_USERS;
		if (!root_key)  // Invalid or unsupported root key name.
			return NULL;
		if (!aIsRemoteRegistry || !colon_pos) // Either caller didn't want it opened, or it doesn't need to be.
			return root_key; // If it's a remote key, this value should only be used by the caller as an indicator.
		// Otherwise, it's a remote computer whose registry the caller wants us to open:
		// It seems best to require the two leading backslashes in case the computer name contains
		// spaces (just in case spaces are allowed on some OSes or perhaps for Unix interoperability, etc.).
		// Therefore, make no attempt to trim leading and trailing spaces from the computer name:
		char computer_name[128];
		strlcpy(computer_name, aBuf, sizeof(computer_name));
		computer_name[colon_pos - aBuf] = '\0';
		HKEY remote_key;
		return (RegConnectRegistry(computer_name, root_key, &remote_key) == ERROR_SUCCESS) ? remote_key : NULL;
	}

	static char *RegConvertRootKey(char *aBuf, size_t aBufSize, HKEY aRootKey)
	{
		// switch() doesn't directly support expression of type HKEY:
		if (aRootKey == HKEY_LOCAL_MACHINE)       strlcpy(aBuf, "HKEY_LOCAL_MACHINE", aBufSize);
		else if (aRootKey == HKEY_CLASSES_ROOT)   strlcpy(aBuf, "HKEY_CLASSES_ROOT", aBufSize);
		else if (aRootKey == HKEY_CURRENT_CONFIG) strlcpy(aBuf, "HKEY_CURRENT_CONFIG", aBufSize);
		else if (aRootKey == HKEY_CURRENT_USER)   strlcpy(aBuf, "HKEY_CURRENT_USER", aBufSize);
		else if (aRootKey == HKEY_USERS)          strlcpy(aBuf, "HKEY_USERS", aBufSize);
		else if (aBufSize)                        *aBuf = '\0'; // Make it be the empty string for anything else.
		// These are either unused or so rarely used (DYN_DATA on Win9x) that they aren't supported:
		// HKEY_PERFORMANCE_DATA, HKEY_PERFORMANCE_TEXT, HKEY_PERFORMANCE_NLSTEXT, HKEY_DYN_DATA
		return aBuf;
	}
	static int RegConvertValueType(char *aValueType)
	{
		if (!stricmp(aValueType, "REG_SZ")) return REG_SZ;
		if (!stricmp(aValueType, "REG_EXPAND_SZ")) return REG_EXPAND_SZ;
		if (!stricmp(aValueType, "REG_MULTI_SZ")) return REG_MULTI_SZ;
		if (!stricmp(aValueType, "REG_DWORD")) return REG_DWORD;
		if (!stricmp(aValueType, "REG_BINARY")) return REG_BINARY;
		return REG_NONE; // Unknown or unsupported type.
	}
	static char *RegConvertValueType(char *aBuf, size_t aBufSize, DWORD aValueType)
	{
		switch(aValueType)
		{
		case REG_SZ: strlcpy(aBuf, "REG_SZ", aBufSize); return aBuf;
		case REG_EXPAND_SZ: strlcpy(aBuf, "REG_EXPAND_SZ", aBufSize); return aBuf;
		case REG_BINARY: strlcpy(aBuf, "REG_BINARY", aBufSize); return aBuf;
		case REG_DWORD: strlcpy(aBuf, "REG_DWORD", aBufSize); return aBuf;
		case REG_DWORD_BIG_ENDIAN: strlcpy(aBuf, "REG_DWORD_BIG_ENDIAN", aBufSize); return aBuf;
		case REG_LINK: strlcpy(aBuf, "REG_LINK", aBufSize); return aBuf;
		case REG_MULTI_SZ: strlcpy(aBuf, "REG_MULTI_SZ", aBufSize); return aBuf;
		case REG_RESOURCE_LIST: strlcpy(aBuf, "REG_RESOURCE_LIST", aBufSize); return aBuf;
		case REG_FULL_RESOURCE_DESCRIPTOR: strlcpy(aBuf, "REG_FULL_RESOURCE_DESCRIPTOR", aBufSize); return aBuf;
		case REG_RESOURCE_REQUIREMENTS_LIST: strlcpy(aBuf, "REG_RESOURCE_REQUIREMENTS_LIST", aBufSize); return aBuf;
		case REG_QWORD: strlcpy(aBuf, "REG_QWORD", aBufSize); return aBuf;
		case REG_SUBKEY: strlcpy(aBuf, "KEY", aBufSize); return aBuf;  // Custom (non-standard) type.
		default: if (aBufSize) *aBuf = '\0'; return aBuf;  // Make it be the empty string for REG_NONE and anything else.
		}
	}

	static DWORD SoundConvertComponentType(char *aBuf, int *aInstanceNumber = NULL)
	{
		char *colon_pos = strchr(aBuf, ':');
		UINT length_to_check = (UINT)(colon_pos ? colon_pos - aBuf : strlen(aBuf));
		if (aInstanceNumber) // Caller wanted the below put into the output parameter.
		{
			if (colon_pos)
			{
				*aInstanceNumber = ATOI(colon_pos + 1);
				if (*aInstanceNumber < 1)
					*aInstanceNumber = 1;
			}
			else
				*aInstanceNumber = 1;
		}
		if (!strlicmp(aBuf, "Master", length_to_check)) return MIXERLINE_COMPONENTTYPE_DST_SPEAKERS;
		if (!strlicmp(aBuf, "Speakers", length_to_check)) return MIXERLINE_COMPONENTTYPE_DST_SPEAKERS;
		if (!strlicmp(aBuf, "Digital", length_to_check)) return MIXERLINE_COMPONENTTYPE_SRC_DIGITAL;
		if (!strlicmp(aBuf, "Line", length_to_check)) return MIXERLINE_COMPONENTTYPE_SRC_LINE;
		if (!strlicmp(aBuf, "Microphone", length_to_check)) return MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE;
		if (!strlicmp(aBuf, "Synth", length_to_check)) return MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER;
		if (!strlicmp(aBuf, "CD", length_to_check)) return MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC;
		if (!strlicmp(aBuf, "Telephone", length_to_check)) return MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE;
		if (!strlicmp(aBuf, "PCSpeaker", length_to_check)) return MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER;
		if (!strlicmp(aBuf, "Wave", length_to_check)) return MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT;
		if (!strlicmp(aBuf, "Aux", length_to_check)) return MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY;
		if (!strlicmp(aBuf, "Analog", length_to_check)) return MIXERLINE_COMPONENTTYPE_SRC_ANALOG;
		return MIXERLINE_COMPONENTTYPE_DST_UNDEFINED;
	}
	static DWORD SoundConvertControlType(char *aBuf)
	{
		// These are the types that seem to correspond to actual sound attributes.  Some of the values
		// are not included here, such as MIXERCONTROL_CONTROLTYPE_FADER, which seems to be a type of
		// sound control rather than a quality of the sound itself.  For performance, put the most
		// often used ones up top:
		if (!stricmp(aBuf, "Vol")) return MIXERCONTROL_CONTROLTYPE_VOLUME;
		if (!stricmp(aBuf, "Volume")) return MIXERCONTROL_CONTROLTYPE_VOLUME;
		if (!stricmp(aBuf, "OnOff")) return MIXERCONTROL_CONTROLTYPE_ONOFF;
		if (!stricmp(aBuf, "Mute")) return MIXERCONTROL_CONTROLTYPE_MUTE;
		if (!stricmp(aBuf, "Mono")) return MIXERCONTROL_CONTROLTYPE_MONO;
		if (!stricmp(aBuf, "Loudness")) return MIXERCONTROL_CONTROLTYPE_LOUDNESS;
		if (!stricmp(aBuf, "StereoEnh")) return MIXERCONTROL_CONTROLTYPE_STEREOENH;
		if (!stricmp(aBuf, "BassBoost")) return MIXERCONTROL_CONTROLTYPE_BASS_BOOST;
		if (!stricmp(aBuf, "Pan")) return MIXERCONTROL_CONTROLTYPE_PAN;
		if (!stricmp(aBuf, "QSoundPan")) return MIXERCONTROL_CONTROLTYPE_QSOUNDPAN;
		if (!stricmp(aBuf, "Bass")) return MIXERCONTROL_CONTROLTYPE_BASS;
		if (!stricmp(aBuf, "Treble")) return MIXERCONTROL_CONTROLTYPE_TREBLE;
		if (!stricmp(aBuf, "Equalizer")) return MIXERCONTROL_CONTROLTYPE_EQUALIZER;
		#define MIXERCONTROL_CONTROLTYPE_INVALID 0 // 0 seems like a safe "undefined" indicator for this type.
		return MIXERCONTROL_CONTROLTYPE_INVALID;
	}

	static int ConvertJoy(char *aBuf, int *aJoystickID = NULL, bool aAllowOnlyButtons = false)
	{
		if (aJoystickID)
			*aJoystickID = 0;  // Set default output value for the caller.
		if (!aBuf || !*aBuf) return JOYCTRL_INVALID;
		char *aBuf_orig = aBuf;
		for (; *aBuf >= '0' && *aBuf <= '9'; ++aBuf); // self-contained loop to find the first non-digit.
		if (aBuf > aBuf_orig) // The string starts with a number.
		{
			int joystick_id = ATOI(aBuf_orig) - 1;
			if (joystick_id < 0 || joystick_id >= MAX_JOYSTICKS)
				return JOYCTRL_INVALID;
			if (aJoystickID)
				*aJoystickID = joystick_id;  // Use ATOI vs. atoi even though hex isn't supported yet.
		}

		if (!strnicmp(aBuf, "Joy", 3))
		{
			if (IsPureNumeric(aBuf + 3, false, false))
			{
				int offset = ATOI(aBuf + 3);
				if (offset < 1 || offset > MAX_JOY_BUTTONS)
					return JOYCTRL_INVALID;
				return JOYCTRL_1 + offset - 1;
			}
		}
		if (aAllowOnlyButtons)
			return JOYCTRL_INVALID;

		// Otherwise:
		if (!stricmp(aBuf, "JoyX")) return JOYCTRL_XPOS;
		if (!stricmp(aBuf, "JoyY")) return JOYCTRL_YPOS;
		if (!stricmp(aBuf, "JoyZ")) return JOYCTRL_ZPOS;
		if (!stricmp(aBuf, "JoyR")) return JOYCTRL_RPOS;
		if (!stricmp(aBuf, "JoyU")) return JOYCTRL_UPOS;
		if (!stricmp(aBuf, "JoyV")) return JOYCTRL_VPOS;
		if (!stricmp(aBuf, "JoyPOV")) return JOYCTRL_POV;
		if (!stricmp(aBuf, "JoyName")) return JOYCTRL_NAME;
		if (!stricmp(aBuf, "JoyButtons")) return JOYCTRL_BUTTONS;
		if (!stricmp(aBuf, "JoyAxes")) return JOYCTRL_AXES;
		if (!stricmp(aBuf, "JoyInfo")) return JOYCTRL_INFO;
		return JOYCTRL_INVALID;
	}

	static TitleMatchModes ConvertTitleMatchMode(char *aBuf)
	{
		if (!aBuf || !*aBuf) return MATCHMODE_INVALID;
		if (*aBuf == '1' && !*(aBuf + 1)) return FIND_IN_LEADING_PART;
		if (*aBuf == '2' && !*(aBuf + 1)) return FIND_ANYWHERE;
		if (*aBuf == '3' && !*(aBuf + 1)) return FIND_EXACT;
		if (!stricmp(aBuf, "FAST")) return FIND_FAST;
		if (!stricmp(aBuf, "SLOW")) return FIND_SLOW;
		return MATCHMODE_INVALID;
	}

	static TransformCmds ConvertTransformCmd(char *aBuf)
	{
		if (!aBuf || !*aBuf) return TRANS_CMD_INVALID;
		if (!stricmp(aBuf, "Asc")) return TRANS_CMD_ASC;
		if (!stricmp(aBuf, "Chr")) return TRANS_CMD_CHR;
		if (!stricmp(aBuf, "Mod")) return TRANS_CMD_MOD;
		if (!stricmp(aBuf, "Pow")) return TRANS_CMD_POW;
		if (!stricmp(aBuf, "Exp")) return TRANS_CMD_EXP;
		if (!stricmp(aBuf, "Sqrt")) return TRANS_CMD_SQRT;
		if (!stricmp(aBuf, "Log")) return TRANS_CMD_LOG;
		if (!stricmp(aBuf, "Ln")) return TRANS_CMD_LN;  // Natural log.
		if (!stricmp(aBuf, "Round")) return TRANS_CMD_ROUND;
		if (!stricmp(aBuf, "Ceil")) return TRANS_CMD_CEIL;
		if (!stricmp(aBuf, "Floor")) return TRANS_CMD_FLOOR;
		if (!stricmp(aBuf, "Abs")) return TRANS_CMD_ABS;
		if (!stricmp(aBuf, "Sin")) return TRANS_CMD_SIN;
		if (!stricmp(aBuf, "Cos")) return TRANS_CMD_COS;
		if (!stricmp(aBuf, "Tan")) return TRANS_CMD_TAN;
		if (!stricmp(aBuf, "ASin")) return TRANS_CMD_ASIN;
		if (!stricmp(aBuf, "ACos")) return TRANS_CMD_ACOS;
		if (!stricmp(aBuf, "ATan")) return TRANS_CMD_ATAN;
		if (!stricmp(aBuf, "BitAnd")) return TRANS_CMD_BITAND;
		if (!stricmp(aBuf, "BitOr")) return TRANS_CMD_BITOR;
		if (!stricmp(aBuf, "BitXOr")) return TRANS_CMD_BITXOR;
		if (!stricmp(aBuf, "BitNot")) return TRANS_CMD_BITNOT;
		if (!stricmp(aBuf, "BitShiftLeft")) return TRANS_CMD_BITSHIFTLEFT;
		if (!stricmp(aBuf, "BitShiftRight")) return TRANS_CMD_BITSHIFTRIGHT;
		return TRANS_CMD_INVALID;
	}

	static MenuCommands ConvertMenuCommand(char *aBuf)
	{
		if (!aBuf || !*aBuf) return MENU_COMMAND_INVALID;
		if (!stricmp(aBuf, "Add")) return MENU_COMMAND_ADD;
		if (!stricmp(aBuf, "Rename")) return MENU_COMMAND_RENAME;
		if (!stricmp(aBuf, "Check")) return MENU_COMMAND_CHECK;
		if (!stricmp(aBuf, "Uncheck")) return MENU_COMMAND_UNCHECK;
		if (!stricmp(aBuf, "ToggleCheck")) return MENU_COMMAND_TOGGLECHECK;
		if (!stricmp(aBuf, "Enable")) return MENU_COMMAND_ENABLE;
		if (!stricmp(aBuf, "Disable")) return MENU_COMMAND_DISABLE;
		if (!stricmp(aBuf, "ToggleEnable")) return MENU_COMMAND_TOGGLEENABLE;
		if (!stricmp(aBuf, "Standard")) return MENU_COMMAND_STANDARD;
		if (!stricmp(aBuf, "NoStandard")) return MENU_COMMAND_NOSTANDARD;
		if (!stricmp(aBuf, "Default")) return MENU_COMMAND_DEFAULT;
		if (!stricmp(aBuf, "NoDefault")) return MENU_COMMAND_NODEFAULT;
		if (!stricmp(aBuf, "Delete")) return MENU_COMMAND_DELETE;
		if (!stricmp(aBuf, "DeleteAll")) return MENU_COMMAND_DELETEALL;
		if (!stricmp(aBuf, "Icon")) return MENU_COMMAND_ICON;
		if (!stricmp(aBuf, "NoIcon")) return MENU_COMMAND_NOICON;
		if (!stricmp(aBuf, "MainWindow")) return MENU_COMMAND_MAINWINDOW;
		if (!stricmp(aBuf, "NoMainWindow")) return MENU_COMMAND_NOMAINWINDOW;
		return MENU_COMMAND_INVALID;
	}

	static ControlCmds ConvertControlCmd(char *aBuf)
	{
		if (!aBuf || !*aBuf) return CONTROL_CMD_INVALID;
		if (!stricmp(aBuf, "Check")) return CONTROL_CMD_CHECK;
		if (!stricmp(aBuf, "Uncheck")) return CONTROL_CMD_UNCHECK;
		if (!stricmp(aBuf, "Enable")) return CONTROL_CMD_ENABLE;
		if (!stricmp(aBuf, "Disable")) return CONTROL_CMD_DISABLE;
		if (!stricmp(aBuf, "Show")) return CONTROL_CMD_SHOW;
		if (!stricmp(aBuf, "Hide")) return CONTROL_CMD_HIDE;
		if (!stricmp(aBuf, "ShowDropDown")) return CONTROL_CMD_SHOWDROPDOWN;
		if (!stricmp(aBuf, "HideDropDown")) return CONTROL_CMD_HIDEDROPDOWN;
		if (!stricmp(aBuf, "TabLeft")) return CONTROL_CMD_TABLEFT;
		if (!stricmp(aBuf, "TabRight")) return CONTROL_CMD_TABRIGHT;
		if (!stricmp(aBuf, "Add")) return CONTROL_CMD_ADD;
		if (!stricmp(aBuf, "Delete")) return CONTROL_CMD_DELETE;
		if (!stricmp(aBuf, "Choose")) return CONTROL_CMD_CHOOSE;
		if (!stricmp(aBuf, "ChooseString")) return CONTROL_CMD_CHOOSESTRING;
		if (!stricmp(aBuf, "EditPaste")) return CONTROL_CMD_EDITPASTE;
		return CONTROL_CMD_INVALID;
	}

	static ControlGetCmds ConvertControlGetCmd(char *aBuf)
	{
		if (!aBuf || !*aBuf) return CONTROLGET_CMD_INVALID;
		if (!stricmp(aBuf, "Checked")) return CONTROLGET_CMD_CHECKED;
		if (!stricmp(aBuf, "Enabled")) return CONTROLGET_CMD_ENABLED;
		if (!stricmp(aBuf, "Visible")) return CONTROLGET_CMD_VISIBLE;
		if (!stricmp(aBuf, "Tab")) return CONTROLGET_CMD_TAB;
		if (!stricmp(aBuf, "FindString")) return CONTROLGET_CMD_FINDSTRING;
		if (!stricmp(aBuf, "Choice")) return CONTROLGET_CMD_CHOICE;
		if (!stricmp(aBuf, "LineCount")) return CONTROLGET_CMD_LINECOUNT;
		if (!stricmp(aBuf, "CurrentLine")) return CONTROLGET_CMD_CURRENTLINE;
		if (!stricmp(aBuf, "CurrentCol")) return CONTROLGET_CMD_CURRENTCOL;
		if (!stricmp(aBuf, "Line")) return CONTROLGET_CMD_LINE;
		if (!stricmp(aBuf, "Selected")) return CONTROLGET_CMD_SELECTED;
		return CONTROLGET_CMD_INVALID;
	}

	static DriveCmds ConvertDriveCmd(char *aBuf)
	{
		if (!aBuf || !*aBuf) return DRIVE_CMD_INVALID;
		if (!stricmp(aBuf, "List")) return DRIVE_CMD_LIST;
		if (!stricmp(aBuf, "FileSystem") || !stricmp(aBuf, "FS")) return DRIVE_CMD_FILESYSTEM;
		if (!stricmp(aBuf, "Label")) return DRIVE_CMD_LABEL;
		if (!strnicmp(aBuf, "SetLabel:", 9)) return DRIVE_CMD_SETLABEL;  // Uses strnicmp() vs. stricmp().
		if (!stricmp(aBuf, "Serial")) return DRIVE_CMD_SERIAL;
		if (!stricmp(aBuf, "Type")) return DRIVE_CMD_TYPE;
		if (!stricmp(aBuf, "Status")) return DRIVE_CMD_STATUS;
		if (!stricmp(aBuf, "Capacity") || !stricmp(aBuf, "Cap")) return DRIVE_CMD_CAPACITY;
		return DRIVE_CMD_INVALID;
	}

	static WinSetAttributes ConvertWinSetAttribute(char *aBuf)
	{
		if (!aBuf || !*aBuf) return WINSET_INVALID;
		if (!stricmp(aBuf, "Trans") || !stricmp(aBuf, "Transparent")) return WINSET_TRANSPARENT;
		if (!stricmp(aBuf, "AlwaysOnTop") || !stricmp(aBuf, "Topmost")) return WINSET_ALWAYSONTOP;
		return WINSET_INVALID;
	}


	static ToggleValueType ConvertOnOff(char *aBuf, ToggleValueType aDefault = TOGGLE_INVALID)
	// Returns aDefault if aBuf isn't either ON, OFF, or blank.
	{
		if (!aBuf || !*aBuf) return NEUTRAL;
		if (!stricmp(aBuf, "ON")) return TOGGLED_ON;
		if (!stricmp(aBuf, "OFF")) return TOGGLED_OFF;
		return aDefault;
	}

	static ToggleValueType ConvertOnOffAlways(char *aBuf, ToggleValueType aDefault = TOGGLE_INVALID)
	// Returns aDefault if aBuf isn't either ON, OFF, ALWAYSON, ALWAYSOFF, or blank.
	{
		if (!aBuf || !*aBuf) return NEUTRAL;
		if (!stricmp(aBuf, "On")) return TOGGLED_ON;
		if (!stricmp(aBuf, "Off")) return TOGGLED_OFF;
		if (!stricmp(aBuf, "AlwaysOn")) return ALWAYS_ON;
		if (!stricmp(aBuf, "AlwaysOff")) return ALWAYS_OFF;
		return aDefault;
	}

	static ToggleValueType ConvertOnOffToggle(char *aBuf, ToggleValueType aDefault = TOGGLE_INVALID)
	// Returns aDefault if aBuf isn't either ON, OFF, TOGGLE, or blank.
	{
		if (!aBuf || !*aBuf) return NEUTRAL;
		if (!stricmp(aBuf, "On")) return TOGGLED_ON;
		if (!stricmp(aBuf, "Off")) return TOGGLED_OFF;
		if (!stricmp(aBuf, "Toggle")) return TOGGLE;
		return aDefault;
	}

	static ToggleValueType ConvertOnOffTogglePermit(char *aBuf, ToggleValueType aDefault = TOGGLE_INVALID)
	// Returns aDefault if aBuf isn't either ON, OFF, TOGGLE, PERMIT, or blank.
	{
		if (!aBuf || !*aBuf) return NEUTRAL;
		if (!stricmp(aBuf, "On")) return TOGGLED_ON;
		if (!stricmp(aBuf, "Off")) return TOGGLED_OFF;
		if (!stricmp(aBuf, "Toggle")) return TOGGLE;
		if (!stricmp(aBuf, "Permit")) return TOGGLE_PERMIT;
		return aDefault;
	}

	static FileLoopModeType ConvertLoopMode(char *aBuf)
	// Returns the file loop mode, or FILE_LOOP_INVALID if aBuf contains an invalid mode.
	{
		if (!aBuf || !*aBuf)
			return FILE_LOOP_FILES_ONLY; // The default mode is used if the param is blank.
		if (strlen(aBuf) > 1)
			return FILE_LOOP_INVALID;
		if (*aBuf == '0')
			return FILE_LOOP_FILES_ONLY;
		if (*aBuf == '1')
			return FILE_LOOP_FILES_AND_FOLDERS;
		if (*aBuf == '2')
			return FILE_LOOP_FOLDERS_ONLY;
		return FILE_LOOP_INVALID;
	}

	static int ConvertMsgBoxResult(char *aBuf)
	// Returns the matching ID, or zero if none.
	{
		if (!aBuf || !*aBuf) return 0;
		// Keeping the most oft-used ones up top helps perf. a little:
		if (!stricmp(aBuf, "YES")) return IDYES;
		if (!stricmp(aBuf, "NO")) return IDNO;
		if (!stricmp(aBuf, "OK")) return IDOK;
		if (!stricmp(aBuf, "CANCEL")) return IDCANCEL;
		if (!stricmp(aBuf, "ABORT")) return IDABORT;
		if (!stricmp(aBuf, "IGNORE")) return IDIGNORE;
		if (!stricmp(aBuf, "RETRY")) return IDRETRY;
		if (!stricmp(aBuf, "TIMEOUT")) return AHK_TIMEOUT;  // Our custom result value.
		return 0;
	}

	static int ConvertRunMode(char *aBuf)
	// Returns the matching WinShow mode, or SW_SHOWNORMAL if none.
	// These are also the modes that AutoIt3 uses.
	{
		if (!aBuf || !*aBuf) return SW_SHOWNORMAL;
		if (!stricmp(aBuf, "MIN")) return SW_MINIMIZE;
		if (!stricmp(aBuf, "MAX")) return SW_MAXIMIZE;
		if (!stricmp(aBuf, "HIDE")) return SW_HIDE;
		return SW_SHOWNORMAL;
	}

	static int ConvertMouseButton(char *aBuf, bool aAllowWheel = true)
	// Returns the matching VK, or zero if none.
	{
		if (!aBuf || !*aBuf) return 0;
		if (!stricmp(aBuf, "LEFT") || !stricmp(aBuf, "L")) return VK_LBUTTON;
		if (!stricmp(aBuf, "RIGHT") || !stricmp(aBuf, "R")) return VK_RBUTTON;
		if (!stricmp(aBuf, "MIDDLE") || !stricmp(aBuf, "M")) return VK_MBUTTON;
		if (!stricmp(aBuf, "X1")) return VK_XBUTTON1;
		if (!stricmp(aBuf, "X2")) return VK_XBUTTON2;
		if (aAllowWheel)
		{
			if (!stricmp(aBuf, "WheelUp") || !stricmp(aBuf, "WU")) return VK_WHEEL_UP;
			if (!stricmp(aBuf, "WheelDown") || !stricmp(aBuf, "WD")) return VK_WHEEL_DOWN;
		}
		return 0;
	}

	static VariableTypeType ConvertVariableTypeName(char *aBuf)
	// Returns the matching type, or zero if none.
	{
		if (!aBuf || !*aBuf) return VAR_TYPE_INVALID;
		if (!stricmp(aBuf, "integer")) return VAR_TYPE_INTEGER;
		if (!stricmp(aBuf, "float")) return VAR_TYPE_FLOAT;
		if (!stricmp(aBuf, "number")) return VAR_TYPE_NUMBER;
		if (!stricmp(aBuf, "time")) return VAR_TYPE_TIME;
		if (!stricmp(aBuf, "date")) return VAR_TYPE_TIME;  // "date" is just an alias for "time".
		if (!stricmp(aBuf, "digit")) return VAR_TYPE_DIGIT;
		if (!stricmp(aBuf, "alpha")) return VAR_TYPE_ALPHA;
		if (!stricmp(aBuf, "alnum")) return VAR_TYPE_ALNUM;
		if (!stricmp(aBuf, "space")) return VAR_TYPE_SPACE;
		return VAR_TYPE_INVALID;
	}

	static ResultType ValidateMouseCoords(char *aX, char *aY)
	{
		if (!aX) aX = "";  if (!aY) aY = "";
		if (!*aX && !*aY) return OK;  // Both are absent, which is the signal to use the current position.
		if (*aX && *aY) return OK; // Both are present (that they are numeric is validated elsewhere).
		return FAIL; // Otherwise, one is blank but the other isn't, which is not allowed.
	}

	static char *LogToText(char *aBuf, size_t aBufSize);
	char *VicinityToText(char *aBuf, size_t aBufSize, int aMaxLines = 15);
	char *ToText(char *aBuf, size_t aBufSize, DWORD aElapsed = 0);

	static void ToggleSuspendState();
	ResultType ChangePauseState(ToggleValueType aChangeTo);
	ResultType Line::BlockInput(bool aEnable);

	Line *PreparseError(char *aErrorText);
	// Call this LineError to avoid confusion with Script's error-displaying functions:
	ResultType LineError(char *aErrorText, ResultType aErrorType = FAIL, char *aExtraInfo = "");

	Line(UCHAR aFileNumber, LineNumberType aFileLineNumber, ActionTypeType aActionType
		, ArgStruct aArg[], ArgCountType aArgc) // Constructor
		: mFileNumber(aFileNumber), mLineNumber(aFileLineNumber), mActionType(aActionType)
		, mAttribute(ATTR_NONE), mArgc(aArgc), mArg(aArg)
		, mPrevLine(NULL), mNextLine(NULL), mRelatedLine(NULL), mParentLine(NULL)
		{}
	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}  // Intentionally does nothing because we're using SimpleHeap for everything.
	void operator delete[](void *aPtr) {}
};



class Label
{
public:
	char *mName;
	Line *mJumpToLine;
	Label *mPrevLabel, *mNextLabel;  // Prev & Next items in linked list.

	Label(char *aLabelName)
		: mName(aLabelName) // Caller gave us a pointer to dynamic memory for this.
		, mJumpToLine(NULL)
		, mPrevLabel(NULL), mNextLabel(NULL)
	{}
	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}
};



class ScriptTimer
{
public:
	Label *mLabel;
	int mPeriod;
	UCHAR mExistingThreads;  // Whether this timer is already running its subroutine.
	DWORD mTimeLastRun;  // TickCount
	bool mEnabled;
	ScriptTimer *mNextTimer;  // Next items in linked list
	ScriptTimer(Label *aLabel)
		#define DEFAULT_TIMER_PERIOD 250
		: mLabel(aLabel), mPeriod(DEFAULT_TIMER_PERIOD), mExistingThreads(0), mTimeLastRun(0)
		, mEnabled(false), mNextTimer(NULL)  // Note that mEnabled must default to false for the counts to be right.
	{}
	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}
};



class UserMenuItem
{
public:
	#define MAX_MENU_LENGTH 100
	char mName[MAX_MENU_LENGTH + 1];
	UINT mMenuID;
	Label *mLabel;
	bool mEnabled, mChecked;
	UserMenuItem *mNextMenuItem;  // Next items in linked list

	UserMenuItem(char *aName, UINT aMenuID, Label *aLabel, HMENU aMenu) // Constructor
		: mMenuID(aMenuID), mLabel(aLabel), mEnabled(true), mChecked(false), mNextMenuItem(NULL)
	{
		strlcpy(mName, aName, sizeof(mName));
		if (aMenu)
			AppendMenu(aMenu, *aName ? MF_STRING : MF_SEPARATOR, aMenuID, aName);
	}

	// Don't overload new and delete operators in this case since we want to use real dynamic memory
	// (since menus can be read in from a file, destroyed and recreated, over and over).
};



class Script
{
private:
	friend class Hotkey;
	Line *mFirstLine, *mLastLine;  // The first and last lines in the linked list.
	Label *mFirstLabel, *mLastLabel;  // The first and last labels in the linked list.
	Var *mFirstVar, *mLastVar;  // The first and last variables in the linked list.
	WinGroup *mFirstGroup, *mLastGroup;  // The first and last variables in the linked list.
	UINT mLineCount, mLabelCount, mVarCount, mGroupCount;

#ifdef AUTOHOTKEYSC
	bool mCustomIcon; // Whether the compiled script uses a custom icon.
#endif;

	// These two track the file number and line number in that file of the line currently being loaded,
	// which simplifies calls to ScriptError() and LineError() (reduces the number of params that must be passed):
	UCHAR mCurrFileNumber;
	LineNumberType mCurrLineNumber;
	bool mNoHotkeyLabels;

	NOTIFYICONDATA mNIC; // For ease of adding and deleting our tray icon.
	HMENU mTrayMenu; // Our tray menu, which should be destroyed upon exiting the program.

#ifdef AUTOHOTKEYSC
	ResultType CloseAndReturn(HS_EXEArc_Read *fp, UCHAR *aBuf, ResultType return_value);
	size_t GetLine(char *aBuf, int aMaxCharsToRead, UCHAR *&aMemFile);
#else
	ResultType CloseAndReturn(FILE *fp, UCHAR *aBuf, ResultType return_value);
	size_t GetLine(char *aBuf, int aMaxCharsToRead, FILE *fp);
#endif
	ResultType IsPreprocessorDirective(char *aBuf);

	ResultType ParseAndAddLine(char *aLineText, char *aActionName = NULL
		, char *aEndMarker = NULL, char *aLiteralMap = NULL, size_t aLiteralMapLength = 0
		, ActionTypeType aActionType = ACT_INVALID, ActionTypeType aOldActionType = OLD_INVALID);
	char *ParseActionType(char *aBufTarget, char *aBufSource, bool aDisplayErrors);
	static ActionTypeType ConvertActionType(char *aActionTypeString);
	static ActionTypeType ConvertOldActionType(char *aActionTypeString);
	ResultType AddLabel(char *aLabelName);
	ResultType AddLine(ActionTypeType aActionType, char *aArg[] = NULL, ArgCountType aArgc = 0, char *aArgMap[] = NULL);
	Var *AddVar(char *aVarName, size_t aVarNameLength, Var *aVarPrev);

	// These aren't in the Line class because I think they're easier to implement
	// if aStartingLine is allowed to be NULL (for recursive calls).  If they
	// were member functions of class Line, a check for NULL would have to
	// be done before dereferencing any line's mNextLine, for example:
	Line *PreparseBlocks(Line *aStartingLine, bool aFindBlockEnd = false, Line *aParentLine = NULL);
	Line *PreparseIfElse(Line *aStartingLine, ExecUntilMode aMode = NORMAL_MODE, AttributeType aLoopTypeFile = ATTR_NONE
		, AttributeType aLoopTypeReg = ATTR_NONE, AttributeType aLoopTypeRead = ATTR_NONE
		, AttributeType aLoopTypeParse = ATTR_NONE);
	HMENU CreateTrayMenu();
	ResultType DestroyTrayMenu()
	{
		if (mTrayMenu)
		{
			// I think DestroyMenu() can fail if an attempt is made to destroy the menu while it is being
			// displayed (but even if it doesn't fail, it seems very bad to try to destroy it then, which
			// is why g_MenuIsVisible is checked just to be sure).
			// But this all should be impossible in our case because the script is in an uninterruptible state
			// while the menu is displayed, which in addition to pausing the current thread (which happens
			// anyway), no new timed or hotkey subroutines can be launched.  Thus, this should rarely if
			// ever happen, which is why no error message is given here:
			//if (g_MenuIsVisible)
			//	return FAIL;
			if (!DestroyMenu(mTrayMenu))
				return FAIL;
			mTrayMenu = NULL;
		}
		return OK;
	}
	void AppendStandardTrayItems();
public:
	Line *mCurrLine;  // Seems better to make this public than make Line our friend.
	char *mThisHotkeyName, *mPriorHotkeyName;

	ScriptTimer *mFirstTimer, *mLastTimer;  // The first and last script timers in the linked list.
	UINT mTimerCount, mTimerEnabledCount;

	UserMenuItem *mFirstMenuItem, *mLastMenuItem;
	UINT mMenuItemCount;

	WIN32_FIND_DATA *mLoopFile;  // The file of the current file-loop, if applicable.
	RegItemStruct *mLoopRegItem; // The registry subkey or value of the current registry enumeration loop.
	LoopReadFileStruct *mLoopReadFile;  // The file whose contents are currently being read by a File-Read Loop.
	char *mLoopField;  // The field of the current string-parsing loop.
	__int64 mLoopIteration; // Signed, since script/ITOA64 aren't designed to handle unsigned.
	DWORD mThisHotkeyStartTime, mPriorHotkeyStartTime;  // Tickcount timestamp of when its subroutine began.
	modLR_type mThisHotkeyModifiersLR;
	char *mFileSpec; // Will hold the full filespec, for convenience.
	char *mFileDir;  // Will hold the directory containing the script file.
	char *mFileName; // Will hold the script's naked file name.
	char *mOurEXE; // Will hold this app's module name (e.g. C:\Program Files\AutoHotkey\AutoHotkey.exe).
	char *mOurEXEDir;  // Same as above but just the containing diretory (for convenience).
	char *mMainWindowTitle; // Will hold our main window's title, for consistency & convenience.
	bool mIsReadyToExecute;
	bool AutoExecSectionIsRunning;
	bool mIsRestart; // The app is restarting rather than starting from scratch.
	bool mIsAutoIt2; // Whether this script is considered to be an AutoIt2 script.
	__int64 mLinesExecutedThisCycle; // Use 64-bit to match the type of g.LinesPerCycle
	int mUninterruptedLineCount; // Similar in purpose to the above, but only 32-bit.
	DWORD mLastScriptRest, mLastPeekTime;

	#define RUNAS_ITEM_SIZE (257 * sizeof(wchar_t))
	wchar_t *mRunAsUser, *mRunAsPass, *mRunAsDomain; // Memory is allocated at runtime, upon first use.

	bool mTrayIncludeStandard;  // Whether to include the tray menu's standard items in the menu.
	UserMenuItem *mTrayMenuDefault;
    
	ResultType Init(char *aScriptFilename, bool aIsRestart);
	ResultType CreateWindows();
	void CreateTrayIcon();
	void UpdateTrayIcon(bool aForceUpdate = false);
	void DisplayTrayMenu();
	ResultType AutoExecSection();
	ResultType PerformMenu(char *aWhichMenu, char *aCommand, char *aMenuItemName, char *aLabel);
	ResultType Edit();
	ResultType Reload(bool aDisplayErrors);
	void ExitApp(char *aBuf = NULL, int ExitCode = 0);
	LineNumberType LoadFromFile();
	ResultType LoadIncludedFile(char *aFileSpec, bool aAllowDuplicateInclude);
	ResultType UpdateOrCreateTimer(Label *aLabel, int aFreq, bool aEnable);
	#define VAR_NAME_LENGTH_DEFAULT 0
	Var *FindOrAddVar(char *aVarName, size_t aVarNameLength = VAR_NAME_LENGTH_DEFAULT, Var *aSearchStart = NULL);
	Var *FindVar(char *aVarName, size_t aVarNameLength = 0, Var **apVarPrev = NULL, Var *aSearchStart = NULL);
	WinGroup *FindOrAddGroup(char *aGroupName);
	ResultType AddGroup(char *aGroupName);
	Label *FindLabel(char *aLabelName);
	ResultType ActionExec(char *aAction, char *aParams = NULL, char *aWorkingDir = NULL
		, bool aDisplayErrors = true, char *aRunShowMode = NULL, HANDLE *aProcess = NULL, bool aUseRunAs = false);
	char *ListVars(char *aBuf, size_t aBufSize);
	char *ListKeyHistory(char *aBuf, size_t aBufSize);

	UserMenuItem *FindMenuItemByID(UINT aID)
	{
		for (UserMenuItem *menu_item = mFirstMenuItem; menu_item != NULL; menu_item = menu_item->mNextMenuItem)
			if (menu_item->mMenuID == aID)
				return menu_item;
		return NULL;
	}


	VarSizeType GetFilename(char *aBuf = NULL)
	{
		if (aBuf)
			strcpy(aBuf, mFileName);
		return (VarSizeType)strlen(mFileName);
	}
	VarSizeType GetFileDir(char *aBuf = NULL)
	{
		char str[MAX_PATH * 2] = "";  // Set default.
		strlcpy(str, mFileDir, sizeof(str));
		size_t length = strlen(str); // Needed not just for AutoIt2.
		// If it doesn't already have a final backslash, namely due to it being a root directory,
		// provide one so that it is backward compatible with AutoIt v2:
		if (mIsAutoIt2 && length && str[length - 1] != '\\')
		{
			str[length++] = '\\';
			str[length] = '\0';
		}
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)length;
	}
	VarSizeType GetFilespec(char *aBuf = NULL)
	{
		if (aBuf)
			sprintf(aBuf, "%s\\%s", mFileDir, mFileName);
		return (VarSizeType)(strlen(mFileDir) + strlen(mFileName) + 1);
	}

	VarSizeType GetLoopFileName(char *aBuf = NULL)
	{
		char *str = "";  // Set default.
		if (mLoopFile)
		{
			// The loop handler already prepended the script's directory in here for us:
			if (str = strrchr(mLoopFile->cFileName, '\\'))
				++str;
			else // No backslash, so just make it the entire file name.
				str = mLoopFile->cFileName;
		}
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}
	VarSizeType GetLoopFileShortName(char *aBuf = NULL)
	{
		char *str = "";  // Set default.
		if (mLoopFile)
		{
			if (   !*(str = mLoopFile->cAlternateFileName)   )
				// Files whose long name is shorter than the 8.3 usually don't have value stored here,
				// so use the long name whenever a short name is unavailable for any reason (could
				// also happen if NTFS has short-name generation disabled?)
				return GetLoopFileName(aBuf);
		}
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}
	VarSizeType GetLoopFileDir(char *aBuf = NULL)
	{
		char *str = "";  // Set default.
		char *last_backslash = NULL;
		if (mLoopFile)
		{
			// The loop handler already prepended the script's directory in here for us:
			str = mLoopFile->cFileName;
			if (last_backslash = strrchr(mLoopFile->cFileName, '\\'))
				*last_backslash = '\0'; // Temporarily terminate.
		}
		VarSizeType length = (VarSizeType)strlen(str);
		if (!aBuf)
		{
			if (last_backslash)
				*last_backslash = '\\';  // Restore the orginal value.
			return length;
		}
		strcpy(aBuf, str);
		if (last_backslash)
			*last_backslash = '\\';  // Restore the orginal value.
		return length;
	}
	VarSizeType GetLoopFileFullPath(char *aBuf = NULL)
	{
		char *str = "";  // Set default.
		if (mLoopFile)
			// The loop handler already prepended the script's directory in here for us:
			str = mLoopFile->cFileName;
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}
	VarSizeType GetLoopFileTimeModified(char *aBuf = NULL)
	{
		char str[64] = "";  // Set default.
		if (mLoopFile)
			FileTimeToYYYYMMDD(str, &mLoopFile->ftLastWriteTime, true);
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}
	VarSizeType GetLoopFileTimeCreated(char *aBuf = NULL)
	{
		char str[64] = "";  // Set default.
		if (mLoopFile)
			FileTimeToYYYYMMDD(str, &mLoopFile->ftCreationTime, true);
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}
	VarSizeType GetLoopFileTimeAccessed(char *aBuf = NULL)
	{
		char str[64] = "";  // Set default.
		if (mLoopFile)
			FileTimeToYYYYMMDD(str, &mLoopFile->ftLastAccessTime, true);
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}
	VarSizeType GetLoopFileAttrib(char *aBuf = NULL)
	{
		char str[64] = "";  // Set default.
		if (mLoopFile)
			FileAttribToStr(str, mLoopFile->dwFileAttributes);
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}
	VarSizeType GetLoopFileSize(char *aBuf, int aDivider)
	{
		char str[128];
		char *target_buf = aBuf ? aBuf : str;
		*target_buf = '\0';  // Set default.
		if (mLoopFile)
		{
			// It's a documented limitation that the size will show as negative if
			// greater than 2 gig, and will be wrong if greater than 4 gig.  For files
			// that large, scripts should use the KB version of this function instead.
			// If a file is over 4gig, set the value to be the maximum size (-1 when
			// expressed as a signed integer, since script variables are based entirely
			// on 32-bit signed integers due to the use of ATOI(), etc.).  UPDATE: 64-bit
			// ints are now standard, so the above is unnecessary:
			//sprintf(str, "%d%", mLoopFile->nFileSizeHigh ? -1 : (int)mLoopFile->nFileSizeLow);
			ULARGE_INTEGER ul;
			ul.HighPart = mLoopFile->nFileSizeHigh;
			ul.LowPart = mLoopFile->nFileSizeLow;
			ITOA64((__int64)(aDivider ? ((unsigned __int64)ul.QuadPart / aDivider) : ul.QuadPart), target_buf);
		}
		return (VarSizeType)strlen(target_buf);
	}

	VarSizeType GetLoopRegType(char *aBuf = NULL)
	{
		char str[256] = "";  // Set default.
		if (mLoopRegItem)
			Line::RegConvertValueType(str, sizeof(str), mLoopRegItem->type);
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}
	VarSizeType GetLoopRegKey(char *aBuf = NULL)
	{
		char str[256] = "";  // Set default.
		if (mLoopRegItem)
			// Use root_key_type, not root_key (which might be a remote vs. local HKEY):
			Line::RegConvertRootKey(str, sizeof(str), mLoopRegItem->root_key_type);
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}
	VarSizeType GetLoopRegSubKey(char *aBuf = NULL)
	{
		char *str = "";  // Set default.
		if (mLoopRegItem)
			str = mLoopRegItem->subkey;
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}
	VarSizeType GetLoopRegName(char *aBuf = NULL)
	{
		char *str = "";  // Set default.
		if (mLoopRegItem)
			str = mLoopRegItem->name; // This can be either the name of a subkey or the name of a value.
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}
	VarSizeType GetLoopRegTimeModified(char *aBuf = NULL)
	{
		char str[64] = "";  // Set default.
		// Only subkeys (not values) have a time.  In addition, Win9x doesn't support retrieval
		// of the time (nor does it store it), so make the var blank in that case:
		if (mLoopRegItem && mLoopRegItem->type == REG_SUBKEY && !g_os.IsWin9x())
			FileTimeToYYYYMMDD(str, &mLoopRegItem->ftLastWriteTime, true);
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}


	VarSizeType GetLoopReadLine(char *aBuf = NULL)
	{
		char *str = "";  // Set default.
		if (mLoopReadFile)
			str = mLoopReadFile->mCurrentLine;
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}

	VarSizeType GetLoopField(char *aBuf = NULL)
	{
		char *str = "";  // Set default.
		if (mLoopField)
			str = mLoopField;
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}

	VarSizeType GetLoopIndex(char *aBuf = NULL)
	{
		char str[64];
		ITOA64(mLoopIteration, str);
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}


	VarSizeType GetThisHotkey(char *aBuf = NULL)
	{
		if (aBuf)
			strcpy(aBuf, mThisHotkeyName);
		return (VarSizeType)strlen(mThisHotkeyName);
	}
	VarSizeType GetPriorHotkey(char *aBuf = NULL)
	{
		if (aBuf)
			strcpy(aBuf, mPriorHotkeyName);
		return (VarSizeType)strlen(mPriorHotkeyName);
	}
	VarSizeType GetTimeSinceThisHotkey(char *aBuf = NULL)
	{
		char str[128];
		// It must be the type of hotkey that has a label because we want the TimeSinceThisHotkey
		// value to be "in sync" with the value of ThisHotkey itself (i.e. use the same method
		// to determine which hotkey is the "this" hotkey):
		if (*mThisHotkeyName)
			// Even if GetTickCount()'s TickCount has wrapped around to zero and the timestamp hasn't,
			// DWORD math still gives the right answer as long as the number of days between
			// isn't greater than about 49.  See MyGetTickCount() for explanation of %d vs. %u.
			// Update: Using 64-bit ints now, so above is obsolete:
			//snprintf(str, sizeof(str), "%d", (DWORD)(GetTickCount() - mThisHotkeyStartTime));
			ITOA64((__int64)(GetTickCount() - mThisHotkeyStartTime), str)  // No semicolon
		else
			strcpy(str, "-1");
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}
	VarSizeType GetTimeSincePriorHotkey(char *aBuf = NULL)
	{
		char str[128];
		if (*mPriorHotkeyName)
			// See MyGetTickCount() for explanation for explanation:
			//snprintf(str, sizeof(str), "%d", (DWORD)(GetTickCount() - mPriorHotkeyStartTime));
			ITOA64((__int64)(GetTickCount() - mPriorHotkeyStartTime), str)
		else
			strcpy(str, "-1");
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}
	VarSizeType MyGetTickCount(char *aBuf = NULL)
	{
		// UPDATE: The below comments are now obsolete in light of having switched over to
		// using 64-bit integers (which aren't that much slower than 32-bit on 32-bit hardware):
		// Known limitation:
		// Although TickCount is an unsigned value, I'm not sure that our EnvSub command
		// will properly be able to compare two tick-counts if either value is larger than
		// INT_MAX.  So if the system has been up for more than about 25 days, there might be
		// problems if the user tries compare two tick-counts in the script using EnvSub.
		// UPDATE: It seems better to store all unsigned values as signed within script
		// variables.  Otherwise, when the var's value is next accessed and converted using
		// ATOI(), the outcome won't be as useful.  In other words, since the negative value
		// will be properly converted by ATOI(), comparing two negative tickcounts works
		// correctly (confirmed).  Even if one of them is negative and the other positive,
		// it will probably work correctly due to the nature of implicit unsigned math.
		// Thus, we use %d vs. %u in the snprintf() call below.
		char str[128];
		ITOA64(GetTickCount(), str);
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}
	VarSizeType GetTimeIdle(char *aBuf = NULL)
	{
		char str[128] = ""; // Set default.
		if (g_os.IsWin2000orLater())
		{
			// Must fetch it at runtime, otherwise the program can't even be launched on Win9x/NT:
			typedef BOOL (WINAPI *MyGetLastInputInfoType)(PLASTINPUTINFO);
			static MyGetLastInputInfoType MyGetLastInputInfo = (MyGetLastInputInfoType)
				GetProcAddress(GetModuleHandle("User32.dll"), "GetLastInputInfo");
			if (MyGetLastInputInfo)
			{
				LASTINPUTINFO lii;
				lii.cbSize = sizeof(lii);
				if (MyGetLastInputInfo(&lii))
					ITOA64(GetTickCount() - lii.dwTime, str);
			}
		}
		if (aBuf)
			strcpy(aBuf, str);
		return (VarSizeType)strlen(str);
	}

	VarSizeType GetTimeIdlePhysical(char *aBuf = NULL);
	VarSizeType ScriptGetCursor(char *aBuf = NULL);
	VarSizeType GetIP(int aAdapterIndex, char *aBuf = NULL);

	VarSizeType GetSpace(VarTypeType aType, char *aBuf = NULL)
	{
		if (!aBuf) return 1;  // i.e. the length of a single space char.
		*(aBuf++) = aType == VAR_SPACE ? ' ' : '\t';
		*aBuf = '\0';
		return 1;
	}

	// Call this SciptError to avoid confusion with Line's error-displaying functions:
	ResultType ScriptError(char *aErrorText, char *aExtraInfo = "");

	#define SOUNDPLAY_ALIAS "AHK_PlayMe"  // Used by destructor and SoundPlay().
	Script();
	~Script();
	// Note that the anchors to any linked lists will be lost when this
	// object goes away, so for now, be sure the destructor is only called
	// when the program is about to be exited, which will thereby reclaim
	// the memory used by the abandoned linked lists (otherwise, a memory
	// leak will result).
};

#endif
