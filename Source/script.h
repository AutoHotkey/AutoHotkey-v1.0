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
#include "Util.h" // for FileTimeToYYYYMMDD()
#include "resources\resource.h"  // For tray icon.
#ifdef AUTOHOTKEYSC
	#include "lib/exearc_read.h"
#endif

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
typedef void *AttributeType;

enum FileLoopModeType {FILE_LOOP_INVALID, FILE_LOOP_FILES_ONLY, FILE_LOOP_FILES_AND_FOLDERS, FILE_LOOP_FOLDERS_ONLY};

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
, ACT_FIRST_COMMAND // i.e the above aren't considered commands for parsing/searching purposes.
, ACT_IFWINEXIST = ACT_FIRST_COMMAND
, ACT_IFWINNOTEXIST, ACT_IFWINACTIVE, ACT_IFWINNOTACTIVE
, ACT_IFINSTRING, ACT_IFNOTINSTRING
, ACT_IFEXIST, ACT_IFNOTEXIST, ACT_IFMSGBOX
, ACT_IF_FIRST = ACT_IFEQUAL, ACT_IF_LAST = ACT_IFMSGBOX  // Keep this updated with any new IFs that are added.
, ACT_MSGBOX, ACT_INPUTBOX, ACT_SPLASHTEXTON, ACT_SPLASHTEXTOFF
, ACT_STRINGLEFT, ACT_STRINGRIGHT, ACT_STRINGMID
, ACT_STRINGTRIMLEFT, ACT_STRINGTRIMRIGHT, ACT_STRINGLOWER, ACT_STRINGUPPER
, ACT_STRINGLEN, ACT_STRINGGETPOS, ACT_STRINGREPLACE
, ACT_ENVSET, ACT_ENVUPDATE
, ACT_RUN, ACT_RUNWAIT
, ACT_GETKEYSTATE
, ACT_SEND, ACT_CONTROLSEND, ACT_CONTROLLEFTCLICK, ACT_CONTROLGETFOCUS, ACT_CONTROLFOCUS
, ACT_CONTROLSETTEXT, ACT_CONTROLGETTEXT
, ACT_SETDEFAULTMOUSESPEED, ACT_MOUSEMOVE, ACT_MOUSECLICK, ACT_MOUSECLICKDRAG, ACT_MOUSEGETPOS
, ACT_STATUSBARGETTEXT
, ACT_STATUSBARWAIT
, ACT_CLIPWAIT
, ACT_SLEEP, ACT_RANDOM
, ACT_GOTO, ACT_GOSUB, ACT_RETURN, ACT_EXIT
, ACT_LOOP, ACT_BREAK, ACT_CONTINUE
, ACT_BLOCK_BEGIN, ACT_BLOCK_END
, ACT_WINACTIVATE, ACT_WINACTIVATEBOTTOM
, ACT_WINWAIT, ACT_WINWAITCLOSE, ACT_WINWAITACTIVE, ACT_WINWAITNOTACTIVE
, ACT_WINMINIMIZE, ACT_WINMAXIMIZE, ACT_WINRESTORE
, ACT_WINHIDE, ACT_WINSHOW
, ACT_WINMINIMIZEALL, ACT_WINMINIMIZEALLUNDO
, ACT_WINCLOSE, ACT_WINKILL, ACT_WINMOVE, ACT_WINMENUSELECTITEM
, ACT_WINSETTITLE, ACT_WINGETTITLE, ACT_WINGETPOS, ACT_WINGETTEXT
// Keep rarely used actions near the bottom for parsing/performance reasons:
, ACT_PIXELGETCOLOR, ACT_PIXELSEARCH
, ACT_GROUPADD, ACT_GROUPACTIVATE, ACT_GROUPDEACTIVATE, ACT_GROUPCLOSE
, ACT_DRIVESPACEFREE, ACT_SOUNDSETWAVEVOLUME, ACT_SOUNDPLAY
, ACT_FILEAPPEND, ACT_FILEREADLINE, ACT_FILECOPY, ACT_FILEMOVE, ACT_FILEDELETE
, ACT_FILECREATEDIR, ACT_FILEREMOVEDIR
, ACT_FILEGETATTRIB, ACT_FILESETATTRIB, ACT_FILEGETTIME, ACT_FILESETTIME
, ACT_FILEGETSIZE, ACT_FILEGETVERSION
, ACT_FILESELECTFILE, ACT_FILESELECTFOLDER
, ACT_INIREAD, ACT_INIWRITE, ACT_INIDELETE
, ACT_REGREAD, ACT_REGWRITE, ACT_REGDELETE
, ACT_SETTITLEMATCHMODE, ACT_SETKEYDELAY, ACT_SETWINDELAY, ACT_SETCONTROLDELAY, ACT_SETBATCHLINES
, ACT_SUSPEND, ACT_PAUSE
, ACT_AUTOTRIM, ACT_STRINGCASESENSE, ACT_DETECTHIDDENWINDOWS, ACT_DETECTHIDDENTEXT
, ACT_SETNUMLOCKSTATE, ACT_SETSCROLLLOCKSTATE, ACT_SETCAPSLOCKSTATE, ACT_SETSTORECAPSLOCKMODE
, ACT_KEYLOG, ACT_LISTLINES, ACT_LISTVARS, ACT_LISTHOTKEYS
, ACT_EDIT, ACT_RELOADCONFIG
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
#define ACT_IS_IF(ActionType) (ActionType >= ACT_IF_FIRST && ActionType <= ACT_IF_LAST)
#define ACT_IS_ASSIGN(ActionType) (ActionType >= ACT_ASSIGN_FIRST && ActionType <= ACT_ASSIGN_LAST)

enum enum_act_old {
  OLD_INVALID = FAIL  // These should both be zero for initialization and function-return-value purposes.
  , OLD_SETENV, OLD_ENVADD, OLD_ENVSUB, OLD_ENVMULT, OLD_ENVDIV
  , OLD_IFEQUAL, OLD_IFNOTEQUAL, OLD_IFGREATER, OLD_IFGREATEROREQUAL, OLD_IFLESS, OLD_IFLESSOREQUAL
  , OLD_LEFTCLICK, OLD_RIGHTCLICK, OLD_LEFTCLICKDRAG, OLD_RIGHTCLICKDRAG
  , OLD_REPEAT, OLD_ENDREPEAT
  , OLD_WINGETACTIVETITLE, OLD_WINGETACTIVESTATS
};


#define ERR_ABORT "  The current hotkey subroutine (or the entire script if"\
	" this isn't a hotkey subroutine) will be aborted."
#define WILL_EXIT "The program will exit."
#define OLD_STILL_IN_EFFECT "The new config file was not loaded; the old config will remain in effect."
#define PLEASE_REPORT "  Please report this as a bug."
#define ERR_UNRECOGNIZED_ACTION "This line does not contain a recognized action."
#define ERR_MISSING_OUTPUT_VAR "This command requires that at least one of its output variables be provided."
#define ERR_ELSE_WITH_NO_IF "This ELSE doesn't appear to belong to any IF-statement."
#define ERR_GROUPADD_LABEL "The target label in parameter #4 does not exist."
#define ERR_WINDOW_PARAM "This command requires that at least one of its window parameters be non-blank."
#define ERR_LOOP_FILE_MODE "If not blank, parameter #2 must be either 0, 1, 2, or a dereferenced variable."
#define ERR_ON_OFF "If not blank, the value must be either ON, OFF, or a dereferenced variable."
#define ERR_ON_OFF_ALWAYS "If not blank, the value must be either ON, OFF, ALWAYSON, ALWAYSOFF, or a dereferenced variable."
#define ERR_ON_OFF_TOGGLE "If not blank, the value must be either ON, OFF, TOGGLE, or a dereferenced variable."
#define ERR_ON_OFF_TOGGLE_PERMIT "If not blank, the value must be either ON, OFF, TOGGLE, PERMIT, or a dereferenced variable."
#define ERR_TITLEMATCHMODE "TitleMatchMode must be either 1, 2, slow, fast, or a dereferenced variable."
#define ERR_TITLEMATCHMODE2 "The variable does not contain a valid TitleMatchMode (the value must be either 1, 2, slow, or fast)." ERR_ABORT
#define ERR_IFMSGBOX "This line specifies an invalid MsgBox result."
#define ERR_RUN_SHOW_MODE "The 3rd parameter must be either blank or one of these words: min, max, hide."
#define ERR_COMPARE_TIMES "Parameter #3 must be either blank, Seconds, Minutes, Hours, Days, or a dereferenced variable."
#define ERR_FILE_TIME "Parameter #3 must be either blank, M, C, A, or a dereferenced variable."
#define ERR_MOUSE_BUTTON "This line specifies an invalid mouse button."
#define ERR_MOUSE_COORD "The X & Y coordinates must be either both absent or both present."
#define ERR_PERCENT "The parameter must be a percentage between 0 and 100, or a dereferenced variable."
#define ERR_MOUSE_SPEED "The Mouse Speed must be a number between 0 and " MAX_MOUSE_SPEED_STR ", blank, or a dereferenced variable."
#define ERR_MEM_ASSIGN "Out of memory while assigning to this variable." ERR_ABORT
#define ERR_VAR_IS_RESERVED "This variable is reserved and cannot be assigned to."
#define ERR_DEFINE_CHAR "The character being defined must not be identical to another special or reserved character."
#define ERR_DEFINE_COMMENT "The comment flag must not be one of the hotkey definition symbols (e.g. ! ^ + $ ~ * < >)."

//----------------------------------------------------------------------------------

struct InputBoxType
{
	char *title;
	char *text;
	Var *output_var;
	char password_char;
	HWND hwnd;
};


typedef UINT LineNumberType;

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

typedef UCHAR ArgCountType;
#define MAX_ARGS 20

ResultType InputBox(Var *aOutputVar, char *aTitle = "", char *aText = "", bool aHideInput = false);
BOOL CALLBACK InputBoxProc(HWND hWndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
BOOL CALLBACK EnumChildFocusFind(HWND aWnd, LPARAM lParam);
BOOL CALLBACK EnumChildGetText(HWND aWnd, LPARAM lParam);
LRESULT CALLBACK MainWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam);
ResultType ShowMainWindow(char *aContents = NULL, bool aJumpToBottom = false);

bool Util_Shutdown(int nFlag);
BOOL Util_ShutdownHandler(HWND hwnd, DWORD lParam);
void Util_WinKill(HWND hWnd);


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
	ResultType PerformLoop(modLR_type aModifiersLR, WIN32_FIND_DATA *apCurrentFile
		, bool &aContinueMainLoop, Line *&aJumpToLine
		, AttributeType aAttr, FileLoopModeType aFileLoopMode, bool aRecurseSubfolders, char *aFilePattern
		, int aIterationLimit, bool aIsInfinite);
	ResultType Perform(modLR_type aModifiersLR, WIN32_FIND_DATA *aCurrentFile = NULL);
	ResultType PerformAssign();
	ResultType DriveSpaceFree(char *aPath);
	ResultType SoundPlay(char *aFilespec, bool aSleepUntilDone);
	ResultType FileSelectFile(char *aOptions, char *aWorkingDir);
	ResultType FileSelectFolder(char *aRootDir, bool aAllowCreateFolder, char *aGreeting);
	ResultType FileCreateDir(char *aDirSpec);
	ResultType FileReadLine(char *aFilespec, char *aLineNumber);
	ResultType FileAppend(char *aFilespec, char *aBuf);
	ResultType FileDelete(char *aFilePattern);
	ResultType FileMove(char *aSource, char *aDest, char *aFlag);
	ResultType FileCopy(char *aSource, char *aDest, char *aFlag);
	static bool Util_CopyFile(const char *szInputSource, const char *szInputDest, bool bOverwrite);
	static void Util_ExpandFilenameWildcard(const char *szSource, const char *szDest, char *szExpandedDest);
	static void Util_ExpandFilenameWildcardPart(const char *szSource, const char *szDest, char *szExpandedDest);
	static bool Util_DoesFileExist(const char *szFilename);

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
	ResultType RegRead(char *aRegKey, char *aRegSubkey, char *aValueName);
	ResultType RegWrite(char *aValueType, char *aRegKey, char *aRegSubkey, char *aValueName, char *aValue);
	ResultType RegDelete(char *aRegKey, char *aRegSubkey, char *aValueName);
	static bool RegRemoveSubkeys(HKEY hRegKey);
	static HKEY RegConvertMainKey(char *aBuf)
	{
		// Returns the HKEY, or NULL if aBuf contains an invalid key name.
		if (!stricmp(aBuf, "HKEY_LOCAL_MACHINE"))  return HKEY_LOCAL_MACHINE;
		if (!stricmp(aBuf, "HKEY_CLASSES_ROOT"))   return HKEY_CLASSES_ROOT;
		if (!stricmp(aBuf, "HKEY_CURRENT_CONFIG")) return HKEY_CURRENT_CONFIG;
		if (!stricmp(aBuf, "HKEY_CURRENT_USER"))   return HKEY_CURRENT_USER;
		if (!stricmp(aBuf, "HKEY_USERS"))          return HKEY_USERS;
		return NULL;
	}

	ResultType PerformShowWindow(ActionTypeType aActionType, char *aTitle = "", char *aText = ""
		, char *aExcludeTitle = "", char *aExcludeText = "");
	ResultType WinMove(char *aTitle, char *aText, char *aX, char *aY
		, char *aWidth = "", char *aHeight = "", char *aExcludeTitle = "", char *aExcludeText = "");
	ResultType WinMenuSelectItem(char *aTitle, char *aText, char *aMenu1, char *aMenu2
		, char *aMenu3, char *aMenu4, char *aMenu5, char *aMenu6, char *aMenu7
		, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlSend(char *aControl, char *aKeysToSend, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText, modLR_type aModifiersLR);
	ResultType ControlLeftClick(char *aControl, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlGetFocus(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlFocus(char *aControl, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlSetText(char *aControl, char *aNewText, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType ControlGetText(char *aControl, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType StatusBarGetText(char *aPart, char *aTitle, char *aText
		, char *aExcludeTitle, char *aExcludeText);
	ResultType StatusBarWait(char *aTextToWaitFor, char *aSeconds, char *aPart, char *aTitle, char *aText
		, char *aInterval, char *aExcludeTitle, char *aExcludeText);
	ResultType WinSetTitle(char *aTitle, char *aText, char *aNewTitle
		, char *aExcludeTitle = "", char *aExcludeText = "");
	ResultType WinGetTitle(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType WinGetText(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType WinGetPos(char *aTitle, char *aText, char *aExcludeTitle, char *aExcludeText);
	ResultType PixelSearch(int aLeft, int aTop, int aRight, int aBottom, int aColor, int aVariation);
	ResultType PixelGetColor(int aX, int aY);

	static ResultType SetToggleState(vk_type aVK, ToggleValueType &ForceLock, char *aToggleText);
	static ResultType MouseClickDrag(vk_type aVK // Which button.
		, int aX1, int aY1, int aX2, int aY2, int aSpeed);
	static ResultType MouseClick(vk_type aVK // Which button.
		, int aX = COORD_UNSPECIFIED, int aY = COORD_UNSPECIFIED // These values signal us not to move the mouse.
		, int aClickCount = 1, int aSpeed = DEFAULT_MOUSE_SPEED, char aEventType = '\0');
	static void MouseMove(int aX, int aY, int aSpeed = DEFAULT_MOUSE_SPEED);
	ResultType MouseGetPos();
	static void MouseEvent(DWORD aEventFlags, DWORD aX = 0, DWORD aY = 0)
	// A small inline to help us remember to use KEYIGNORE so that our own mouse
	// events won't be falsely detected as hotkeys by the hooks (if they are installed).
	{
		mouse_event(aEventFlags, aX, aY, 0, KEYIGNORE);
	}

public:
	#define LINE_LOG_SIZE 50
	static Line *sLog[LINE_LOG_SIZE];
	static int sLogNext;

	ActionTypeType mActionType; // What type of line this is.
	LineNumberType mFileLineNumber;  // The line number in the file from which the script was loaded, for debugging.
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
	#define LINE_ARG1 (line->mArgc > 0 ? line->sArgDeref[0] : "")
	#define LINE_ARG2 (line->mArgc > 1 ? line->sArgDeref[1] : "")
	#define LINE_ARG3 (line->mArgc > 2 ? line->sArgDeref[2] : "")
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

	ResultType ExecUntil(ExecUntilMode aMode, modLR_type aModifiersLR, Line **apJumpToLine = NULL
		, WIN32_FIND_DATA *aCurrentFile = NULL);

	Var *ResolveVarOfArg(int aArgIndex, bool aCreateIfNecessary = true);
	ResultType ExpandArgs();
	VarSizeType GetExpandedArgSize(bool aCalcDerefBufSize);
	char *ExpandArg(char *aBuf, int aArgIndex);

	bool FileIsFilteredOut(WIN32_FIND_DATA &aCurrentFile, FileLoopModeType aFileLoopMode, char *aFilePath);

	ResultType SetJumpTarget(bool aIsDereferenced);
	ResultType IsJumpValid(Line *aDestination);

	static ArgTypeType ArgIsVar(ActionTypeType aActionType, int aArgIndex);
	static int ConvertEscapeChar(char *aFilespec, char aOldChar, char aNewChar);
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

	bool ArgMustBeDereferenced(Var *aVar, int aArgIndexToExclude)
	{
		if (aVar->mType == VAR_CLIPBOARD)
			// Even if the clipboard is both an input and an output var, it still
			// doesn't need to be dereferenced into the temp buffer because the
			// clipboard has two buffers of its own.  The only exception is when
			// the clipboard has only files on it, in which case those files need
			// to be converted into plain text:
			return CLIPBOARD_CONTAINS_ONLY_FILES;
		if (aVar->mType != VAR_NORMAL || !aVar->Length())
			// Reserved vars must always be dereferenced due to their volatile nature.
			// Normal vars of length zero are dereferenced because they might exist
			// as system environment variables, whose contents are also potentially
			// volatile (i.e. they are sometimes changed by outside forces):
			return true;
		// Since the above didn't return, we know that this is a NORMAL input var of
		// non-zero length.  Such input vars only need to be dereferenced if they are
		// also used as an output var by the current script line:
		for (int iArg = 0; iArg < mArgc; ++iArg)
			if (iArg != aArgIndexToExclude && mArg[iArg].type == ARG_TYPE_OUTPUT_VAR
				&& ResolveVarOfArg(iArg, false) == aVar)
				return true;
		// Otherwise:
		return false;
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
		case ACT_SETWINDELAY:
		case ACT_SETCONTROLDELAY:
		case ACT_SETBATCHLINES:
		case ACT_RANDOM:
		case ACT_WINMOVE:
			return true;
		// Since mouse coords are relative to the foreground window, they can be negative:
		case ACT_MOUSECLICK:
			return (aArgNum == 2 || aArgNum == 3);
		case ACT_MOUSECLICKDRAG:
			return (aArgNum >= 2 && aArgNum <= 5);  // Allow dragging to/from negative coordinates.
		case ACT_MOUSEMOVE:
			return (aArgNum == 1 || aArgNum == 2);
		case ACT_PIXELGETCOLOR:
			return (aArgNum == 2 || aArgNum == 3);
		case ACT_PIXELSEARCH:
			return (aArgNum >= 3 || aArgNum <= 7); // i.e. Color values can be negative, but the last param cannot.
		}
		return false;  // Since above didn't return, negative is not allowed.
	}

	static TitleMatchModes ConvertTitleMatchMode(char *aBuf)
	{
		if (!aBuf || !*aBuf) return MATCHMODE_INVALID;
		if (*aBuf == '1' && !*(aBuf + 1)) return FIND_IN_LEADING_PART;
		if (*aBuf == '2' && !*(aBuf + 1)) return FIND_ANYWHERE;
		if (!stricmp(aBuf, "FAST")) return FIND_FAST;
		if (!stricmp(aBuf, "SLOW")) return FIND_SLOW;
		return MATCHMODE_INVALID;
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

	static int ConvertMouseButton(char *aBuf)
	// Returns the matching VK, or zero if none.
	{
		if (!aBuf || !*aBuf) return 0;
		if (!stricmp(aBuf, "LEFT")) return VK_LBUTTON;
		if (!stricmp(aBuf, "RIGHT")) return VK_RBUTTON;
		if (!stricmp(aBuf, "MIDDLE")) return VK_MBUTTON;
		return 0;
	}

	static ResultType ValidateMouseCoords(char *aX, char *aY)
	{
		if (!aX) aX = "";  if (!aY) aY = "";
		if (!*aX && !*aY) return OK;  // Both are absent, which is the signal to use the current position.
		if (*aX && *aY) return OK; // Both are present (that they are numeric is validated elsewhere).
		return FAIL; // Otherwise, one is blank but the other isn't, which is not allowed.
	}

	void Log()
	{
		// Probably doesn't need to be thread-safe or recursion-safe?
		sLog[sLogNext++] = this;
		if (sLogNext >= LINE_LOG_SIZE)
			sLogNext = 0;
	}
	static char *LogToText(char *aBuf, size_t aBufSize);
	char *VicinityToText(char *aBuf, size_t aBufSize, int aMaxLines = 15);
	char *ToText(char *aBuf, size_t aBufSize, bool aAppendNewline = false);

	static void ToggleSuspendState();
	ResultType ChangePauseState(ToggleValueType aChangeTo);

	Line *PreparseError(char *aErrorText);
	// Call this LineError to avoid confusion with Script's error-displaying functions:
	ResultType LineError(char *aErrorText, ResultType aErrorType = FAIL, char *aExtraInfo = "");

	Line(LineNumberType aFileLineNumber, ActionTypeType aActionType, ArgStruct aArg[], ArgCountType aArgc) // Constructor
		: mActionType(aActionType), mFileLineNumber(aFileLineNumber), mAttribute(ATTR_NONE)
		, mArgc(aArgc), mArg(aArg)
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



class Script
{
private:
	friend class Hotkey;
	Line *mFirstLine, *mLastLine;  // The first and last lines in the linked list.
	Label *mFirstLabel, *mLastLabel;  // The first and last labels in the linked list.
	Var *mFirstVar, *mLastVar;  // The first and last variables in the linked list.
	WinGroup *mFirstGroup, *mLastGroup;  // The first and last variables in the linked list.
	UINT mLineCount, mLabelCount, mVarCount, mGroupCount;
	LineNumberType mFileLineCount;  // How many physical lines are in the file.
	NOTIFYICONDATA mNIC; // For ease of adding and deleting our tray icon.

#ifdef AUTOHOTKEYSC
	int CloseAndReturn(HS_EXEArc_Read *fp, UCHAR *aBuf, int return_value);
	size_t GetLine(char *aBuf, int aMaxCharsToRead, UCHAR *&aMemFile);
#else
	int CloseAndReturn(FILE *fp, UCHAR *aBuf, int return_value);
	size_t GetLine(char *aBuf, int aMaxCharsToRead, FILE *fp);
#endif
	ResultType IsPreprocessorDirective(char *aBuf);

	ResultType ParseAndAddLine(char *aLineText, char *aActionName = NULL, char *aEndMarker = NULL
		, char *aLiteralMap = NULL, size_t aLiteralMapLength = 0
		, ActionTypeType aActionType = ACT_INVALID, ActionTypeType aOldActionType = OLD_INVALID);
	char *ParseActionType(char *aBufTarget, char *aBufSource, bool aDisplayErrors);
	static ActionTypeType ConvertActionType(char *aActionTypeString);
	static ActionTypeType ConvertOldActionType(char *aActionTypeString);
	ResultType AddLabel(char *aLabelName);
	ResultType AddLine(ActionTypeType aActionType, char *aArg[] = NULL, ArgCountType aArgc = 0
		, char *aArgMap[] = NULL);
	ResultType AddVar(char *aVarName, size_t aVarNameLength = 0);

	// These aren't in the Line class because I think they're easier to implement
	// if aStartingLine is allowed to be NULL (for recursive calls).  If they
	// were member functions of class Line, a check for NULL would have to
	// be done before dereferencing any line's mNextLine, for example:
	Line *PreparseBlocks(Line *aStartingLine, int aFindBlockEnd = 0, Line *aParentLine = NULL);
	Line *PreparseIfElse(Line *aStartingLine, ExecUntilMode aMode = NORMAL_MODE
		, AttributeType aLoopType = ATTR_NONE);
public:
	Line *mCurrLine;  // Seems better to make this public than make Line our friend.
	Label *mThisHotkeyLabel, *mPriorHotkeyLabel;
	WIN32_FIND_DATA *mLoopFile;  // The file of the current file-loop, if applicable.
	DWORD mThisHotkeyStartTime, mPriorHotkeyStartTime;  // Tickcount timestamp of when its subroutine began.
	char *mFileSpec; // Will hold the full filespec, for convenience.
	char *mFileDir;  // Will hold the directory containing the script file.
	char *mFileName; // Will hold the script's naked file name.
	char *mOurEXE; // Will hold this app's module name (e.g. C:\Program Files\AutoHotkey\AutoHotkey.exe).
	char *mMainWindowTitle; // Will hold our main window's title, for consistency & convenience.
	bool mIsReadyToExecute;
	bool mIsRestart; // The app is restarting rather than starting from scratch.
	bool mIsAutoIt2; // Whether this script is considered to be an AutoIt2 script.
	LineNumberType mLinesExecutedThisCycle; // Not tracking this separately for every recursed subroutine.
	DWORD mLastSleepTime; // Track MsgSleep() from any and all sources to pump messages more consistently.

	ResultType Init(char *aScriptFilename, bool aIsRestart);
	ResultType CreateWindows(HINSTANCE hInstance);
	void UpdateTrayIcon();
	ResultType Edit();
	ResultType Reload();
	void ExitApp(char *aBuf = NULL, int ExitCode = 0);
	int LoadFromFile();
	Var *FindOrAddVar(char *aVarName, size_t aVarNameLength = 0);
	Var *FindVar(char *aVarName, size_t aVarNameLength = 0);
	ResultType ExecuteFromLine1()
	{
		if (!mIsReadyToExecute)
			return FAIL;
		if (mFirstLine != NULL)
			return mFirstLine->ExecUntil(UNTIL_RETURN, 0);
		return OK;
	}
	WinGroup *FindOrAddGroup(char *aGroupName);
	ResultType AddGroup(char *aGroupName);
	Label *FindLabel(char *aLabelName);
	ResultType ActionExec(char *aAction, char *aParams = NULL, char *aWorkingDir = NULL
		, bool aDisplayErrors = true, char *aRunShowMode = NULL, HANDLE *aProcess = NULL);
	char *ListVars(char *aBuf, size_t aBufSize);

	int GetFilename(char *aBuf = NULL)
	{
		VarSizeType length = (VarSizeType)strlen(mFileName);
		if (!aBuf) return length;
		strcpy(aBuf, mFileName);
		aBuf += length;
		return length;
	}
	int GetFileDir(char *aBuf = NULL)
	{
		VarSizeType length = (VarSizeType)strlen(mFileDir);
		if (!aBuf) return length;
		strcpy(aBuf, mFileDir);
		aBuf += length;
		return length;
	}
	int GetFilespec(char *aBuf = NULL)
	{
		VarSizeType length = (VarSizeType)(strlen(mFileDir) + strlen(mFileName) + 1);
		if (!aBuf) return length;
		sprintf(aBuf, "%s\\%s", mFileDir, mFileName);
		aBuf += length;
		return length;
	}

	int GetLoopFileName(char *aBuf = NULL)
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
		VarSizeType length = (VarSizeType)strlen(str);
		if (!aBuf) return length;
		strcpy(aBuf, str);
		aBuf += length;
		return length;
	}
	int GetLoopFileShortName(char *aBuf = NULL)
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
		VarSizeType length = (VarSizeType)strlen(str);
		if (!aBuf) return length;
		strcpy(aBuf, str);
		aBuf += length;
		return length;
	}
	int GetLoopFileDir(char *aBuf = NULL)
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
		aBuf += length;
		if (last_backslash)
			*last_backslash = '\\';  // Restore the orginal value.
		return length;
	}
	int GetLoopFileFullPath(char *aBuf = NULL)
	{
		char *str = "";  // Set default.
		if (mLoopFile)
			// The loop handler already prepended the script's directory in here for us:
			str = mLoopFile->cFileName;
		VarSizeType length = (VarSizeType)strlen(str);
		if (!aBuf) return length;
		strcpy(aBuf, str);
		aBuf += length;
		return length;
	}
	int GetLoopFileTimeModified(char *aBuf = NULL)
	{
		char str[64] = "";  // Set default.
		if (mLoopFile)
			FileTimeToYYYYMMDD(str, &mLoopFile->ftLastWriteTime, true);
		VarSizeType length = (VarSizeType)strlen(str);
		if (!aBuf) return length;
		strcpy(aBuf, str);
		aBuf += length;
		return length;
	}
	int GetLoopFileTimeCreated(char *aBuf = NULL)
	{
		char str[64] = "";  // Set default.
		if (mLoopFile)
			FileTimeToYYYYMMDD(str, &mLoopFile->ftCreationTime, true);
		VarSizeType length = (VarSizeType)strlen(str);
		if (!aBuf) return length;
		strcpy(aBuf, str);
		aBuf += length;
		return length;
	}
	int GetLoopFileTimeAccessed(char *aBuf = NULL)
	{
		char str[64] = "";  // Set default.
		if (mLoopFile)
			FileTimeToYYYYMMDD(str, &mLoopFile->ftLastAccessTime, true);
		VarSizeType length = (VarSizeType)strlen(str);
		if (!aBuf) return length;
		strcpy(aBuf, str);
		aBuf += length;
		return length;
	}
	int GetLoopFileAttrib(char *aBuf = NULL)
	{
		char str[64] = "";  // Set default.
		if (mLoopFile)
			FileAttribToStr(str, mLoopFile->dwFileAttributes);
		VarSizeType length = (VarSizeType)strlen(str);
		if (!aBuf) return length;
		strcpy(aBuf, str);
		aBuf += length;
		return length;
	}
	int GetLoopFileSize(char *aBuf = NULL)
	{
		char str[64] = "";  // Set default.
		if (mLoopFile)
			// It's a documented limitation that the size will show as negative if
			// greater than 2 gig, and will be wrong if greater than 4 gig.  For files
			// that large, scripts should use the KB version of this function instead.
			// If a file is over 4gig, set the value to be the maximum size (-1 when
			// expressed as a signed integer, since script variables are based entirely
			// on 32-bit signed integers due to the use of atoi(), etc.):
			sprintf(str, "%d%", mLoopFile->nFileSizeHigh ? -1 : (int)mLoopFile->nFileSizeLow);
		VarSizeType length = (VarSizeType)strlen(str);
		if (!aBuf) return length;
		strcpy(aBuf, str);
		aBuf += length;
		return length;
	}
	int GetLoopFileSizeKB(char *aBuf = NULL)
	{
		char str[128] = "";  // Set default.
		if (mLoopFile)
		{
			ULARGE_INTEGER ul;
			ul.LowPart = mLoopFile->nFileSizeLow;
			ul.HighPart = mLoopFile->nFileSizeHigh;
			// Always use signed (%d) vs. unsigned (%u).  See comment in GetLoopFileSize():
			sprintf(str, "%d%", (int)((unsigned __int64)ul.QuadPart/1024));
		}
		VarSizeType length = (VarSizeType)strlen(str);
		if (!aBuf) return length;
		strcpy(aBuf, str);
		aBuf += length;
		return length;
	}

	int GetThisHotkey(char *aBuf = NULL)
	{
		char *str = "";  // Set default.
		if (mThisHotkeyLabel)
			str = mThisHotkeyLabel->mName;
		VarSizeType length = (VarSizeType)strlen(str);
		if (!aBuf) return length;
		strcpy(aBuf, str);
		aBuf += length;
		return length;
	}
	int GetPriorHotkey(char *aBuf = NULL)
	{
		char *str = "";  // Set default.
		if (mPriorHotkeyLabel)
			str = mPriorHotkeyLabel->mName;
		VarSizeType length = (VarSizeType)strlen(str);
		if (!aBuf) return length;
		strcpy(aBuf, str);
		aBuf += length;
		return length;
	}
	int GetTimeSinceThisHotkey(char *aBuf = NULL)
	{
		char str[128];
		// It must be the type of hotkey that has a label because we want the TimeSinceThisHotkey
		// value to be "in sync" with the value of ThisHotkey itself (i.e. use the same method
		// to determine which hotkey is the "this" hotkey):
		if (mThisHotkeyLabel)
			// Even if GetTickCount()'s TickCount has wrapped around to zero and the timestamp hasn't,
			// DWORD math still gives the right answer as long as the number of days between
			// isn't greater than about 49.  See MyGetTickCount() for explanation of %d vs. %u:
			snprintf(str, sizeof(str), "%d", (DWORD)(GetTickCount() - mThisHotkeyStartTime));
		else
			strcpy(str, "-1");
		VarSizeType length = (VarSizeType)strlen(str);
		if (!aBuf) return length;
		strcpy(aBuf, str);
		aBuf += length;
		return length;
	}
	int GetTimeSincePriorHotkey(char *aBuf = NULL)
	{
		char str[128];
		if (mPriorHotkeyLabel)
			// See MyGetTickCount() for explanation of %d vs. %u:
			snprintf(str, sizeof(str), "%d", (DWORD)(GetTickCount() - mPriorHotkeyStartTime));
		else
			strcpy(str, "-1");
		VarSizeType length = (VarSizeType)strlen(str);
		if (!aBuf) return length;
		strcpy(aBuf, str);
		aBuf += length;
		return length;
	}
	int MyGetTickCount(char *aBuf = NULL)
	{
		// Known limitation:
		// Although TickCount is an unsigned value, I'm not sure that our EnvSub command
		// will properly be able to compare two tick-counts if either value is larger than
		// INT_MAX.  So if the system has been up for more than about 25 days, there might be
		// problems if the user tries compare two tick-counts in the script using EnvSub.
		// UPDATE: It seems better to store all unsigned values as signed within script
		// variables.  Otherwise, when the var's value is next accessed and converted using
		// atoi(), the outcome won't be as useful.  In other words, since the negative value
		// will be properly converted by atoi(), comparing two negative tickcounts works
		// correctly (confirmed).  Even if one of them is negative and the other positive,
		// it will probably work correctly due to the nature of implicit unsigned math.
		// Thus, we use %d vs. %u in the snprintf() call below.
		char str[128];
		snprintf(str, sizeof(str), "%d", GetTickCount());
		VarSizeType length = (VarSizeType)strlen(str);
		if (!aBuf) return length;
		strcpy(aBuf, str);
		aBuf += length;
		return length;
	}
	int GetSpace(char *aBuf = NULL)
	{
		if (!aBuf) return 1;  // i.e. the length of a single space char.
		*(aBuf++) = ' ';
		*aBuf = '\0';
		return 1;
	}

	// Call this SciptError to avoid confusion with Line's error-displaying functions:
	ResultType ScriptError(char *aErrorText, char *aExtraInfo = "");
	void ShowInEditor();

	Script();
	// Note that the anchors to any linked lists will be lost when this
	// object goes away, so for now, be sure the destructor is only called
	// when the program is about to be exited, which will thereby reclaim
	// the memory used by the abandoned linked lists (otherwise, a memory
	// leak will result).
};

#endif
