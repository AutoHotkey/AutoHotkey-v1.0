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

#ifndef var_h
#define var_h

#include "defines.h"
#include "SimpleHeap.h"
#include "clipboard.h"
#include "util.h" // for strlcpy() & snprintf()
EXTERN_CLIPBOARD;

#define MAX_ALLOC_SIMPLE 64  // Do not decrease this much since it is used for the sizing of some built-in variables.
#define SMALL_STRING_LENGTH (MAX_ALLOC_SIMPLE - 1)  // The largest string that can fit in the above.
#define DEREF_BUF_EXPAND_INCREMENT (32 * 1024)
#define ERRORLEVEL_NONE "0"
#define ERRORLEVEL_ERROR "1"
#define ERRORLEVEL_ERROR2 "2"

enum AllocMethod {ALLOC_NONE, ALLOC_SIMPLE, ALLOC_MALLOC};
enum VarTypes
{
  VAR_NORMAL // It's probably best of this one is first (zero).
, VAR_CLIPBOARD
, VAR_YYYY, VAR_MM, VAR_MMMM, VAR_MMM, VAR_DD, VAR_YDAY, VAR_YWEEK, VAR_WDAY, VAR_DDDD, VAR_DDD
, VAR_HOUR, VAR_MIN, VAR_SEC, VAR_TICKCOUNT, VAR_NOW, VAR_NOWUTC
, VAR_WORKINGDIR, VAR_BATCHLINES
, VAR_TITLEMATCHMODE, VAR_TITLEMATCHMODESPEED, VAR_DETECTHIDDENWINDOWS, VAR_DETECTHIDDENTEXT
, VAR_AUTOTRIM, VAR_STRINGCASESENSE, VAR_FORMATINTEGER, VAR_FORMATFLOAT
, VAR_KEYDELAY, VAR_WINDELAY, VAR_CONTROLDELAY, VAR_MOUSEDELAY, VAR_DEFAULTMOUSESPEED
, VAR_ICONHIDDEN, VAR_ICONTIP, VAR_ICONFILE, VAR_ICONNUMBER, VAR_EXITREASON
, VAR_OSTYPE, VAR_OSVERSION, VAR_LANGUAGE, VAR_COMPUTERNAME, VAR_USERNAME
, VAR_WINDIR, VAR_PROGRAMFILES, VAR_DESKTOP, VAR_DESKTOPCOMMON, VAR_STARTMENU, VAR_STARTMENUCOMMON
, VAR_PROGRAMS, VAR_PROGRAMSCOMMON, VAR_STARTUP, VAR_STARTUPCOMMON, VAR_MYDOCUMENTS
, VAR_ISADMIN, VAR_CURSOR, VAR_CARETX, VAR_CARETY
, VAR_SCREENWIDTH, VAR_SCREENHEIGHT
, VAR_IPADDRESS1, VAR_IPADDRESS2, VAR_IPADDRESS3, VAR_IPADDRESS4
, VAR_SCRIPTNAME, VAR_SCRIPTDIR, VAR_SCRIPTFULLPATH
, VAR_LOOPFILENAME, VAR_LOOPFILESHORTNAME, VAR_LOOPFILEDIR, VAR_LOOPFILEFULLPATH
, VAR_LOOPFILETIMEMODIFIED, VAR_LOOPFILETIMECREATED, VAR_LOOPFILETIMEACCESSED
, VAR_LOOPFILEATTRIB, VAR_LOOPFILESIZE, VAR_LOOPFILESIZEKB, VAR_LOOPFILESIZEMB
, VAR_LOOPREGTYPE, VAR_LOOPREGKEY, VAR_LOOPREGSUBKEY, VAR_LOOPREGNAME, VAR_LOOPREGTIMEMODIFIED
, VAR_LOOPREADLINE, VAR_LOOPFIELD, VAR_INDEX
, VAR_THISMENUITEM, VAR_THISMENUITEMPOS, VAR_THISMENU, VAR_THISHOTKEY, VAR_PRIORHOTKEY
, VAR_TIMESINCETHISHOTKEY, VAR_TIMESINCEPRIORHOTKEY
, VAR_ENDCHAR, VAR_GUI, VAR_GUICONTROL, VAR_GUICONTROLEVENT, VAR_GUIWIDTH, VAR_GUIHEIGHT
, VAR_TIMEIDLE, VAR_TIMEIDLEPHYSICAL
, VAR_SPACE, VAR_TAB, VAR_AHKVERSION
};
#define VAR_IS_RESERVED(var) (var->mType != VAR_NORMAL && var->mType != VAR_CLIPBOARD)

typedef UCHAR VarTypeType;     // UCHAR vs. VarTypes to save memory.
typedef UCHAR AllocMethodType; // UCHAR vs. AllocMethod to save memory.
typedef DWORD VarSizeType;  // Up to 4 gig if sizeof(UINT) is 4.  See next line.
#define VARSIZE_MAX MAXDWORD
#define VARSIZE_ERROR VARSIZE_MAX
#define MAX_FORMATTED_NUMBER_LENGTH 255

class Var
{
private:
	char *mContents;
	VarSizeType mLength;   // How much is actually stored in it currently, excluding the zero terminator.
	VarSizeType mCapacity; // In bytes.  Includes the space for the zero terminator.
	AllocMethodType mHowAllocated; // Keep adjacent/contiguous with the below.
public:
	VarTypeType mType;  // Keep adjacent/contiguous with the above.
	// Testing shows that due to data alignment, keeping mType adjacent to the other less-than-4-size member
	// above it reduces size of each object by 4 bytes.
	char *mName;    // The name of the var.
	Var *mNextVar;  // Next item in linked list.
	static char sEmptyString[1]; // A special, non-constant memory area for empty variables.

	// Convert to unsigned 64-bit to support for 64-bit pointers.  Since most script operations --
	// such as addition and comparison -- read strings in as signed 64-bit, it is documented that
	// math and other numerical operations should never be performed on these while they exist
	// as strings in script variables:
	//#define ASSIGN_HWND_TO_VAR(var, hwnd) var->Assign((unsigned __int64)hwnd)
	// UPDATE: Always assign as hex for better compatibility with Spy++ and other apps that
	// report window handles:
	ResultType AssignHWND(HWND aWnd)
	{
		char buf[64];
		*buf = '0';
		*(buf + 1) = 'x';
		_ui64toa((unsigned __int64)aWnd, buf + 2, 16);
		return Assign(buf);
	}

	ResultType Assign(int aValueToAssign);
	ResultType Assign(DWORD aValueToAssign);
	ResultType Assign(__int64 aValueToAssign);
	//ResultType Assign(unsigned __int64 aValueToAssign);
	ResultType Assign(double aValueToAssign);
	ResultType Assign(char *aBuf = NULL, VarSizeType aLength = VARSIZE_MAX, bool aTrimIt = false);
	VarSizeType Get(char *aBuf = NULL);
	static ResultType ValidateName(char *aName, bool aIsRuntime = false);
	VarSizeType Capacity() {return mCapacity;}
	char *ToText(char *aBuf, size_t aBufSize, bool aAppendNewline)
	// Translates this var into its text equivalent, putting the result into aBuf andp
	// returning the position in aBuf of its new string terminator.
	{
		if (!aBuf) return NULL;
		char *aBuf_orig = aBuf;
		snprintf(aBuf, BUF_SPACE_REMAINING, "%s[%u of %u]: %-1.60s%s", mName
			, mLength, mCapacity ? (mCapacity - 1) : 0  // Use -1 since it makes more sense to exclude the terminator.
			, mContents, mLength > 60 ? "..." : "");
		aBuf += strlen(aBuf);
		if (aAppendNewline && BUF_SPACE_REMAINING >= 2)
		{
			*aBuf++ = '\r';
			*aBuf++ = '\n';
			*aBuf = '\0';
		}
		return aBuf;
	}

	char *Contents()
	{
		if (mType == VAR_NORMAL)
			return mContents;
		if (mType == VAR_CLIPBOARD)
			// The returned value will be a writable mem area if clipboard is open for write.
			// Otherwise, the clipboard will be opened physically, if it isn't already, and
			// a pointer to its contents returned to the caller:
			return g_clip.Contents();
		// For reserved vars.  Probably better than returning NULL:
		return "Unsupported script variable type.";
	}

	VarSizeType &Length() // Returns a reference so that caller can use this function as an lvalue.
	{
		if (mType == VAR_NORMAL)
			return mLength;
		static VarSizeType length; // Must be static so that caller can use its contents.
		if (mType == VAR_CLIPBOARD)
			// Since the length of the clipboard isn't normally tracked, we just return a
			// temporary storage area for the caller to use.  Note: This approach is probably
			// not thread-safe, but currently there's only one thread so it's not an issue.
			return length;
		// For reserved vars do the same thing as above, but this function should never
		// be called for them:
		return length;  // Should never be reached?
	}

	ResultType Close()
	{
		if (mType == VAR_CLIPBOARD && g_clip.IsReadyForWrite())
			return g_clip.Commit(); // Writes the new clipboard contents to the clipboard and closes it.
		return OK; // In all other cases.
	}

	Var(char *aVarName, VarTypeType aType = VAR_NORMAL)
		// The caller must ensure that aVarName is non-null.
		: mName(aVarName) // Caller gave us a pointer to dynamic memory for this (or static in the case of ResolveVarOfArg()).
		, mLength(0)
		, mCapacity(0)
		, mHowAllocated(ALLOC_NONE)
		, mType(aType)
		// This initial empty-string value may be relied upon (i.e. don't make it NULL).
		// In addition, it's safer to make it modifiable rather than a constant such
		// as "" so that any caller that accesses the memory via Contents() can write
		// to it (e.g. it can do *buf = '\0').  In some sense this is a little scary
		// because if anything misbehaves and sets this value to be something other
		// than '\0', all empty variables will suddenly take on the wrong value and
		// there will suddenly be a lot of buffer overflows probably, since those
		// variables will no longer be terminated.  However, you could argue that this
		// is a good thing because any bug that ever does that is more likely to cause
		// a crash right away rather than go silently undetected for a long time:
		, mContents(sEmptyString)
		, mNextVar(NULL)
	{}
	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}
};


#endif
