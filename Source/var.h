/*
AutoHotkey

Copyright 2003-2005 Chris Mallett

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
// The following values have a special nature unlike the other types. They need to be kept first for so that
// VAR_FIRST_NON_BYREF can be used to determine whether a variable is BYREF or not.
// VAR_BYREF is used when an ALIAS variable doesn't yet have a target, which avoids the need to check
// whether mAliasFor is NULL in many places.  In other words, a var of type VAR_ALIAS always has a valid mAliasFor.
  VAR_INVALID, VAR_BYREF, VAR_ALIAS, VAR_FIRST_NON_BYREF
, VAR_NORMAL = VAR_FIRST_NON_BYREF  // Most variables are this type.
, VAR_CLIPBOARD
, VAR_LAST_UNRESERVED = VAR_CLIPBOARD  // Keep this in sync with any changes to the set of unreserved variables.
#define VAR_IS_RESERVED(var) (var->Type() > VAR_LAST_UNRESERVED)
, VAR_CLIPBOARDALL // Must be reserved because it's not designed to be writable.
, VAR_TRUE, VAR_FALSE
, VAR_YYYY, VAR_MM, VAR_MMMM, VAR_MMM, VAR_DD, VAR_YDAY, VAR_YWEEK, VAR_WDAY, VAR_DDDD, VAR_DDD
, VAR_HOUR, VAR_MIN, VAR_SEC, VAR_MSEC, VAR_TICKCOUNT, VAR_NOW, VAR_NOWUTC
, VAR_WORKINGDIR, VAR_BATCHLINES
, VAR_TITLEMATCHMODE, VAR_TITLEMATCHMODESPEED, VAR_DETECTHIDDENWINDOWS, VAR_DETECTHIDDENTEXT
, VAR_AUTOTRIM, VAR_STRINGCASESENSE, VAR_FORMATINTEGER, VAR_FORMATFLOAT
, VAR_KEYDELAY, VAR_WINDELAY, VAR_CONTROLDELAY, VAR_MOUSEDELAY, VAR_DEFAULTMOUSESPEED, VAR_ISSUSPENDED
, VAR_ICONHIDDEN, VAR_ICONTIP, VAR_ICONFILE, VAR_ICONNUMBER, VAR_EXITREASON
, VAR_OSTYPE, VAR_OSVERSION, VAR_LANGUAGE, VAR_COMPUTERNAME, VAR_USERNAME
, VAR_WINDIR, VAR_PROGRAMFILES, VAR_DESKTOP, VAR_DESKTOPCOMMON, VAR_STARTMENU, VAR_STARTMENUCOMMON
, VAR_PROGRAMS, VAR_PROGRAMSCOMMON, VAR_STARTUP, VAR_STARTUPCOMMON, VAR_MYDOCUMENTS
, VAR_ISADMIN, VAR_CURSOR, VAR_CARETX, VAR_CARETY
, VAR_SCREENWIDTH, VAR_SCREENHEIGHT
, VAR_IPADDRESS1, VAR_IPADDRESS2, VAR_IPADDRESS3, VAR_IPADDRESS4
, VAR_SCRIPTNAME, VAR_SCRIPTDIR, VAR_SCRIPTFULLPATH, VAR_LINENUMBER, VAR_LINEFILE, VAR_ISCOMPILED
, VAR_LOOPFILENAME, VAR_LOOPFILESHORTNAME, VAR_LOOPFILEDIR
, VAR_LOOPFILEFULLPATH, VAR_LOOPFILELONGPATH, VAR_LOOPFILESHORTPATH
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

typedef UCHAR VarTypeType;     // UCHAR vs. VarTypes to save memory.
typedef UCHAR AllocMethodType; // UCHAR vs. AllocMethod to save memory.
typedef UCHAR VarAttribType;   // Same.
typedef DWORD VarSizeType;     // Up to 4 gig if sizeof(UINT) is 4.  See next line.
#define VARSIZE_MAX MAXDWORD
#define VARSIZE_ERROR VARSIZE_MAX
#define MAX_FORMATTED_NUMBER_LENGTH 255 // Large enough to allow custom zero or space-padding via %10.2f, etc.  But not too large because some things might rely on this being fairly small.


class Var; // Forward declaration.
struct VarBkp // This should be kept in sync with any changes to the Var class.  See Var for comments.
{
	Var *mVar; // Used to save the target var to which these backed up contents will later be restored.
	char *mContents;
	union
	{
		VarSizeType mLength;
		Var *mAliasFor;
	};
	VarSizeType mCapacity;
	AllocMethodType mHowAllocated;
	VarAttribType mAttrib;
	// Not needed in the backup:
	//bool mIsLocal;
	//VarTypeType mType;
	//char *mName;
};

class Var
{
private:
	// Keep VarBkp (above) in sync with any changes made to the members here.
	char *mContents;
	union
	{
		VarSizeType mLength;  // How much is actually stored in it currently, excluding the zero terminator.
		Var *mAliasFor;       // The variable for which this variable is an alias.
	};
	VarSizeType mCapacity; // In bytes.  Includes the space for the zero terminator.
	AllocMethodType mHowAllocated; // Keep adjacent/contiguous with the below to save memory.
	#define VAR_ATTRIB_BINARY_CLIP  0x01
	#define VAR_ATTRIB_PARAM        0x02 // Currently unused.
	#define VAR_ATTRIB_STATIC       0x04 // Next in series would be 0x08, 0x10, etc.
	VarAttribType mAttrib;  // Bitwise combination of the above flags.
	bool mIsLocal;
	VarTypeType mType;  // Keep adjacent/contiguous with the above.

public:
	// Testing shows that due to data alignment, keeping mType adjacent to the other less-than-4-size member
	// above it reduces size of each object by 4 bytes.
	char *mName;    // The name of the var.
	static char sEmptyString[1]; // A special, non-constant memory area for empty variables.


	void Backup(VarBkp &aVarBkp)
	// This method is used rather than struct copy (=) because it's of expected higher performance than
	// using the Var::constructor to make a copy of each var.  Also note that something like memcpy()
	// can't be used on Var objects since they're not POD (e.g. they have a contructor and they have
	// private members).
	{
		aVarBkp.mVar = this; // Allows the Restore() to always know its target without searching.
		aVarBkp.mContents = mContents;
		aVarBkp.mLength = mLength;
		aVarBkp.mCapacity = mCapacity;
		aVarBkp.mHowAllocated = mHowAllocated; // This might be ALLOC_SIMPLE or ALLOC_NONE if backed up variable was at the lowest layer of the call stack.
		aVarBkp.mAttrib = mAttrib;
		// Once the backup is made, Free() is not called because the whole point of the backup is to
		// preserve the original memory/contents of each variable.  Instead, clear the variable
		// completely and set it up to become ALLOC_MALLOC in case anything actually winds up using
		// the variable prior to the restoration of the backup.  In other words, ALLOC_SIMPLE and NONE
		// retained (if present) because that would cause a memory leak when multiple layers are all
		// allowed to use ALLOC_SIMPLE yet none are ever able to free it (the bottommost layer is
		// allowed to use ALLOC_SIMPLE because that's a fixed/constant amount of memory gets freed
		// when the program exits).
		// Reset this variable to create a "new layer" for it, keeping its backup intact but allowing
		// this variable (or formal parameter) to be given a new value in the future:
		if (mAttrib & VAR_ATTRIB_STATIC) // By definition, static variables retain their contents between calls.
			return;
		mContents = sEmptyString;
		mLength = 0;
		mCapacity = 0;
		mHowAllocated = ALLOC_MALLOC;
		mAttrib &= ~VAR_ATTRIB_BINARY_CLIP;  // But the VAR_ATTRIB_PARAM/STATIC flags are unaltered.
	}

	void Restore(VarBkp &aVarBkp)
	// Caller must ensure that Free() has been called for this variable prior calling Restore(), since otherwise
	// there would be a memory leak.
	{
		if (mAttrib & VAR_ATTRIB_STATIC)
			return;
		mContents = aVarBkp.mContents;
		mLength = aVarBkp.mLength;
		mCapacity = aVarBkp.mCapacity;
		mHowAllocated = aVarBkp.mHowAllocated; // This might be ALLOC_SIMPLE or ALLOC_NONE if backed up variable was at the lowest layer of the call stack.
		mAttrib = aVarBkp.mAttrib;
	}

	ResultType AssignHWND(HWND aWnd)
	{
		// Convert to unsigned 64-bit to support for 64-bit pointers.  Since most script operations --
		// such as addition and comparison -- read strings in as signed 64-bit, it is documented that
		// math and other numerical operations should never be performed on these while they exist
		// as strings in script variables:
		//#define ASSIGN_HWND_TO_VAR(var, hwnd) var->Assign((unsigned __int64)hwnd)
		// UPDATE: Always assign as hex for better compatibility with Spy++ and other apps that
		// report window handles:
		char buf[64];
		*buf = '0';
		buf[1] = 'x';
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

	#define VAR_ALWAYS_FREE         0
	#define VAR_NEVER_FREE          1
	#define VAR_FREE_IF_LARGE       2
	#define VAR_FREE_EXCLUDE_STATIC 3
	void Free(int aWhenToFree = VAR_ALWAYS_FREE, bool aExcludeAliases = false);

	static ResultType ValidateName(char *aName, bool aIsRuntime = false, bool aDisplayError = true);

	char *ToText(char *aBuf, int aBufSize, bool aAppendNewline)
	// aBufSize is an int so that any negative values passed in from caller are not lost.
	// Caller has ensured that aBuf isn't NULL.
	// Translates this var into its text equivalent, putting the result into aBuf andp
	// returning the position in aBuf of its new string terminator.
	{
		if (mType == VAR_ALIAS) // For simplicity and reduced code size, just make a recursive call to self.
			return mAliasFor->ToText(aBuf, aBufSize, aAppendNewline);
		char *aBuf_orig = aBuf;
		aBuf += snprintf(aBuf, BUF_SPACE_REMAINING, "%s[%u of %u]: %-1.60s%s", mName
			, mLength, mCapacity ? (mCapacity - 1) : 0  // Use -1 since it makes more sense to exclude the terminator.
			, mContents, mLength > 60 ? "..." : "");
		if (aAppendNewline && BUF_SPACE_REMAINING >= 2)
		{
			*aBuf++ = '\r';
			*aBuf++ = '\n';
			*aBuf = '\0';
		}
		return aBuf;
	}

	// These rely on the fact that currently, aliases can't point to other aliases:
	VarTypeType Type()
	{
		return (mType == VAR_ALIAS) ? mAliasFor->mType : mType;
	}
	bool IsByRef()
	{
		return mType < VAR_FIRST_NON_BYREF;
	}
	bool IsLocal()
	{
		// Callers want to know whether this variable is local, even if it's a local alias for a global:
		return mIsLocal;
		//return (mType == VAR_ALIAS) ? mAliasFor->mIsLocal : mIsLocal;
	}
	bool IsBinaryClip()
	{
		return (mType == VAR_ALIAS ? mAliasFor->mAttrib : mAttrib) & VAR_ATTRIB_BINARY_CLIP;
	}
	void OverwriteAttrib(VarAttribType aAttrib)
	{
		if (mType == VAR_ALIAS)
			mAliasFor->mAttrib = aAttrib;
		else
			mAttrib = aAttrib;
	}
	VarSizeType Capacity() // Capacity includes the zero terminator.
	{
		return (mType == VAR_ALIAS) ? mAliasFor->mCapacity : mCapacity;
	}

	VarSizeType &Length()
	// This should not be called to discover a non-NORMAL var's length because the length
	// of most such variables aren't knowable without calling Get() on them.
	// Returns a reference so that caller can use this function as an lvalue.
	{
		if (mType == VAR_ALIAS) // For simplicity and reduced code size, just make a recursive call to self.
			return mAliasFor->Length();
		if (mType == VAR_NORMAL)
			return mLength;
		// Since the length of the clipboard isn't normally tracked, we just return a
		// temporary storage area for the caller to use.  Note: This approach is probably
		// not thread-safe, but currently there's only one thread so it's not an issue.
		// For reserved vars do the same thing as above, but this function should never
		// be called for them:
		static VarSizeType length; // Must be static so that caller can use its contents.
		return length;
	}

	char *Contents()
	{
		if (mType == VAR_ALIAS) // For simplicity and reduced code size, just make a recursive call to self.
			return mAliasFor->Contents();
		if (mType == VAR_NORMAL)
			return mContents;
		if (mType == VAR_CLIPBOARD)
			// The returned value will be a writable mem area if clipboard is open for write.
			// Otherwise, the clipboard will be opened physically, if it isn't already, and
			// a pointer to its contents returned to the caller:
			return g_clip.Contents();
		return ""; // For reserved vars (but this method should probably never be called for them).
	}

	Var *ResolveAlias()
	{
		return (mType == VAR_ALIAS) ? mAliasFor : this; // Return target if it's an alias, or itself if not.
	}

	void UpdateAlias(Var *aTargetVar)
	// Caller must ensure that this variable is VAR_BYREF or VAR_ALIAS and that aTargetVar isn't NULL.
	{
		// Resolve aliases-to-aliases for performance and to increase the expectation of
		// reliability since a chain of aliases-to-aliases might break if an alias in
		// the middle is ever allowed to revert to a non-alias (or gets deleted).
		// A caller may ask to create an alias to an alias when a function calls another
		// function and passes to it one of its own byref-params.
		while (aTargetVar->mType == VAR_ALIAS)
			aTargetVar = aTargetVar->mAliasFor;

		// The following is done only after the above in case there's ever a way for the above
		// to circle back to become this variable.
		// Short-circuit potential infinite loops in other methods by refusing to change an alias
		// to point to itself.  This case probably only matters the first time this alias is
		// updated (i.e. while it's still of type VAR_BYREF), since once it's a valid alias,
		// changing it to point to itself would be handled correctly by the code further below
		// (it would mAliasFor to be the same value that it had previously):
		if (aTargetVar == this)
			return;

		mAliasFor = aTargetVar; // Should always be non-NULL due to various checks elsewhere.
		mType = VAR_ALIAS; // Only actually needed the first time (to convert it from VAR_BYREF).
	}

	ResultType Close(bool aIsBinaryClip = false)
	{
		if (mType == VAR_ALIAS) // For simplicity and reduced code size, just make a recursive call to self.
			return mAliasFor->Close(aIsBinaryClip);
		if (mType == VAR_CLIPBOARD && g_clip.IsReadyForWrite())
			return g_clip.Commit(); // Writes the new clipboard contents to the clipboard and closes it.
		// The binary-clip attribute is also reset here for cases where a caller uses a variable without
		// having called Assign() to resize it first, which can happen if the variable's capacity is already
		// sufficient to hold the desired contents.
		if (aIsBinaryClip)
			mAttrib |= VAR_ATTRIB_BINARY_CLIP;
		else
			mAttrib &= ~VAR_ATTRIB_BINARY_CLIP;
		return OK; // In all other cases.
	}

	Var(char *aVarName, VarTypeType aType, bool aIsLocal)  // Not currently needed (e.g. for VAR_ATTRIB_PARAM): , VarAttribType aAttrib = 0)  [also would need to change mAttrib(aAttrib)] in initializier]
		// The caller must ensure that aVarName is non-null.
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
		: mContents(sEmptyString)
		, mLength(0) // This also initializes mAliasFor within the same union.
		, mCapacity(0), mHowAllocated(ALLOC_NONE)
		, mAttrib(0), mIsLocal(aIsLocal), mType(aType)
		, mName(aVarName) // Caller gave us a pointer to dynamic memory for this (or static in the case of ResolveVarOfArg()).
	{}
	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}
};

#endif
