/*
AutoHotkey

Copyright 2003-2006 Chris Mallett (support@autohotkey.com)

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
#define DEREF_BUF_EXPAND_INCREMENT (32 * 1024) // 32 seems to work well, so might not be ever a good enough reason to ever increase it.
#define ERRORLEVEL_NONE "0"
#define ERRORLEVEL_ERROR "1"
#define ERRORLEVEL_ERROR2 "2"

enum AllocMethod {ALLOC_NONE, ALLOC_SIMPLE, ALLOC_MALLOC};
enum VarTypes
{
  VAR_INVALID
// VAR_BYREF is used when a VAR_ALIAS variable doesn't yet have a target, which avoids the need to check whether
// mAliasFor is NULL in many places.  In other words, a var of type VAR_ALIAS always has a non-NULL mAliasFor.
, VAR_BYREF // VAR_BYREF and VAR_ALIAS are for internal use (external users of variables never see them). VAR_BYREF is
, VAR_ALIAS // never seen because it has always become VAR_ALIAS by the time anyone uses them.  ALIAS is never seen because external users call Var::Type(), which automatically resolves ALIAS to some other type.
// ABOVE MUST BE KEPT FIRST AND MUST BE KEPT ADJACENT TO VAR_FIRST_NON_BYREF BELOW. 
, VAR_FIRST_NON_BYREF, VAR_NORMAL = VAR_FIRST_NON_BYREF  // Most variables are this type.
, VAR_CLIPBOARD
, VAR_LAST_UNRESERVED = VAR_CLIPBOARD  // Keep this in sync with any changes to the set of unreserved variables.
#define VAR_IS_RESERVED(var) ((var).Type() > VAR_LAST_UNRESERVED)
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
, VAR_COMSPEC, VAR_WINDIR, VAR_TEMP, VAR_PROGRAMFILES, VAR_APPDATA, VAR_APPDATACOMMON, VAR_DESKTOP, VAR_DESKTOPCOMMON
, VAR_STARTMENU, VAR_STARTMENUCOMMON
, VAR_PROGRAMS, VAR_PROGRAMSCOMMON, VAR_STARTUP, VAR_STARTUPCOMMON, VAR_MYDOCUMENTS
, VAR_ISADMIN, VAR_CURSOR, VAR_CARETX, VAR_CARETY
, VAR_SCREENWIDTH, VAR_SCREENHEIGHT
, VAR_IPADDRESS1, VAR_IPADDRESS2, VAR_IPADDRESS3, VAR_IPADDRESS4
, VAR_SCRIPTNAME, VAR_SCRIPTDIR, VAR_SCRIPTFULLPATH, VAR_LINENUMBER, VAR_LINEFILE, VAR_ISCOMPILED
, VAR_LOOPFILENAME, VAR_LOOPFILESHORTNAME, VAR_LOOPFILEEXT, VAR_LOOPFILEDIR
, VAR_LOOPFILEFULLPATH, VAR_LOOPFILELONGPATH, VAR_LOOPFILESHORTPATH
, VAR_LOOPFILETIMEMODIFIED, VAR_LOOPFILETIMECREATED, VAR_LOOPFILETIMEACCESSED
, VAR_LOOPFILEATTRIB, VAR_LOOPFILESIZE, VAR_LOOPFILESIZEKB, VAR_LOOPFILESIZEMB
, VAR_LOOPREGTYPE, VAR_LOOPREGKEY, VAR_LOOPREGSUBKEY, VAR_LOOPREGNAME, VAR_LOOPREGTIMEMODIFIED
, VAR_LOOPREADLINE, VAR_LOOPFIELD, VAR_INDEX
, VAR_THISMENUITEM, VAR_THISMENUITEMPOS, VAR_THISMENU, VAR_THISHOTKEY, VAR_PRIORHOTKEY
, VAR_TIMESINCETHISHOTKEY, VAR_TIMESINCEPRIORHOTKEY
, VAR_ENDCHAR, VAR_LASTERROR, VAR_GUI, VAR_GUICONTROL, VAR_GUICONTROLEVENT, VAR_EVENTINFO
, VAR_GUIWIDTH, VAR_GUIHEIGHT, VAR_GUIX, VAR_GUIY
, VAR_TIMEIDLE, VAR_TIMEIDLEPHYSICAL
, VAR_SPACE, VAR_TAB, VAR_AHKVERSION, VAR_AHKPATH
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
	VarTypeType mType;  // Keep adjacent/contiguous with the above due to struct alignment, to save memory.

public:
	// Testing shows that due to data alignment, keeping mType adjacent to the other less-than-4-size member
	// above it reduces size of each object by 4 bytes.
	char *mName;    // The name of the var.

	// sEmptyString is a special *writable* memory area for empty variables (those with zero capacity).
	// Although making it writable does make buffer overflows difficult to detect and analyze (since they
	// tend to corrupt the program's static memory pool), the advantages in maintainability and robustness
	// see to far outweigh that.  For example, it avoids having to constantly think about whether
	// *Contents()='\0' is safe. The sheer number of places that's avoided is a great relief, and it also
	// cuts down on code size due to not having to always check Capacity() and/or create more functions to
	// protect from writing to read-only strings, which would hurt performance.
	// The biggest offender of buffer overflow in sEmptyString is DllCall, which happens most frequently
	// when a script forgets to call VarSetCapacity before psssing a buffer to some function that writes a
	// string to it.  That drawback has been addressed by passing the read-only empty string in place of
	// sEmptyString in DllCall(), which forces an exception to occur immediately, which is caught by the
	// exception handler there.
	static char sEmptyString[1]; // See above.

	ResultType AssignHWND(HWND aWnd);
	ResultType Assign(DWORD aValueToAssign);
	ResultType Assign(int aValueToAssign);
	ResultType Assign(__int64 aValueToAssign);
	//ResultType Assign(unsigned __int64 aValueToAssign);
	ResultType Assign(double aValueToAssign);
	ResultType Assign(ExprTokenType &aToken);
	ResultType AssignClipboardAll();
	ResultType AssignBinaryClip(Var &aSourceVar);
	ResultType Assign(char *aBuf = NULL, VarSizeType aLength = VARSIZE_MAX, bool aExactSize = false, bool aObeyMaxMem = true);
	VarSizeType Get(char *aBuf = NULL);

	// Not an enum so that it can be global more easily:
	#define VAR_ALWAYS_FREE                    0 // This item and the next must be first and numerically adjacent to
	#define VAR_ALWAYS_FREE_BUT_EXCLUDE_STATIC 1 // each other so that VAR_ALWAYS_FREE_LAST covers only them.
	#define VAR_ALWAYS_FREE_LAST               2 // Never actually passed as a parameter, just a placeholder (see above comment).
	#define VAR_NEVER_FREE                     3
	#define VAR_FREE_IF_LARGE                  4
	void Free(int aWhenToFree = VAR_ALWAYS_FREE, bool aExcludeAliases = false);
	ResultType AppendIfRoom(char *aStr, VarSizeType aLength);
	void AcceptNewMem(char *aNewMem, VarSizeType aLength);
	void SetLengthFromContents();

	static ResultType BackupFunctionVars(Func &aFunc, VarBkp *&aVarBackup, int &aVarBackupCount);
	void Backup(VarBkp &aVarBkp);
	static void FreeAndRestoreFunctionVars(Func &aFunc, VarBkp *&aVarBackup, int &aVarBackupCount);

	#define DISPLAY_NO_ERROR   0  // Must be zero.
	#define DISPLAY_VAR_ERROR  1
	#define DISPLAY_FUNC_ERROR 2
	static ResultType ValidateName(char *aName, bool aIsRuntime = false, int aDisplayError = DISPLAY_VAR_ERROR);

	char *ToText(char *aBuf, int aBufSize, bool aAppendNewline)
	// aBufSize is an int so that any negative values passed in from caller are not lost.
	// Caller has ensured that aBuf isn't NULL.
	// Translates this var into its text equivalent, putting the result into aBuf andp
	// returning the position in aBuf of its new string terminator.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		// v1.0.44.14: Changed it so that ByRef/Aliases report their own name rather than the target's/caller's
		// (it seems more useful and intuitive).
		char *aBuf_orig = aBuf;
		aBuf += snprintf(aBuf, BUF_SPACE_REMAINING, "%s[%u of %u]: %-1.60s%s", mName // mName not var.mName (see comment above).
			, var.mLength, var.mCapacity ? (var.mCapacity - 1) : 0  // Use -1 since it makes more sense to exclude the terminator.
			, var.mContents, var.mLength > 60 ? "..." : "");
		if (aAppendNewline && BUF_SPACE_REMAINING >= 2)
		{
			*aBuf++ = '\r';
			*aBuf++ = '\n';
			*aBuf = '\0';
		}
		return aBuf;
	}

	__forceinline VarTypeType Type()
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		return (mType == VAR_ALIAS) ? mAliasFor->mType : mType;
	}

	__forceinline bool IsByRef()
	{
		return mType < VAR_FIRST_NON_BYREF;
	}

	__forceinline bool IsLocal()
	{
		// Since callers want to know whether this variable is local, even if it's a local alias for a
		// global, don't use the method below:
		//    return (mType == VAR_ALIAS) ? mAliasFor->mIsLocal : mIsLocal;
		return mIsLocal;
	}

	__forceinline bool IsNonStaticLocal()
	{
		// Since callers want to know whether this variable is local, even if it's a local alias for a
		// global, don't use resolve VAR_ALIAS.
		// Even a ByRef local is considered local here because callers are interested in whether this
		// variable can vary from call to call to the same function (and a ByRef can vary in what it
		// points to).  Variables that vary can thus be altered by the backup/restore process.
		return mIsLocal && !(mAttrib & VAR_ATTRIB_STATIC);
	}

	__forceinline bool IsBinaryClip()
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		return (mType == VAR_ALIAS ? mAliasFor->mAttrib : mAttrib) & VAR_ATTRIB_BINARY_CLIP;
	}

	void OverwriteAttrib(VarAttribType aAttrib)
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		if (mType == VAR_ALIAS)
			mAliasFor->mAttrib = aAttrib;
		else
			mAttrib = aAttrib;
	}

	VarSizeType Capacity() // __forceinline() on Capacity, Length, and/or Contents bloats the code and reduces performance.
	// Capacity includes the zero terminator (though if capacity is zero, there will also be a zero terminator in mContents due to it being "").
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		// Fix for v1.0.37: Callers want the clipboard's capacity returned, if it has a capacity.  This is
		// because Capacity() is defined as being the size available in Contents(), which for the clipboard
		// would be a pointer to the clipboard-buffer-to-be-written (or zero if none).
		return var.mType == VAR_CLIPBOARD ? g_clip.mCapacity : var.mCapacity;
	}

	VarSizeType &Length() // __forceinline() on Capacity, Length, and/or Contents bloats the code and reduces performance.
	// This should not be called to discover a non-NORMAL var's length (nor that of an environment variable)
	// because their lengths aren't knowable without calling Get().
	// Returns a reference so that caller can use this function as an lvalue.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		if (var.mType == VAR_NORMAL)
			return var.mLength;
		// Since the length of the clipboard isn't normally tracked, we just return a
		// temporary storage area for the caller to use.  Note: This approach is probably
		// not thread-safe, but currently there's only one thread so it's not an issue.
		// For reserved vars do the same thing as above, but this function should never
		// be called for them:
		static VarSizeType length; // Must be static so that caller can use its contents. See above.
		return length;
	}

	VarSizeType LengthIgnoreBinaryClip()
	// Returns 0 for types other than VAR_NORMAL and VAR_CLIPBOARD.
	// IMPORTANT: Environment variables aren't supported here, so caller must either want such
	// variables treated as blank, or have already checked that they're not environment variables.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		// Return the apparent length of the string (i.e. the position of its first binary zero).
		return (var.mType == VAR_NORMAL && !(var.mAttrib & VAR_ATTRIB_BINARY_CLIP))
			? var.mLength : strlen(var.Contents()); // Use Contents() vs. mContents to support VAR_CLIPBOARD.
	}

	char *Contents() // __forceinline() on Capacity, Length, and/or Contents bloats the code and reduces performance.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		if (var.mType == VAR_NORMAL)
			return var.mContents;
		if (var.mType == VAR_CLIPBOARD)
			// The returned value will be a writable mem area if clipboard is open for write.
			// Otherwise, the clipboard will be opened physically, if it isn't already, and
			// a pointer to its contents returned to the caller:
			return g_clip.Contents();
		return sEmptyString; // For reserved vars (but this method should probably never be called for them).
	}

	__forceinline Var *ResolveAlias()
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		return (mType == VAR_ALIAS) ? mAliasFor : this; // Return target if it's an alias, or itself if not.
	}

	__forceinline void UpdateAlias(Var *aTargetVar) // __forceinline because it's currently only called from one place.
	// Caller must ensure that this variable is VAR_BYREF or VAR_ALIAS and that aTargetVar isn't NULL.
	{
		// BELOW IS THE MEANS BY WHICH ALIASES AREN'T ALLOWED TO POINT TO OTHER ALIASES, ONLY DIRECTLY TO THE TARGET VAR.
		// Resolve aliases-to-aliases for performance and to increase the expectation of
		// reliability since a chain of aliases-to-aliases might break if an alias in
		// the middle is ever allowed to revert to a non-alias (or gets deleted).
		// A caller may ask to create an alias to an alias when a function calls another
		// function and passes to it one of its own byref-params.
		while (aTargetVar->mType == VAR_ALIAS)
			aTargetVar = aTargetVar->mAliasFor;

		// The following is done only after the above in case there's ever a way for the above
		// to circle back to become this variable.
		// Prevent potential infinite loops in other methods by refusing to change an alias
		// to point to itself.  This case probably only matters the first time this alias is
		// updated (i.e. while it's still of type VAR_BYREF), since once it's a valid alias,
		// changing it to point to itself would be handled correctly by the code further below
		// (it would set mAliasFor to be the same value that it had previously):
		if (aTargetVar == this)
			return;

		mAliasFor = aTargetVar; // Should always be non-NULL due to various checks elsewhere.
		mType = VAR_ALIAS; // Only actually needed the first time (to convert it from VAR_BYREF).
	}

	ResultType Close(bool aIsBinaryClip = false)
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		if (var.mType == VAR_CLIPBOARD && g_clip.IsReadyForWrite())
			return g_clip.Commit(); // Writes the new clipboard contents to the clipboard and closes it.
		// The binary-clip attribute is also reset here for cases where a caller uses a variable without
		// having called Assign() to resize it first, which can happen if the variable's capacity is already
		// sufficient to hold the desired contents.
		if (aIsBinaryClip)
			var.mAttrib |= VAR_ATTRIB_BINARY_CLIP;
		else
			var.mAttrib &= ~VAR_ATTRIB_BINARY_CLIP;
		return OK; // In all other cases.
	}

	// Constructor:
	Var(char *aVarName, VarTypeType aType, bool aIsLocal)  // Not currently needed (e.g. for VAR_ATTRIB_PARAM): , VarAttribType aAttrib = 0)  [also would need to change mAttrib(aAttrib)] in initializier]
		// The caller must ensure that aVarName is non-null.
		: mCapacity(0), mContents(sEmptyString) // Invariant: Anyone setting mCapacity to 0 must also set mContents to the empty string.
		, mLength(0) // This also initializes mAliasFor within the same union.
		, mHowAllocated(ALLOC_NONE)
		, mAttrib(0), mIsLocal(aIsLocal), mType(aType)
		, mName(aVarName) // Caller gave us a pointer to dynamic memory for this (or static in the case of ResolveVarOfArg()).
	{}
	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}
};

#endif
