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
EXTERN_CLIPBOARD;

// Something big enough to be flexible, yet small enough to not be a problem on 99% of systems:
#define MAX_ALLOC_MALLOC (16 * 1024 * 1024)
#define MAX_ALLOC_SIMPLE 64
#define DEREF_BUF_MAX MAX_ALLOC_MALLOC
#define DEREF_BUF_EXPAND_INCREMENT (32 * 1024)
#define ERRORLEVEL_NONE "0"
#define ERRORLEVEL_ERROR "1"
#define ERRORLEVEL_ERROR2 "2"

enum AllocMethod {ALLOC_NONE, ALLOC_SIMPLE, ALLOC_MALLOC};
enum VarTypes
{
  VAR_NORMAL // It's probably best of this one is first (zero).
, VAR_CLIPBOARD
, VAR_SEC, VAR_MIN, VAR_HOUR, VAR_MDAY, VAR_MON, VAR_YEAR, VAR_WDAY, VAR_YDAY
, VAR_WORKINGDIR, VAR_NUMBATCHLINES
, VAR_OSTYPE, VAR_OSVERSION
, VAR_SCRIPTNAME, VAR_SCRIPTDIR, VAR_SCRIPTFULLPATH
, VAR_LOOPFILENAME, VAR_LOOPFILESHORTNAME, VAR_LOOPFILEDIR, VAR_LOOPFILEFULLPATH
, VAR_LOOPFILETIMEMODIFIED, VAR_LOOPFILETIMECREATED, VAR_LOOPFILETIMEACCESSED
, VAR_LOOPFILEATTRIB, VAR_LOOPFILESIZE, VAR_LOOPFILESIZEKB, VAR_LOOPFILESIZEMB
, VAR_THISHOTKEY, VAR_PRIORHOTKEY, VAR_TIMESINCETHISHOTKEY, VAR_TIMESINCEPRIORHOTKEY
, VAR_TICKCOUNT
, VAR_SPACE
};
#define VAR_IS_RESERVED(var) (var->mType != VAR_NORMAL && var->mType != VAR_CLIPBOARD)

typedef UCHAR VarTypeType;     // UCHAR vs. VarTypes to save memory.
typedef UCHAR AllocMethodType; // UCHAR vs. AllocMethod to save memory.
typedef DWORD VarSizeType;  // Up to 4 gig if sizeof(UINT) is 4.  See next line.
#define VARSIZE_MAX MAXDWORD
#define MAX_FORMATTED_NUMBER_LENGTH 255

class Var
{
private:
	char *mContents;
	VarSizeType mLength;   // How much is actually stored in it currently, excluding the zero terminator.
	VarSizeType mCapacity; // In bytes.  Includes the space for the zero terminator.
	AllocMethodType mHowAllocated;
public:
	char *mName;    // The name of the var.
	VarTypeType mType;
	Var *mNextVar;  // Next item in linked list.

	ResultType Assign(int aValueToAssign);
	ResultType Assign(__int64 aValueToAssign);
	ResultType Assign(double aValueToAssign);
	ResultType Assign(char *aBuf = NULL, VarSizeType aLength = VARSIZE_MAX, bool aTrimIt = false);
	VarSizeType Get(char *aBuf = NULL);
	static ResultType ValidateName(char *aName);
	VarSizeType Capacity() {return mCapacity;}
	char *ToText(char *aBuf, size_t aBufSize, bool aAppendNewline)
	// Translates this var into its text equivalent, putting the result into aBuf and
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
		// The caller must ensure that aVarName is non-null and non-empty-string.
		: mName(aVarName) // Caller gave us a pointer to dynamic memory for this.
		, mLength(0)
		, mCapacity(0)
		, mHowAllocated(ALLOC_NONE)
		, mType(aType)
		, mContents("") // This initial empty-string value may be relied upon.
		, mNextVar(NULL)
	{}
	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}
};


#endif
