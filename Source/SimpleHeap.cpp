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

#include "SimpleHeap.h"
#include "globaldata.h" // for g_script, so that errors can be centrally reported here.

// Static member data:
SimpleHeap *SimpleHeap::sFirst = NULL;
SimpleHeap *SimpleHeap::sLast  = NULL;
UINT SimpleHeap::sBlockCount   = 0;

char *SimpleHeap::Malloc(char *aBuf)
{
	char *new_buf = "";
	if (!aBuf || !*aBuf)
		return new_buf; // Return the constant empty string to the caller.
	if (   !(new_buf = SimpleHeap::Malloc(strlen(aBuf) + 1))   ) // +1 for the zero terminator.
	{
		g_script.ScriptError("SimpleHeap::Malloc(buf): Out of memory.", aBuf);
		return NULL;
	}
	strcpy(new_buf, aBuf);
	return new_buf;
}



char *SimpleHeap::Malloc(size_t aSize)
// Seems okay to return char* for convenience, since that's the type most often used.
// This could be made more memory efficient by searching old blocks for sufficient
// free space to handle <size> prior to creating a new block.  But the whole point
// of this class is that it's only called to allocate relatively small objects,
// such as the lines of text in a script file.  The length of such lines is typically
// around 80, and only rarely would exceed 1000.  Trying to find memory in old blocks
// seems like a bad trade-off compared to the performance impact of traversing a
// potentially large linked list or maintaining and traversing an array of
// "under-utilized" blocks.
{
	if (aSize < 1 || aSize > BLOCK_SIZE)
		return NULL;
	if (NULL == sFirst) // We need at least one block to do anything, so create it.
		if (NULL == (sFirst = new SimpleHeap))
			return NULL;
	if (aSize > sLast->mSpaceAvailable)
		if (NULL == (sLast->mNextBlock = new SimpleHeap))
			return NULL;
	char *return_address = sLast->mFreeMarker;
	sLast->mFreeMarker += aSize;
	sLast->mSpaceAvailable -= aSize;
	return return_address;
}



void SimpleHeap::DeleteAll()
// See Hotkey::AllDestructAndExit for comments about why this isn't actually called.
{
	SimpleHeap *next, *curr;
	for (curr = sFirst; curr != NULL;)
	{
		next = curr->mNextBlock;  // Save this member's value prior to deleting the object.
		delete curr;
		curr = next;
	}
}



SimpleHeap::SimpleHeap()  // Construct a new block.
	: mNextBlock(NULL)
	, mFreeMarker(mBlock)  // Starts off pointing to the first byte in the new block.
	, mSpaceAvailable(BLOCK_SIZE)
{
	// Initialize static members here rather than in initializer list, to avoid
	// compiler warning:
	sLast = this;  // Constructing a new block always results in it becoming the current block.
	++sBlockCount;
}


SimpleHeap::~SimpleHeap()
// This destructor is currently never called because all instances of the object are created
// with "new", yet none are ever destroyed with "delete".  As an alternative to this behavior
// the delete method should recursively delete mNextBlock, if it's non-NULL, prior to
// returning.  It seems unnecessary to do this, however, since the whole idea behind this
// class is that it's a simple implementation of one-time, persistent memory allocation.
// It's not intended to permit deallocation and subsequent reclamation of freed fragments
// within the collection of blocks.  When the program exits, all memory dynamically
// allocated by the constructor and any other methods that call "new" will be reclaimed
// by the OS.  UPDATE: This is now called by static method DeleteAll().
{
	return;
}
