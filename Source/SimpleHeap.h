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

#ifndef SimpleHeap_h
#define SimpleHeap_h

#include "stdafx.h" // pre-compiled headers

// May greatly improve the efficiency of dynamic memory for callers that would otherwise want to do many
// small new's or malloc's.  Savings of both RAM space overhead and performance are achieved.  In addition,
// the OS's overall memory fragmentation may be reduced, especially if the apps uses this class over
// a long period of time (hours or days).

// The size of each block in bytes.  Use a size that's a good compromise
// of avg. wastage vs. reducing memory fragmentation and overhead.
// But be careful never to reduce it to something less than LINE_SIZE
// (the maximum line length that can be loaded -- currently 16K), otherwise,
// memory for that line might be impossible to allocate.
// Update: reduced it from 64K to 32K since many scripts tend to be small.
#define BLOCK_SIZE (32 * 1024)

class SimpleHeap
{
private:
	char mBlock[BLOCK_SIZE]; // This object's memory block.  Although private, its contents are public.
	char *mFreeMarker;  // Address inside the above block of the first unused byte.
	size_t mSpaceAvailable;
	static UINT sBlockCount;
	static SimpleHeap *sFirst, *sLast;  // The first and last objects in the linked list.
	static char *sMostRecentlyAllocated; // For use with Delete().
	SimpleHeap *mNextBlock;  // The object after this one in the linked list; NULL if none.

	SimpleHeap();  // Private constructor, since we want only the static methods to be able to create new objects.
	~SimpleHeap();
public:
	static UINT GetBlockCount() {return sBlockCount;}
	static char *Malloc(char *aBuf); // Return a block of memory to the caller and copy aBuf into it.
	static char *Malloc(size_t aSize); // Return a block of memory to the caller.
	static void Delete(void *aPtr);
	static void DeleteAll();
};

#endif
