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

// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#define WIN32_LEAN_AND_MEAN		// Exclude rarely-used stuff from Windows headers

// Windows Header Files:
// Necessary to do this prior to including windows.h so that NT functions are unlocked:
// UPDATE: Using 0x0500 now so that VK_XBUTTON1 and 2 can be supported:
#define _WIN32_WINNT 0x0500



// AutoIt3: Only use the includes here if running Visual C - it is really slow under
// MingW to use this method so we will comment out the headers unless the
// compiler is VC (and manually include the required headers in the source files)
#ifdef _MSC_VER
	// C RunTime Header Files
	#include <stdio.h>
	#include <stdlib.h>
	#include <time.h> // for the c-lib time functions, to support A_YDAY, etc.
	#include <sys/timeb.h> // for _timeb struct
	#include <stdarg.h> // used by snprintfcat()
	#include <limits.h>  // for UINT_MAX, UCHAR_MAX, etc.
	//#include <malloc.h>
	//#include <memory.h>
	//#include <tchar.h>

	#include <windows.h>
	#include <commctrl.h> // for status bar functions. Must be included after <windows.h>.
	#include <shellapi.h>  // for ShellExecute()
	#include <shlobj.h>  // for SHGetMalloc()
	#include <mmsystem.h> // for mciSendString() and waveOutSetVolume()
	#include <commdlg.h> // for OPENFILENAME

	// It's probably best not to do these, because I think they would then be included
	// for everything, even modules that don't need it, which might result in undesired
	// dependencies sneaking in, or subtle naming conflicts:
	// ...
	//#include "defines.h"
	//#include "application.h"
	//#include "globaldata.h"
	//#include "window.h"  // Not to be confused with "windows.h"
	//#include "util.h"
	//#include "SimpleHeap.h"
#endif
