// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#define _WIN32_WINNT 0x0501

#include <WinSDKVer.h>
#include <SDKDDKVer.h>

#include <stdio.h>
#include <tchar.h>
#include <io.h>
#include <fcntl.h>

#include <windows.h>

#include <comutil.h>
#include <comdef.h>

#include <ShlObj.h>

#include <array>
#include <set>

#include <string>
#include <algorithm>

#include <iostream>
#include <sstream>
#include <fstream>
#include <iomanip>

#include <thread>
#include <mutex>
#include <future>


#include "vdi/vdi.h"
#include "vdi/vdierror.h"

//#import "msado60_Backcompat.tlb" rename("EOF", "eof") rename("BOF", "bof")

#include "msado60_backcompat.tlh"

#include <assert.h>



