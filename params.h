#pragma once

#include <string>
#include <iostream>
#include <sstream>

#include "util.h"

struct paramflags
{
	bool noelevate = false;
	bool test = false;
};

struct params
{
	std::wstring instance;
	std::wstring command;
	std::wstring subcommand;
	std::wstring database;
	std::wstring device;
	std::wstring from;
	std::wstring to;

	std::wstring as;
	std::wstring username;
	std::wstring password;

	HRESULT hr = S_OK;
	std::wstring errorMessage;

	paramflags flags;

	friend std::wostream& operator<<(std::wostream& o, const params& p);

	DWORD timeout = 10 * 1000;

	bool isRestore() const
	{
		return iequals(command, L"restore");
	}

	bool isBackup() const
	{
		return iequals(command, L"backup");
	}

	bool isBackupOrRestore() const
	{
		return isBackup() || isRestore();
	}

	bool isPipe() const
	{
		return iequals(command, L"pipe");
	}
};


#ifdef _DEBUG
bool TestParseParams();
#endif

params ParseParams(int argc, wchar_t* argv[], bool quiet);
std::wstring MakeParams(const params& p);
