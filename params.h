#pragma once

#include <string>
#include <iostream>
#include <sstream>

#include "util.h"

struct paramflags
{
	bool noelevate = false;
	bool test = false;
	std::string tee;
};

struct params
{
	std::string instance;
	std::string command;
	std::string subcommand;
	std::string database;
	std::string device;
	std::string from;
	std::string to;
	std::string sql;

	std::string as;
	std::string username;
	std::string password;

	HRESULT hr = S_OK;
	std::string errorMessage;

	paramflags flags;

	friend std::ostream& operator<<(std::ostream& o, const params& p);

	DWORD timeout = 10 * 1000;

	bool isRestore() const
	{
		return iequals(command, "restore");
	}

	bool isBackup() const
	{
		return iequals(command, "backup");
	}

	bool isBackupOrRestore() const
	{
		return isBackup() || isRestore();
	}

	bool isPipe() const
	{
		return iequals(command, "pipe");
	}
};


#ifdef _DEBUG
bool TestParseParams();
#endif

params ParseParams(int argc, const char* argv[], bool quiet);
std::string MakeParams(const params& p);
