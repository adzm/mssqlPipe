#include "stdafx.h"

#include "params.h"

#include "util.h"


void showUsage()
{
	std::wcerr << LR"(
Usage:

mssqlPipe [instance] [as username[:password]] (backup|restore|pipe) ... 

... backup [database] dbname [to filename]
... restore [database] dbname [from filename] [to filepath] [with replace]
... restore filelistonly [from filename]
... pipe (to|from) devicename [(to|from) filename]

stdin or stdout will be used if no filenames specified. Windows authentication
(SSPI) will be used if [as username[:password]] is not specified.

Examples:

mssqlPipe myinstance backup AdventureWorks to AdventureWorks.bak
mssqlPipe myinstance as sa:hunter2 backup AdventureWorks > AdventureWorks.bak
mssqlPipe backup database AdventureWorks | 7za a AdventureWorks.xz -txz -si
7za e AdventureWorks.xz -so | mssqlPipe restore AdventureWorks to c:/db/
mssqlPipe restore AdventureWorks from AdventureWorks.bak with replace
mssqlPipe pipe from VirtualDevice42 > output.bak
mssqlPipe pipe to VirtualDevice42 < input.bak

Happy piping!
)";

	std::wcerr << std::endl;
}

inline std::wstring escapeCommandLine(std::wstring str)
{
	if (std::wstring::npos == str.find_first_of(L" \"")) {
		return str;
	}

	size_t quoteCount = count(begin(str), end(str), L'\"');

	std::wstring q;
	q.reserve(2 + str.size() + (quoteCount * 3));

	q += L'\"';

	for (auto c : str) {
		q.push_back(c);
		if (c == L'\"') {
			// triple quotes should work even outside of another quoted field
			q += L"\"\"\"";
		}
	}

	q += L'\"';

	return q;
}

std::wstring makeCommandLineDiff(std::wstring cmdLineA, std::wstring cmdLineB)
{
	auto extractArgs = [](std::wstring str) {
		std::vector<std::wstring> args;

		int argc = 0;
		wchar_t** argv = ::CommandLineToArgvW(str.c_str(), &argc);

		args.reserve(argc);

		for (int i = 0; i < argc; ++i) {
			args.push_back(argv[i]);
		}

		::LocalFree(argv);

		return args;
	};

	auto argsA = extractArgs(cmdLineA);
	auto argsB = extractArgs(cmdLineB);

	std::wstring result;
	std::wstring postResult;

	auto itA = argsA.begin();
	auto itB = argsB.begin();

	while (itA < argsA.end() || itB < argsB.end()) {

		std::wstring a;
		std::wstring b;

		if (itA < argsA.end()) {
			a = *itA;
		}

		if (itB < argsB.end()) {
			b = *itB;
		}

		if (iequals(a, b)) {
			++itA;
			++itB;
			//result += L".";
			continue;
		}
		else {
			break;
		}
	}

	if (itA == argsA.end() && itB == argsB.end()) {
		return result;
	}

	// we are at a diff!

	// now find the tail
	auto itA_end = argsA.end() - 1;
	auto itB_end = argsB.end() - 1;

	while (itA_end >= itA && itB_end >= itB) {

		std::wstring a = *itA_end;
		std::wstring b = *itB_end;

		if (iequals(a, b)) {
			--itA_end;
			--itB_end;
			//postResult += L".";
			continue;
		}
		else {
			break;
		}
	}

	result += L"(`";
	while (itA <= itA_end) {
		result += *itA;
		if (itA < itA_end) {
			result += L" ";
		}
		++itA;
	}
	result += L"`, `";
	while (itB <= itB_end) {
		result += *itB;
		if (itB < itB_end) {
			result += L" ";
		}
		++itB;
	}
	result += L"`)";

	result += postResult;

	return result;
}

params ParseParams(int argc, wchar_t* argv[], bool quiet)
{
	params p;

	auto invalidArgs = [&](const wchar_t* msg, const wchar_t* arg = nullptr, HRESULT hr = E_INVALIDARG)
	{
		if (!quiet) {
			std::wcerr << msg;

			if (arg && *arg) {
				std::wcerr << L" `" << arg << L"`";
			}

			std::wcerr << std::endl;

			showUsage();
		}

		if (msg && *msg) {
			if (!p.errorMessage.empty()) {
				p.errorMessage += L"\r\n";
			}
			p.errorMessage += msg;
			if (arg && *arg) {
				p.errorMessage += L" `";
				p.errorMessage += arg;
				p.errorMessage += L"`";
			}
		}

		if (!hr) {
			hr = E_FAIL;
		}

		p.hr = hr;

		return p;
	};

	if (argc < 2) {
		if (!quiet) {
			showUsage();
		}
		return p;
	}


	p.device = make_guid();

	wchar_t** arg = argv;
	wchar_t** argEnd = arg + argc;
	wchar_t** argFirst = arg + 1;
	wchar_t** argVerb = nullptr;
	wchar_t** argVerbEnd = nullptr;

	++arg;

	// mssqlpipe [instance] [as username[:password]] (backup|restore [filelistonly]) [database] dbname [ [from|to] filename|path ] ...

	// first off we will find the verb!
	{
		while (arg < argEnd) {
			auto sz = *arg;
			if (iequals(sz, L"backup")) {
				argVerb = arg++;
				p.command = ToLower(sz);
				if (arg < argEnd && iequals(*arg, L"database")) {
					++arg;
				}
				break;
			}
			else if (iequals(sz, L"restore")) {
				argVerb = arg++;
				p.command = ToLower(sz);
				if (arg < argEnd && iequals(*arg, L"database")) {
					++arg;
				}
				if (arg < argEnd && iequals(*arg, L"filelistonly")) {
					p.subcommand = ToLower(*arg);
					++arg;
				}
				break;
			}
			else if (iequals(sz, L"pipe")) {
				argVerb = arg++;
				p.command = ToLower(sz);
				break;
			}

			++arg;
		}

		argVerbEnd = arg;
	}

	// if argVerb is null, we are doomed

	if (!argVerb) {
		return invalidArgs(L"missing verb");
	}

	// so now we have two ranges: 
	// [argFirst...argVerb) and [argVerbEnd...argEnd)

	// start by parsing the first range, which has instance and authentication
	{
		arg = argFirst;

		// first is instance
		if (arg < argVerb && !iequals(*arg, L"as")) {
			p.instance = *arg;
			++arg;
		}
		if (arg < argVerb) {
			if (iequals(*arg, L"as")) {
				++arg;
				if (arg < argVerb) {
					p.as = *arg;
					++arg;

					// parse creds into username and password
					size_t split = p.as.find(L':');
					if (split == std::wstring::npos) {
						p.username = p.as;
					}
					else {
						p.username = p.as.substr(0, split);
						p.password = p.as.substr(split + 1);
					}
				}
				else {
					return invalidArgs(L"invalid authentication");
				}
			}
		}

		if (arg < argVerb) {
			return invalidArgs(L"invalid instance or authentication");
		}

		while (arg < argVerb) {
			auto sz = *arg;
			if (iequals(sz, L"as")) {
			}
			else {
				p.instance = sz;
				++arg;
			}
		}
	}

	// now we can parse specific options for the verbs

	arg = argVerbEnd;


	if (p.isPipe()) {
		if (arg >= argEnd) {
			return invalidArgs(L"arguments missing");
		}

		// pipe can have either to or from followed by device name

		if (arg < argEnd && iequals(*arg, L"to")) {
			p.subcommand = ToLower(*arg);
		}
		else if (arg < argEnd && iequals(*arg, L"from")) {
			p.subcommand = ToLower(*arg);
		}
		else {
			return invalidArgs(L"pipe requires `to` or `from` and a device name");
		}

		++arg;

		if (arg >= argEnd) {
			return invalidArgs(L"pipe requires a device name");
		}

		p.device = *arg;

		++arg;

		if (arg < argEnd && p.subcommand == L"from" && iequals(*arg, L"to")) {
			++arg;

			if (arg >= argEnd) {
				return invalidArgs(L"missing file name");
			}

			p.to = *arg;
			++arg;
		}
		else if (arg < argEnd && p.subcommand == L"to" && iequals(*arg, L"from")) {
			++arg;

			if (arg >= argEnd) {
				return invalidArgs(L"missing file name");
			}

			p.from = *arg;
			++arg;
		}

		if (arg < argEnd) {
			return invalidArgs(L"extra args at end");
		}
	}
	else if (p.isRestore() && p.subcommand == L"filelistonly") {
		// filelistonly 

		if (arg < argEnd && iequals(*arg, L"from")) {
			++arg;

			if (arg >= argEnd) {
				return invalidArgs(L"missing file name");
			}

			p.from = *arg;
			++arg;
		}

		if (arg < argEnd) {
			return invalidArgs(L"extra args at end");
		}
	}
	else if (p.isBackupOrRestore()) {

		// parse database name
		if (arg < argEnd) {
			p.database = *arg;
			++arg;
		}

		if (p.isBackup()) {

			// to file
			if (arg < argEnd && iequals(*arg, L"to")) {
				++arg;

				if (arg >= argEnd) {
					return invalidArgs(L"missing file name");
				}

				p.to = *arg;
				++arg;
			}

			// TODO with copy only, not copy only, etc?

			if (arg < argEnd && iequals(*arg, L"with")) {
				++arg;

				if (arg >= argEnd) {
					return invalidArgs(L"missing option after with");
				}

				// no with options for backup yet

				return invalidArgs(L"invalid with option");
			}

			if (arg < argEnd) {
				return invalidArgs(L"extra args at end");
			}

		}
		else if (p.isRestore()) {

			// from file
			if (arg < argEnd && iequals(*arg, L"from")) {
				++arg;

				if (arg >= argEnd) {
					return invalidArgs(L"missing file name");
				}

				p.from = *arg;
				++arg;
			}

			// to path
			if (arg < argEnd && iequals(*arg, L"to")) {
				++arg;

				if (arg >= argEnd) {
					return invalidArgs(L"missing restore path");
				}

				p.to = *arg;
				++arg;
			}

			// with replace

			if (arg < argEnd && iequals(*arg, L"with")) {
				++arg;

				if (arg >= argEnd) {
					return invalidArgs(L"missing option after with");
				}

				if (iequals(*arg, L"replace")) {
					p.subcommand = ToLower(*arg);
					++arg;
				}
				else {
					return invalidArgs(L"invalid with option");
				}
			}
		}
		else {
			return invalidArgs(L"something horrible");
		}
	}

	return p;
}

std::wstring MakeParams(const params& p)
{
	std::wstring commandLine;
	commandLine.reserve(260);

	auto append = [&commandLine](std::wstring str) {
		if (str.empty()) {
			return;
		}

		if (!commandLine.empty()) {
			commandLine += L" ";
		}

		commandLine += escapeCommandLine(str);
	};

	if (!p.instance.empty()) {
		append(p.instance);
	}
	if (!p.as.empty()) {
		append(L"as");
		append(p.as);
	}

	assert(!p.command.empty());

	append(p.command);
	if (p.isRestore() && p.subcommand == L"filelistonly") {
		append(p.subcommand);

		if (!p.from.empty()) {
			append(L"from");
			append(p.from);
		}
	}
	else if (p.isPipe()) {

		assert(!p.subcommand.empty());

		append(p.subcommand);

		assert(!p.device.empty());

		append(p.device);

		assert((p.to.empty() && p.from.empty()) || (p.to.empty() && !p.from.empty()) || (p.from.empty() && !p.to.empty()));

		if (!p.to.empty()) {
			append(L"to");
			append(p.to);
		}
		if (!p.from.empty()) {
			append(L"from");
			append(p.from);
		}
	}
	else if (p.isBackup()) {

		assert(!p.database.empty());

		append(p.database);

		if (!p.to.empty()) {
			append(L"to");
			append(p.to);
		}

		assert(p.from.empty());
	}
	else if (p.isRestore()) {

		assert(!p.database.empty());

		append(p.database);

		if (!p.from.empty()) {
			append(L"from");
			append(p.from);
		}

		if (!p.to.empty()) {
			append(L"to");
			append(p.to);
		}

		if (p.subcommand == L"replace") {
			append(L"with");
			append(p.subcommand);
		}
	}

	return commandLine;
}

std::wostream& operator<<(std::wostream& o, const params& p)
{
	o << L"(";
	if (S_OK != p.hr) {
		o << L"hr=" << p.hr << L";";
	}
	if (!p.instance.empty()) {
		o << L"instance=" << p.instance << L";";
	}
	if (!p.command.empty()) {
		o << L"command=" << p.command << L";";
	}
	if (!p.subcommand.empty()) {
		o << L"subcommand=" << p.subcommand << L";";
	}
	if (!p.database.empty()) {
		o << L"database=" << p.database << L";";
	}
	if (!p.from.empty()) {
		o << L"from=" << p.from << L";";
	}
	if (!p.to.empty()) {
		o << L"to=" << p.to << L";";
	}
	if (!p.as.empty()) {
		o << L"as=" << p.as << L";";
	}
	if (!p.username.empty()) {
		o << L"username=" << p.username << L";";
	}
	if (!p.password.empty()) {
		o << L"password=" << p.password << L";";
	}

	if (p.isPipe()) {
		o << L"device=" << p.device << L";";
	}

	o << L")";
	return o;
}

#ifdef _DEBUG
bool TestParseParams()
{
	auto test = [](const wchar_t* cmdLine) {
		int argc = 0;
		wchar_t** argv = ::CommandLineToArgvW(cmdLine, &argc);

		if (!argv) {
			return false;
		}

		auto p = ParseParams(argc, argv, true);

		::LocalFree(argv);

		auto rebuilt = L"mssqlPipe " + MakeParams(p);

		bool succeeded = SUCCEEDED(p.hr);
		bool matches = iequals(cmdLine, rebuilt) || iequals(L"(`database`, ``)", makeCommandLineDiff(cmdLine, rebuilt));

		argc = 0;
		argv = ::CommandLineToArgvW(rebuilt.c_str(), &argc);

		if (!argv) {
			return false;
		}

		auto p_rebuilt = ParseParams(argc, argv, true);

		::LocalFree(argv);

		auto p_rebuilt_rebuilt = L"mssqlPipe " + MakeParams(p_rebuilt);

		bool rebuilt_matches = iequals(rebuilt, p_rebuilt_rebuilt);

		if (!succeeded || !matches || !rebuilt_matches) {
			if (!succeeded) {
				std::wcerr << L"FAILED! " << p.errorMessage << L" (0x" << std::hex << p.hr << std::dec << L")" << std::endl;
			}
			if (!matches) {
				std::wcerr << L"MISMATCH! " << makeCommandLineDiff(cmdLine, rebuilt) << std::endl;
			}
			if (!rebuilt_matches) {
				std::wcerr << L"REBUILT MISMATCH! " << makeCommandLineDiff(rebuilt, p_rebuilt_rebuilt) << std::endl;
			}
			std::wcerr << L">\t`" << cmdLine << L"`" << std::endl;
			std::wcerr << L"<\t`" << rebuilt << L"`" << std::endl;
			std::wcerr << L"=\t" << p << std::endl;
			if (!rebuilt_matches) {
				std::wcerr << L"*<\t`" << p_rebuilt_rebuilt << L"`" << std::endl;
				std::wcerr << L"*=\t" << p_rebuilt << std::endl;
			}
			std::wcerr << std::endl;
			return false;
		}

		return true;
	};

	// arguments before verb
	if (!test(L"mssqlPipe backup AdventureWorks")) { return false; }
	if (!test(L"mssqlPipe myinstance backup AdventureWorks")) { return false; }
	if (!test(L"mssqlPipe myinstance as sa:hunter2 backup AdventureWorks")) { return false; }

	// backup
	if (!test(L"mssqlPipe backup AdventureWorks")) { return false; }
	if (!test(L"mssqlPipe backup database AdventureWorks")) { return false; }
	if (!test(L"mssqlPipe backup AdventureWorks to z:/db")) { return false; }

	// restore
	if (!test(L"mssqlPipe restore AdventureWorks")) { return false; }
	if (!test(L"mssqlPipe restore database AdventureWorks")) { return false; }
	if (!test(L"mssqlPipe restore AdventureWorks from z:/db/AdventureWorks.bak")) { return false; }
	if (!test(L"mssqlPipe restore AdventureWorks with replace")) { return false; }
	if (!test(L"mssqlPipe restore AdventureWorks from z:/db/AdventureWorks.bak with replace")) { return false; }

	// restore filelistonly
	if (!test(L"mssqlPipe restore filelistonly")) { return false; }
	if (!test(L"mssqlPipe restore filelistonly from z:/db/AdventureWorks.bak")) { return false; }

	// pipe
	if (!test(L"mssqlPipe pipe to VirtualDevice")) { return false; }
	if (!test(L"mssqlPipe pipe from VirtualDevice")) { return false; }
	if (!test(L"mssqlPipe pipe to VirtualDevice from AdventureWorks.bak")) { return false; }
	if (!test(L"mssqlPipe pipe from VirtualDevice to AdventureWorks.bak")) { return false; }

	return true;
}
#endif