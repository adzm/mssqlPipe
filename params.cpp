#include "stdafx.h"

#include "params.h"

#include "util.h"


void showUsage()
{
	std::cerr << R"(
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

	std::cerr << std::endl;
}

inline std::string escapeCommandLine(std::string str)
{
	if (std::string::npos == str.find_first_of(" \"")) {
		return str;
	}

	size_t quoteCount = count(begin(str), end(str), L'\"');

	std::string q;
	q.reserve(2 + str.size() + (quoteCount * 3));

	q += L'\"';

	for (auto c : str) {
		q.push_back(c);
		if (c == L'\"') {
			// triple quotes should work even outside of another quoted field
			q += "\"\"\"";
		}
	}

	q += L'\"';

	return q;
}

std::string makeCommandLineDiff(std::string cmdLineA, std::string cmdLineB)
{
	auto extractArgs = [](std::string str) {
		std::vector<std::string> args;

		int argc = 0;
		char** argv = ::CommandLineToArgvW(str.c_str(), &argc);

		args.reserve(argc);

		for (int i = 0; i < argc; ++i) {
			args.push_back(argv[i]);
		}

		::LocalFree(argv);

		return args;
	};

	auto argsA = extractArgs(cmdLineA);
	auto argsB = extractArgs(cmdLineB);

	std::string result;
	std::string postResult;

	auto itA = argsA.begin();
	auto itB = argsB.begin();

	while (itA < argsA.end() || itB < argsB.end()) {

		std::string a;
		std::string b;

		if (itA < argsA.end()) {
			a = *itA;
		}

		if (itB < argsB.end()) {
			b = *itB;
		}

		if (iequals(a, b)) {
			++itA;
			++itB;
			//result += ".";
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

		std::string a = *itA_end;
		std::string b = *itB_end;

		if (iequals(a, b)) {
			--itA_end;
			--itB_end;
			//postResult += ".";
			continue;
		}
		else {
			break;
		}
	}

	result += "(`";
	while (itA <= itA_end) {
		result += *itA;
		if (itA < itA_end) {
			result += " ";
		}
		++itA;
	}
	result += "`, `";
	while (itB <= itB_end) {
		result += *itB;
		if (itB < itB_end) {
			result += " ";
		}
		++itB;
	}
	result += "`)";

	result += postResult;

	return result;
}

params ParseSqlParams(int argc, const char* argv[], bool quiet)
{
	params p;

	auto invalidArgs = [&](const char* msg, const char* arg = nullptr, HRESULT hr = E_INVALIDARG)
	{
		if (!quiet) {
			std::cerr << msg;

			if (arg && *arg) {
				std::cerr << " `" << arg << "`";
			}

			std::cerr << std::endl;

			showUsage();
		}

		if (msg && *msg) {
			if (!p.errorMessage.empty()) {
				p.errorMessage += "\r\n";
			}
			p.errorMessage += msg;
			if (arg && *arg) {
				p.errorMessage += " `";
				p.errorMessage += arg;
				p.errorMessage += "`";
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

	const char** arg = argv;
	const char** argEnd = arg + argc;
	const char** argFirst = arg + 1;

	const char** argVerb = nullptr;
	const char** argVerbEnd = nullptr;

	++arg;

	// mssqlpipe [instance] [as username[:password]] (backup|restore [filelistonly]) [database] dbname [ [from|to] filename|path ] ...

	// first off we will find the verb!
	{
		while (arg < argEnd) {
			auto sz = *arg;
			if (iequals(sz, "backup")) {
				argVerb = arg++;
				p.command = ToLower(sz);
				if (arg < argEnd && iequals(*arg, "database")) {
					++arg;
				}
				break;
			}
			else if (iequals(sz, "restore")) {
				argVerb = arg++;
				p.command = ToLower(sz);
				if (arg < argEnd && iequals(*arg, "database")) {
					++arg;
				}
				if (arg < argEnd && iequals(*arg, "filelistonly")) {
					p.subcommand = ToLower(*arg);
					++arg;
				}
				break;
			}
			else if (iequals(sz, "pipe")) {
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
		return invalidArgs("missing verb");
	}

	// so now we have two ranges: 
	// [argFirst...argVerb) and [argVerbEnd...argEnd)

	// start by parsing the first range, which has instance and authentication
	{
		arg = argFirst;

		// first is instance
		if (arg < argVerb && !iequals(*arg, "as")) {
			p.instance = *arg;
			++arg;
		}
		if (arg < argVerb) {
			if (iequals(*arg, "as")) {
				++arg;
				if (arg < argVerb) {
					p.as = *arg;
					++arg;

					// parse creds into username and password
					size_t split = p.as.find(L':');
					if (split == std::string::npos) {
						p.username = p.as;
					}
					else {
						p.username = p.as.substr(0, split);
						p.password = p.as.substr(split + 1);
					}
				}
				else {
					return invalidArgs("invalid authentication");
				}
			}
		}

		if (arg < argVerb) {
			return invalidArgs("invalid instance or authentication");
		}

		while (arg < argVerb) {
			auto sz = *arg;
			if (iequals(sz, "as")) {
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
			return invalidArgs("arguments missing");
		}

		// pipe can have either to or from followed by device name

		if (arg < argEnd && iequals(*arg, "to")) {
			p.subcommand = ToLower(*arg);
		}
		else if (arg < argEnd && iequals(*arg, "from")) {
			p.subcommand = ToLower(*arg);
		}
		else {
			return invalidArgs("pipe requires `to` or `from` and a device name");
		}

		++arg;

		if (arg >= argEnd) {
			return invalidArgs("pipe requires a device name");
		}

		p.device = *arg;

		++arg;

		if (arg < argEnd && p.subcommand == "from" && iequals(*arg, "to")) {
			++arg;

			if (arg >= argEnd) {
				return invalidArgs("missing file name");
			}

			p.to = *arg;
			++arg;
		}
		else if (arg < argEnd && p.subcommand == "to" && iequals(*arg, "from")) {
			++arg;

			if (arg >= argEnd) {
				return invalidArgs("missing file name");
			}

			p.from = *arg;
			++arg;
		}

		if (arg < argEnd) {
			return invalidArgs("extra args at end");
		}
	}
	else if (p.isRestore() && p.subcommand == "filelistonly") {
		// filelistonly 

		if (arg < argEnd && iequals(*arg, "from")) {
			++arg;

			if (arg >= argEnd) {
				return invalidArgs("missing file name");
			}

			p.from = *arg;
			++arg;
		}

		if (arg < argEnd) {
			return invalidArgs("extra args at end");
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
			if (arg < argEnd && iequals(*arg, "to")) {
				++arg;

				if (arg >= argEnd) {
					return invalidArgs("missing file name");
				}

				p.to = *arg;
				++arg;
			}

			// TODO with copy only, not copy only, etc?

			if (arg < argEnd && iequals(*arg, "with")) {
				++arg;

				if (arg >= argEnd) {
					return invalidArgs("missing option after with");
				}

				// no with options for backup yet

				return invalidArgs("invalid with option");
			}

			if (arg < argEnd) {
				return invalidArgs("extra args at end");
			}

		}
		else if (p.isRestore()) {

			// from file
			if (arg < argEnd && iequals(*arg, "from")) {
				++arg;

				if (arg >= argEnd) {
					return invalidArgs("missing file name");
				}

				p.from = *arg;
				++arg;
			}

			// to path
			if (arg < argEnd && iequals(*arg, "to")) {
				++arg;

				if (arg >= argEnd) {
					return invalidArgs("missing restore path");
				}

				p.to = *arg;
				++arg;
			}

			// with replace

			if (arg < argEnd && iequals(*arg, "with")) {
				++arg;

				if (arg >= argEnd) {
					return invalidArgs("missing option after with");
				}

				if (iequals(*arg, "replace")) {
					p.subcommand = ToLower(*arg);
					++arg;
				}
				else {
					return invalidArgs("invalid with option");
				}
			}
		}
		else {
			return invalidArgs("something horrible");
		}
	}

	return p;
}

params ParseParams(int argc, const char* argv[], bool quiet)
{
	// exclude all --special flags

	paramflags flags;
	std::vector<const char*> args;
	args.reserve(argc);

	args.push_back(argv[0]);
	for (int i = 1; i < argc; ++i) {
		const char* arg = argv[i];
		if (arg && arg[0] && arg[1] && arg[0] == L'-' && arg[1] == L'-') {
			if (iequals(arg, "--noelevate")) {
				flags.noelevate = true;
			} else if (iequals(arg, "--test")) {
				flags.test = true;
			} else if (iequals(arg, "--tee")) {
				++i;
				if (i < argc) {
					flags.tee = argv[i];
				}
			} else {
				args.push_back(arg);
			}
		} else {
			args.push_back(arg);
		}		
	}

	params p = ParseSqlParams(static_cast<int>(args.size()), static_cast<const char**>(&args[0]), quiet);

	p.flags = flags;
	
	return p;
}

std::string MakeParams(const params& p)
{
	std::string commandLine;
	commandLine.reserve(260);

	auto append = [&commandLine](std::string str) {
		if (str.empty()) {
			return;
		}

		if (!commandLine.empty()) {
			commandLine += " ";
		}

		commandLine += escapeCommandLine(str);
	};

	if (p.flags.noelevate) {
		append("--noelevate");
	}
	if (p.flags.test) {
		append("--test");
	}
	if (!p.flags.tee.empty()) {
		append("--tee");
		append(p.flags.tee);
	}

	if (!p.instance.empty()) {
		append(p.instance);
	}
	if (!p.as.empty()) {
		append("as");
		append(p.as);
	}

	assert(!p.command.empty());

	append(p.command);
	if (p.isRestore() && p.subcommand == "filelistonly") {
		append(p.subcommand);

		if (!p.from.empty()) {
			append("from");
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
			append("to");
			append(p.to);
		}
		if (!p.from.empty()) {
			append("from");
			append(p.from);
		}
	}
	else if (p.isBackup()) {

		assert(!p.database.empty());

		append(p.database);

		if (!p.to.empty()) {
			append("to");
			append(p.to);
		}

		assert(p.from.empty());
	}
	else if (p.isRestore()) {

		assert(!p.database.empty());

		append(p.database);

		if (!p.from.empty()) {
			append("from");
			append(p.from);
		}

		if (!p.to.empty()) {
			append("to");
			append(p.to);
		}

		if (p.subcommand == "replace") {
			append("with");
			append(p.subcommand);
		}
	}

	return commandLine;
}

std::wostream& operator<<(std::wostream& o, const params& p)
{
	o << "(";
	if (S_OK != p.hr) {
		o << "hr=" << p.hr << ";";
	}
	if (!p.instance.empty()) {
		o << "instance=" << p.instance << ";";
	}
	if (!p.command.empty()) {
		o << "command=" << p.command << ";";
	}
	if (!p.subcommand.empty()) {
		o << "subcommand=" << p.subcommand << ";";
	}
	if (!p.database.empty()) {
		o << "database=" << p.database << ";";
	}
	if (!p.from.empty()) {
		o << "from=" << p.from << ";";
	}
	if (!p.to.empty()) {
		o << "to=" << p.to << ";";
	}
	if (!p.as.empty()) {
		o << "as=" << p.as << ";";
	}
	if (!p.username.empty()) {
		o << "username=" << p.username << ";";
	}
	if (!p.password.empty()) {
		o << "password=" << p.password << ";";
	}

	if (p.isPipe()) {
		o << "device=" << p.device << ";";
	}

	o << ")";
	return o;
}

#ifdef _DEBUG
bool TestParseParams()
{
	auto test = [](const char* cmdLine) {
		int argc = 0;
		char** argv = ::CommandLineToArgvW(cmdLine, &argc);

		if (!argv) {
			return false;
		}

		auto p = ParseParams(argc, argv, true);

		::LocalFree(argv);

		auto rebuilt = "mssqlPipe " + MakeParams(p);

		bool succeeded = SUCCEEDED(p.hr);
		bool matches = iequals(cmdLine, rebuilt) || iequals("(`database`, ``)", makeCommandLineDiff(cmdLine, rebuilt));

		argc = 0;
		argv = ::CommandLineToArgvW(rebuilt.c_str(), &argc);

		if (!argv) {
			return false;
		}

		auto p_rebuilt = ParseParams(argc, argv, true);

		::LocalFree(argv);

		auto p_rebuilt_rebuilt = "mssqlPipe " + MakeParams(p_rebuilt);

		bool rebuilt_matches = iequals(rebuilt, p_rebuilt_rebuilt);

		if (!succeeded || !matches || !rebuilt_matches) {
			if (!succeeded) {
				std::cerr << "FAILED! " << p.errorMessage << " (0x" << std::hex << p.hr << std::dec << ")" << std::endl;
			}
			if (!matches) {
				std::cerr << "MISMATCH! " << makeCommandLineDiff(cmdLine, rebuilt) << std::endl;
			}
			if (!rebuilt_matches) {
				std::cerr << "REBUILT MISMATCH! " << makeCommandLineDiff(rebuilt, p_rebuilt_rebuilt) << std::endl;
			}
			std::cerr << ">\t`" << cmdLine << "`" << std::endl;
			std::cerr << "<\t`" << rebuilt << "`" << std::endl;
			std::cerr << "=\t" << p << std::endl;
			if (!rebuilt_matches) {
				std::cerr << "*<\t`" << p_rebuilt_rebuilt << "`" << std::endl;
				std::cerr << "*=\t" << p_rebuilt << std::endl;
			}
			std::cerr << std::endl;
			return false;
		}

		return true;
	};

	// arguments before verb
	if (!test("mssqlPipe backup AdventureWorks")) { return false; }
	if (!test("mssqlPipe myinstance backup AdventureWorks")) { return false; }
	if (!test("mssqlPipe myinstance as sa:hunter2 backup AdventureWorks")) { return false; }

	// backup
	if (!test("mssqlPipe backup AdventureWorks")) { return false; }
	if (!test("mssqlPipe backup database AdventureWorks")) { return false; }
	if (!test("mssqlPipe backup AdventureWorks to z:/db")) { return false; }

	// restore
	if (!test("mssqlPipe restore AdventureWorks")) { return false; }
	if (!test("mssqlPipe restore database AdventureWorks")) { return false; }
	if (!test("mssqlPipe restore AdventureWorks from z:/db/AdventureWorks.bak")) { return false; }
	if (!test("mssqlPipe restore AdventureWorks with replace")) { return false; }
	if (!test("mssqlPipe restore AdventureWorks from z:/db/AdventureWorks.bak with replace")) { return false; }

	// restore filelistonly
	if (!test("mssqlPipe restore filelistonly")) { return false; }
	if (!test("mssqlPipe restore filelistonly from z:/db/AdventureWorks.bak")) { return false; }

	// pipe
	if (!test("mssqlPipe pipe to VirtualDevice")) { return false; }
	if (!test("mssqlPipe pipe from VirtualDevice")) { return false; }
	if (!test("mssqlPipe pipe to VirtualDevice from AdventureWorks.bak")) { return false; }
	if (!test("mssqlPipe pipe from VirtualDevice to AdventureWorks.bak")) { return false; }

	return true;
}
#endif