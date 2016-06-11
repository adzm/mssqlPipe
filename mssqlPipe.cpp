// mssqlPipe.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "vdi/vdiguid.h"

///

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

///

std::mutex outputMutex_;

///

inline std::wstring escape(std::wstring str)
{
	size_t quoteCount = count(begin(str), end(str), L'\'');

	if (!quoteCount) {
		return str;
	}
	
	std::wstring q;
	q.reserve(str.size() + quoteCount);

	for (auto c : str) {		
		q.push_back(c);
		if (c == L'\'') {
			q.push_back(L'\'');
		}
	}

	return q;
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

///

inline std::wstring ToLower(std::wstring str)
{
	transform(str.begin(), str.end(), str.begin(), towlower);
	return str;
}

inline int icmp(const std::wstring& a, const std::wstring& b)
{
	return _wcsicmp(a.c_str(), b.c_str());
}
inline int icmp(const wchar_t* a, const std::wstring& b)
{
	return _wcsicmp(a, b.c_str());
}
inline int icmp(const std::wstring& a, const wchar_t* b)
{
	return _wcsicmp(a.c_str(), b);
}
inline int icmp(const wchar_t* a, const wchar_t* b)
{
	return _wcsicmp(a, b);
}

template<typename A, typename B>
bool iequals(A&& a, B&& b)
{
	return 0 == icmp(std::forward<A>(a), std::forward<B>(b));
}

struct iless_predicate
{
	template<typename A, typename B>
	bool operator()(A&& a, B&& b)
	{
		return icmp(std::forward<A>(a), std::forward<B>(b)) < 0;
	}
};

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

//std::wstring makeInlineDiff(std::wstring l, std::wstring r)
//{
//	auto charEqual = [](wchar_t a, wchar_t b) {
//		if (a == b) {
//			return true;
//		} else if (iswascii(a) && iswascii(b)) {
//			return towlower(a) == towlower(b);
//		} else {
//			return a == b;
//		}
//	};
//
//	size_t ix = 0;
//
//	// find the ix of first diff
//	for (ix = 0; ix < l.size() && ix < r.size(); ++ix) {
//		if (!charEqual(l[ix], r[ix])) {
//			break;
//		}
//	}
//
//	size_t ixFirstDiff = ix;
//
//	// find the ix of last diff
//	for (ix = 0; ix < l.size() && ix < r.size(); ++ix) {
//		if (!charEqual(l[l.size() - 1 - ix], r[r.size() - 1 - ix])) {
//			break;
//		}
//	}
//
//	size_t ixLastDiff = ix; // offset from .size()
//
//	//
//	
//	// strip equal sections from end
//	if (ixLastDiff != 0) {
//		l = l.substr(0, l.size() - 1 - ix);
//		r = r.substr(0, r.size() - 1 - ix);
//	}
//
//	if (ixFirstDiff != 0) {
//		l = l.substr(ixFirstDiff);
//		r = r.substr(ixFirstDiff);
//	}
//
//	// now we have a center section to process
//	std::wstring inlineDiff;
//	{
//
//	}
//
//	{
//		std::wstring str;
//		if (ixFirstDiff != 0) {
//			str += L"...";
//		}
//
//		str += inlineDiff;
//
//		if (ixLastDiff != 0) {
//			str += L"...";
//		}
//
//		return str;
//	}
//}

template<typename _IIID>
void ComIssueError(HRESULT hr, const _com_ptr_t<_IIID>& p)
{
	_com_issue_errorex(hr, p.GetInterfacePtr(), p.GetIID());
}

template<typename T>
void ComIssueError(HRESULT hr, const T* p)
{
	_com_issue_errorex(hr, p, __uuidof(p));
}

void ComIssueError(HRESULT hr)
{
	_com_issue_error(hr);
}

template<typename _IIID>
void ComEnsure(HRESULT hr, const _com_ptr_t<_IIID>& p)
{
	if (FAILED(hr)) {
		ComIssueError(hr, p);
	}
}

template<typename T>
void ComEnsure(HRESULT hr, const T* p)
{
	if (FAILED(hr)) {
		ComIssueError(hr, p);
	}
}


void ComEnsure(HRESULT hr)
{
	if (FAILED(hr)) {
		ComIssueError(hr);
	}
}

///

std::wstring make_guid()
{
	std::wstring guid;
	GUID guidRaw = { 0 };
	::CoCreateGuid(&guidRaw);
	guid.resize(38);
	::StringFromGUID2(guidRaw, &guid[0], guid.size() + 1);
	return guid;
}

///

struct InputFile
{
	InputFile(HANDLE hFile, size_t buflen)
		: hFile(hFile)
		, membufReserved(buflen)
	{
		if (membufReserved) {			
			membuf.reset(new BYTE[membufReserved]);
		}
	}

	void fillBuffer()
	{
		if (!membuf) {
			return;
		}
	}

	HRESULT resetPos()
	{
		if (streamPos > membufLen) {
			return E_FAIL;
		}

		membufPos = 0;
		return S_OK;
	}

	DWORD read(BYTE* buf, DWORD len)
	{
		DWORD totalRead = 0;
		if (membuf) {
			// if first read, fill the buffer
			if (membufLen == 0 && membufPos == 0 && streamPos == 0 && membufReserved != 0) {
				DWORD bytesRead = 0;

				while (bytesRead < membufReserved) {
					DWORD dwBytes = 0;
					BOOL ret = ::ReadFile(hFile, membuf.get() + bytesRead, membufReserved - bytesRead, &dwBytes, nullptr);
					bytesRead += dwBytes;
					if (!ret || (dwBytes == 0)) {
						break;
					}
				}

				streamPos += bytesRead;

				membufLen = bytesRead;
				membufPos = 0;
			}

			DWORD membufAvail = membufLen - membufPos;

			if (membufAvail) {
				DWORD bytesToRead = len;
				if (bytesToRead > membufAvail) {
					bytesToRead = membufAvail;
				}
				memcpy(buf, membuf.get() + membufPos, bytesToRead);
				membufPos += bytesToRead;
				totalRead += bytesToRead;

				len -= bytesToRead;
				buf += bytesToRead;
			}
			else {
				membuf.reset();
			}
		}

		DWORD streamRead = 0;

		while (streamRead < len) {
			DWORD dwBytes = 0;
			BOOL ret = ::ReadFile(hFile, buf, len, &dwBytes, nullptr);
			streamRead += dwBytes;
			if (!ret || (dwBytes == 0)) {
				break;
			}
		}

		streamPos += streamRead;
		totalRead += streamRead;

		return totalRead;
	}

protected:
	HANDLE hFile = nullptr;
	std::unique_ptr<BYTE[]> membuf;
	size_t membufReserved = 0;
	size_t membufLen = 0;
	size_t membufPos = 0;
	size_t streamPos = 0;
};

struct OutputFile
{
	explicit OutputFile(HANDLE hFile)
		: hFile(hFile)
	{		
	}

	DWORD write(void* buf, DWORD len)
	{
		DWORD dwBytesWritten = 0;
		while (dwBytesWritten < len) {
			DWORD dwBytes = 0;
			BOOL ret = ::WriteFile(hFile, static_cast<BYTE*>(buf) + dwBytesWritten, len - dwBytesWritten, &dwBytes, nullptr);
			dwBytesWritten += dwBytes;
			if (!ret || (dwBytes == 0)) {
				break;
			}
		}
		return dwBytesWritten;
	}

	void flush()
	{
		// not necessary with the win32 streams
		//fflush(file);
	}

protected:
	HANDLE hFile = nullptr;
};

///

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

	friend std::wostream& operator<<(std::wostream& o, const params& p)
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

///

_COM_SMARTPTR_TYPEDEF(IClientVirtualDeviceSet2, __uuidof(IClientVirtualDeviceSet2));
_COM_SMARTPTR_TYPEDEF(IClientVirtualDevice, __uuidof(IClientVirtualDevice));

struct CoInit
{
	CoInit(DWORD dwCoInit = COINIT_MULTITHREADED)
	{		
		::CoInitializeEx(NULL, dwCoInit);
	}

	~CoInit()
	{
		::CoUninitialize();
	}
};

HRESULT processPipeRestore(IClientVirtualDevice* pDevice, InputFile& file, bool quiet = false)
{
	if (!pDevice) {
		return E_INVALIDARG;
	}

	if (!quiet) {			
		std::unique_lock<std::mutex> lock(outputMutex_);
		std::wcerr << L"\nProcessing... " << std::endl;
	}

	DWORD ticksBegin = ::GetTickCount();
	DWORD ticksLastStatus = ticksBegin;

	static const DWORD timeout = 10 * 60 * 1000;
	HRESULT hr = S_OK;

	__int64 totalBytes = 0;
	DWORD totalCommands = 0;

	for (;;) {
		VDC_Command* pCmd = nullptr;		
		DWORD completionCode = 0;
		DWORD bytesTransferred = 0;

		/*If a client is using asynchronous I/O, it must ensure that a mechanism exists to complete outstanding requests when it is blocked in a 
		IClientVirtualDevice::GetCommand call. Because GetCommand waits in an Alertable state, that use of Alertable I/O is one such technique. 
		In this case, the operating system calls back to a completion routine set up by the client.*/
		hr = pDevice->GetCommand(timeout, &pCmd);

		if (!SUCCEEDED(hr)) {
			switch (hr) {
				case VD_E_CLOSE:
					// normal, we closed.
					hr = 0;
					break;
				case VD_E_TIMEOUT:
					break;
				case VD_E_ABORT:
					break;
				default:
					break;
			}
			break; // exit loop
		}

		++totalCommands;

		switch (pCmd->commandCode) {
		case VDC_Read:
			while (bytesTransferred < pCmd->size) {
				bytesTransferred += file.read(pCmd->buffer + bytesTransferred, pCmd->size - bytesTransferred);
			}

			totalBytes += bytesTransferred;

			break;
		case VDC_Write:
			completionCode = ERROR_NOT_SUPPORTED;
			break;
		case VDC_Flush:
			completionCode = ERROR_NOT_SUPPORTED;
			break;
		case VDC_ClearError:
			break;
		default:
			completionCode = ERROR_NOT_SUPPORTED;
			break;
		}

		hr = pDevice->CompleteCommand(pCmd, completionCode, bytesTransferred, 0);

		if (!SUCCEEDED(hr)) {
			// error message?
			break;
		}
				
		if (!quiet && totalBytes && (30000 < ::GetTickCount() - ticksLastStatus)) {
			ticksLastStatus = ::GetTickCount();
			DWORD ticksTotal = ticksLastStatus - ticksBegin;
			DWORD seconds = ticksTotal / 1000;
			if (!seconds) {
				seconds = 1;
			}
			__int64 bytesPerSec = (totalBytes / seconds);
			DWORD commandsPerSec = (totalCommands / seconds);
			
			std::unique_lock<std::mutex> lock(outputMutex_);
			std::wcerr << L"\nProcessing... " << totalBytes
				<< L" bytes in " << seconds << L" seconds (" << bytesPerSec << L" bytes/sec)"
				<< std::endl;
		}
	}

	if (!quiet && totalBytes) {
		DWORD ticksTotal = ::GetTickCount() - ticksBegin;
		DWORD seconds = ticksTotal / 1000;
		if (!seconds) {
			seconds = 1;
		}
		__int64 bytesPerSec = (totalBytes / seconds);
		DWORD commandsPerSec = (totalCommands / seconds);
		
		std::unique_lock<std::mutex> lock(outputMutex_);
		std::wcerr << L"\nTotal " << totalBytes
			<< L" bytes in " << seconds << L" seconds (" << bytesPerSec << L" bytes/sec) and "
			<< totalCommands << L" commands (" << commandsPerSec << L" command/sec)"
			<< std::endl;
	}
	
	return hr;
}

HRESULT processPipeBackup(IClientVirtualDevice* pDevice, OutputFile& file, bool quiet = false)
{
	if (!pDevice) {
		return E_INVALIDARG;
	}

	if (!quiet) {			
		std::unique_lock<std::mutex> lock(outputMutex_);
		std::wcerr << L"\nProcessing... " << std::endl;
	}

	DWORD ticksBegin = ::GetTickCount();
	DWORD ticksLastStatus = ticksBegin;

	static const DWORD timeout = 10 * 60 * 1000;
	HRESULT hr = S_OK;

	__int64 totalBytes = 0;
	DWORD totalCommands = 0;

	for (;;) {
		VDC_Command* pCmd = nullptr;		
		DWORD completionCode = 0;
		DWORD bytesTransferred = 0;

		/*If a client is using asynchronous I/O, it must ensure that a mechanism exists to complete outstanding requests when it is blocked in a 
		IClientVirtualDevice::GetCommand call. Because GetCommand waits in an Alertable state, that use of Alertable I/O is one such technique. 
		In this case, the operating system calls back to a completion routine set up by the client.*/
		hr = pDevice->GetCommand(timeout, &pCmd);

		if (!SUCCEEDED(hr)) {
			switch (hr) {
				case VD_E_CLOSE:
					// normal, we closed.
					hr = 0;
					break;
				case VD_E_TIMEOUT:
					break;
				case VD_E_ABORT:
					break;
				default:
					break;
			}
			break; // exit loop
		}

		++totalCommands;

		switch (pCmd->commandCode) {
		case VDC_Read:
			completionCode = ERROR_NOT_SUPPORTED;
			break;
		case VDC_Write:
			while (bytesTransferred < pCmd->size) {
				bytesTransferred += file.write(pCmd->buffer + bytesTransferred, pCmd->size - bytesTransferred);
			}

			totalBytes += bytesTransferred;

			break;
		case VDC_Flush:
			file.flush();
			break;
		case VDC_ClearError:
			break;
		default:
			completionCode = ERROR_NOT_SUPPORTED;
			break;
		}

		hr = pDevice->CompleteCommand(pCmd, completionCode, bytesTransferred, 0);

		if (!SUCCEEDED(hr)) {
			// error message?
			break;
		}
				
		if (!quiet && totalBytes && (30000 < ::GetTickCount() - ticksLastStatus)) {
			ticksLastStatus = ::GetTickCount();
			DWORD ticksTotal = ticksLastStatus - ticksBegin;
			DWORD seconds = ticksTotal / 1000;
			if (!seconds) {
				seconds = 1;
			}
			__int64 bytesPerSec = (totalBytes / seconds);
			DWORD commandsPerSec = (totalCommands / seconds);
			
			std::unique_lock<std::mutex> lock(outputMutex_);
			std::wcerr << L"\nProcessing... " << totalBytes
				<< L" bytes in " << seconds << L" seconds (" << bytesPerSec << L" bytes/sec)"
				<< std::endl;
		}
	}

	if (!quiet && totalBytes) {
		DWORD ticksTotal = ::GetTickCount() - ticksBegin;
		DWORD seconds = ticksTotal / 1000;
		if (!seconds) {
			seconds = 1;
		}
		__int64 bytesPerSec = (totalBytes / seconds);
		DWORD commandsPerSec = (totalCommands / seconds);

		std::unique_lock<std::mutex> lock(outputMutex_);
		std::wcerr << L"\nTotal " << totalBytes
			<< L" bytes in " << seconds << L" seconds (" << bytesPerSec << L" bytes/sec) and "
			<< totalCommands << L" commands (" << commandsPerSec << L" command/sec)"
			<< std::endl;
	}
	
	return hr;
}

struct VirtualDevice
{
	IClientVirtualDeviceSet2Ptr pSet;
	IClientVirtualDevicePtr pDevice;
		
	VDConfig config;
	
	std::wstring instance;
	std::wstring name;

	VirtualDevice(std::wstring instance, std::wstring name)
		: instance(std::move(instance))
		, name(std::move(name))
		, config({ 1 })
	{}

	~VirtualDevice()
	{
		Close();
	}

	HRESULT Close()
	{
		HRESULT hr = 0;
		if (pSet) {			
			hr = pSet->Close();
		}
		pDevice = nullptr;
		pSet = nullptr;
		return hr;
	}

	HRESULT Abort()
	{
		if (!pSet) {
			return S_FALSE;
		}
		return pSet->SignalAbort();
	}
		
	HRESULT Create()
	{
		HRESULT hr = S_OK;
	
		hr = pSet.CreateInstance(CLSID_MSSQL_ClientVirtualDeviceSet);
		if (!SUCCEEDED(hr)) {
			std::unique_lock<std::mutex> lock(outputMutex_);
			std::wcerr << L"Failed to cocreate device set: " << std::hex << hr << std::dec << std::endl;
			return hr;
		}

		const wchar_t* wInstance = instance.empty() ? nullptr : (const wchar_t*)instance.c_str();

		hr = pSet->CreateEx(wInstance, name.c_str(), &config);
		if (!SUCCEEDED(hr)) {
			std::unique_lock<std::mutex> lock(outputMutex_);
			std::wcerr << L"Failed to create device set: " << std::hex << hr << std::dec << std::endl;
			return hr;
		}
		else {			
		}

		return hr;
	}

	HRESULT Open(DWORD dwTimeout)
	{
		HRESULT hr = 0;

		hr = pSet->GetConfiguration(dwTimeout, &config);
		if (!SUCCEEDED(hr)) {
			pSet->Close();
			std::unique_lock<std::mutex> lock(outputMutex_);
			std::wcerr << L"Failed to initialize backup operation: " << std::hex << hr << std::dec << std::endl;
			return hr;
		}

		hr = pSet->OpenDevice(name.c_str(), &pDevice);
		if (!SUCCEEDED(hr)) {
			pSet->Close();
			std::unique_lock<std::mutex> lock(outputMutex_);
			std::wcerr << L"Failed to open backup device: " << std::hex << hr << std::dec << std::endl;
			return hr;
		}

		return hr;
	}
};

std::wstring EscapeConnectionStringValue(const std::wstring& val)
{
	auto pos = val.begin();

	bool hasSpace = false;
	int singleCount = 0;
	int doubleCount = 0;
	int otherCount = 0;

	if (!val.empty()) {
		if ((::iswspace(val.front())) || (::iswspace(val.back()))) {
			hasSpace = true;
		}
	}
	while (pos < val.end()) {
		switch (*pos++) {
		case L';':
			++otherCount;
			break;
		case L'\'':
			++singleCount;
			break;
		case L'\"':
			++doubleCount;
			break;
		}
	}

	int totalCount = otherCount + singleCount + doubleCount;

	if (!totalCount && !hasSpace) {
		return val;
	}

	std::wstring esc;
	esc.reserve(val.length() + 2 + (totalCount * 2));

	wchar_t c = L'\'';
	if (singleCount && !doubleCount) {
		c = L'\"';
	}
	esc.push_back(c);

	pos = val.begin();
	while (pos < val.end()) {
		esc.push_back(*pos);
		if (*pos == c) {
			esc.push_back(c);
		}
		++pos;
	}

	esc.push_back(c);

	return esc;
}

std::wstring MakeConnectionString(const std::wstring& instance, const std::wstring& username, const std::wstring& password)
{
	std::wstring connectionString;

	connectionString.reserve(256);
		
	// SQLNCLI might be a better option, but SQLOLEDB is ubiquitous. 
	connectionString = L"Provider=SQLOLEDB;Initial Catalog=master;";
	
	if (username.empty()) {
		connectionString += L"Integrated Security=SSPI;";
	}
	else {
		connectionString += L"User ID=";
		connectionString += EscapeConnectionStringValue(username);
		connectionString += L";Password=";
		connectionString += EscapeConnectionStringValue(password);
		connectionString += L";";
	}

	connectionString += L"Data Source=";

	std::wstring dataSource = L"lpc:.";
	if (!instance.empty()) {
		dataSource += L"\\";
		dataSource += instance;
	}

	connectionString += EscapeConnectionStringValue(dataSource);
	connectionString += L";";

	return connectionString;
}

ADODB::_ConnectionPtr Connect(std::wstring connectionString)
{
	try {
		ADODB::_ConnectionPtr pCon(__uuidof(ADODB::Connection));

		pCon->Open(connectionString.c_str(), L"", L"", ADODB::adConnectUnspecified);

		pCon->CommandTimeout = 0;

		pCon->CursorLocation = ADODB::adUseServer;

		return pCon;
	}
	catch (_com_error& e) {
		std::unique_lock<std::mutex> lock(outputMutex_);
		std::wcerr << L"Could not connect to " << connectionString << std::endl;
		std::wcerr << std::hex << e.Error() << std::dec << L": " << e.ErrorMessage() << std::endl;
		std::wcerr << e.Description() << std::endl;

		return nullptr;
	}
}

inline std::wostream& operator<<(std::wostream& o, const _bstr_t& bstr)
{
	const wchar_t* str = (const wchar_t*)bstr;
	if (str) {
		return o << str;
	}
	else {
		return o;
	}
}

inline std::ostream& operator<<(std::ostream& o, const _bstr_t& bstr)
{
	const char* str = (const char*)bstr;
	if (str) {
		return o << str;
	}
	else {
		return o;
	}
}

void traceAdoErrors(ADODB::_Connection* pCon) 
{
	auto errors = pCon->Errors;
	if (!errors) {
		return;
	}

	long count = errors->Count;
	if (!count) {
		return;
	}

	{
		std::unique_lock<std::mutex> lock(outputMutex_);
		for (long i = 0; i < count; ++i) {
			auto error = errors->Item[i];
			if (!error) {
				continue;
			}
			std::wcerr
				<< error->Description
				<< L"\t" << std::hex << error->Number << std::dec
				<< L"\t" << error->NativeError
				<< L"\t" << error->SQLState
				<< std::endl;
			//std::wcerr 
			//	<< L"[" << i << L"]"
			//	<< L"\t" << std::hex << error->Number << std::dec
			//	<< L"\t" << error->NativeError
			//	<< L"\t" << error->SQLState
			//	<< L"\t" << error->Description
			//	<< L"\t" << error->Source
			//	<< std::endl;
		}
	}

	errors->Clear();
}

void traceAdoRecordset(ADODB::_Recordset* pRs)
{
	if (!pRs->State || pRs->eof) {
		return;
	}

	auto fields = pRs->Fields;
	if (!fields) {
		return;
	}
	long count = fields->Count;

	std::unique_lock<std::mutex> lock(outputMutex_);
	for (long x = 0; x < count; ++x) {
		auto field = fields->Item[x];
		if (!field) {
			continue;
		}
		auto name = field->Name;
		if (!name.length()) {
			std::wcerr << L"\t[" << x << L"]";
		}
		else {
			std::wcerr << L"\t[" << field->Name << L"]";
		}
	}
	std::wcerr << std::endl;

	while (!pRs->eof) {
		for (long x = 0; x < count; ++x) {
			_variant_t var = pRs->Collect[x];
			HRESULT hrChange = 0;
			if (var.vt & VT_ARRAY) {
				hrChange = E_FAIL;
			}
			else {
				hrChange = ::VariantChangeType(&var, &var, 0, VT_BSTR);
			}

			if (SUCCEEDED(hrChange) && var.vt == VT_BSTR) {
				std::wcerr << L"\t" << var.bstrVal;
			}
			else if (var.vt == VT_NULL) {
				std::wcerr << L"\t";
			}
			else {
				std::wcerr << L"\t?";
			}
		}
		std::wcerr << std::endl;
		pRs->MoveNext();
	}
}

struct DbFile
{
	std::wstring logicalName;
	std::wstring physicalName;
	std::wstring type;
};

HRESULT RunPrepareRestoreDatabase(params p, InputFile& inputFile, std::wstring& dataPath, std::wstring& logPath, std::vector<DbFile>& fileList, bool quiet)
{
	HRESULT hr = 0;

	std::wstring connectionString = MakeConnectionString(p.instance, p.username, p.password);

	std::wstring sql;
	{
		std::wostringstream o;
		o << L"restore filelistonly from virtual_device=N'" << escape(p.device) << L"';";
		sql = o.str();
	}

	VirtualDevice vd(p.instance, p.device);
	hr = vd.Create();
	if (!SUCCEEDED(hr)) {
		return hr;
	}

	auto pipeResult = std::async([&vd, &inputFile, &p, quiet]{
		CoInit comInit;

		vd.Open(p.timeout);
		return processPipeRestore(vd.pDevice, inputFile, quiet);
	});

	auto adoResult = std::async([&connectionString, &sql, &fileList]{
		CoInit comInit;
		
		try {
			ADODB::_ConnectionPtr pCon = Connect(connectionString);
			if (!pCon) {
				return E_FAIL;
			}

			ADODB::_RecordsetPtr pRs(__uuidof(ADODB::Recordset));
			pRs->CursorLocation = ADODB::adUseServer;
			pRs->Open(sql.c_str(), (IDispatch*)pCon, ADODB::adOpenForwardOnly, ADODB::adLockReadOnly, ADODB::adCmdText);

			traceAdoErrors(pCon);

			for (; pRs && !pRs->eof; pRs->MoveNext()) {
				DbFile f;
				f.logicalName = pRs->Fields->Item[L"LogicalName"]->Value.bstrVal;
				f.physicalName = pRs->Fields->Item[L"PhysicalName"]->Value.bstrVal;
				f.type = pRs->Fields->Item[L"Type"]->Value.bstrVal;

				fileList.push_back(f);
			}
			traceAdoErrors(pCon);

			pCon->Close();

			return S_OK;
		}
		catch (_com_error& e) {			
			std::unique_lock<std::mutex> lock(outputMutex_);
			std::wcerr << std::hex << e.Error() << std::dec << L": " << e.ErrorMessage() << std::endl;
			std::wcerr << e.Description() << std::endl;
			return e.Error();
		}
	});
	
	auto adoPathsResult = std::async([&connectionString, &dataPath, &logPath]{
		CoInit comInit;
		
		try {
			ADODB::_ConnectionPtr pCon = Connect(connectionString);
			if (!pCon) {
				return E_FAIL;
			}
			
			// InstanceDefaultDataPath and InstanceDefaultLogPath are SQL2012+
			auto query = LR"(
;with database_info as (
select 
	substring(physical_name, 1, len(physical_name) - charindex('\', reverse(physical_name))) as physical_path
	, case when database_id in (db_id('master'), db_id('msdb'), db_id('tempdb'), db_id('model')) then 1 else 0 end as is_system
	, type
from sys.master_files
)
select 
	coalesce( convert(nvarchar(512), serverproperty('InstanceDefaultDataPath')), (
		select top 1 physical_path from database_info
		where type = 0
		group by physical_path, is_system
		order by is_system, count(*), physical_path
	)) as DefaultData
	, 
	coalesce( convert(nvarchar(512), serverproperty('InstanceDefaultLogPath')), (
		select top 1 physical_path from database_info
		where type = 1
		group by physical_path, is_system
		order by is_system, count(*), physical_path
	)) as DefaultLog
;
)";

			ADODB::_RecordsetPtr pRs(__uuidof(ADODB::Recordset));
			pRs->CursorLocation = ADODB::adUseServer;
			pRs->Open(query, (IDispatch*)pCon, ADODB::adOpenForwardOnly, ADODB::adLockReadOnly, ADODB::adCmdText);

			traceAdoErrors(pCon);

			if (!pRs->eof) {
				
				dataPath = pRs->Fields->Item[L"DefaultData"]->Value.bstrVal;
				logPath = pRs->Fields->Item[L"DefaultLog"]->Value.bstrVal;

				traceAdoErrors(pCon);
			}

			pCon->Close();

			return S_OK;
		}
		catch (_com_error& e) {			
			std::unique_lock<std::mutex> lock(outputMutex_);
			std::wcerr << std::hex << e.Error() << std::dec << L": " << e.ErrorMessage() << std::endl;
			std::wcerr << e.Description() << std::endl;
			return e.Error();
		}
	});

	HRESULT hrAdoPaths = adoPathsResult.get();

	if (!SUCCEEDED(hrAdoPaths)) {
		if (!hr) {
			hr = hrAdoPaths;
		}
		vd.Abort();
	}

	HRESULT hrAdo = adoResult.get();

	if (!SUCCEEDED(hrAdo)) {
		if (!hr) {
			hr = hrAdo;
		}
		vd.Abort();
	}

	HRESULT hrPipe = pipeResult.get();
	if (!SUCCEEDED(hrPipe)) {
		if (!hr) {
			hr = hrAdo;
		}
	}

	return hr;
}

HRESULT RunRestoreDatabase(params p, std::wstring sql, InputFile& inputFile, bool quiet)
{
	HRESULT hr = 0;

	std::wstring connectionString = MakeConnectionString(p.instance, p.username, p.password);

	if (!quiet) {
		std::unique_lock<std::mutex> lock(outputMutex_);
		std::wcerr << L"Restoring via virtual device " << p.device << std::endl;
	}

	VirtualDevice vd(p.instance, p.device);
	hr = vd.Create();
	if (!SUCCEEDED(hr)) {
		return hr;
	}

	auto pipeResult = std::async([&vd, &inputFile, &p, quiet]{
		CoInit comInit;

		vd.Open(p.timeout);
		return processPipeRestore(vd.pDevice, inputFile, quiet);
	});

	auto adoResult = std::async([&connectionString, &sql]{
		CoInit comInit;
		
		try {
			ADODB::_ConnectionPtr pCon = Connect(connectionString);
			if (!pCon) {
				return E_FAIL;
			}

			ADODB::_RecordsetPtr pRs(__uuidof(ADODB::Recordset));
			pRs->CursorLocation = ADODB::adUseServer;
			pRs->Open(sql.c_str(), (IDispatch*)pCon, ADODB::adOpenForwardOnly, ADODB::adLockReadOnly, ADODB::adCmdText);

			traceAdoErrors(pCon);

			while (pRs) {
				traceAdoRecordset(pRs);

				traceAdoErrors(pCon);
				_variant_t varAffected;
				pRs = pRs->NextRecordset(&varAffected);

				traceAdoErrors(pCon);
			}

			pCon->Close();

			return S_OK;
		}
		catch (_com_error& e) {			
			std::unique_lock<std::mutex> lock(outputMutex_);
			std::wcerr << std::hex << e.Error() << std::dec << L": " << e.ErrorMessage() << std::endl;
			std::wcerr << e.Description() << std::endl;
			return e.Error();
		}
	});

	HRESULT hrAdo = adoResult.get();

	if (!SUCCEEDED(hrAdo)) {
		if (!hr) {
			hr = hrAdo;
		}
		vd.Abort();
	}

	HRESULT hrPipe = pipeResult.get();
	if (!SUCCEEDED(hrPipe)) {
		if (!hr) {
			hr = hrAdo;
		}
	}

	return hr;
}

std::wstring BuildRestoreCommand(params p, std::wstring dataPath, std::wstring logPath, const std::vector<DbFile>& fileList)
{
	if (!p.to.empty()) {
		dataPath = p.to;
		logPath = p.to;
	}

	if (!dataPath.empty() && dataPath.back() != L'\\') {
		dataPath += L'\\';
	}
	if (!logPath.empty() && logPath.back() != L'\\') {
		logPath += L'\\';
	}

	std::wostringstream o;
	o << L"restore database [" << escape(p.database) << L"] from virtual_device=N'" << escape(p.device) << L"' with ";
	
	if (iequals(p.subcommand, L"replace")) {
		o << L"replace, ";
	}

	// build moves?
	if (!fileList.empty()) {
		std::set<std::wstring, iless_predicate> fileRoots;
		for (auto file : fileList) {
			std::wstring targetBase;
			if (iequals(file.type, L"D")) {
				targetBase = dataPath + p.database + L"_dat";
			}
			else {
				targetBase = logPath + p.database + L"_log";
			}
			
			std::wstring target = targetBase;
			size_t ordinal = 1;
			while (fileRoots.count(target)) {
				++ordinal;
				std::wostringstream newTarget;
				newTarget << target << ordinal;
				target = newTarget.str();
			}
			fileRoots.insert(target);
						
			if (iequals(file.type, L"D")) {
				target += L".mdf";
			}
			else {
				target += L".ldf";
			}

			o << L"move N'" << escape(file.logicalName) << L"' to N'" << escape(target) << L"', ";
		}
	}
		
	o << L"nounload"; // noop basically, so i dont have to deal with trailing comma

	o << L";";
	
	return o.str();
}

HRESULT RunRestore(params p, HANDLE hFile)
{
	HRESULT hr = S_OK;

	InputFile inputFile(hFile, 0x10000);

	if (iequals(p.subcommand, L"filelistonly")) {
		
		std::wostringstream o;
		o << L"restore filelistonly from virtual_device=N'" << escape(p.device) << L"';";

		hr = RunRestoreDatabase(p, o.str(), inputFile, true);
	
		if (!SUCCEEDED(hr)) {
			std::unique_lock<std::mutex> lock(outputMutex_);
			std::wcerr << L"RunRestoreDatabase generic failed with " << std::hex << hr << std::dec << std::endl;
			return hr;
		}

		return hr;
	}

	std::vector<DbFile> fileList;
	std::wstring dataPath;
	std::wstring logPath;

	{
		params altp = p;
		altp.device = make_guid();

		hr = RunPrepareRestoreDatabase(altp, inputFile, dataPath, logPath, fileList, true);
		if (!SUCCEEDED(hr)) {
			std::unique_lock<std::mutex> lock(outputMutex_);
			std::wcerr << L"RunRestoreFileListOnly failed with " << std::hex << hr << std::dec << std::endl;
			return hr;
		}
	}

	hr = inputFile.resetPos();
	if (!SUCCEEDED(hr)) {
		std::unique_lock<std::mutex> lock(outputMutex_);
		std::wcerr << L"RunRestore failed to reset buffered input pos with " << std::hex << hr << std::dec << std::endl;
		return hr;
	}
	
	auto sql = BuildRestoreCommand(p, dataPath, logPath, fileList);

	hr = RunRestoreDatabase(p, sql, inputFile, false);
	
	if (!SUCCEEDED(hr)) {
		std::unique_lock<std::mutex> lock(outputMutex_);
		std::wcerr << L"RunRestoreDatabase failed with " << std::hex << hr << std::dec << std::endl;
		return hr;
	}

	return hr;
}

HRESULT RunBackup(params p, HANDLE hFile)
{
	HRESULT hr = 0;

	std::wstring connectionString = MakeConnectionString(p.instance, p.username, p.password);

	std::wstring sql;
	{
		std::wostringstream o;
		// always do copy_only, could be an option in the future
		o << L"backup database [" << escape(p.database) << L"] to virtual_device=N'" << escape(p.device) << L"' with copy_only;";
		sql = o.str();
	}

	{
		std::unique_lock<std::mutex> lock(outputMutex_);
		std::wcerr << L"Backing up via virtual device " << p.device << std::endl;
	}

	VirtualDevice vd(p.instance, p.device);
	hr = vd.Create();
	if (!SUCCEEDED(hr)) {
		return hr;
	}

	OutputFile outputFile(hFile);

	auto pipeResult = std::async([&vd, &outputFile, &p]{
		CoInit comInit;

		vd.Open(p.timeout);
		return processPipeBackup(vd.pDevice, outputFile);
	});

	auto adoResult = std::async([&connectionString, &sql]{
		CoInit comInit;
		
		try {
			ADODB::_ConnectionPtr pCon = Connect(connectionString);
			if (!pCon) {
				return E_FAIL;
			}

			ADODB::_RecordsetPtr pRs(__uuidof(ADODB::Recordset));
			pRs->CursorLocation = ADODB::adUseServer;
			pRs->Open(sql.c_str(), (IDispatch*)pCon, ADODB::adOpenForwardOnly, ADODB::adLockReadOnly, ADODB::adCmdText);

			traceAdoErrors(pCon);

			while (pRs) {
				traceAdoRecordset(pRs);

				traceAdoErrors(pCon);
				_variant_t varAffected;
				pRs = pRs->NextRecordset(&varAffected);

				traceAdoErrors(pCon);
			}

			pCon->Close();

			return S_OK;
		}
		catch (_com_error& e) {			
			std::unique_lock<std::mutex> lock(outputMutex_);
			std::wcerr << std::hex << e.Error() << std::dec << L": " << e.ErrorMessage() << std::endl;
			std::wcerr << e.Description() << std::endl;
			return e.Error();
		}
	});

	HRESULT hrAdo = adoResult.get();

	if (!SUCCEEDED(hrAdo)) {
		if (!hr) {
			hr = hrAdo;
		}
		vd.Abort();
	}

	HRESULT hrPipe = pipeResult.get();
	if (!SUCCEEDED(hrPipe)) {
		if (!hr) {
			hr = hrAdo;
		}
	}

	return hr;
}

HRESULT RunPipe(params p, HANDLE hFile)
{
	HRESULT hr = 0;

	{
		std::unique_lock<std::mutex> lock(outputMutex_);
		std::wcerr << L"Piping " << p.subcommand << L" virtual device " << p.device << std::endl;
	}

	VirtualDevice vd(p.instance, p.device);
	hr = vd.Create();
	if (!SUCCEEDED(hr)) {
		return hr;
	}

	// in pipe mode, use a much larger than the default timeout, say 5 minutes
	DWORD pipeTimeout = 5 * 60 * 1000;

	if (iequals(p.subcommand, L"from")) {
		OutputFile outputFile(hFile);

		auto pipeResult = std::async([&vd, &outputFile, pipeTimeout]{
			CoInit comInit;

			vd.Open(pipeTimeout);
			return processPipeBackup(vd.pDevice, outputFile);
		});

		hr = pipeResult.get();
	}
	else if (iequals(p.subcommand, L"to")) {

		InputFile inputFile(hFile, 0); // no buffering

		auto pipeResult = std::async([&vd, &inputFile, pipeTimeout]{
			CoInit comInit;

			vd.Open(pipeTimeout);
			return processPipeRestore(vd.pDevice, inputFile);
		});

		hr = pipeResult.get();
	}

	return hr;
}		

HRESULT Run(params p)
{
	HANDLE hFile = nullptr;

	HANDLE hStdIn = ::GetStdHandle(STD_INPUT_HANDLE);
	HANDLE hStdOut = ::GetStdHandle(STD_OUTPUT_HANDLE);
	HANDLE hStdErr = ::GetStdHandle(STD_ERROR_HANDLE);

	if (p.isBackup()) {
		if (p.to.empty()) {
			hFile = hStdOut;
		}
		else {
			hFile = ::CreateFile(p.to.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
			if (!hFile || INVALID_HANDLE_VALUE == hFile) {
				DWORD ret = ::GetLastError();
				std::wcerr << ret << L": Failed to open " << p.to << std::endl;
				return E_FAIL;
			}
		}
	}
	else if (p.isRestore()) {
		if (!p.to.empty() && INVALID_FILE_ATTRIBUTES == ::GetFileAttributes(p.to.c_str())) {
			int ret = ::SHCreateDirectoryEx(nullptr, p.to.c_str(), nullptr);
			if (ret && ret != ERROR_FILE_EXISTS && ret != ERROR_ALREADY_EXISTS) {
				std::wcerr << ret << L": Failed to create restore to path. Continuing (SQL Server may have access)... " << p.from << std::endl;
			}
		}
		if (p.from.empty()) {
			hFile = hStdIn;
		}
		else {
			hFile = ::CreateFile(p.from.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
			if (!hFile || INVALID_HANDLE_VALUE == hFile) {
				DWORD ret = ::GetLastError();
				std::wcerr << ret << L": Failed to open " << p.from << std::endl;
				return E_FAIL;
			}
		}
	}
	else if (p.isPipe()) {
		if (iequals(p.subcommand, L"to")) {
			if (p.from.empty()) {
				hFile = hStdIn;
			}
			else {
				hFile = ::CreateFile(p.from.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
				if (!hFile || INVALID_HANDLE_VALUE == hFile) {
					DWORD ret = ::GetLastError();
					std::wcerr << ret << L": Failed to open " << p.from << std::endl;
					return E_FAIL;
				}
			}
		}
		else if (iequals(p.subcommand, L"from")) {
			if (p.to.empty()) {
				hFile = hStdOut;
			}
			else {
				hFile = ::CreateFile(p.to.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
				if (!hFile || INVALID_HANDLE_VALUE == hFile) {
					DWORD ret = ::GetLastError();
					std::wcerr << ret << L": Failed to open " << p.to << std::endl;
					return E_FAIL;
				}
			}
		}
	}

	if (!hFile) {
		std::wcerr << L"missing file" << std::endl;
		return E_FAIL;
	}

	HRESULT hr = S_OK;
	
	if (p.isBackup()) {
		hr = RunBackup(p, hFile);
	}
	else if (p.isRestore()) {
		hr = RunRestore(p, hFile);
	}
	else if (p.isPipe()) {
		hr = RunPipe(p, hFile);
	}
	else {		
		std::wcerr << L"unexpected command " << p.command << std::endl;
		hr = E_FAIL;
	}
	
	if (hFile != hStdOut && hFile != hStdIn && hFile != hStdErr) {
		::CloseHandle(hFile);
		hFile = nullptr;
	}

	return hr;
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
		} else if (arg < argEnd && p.subcommand == L"to" && iequals(*arg, L"from")) {
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
	} else if (p.isPipe()) {

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

		if (!succeeded || !matches) {
			if (!succeeded) {
				std::wcerr << L"FAILED! " << p.errorMessage << L" (0x" << std::hex << p.hr << std::dec << L")" << std::endl;
			}
			if (!matches) {
				std::wcerr << L"MISMATCH! " << makeCommandLineDiff(cmdLine, rebuilt) << std::endl;
			}
			std::wcerr << L">\t`" << cmdLine << L"`" << std::endl;
			std::wcerr << L"<\t`" << rebuilt << L"`" << std::endl;
			std::wcerr << L"=\t" << p << std::endl;
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

int wmain(int argc, wchar_t* argv[])
{
	CoInit comInit;

#ifdef _DEBUG
	bool waitForDebugger = true;
	waitForDebugger = false;
	if (waitForDebugger && !::IsDebuggerPresent()) {
		std::wcerr << L"Waiting for debugger..." << std::endl;
		::Sleep(10000);
	}
#endif
#ifdef _DEBUG
	assert(TestParseParams());
#endif
		
	std::wcerr << L"\nmssqlPipe v1.0.1\n" << std::endl;

	params p = ParseParams(argc, argv, false);
	
	if (SUCCEEDED(p.hr)) {
		p.hr = Run(p);
	}

	if (E_ACCESSDENIED == p.hr) {
		std::wcerr << L"Run failed with E_ACCESSDENIED; mssqlPipe may need to run as an administrator." << std::endl;
	}

#ifdef _DEBUG
	::Sleep(5000);
#endif

	return p.hr;
}

