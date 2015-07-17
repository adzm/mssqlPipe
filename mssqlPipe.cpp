// mssqlPipe.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "vdi/vdiguid.h"

///

void showUsage()
{
	std::wcerr << std::endl;
	std::wcerr << L"Usage:" << std::endl;
	std::wcerr << std::endl;
	std::wcerr << L"mssqlPipe [instance] backup [database] dbname\n\t[to filename]" << std::endl;
	std::wcerr << L"mssqlPipe [instance] restore [database] dbname\n\t[from filename] [to filepath] [with replace]" << std::endl;
	std::wcerr << L"mssqlPipe [instance] restore filelistonly\n\t[from filename]" << std::endl;
	std::wcerr << L"mssqlPipe [instance] pipe\n\tto devicename\n\tfrom devicename" << std::endl;
	std::wcerr << std::endl;
	std::wcerr << L"stdin or stdout will be used if no filenames specified." << std::endl;
	std::wcerr << std::endl;
	std::wcerr << L"Examples:" << std::endl;
	std::wcerr << std::endl;
	std::wcerr << L"mssqlPipe myinstance backup AdventureWorks to AdventureWorks.bak" << std::endl;
	std::wcerr << L"mssqlPipe myinstance backup AdventureWorks > AdventureWorks.bak" << std::endl;
	std::wcerr << L"mssqlPipe backup database AdventureWorks | 7za a AdventureWorks.xz -si" << std::endl;
	std::wcerr << L"7za e AdventureWorks.xz -so | mssqlPipe restore AdventureWorks to c:\\db\\" << std::endl;
	std::wcerr << L"mssqlPipe restore AdventureWorks from AdventureWorks.bak with replace" << std::endl;
	std::wcerr << L"mssqlPipe pipe from VirtualDevice42 > output.bak" << std::endl;
	std::wcerr << L"mssqlPipe pipe to VirtualDevice42 < input.bak" << std::endl;
	std::wcerr << std::endl;
	std::wcerr << L"Happy piping!" << std::endl;
	std::wcerr << std::endl;
}

///

std::mutex outputMutex_;

///

inline std::wstring escape(const std::wstring& str)
{
	size_t quoteCount = count(begin(str), end(str), L'\'');
	
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

HRESULT invalidArgs(const wchar_t* msg, const wchar_t* arg = nullptr, HRESULT hr = E_INVALIDARG)
{
	std::wcerr << msg;

	if (arg && *arg) {
		std::wcerr << L" `" << arg << L"`";
	} 

	std::wcerr << std::endl;

	showUsage();

	return hr;
};

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
	InputFile(FILE* f, size_t buflen)
		: file(f)
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
				size_t bytesRead = fread(membuf.get(), 1, membufReserved, file);

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
		DWORD streamRead = fread(buf, 1, len, file);
		streamPos += streamRead;
		totalRead += streamRead;

		return totalRead;
	}

protected:
	FILE* file = nullptr;
	std::unique_ptr<BYTE[]> membuf;
	size_t membufReserved = 0;
	size_t membufLen = 0;
	size_t membufPos = 0;
	size_t streamPos = 0;
};

struct OutputFile
{
	explicit OutputFile(FILE* f)
		: file(f)
	{		
	}

	DWORD write(void* buf, DWORD len)
	{
		return fwrite(buf, 1, len, file);
	}

	void flush()
	{
		fflush(file);
	}

protected:
	FILE* file = nullptr;
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
			std::wcerr << L"Note mssqlPipe must be run as administrator!" << std::endl;
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

std::wstring MakeConnectionString(std::wstring instance)
{
	std::wstring connectionString;

	connectionString.reserve(128);
		
	// SQLNCLI might be a better option, but SQLOLEDB is ubiquitous. 
	connectionString = L"Provider=SQLOLEDB;Initial Catalog=master;Integrated Security=SSPI;Data Source=lpc:.";

	if (!instance.empty()) {
		connectionString += L"\\";
		connectionString += instance;
	}
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
		std::wcerr << e.Error() << L": " << e.ErrorMessage() << std::endl;
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
			std::wcerr << L"[" << i << L"]"
				<< L"\t" << std::hex << error->Number << std::dec
				<< L"\t" << error->NativeError
				<< L"\t" << error->SQLState
				<< L"\t" << error->Description
				<< L"\t" << error->Source
				<< std::endl;
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

	std::wstring connectionString = MakeConnectionString(p.instance);

	std::wstring sql;
	{
		std::wostringstream o;
		o << L"restore filelistonly from virtual_device=N'" << escape(p.device) << L"';";
		sql = o.str();
	}

	VirtualDevice vd(p.instance, p.device);
	hr = vd.Create();

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
			std::wcerr << e.Error() << L": " << e.ErrorMessage() << std::endl;
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
			std::wcerr << e.Error() << L": " << e.ErrorMessage() << std::endl;
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

	std::wstring connectionString = MakeConnectionString(p.instance);

	VirtualDevice vd(p.instance, p.device);
	hr = vd.Create();
	if (!quiet) {
		std::unique_lock<std::mutex> lock(outputMutex_);
		std::wcerr << L"Restoring via virtual device " << p.device << std::endl;
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
			std::wcerr << e.Error() << L": " << e.ErrorMessage() << std::endl;
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

HRESULT RunRestore(params p, FILE* file)
{
	HRESULT hr = S_OK;

	InputFile inputFile(file, 0x10000);

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

HRESULT RunBackup(params p, FILE* file)
{
	HRESULT hr = 0;

	std::wstring connectionString = MakeConnectionString(p.instance);

	std::wstring sql;
	{
		std::wostringstream o;
		// always do copy_only, could be an option in the future
		o << L"backup database [" << escape(p.database) << L"] to virtual_device=N'" << escape(p.device) << L"' with copy_only;";
		sql = o.str();
	}

	VirtualDevice vd(p.instance, p.device);
	hr = vd.Create();
	{
		std::unique_lock<std::mutex> lock(outputMutex_);
		std::wcerr << L"Backing up via virtual device " << p.device << std::endl;
	}

	OutputFile outputFile(file);

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
			std::wcerr << e.Error() << L": " << e.ErrorMessage() << std::endl;
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

HRESULT RunPipe(params p, FILE* file)
{
	HRESULT hr = 0;

	VirtualDevice vd(p.instance, p.device);
	hr = vd.Create();
	{
		std::unique_lock<std::mutex> lock(outputMutex_);
		std::wcerr << L"Piping " << p.subcommand << L" virtual device " << p.device << std::endl;
	}

	// in pipe mode, use a much larger than the default timeout, say 5 minutes
	DWORD pipeTimeout = 5 * 60 * 1000;

	if (iequals(p.subcommand, L"from")) {
		OutputFile outputFile(file);

		auto pipeResult = std::async([&vd, &outputFile, pipeTimeout]{
			CoInit comInit;

			vd.Open(pipeTimeout);
			return processPipeBackup(vd.pDevice, outputFile);
		});

		hr = pipeResult.get();
	}
	else if (iequals(p.subcommand, L"to")) {

		InputFile inputFile(file, 0); // no buffering

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
	FILE* file = nullptr;
	if (p.isBackup()) {
		if (p.to.empty()) {
			file = stdout;
		}
		else {
			int ret = _wfopen_s(&file, p.to.c_str(), L"wb");
			if (!file) {
				std::wcerr << ret << L": Failed to open " << p.to << std::endl;
				return E_FAIL;
			}
		}
	}
	else if (p.isRestore()) {
		if (p.from.empty()) {
			file = stdin;
		}
		else {
			int ret = _wfopen_s(&file, p.from.c_str(), L"rb");
			if (!file) {
				std::wcerr << ret << L": Failed to open " << p.from << std::endl;
				return E_FAIL;
			}
		}
	}
	else if (p.isPipe()) {
		if (iequals(p.subcommand, L"to")) {
			file = stdin;
		}
		else if (iequals(p.subcommand, L"from")) {
			file = stdout;
		}
	}

	if (!file) {
		std::wcerr << L"missing file" << std::endl;
		return E_FAIL;
	}
	
	_setmode(_fileno(file), _O_BINARY);

	HRESULT hr = S_OK;
	
	if (p.isBackup()) {
		RunBackup(p, file);
	}
	else if (p.isRestore()) {
		RunRestore(p, file);
	}
	else if (p.isPipe()) {
		RunPipe(p, file);
	}
	else {		
		std::wcerr << L"unexpected command " << p.command << std::endl;
		return E_FAIL;
	}
	
	if (file != stdout && file != stdin) {
		fclose(file);
	}

	return hr;
}

int wmain(int argc, wchar_t* argv[])
{
	CoInit comInit;
	
	std::wcerr << L"\nmssqlPipe v1.0\n" << std::endl;
	if (argc < 2) {
		showUsage();
		return E_INVALIDARG;
	}

#ifdef _DEBUG
	bool waitForDebugger = true;
	waitForDebugger = false;
	if (waitForDebugger && !::IsDebuggerPresent()) {
		std::wcerr << L"Waiting for debugger..." << std::endl;
		::Sleep(10000);
	}
#endif

	params p;
		
	p.device = make_guid();

	wchar_t** arg = argv;
	wchar_t** argEnd = arg + argc;

	++arg;

	// mssqlpipe [instance] (backup|restore [filelistonly]) [database] dbname [ [from|to] filename|path ]

	auto parseCommand = [&](const wchar_t* sz){
		if (iequals(sz, L"backup")) {			
			p.command = ToLower(sz);
			return true;
		}
		else if (iequals(sz, L"restore")) {
			p.command = ToLower(sz);
			return true;
		}
		else if (iequals(sz, L"pipe")) {
			p.command = ToLower(sz);
			return true;
		}
		else {
			return false;
		}
	};

	if (arg >= argEnd) {
		return invalidArgs(L"missing command");
	}

	if (parseCommand(*arg)) {
		++arg;
	} else {
		p.instance = *arg;

		++arg;

		if (arg >= argEnd) {
			return invalidArgs(L"missing command");
		}

		if (!parseCommand(*arg)) {
			return invalidArgs(L"invalid command", *arg);
		}

		++arg;
	}

	if (arg < argEnd && iequals(*arg, L"filelistonly")) {
		if (!p.isRestore()) {
			return invalidArgs(L"filelistonly only valid for restore");
		}
		p.subcommand = *arg;
		++arg;
	}
	else if (p.isBackupOrRestore()) {
		if (arg < argEnd && iequals(*arg, L"database")) {
			++arg;
		}

		if (arg >= argEnd) {
			return invalidArgs(L"missing database name");
		}

		p.database = *arg;
		++arg;
	}
	
	// [from|to]

	if (arg < argEnd) {

		// backup can only do 'to'
		if (p.isBackup()) {
			if (iequals(*arg, L"to")) {
				++arg;
			
				if (arg >= argEnd) {
					return invalidArgs(L"backup to missing local path");
				}

				p.to = *arg;
				++arg;
			}
		}
		else if (p.isRestore()) {
			// restore can have [from] then [to]
			if (arg < argEnd && iequals(*arg, L"from")) {
				++arg;
			
				if (arg >= argEnd) {
					return invalidArgs(L"restore from: missing local path or file name");
				}

				p.from = *arg;
				++arg;
			}

			if (arg < argEnd && iequals(*arg, L"to")) {
				++arg;
			
				if (arg >= argEnd) {
					return invalidArgs(L"restore to: missing local path");
				}

				p.to = *arg;
				++arg;
				
				::SHCreateDirectoryEx(nullptr, p.to.c_str(), nullptr);
			}

			// then [with] [replace]
			if (arg < argEnd && iequals(*arg, L"with")) {
				++arg;
			}

			if (arg < argEnd) {
				if (iequals(*arg, L"replace")) {
					p.subcommand = *arg;
				}
				else {
					return invalidArgs(L"unexpected argument after restore", *arg);
				}
			}
		}
		else if (p.isPipe()) {
			// pipe can have either to or from followed by device name
			
			if (iequals(*arg, L"to")) {
				p.to = *arg;
			}
			else if (iequals(*arg, L"from")) {
				p.from = *arg;
			}
			else {
				return invalidArgs(L"pipe requires `to` or `from` and a device name");
			}

			p.subcommand = ToLower(*arg);

			++arg;
			
			if (arg >= argEnd) {
				return invalidArgs(L"pipe requires a device name");
			}

			p.device = *arg;

			++arg;
		}
	}

	HRESULT hr = Run(p);
		
	return hr;
}

