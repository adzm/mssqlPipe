#pragma once

#include <string>
#include <algorithm>


inline std::wstring ToLower(std::wstring str)
{
	std::transform(str.begin(), str.end(), str.begin(), towlower);
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

/****/

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


/****/

inline std::wstring make_guid()
{
	std::wstring guid;
	GUID guidRaw = { 0 };
	::CoCreateGuid(&guidRaw);
	guid.resize(38);
	::StringFromGUID2(guidRaw, &guid[0], guid.size() + 1);
	return guid;
}

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

inline void ComIssueError(HRESULT hr)
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


inline void ComEnsure(HRESULT hr)
{
	if (FAILED(hr)) {
		ComIssueError(hr);
	}
}


/****/

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



/****/


