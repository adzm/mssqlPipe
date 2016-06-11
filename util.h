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

template <typename char_type,
	typename traits = std::char_traits<char_type> >
	class basic_teebuf :
	public std::basic_streambuf<char_type, traits>
{
public:
	typedef typename traits::int_type int_type;

	basic_teebuf(std::basic_streambuf<char_type, traits>* sb1,
		std::basic_streambuf<char_type, traits>* sb2)
		: sb1(sb1)
		, sb2(sb2)
	{
	}

private:
	virtual int sync() override
	{
		int const r1 = sb1->pubsync();
		int const r2 = sb2->pubsync();
		return r1 == 0 && r2 == 0 ? 0 : -1;
	}

	virtual int_type overflow(int_type c) override
	{
		int_type const eof = traits::eof();

		if (traits::eq_int_type(c, eof))
		{
			return traits::not_eof(c);
		}
		else
		{
			char_type const ch = traits::to_char_type(c);
			int_type const r1 = sb1->sputc(ch);
			int_type const r2 = sb2->sputc(ch);

			return
				traits::eq_int_type(r1, eof) ||
				traits::eq_int_type(r2, eof) ? eof : c;
		}
	}

private:
	std::basic_streambuf<char_type, traits>* sb1;
	std::basic_streambuf<char_type, traits>* sb2;
};

typedef basic_teebuf<char> teebuf;
typedef basic_teebuf<wchar_t> wteebuf;

template <typename char_type,
	typename traits = std::char_traits<char_type> >
class basic_teestream : public std::basic_ostream<char_type, traits>
{
public:
	basic_teestream(std::basic_ostream<char_type, traits>& o1, std::basic_ostream<char_type, traits>& o2)
		: std::basic_ostream<char_type, traits>(&tbuf)
		, tbuf(o1.rdbuf(), o2.rdbuf())
	{
	}

protected:
	basic_teebuf<char_type, traits> tbuf;
};

typedef basic_teestream<char> teestream;
typedef basic_teestream<wchar_t> wteestream;

/****/