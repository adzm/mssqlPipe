#pragma once

#include <string>
#include <algorithm>

#include <nowide/convert.hpp>

using nowide::widen;
using nowide::narrow;

inline std::string ToLower(std::string str)
{
	std::transform(str.begin(), str.end(), str.begin(), towlower);
	return str;
}

inline int icmp(const std::string& a, const std::string& b)
{
	return _stricmp(a.c_str(), b.c_str());
}
inline int icmp(const char* a, const std::string& b)
{
	return _stricmp(a, b.c_str());
}
inline int icmp(const std::string& a, const char* b)
{
	return _stricmp(a.c_str(), b);
}
inline int icmp(const char* a, const char* b)
{
	return _stricmp(a, b);
}

template<typename A, typename B>
bool iequals(A&& a, B&& b)
{
	return 0 == icmp(std::forward<A>(a), std::forward<B>(b));
}

struct iless_predicate
{
	template<typename A, typename B>
	bool operator()(A&& a, B&& b) const
	{
		return icmp(std::forward<A>(a), std::forward<B>(b)) < 0;
	}
};

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

inline std::string escape(std::string str)
{
	size_t quoteCount = count(begin(str), end(str), '\'');

	if (!quoteCount) {
		return str;
	}

	std::string q;
	q.reserve(str.size() + quoteCount);

	for (auto c : str) {
		q.push_back(c);
		if (c == '\'') {
			q.push_back('\'');
		}
	}

	return q;
}

/****/

inline std::vector<std::string> make_argv(int argc, wchar_t** argv)
{
	std::vector<std::string> args;
	args.reserve(argc);
	for (int i = 0; i < argc; ++i) {
		args.push_back(narrow(argv[i]));
	}
	return args;
}

inline std::vector<std::string> make_argv(const char* cmdLine)
{
	auto wcmdLine = widen(cmdLine);
	int argc = 0;
	wchar_t** argv = ::CommandLineToArgvW(wcmdLine.c_str(), &argc);

	std::vector<std::string> args;
	args.reserve(argc);

	for (int i = 0; i < argc; ++i) {
		args.push_back(narrow(argv[i]));
	}

	::LocalFree(argv);

	return args;
}

inline std::vector<const char*> make_argv_ptrs(const std::vector<std::string>& args)
{
	std::vector<const char*> argv;
	argv.reserve(args.size());

	for (auto&& arg : args) {
		argv.push_back(arg.c_str());
	}

	return argv;
}

/****/

inline std::string make_guid()
{
	wchar_t guid[39] = { 0 };
	GUID guidRaw = { 0 };
	::CoCreateGuid(&guidRaw);
	::StringFromGUID2(guidRaw, guid, _countof(guid));
	return narrow(guid);
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
typedef basic_teebuf<char> wteebuf;

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
typedef basic_teestream<char> wteestream;

/****/