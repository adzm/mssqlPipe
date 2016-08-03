#pragma once

struct pipestat
{
	__int64 totalBytes = 0;

	DWORD ticksBegin = 0;
	DWORD ticksLastStatus = 0;

	std::mutex& outputMutex;
	bool quiet = false;

	pipestat(std::mutex& outputMutex, bool quiet)
		: ticksBegin(::GetTickCount())
		, outputMutex(outputMutex)
		, quiet(quiet)
	{
		ticksLastStatus = ticksBegin;
	}

	void accumulate(__int64 len)
	{
		totalBytes += len;

		if (quiet || !totalBytes) {
			return;
		}

		DWORD tick = ::GetTickCount();
		DWORD ticksTotal = tick - ticksBegin;
		DWORD ticksSince = tick - ticksLastStatus;

		DWORD threshold = 1000;

		if (ticksTotal <= 15000) {
			threshold = 1000;
		} else if (ticksTotal <= 25000) {
			threshold = 2000;
		} else if (ticksTotal < 60000) {
			threshold = 5000;
		} else {
			threshold = 10000;
		}

		if (ticksSince < threshold) {
			return;
		}

		ticksLastStatus = tick;

		display("Processing... ");
	}

	void finalize()
	{
		if (quiet || !totalBytes) {
			return;
		}

		ticksLastStatus = ::GetTickCount();

		display("Total ");
	}

	void display(const char* prefix)
	{
		DWORD ticksTotal = ticksLastStatus - ticksBegin;
		
		double totalSeconds = ticksTotal / 1000.0;
		if (totalSeconds == 0.0) {
			totalSeconds = 0.1;
		}
		double totalKilobytes = totalBytes / 1024.0;
		double kilobytesPerSec = totalKilobytes / totalSeconds;

		std::unique_lock<std::mutex> lock(outputMutex);

		if (totalSeconds <= 1.0) {
			nowide::cerr << prefix << std::setw(9) << static_cast<int>(totalKilobytes)
				<< " kb in " << " < 1 second"
				<< std::endl;
		}
		else {
			nowide::cerr << prefix << std::setw(9) << static_cast<int>(totalKilobytes)
				<< " kb in " << std::setw(4) << static_cast<int>(totalSeconds) << " seconds ("
				<< std::setw(6) << static_cast<int>(kilobytesPerSec) << " kb/sec)"
				<< std::endl;
		}
	}
};

