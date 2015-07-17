# mssqlPipe v1.0.0
SQL Server command line backup and restore through pipes using stdin and stdout

mssqlPipe 

Usage:

    mssqlPipe [instance] backup [database] dbname [to filename]
    mssqlPipe [instance] restore [database] dbname [from filename] [to filepath] [with replace]
    mssqlPipe [instance] restore filelistonly [from filename]
    mssqlPipe [instance] pipe to devicename
    mssqlPipe [instance] pipe from devicename

stdin or stdout will be used if no filenames specified.

When restoring, default database and log path will be determined for SQL 2012+, otherwise will guess based on existing databases if filepath not specified. Files will be renamed like database_dat.mdf and database_log.ldf. Existing databases will _not_ be overwritten unless you use `with replace`!

Examples:

    mssqlPipe myinstance backup AdventureWorks to AdventureWorks.bak
    mssqlPipe myinstance backup AdventureWorks > AdventureWorks.bak
    mssqlPipe backup database AdventureWorks | 7za a AdventureWorks.xz -si
    7za e AdventureWorks.xz -so | mssqlPipe restore AdventureWorks to c:\db\
    mssqlPipe restore AdventureWorks from AdventureWorks.bak with replace
    mssqlPipe pipe from VirtualDevice42 > output.bak
    mssqlPipe pipe to VirtualDevice42 < input.bak
    mssqlPipe sql2008 backup AdventureWorksOld | mssqlPipe sql2012 restore AdventureWorksOld
    curl -u adzm:hunter2 sftp://adzm.net/backup.xz | 7za e -txz -so -si nul | mssqlPipe restore AdventureWorks

If you need to do anything fancy, use the pipe verb with a devicename of your choosing and run a query manually.

I was inspired by the VDI sample; and apparently so was [sqlpipe](https://github.com/duncansmart/sqlpipe)

Happy piping!