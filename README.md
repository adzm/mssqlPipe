# mssqlPipe v1.2.0

SQL Server command line backup and restore through pipes using stdin and stdout!

## Features

- Familiar T-SQL-like syntax.
- If piping to a non-elevated process, a new elevated process is spawned automatically, redirecting input and output of the original process!
- Automatically renames and moves database files to a specified path.
- Low-level pipe command for complex or customized usage.
- Easily scriptable and automatable backup and restore; errorlevel is set appropriately on process exit.

## Syntax

Syntax is intended to be similar to T-SQL syntax that users of this tool are likely already familiar with.

    mssqlPipe [instance] [as username[:password]] (backup|restore|pipe) ... 

The options `instance` and `as username[:password]` are common to all verbs. Windows authentication (SSPI) will be used if a username is not supplied.

### backup

    mssqlPipe backup [database] dbname [to filename]

### restore

    mssqlPipe restore [database] dbname [from filename] [to filepath] [with replace]
    mssqlPipe restore filelistonly [from filename]

### pipe

If you need to do anything special, `pipe` will simply create the virtual device on the SQL server. Use `to` to pipe stdin to the virtual device, or `from` to pipe the virtual device to stdout.

    mssqlPipe pipe to devicename
    mssqlPipe pipe from devicename
    mssqlPipe pipe to devicename from filename
    mssqlPipe pipe from devicename to filename

## Usage

You can optionally use the word `database` after `backup` and `restore`, like T-SQL `BACKUP` command. Unless your database is literally named 'database'.

stdin or stdout will be used if no filenames specified.

`[as username[:password]]` can be omitted. Windows authentication (SSPI) will be used instead.

When restoring, default database and log path will be determined for SQL 2012+, otherwise will guess based on existing databases if filepath not specified. Files will be renamed like database_dat.mdf and database_log.ldf. Existing databases will _not_ be overwritten unless you use `with replace`!

## Examples

    mssqlPipe myinstance backup AdventureWorks to AdventureWorks.bak
    mssqlPipe myinstance as sa:hunter2 backup AdventureWorks > AdventureWorks.bak
    mssqlPipe backup database AdventureWorks | 7za a AdventureWorks.xz -txz -si
    7za e AdventureWorks.xz -so | mssqlPipe restore AdventureWorks to c:/db/
    mssqlPipe restore AdventureWorks from AdventureWorks.bak with replace
    mssqlPipe pipe from VirtualDevice42 > output.bak
    mssqlPipe pipe to VirtualDevice42 < input.bak
    mssqlPipe pipe to VirtualDevice42 from input.bak
    mssqlPipe sql2008 backup AdventureWorksOld | mssqlPipe sql2012 restore AdventureWorksOld
    curl -u adzm:hunter2 sftp://adzm.net/backup.xz | 7za e -txz -so -si nul | mssqlPipe restore AdventureWorks

If you need to do anything fancy, use the pipe verb with a devicename of your choosing and run a query manually.

## Feedback

Please report and issues and feature requests!

I was inspired by the VDI sample; and apparently so was [sqlpipe](https://github.com/duncansmart/sqlpipe)

Happy piping! :rainbow:
