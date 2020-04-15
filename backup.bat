@echo off
REM ----------------------------------------------------------------
REM Create a directory to save mysql backup files if not already exists REM ----------------------------------------------------------------

IF NOT EXIST "C:\MYSQLBACKUPS" mkdir C:\MYSQLBACKUPS

REM ----------------------------------------------------------------
REM append date and time to mysqldump files
REM ----------------------------------------------------------------

SET dt=%date:~-4%_%date:~3,2%_%date:~0,2%_%time:~0,2%_%time:~3,2%_%time:~6,2%


set bkupfilename=%dt%.sql

REM ----------------------------------------------------------------
REM Display some message on the screen about the backup
REM ----------------------------------------------------------------
ECHO Starting Backup of MySQL Database
ECHO Backup is going to save in C:\MYSQLBACKUPS\ folder.
ECHO Please wait ...

REM ----------------------------------------------------------------
REM mysqldump backup command. append date and time in filename
REM ----------------------------------------------------------------

"C:\Program Files\MySQL\MySQL Server 5.7\bin\mysqldump.exe"  --routines -u root -proot church> C:\MYSQLBACKUPS\"mysqldb_%bkupfilename%"

REM ----------------------------------------------------------------
REM delete mysqldump backups older than 60 days
REM ----------------------------------------------------------------

ECHO.
ECHO Trying to find and delete backups older than 90 days if found.
ECHO And the result is:
forfiles /p C:\MYSQLBACKUPS /s /m *.* /d -3 /c "cmd /c del @file : date >= 60days"

ECHO.
ECHO Backup completed!
ECHO Backup saved in C:\MYSQLBACKUPS\ 
ECHO Thank You for backing up! 
ECHO - Regards, Admin!
ECHO.
ECHO I am about to show you the backup files.

PAUSE

REM Show user the backup files
EXPLORER C:\MYSQLBACKUPS\
EXIT

REM Author: JK
REM Kohima, Nagaland
REM Modified: OCtober 7, 2016 