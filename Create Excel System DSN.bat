@echo off
REM Batch file to configure a System DSN for a Excel file dropped on this script
REM Patrick Mackaaij, 03-03-2008 v1.1
REM
REM Parameters not passed to the System DSN at the moment:
REM - Excel version (FIL)
REM - Default directory without trailing slash (DefaultDir)

REM =========================================== START Check code
REM Check wether the dropped file is an Excel file (based on the extension ".xls")
if NOT "%~x1"==".xls" echo Received filename: %1 & echo This filename does not have the extension ".xls". & echo This file is not expected to be an Excel file so the script will be aborted. & color 4f & pause & goto :eof
REM =========================================== END Check code

REM =========================================== START CreateDSN code
set Filename=%~n1
set Fullname=%~f1

REM Configure System DSN using command line tool odbcconf
REM The filename is used as the DSN name
REM If a DSN name already exists it is overwritten
REM DSN is created as ReadOnly
odbcconf configsysdsn "Microsoft Excel Driver (*.xls)" "DSN=%Filename%|DBQ=%Fullname%|ReadOnly=01|Description=Configure System DSN's for Excel files (ODBC)"
REM =========================================== END CreateDSN code

REM =========================================== START End code
:quit
REM Open ODBC management so the DSN can be removed after use
start odbcad32
REM =========================================== END End code