@echo off
REM Batch file to configure a System DSN for Excel files dropped on this script
REM Opens Windows ODBC management afterwards
REM Note: Overwrites existing DSN's
REM
REM Patrick Mackaaij, 03-03-2008 v1.2
REM
REM Parameters not passed to the DSN at the moment:
REM - Default directory without trailing slash (DefaultDir)
REM - Excel version (FIL)

REM =========================================== START Check code
REM Explain this batch file if no file is dropped
if %1.==. echo Batch file to configure a System DSN for Excel files & echo. & echo This batch file creates a System DSN for each Excel file dropped on it. & echo The DSN's name is based on the filename. & echo Afterwards Windows ODBC management is started. & echo Note: Existing DSN's will be overwritten. & echo. & echo So drop some Excel files on this batch! & pause & goto :eof
REM =========================================== END Check code

REM =========================================== START Main code
:main
REM Check wether the dropped file is an Excel file (based on the extension ".xls")
if NOT "%~x1"==".xls" echo WARNING ONLY: No DSN is created for the following filename: & echo %1 & echo This filename does not have the extension ".xls". & echo This file is not expected to be an Excel file. & color 4f & pause & goto :shift

REM Parse the passed filename (which includes a full path) to only the filename
REM The filename is used as the DSN name
set Filename=%~n1
set Fullname=%1

REM Code to dequote the passed filename, otherwise odbcconf will interpreted the DBQ until the first space
REM Source: http://www.ss64.com/ntsyntax/esc.html
SET Fullname=###%Fullname%###
SET Fullname=%Fullname:"###=%
SET Fullname=%Fullname:###"=%
SET Fullname=%Fullname:###=%

REM Configure System DSN using command line tool odbcconf
REM The filename is used as the DSN name
REM If a DSN name already exists it is overwritten
REM DSN is created as ReadOnly
start /wait odbcconf configsysdsn "Microsoft Excel Driver (*.xls)" "DSN=%Filename%|DBQ=%Fullname%|ReadOnly=01|Description=Configure System DSN's for Excel files (ODBC)"
if NOT "%errorlevel%"=="0" echo ERROR: NO DSN CREATED for the following file: & echo %1 & echo Maybe the filename contains illegal characters like parentheses: (). & color 4f & pause

:shift
REM Shift parameters so all dropped files are handled
shift
REM If there are more files to be handled do all of it again
if not %1.==. goto :main
goto :quit
REM =========================================== END Main code

REM =========================================== START End code
:quit
REM Open ODBC management so the DSN can be removed after use
start odbcad32
REM =========================================== END End code