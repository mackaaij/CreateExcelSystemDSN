@echo off
REM Batch file to configure a System DSN for Excel files dropped on this script
REM (or looks for .xls files in the current folder)
REM Opens Windows ODBC management afterwards
REM Note: Overwrites existing DSN's
REM
REM Patrick Mackaaij, 04-03-2008
REM
REM Parameters not passed to the DSN at the moment:
REM - Default directory without trailing slash (DefaultDir)
REM - Excel version (FIL)
title Create Excel System DSN v1.3

REM Setting to check whether a DSN was created.
REM Windows ODBC management is only opened if a DSN was created.
set CreatedDSN=false

REM =========================================== START Check code
REM If no files are dropped check if .xls files exist in the current folder
if [%1]==[] goto :SearchXLS

REM Otherwise parse the variables passed (filenames)
:ParseVariables
REM Check wether the dropped file is an Excel file (based on the extension ".xls")
if NOT [%~x1]==[.xls] echo WARNING ONLY: No DSN is created for the following filename: & echo %1 & echo This filename does not have the extension ".xls". & echo This file is not expected to be an Excel file. & color 4f & pause & goto :shift
REM If the extension is .XLS...
REM Set the passed filename (which includes a full path) to only the filename in currentFilename
REM (The filename is used as the DSN name)
REM And set the currentFullname to the full path
set currentFilename=%~n1
set currentFullname=%1
REM Call the :createDSN routine
call :createDSN

:shift
REM Shift parameters so all dropped files are handled
shift
REM If there are more files to be handled do all of it again
if not [%1]==[] goto :ParseVariables
goto :quit

REM =========================================== END Check code

REM =========================================== START Main createDSN code (called routine)
:createDSN
REM Code to dequote the passed filename, otherwise odbcconf will interpreted the DBQ until the first space
REM Source: http://www.ss64.com/ntsyntax/esc.html
SET Fullname=###%currentFullname%###
SET Fullname=%currentFullname:"###=%
SET Fullname=%currentFullname:###"=%
SET Fullname=%currentFullname:###=%

REM Configure System DSN using command line tool odbcconf
REM The filename is used as the DSN name
REM If a DSN name already exists it is overwritten
REM DSN is created as ReadOnly
odbcconf configsysdsn "Microsoft Excel Driver (*.xls)" "DSN=%currentFilename%|DBQ=%currentFullname%|ReadOnly=01|Description=Created by: Create Excel System DSN.bat"
if NOT "%errorlevel%"=="0" echo ERROR: NO DSN CREATED for the following file: & echo %1 & echo Maybe the filename contains illegal characters like parentheses: (). & color 4f & pause

if "%errorlevel%"=="0" set CreatedDSN=true

REM Exit this loop
goto :eof

REM =========================================== END Main createDSN code (called routine)

REM =========================================== START SearchXLS code
:SearchXLS
REM Check whether the number of .XLS files equals 0. If so, abort script.
for /F %%i in ('dir /b *.xls ^| find /c /v ""') do set numberoffiles=%%i
if %numberoffiles%==0 echo Batch file to configure a System DSN for Excel files & echo. & echo This batch file creates a System DSN for each Excel file dropped on it. & echo (or looks for .xls files in the current folder) & echo The DSN's name is based on the filename. & echo Afterwards Windows ODBC management is started. & echo Note: Existing DSN's will be overwritten. & echo. & echo So drop some Excel files on this batch! & echo (or run it from a folder containing .xls files) & pause & goto :eof

REM List all .xls files in the current directory using: 'dir /b /on *.xls'
REM "delims=" preserves the full filename which would otherwise brake on the first space
REM Pass the filenames (without extension) to: createDSN
for /f "delims=" %%i in ('dir /b /on *.xls') do (set currentFilename=%%~ni) & (set currentFullname=%%~fi) & call :createDSN
REM =========================================== END SearchXLS code

REM =========================================== START End code
:quit
REM Open ODBC management so the DSN can be removed after use
if NOT %CreatedDSN%==false start odbcad32
REM =========================================== END End code