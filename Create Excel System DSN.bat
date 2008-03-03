@echo off
REM Batch file to list all the Excel files in the current directory and configure System DSN's for each of them
REM Patrick Mackaaij, 03-03-2008 v1.00
REM
REM Parameters not passed at the moment:
REM - Default directory without trailing slash (DefaultDir)
REM - Excel version (FIL)

REM =========================================== START Main code
REM List all .xls files in the current directory using: 'dir /b /on *.xls'
REM Pass the filenames (without extension) to: createDSN
for /f %%i in ('dir /b /on *.xls') do (set currentFilename=%%~ni) & (set currentFullname=%%~fi) & call :createDSN

REM Quit after the for-loop
goto :quit
REM =========================================== END Main code

REM =========================================== START CreateDSN loop code
:createDSN
REM Configure System DSN using command line tool odbcconf
REM The filename is used as the DSN name
REM If a DSN name already exists it is overwritten
REM DSN is created as ReadOnly
odbcconf configsysdsn "Microsoft Excel Driver (*.xls)" "DSN=%currentFilename%|DBQ=%currentFullname%|ReadOnly=01|Description=Configure System DSN's for Excel files (ODBC)"

REM Exit this loop
goto :eof
REM =========================================== END CreateDSN loop code

REM =========================================== START End code
:quit
REM Open ODBC management so the DSN can be removed after use
start odbcad32
REM =========================================== END End code