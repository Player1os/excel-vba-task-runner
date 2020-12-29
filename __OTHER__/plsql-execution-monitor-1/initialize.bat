@echo off

:: Initialize the SQL*PLUS script environment.
call H:\sql-plus\initialize "H:\projects\plsql-execution-monitor"

:: Create the script's temp directory and set the corresponding environment variable.
set APP_TEMP_DIRECTORY_PATH=H:\temp\plsql_execution_monitor
mkdir "%APP_TEMP_DIRECTORY_PATH%"

:: Invoke the verify script.
call verify

:: Delete the script's temp directory and unset the corresponding environment variable.
rmdir "%APP_TEMP_DIRECTORY_PATH%"
set APP_TEMP_DIRECTORY_PATH=

:: Terminate the SQL*PLUS script environment.
call H:\sql-plus\terminate
