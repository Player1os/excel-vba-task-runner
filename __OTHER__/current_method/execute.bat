:: Set the script runtime datetime.
for /F "usebackq tokens=1,2 delims==" %%i in (`wmic os get LocalDateTime /VALUE 2^>NUL`) do if '.%%i.'=='.LocalDateTime.' set APP_DATETIME=%%j
set APP_DATETIME=%APP_DATETIME:~0,4%_%APP_DATETIME:~4,2%_%APP_DATETIME:~6,2%-%APP_DATETIME:~8,2%_%APP_DATETIME:~10,2%_%APP_DATETIME:~12,6%

:: Set the directory paths.
set APP_DATABASE_CONNECTION_DIRECTORY_PATH=H:\database-connection
set APP_PROJECT_DIRECTORY_PATH=H:\project
set APP_LOG_DIRECTORY_PATH=H:\log\%2\%APP_DATETIME%
set APP_SCRIPT_DIRECTORY_PATH=H:\script
set APP_TEMP_DIRECTORY_PATH=H:\temp

:: Set the report recipient email address parameter.
set APP_REPORT_RECIPIENT_EMAIL_ADDRESS=user@example.com

:: Create the temp directory.
mkdir "%APP_TEMP_DIRECTORY_PATH%"

:: Generate and load the sql-plus connection parameters according to the first argument.
call "%APP_SCRIPT_DIRECTORY_PATH%\generate_sql_plus_parameters" %1
call "%APP_TEMP_DIRECTORY_PATH%\sql_plus_parameters"

:: Store the current working directory and change it to the second argument.
set APP_ORIGINAL_WORKING_DIRECTORY=%CD%
cd "%APP_PROJECT_DIRECTORY_PATH%\%2"

:: Invoke the project's script.
mkdir "%APP_LOG_DIRECTORY_PATH%"
set APP_OUTPUT_LOG_FILE_PATH=%APP_LOG_DIRECTORY_PATH%\Output.log
set APP_ERROR_LOG_FILE_PATH=%APP_LOG_DIRECTORY_PATH%\Error.log
echo * Executing execute.bat of project %CD% * >%APP_OUTPUT_LOG_FILE_PATH%
echo * BEGIN: %date% - %time% * >>%APP_OUTPUT_LOG_FILE_PATH%
call execute >>%APP_OUTPUT_LOG_FILE_PATH% 2>%APP_ERROR_LOG_FILE_PATH%
echo * END: %date% - %time% * >>%APP_OUTPUT_LOG_FILE_PATH%

:: Check for and handle any errors the project's script may have run into.
call "%APP_SCRIPT_DIRECTORY_PATH%\handle_errors" %2

:: Clear the internal script variables.
set APP_OUTPUT_LOG_FILE_PATH=
set APP_ERROR_LOG_FILE_PATH=

:: Restore and clear the stored current working directory.
cd %APP_ORIGINAL_WORKING_DIRECTORY%
set APP_ORIGINAL_WORKING_DIRECTORY=

:: Clear the sql-plus connection parameters.
set SQL_PLUS_USERNAME=
set SQL_PLUS_DATABASE=
set SQL_PLUS_PASSWORD_FILE_PATH=
set NLS_LANG=

:: Remove the temp directory.
rmdir /s /q "%APP_TEMP_DIRECTORY_PATH%"

:: Clear the report recipient email address parameter.
set APP_REPORT_RECIPIENT_EMAIL_ADDRESS=

:: Clear the directory paths.
set APP_DATABASE_CONNECTION_DIRECTORY_PATH=
set APP_PROJECT_DIRECTORY_PATH=
set APP_LOG_DIRECTORY_PATH=
set APP_SCRIPT_DIRECTORY_PATH=
set APP_TEMP_DIRECTORY_PATH=

:: Clear the script runtime datetime.
set APP_DATETIME=
