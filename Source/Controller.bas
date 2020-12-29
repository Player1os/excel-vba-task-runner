Attribute VB_Name = "Controller"
Option Explicit
Option Private Module

' Requires reference: MSXML2
' Requires reference: Scripting
' Requires external module: ModMail
' Requires module: ModFileSystem
' Requires module: Runtime
' Requires module: ThisUserForm
' Requires module: ThisWorkbook

Private vConfig As Dictionary

Private Sub pLoadConfig()
    ' Early exit if the config has already been loaded.
    If Not (vConfig Is Nothing) Then
        Exit Sub
    End If

    ' Initialize the result.
    Set vConfig = New Dictionary

    ' Load the xml file.
    With New MSXML2.DOMDocument60
        ' Configure to load files asynchronously.
        .async = False

        ' Load the build configuration xml file.
        Call .Load(Runtime.FileSystemObject().BuildPath(ThisWorkbook.Path, "Config.xml"))

        ' Load the root xml node.
        With .SelectSingleNode("config")
            With .SelectSingleNode("run-attempt")
                Call vConfig.Add("run-attempt/maximum-count", CLng(.SelectSingleNode("maximum-count").Text))
                Call vConfig.Add("run-attempt/cooldown-seconds", CLng(.SelectSingleNode("cooldown-seconds").Text))
            End With
            With .SelectSingleNode("log")
                Call vConfig.Add("log/directory-path", .SelectSingleNode("directory-path").Text)
                Call vConfig.Add("log/maximum-record-count", CLng(.SelectSingleNode("maximum-record-count").Text))
            End With
            With .SelectSingleNode("error-report")
                Call vConfig.Add("error-report/sender-address", .SelectSingleNode("sender-address").Text)
                Call vConfig.Add("error-report/recipient-address", .SelectSingleNode("recipient-address").Text)
            End With
        End With

        ' Validate the field values.
        If vConfig("run-attempt/maximum-count") < 1 Then
            Call Runtime.RaiseError("Controller.pLoadConfig", "The 'run-attempt/maximum-count' field must be a positive integer.")
        End If
        If vConfig("run-attempt/cooldown-seconds") < 1 Then
            Call Runtime.RaiseError("Controller.pLoadConfig", "The 'run-attempt/cooldown-seconds' field must be a positive integer.")
        End If
        If Not ModFileSystem.IsValidPath(vConfig("log/directory-path")) Then
            Call Runtime.RaiseError("Controller.pLoadConfig", "The 'log/directory-path' field must be a valid file system path.")
        End If
        If vConfig("log/maximum-record-count") < 1 Then
            Call Runtime.RaiseError("Controller.pLoadConfig", "The 'log/maximum-record-count' field must be a positive integer.")
        End If
        If Not ModMail.IsValidAddress(vConfig("error-report/sender-address")) Then
            Call Runtime.RaiseError("Controller.pLoadConfig", "The 'error-report/sender-address' field must be a valid email address.")
        End If
        If Not ModMail.IsValidAddress(vConfig("error-report/recipient-address")) Then
            Call Runtime.RaiseError("Controller.pLoadConfig", "The 'error-report/recipient-address' field must be a valid email address.")
        End If
    End With
End Sub

Private Sub pHandleUnreportedError()
    ' Declare local variables.
    Dim vRecipients As New Dictionary
    Dim vErrorLogFilePath As String

    ' Load the file system object.
    With Runtime.FileSystemObject()
        ' Determine the path to the error log file.
        vErrorLogFilePath = .BuildPath(ThisWorkbook.Path, "Error.log")

        ' Early exit if the error log file does not exist.
        If Not .FileExists(vErrorLogFilePath) Then
            Exit Sub
        End If

        ' Send the unreported error by mail.
        Call vRecipients.Add("To", vConfig("error-report/recipient-address"))
        Call ModMail.Send(vRecipients, _
            "[" & Runtime.ProjectName() & "] An unreported error has been detected", _
            Runtime.ReadFile(vErrorLogFilePath), _
            vConfig("error-report/sender-address"))

        ' Delete the error log file.
        Call .DeleteFile(vErrorLogFilePath)
    End With
End Sub

Private Function pTimestamp()
    ' TODO: Refactor.
'    :: Store the current time and date.
'    set APP_TIME=%TIME%
'    set APP_DATE=%DATE%
'
'    :: Determine the current year.
'    set APP_YEAR=%APP_DATE:~-4%
'
'    :: Determine the current month.
'    set APP_MONTH=%APP_DATE:~-7,2%
'    if "%APP_MONTH:~0,1%" == " " (
'        set APP_MONTH=0%APP_MONTH:~1,1%
'    )
'
'    :: Determine the current day.
'    set APP_DAY=%APP_DATE:~-10,2%
'    if "%APP_DAY:~0,1%" == " " (
'        set APP_DAY=0%APP_DAY:~1,1%
'    )
'
'    :: Determine the current hour.
'    set APP_HOUR=%APP_TIME:~0,2%
'    if "%APP_HOUR:~0,1%" == " " (
'        set APP_HOUR=0%APP_HOUR:~1,1%
'    )
'
'    :: Determine the current minute.
'    set APP_MINUTE=%APP_TIME:~3,2%
'    if "%APP_MINUTE:~0,1%" == " " (
'        set APP_MINUTE=0%APP_MINUTE:~1,1%
'    )
'
'    :: Determine the current second.
'    set APP_SECOND=%APP_TIME:~6,2%
'    if "%APP_SECOND:~0,1%" == " " (
'        set APP_SECOND=0%APP_SECOND:~1,1%
'    )
'
'    :: Combine the collected parts into a timestamp.
'    set APP_TIMESTAMP=%APP_YEAR%%APP_MONTH%%APP_DAY%_%APP_HOUR%%APP_MINUTE%%APP_SECOND%
'
'    :: Clear the collected parts.
'    set APP_SECOND=
'    set APP_MINUTE=
'    set APP_HOUR=
'    set APP_DAY=
'    set APP_MONTH=
'    set APP_YEAR=
'
'    :: Clear the stored time and date.
'    set APP_DATE=
'    set APP_TIME=
End Function

Private Sub pSendErrorMail()
    ' TODO: Refactor.
'    ' Load external Modules.
'    Dim vOutlookApplication: Set vOutlookApplication = CreateObject("Outlook.Application")
'    Dim vWScriptShell: Set vWScriptShell = CreateObject("WScript.Shell")
'
'    ' Define constants.
'    Const vOlMailItem = 0
'    Const vVbCritical = 16
'
'    ' Disable error handling.
'    On Error Resume Next
'
'    ' Allocate an email item.
'    Dim vMailItem: Set vMailItem = vOutlookApplication.CreateItem(vOlMailItem)
'    Call vOutlookApplication.Session.Logon
'
'    ' Prepare the email and send it.
'    vMailItem.To = vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_ERROR_MAIL_RECIPIENT%")
'    vMailItem.Subject = "[Task Runner] Failed to execute task '" & vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_TASK_NAME%") & "'"
'    vMailItem.HTMLBody = "<div style=""font-family: Arial; font-size: 10pt"">" _
'        & "<p><b>Error message:</b> <code style=""background-color: #eee; color: #c00"">" _
'        & WScript.Arguments(0) _
'        & "</code></p>" _
'        & "<p><b>Task name:</b> <code style=""background-color: #eee; color: #c00"">"
'    vMailItem.HTMLBody = vMailItem.HTMLBody _
'        & vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_TASK_NAME%") _
'        & "</code></p>" _
'        & "<p><b>User name:</b> <code style=""background-color: #eee; color: #c00"">"
'    vMailItem.HTMLBody = vMailItem.HTMLBody _
'        & vWScriptShell.ExpandEnvironmentStrings("%USERNAME%") _
'        & "</code></p>" _
'        & "<p><b>Machine name:</b> <code style=""background-color: #eee; color: #c00"">"
'    vMailItem.HTMLBody = vMailItem.HTMLBody _
'        & vWScriptShell.ExpandEnvironmentStrings("%COMPUTERNAME%") _
'        & "</code></p>" _
'        & "<p><b>Script file path:</b> <code style=""background-color: #eee; color: #c00"">"
'    vMailItem.HTMLBody = vMailItem.HTMLBody _
'        & vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_SCRIPT_FILE_PATH%") _
'        & "</code></p>" _
'        & "<p><b>Start timestamp:</b> <code style=""background-color: #eee; color: #c00"">"
'    vMailItem.HTMLBody = vMailItem.HTMLBody _
'        & vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_START_TIMESTAMP%") _
'        & "</code></p>" _
'        & "<p><b>End timestamp:</b> <code style=""background-color: #eee; color: #c00"">"
'    vMailItem.HTMLBody = vMailItem.HTMLBody _
'        & vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_END_TIMESTAMP%") _
'        & "</code></p>" _
'        & "<p><b>Return code:</b> <code style=""background-color: #eee; color: #c00"">"
'    vMailItem.HTMLBody = vMailItem.HTMLBody _
'        & vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_RETURN_CODE%") _
'        & "</code></p>" _
'        & "<p><b>Log directory:</b> <code style=""background-color: #eee; color: #c00"">"
'    vMailItem.HTMLBody = vMailItem.HTMLBody _
'        & vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%") _
'        & "</code></p>" _
'        & "</div>"
'    Call vMailItem.Send
'
'    ' Disconnect and disable the imported libraries.
'    Call vOutlookApplication.Session.Logoff
'
'    ' Check whether reporting the error was finished without error, otherwise display a message box to the user.
'    If Err.Number <> 0 Then
'        Call MsgBox("An unexpected error had occured while attemtping to report an error regarding task '" _
'            & vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_TASK_NAME%") & "'", vVbCritical, "Task runner")
'    End If
'
'    ' Reenable error handling.
'    On Error GoTo 0
End Sub

Private Sub pRemoveExtraLogs()
    ' TODO: Refactor.
'    ' Load external modules.
'    Dim vFileSystemObject: Set vFileSystemObject = CreateObject("Scripting.FileSystemObject")
'    Dim vWscriptShell: Set vWscriptShell = CreateObject("WScript.Shell")
'
'    ' Define a collection of folder paths.
'    Dim vLogFolderPaths: Set vLogFolderPaths = CreateObject("Scripting.Dictionary")
'
'    ' Load the runtime parameters.
'    Dim vTaskLogDirectoryPath: vTaskLogDirectoryPath = _
'        vWscriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_TASK_LOG_DIRECTORY_PATH%")
'    Dim vMaximumTaskLogCount: vMaximumTaskLogCount = _
'        CLng(vWscriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_MAXIMUM_TASK_LOG_COUNT%"))
'
'    ' Store the current subdirectories of the specified task's log directory.
'    Dim vLogFolder
'    Dim vLogFolderIndex
'    vLogFolderIndex = 1
'    For Each vLogFolder In vFileSystemObject.GetFolder(vTaskLogDirectoryPath).SubFolders
'        Call vLogFolderPaths.Add(vLogFolderIndex, vLogFolder.Path)
'        vLogFolderIndex = vLogFolderIndex + 1
'    Next
'
'    ' Remove the oldest extraneous log subdirectories of the specified task's log directory.
'    For vLogFolderIndex = 1 To vLogFolderPaths.Count - vMaximumTaskLogCount
'        Call vFileSystemObject.DeleteFolder(vLogFolderPaths(vLogFolderIndex))
'    Next
End Sub

Private Sub pStartPathHandler( _
    ByRef vParameters As Dictionary _
)
    ' Declare local variables.
    Dim vTaskName As String
    Dim vScriptFilePath As String

    ' Validate the task_name parameter.
    If Not vParameters.Exists("task_name") Then
        Call Runtime.RaiseError("Controller.Navigate", "The 'task_name' parameter must be set.")
    End If
    If vParameters("task_name") = vbNullString Then
        Call Runtime.RaiseError("Controller.Navigate", "The 'task_name' parameter must be non-empty.")
    End If

    ' Validate the script_file_path parameter.
    If Not vParameters.Exists("script_file_path") Then
        Call Runtime.RaiseError("Controller", "The 'script_file_path' parameter must be set.")
    End If
    If vParameters("script_file_path") = vbNullString Then
        Call Runtime.RaiseError("Controller", "The 'script_file_path' parameter must be non-empty.")
    End If

    Dim vRunAttemptCount As Long

    vStartDateTimeStamp = Runtime.DateTimeStamp(Now)

    Do While Not Runtime.FileSystemObject().FileExists(vParameters("script_file_path"))
        Call WScript.Sleep(vConfig("run-attempt/cooldown-seconds"))

        vRunAttemptCount = vRunAttemptCount + 1
        If vRunAttemptCount > vConfig("run-attempt/maximum-count") Then
            ' TODO. Raise an error.
        End If
    Loop

    With Runtime.WScriptShell().Exec(vParameters("script_file_path"))
        ' Log that the task is running with the generated pid by creating a file.
        ' TODO.

        ' Give the subprocess time to execute.
        Do While .Status <> WshFinished
            Call WScript.Sleep(100)
        Loop

        ' Delete the log of the generated pid.
        ' TODO.
    End With

    vEndDateTimeStamp = Runtime.DateTimeStamp(Now)



    ' TODO: Implement.
    ' - Must be able to kill all processed created for tasks with the same name.
    ' - Must be able to retry running the task if the script is not available.

'    :: Set the iteration counter.
'    set APP_TASK_RUNNER_ITERATION_COUNTER=0
'
'    :: Store the start timestamp.
'    call "%~dp0timestamp.bat"
'    set APP_TASK_RUNNER_START_TIMESTAMP=%APP_TIMESTAMP%
'    set APP_TIMESTAMP=
'
'    :: Determine the current task log directory.
'    set APP_TASK_RUNNER_TASK_LOG_DIRECTORY_PATH=%APP_TASK_RUNNER_LOG_DIRECTORY_PATH%\%APP_TASK_RUNNER_TASK_NAME%
'    set APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH=%APP_TASK_RUNNER_TASK_LOG_DIRECTORY_PATH%\%APP_TASK_RUNNER_START_TIMESTAMP%
'
'    :: Check whether the target script is available.
'    :check_availability
'    if not exist %APP_TASK_RUNNER_SCRIPT_FILE_PATH% (
'        :: Wait for the specified amount of time.
'        ping 127.0.0.1 -n %APP_TASK_RUNNER_WAIT_SECONDS% >nul
'
'        :: Increment the iteration counter.
'        set /a APP_TASK_RUNNER_ITERATION_COUNTER=%APP_TASK_RUNNER_ITERATION_COUNTER%+1
'
'        :: Check whether the iteration count was reached.
'        if not %APP_TASK_RUNNER_ITERATION_COUNTER% lss %APP_TASK_RUNNER_ITERATION_COUNT% (
'            :: Reset the return code, current log directory path and end timestamp.
'            set APP_TASK_RUNNER_RETURN_CODE=N/A
'            set APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH=N/A
'            set APP_TASK_RUNNER_END_TIMESTAMP=N/A
'
'            :: Report the error to the user.
'            cscript /NoLogo "%~dp0\send_error_mail.vbs" "The '%APP_TASK_RUNNER_TASK_NAME%' task script file was not found to be available."
'
'            :: Jump to the termination section.
'            goto :terminate
'        )
'
'        :: Jump back to the iteration condition.
'        goto :check_availability
'    )
'
'    :: Create the current task log directory.
'    mkdir "%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%"
'
'    :: Execute the script in the submitted file path.
'    call "%APP_TASK_RUNNER_SCRIPT_FILE_PATH%" ^
'        >"%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%\out.log" ^
'        2>"%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%\err.log"
'
'    :: Store the return code.
'    set APP_TASK_RUNNER_RETURN_CODE=%ERRORLEVEL%
'
'    :: Store the end timestamp.
'    call "%~dp0timestamp.bat"
'    set APP_TASK_RUNNER_END_TIMESTAMP=%APP_TIMESTAMP%
'    set APP_TIMESTAMP=
'
'    :: Output additional information about the execution of the script.
'    echo User name: %USERNAME% >"%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%\info.log"
'    echo Machine name: %COMPUTERNAME% >>"%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%\info.log"
'    echo Script file path: %APP_TASK_RUNNER_SCRIPT_FILE_PATH% >>"%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%\info.log"
'    echo Start timestamp: %APP_TASK_RUNNER_START_TIMESTAMP% >>"%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%\info.log"
'    echo End timestamp: %APP_TASK_RUNNER_END_TIMESTAMP% >>"%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%\info.log"
'    echo Return code: %APP_TASK_RUNNER_RETURN_CODE% >>"%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%\info.log"
'
'    :: Remove any extraenous log files, that exceed the specified limit.
'    cscript /NoLogo "%~dp0\remove_extra_logs.vbs"
'
'    :: Check whether an error had occurred.
'    if %APP_TASK_RUNNER_RETURN_CODE% neq 0 (
'        :: Report the error to the user.
'        cscript /NoLogo "%~dp0\send_error_mail.vbs" "The '%APP_TASK_RUNNER_TASK_NAME%' task script has returned a non-zero status code."
'
'        :: Jump to the termination section.
'        goto :terminate
'    )
'
'    :terminate
'
'    :: Clear runtime parameters.
'    set APP_TASK_RUNNER_RETURN_CODE=
'    set APP_TASK_RUNNER_END_TIMESTAMP=
'    set APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH=
'    set APP_TASK_RUNNER_TASK_LOG_DIRECTORY_PATH=
'    set APP_TASK_RUNNER_START_TIMESTAMP=
'    set APP_TASK_RUNNER_ITERATION_COUNTER=
'    set APP_TASK_RUNNER_SCRIPT_FILE_PATH=
End Sub

Private Sub pKillPathHandler( _
    ByRef vParameters As Dictionary _
)
    ' Declare local variables.
    Dim vTaskName As String

    ' Validate the task_name parameter.
    If Not vParameters.Exists("task_name") Then
        Call Runtime.RaiseError("Controller.pKillPathHandler", "The 'task_name' parameter must be set.")
    End If
    If vParameters("task_name") = vbNullString Then
        Call Runtime.RaiseError("Controller.pKillPathHandler", "The 'task_name' parameter must be non-empty.")
    End If

    ' TODO: Implement.
End Sub

Public Sub Navigate( _
    ByRef vPath As String, _
    ByRef vParameters As Scripting.Dictionary _
)
    With Runtime.FileSystemObject()
        Call Runtime.WriteFile(.BuildPath(.BuildPath(ThisWorkbook.Path, "Temp"), "test2.txt"), Runtime.ReadFile(.BuildPath(.BuildPath(ThisWorkbook.Path, "Temp"), "test2.txt")))
        'Call ThisUserForm.SetInnerHtml("<p>" & ModFileSystem.ReadFile(.BuildPath(.BuildPath(ThisWorkbook.Path, "Temp"), "test.txt")) & "</p>")
    End With

'    ' Load the configuration, if not yet loaded.
'    Call pLoadConfig
'
'    ' Handle previously unreported error, if found.
'    Call pHandleUnreportedError
'
'    ' Route behaviour based on path and execute the suitable handler.
'    Select Case vPath
'        Case "start"
'            Call pStartPathHandler(vParameters)
'        Case "kill"
'            Call pKillPathHandler(vParameters)
'        Case Else
'            Call Runtime.RaiseError("Controller.Navigate", "An undefined path '" & vPath & "' has been submitted.")
'    End Select

'    ' Declare local variables.
'    Dim vValues As New Scripting.Dictionary
'    Dim vName As Variant
'    Dim vOutput As String
'
'    ' Fill the output values.
'    With vValues
'        Call .Add("Date & Time", Runtime.DateTimeStamp(Now))
'        Call .Add("User @ Computer", Runtime.Username() & "@" & Runtime.ComputerName())
'        Call .Add("Navigate Path", Runtime.GenerateNavigatePath(vPath, vParameters))
'    End With
'
'    ' If debug mode is enabled, print the current state to the immediate window.
'    If Runtime.IsDebugModeEnabled() Then
'        Debug.Print "===================== Output ====================="
'        For Each vName In vValues.Keys()
'            Debug.Print Trim("[" & CStr(vName) & "] " & CStr(vValues(vName)))
'        Next
'    End If
'
'    ' Check whether the application is running in background mode.
'    If Runtime.IsBackgroundModeEnabled() Then
'        ' Output the current state to a file.
'        For Each vName In vValues.Keys()
'            vOutput = vOutput & Trim("[" & CStr(vName) & "] " & CStr(vValues(vName))) & vbLf
'        Next
'        With Runtime.FileSystemObject()
'            Call Runtime.AppendFile(.BuildPath(ThisWorkbook.Path, "Output.log"), vOutput, vbLf)
'        End With
'    Else
'        ' Display the current state on the loaded html page.
'        For Each vName In vValues.Keys()
'            vOutput = vOutput & "<p><b>" & CStr(vName) & "</b>: <code>" & CStr(vValues(vName)) & "</code></p>"
'        Next
'        Call ThisUserForm.SetInnerHtml("<h1>Navigate</h1>" & vOutput)
'    End If
End Sub

Public Sub HandleError( _
    ByRef vPath As String, _
    ByRef vParameters As Dictionary, _
    ByRef vErrorMessage As String _
)
    ' Declare local variables.
    Dim vValues As New Scripting.Dictionary
    Dim vName As Variant
    Dim vOutput As String

    ' Fill the output values.
    With vValues
        Call .Add("Date & Time", Runtime.DateTimeStamp(Now))
        Call .Add("User @ Computer", Runtime.Username() & "@" & Runtime.ComputerName())
        Call .Add("Navigate Path", Runtime.GenerateNavigatePath(vPath, vParameters))
        Call .Add("Error Number", CStr(Err.Number))
        Call .Add("Error Source", Err.Source)
        Call .Add("Error Description", Err.Description)
    End With

    ' If debug mode is enabled, print the current state to the immediate window.
    If Runtime.IsDebugModeEnabled() Then
        Debug.Print "===================== Error ======================"
        For Each vName In vValues.Keys()
            Debug.Print Trim("[" & CStr(vName) & "] " & CStr(vValues(vName)))
        Next
        Debug.Print Trim("[Error Message] " & vErrorMessage)
    End If

    ' Check whether the application is running in background mode.
    If Runtime.IsBackgroundModeEnabled() Then
        ' Output the current state to a file.
        For Each vName In vValues.Keys()
            vOutput = vOutput & Trim("[" & CStr(vName) & "] " & CStr(vValues(vName))) & vbLf
        Next
        vOutput = vOutput & Trim("[Error Message] " & vErrorMessage) & vbLf
        With Runtime.FileSystemObject()
            Call Runtime.AppendFile(.BuildPath(ThisWorkbook.Path, "Error.log"), vOutput, vbLf)
        End With
    Else
        ' Display the current state on the loaded html page.
        For Each vName In vValues.Keys()
            vOutput = vOutput & "<p><b>" & CStr(vName) & "</b>: <code>" & CStr(vValues(vName)) & "</code></p>"
        Next
        Call ThisUserForm.SetInnerHtml("<h1>Navigate</h1>" & vOutput)

        ' Show dialog box with an error message.
        Call MsgBox(IIf(vErrorMessage = vbNullString, "An unknown unexpected error had occurred.", vErrorMessage), _
            vbCritical, "Error Message")
    End If
'    ' If debug mode is enabled, print the current state to the immediate window.
'    If Runtime.IsDebugModeEnabled() Then
'        Debug.Print "===================== Error ======================"
'        Debug.Print "[Date & Time] " & Format(Now, "yyyy-mm-dd Hh:Nn:Ss")
'        Debug.Print "[User @ Computer] " & Runtime.Username() & "@" & Runtime.ComputerName()
'        Debug.Print "[Navigate Path] " & Runtime.GenerateNavigatePath(vPath, vParameters)
'        Debug.Print "[Error Number] " & CStr(Err.Number)
'        Debug.Print "[Error Source] " & Err.Source
'        Debug.Print "[Error Description] " & Err.Description
'        Debug.Print "[Error Message] " & vErrorMessage
'    Else
'        ' TODO Send Mail.
'
'
'        ' Display the path and parameters on the loaded html page.
'        Call ThisUserForm.SetInnerHtml("<h1>Handle Error</h1>" _
'            & "<p><b>Date & Time</b>: <code>" & Format(Now, "yyyy-mm-dd Hh:Nn:Ss") & "</code></p>" _
'            & "<p><b>User @ Computer</b>: <code>" & Runtime.Username() & "@" & Runtime.ComputerName() & "</code></p>" _
'            & "<p><b>Navigate Path</b>: <code>" & Runtime.GenerateNavigatePath(vPath, vParameters) & "</code></p>" _
'            & "<p><b>Error Number</b>: <code>" & CStr(Err.Number) & "</code></p>" _
'            & "<p><b>Error Source</b>: <code>" & Err.Source & "</code></p>" _
'            & "<p><b>Error Description</b>: <code>" & Err.Description & "</code></p>")
'
'        Call MsgBox(IIf(vErrorMessage = vbNullString, "An unknown unexpected error had occurred.", vErrorMessage), _
'            vbCritical, "Error Message")
'
'        ' TODO Handle in case of failure.
'        ' Output the current timestamp to a file.
'        With Runtime.FileSystemObject()
'            With .OpenTextFile(.BuildPath(ThisWorkbook.Path, "Error.log"), ForAppending, True)
'                Call .WriteLine("[Date & Time] " & Format(Now, "yyyy-mm-dd Hh:Nn:Ss"))
'                Call .WriteLine("[User @ Computer] " & Runtime.Username() & "@" & Runtime.ComputerName())
'                Call .WriteLine("[Navigate Path] " & Runtime.GenerateNavigatePath(vPath, vParameters))
'                Call .WriteLine("[Error Number] " & CStr(Err.Number))
'                Call .WriteLine("[Error Source] " & Err.Source)
'                Call .WriteLine("[Error Description] " & Err.Description)
'                Call .WriteLine("[Error Message] " & vErrorMessage)
'                Call .WriteLine
'                Call .Close
'            End With
'        End With
'    End If
End Sub

Public Sub ExecuteTestCase( _
    ByRef vModuleName As String, _
    ByRef vCaseName As String _
)
    Select Case vModuleName
        Case Else
            Call Runtime.RaiseUndefinedTestModuleHandler
    End Select
End Sub

'''''''''''''''''''''''
'                     '
' Procedure Template: '
'                     '
'''''''''''''''''''''''

' Public [Sub | Function] ProcedureName()
'     ' Declare local variables.
'     ' TODO: Implement.

'     ' Setup error handling.
'     On Error GoTo HandleError:

'     ' Allocate resources.
'     ' TODO: Implement.

'     ' Implement the application logic.
'     ' TODO: Implement.

' Terminate:
'     ' Reset error handling.
'     On Error GoTo 0

'     ' Release all allocated resources if needed.
'     ' TODO: Implement.

'     ' Re-raise any stored error.
'     Call Runtime.ReRaiseError

'     ' Exit the procedure.
'     Exit [Sub | Function]

' HandleError:
'     ' Store the error for further handling.
'     Call Runtime.StoreError

'     ' TODO: Verify whether the error should be re-raised.

'     ' Resume to procedure termination.
'     Resume Terminate:
' End [Sub | Function]
