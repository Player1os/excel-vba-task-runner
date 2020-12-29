Option Explicit

' Declare local variables.
Dim vExecuteScriptFilePath
Dim vTaskName

' Determine th path to the execute script file.
With CreateObject("Scripting.FileSystemObject")
	vExecuteScriptFilePath = .BuildPath(.GetParentFolderName(.GetParentFolderName(WScript.ScriptFullName)), "Execute.vbs")
End With

' Load the wscript shell object.
With CreateObject("WScript.Shell")
	' Load the user defined parameters.
	If WScript.Arguments.Count <> 1 Then
		vTaskName = WScript.Arguments(0)
	Else
		vTaskName = InputBox("Enter the name of the task to kill")
	End If

	' Set the navigate path environment variable to the user's input.
	.Environment("PROCESS")("APP_NAVIGATE_PATH") = "kill?task_name=" & vTaskName

	' Run the execute script.
	Call .Run(vExecuteScriptFilePath, 0, False)
End With
