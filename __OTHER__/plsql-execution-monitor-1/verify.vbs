Option Explicit

' Configure the message box title.
vMsgBoxTitle = "PL/SQL Execution Monitor"

' Trigger a message box for each line on the standard input.
Do While Not WScript.StdIn.AtEndOfStream
	Call MsgBox(WScript.StdIn.ReadLine(), 48, vMsgBoxTitle)
Loop
