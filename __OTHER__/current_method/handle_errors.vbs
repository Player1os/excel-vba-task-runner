Option Explicit

' Define external object constants.
Const olMailItem = 0

Sub SendMail( _
	vRecipient, _
	vSubject, _
	vBody _
)
	' Initialize the outlook application.
	With CreateObject("Outlook.Application")
		' Log on into the session.
		Call .Session.Logon

		' Create a new mail item.
		With .CreateItem(olMailItem)
			' Set the recipient of the message.
			.To = vRecipient

			' Set the message subject and body.
			.Subject = vSubject
			.Body = vBody

			' Send the prepared message.
			Call .Send
		End With

		' Log off from the session.
		Call .Session.Logoff
	End With
End Sub

' Declare the parameter variables.
Dim vErrorLogFilePath
Dim vReportRecipientEmailAddress
Dim vTaskName

' Store the relevant environment variables and script arguments into the parameter variables.
With CreateObject("WScript.Shell").Environment("PROCESS")
	vErrorLogFilePath = .Item("APP_ERROR_LOG_FILE_PATH")
	vReportRecipientEmailAddress = .Item("APP_REPORT_RECIPIENT_EMAIL_ADDRESS")
	vTaskName = WScript.Arguments(0)
End With

' Initialize the file system object external object.
With CreateObject("Scripting.FileSystemObject")
	If (.GetFile(vErrorLogFilePath).Size > 0) Then
		Call SendMail( _
			vReportRecipientEmailAddress, _
			"[Task Runner] Error detected", _
			"The task '" & vTaskName & "' has produced a non-empty error log, located at '" & vErrorLogFilePath & "'." _
		)
	End If
End With
