Option Explicit

' Define external object constants.
Const adSaveCreateOverWrite = 2
Const adTypeBinary = 1
Const adTypeText = 2

' Declare the string decode function.
Function Decode( _
	vValue _
)
	' Declare local variables.
	Dim vTemporaryXmlElement

	' Create a virtual xml element with the binary base64 datatype.
	Set vTemporaryXmlElement = CreateObject("MSXML2.DOMDocument.6.0").CreateElement("base64")
	vTemporaryXmlElement.DataType = "bin.base64"
	vTemporaryXmlElement.Text = vValue

	' Create and open a binary stream.
	With CreateObject("ADODB.Stream")
		.Type = adTypeBinary
		Call .Open

		' Write the binary data into the stream.
		Call .Write(vTemporaryXmlElement.nodeTypedValue)

		' Reset the binary stream to the begining and transform it into a text stream with the utf-8 charset.
		.Position = 0
		.Type = adTypeText
		.Charset = "UTF-8"

		' Read the stream's content into the result.
		Decode = .ReadText()

		' Close the stream.
		Call .Close
	End With
End Function

' Declare the file write function.
Sub WriteFile( _
	vPath, _
	vText _
)
	' Declare local variables.
	Dim vFileStream

	' Configure the file stream as a binary stream.
	Set vFileStream = CreateObject("ADODB.Stream")
	With vFileStream
		.Type = adTypeBinary
		Call .Open
	End With

	' Initialize a temporary stream.
	With CreateObject("ADODB.Stream")
		' Configure and open the temporary stream for text data using the utf-8 charset.
		.Type = adTypeText
		.Charset = "UTF-8"
		Call .Open

		' Write the specified text into the temporary stream.
		Call .WriteText(vText)

		' Reset the temporary stream to work with binary data and skip the created byte order mark.
		.Position = 0
		.Type = adTypeBinary
		.Position = 3

		' Copy the contents of the temporary stream from its current position to the file stream.
		Call .CopyTo(vFileStream)

		' Close the temporary stream.
		Call .Close
	End With

	' Load the file stream.
	With vFileStream
		' Save the contents of the file stream into a file located at the specified path.
		Call .SaveToFile(vPath, adSaveCreateOverWrite)

		' Close the stream.
		Call .Close
	End With

	' Exit the procedure.
	Exit Sub
End Sub

' Declare the parameter variables.
Dim vTempDirectoryPath
Dim vDatabaseConnectionDirectoryPath
Dim vDatabaseConnectionConfigurationName

' Store the relevant environment variables and script arguments into the parameter variables.
With CreateObject("WScript.Shell").Environment("PROCESS")
	vTempDirectoryPath = .Item("APP_TEMP_DIRECTORY_PATH")
	vDatabaseConnectionDirectoryPath = .Item("APP_DATABASE_CONNECTION_DIRECTORY_PATH")
	vDatabaseConnectionConfigurationName = WScript.Arguments(0)
End With

' Declare global variables.
Dim vFileSystemObject
Dim vPasswordFilePath

' Initialize the file system object external object.
Set vFileSystemObject = CreateObject("Scripting.FileSystemObject")

' Create an xml dom document.
With CreateObject("MSXML2.DOMDocument.6.0")
	' Configure to load files asynchronously.
	.async = False

	' Load the database connection configuration xml file with the submitted name.
	Call .load(vFileSystemObject.BuildPath(vDatabaseConnectionDirectoryPath, vDatabaseConnectionConfigurationName & ".xml"))

	' Load the root xml node.
	With .selectSingleNode("database-connection")
		' Decode and store the clear text password in a temporary text file.
		vPasswordFilePath = vFileSystemObject.BuildPath(vTempDirectoryPath, "sql_plus_password")
		Call WriteFile(vPasswordFilePath, Decode(.selectSingleNode("encoded-password").Text) & vbCrLf)

		' Generate the sql-plus parameters batch script.
		Call WriteFile( _
			vFileSystemObject.BuildPath(vTempDirectoryPath, "sql_plus_parameters.bat"), _
			"set SQL_PLUS_DATABASE=" & .selectSingleNode("database-name").Text & vbCrLf _
			& "set SQL_PLUS_USERNAME=" & .selectSingleNode("username").Text & vbCrLf _
			& "set SQL_PLUS_PASSWORD_FILE_PATH=" & vPasswordFilePath & vbCrLf _
			& "set NLS_LANG=AMERICAN_AMERICA.AL32UTF8" _
		)
	End With
End With
