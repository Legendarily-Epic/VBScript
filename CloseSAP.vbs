'//======================\\
'//== Information Only ==\\
'//======================\\
'   Synopsis: Kills the identified process on local machine
'==========================


'//-- Set-up Variables --\\
'//----------------------\\
Dim objShell, intCountdown, intInput
Set objShell = CreateObject("wscript.shell")
intCountdown = 15

'//-- Perform Actions --\\
'//---------------------\\

Select Case objShell.Popup("Warning!" & vbCrlf & vbCrlf & "All active sessions of SAP will be terminated. " _
						    & "Save your work now to prevent data loss. " _
							& "Would you like to proceed? " _
							& "No response will be assumed ""Yes"" in " & intCountdown & " seconds.",intCountdown,"" _
							& "Title",48+4+4096) '48=excl,4=yes/no,4096=TopWindow
	Case 6 'Yes
		KillProcess "'saplogon.exe'"
	Case 7 'No
	Case -1 'Timed-out
		KillProcess "'saplogon.exe'"
End Select



'//==========================\\
'//== Supporting Functions ==\\
'//==========================\\

Function KillProcess(strProcessKill)

	'//-- Set-up Variables --\\
	'//----------------------\\

	Dim objWMIService, objProcess, colProcess, strComputer
	Dim objShell, strMachine, strUser
	Dim fso, strDirectory
	'strProcessKill = "'saplogon.exe'"
	strComputer = "."

	'related to the shell
	Set objShell = CreateObject("wscript.shell")
	strMachine = objShell.ExpandEnvironmentStrings("%ComputerName%")
	strUser = objShell.ExpandEnvironmentStrings("%UserName%")

	'related to WMI
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = " & strProcessKill)

	'related to fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	strDirectory = "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
	If fso.FileExists(strDirectory) then
		strDirectory = strDirectory  '32-bit version
	Else
		strDirectory = "C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe"  '64-bit version
	End If

	'//-- Perform Actions --\\
	'//---------------------\\

	For Each objProcess in colProcess
		objProcess.Terminate()
	Next

	objShell.Popup "Previously running instances of " & strProcessKill & " are now terminated. " & vbCrlf & vbCrlf _
				   & "Machine: " & strMachine & vbCrlf _
				   & "User: " & UCase(strUser) & vbCrlf _
				   & "Directory: " & strDirectory,5,,48+4096


	'//--Deallocate objects--\\
	'//----------------------\\

	Set objWMIService = Nothing
	Set objProcess = Nothing
	Set colProcess = Nothing
	Set objShell = Nothing
	Set fso = Nothing

End Function
