'example strProcess (excel.exe, saplogon.exe)

Function KillProcess(strProcess)

	'//-- Set-up Variables --\\
	'//----------------------\\

	Dim objWMIService, objProcess, colProcess, strComputer
	strComputer = "."

	'related to WMI
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = " & strProcess)

	'//-- Perform Actions --\\
	'//---------------------\\

	For Each objProcess in colProcess
		objProcess.Terminate()
	Next


End Function
