'//======================\\
'//== Information Only ==\\
'//======================\\
'   Synopsis: Opens SAP, runs report on disputes incoming, and exports results to Excel
'==========================

KillProcess  "'saplogon.exe'" 'Closes any SAP instance if open.

'//-- Set variables --\\
'//-------------------\\

Dim objShell
Dim SapGui, appSAP, Connection, session, sessionNew, SAPtrans, SAPvariant
Dim appExcel, wb, ws, wbpath
Dim sapPath
Dim fso

DateStart = "1/1/" & Year(Date)

wbpath = " " 'Folder path you want to save excel file to
wb = ".xlsx" 'Name you want for the excel file
sapPath64 = """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe""" 'running on Win7 64bit, 32bit version
sapPath32 = """C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe""" 'running on Win7 32bit, 32bit version
SAPtrans = "UDM_DISPUTE"

Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

'Get proper previous date. Normally we use the previous day but on Monday we want all the way back to Friday.
If Weekday(Date)=2 then 'If week day is Monday
    DateEnd = Date-3
Else
    DateEnd = Date
End if


'//-- check to see if SAP is already running --\\
'// -- can't login to SAP twice ---------------\\
'----------------------------------------------\\

If IsProcessRunning("saplogon.exe") = True Then

	'set SAP objects
	Set SapGui = GetObject("SAPGUI")
	Set appSAP = SapGui.GetScriptingEngine
	Set Connection = appSAP.Connections(0)
	Set session = Connection.Children(Connection.Children.Count - 1) 'must do -1 to convert to option base 0

	'initialize the CreateSession event
	WScript.ConnectObject appSAP, "Engine_"

	'create new session
	with session
		.findbyid("wnd[0]").SendVKey 74 'Ctrl+N for create session
	end with

	'set the session variables
	wscript.sleep 8000
	Set sessionNew = Connection.Children(Connection.Children.Count - 1) 'must do -1 to convert to option base 0
	'wscript.echo Connection.Children.Count
	Set session = sessionNew

Else
	'launch SAP
	'If fso.fileexists(sapPath64) then
		objShell.Run sapPath64
	'Else
		'objShell.Run sapPath32
	'End if

	wscript.Sleep 5000

	'set SAP objects
	Set SapGui = GetObject("SAPGUI")
	Set appSAP = SapGui.GetScriptingEngine
	Set Connection = appSAP.OpenConnection("-  ARP [Aero_Prod_ERP]",True)
	Set session = Connection.Children(0)

End If

'SAP screen inputs
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = SAPtrans
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell/shellcont[1]/shell").doubleClickNode "          2" 'Double click the Find Dispute Node
If Weekday(Date)=2 then 'If weekday is Monday search last Friday through Sunday, else just search previous day.
    session.findById("wnd[0]/usr/cntlCLFRM_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont/shell").setCurrentCell 9,"SEL_ICON2"
    session.findById("wnd[0]/usr/cntlCLFRM_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont/shell").pressButtonCurrentCell 'Press the multiple selections button
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL").select 'Select "Select Ranges" tab
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-ILOW_I[1,0]").text = Date-3 'Enter earliest date
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").text = Date-1 'Enter prevous date
    session.findById("wnd[1]/tbar[0]/btn[8]").press 'Click the execute to button to accept date range
Else
    session.findById("wnd[0]/usr/cntlCLFRM_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont/shell").modifyCell 9,"VALUE2",Date-1 'Enter previous day
End if
session.findById("wnd[0]/usr/cntlCLFRM_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont/shell").modifyCell 12,"VALUE1","10000" 'Change Restrict Hits to 10,000
session.findById("wnd[0]/usr/cntlCLFRM_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell/shellcont[1]/shell").pressButton "DO_QUERY" 'Cick the Search button

'Exporting
session.findById("wnd[0]/usr/cntlCLFRM_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").contextMenu
session.findById("wnd[0]/usr/cntlCLFRM_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/radRB_OTHERS").setFocus
session.findById("wnd[1]/usr/radRB_OTHERS").select
session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus
session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "08"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select 'Select table
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press

wscript.sleep 5000

'Excel Procedure
objShell.AppActivate("Microsoft Excel") 						 'Worksheet in "some default name here" (1)
Set appExcel = GetObject(,"Excel.Application")
With appExcel
	.DisplayAlerts = False
	.Workbooks(appExcel.Workbooks.Count).saveas wbpath & wb
	.Quit
End with

wscript.sleep 5000  'give it time to save before clicking the sap prompt

'SAP close-out procedure
with session
	'.findById("wnd[1]/tbar[0]/btn[0]").press 			         'prompt that reminds me to save excel - must be clicked **ONLY IF** appExcel.quit not done above
	'.findById("wnd[0]/tbar[0]/btn[15]").press				     'yellow "exit" button to move back one
	.findById("wnd[0]/tbar[0]/btn[15]").press                    'yellow "exit" button to move back one - back @ main menu
	'.findById("wnd[0]/mbar/menu[4]/menu[12]").select             'System -> Log off
	'.findById("wnd[1]/usr/btnSPOP-OPTION1").press                'confirm log off
end with


'Deallocate variables
Set objShell = Nothing
Set fso = Nothing
Set SapGui = Nothing
Set appSAP = Nothing
Set Connection = Nothing
Set session = Nothing
Set appExcel = Nothing
Set sessionNew = Nothing

KillProcess  "'saplogon.exe'"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function IsProcessRunning(strProcess)
'http://stackoverflow.com/questions/19794726/vb-script-how-to-tell-if-a-program-is-already-running
    Dim Process, strObject, strComputer
    IsProcessRunning = False
	strComputer = "."
    strObject   = "winmgmts://" & strComputer
    For Each Process in GetObject(strObject).InstancesOf("win32_process")
    If UCase(Process.name) = UCase(strProcess) Then
        IsProcessRunning = True
        Exit Function
    End If
    Next
End Function
