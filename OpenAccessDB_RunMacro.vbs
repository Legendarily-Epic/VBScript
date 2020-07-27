'//======================\\
'//== Information Only ==\\
'//======================\\
'   Synopsis: Starts up the Database & associated scripts
'==========================


'//-- Set-up Variables --\\
'//----------------------\\

Dim appAccess
Dim objShell
Dim timestart, timeend

Set appAccess = CreateObject("Access.Application")
Set objShell = CreateObject("Wscript.Shell")

strSource = ".accdb" 'path location of access database

timestart = time


'//--Run Sub that Runs Daily Update Macro\\
'//--------------------------------------\\

with appAccess
	.OpenCurrentDatabase strSource
	.visible = true
	.run "DashboardRefresh" 'name of sub created in db
	.quit
End With

timeend = time
objShell.popup "OTTR Dashboard Refresh: " & vbCRLF & _
			  "From start to finish, total elapsed time was " & datediff("s",timestart, timeend)/60 & " minutes.",,,1+4096


'//--Deallocate variables\\
'//----------------------\\

Set appAccess = Nothing
Set objShell = Nothing
Set timestart = Nothing
Set timeend = Nothing
