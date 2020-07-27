'//======================\\
'//== Information Only ==\\
'//======================\\
'   Synopsis: With SAP already open, runs transaction and inputs dates for current week
'==========================

Dim SapGui, appSAP, Connection, session, SAPTrans
Dim WeekStartDate, WeekEndDate

Select Case WeekDay(Date)
	Case 1 'Sunday
		WeekStartDate = Date
		WeekEndDate = Date + 6
	Case 2 'Monday
		WeekStartDate = Date -1
		WeekEndDate = Date + 5
	Case 3 'Tuesday
		WeekStartDate = Date -2
		WeekEndDate = Date + 4
	Case 4 'Wednesday
		WeekStartDate = Date -3
		WeekEndDate = Date + 3
	Case 5 'Thursday
		WeekStartDate = Date -4
		WeekEndDate = Date + 2
	Case 6 'Friday
		WeekStartDate = Date -5
		WeekEndDate = Date + 1
	Case 7 'Saturday
		WeekStartDate = Date - 6
		WeekEndDate = Date
End Select

SAPTrans = "YAFSHIPMENT2"

'set SAP objects
	Set SapGui = GetObject("SAPGUI")
	Set appSAP = SapGui.GetScriptingEngine

'create new session
	Set Connection = appSAP.Connections(0)
	Set session = Connection.Children(Connection.Children.Count - 1) 'must do -1 to convert to option base 0
	session.CreateSession
	wscript.sleep 2000
	'reset session variable to catch the new session enumeration
	Set session = Connection.Children(Connection.Children.Count - 1) 'must do -1 to convert to option base 0


session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = SAPTrans
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_DATE-LOW").text = WeekStartDate
session.findById("wnd[0]/usr/ctxtS_DATE-HIGH").text = WeekEndDate
session.findById("wnd[0]/usr/ctxtS_DATE-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtS_DATE-HIGH").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
