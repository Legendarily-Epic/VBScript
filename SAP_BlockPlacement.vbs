'//======================\\
'//== Information Only ==\\
'//======================\\
'   Synopsis: Opens SAP and OE/SP price deviation flag file and blocks items/schedule lines in SAP. Write back status to excel file, then sends file via email to stakeholders.
'==========================


'//-- Set variables --\\
'//-------------------\\

Dim objShell
Dim SapGui, appSAP, Connection, session, sessionNew, SAPtrans, SAPvariant
Dim appExcel, wb, ws, wbpath
Dim sapPath
Dim fso

sapPath64 = """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe""" 'running on Win7 64bit, 32bit version
sapPath32 = """C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe""" 'running on Win7 32bit, 32bit version

Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

'Ensures SAP is closed. Easier if it just closes and starts new session.
KillProcess "'saplogon.exe'"


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

	'Set connection - select connection name as naming can be different depending on GUI
	Set Connection = appSAP.OpenConnection("-  ARP [Aero_Prod_ERP]",True)
	'Set Connection = appSAP.OpenConnection("- ARP - Aero_Prod_ERP",True)
    'Set Connection = appSAP.OpenConnection("ARV - VCCT",True) 'Used for testing
	Set session = Connection.Children(0)

End If


'Open Excel file
Set appExcel = CreateObject("Excel.Application")
appExcel.Visible = True
appExcel.Workbooks.Open ".xlsx" 'Add path a file name for that which you are trying to open
appExcel.Visible = True
appExcel.DisplayAlerts = False
appExcel.Worksheets("Block").Activate 'Choose the sheet in the file that you wish to activate

Set wb = appExcel.ActiveWorkbook
Set ws = wb.Sheets("Block") 'Choose the sheet in the file that you wish to activate

'Declare variables
Dim ExcelRowNum
Dim SalesOrderStatus
Dim SalesOrder
Dim SalesOrderItem
Dim SalesOrderScheduleLine
Dim ScrollbarPosition
Dim FlagReason

'Creates new column headers in excel file
wb.ActiveSheet.Cells(1, 4).Value = "Macro Sales Order Status"
wb.ActiveSheet.Cells(1, 5).Value = "Pre-existing Block Status"
wb.ActiveSheet.Cells(1, 6).Value = "Aged Item Status"
wb.ActiveSheet.Cells(1, 7).Value = "Text Notes"

'Assign starting variable
ExcelRowNum = 2 'Row 1 is header
SalesOrderStatus = wb.ActiveSheet.Cells(ExcelRowNum, 4)

'Searches for a line item that doesn't have a status yet (macro hasn't completed). This is done in case you have to abort macro manually. I can restart macro where it left off in the file instead starting from the begining
Do Until SalesOrderStatus = ""
    ExcelRowNum = ExcelRowNum + 1
    SalesOrderStatus = wb.ActiveSheet.Cells(ExcelRowNum, 4)
Loop

'Assign rest of variables now that first row to work is known
SalesOrder = wb.ActiveSheet.Cells(ExcelRowNum, 1)
SalesOrderItem = wb.ActiveSheet.Cells(ExcelRowNum, 2)
SalesOrderScheduleLine = 0
ScrollbarPosition = 0

'Enters VA02 - Sales Order change
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "VA02" 'transaction T-code box
session.findById("wnd[0]").sendVKey 0 'Enter

'Begin Loop of each notification number
Do While SalesOrder <> ""
On Error Resume Next

    'Enter Sales Order
    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = SalesOrder 'Enters the sales number
    session.findById("wnd[0]").sendVKey 0 'Presses Enter

    'If error when trying to get into Sales Order (like no authorization) then will log the error
    If session.findById("wnd[0]/sbar").Text <> "" and session.findById("wnd[0]/sbar").Text <> "Consider the subsequent documents" then
		wb.ActiveSheet.Cells(ExcelRowNum, 4).Value = session.findById("wnd[0]/sbar").Text
        session.findById("wnd[0]").sendVKey 0 'Press Enter through any warning message
    Else

        'Sometimes a window pops up, usually an order block window, this will just close it if pops up to continue
        If not session.findById("wnd[1]/tbar[0]/btn[0]", False) is nothing then
            session.findById("wnd[1]/tbar[0]/btn[0]").press
        End if

        'Enter Sales Order Item
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4054/btnBT_POPO").press 'Press item to top button
        session.findById("wnd[1]/usr/txtRV45A-POSNR").text = SalesOrderItem 'Enter Sales Order Item
        session.findById("wnd[1]/tbar[0]/btn[0]").press 'Press Green checkmark button
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,0]").setFocus 'Set Focus on top item
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,0]").caretPosition = 5 'Place cursor in item number
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4054/btnBT_PEIN").press 'Press the Schedule lines for item button

        'Place block on all available schedule lines
        For row = 0 to (session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN").VerticalScrollbar.Maximum)-1 'Only goes through available lines in SAP table
            If Year(session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN/ctxtRV45A-ETDAT[1," & SalesOrderScheduleLine & "]").text) > Year(Date) then 'Only place block on schedule lines for this year
                'Handles scrolling for greater than 20 schedule lines
                If SalesOrderScheduleLine = 21 then
                    ScrollbarPosition = ScrollbarPosition + 21
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN").verticalScrollbar.position = ScrollbarPosition
                    SalesOrderScheduleLine = 0
                End if
                SalesOrderScheduleLine = SalesOrderScheduleLine + 1
            Else
                'Handles scrolling for greater than 20 schedule lines
                If SalesOrderScheduleLine = 21 then
                    ScrollbarPosition = ScrollbarPosition + 21
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN").verticalScrollbar.position = ScrollbarPosition
                    SalesOrderScheduleLine = 0
                End if

                'Places block or writes back block that already exists
                If session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN/cmbVBEP-LIFSP[6," & SalesOrderScheduleLine & "]").key <> " " and session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN/txtVBEP-WMENG[2," & SalesOrderScheduleLine & "]").text <> "0" Then 'If Block already exists
                    wb.ActiveSheet.Cells(ExcelRowNum, 5).Value = wb.ActiveSheet.Cells(ExcelRowNum, 5).Value & "L" & SalesOrderScheduleLine+1 & "-" & session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN/cmbVBEP-LIFSP[6," & SalesOrderScheduleLine & "]").key & ":" 'Write line number and block key back to excel file
                Else
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN/cmbVBEP-LIFSP[6," & SalesOrderScheduleLine & "]").key = "ZJ" 'Set ZJ Delivery Block
                End If
                SalesOrderScheduleLine = SalesOrderScheduleLine + 1
            End If
        Next

        SalesOrderScheduleLine = 0
        ScrollbarPosition = 0
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Press Green Back button
        'Sometimes a window pops up, "No valid Storage Location for Valuation Type NEW", this will just close it if pops up to continue.
        If not session.findById("wnd[1]/tbar[0]/btn[0]", False) is nothing then
            session.findById("wnd[1]/tbar[0]/btn[0]").press
        End if
        wb.ActiveSheet.Cells(ExcelRowNum, 4).Value = "Blocks placed or already in place"

        'Repeat for other items on same Sales Order
        Do While wb.ActiveSheet.Cells((ExcelRowNum + 1), 1) = SalesOrder 'If sales order matches next sales order then must be same sales order just different item. Want to update all items before saving sales order.
            ExcelRowNum = ExcelRowNum + 1
            SalesOrderItem = wb.ActiveSheet.Cells(ExcelRowNum, 2)

            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4054/btnBT_POPO").press 'Press item to top button
            session.findById("wnd[1]/usr/txtRV45A-POSNR").text = SalesOrderItem 'Enter Sales Order Item
            session.findById("wnd[1]/tbar[0]/btn[0]").press 'Press Green checkmark button
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,0]").setFocus 'Set Focus on top item
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,0]").caretPosition = 5 'Place cursor in item number
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4054/btnBT_PEIN").press 'Press the Schedule lines for item button
            'Place block on all available schedule lines
            For row = 0 to (session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN").VerticalScrollbar.Maximum)-1 'Only goes through available lines in SAP table
                If Year(session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN/ctxtRV45A-ETDAT[1," & SalesOrderScheduleLine & "]").text) > Year(Date) then 'Only place block on schedule lines for this year
                    'Handles scrolling for greater than 20 schedule lines
                    If SalesOrderScheduleLine = 21 then
                        ScrollbarPosition = ScrollbarPosition + 21
                        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN").verticalScrollbar.position = ScrollbarPosition
                        SalesOrderScheduleLine = 0
                    End if
                    SalesOrderScheduleLine = SalesOrderScheduleLine + 1
                Else
                    'Handles scrolling for greater than 20 schedule lines
                    If SalesOrderScheduleLine = 21 then
                        ScrollbarPosition = ScrollbarPosition + 21
                        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN").verticalScrollbar.position = ScrollbarPosition
                        SalesOrderScheduleLine = 0
                    End if

                    'Places block or writes back block that already exists
                    If session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN/cmbVBEP-LIFSP[6," & SalesOrderScheduleLine & "]").key <> " " and session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN/txtVBEP-WMENG[2," & SalesOrderScheduleLine & "]").text <> "0" Then 'If Block already exists
                        wb.ActiveSheet.Cells(ExcelRowNum, 5).Value = wb.ActiveSheet.Cells(ExcelRowNum, 5).Value & "L" & SalesOrderScheduleLine+1 & "-" & session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN/cmbVBEP-LIFSP[6," & SalesOrderScheduleLine & "]").key & ":"
                    Else
                        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/tblSAPMV45ATCTRL_PEIN/cmbVBEP-LIFSP[6," & SalesOrderScheduleLine & "]").key = "ZJ" 'Set ZJ Delivery Block
                    End If
                    SalesOrderScheduleLine = SalesOrderScheduleLine + 1
                End If
            Next

            SalesOrderScheduleLine = 0
            ScrollbarPosition = 0
            session.findById("wnd[0]/tbar[0]/btn[3]").press 'Press Green Back button
            'Sometimes a window pops up, "No valid Storage Location for Valuation Type NEW", this will just close it if pops up to continue.
            If not session.findById("wnd[1]/tbar[0]/btn[0]", False) is nothing then
                session.findById("wnd[1]/tbar[0]/btn[0]").press
            End if
            wb.ActiveSheet.Cells(ExcelRowNum, 4).Value = "Blocks placed or already in place"
        Loop

        'Enter text notes
        session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press 'Presses the "Display Header Doc Details" button
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10").select 'Selects the "Texts" tab
        'If Text notes are formatted then there are different steps in order to enter notes.
        If session.findById("wnd[0]/sbar").Text = "Text is formatted -> Details" Then
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/btnTP_DETAIL").press 'Press Detail Button
            session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,1]").setFocus
            session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,1]").caretPosition = 1
            session.findById("wnd[0]/tbar[1]/btn[5]").press 'Press Insert Button
            session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,1]").text = "Order flagged by Price Deviation Tool (PDT) for " 'Enter Text
            session.findById("wnd[0]").sendVKey 0 'Press Enter to go to next line
            session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,2]").text = wb.ActiveSheet.Cells(ExcelRowNum, 3).Value 'Enter Next line text
            session.findById("wnd[0]").sendVKey 0 'Press Enter to go to next line
            session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,3]").text = "on " & Date & " and ZJ Blocks placed. " 'Enter Next line text
            session.findById("wnd[0]").sendVKey 0 'Press Enter to go to next line
            session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,4]").text = "Contact Nicole Ortiz (H242630) for any questions." 'Enter Next line text
            session.findById("wnd[0]/tbar[1]/btn[5]").press 'Press End Insertion button
            session.findById("wnd[0]/tbar[0]/btn[3]").press 'Press green arrow back
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        Else
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem "0002","Column1" 'Selects the Internal Notes Node
            'Checks to see if text notes for same status already exists so we don't duplicate notes
            If InStr(session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text, wb.ActiveSheet.Cells(ExcelRowNum, 3)) = 0 then
                'Checks to see if Notes exists or not. If not then append notes, if yes then add a couple returns to give new note space below previous note.
                If session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text <> "" + vbCr + "" then
                    PreviousTexts = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text 'Copies previous texts
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = "Order flagged by Price Deviation Tool (PDT) for " & wb.ActiveSheet.Cells(ExcelRowNum, 3).Value & " on " & Date & " and ZJ Cust Info Required block(s) placed. Contact Nicole Ortiz (H242630) for any questions." & vbCr & vbCr & PreviousTexts 'Appends ZJ placment notes to previous texts
                    PreviousTexts = ""
                Else
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes 0,0
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = "Order flagged by Price Deviation Tool (PDT) for " & wb.ActiveSheet.Cells(ExcelRowNum, 3).Value & " on " & Date & " and ZJ Cust Info Required block(s) placed. Contact Nicole Ortiz (H242630) for any questions." 'Appends ZJ placment notes
                End IF
            Else
                wb.ActiveSheet.Cells(ExcelRowNum, 6).Value = "Notes already exists"
                wb.ActiveSheet.Cells(ExcelRowNum, 7).Value = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text 'Write back text notes to excel file
            End If
        End IF
        'Save Sales Order
        session.findById("wnd[0]/tbar[0]/btn[11]").press 'Press the save button

        'Sometimes a window pops up, changes, this will press continue to save the sales order.
        If not session.findById("wnd[1]/usr/btnCONTINUE", False) is nothing then
            session.findById("wnd[1]/usr/btnCONTINUE").press
        End if

        'Sometimes a window pops up, save incomplete document, this will just close it if pops up to continue.
        If not session.findById("wnd[1]/usr/btnSPOP-VAROPTION1", False) is nothing then
            session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
        End if

        'Sometimes a window pops up, "Credit Check SAP Credit Management Failed", this will just close it if pops up to continue.
        If not session.findById("wnd[1]/tbar[0]/btn[0]", False) is nothing then
            session.findById("wnd[1]/tbar[0]/btn[0]").press
        End if

        'Sometimes a window pops up, "Reference Number of a Document from back Error", this will just close it if pops up to continue.
        If not session.findById("wnd[1]/tbar[0]/btn[0]", False) is nothing then
            session.findById("wnd[1]/tbar[0]/btn[0]").press
        End if

        'Sometimes a window pops up, save incomplete document, this will just close it if pops up to continue.
        If not session.findById("wnd[1]/usr/btnSPOP-VAROPTION1", False) is nothing then
            session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
        End if

        'Sometimes a window pops up, "Order receipt/delivery not possible, credit customer blocked", this will just close it if pops up to continue.
        If session.findById("wnd[0]/sbar").Text = "Order receipt/delivery not possible, credit customer blocked" then
            errMsg = session.findById("wnd[0]/sbar").Text
            'session.findById("wnd[0]").sendVKey 0 'Presses Enter to select the "OK" button
            session.findById("wnd[0]/tbar[0]/btn[3]").press 'Press green back button to get out of header details
            session.findById("wnd[0]/tbar[0]/btn[3]").press 'Press green back button to return to VA02 main menu/sales order input
            session.findById("wnd[1]/usr/btnSPOP-OPTION2").press 'Chooses to not save when with pop-up
        End if

        'Sometimes a window pops up, "Negative net value is not permitted on line item 1000", this will just close it if pops up to continue.
        If session.findById("wnd[0]/sbar").Text = "Negative net value is not permitted on line item 1000" then
            errMsg = session.findById("wnd[0]/sbar").Text
            session.findById("wnd[1]/usr/btnSPOP-OPTION2").press 'Presses Enter to select the "OK" button
            session.findById("wnd[0]/tbar[0]/btn[3]").press 'Press green back button to get out of header details
            session.findById("wnd[0]/tbar[0]/btn[3]").press 'Press green back button to return to VA02 main menu/sales order input
            session.findById("wnd[1]/usr/btnSPOP-OPTION2").press 'Chooses to not save when with pop-up
        End if

        If errMsg = "" then
            wb.ActiveSheet.Cells(ExcelRowNum, 4).Value = session.findById("wnd[0]/sbar").Text
        Else
            wb.ActiveSheet.Cells(ExcelRowNum, 4).Value = errMsg
        End If

    End IF

    ExcelRowNum = ExcelRowNum + 1
    SalesOrder = wb.ActiveSheet.Cells(ExcelRowNum, 1)
    SalesOrderItem = wb.ActiveSheet.Cells(ExcelRowNum, 2)
    SalesOrderScheduleLine = 0
    ScrollbarPosition = 0

Loop

    'Excel close-out procedure
    wb.Save
    wb.Close
    appExcel.Quit

    'SAP close-out procedure
    ' with session
        ' .findById("wnd[0]/tbar[0]/btn[15]").press		     'yellow "exit" button to move back one
        ' .findById("wnd[0]/tbar[0]/btn[15]").press            'yellow "exit" button to move back one - back @ main menu
        ' .findById("wnd[0]/mbar/menu[4]/menu[12]").select     'System -> Log off
        ' .findById("wnd[1]/usr/btnSPOP-OPTION1").press        'confirm log off
    ' end with

    KillProcess "'saplogon.exe'"

    EmailStakeholders 'Sends the automated email that the macro has completed and attaches the file for review

Set appExcel = Nothing
Set wb = Nothing
Set ws = Nothing
Set SapGuiAuto = Nothing
Set application = Nothing
Set connection = Nothing
Set session = Nothing


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'//==========================\\
'//== Supporting Functions ==\\
'//==========================\\

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function EmailStakeholders()

Dim rpt01
'Dim rpt02

rpt01 = ".xlsx" 'Add path a file name for that which you are trying to open
'rpt02 = ""


Dim appOutlook
Dim objMailItem
Dim arrStakeholders()
Dim arrStakeholderscc()

ReDim arrStakeholders(6)
    arrStakeholders(0) = "1@mail.com;"
    arrStakeholders(1) = "2@mail.com;"
    arrStakeholders(2) = ""
    arrStakeholders(3) = ""
    arrStakeholders(4) = ""
    arrStakeholders(5) = ""
    ' arrStakeholders(6) = ";"
    ' arrStakeholders(7) = ";"
    ' arrStakeholders(8) = ";"
    ' arrStakeholders(9) = ";"
    ' arrStakeholders(10) = ";"
    ' arrStakeholders(11) = ";"
    'arrStakeholders(12) = ";"
    'arrStakeholders(13) = ";"
    'arrStakeholders(14) = ";"


ReDim arrStakeholderscc(6)
    arrStakeholderscc(0) = "3@mail.com;"
    arrStakeholderscc(1) = "4@mail.com;"
    arrStakeholderscc(2) = ""
    arrStakeholderscc(3) = ""
    arrStakeholderscc(4) = ""
    arrStakeholderscc(5) = ""
    'arrStakeholderscc(6) = ";"
    'arrStakeholderscc(7) = ""
    'arrStakeholderscc(8) = ""
    'arrStakeholderscc(9) = ""
    'arrStakeholderscc(10) = ""


Set appOutlook = CreateObject("Outlook.Application")
Set objMailItem = appOutlook.CreateItem(olMailItem)

With objMailItem
    .SentOnBehalfOfName = "BizOps@mail.com"
    .Subject = "OE/SP PDT auto block macro has completed " & Now 'Enter your own subject as example provided
    .Attachments.Add rpt01
    '.Attachments.Add rpt02
    .To = Join(arrStakeholders)
    .CC = Join(arrStakeholderscc)
    'Add you own body message below. Example has been provided.
    .Body = "***This is an automated message***" & vbCr & vbCr & "The OE/SP PDT auto block macro has completed. Attached is todays file the macro ran against and updated for review"
    .Send
End With


End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
