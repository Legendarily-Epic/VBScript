'//--Information Only--\\
'//--------------------\\
'   Synopsis: Exports qry to Excel to e-mail stakeholders
'   Ext Refs: Microsoft Excel 12.0 Object Library


'//--Declare variables & set objects--\\
'//-----------------------------------\\

Dim db As Database, rs1 As Recordset
Dim ws As Worksheet, wb As Workbook
Dim appExcel As New Excel.Application
Dim saveroot As String, save1 As String
Dim appOutlook As New Outlook.Application, objMailItem As MailItem
Dim arrStakeholders() As String
Dim arrStakeholderscc() As String


Set db = CurrentDb
Set appExcel = CreateObject("Excel.Application")
Set rs1 = db.OpenRecordset("tbl_01") 'query you want to use

saveroot = "" 'add path to folder location you want to save the excel file to
save1 = saveroot & ".xlsx" 'add the name of the file you wish to create


'//--Create Excel instance & copy records--\\
'//-------------- rs1 ---------------------\\
'//----------------------------------------\\

With appExcel
.Visible = True
.DisplayAlerts = False
.Workbooks.Add
End With

'can't set these until workbook is added b/c you can't count/set something that does not exist
Set wb = appExcel.Workbooks(appExcel.Workbooks.Count)
Set ws = wb.Worksheets(appExcel.Sheets.Count)

For i = 1 To rs1.Fields.Count
    ws.Cells(1, i) = rs1.Fields(i - 1).Name
Next

With ws
  .Activate
  .Cells(2, 1).CopyFromRecordset rs1
  .Columns.AutoFit
  .SaveAs save1
End With

appExcel.Quit


'//--Send rs1 Email to StakeHolders--\\
'//----------------------------------\\

ReDim arrStakeholders(1)
    arrStakeholders(0) = "1@mail.com;"
    arrStakeholders(1) = "2@mail.com;"


ReDim arrStakeholderscc(1)
    arrStakeholderscc(0) = "3@mail.com;"
    arrStakeholderscc(0) = "4@mail.com;"


Set appOutlook = CreateObject("Outlook.Application")
Set objMailItem = appOutlook.CreateItem(olMailItem)

With objMailItem
    .Subject = " " & Now 'add subject of email
    .Attachments.Add save1
    .To = Join(arrStakeholders)
    .CC = Join(arrStakeholderscc)
    .Body = "This is an automated message. Attached is " 'elaborate further on what you are sending, why, and any further action required
    .Send
End With


'//--Deallocate Objects--\\
'//----------------------\\

Set db = Nothing
Set rs1 = Nothing
Set appExcel = Nothing
Set ws = Nothing
Set wb = Nothing
Set appOutlook = Nothing
Set objMailItem = Nothing

End Function
