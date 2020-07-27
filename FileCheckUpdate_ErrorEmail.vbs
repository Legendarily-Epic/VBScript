
''''Checking data source files from BO4 ran successfully''''

Set fs = CreateObject("Scripting.FileSystemObject")

'Stationary Orders BO4 Reports
Set NonRepairs_File = fs.GetFile(".xlsx")
Set Repairs_File = fs.GetFile(".xlsx")
Set NonRepairs_BackupFile = fs.GetFile(".xlsx")
Set Repairs_BackupFile = fs.GetFile(".xlsx")

Dim Error_Message
Error_Message=""

'Check to see if files were updated and can be copied to Shared Drive location for tableau extract updates.

'NonRepairs - 1st checks weekly 9p schedule, if not there then it goes to the normal schedue location and sees if an update from 8PM (2000) ran successfully.
If FormatDateTime(NonRepairs_File.DateLastModified,2) <> FormatDateTime(Date,2) THEN
	If Day(NonRepairs_BackupFile.DateLastModified) & Hour(NonRepairs_BackupFile.DateLastModified) <> Day(Date) & 20 Then
		Error_Message = Error_Message & vbCr & vbCr & "NonRepairs failed to refresh sucessfully out of BO4"
	Else
		NonRepairs_BackupFile.Copy (".xlsx")
	End If
Else
    NonRepairs_File.Copy (".xlsx")
End If

'Repairs - 1st checks weekly 9p schedule, if not there then it goes to the normal schedue location and sees if an update from 8PM (2000) ran successfully.
If FormatDateTime(Repairs_File.DateLastModified,2) <> FormatDateTime(Date,2) Then
	If Day(Repairs_BackupFile.DateLastModified) & Hour(Repairs_BackupFile.DateLastModified) <> Day(Date)& 20 Then
		Error_Message = Error_Message & vbCr & vbCr & "Repairs failed to refresh sucessfully out of BO4"
	Else
		Repairs_BackupFile.Copy (".xlsx")
	End If
Else
    Repairs_File.Copy (".xlsx")
End If


If Error_Message <>"" Then
    Email_Error_Message
Else
	Email_Success
End If


Set fs = Nothing
Set NonRepairs_File = Nothing
Set Repairs_File = Nothing
Set NonRepairs_BackupFile = Nothing
Set Repairs_BackupFile = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function Email_Error_Message()

Dim appOutlook
Dim objMailItem
Dim arrStakeholders()
Dim arrStakeholderscc()

ReDim arrStakeholders(6)
    arrStakeholders(0) = "1@mail.com;"
    arrStakeholders(1) = ""
    arrStakeholders(2) = ""
    arrStakeholders(3) = ""
    arrStakeholders(4) = ""
    arrStakeholders(5) = ""
    'arrStakeholders(6) = ";"
    'arrStakeholders(7) = ";"
    'arrStakeholders(8) = ";"
    'arrStakeholders(9) = ";"
    'arrStakeholders(10) = ";"
    'arrStakeholders(11) = ";"
    'arrStakeholders(12) = ";"
    'arrStakeholders(13) = ";"
    'arrStakeholders(14) = ";"

Set appOutlook = CreateObject("Outlook.Application")
Set objMailItem = appOutlook.CreateItem(olMailItem)

With objMailItem
    .SentOnBehalfOfName = "BizOps@mail.com" 'added to send from group box if you have permissions
    .Subject = "FAILURE: ASP EOW BO4 failure " & Now
    .To = Join(arrStakeholders)
    .Body = "***This is an automated message***" & vbCr & vbCr & Error_Message
    .Send
End With


Set appOutlook = Nothing
Set objMailItem = Nothing


End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function Email_Success()

Set fs = CreateObject("Scripting.FileSystemObject")

'Stationary Orders BO4 Reports
Set NonRepairs_Final_File = fs.GetFile(".xlsx")
Set Repairs_Final_File = fs.GetFile(".xlsx")

Dim appOutlook
Dim objMailItem
Dim arrStakeholders()

ReDim arrStakeholders(6)
    arrStakeholders(0) = "1@mail.com;"
    arrStakeholders(1) = ""
    arrStakeholders(2) = ""
    arrStakeholders(3) = ""
    arrStakeholders(4) = ""
    arrStakeholders(5) = ""
    'arrStakeholders(6) = ";"
    'arrStakeholders(7) = ";"
    'arrStakeholders(8) = ";"
    'arrStakeholders(9) = ";"
    'arrStakeholders(10) = ";"
    'arrStakeholders(11) = ";"
    'arrStakeholders(12) = ";"
    'arrStakeholders(13) = ";"
    'arrStakeholders(14) = ";"


Set appOutlook = CreateObject("Outlook.Application")
Set objMailItem = appOutlook.CreateItem(olMailItem)

With objMailItem
    .SentOnBehalfOfName = "BizOps@mail.com"
    .Subject = "SUCCESS: ASP EOW files have been successfully updated and moved " & Now
    .To = Join(arrStakeholders)
    .Body = "***This is an automated message***" & vbCr & vbCr & "EOW ASP files have been sucessfully updated and moved."& vbCr & vbCr & NonRepairs_Final_File.Name & " " & NonRepairs_Final_File.DateLastModified & vbCr & Repairs_Final_File.Name & " " & Repairs_Final_File.DateLastModified
    .Send
End With

End Function

Set fs = Nothing
Set NonRepairs_Final_File = Nothing
Set Repairs_Final_File = Nothing
Set appOutlook = Nothing
Set objMailItem = Nothing
