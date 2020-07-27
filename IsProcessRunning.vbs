'example strProcess (excel.exe, saplogon.exe)

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
