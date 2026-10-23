Function VidAndPidByPrinterName(PrinterName As String) As String
    On Error GoTo 10
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPEntity where Name like '%" & PrinterName & "%'")
    
    Dim vidPos As Long, pidPos As Long
    Dim vid As String, pid As String
    
    For Each objItem In colItems
        vidPos = 0
        pidPos = 0
        vidPos = InStr(1, objItem.deviceId, "VID_", vbTextCompare)
        pidPos = InStr(1, objItem.deviceId, "PID_", vbTextCompare)
'        Debug.Print objItem.Name & " - " & objItem.deviceID
    
        If vidPos > 0 And pidPos > 0 Then
            vid = LCase(Mid(objItem.deviceId, vidPos, 8)) ' VID_XXXX = 8 символов
            pid = LCase(Mid(objItem.deviceId, pidPos, 8)) ' PID_XXXX = 8 символов
            VidAndPidByPrinterName = vid & "&" & pid
            Exit Function
        End If
    Next
10
End Function
