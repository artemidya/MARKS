Property Get VidAndPid(OutPutName As String) As String
    Dim i As Long
    i = GetIndex(OutPutName)
    
    With printers(i)
        If .VidAndPid = "" Then
           .VidAndPid = VidAndPidByPrinterName(OutPutName)
        End If
        VidAndPid = .VidAndPid
    End With
End Property
