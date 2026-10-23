Property Get VidAndPidByIndex(ByVal index As Long) As String
    With printers(index)
        If .VidAndPid = "" Then
           .VidAndPid = VidAndPidByPrinterName(.name)
        End If
        VidAndPidByIndex = .VidAndPid
    End With
End Property
