Property Get VidAndPid(OutPutName As String) As String
    Dim index As Long
    index = GetIndex(OutPutName)

    With Interfaces(index)
        If Not .isComPortAdapter Then
            VidAndPid = .VidAndPid
        Else 'if com-port - need check vidAndpid from usb interfaces
            VidAndPid = CheckCOMVidAndPid(.VidAndPid)
        End If
    End With
End Property
