Property Get VidAndPidByIndex(ByVal index As Long) As String
    With Interfaces(index)
        If Not .isComPortAdapter Then
            VidAndPidByIndex = .VidAndPid
        Else 'if com-port - need check vidAndpid from usb interfaces
            VidAndPidByIndex = CheckCOMVidAndPid(.VidAndPid)
        End If
    End With
End Property
