Private Function GetVidAndPidFromVirtualPorts(ByVal Port As String) As String
    Port = UCase(RemoveTrailingColon(Port))
    Dim q As Long
    For q = 0 To Interfaces.count - 1
        If Interfaces.isComOrLPT_Adapter(q) Then
            If UCase(Interfaces.PortName(q)) = Port Then
                GetVidAndPidFromVirtualPorts = Interfaces.VidAndPidByIndex(q)
                Exit Function
            End If
        End If
    Next
End Function
