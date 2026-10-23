Property Get VidAndPidByIndex(ByVal index As Long) As String
    If index < printers.count Then
        VidAndPidByIndex = printers.VidAndPidByIndex(index)
    ElseIf index < (printers.count + Interfaces.count) Then
        index = index - printers.count
        VidAndPidByIndex = Interfaces.VidAndPidByIndex(index)
    ElseIf index < (printers.count + Interfaces.count + Ports.count) Then
        'Ports
        index = index - printers.count - Interfaces.count
        VidAndPidByIndex = GetVidAndPidFromVirtualPorts(Ports.Names(index))
    Else
    End If
End Property
