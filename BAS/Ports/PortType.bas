Property Get PortType(OutPutName As String) As PortType
    Dim i As Long
    i = GetIndex(OutPutName)
    If UCase(left(Ports(i).name, 3)) = "COM" Then
        PortType = ComPort
    ElseIf UCase(left(Ports(i).name, 3)) = "LPT" Then
        PortType = LPTport
    Else
        PortType = UnknownPort
    End If
End Property
