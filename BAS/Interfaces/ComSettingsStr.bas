Property Get ComSettingsStr(OutPutName As String) As String
    On Error GoTo 10
    Dim i As Long:    i = GetIndex(OutPutName)
    Dim RegData1 As New RegData
    With Interfaces(i)
        If .isComPortAdapter And InStr(1, .PortName, "COM", vbTextCompare) Then
            ComSettingsStr = RegData1.ComSettingsStr(RemoveTrailingColon(Trim(.PortName)))
        End If
    End With
10
End Property
