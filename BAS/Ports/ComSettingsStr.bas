Property Get ComSettingsStr(OutPutName As String) As String
    On Error GoTo 10
    Dim i As Long:    i = GetIndex(OutPutName)
    Dim RegData1 As New RegData
    With Ports(i)
        If InStr(1, .name, "COM", vbTextCompare) Then
            ComSettingsStr = RegData1.ComSettingsStr(RemoveTrailingColon(Trim(.name)))
        End If
    End With
10
End Property
