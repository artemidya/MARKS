Private Function ExtractExparedId(devicePath As String) As String
    On Error GoTo 10
    Dim parts() As String
    parts = Split(devicePath, "#")
    
    If UBound(parts) > 1 Then
        ExtractExparedId = parts(2)
    End If
    
    Exit Function
10:
End Function
