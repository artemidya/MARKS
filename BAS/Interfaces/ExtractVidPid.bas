Private Function ExtractVidPid(devicePath As String) As String
    On Error GoTo 10
    Dim parts() As String
    parts = Split(devicePath, "#")
    
    If UBound(parts) > 0 Then
        ExtractVidPid = parts(1)
    End If
    
    Exit Function
10:
End Function
