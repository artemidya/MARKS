Private Function ConvertLng(t As String) As Long
    On Error GoTo 10
    ConvertLng = CLng(t)
    Exit Function
10
    Command4.Enabled = False
End Function
