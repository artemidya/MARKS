Private Function ConvertTextToLong(ByRef value As String) As Long
    On Error GoTo 10
    
    ConvertTextToLong = CLng(val(Replace$(value, ",", ".")) * 1000)
    Exit Function
10
End Function
