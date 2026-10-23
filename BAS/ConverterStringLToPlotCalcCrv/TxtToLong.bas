Private Function TxtToLong(ByRef txt As String) As Long
    On Error Resume Next
    TxtToLong = CLng(txt)
End Function
