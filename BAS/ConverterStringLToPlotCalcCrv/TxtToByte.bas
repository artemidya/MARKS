Private Function TxtToByte(ByRef txt As String) As Byte
    On Error Resume Next
    TxtToByte = CByte(txt)
End Function
