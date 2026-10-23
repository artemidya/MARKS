Public Function StringToLongArray(text As String) As Long()
    text = Trim(text)
    Dim arrStr() As String
    arrStr = Split(text, ";")
    If UBound(arrStr) > -1 Then
        Dim arr() As Long
        ReDim arr(UBound(arrStr))
        Dim q As Long
        For q = 0 To UBound(arrStr)
            arr(q) = StringToLong(arrStr(q))
        Next
    End If
    StringToLongArray = arr
End Function
