Public Function StringToLong(text As String) As Long
    On Error GoTo 10
    Dim V As Double
    V = val(text)
    If V > 2147483647 Then V = 2147483647
    If V < -2147483646 Then V = -2147483646
    StringToLong = CLng(V)
10
End Function
