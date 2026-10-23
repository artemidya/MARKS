Private Function Angel(ByRef x1 As Long, ByRef y1 As Long, ByRef x2 As Long, ByRef y2 As Long) As Double
    If x1 = x2 Then
        If y2 > y1 Then Angel = 3.141593
    Else
        Angel = (y2 - y1) / (x2 - x1)
        Angel = Atn(Angel) + 1.570796
        If (x2 - x1) < 0 Then Angel = Angel + 3.141593
    End If
End Function
