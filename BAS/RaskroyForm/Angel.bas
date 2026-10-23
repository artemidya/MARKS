Private Function Angel(ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double) As Double
    If x1 = x2 Then
        If y2 > y1 Then Angel = 3.141593
        Exit Function
    End If
    Angel = (y2 - y1) / (x2 - x1)
    Angel = Atn(Angel) + 1.570796
    If (x2 - x1) < 0 Then Angel = Angel + 3.141593
End Function
