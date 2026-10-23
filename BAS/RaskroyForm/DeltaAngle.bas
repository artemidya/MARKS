Private Function DeltaAngle(ByRef a1 As Double, ByRef a2 As Double) As Double
    DeltaAngle = a2 - a1
    If DeltaAngle > 3.141593 Then
        DeltaAngle = DeltaAngle - 6.283186
    ElseIf DeltaAngle < -3.141593 Then
        DeltaAngle = DeltaAngle + 6.283186
    End If
End Function
