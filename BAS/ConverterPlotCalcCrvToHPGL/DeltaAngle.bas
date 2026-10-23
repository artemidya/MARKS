Private Function DeltaAngle(ByRef a1 As Double, ByRef a2 As Double) As Double
    DeltaAngle = a2 - a1
    If DeltaAngle > 3.14159265358979 Then
        DeltaAngle = DeltaAngle - 6.28318530717959
    ElseIf DeltaAngle < -3.14159265358979 Then
        DeltaAngle = DeltaAngle + 6.28318530717959
    End If
End Function
