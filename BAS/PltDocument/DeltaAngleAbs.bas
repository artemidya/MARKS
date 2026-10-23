Private Function DeltaAngleAbs(ByRef a1 As Double, ByRef a2 As Double) As Double
    DeltaAngleAbs = a2 - a1
    If DeltaAngleAbs > 3.141593 Then
        DeltaAngleAbs = DeltaAngleAbs - 6.283186
    ElseIf DeltaAngleAbs < -3.141593 Then
        DeltaAngleAbs = DeltaAngleAbs + 6.283186
    End If
    DeltaAngleAbs = Abs(DeltaAngleAbs)
End Function
