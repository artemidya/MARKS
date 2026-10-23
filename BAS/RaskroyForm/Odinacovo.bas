Private Function Odinacovo(ByRef RCurves1 As RCurve, ByRef RCurves2 As RCurve) As Boolean
    'Проверяет одинаковый ли объект с RCurves1 и RCurves2
    Dim qq As Long, qq2 As Long
    If UBound(RCurves2.RPaths) <> UBound(RCurves1.RPaths) Then Exit Function
    If RCurves1.MaxLenght <> RCurves2.MaxLenght Then Exit Function
    
    For qq = 1 To UBound(RCurves1.RPaths)
    
        If UBound(RCurves2.RPaths(qq).RSegments) <> UBound(RCurves1.RPaths(qq).RSegments) Then Exit Function
        For qq2 = 1 To UBound(RCurves1.RPaths(qq).RSegments)
            If Int(RCurves1.RPaths(qq).RSegments(qq2).ECPx) <> Int(RCurves2.RPaths(qq).RSegments(qq2).ECPx) Then Exit Function
            If Int(RCurves1.RPaths(qq).RSegments(qq2).ECPy) <> Int(RCurves2.RPaths(qq).RSegments(qq2).ECPy) Then Exit Function
            If Int(RCurves1.RPaths(qq).RSegments(qq2).SCPx) <> Int(RCurves2.RPaths(qq).RSegments(qq2).SCPx) Then Exit Function
            If Int(RCurves1.RPaths(qq).RSegments(qq2).SCPy) <> Int(RCurves2.RPaths(qq).RSegments(qq2).SCPy) Then Exit Function
            If Int(RCurves1.RPaths(qq).RSegments(qq2).SNPx) <> Int(RCurves2.RPaths(qq).RSegments(qq2).SNPx) Then Exit Function
            If Int(RCurves1.RPaths(qq).RSegments(qq2).SNPy) <> Int(RCurves2.RPaths(qq).RSegments(qq2).SNPy) Then Exit Function
        Next
    Next
    Odinacovo = True
End Function
