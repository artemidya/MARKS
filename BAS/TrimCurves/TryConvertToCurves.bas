Private Function TryConvertToCurves(shape As VGCore.shape) As Boolean
    On Error GoTo 10
    shape.ConvertToCurves
    If shape.Type = cdrCurveShape Then
        TryConvertToCurves = True
    ElseIf shape.Type = cdrGroupShape Then
        TrimShapes shape.shapes
    End If
10
End Function
