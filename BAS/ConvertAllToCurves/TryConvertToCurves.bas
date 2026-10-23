Private Function TryConvertToCurves(shape As VGCore.shape) As Boolean
    On Error GoTo 10
    shape.ConvertToCurves
    TryConvertToCurves = True
10
End Function
