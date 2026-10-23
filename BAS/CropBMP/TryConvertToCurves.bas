Private Function TryConvertToCurves(shape As VGCore.shape, NewCrv As VGCore.curve) As Boolean
    Dim s2 As VGCore.shape
    Set s2 = shape.Clone
    On Error GoTo 10
    s2.ConvertToCurves
    If s2.Type = cdrCurveShape Then
        TryConvertToCurves = True
        Set NewCrv = Application.CreateCurve
        NewCrv.AppendCurve s2.curve
    End If
10
    s2.Delete
End Function
