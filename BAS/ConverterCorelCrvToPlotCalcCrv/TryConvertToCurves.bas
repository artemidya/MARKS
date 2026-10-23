Private Sub TryConvertToCurves(shape As VGCore.shape, mx As waMatrix)
    Application.Optimization = True
    Dim s2 As VGCore.shape
    Set s2 = shape.Clone
    On Error GoTo 10
    s2.ConvertToCurves
    If s2.Type = cdrCurveShape Then
        GetColor s2
        GetCurve s2.DisplayCurve, mx
    ElseIf s2.Type = cdrGroupShape Then
        GetCurvesCycle s2.shapes, mx
    End If
10
    s2.Delete
    Application.Optimization = False
End Sub
